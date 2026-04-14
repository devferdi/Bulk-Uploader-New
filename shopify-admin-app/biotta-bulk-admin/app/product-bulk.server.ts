import { randomUUID } from "node:crypto";
import { spawn, spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const APP_DIR = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(APP_DIR, "../../..");
const WORKER_SCRIPT_PATH = path.join(
  REPO_ROOT,
  "backend",
  "product_bulk_worker.py",
);
const DEFAULT_PYTHON_COMMAND = "python3";
const WORKER_API_VERSION = "2026-01";
const XLSX_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const PRODUCT_JOB_TTL_MS = 30 * 60 * 1000;

type BulkResource =
  | "products"
  | "metaobjects"
  | "collections"
  | "file-alt-texts";
type ProductJobAction = "download" | "upload";
type ProductWorkerAction =
  | "products-download"
  | "products-upload"
  | "metaobjects-download"
  | "metaobjects-upload"
  | "collections-download"
  | "collections-upload"
  | "file-alt-texts-download"
  | "file-alt-texts-upload";
type ProductJobStatus = "queued" | "running" | "completed" | "failed";

type ProductWorkerResult = {
  action?: ProductWorkerAction | ProductJobAction;
  error?: string;
  file_path?: string | null;
};

type ProductWorkerOptions = {
  accessToken: string;
  action: ProductWorkerAction;
  inputFilePath?: string;
  outputDir?: string;
  shop: string;
};

type ProductSpreadsheetOptions = {
  accessToken: string;
  shop: string;
};

type ProductUploadOptions = ProductSpreadsheetOptions & {
  file: File;
};

type ProductJob = {
  action: ProductJobAction;
  createdAt: number;
  error: string | null;
  fileName: string | null;
  filePath: string | null;
  id: string;
  resource: BulkResource;
  status: ProductJobStatus;
  tempDir: string | null;
  updatedAt: number;
};

export type ProductJobSummary = {
  action: ProductJobAction;
  createdAt: string;
  error?: string;
  fileName?: string;
  jobId: string;
  resource: BulkResource;
  status: ProductJobStatus;
  updatedAt: string;
};

type ProductJobGlobal = typeof globalThis & {
  __biottaProductJobs?: Map<string, ProductJob>;
};

const globalProductJobs = globalThis as ProductJobGlobal;
const productJobs =
  globalProductJobs.__biottaProductJobs ??
  new Map<string, ProductJob>();

if (!globalProductJobs.__biottaProductJobs) {
  globalProductJobs.__biottaProductJobs = productJobs;
}

function sanitizeFileName(fileName: string) {
  const normalizedName = fileName.trim() || "shopify_products.xlsx";
  return normalizedName.replace(/[^a-zA-Z0-9._-]/g, "_");
}

function pythonHasSpreadsheetDependencies(pythonExecutable: string) {
  const result = spawnSync(
    pythonExecutable,
    ["-c", "import pandas, openpyxl, requests, bs4, lxml, openai"],
    {
      cwd: REPO_ROOT,
      env: {
        ...process.env,
        PYTHONIOENCODING: "utf-8",
        PYTHONUNBUFFERED: "1",
      },
      stdio: "ignore",
    },
  );

  return result.status === 0;
}

function getPythonExecutable() {
  const configuredExecutable = process.env.PRODUCT_BULK_PYTHON?.trim();

  if (configuredExecutable) {
    return configuredExecutable;
  }

  const candidatePaths = [
    path.join(REPO_ROOT, "backend", ".venv", "bin", "python"),
    path.join(REPO_ROOT, ".venv", "bin", "python"),
    path.join(REPO_ROOT, "venv", "bin", "python"),
  ];

  for (const candidatePath of candidatePaths) {
    if (
      existsSync(candidatePath) &&
      pythonHasSpreadsheetDependencies(candidatePath)
    ) {
      return candidatePath;
    }
  }

  return DEFAULT_PYTHON_COMMAND;
}

function parseWorkerResult(stdout: string) {
  const resultLine = stdout
    .trim()
    .split(/\r?\n/)
    .reverse()
    .find((line) => line.startsWith("WORKER_RESULT_JSON="));

  if (!resultLine) {
    return null;
  }

  const rawJson = resultLine.slice("WORKER_RESULT_JSON=".length);
  return JSON.parse(rawJson) as ProductWorkerResult;
}

function formatWorkerError(message: string, stdout: string, stderr: string) {
  const details = [stderr.trim(), stdout.trim()].filter(Boolean).join("\n\n");
  return details ? `${message}\n\n${details}` : message;
}

async function runProductWorker({
  accessToken,
  action,
  inputFilePath,
  outputDir,
  shop,
}: ProductWorkerOptions) {
  const args = [
    WORKER_SCRIPT_PATH,
    action,
    "--shop",
    shop,
    "--access-token",
    accessToken,
    "--api-version",
    WORKER_API_VERSION,
    "--script-dir",
    REPO_ROOT,
  ];

  if (outputDir) {
    args.push("--output-dir", outputDir);
  }

  if (inputFilePath) {
    args.push("--file", inputFilePath);
  }

  return await new Promise<{ filePath: string }>((resolve, reject) => {
    const pythonExecutable = getPythonExecutable();
    const worker = spawn(pythonExecutable, args, {
      cwd: REPO_ROOT,
      env: {
        ...process.env,
        PYTHONIOENCODING: "utf-8",
        PYTHONUNBUFFERED: "1",
      },
    });

    let stdout = "";
    let stderr = "";

    worker.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });

    worker.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });

    worker.on("error", (error) => {
      reject(
        new Error(
          formatWorkerError(
            `Failed to start the product spreadsheet worker with ${pythonExecutable}: ${error.message}`,
            stdout,
            stderr,
          ),
        ),
      );
    });

    worker.on("close", (code) => {
      let workerResult: ProductWorkerResult | null = null;

      try {
        workerResult = parseWorkerResult(stdout);
      } catch (error) {
        reject(
          new Error(
            formatWorkerError(
              `Could not parse the worker response: ${
                error instanceof Error ? error.message : "Unknown error"
              }`,
              stdout,
              stderr,
            ),
          ),
        );
        return;
      }

      if (workerResult?.error) {
        reject(
          new Error(formatWorkerError(workerResult.error, stdout, stderr)),
        );
        return;
      }

      if (code !== 0) {
        reject(
          new Error(
            formatWorkerError(
              `The product spreadsheet worker exited with code ${code}.`,
              stdout,
              stderr,
            ),
          ),
        );
        return;
      }

      if (!workerResult?.file_path) {
        reject(
          new Error(
            formatWorkerError(
              "The product spreadsheet worker did not return a file path.",
              stdout,
              stderr,
            ),
          ),
        );
        return;
      }

      resolve({ filePath: workerResult.file_path });
    });
  });
}

function buildSpreadsheetResponse(fileBuffer: Buffer, fileName: string) {
  return new Response(new Uint8Array(fileBuffer), {
    headers: {
      "Cache-Control": "no-store",
      "Content-Disposition": `attachment; filename="${fileName}"`,
      "Content-Type": XLSX_CONTENT_TYPE,
    },
  });
}

async function removeJobArtifacts(job: ProductJob) {
  if (!job.tempDir) {
    return;
  }

  await rm(job.tempDir, { force: true, recursive: true }).catch(() => undefined);
}

async function pruneExpiredJobs() {
  const expirationCutoff = Date.now() - PRODUCT_JOB_TTL_MS;

  for (const [jobId, job] of productJobs.entries()) {
    if (job.updatedAt >= expirationCutoff) {
      continue;
    }

    await removeJobArtifacts(job);
    productJobs.delete(jobId);
  }
}

function serializeProductJob(job: ProductJob): ProductJobSummary {
  return {
    action: job.action,
    createdAt: new Date(job.createdAt).toISOString(),
    ...(job.error ? { error: job.error } : {}),
    ...(job.fileName ? { fileName: job.fileName } : {}),
    jobId: job.id,
    resource: job.resource,
    status: job.status,
    updatedAt: new Date(job.updatedAt).toISOString(),
  };
}

function createProductJob(resource: BulkResource, action: ProductJobAction) {
  const now = Date.now();
  const job: ProductJob = {
    action,
    createdAt: now,
    error: null,
    fileName: null,
    filePath: null,
    id: randomUUID(),
    resource,
    status: "queued",
    tempDir: null,
    updatedAt: now,
  };

  productJobs.set(job.id, job);
  return job;
}

function updateProductJob(
  jobId: string,
  updates: Partial<Omit<ProductJob, "action" | "createdAt" | "id">>,
) {
  const existingJob = productJobs.get(jobId);

  if (!existingJob) {
    return null;
  }

  const updatedJob: ProductJob = {
    ...existingJob,
    ...updates,
    updatedAt: Date.now(),
  };

  productJobs.set(jobId, updatedJob);
  return updatedJob;
}

async function runProductJobInBackground(
  jobId: string,
  {
    accessToken,
    action,
    resource,
    file,
    shop,
  }: ProductSpreadsheetOptions & {
    action: ProductJobAction;
    file?: File;
    resource: BulkResource;
  },
) {
  let tempDir: string | null = null;

  try {
    tempDir = await mkdtemp(
      path.join(
        tmpdir(),
        `biotta-${resource}-${action}-`,
      ),
    );

    updateProductJob(jobId, {
      error: null,
      status: "running",
      tempDir,
    });

    let inputFilePath: string | undefined;

    if (action === "upload") {
      if (!file) {
        throw new Error("Choose an .xlsx spreadsheet before uploading.");
      }

      const inputFileName = sanitizeFileName(file.name);
      inputFilePath = path.join(tempDir, inputFileName);
      const fileBuffer = Buffer.from(await file.arrayBuffer());

      await writeFile(inputFilePath, fileBuffer);
    }

    const { filePath } = await runProductWorker({
      accessToken,
      action: `${resource}-${action}`,
      inputFilePath,
      outputDir: action === "download" ? tempDir : undefined,
      shop,
    });

    updateProductJob(jobId, {
      error: null,
      fileName: sanitizeFileName(path.basename(filePath)),
      filePath,
      status: "completed",
      tempDir,
    });
  } catch (error) {
    updateProductJob(jobId, {
      error:
        error instanceof Error
          ? error.message
          : "The product spreadsheet job failed.",
      status: "failed",
      tempDir,
    });
  }
}

async function startBulkDownloadJob(
  resource: BulkResource,
  {
    accessToken,
    shop,
  }: ProductSpreadsheetOptions,
) {
  await pruneExpiredJobs();

  const job = createProductJob(resource, "download");
  void runProductJobInBackground(job.id, {
    accessToken,
    action: "download",
    resource,
    shop,
  });

  return serializeProductJob(job);
}

async function startBulkUploadJob(
  resource: BulkResource,
  {
    accessToken,
    file,
    shop,
  }: ProductUploadOptions,
) {
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    throw new Error("Please upload an .xlsx spreadsheet.");
  }

  await pruneExpiredJobs();

  const job = createProductJob(resource, "upload");
  void runProductJobInBackground(job.id, {
    accessToken,
    action: "upload",
    file,
    resource,
    shop,
  });

  return serializeProductJob(job);
}

export async function startProductDownloadJob(options: ProductSpreadsheetOptions) {
  return startBulkDownloadJob("products", options);
}

export async function startProductUploadJob(options: ProductUploadOptions) {
  return startBulkUploadJob("products", options);
}

export async function startMetaobjectDownloadJob(options: ProductSpreadsheetOptions) {
  return startBulkDownloadJob("metaobjects", options);
}

export async function startMetaobjectUploadJob(options: ProductUploadOptions) {
  return startBulkUploadJob("metaobjects", options);
}

export async function startCollectionDownloadJob(options: ProductSpreadsheetOptions) {
  return startBulkDownloadJob("collections", options);
}

export async function startCollectionUploadJob(options: ProductUploadOptions) {
  return startBulkUploadJob("collections", options);
}

export async function startFileAltTextsDownloadJob(
  options: ProductSpreadsheetOptions,
) {
  return startBulkDownloadJob("file-alt-texts", options);
}

export async function startFileAltTextsUploadJob(options: ProductUploadOptions) {
  return startBulkUploadJob("file-alt-texts", options);
}

export async function getProductJobSummary(jobId: string) {
  await pruneExpiredJobs();

  const job = productJobs.get(jobId);
  return job ? serializeProductJob(job) : null;
}

export async function getBulkJobSummary(jobId: string) {
  return getProductJobSummary(jobId);
}

export async function buildProductJobFileResponse(jobId: string) {
  await pruneExpiredJobs();

  const job = productJobs.get(jobId);

  if (!job) {
    throw new Error("This product spreadsheet job was not found or has expired.");
  }

  if (job.status !== "completed" || !job.filePath) {
    throw new Error("This product spreadsheet job has not finished yet.");
  }

  const fileBuffer = await readFile(job.filePath);
  return buildSpreadsheetResponse(
    fileBuffer,
    job.fileName ?? sanitizeFileName(path.basename(job.filePath)),
  );
}

export async function buildBulkJobFileResponse(jobId: string) {
  return buildProductJobFileResponse(jobId);
}
