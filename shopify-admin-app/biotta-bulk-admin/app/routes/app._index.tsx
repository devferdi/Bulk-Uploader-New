import { useState } from "react";
import type { ChangeEvent, CSSProperties } from "react";
import type { HeadersFunction, LoaderFunctionArgs } from "react-router";
import { useRouteError } from "react-router";
import { boundary } from "@shopify/shopify-app-react-router/server";

import { authenticate } from "../shopify.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  await authenticate.admin(request);
  return null;
};

const JOB_POLL_INTERVAL_MS = 1500;
const JOB_POLL_TIMEOUT_MS = 10 * 60 * 1000;

type ModuleId =
  | "products"
  | "metaobjects"
  | "collections"
  | "file-alt-texts";
type BulkJobAction = "download" | "upload";
type BulkJobStatus = "queued" | "running" | "completed" | "failed";
type BulkJobSummary = {
  action: BulkJobAction;
  error?: string;
  fileName?: string;
  jobId: string;
  status: BulkJobStatus;
};
type BulkWorkflowState = {
  isDownloading: boolean;
  isUploading: boolean;
  selectedFile: File | null;
};
type WorkflowState = Record<ModuleId, BulkWorkflowState>;
type BulkWorkflowModule = {
  defaultDownloadFileName: string;
  defaultUploadFileName: string;
  downloadButtonLabel: string;
  downloadDescription: string;
  downloadEndpoint: string;
  downloadFinishedError: string;
  downloadStartError: string;
  downloadSuccessMessage: string;
  dropZoneLabel: string;
  fileRouteBasePath: string;
  id: ModuleId;
  jobRouteBasePath: string;
  kicker: string;
  summary: string;
  title: string;
  uploadButtonLabel: string;
  uploadDescription: string;
  uploadEndpoint: string;
  uploadFinishedError: string;
  uploadStartError: string;
  uploadSuccessMessage: string;
  workerFileDownloadError: string;
};

type ShopifyToastOptions = {
  duration?: number;
  isError?: boolean;
};

type ShopifyGlobal = {
  toast?: {
    show: (message: string, options?: ShopifyToastOptions) => void;
  };
};

const BULK_WORKFLOW_MODULES: BulkWorkflowModule[] = [
  {
    id: "products",
    kicker: "Products",
    title: "Catalog spreadsheet",
    summary:
      "Download product data, variants, images, metafields, and inventory columns into one workbook.",
    downloadDescription:
      "Create a fresh XLSX workbook from Shopify before making bulk changes.",
    uploadDescription:
      "Upload the edited workbook to apply product changes with the existing Python sheet logic.",
    dropZoneLabel: "Products spreadsheet",
    downloadButtonLabel: "Download workbook",
    uploadButtonLabel: "Upload workbook",
    downloadEndpoint: "/app/products/download",
    uploadEndpoint: "/app/products/upload",
    jobRouteBasePath: "/app/products/jobs",
    fileRouteBasePath: "/app/products/jobs",
    defaultDownloadFileName: "shopify_products_bulk.xlsx",
    defaultUploadFileName: "shopify_products_bulk_updated.xlsx",
    downloadStartError: "Could not start the products download job.",
    downloadFinishedError: "Could not finish the products download job.",
    uploadStartError: "Could not start the products upload job.",
    uploadFinishedError: "Could not finish the products upload job.",
    workerFileDownloadError:
      "Could not download the completed products spreadsheet.",
    downloadSuccessMessage: "Products spreadsheet downloaded",
    uploadSuccessMessage: "Upload complete. Updated spreadsheet downloaded",
  },
  {
    id: "metaobjects",
    kicker: "Metaobjects",
    title: "Content models workbook",
    summary:
      "Export metaobject records, definitions, handles, status, and field values into a single file.",
    downloadDescription:
      "Pull the latest metaobject records into XLSX so edits start from the current store state.",
    uploadDescription:
      "Upload the edited workbook to create or update records, handles, statuses, and field values.",
    dropZoneLabel: "Metaobjects spreadsheet",
    downloadButtonLabel: "Download workbook",
    uploadButtonLabel: "Upload workbook",
    downloadEndpoint: "/app/metaobjects/download",
    uploadEndpoint: "/app/metaobjects/upload",
    jobRouteBasePath: "/app/bulk-jobs",
    fileRouteBasePath: "/app/bulk-jobs",
    defaultDownloadFileName: "shopify_metaobjects.xlsx",
    defaultUploadFileName: "shopify_metaobjects_updated.xlsx",
    downloadStartError: "Could not start the metaobjects download job.",
    downloadFinishedError: "Could not finish the metaobjects download job.",
    uploadStartError: "Could not start the metaobjects upload job.",
    uploadFinishedError: "Could not finish the metaobjects upload job.",
    workerFileDownloadError:
      "Could not download the completed metaobjects spreadsheet.",
    downloadSuccessMessage: "Metaobjects spreadsheet downloaded",
    uploadSuccessMessage:
      "Upload complete. Updated metaobjects spreadsheet downloaded",
  },
  {
    id: "collections",
    kicker: "Collections",
    title: "Collection rules workbook",
    summary:
      "Manage manual and smart collections, products, conditions, images, and collection metafields.",
    downloadDescription:
      "Create an XLSX snapshot of your current custom and smart collections before editing.",
    uploadDescription:
      "Upload the workbook after editing to apply collection updates back into Shopify.",
    dropZoneLabel: "Collections spreadsheet",
    downloadButtonLabel: "Download workbook",
    uploadButtonLabel: "Upload workbook",
    downloadEndpoint: "/app/collections/download",
    uploadEndpoint: "/app/collections/upload",
    jobRouteBasePath: "/app/bulk-jobs",
    fileRouteBasePath: "/app/bulk-jobs",
    defaultDownloadFileName: "shopify_collections.xlsx",
    defaultUploadFileName: "shopify_collections_updated.xlsx",
    downloadStartError: "Could not start the collections download job.",
    downloadFinishedError: "Could not finish the collections download job.",
    uploadStartError: "Could not start the collections upload job.",
    uploadFinishedError: "Could not finish the collections upload job.",
    workerFileDownloadError:
      "Could not download the completed collections spreadsheet.",
    downloadSuccessMessage: "Collections spreadsheet downloaded",
    uploadSuccessMessage:
      "Upload complete. Updated collections spreadsheet downloaded",
  },
  {
    id: "file-alt-texts",
    kicker: "Files",
    title: "Alt text workbook",
    summary:
      "Review uploaded files, current alt text, file IDs, and URLs, then bulk-update alt text in one pass.",
    downloadDescription:
      "Export your files into XLSX so the team can review or rewrite alt text offline.",
    uploadDescription:
      "Edit the `New Alt Text` column in the workbook, then upload it to update Shopify file alt text.",
    dropZoneLabel: "Files alt-text spreadsheet",
    downloadButtonLabel: "Download workbook",
    uploadButtonLabel: "Upload workbook",
    downloadEndpoint: "/app/file-alt-texts/download",
    uploadEndpoint: "/app/file-alt-texts/upload",
    jobRouteBasePath: "/app/bulk-jobs",
    fileRouteBasePath: "/app/bulk-jobs",
    defaultDownloadFileName: "shopify_uploaded_files_alt_texts.xlsx",
    defaultUploadFileName: "shopify_uploaded_files_alt_texts_updated.xlsx",
    downloadStartError: "Could not start the files alt-text download job.",
    downloadFinishedError: "Could not finish the files alt-text download job.",
    uploadStartError: "Could not start the files alt-text upload job.",
    uploadFinishedError: "Could not finish the files alt-text upload job.",
    workerFileDownloadError:
      "Could not download the completed files alt-text spreadsheet.",
    downloadSuccessMessage: "Files alt-text spreadsheet downloaded",
    uploadSuccessMessage:
      "Upload complete. Updated files alt-text spreadsheet downloaded",
  },
];

const dashboardGridStyle: CSSProperties = {
  alignItems: "start",
  display: "grid",
  gap: "1rem",
  gridTemplateColumns: "repeat(auto-fit, minmax(28rem, 1fr))",
};

const moduleCardStyle: CSSProperties = {
  background: "#ffffff",
  border: "1px solid rgba(15, 23, 42, 0.08)",
  borderRadius: "20px",
  boxShadow: "0 14px 34px rgba(15, 23, 42, 0.06)",
  display: "grid",
  gap: "1rem",
  padding: "1.25rem",
};

const moduleHeaderStyle: CSSProperties = {
  display: "grid",
  gap: "0.5rem",
};

const moduleKickerStyle: CSSProperties = {
  color: "#667085",
  fontSize: "0.75rem",
  fontWeight: 700,
  letterSpacing: "0.08em",
  margin: 0,
  textTransform: "uppercase",
};

const moduleTitleStyle: CSSProperties = {
  color: "#101828",
  fontSize: "1.25rem",
  lineHeight: 1.2,
  margin: 0,
};

const moduleSummaryStyle: CSSProperties = {
  color: "#475467",
  fontSize: "0.95rem",
  lineHeight: 1.6,
  margin: 0,
};

const moduleActionGridStyle: CSSProperties = {
  display: "grid",
  gap: "0.875rem",
  gridTemplateColumns: "repeat(auto-fit, minmax(16rem, 1fr))",
};

const moduleActionCardStyle: CSSProperties = {
  background: "linear-gradient(180deg, #ffffff 0%, #f8fafc 100%)",
  border: "1px solid rgba(15, 23, 42, 0.08)",
  borderRadius: "16px",
  display: "grid",
  gap: "0.75rem",
  padding: "1rem",
};

const moduleActionLabelStyle: CSSProperties = {
  color: "#667085",
  fontSize: "0.75rem",
  fontWeight: 700,
  letterSpacing: "0.04em",
  margin: 0,
  textTransform: "uppercase",
};

const moduleActionTitleStyle: CSSProperties = {
  color: "#101828",
  fontSize: "1rem",
  lineHeight: 1.3,
  margin: 0,
};

const pageShellStyle: CSSProperties = {
  background:
    "radial-gradient(circle at top left, rgba(59, 130, 246, 0.08), transparent 30%), linear-gradient(180deg, #f8fafc 0%, #eef2f7 100%)",
  minHeight: "100vh",
  padding: "1.5rem",
};

const pageInnerStyle: CSSProperties = {
  display: "grid",
  gap: "1.5rem",
  margin: "0 auto",
  maxWidth: "1400px",
};

const pageHeadingStyle: CSSProperties = {
  color: "#101828",
  fontSize: "2rem",
  lineHeight: 1.1,
  margin: 0,
};

const pageSubheadingStyle: CSSProperties = {
  color: "#475467",
  fontSize: "1rem",
  lineHeight: 1.6,
  margin: 0,
  maxWidth: "52rem",
};

const actionBodyStyle: CSSProperties = {
  display: "grid",
  gap: "0.75rem",
};

const bodyCopyStyle: CSSProperties = {
  color: "#475467",
  fontSize: "0.95rem",
  lineHeight: 1.6,
  margin: 0,
};

const inputLabelStyle: CSSProperties = {
  color: "#344054",
  fontSize: "0.875rem",
  fontWeight: 600,
  margin: 0,
};

const fileInputStyle: CSSProperties = {
  background: "#ffffff",
  border: "1px solid #d0d5dd",
  borderRadius: "12px",
  color: "#101828",
  fontSize: "0.95rem",
  padding: "0.875rem 1rem",
  width: "100%",
};

const buttonRowStyle: CSSProperties = {
  display: "flex",
  flexWrap: "wrap",
  gap: "0.75rem",
};

const primaryButtonStyle: CSSProperties = {
  appearance: "none",
  background: "#111827",
  border: "1px solid #111827",
  borderRadius: "999px",
  color: "#ffffff",
  cursor: "pointer",
  fontSize: "0.95rem",
  fontWeight: 600,
  minHeight: "2.75rem",
  padding: "0.75rem 1.25rem",
};

const secondaryButtonStyle: CSSProperties = {
  appearance: "none",
  background: "#ffffff",
  border: "1px solid #d0d5dd",
  borderRadius: "999px",
  color: "#101828",
  cursor: "pointer",
  fontSize: "0.95rem",
  fontWeight: 600,
  minHeight: "2.75rem",
  padding: "0.75rem 1.25rem",
};

const selectedFileStyle: CSSProperties = {
  color: "#0f172a",
  fontSize: "0.9rem",
  lineHeight: 1.5,
  margin: 0,
};

function createInitialWorkflowState(): WorkflowState {
  return {
    collections: {
      isDownloading: false,
      isUploading: false,
      selectedFile: null,
    },
    "file-alt-texts": {
      isDownloading: false,
      isUploading: false,
      selectedFile: null,
    },
    metaobjects: {
      isDownloading: false,
      isUploading: false,
      selectedFile: null,
    },
    products: {
      isDownloading: false,
      isUploading: false,
      selectedFile: null,
    },
  };
}

function delay(ms: number) {
  return new Promise<void>((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

function getDownloadFileName(response: Response, fallbackFileName: string) {
  const headerValue = response.headers.get("content-disposition") ?? "";
  const utf8Match = headerValue.match(/filename\*=UTF-8''([^;]+)/i);
  const asciiMatch = headerValue.match(/filename="([^"]+)"/i);
  const matchedFileName = utf8Match?.[1] ?? asciiMatch?.[1];

  if (!matchedFileName) {
    return fallbackFileName;
  }

  try {
    return decodeURIComponent(matchedFileName);
  } catch {
    return matchedFileName;
  }
}

async function getErrorMessage(response: Response, fallbackMessage: string) {
  const contentType = response.headers.get("content-type") ?? "";

  if (contentType.includes("application/json")) {
    const payload = await response.json().catch(() => null);
    if (payload && typeof payload.error === "string") {
      return payload.error;
    }
  }

  const responseText = await response.text().catch(() => "");
  return responseText.trim() || fallbackMessage;
}

function triggerFileDownload(blob: Blob, fileName: string) {
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");

  anchor.href = objectUrl;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();

  window.setTimeout(() => {
    URL.revokeObjectURL(objectUrl);
  }, 0);
}

function showToast(message: string, options?: ShopifyToastOptions) {
  if (typeof window === "undefined") {
    return;
  }

  const shopifyGlobal = (window as Window & { shopify?: ShopifyGlobal }).shopify;
  shopifyGlobal?.toast?.show(message, options);
}

export default function Index() {
  const [workflowState, setWorkflowState] = useState<WorkflowState>(
    createInitialWorkflowState,
  );

  function updateWorkflowState(
    moduleId: ModuleId,
    nextState: Partial<BulkWorkflowState>,
  ) {
    setWorkflowState((currentState) => ({
      ...currentState,
      [moduleId]: {
        ...currentState[moduleId],
        ...nextState,
      },
    }));
  }

  async function startBulkJob(
    endpoint: string,
    init: RequestInit,
    fallbackMessage: string,
  ) {
    const response = await fetch(endpoint, init);

    if (!response.ok) {
      throw new Error(await getErrorMessage(response, fallbackMessage));
    }

    const payload = (await response.json().catch(() => null)) as
      | BulkJobSummary
      | null;

    if (!payload || typeof payload.jobId !== "string") {
      throw new Error(fallbackMessage);
    }

    return payload;
  }

  async function waitForBulkJob(
    module: BulkWorkflowModule,
    job: BulkJobSummary,
    fallbackMessage: string,
  ) {
    const deadline = Date.now() + JOB_POLL_TIMEOUT_MS;

    while (Date.now() < deadline) {
      const response = await fetch(`${module.jobRouteBasePath}/${job.jobId}`);

      if (!response.ok) {
        throw new Error(await getErrorMessage(response, fallbackMessage));
      }

      const payload = (await response.json().catch(() => null)) as
        | BulkJobSummary
        | null;

      if (!payload || typeof payload.status !== "string") {
        throw new Error(fallbackMessage);
      }

      if (payload.status === "completed") {
        return payload;
      }

      if (payload.status === "failed") {
        throw new Error(payload.error || fallbackMessage);
      }

      await delay(JOB_POLL_INTERVAL_MS);
    }

    throw new Error(
      "The spreadsheet job is still running. Please try again in a moment.",
    );
  }

  async function downloadCompletedJobFile(
    module: BulkWorkflowModule,
    jobId: string,
    fallbackFileName: string,
  ) {
    const response = await fetch(`${module.fileRouteBasePath}/${jobId}/file`);

    if (!response.ok) {
      throw new Error(
        await getErrorMessage(response, module.workerFileDownloadError),
      );
    }

    const spreadsheetBlob = await response.blob();
    const fileName = getDownloadFileName(response, fallbackFileName);

    return { fileName, spreadsheetBlob };
  }

  async function handleDownload(module: BulkWorkflowModule) {
    updateWorkflowState(module.id, { isDownloading: true });

    try {
      const job = await startBulkJob(
        module.downloadEndpoint,
        { method: "POST" },
        module.downloadStartError,
      );
      const completedJob = await waitForBulkJob(
        module,
        job,
        module.downloadFinishedError,
      );
      const { fileName, spreadsheetBlob } = await downloadCompletedJobFile(
        module,
        completedJob.jobId,
        completedJob.fileName || module.defaultDownloadFileName,
      );

      triggerFileDownload(spreadsheetBlob, fileName);
      showToast(module.downloadSuccessMessage);
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : "Could not download the workbook.";

      showToast(message, { duration: 5000, isError: true });
    } finally {
      updateWorkflowState(module.id, { isDownloading: false });
    }
  }

  function handleFileSelection(
    moduleId: ModuleId,
    event: ChangeEvent<HTMLInputElement>,
  ) {
    const files = Array.from(event.currentTarget.files ?? []);
    const file = files[0] ?? null;

    updateWorkflowState(moduleId, { selectedFile: file });
  }

  async function handleUpload(
    module: BulkWorkflowModule,
  ) {
    const selectedFile = workflowState[module.id].selectedFile;

    if (!selectedFile) {
      showToast("Choose an .xlsx spreadsheet before uploading.", {
        duration: 5000,
        isError: true,
      });
      return;
    }

    if (!selectedFile.name.toLowerCase().endsWith(".xlsx")) {
      showToast("Please upload an .xlsx spreadsheet.", {
        duration: 5000,
        isError: true,
      });
      return;
    }

    updateWorkflowState(module.id, { isUploading: true });

    try {
      const formData = new FormData();
      formData.append("spreadsheet", selectedFile);

      const job = await startBulkJob(
        module.uploadEndpoint,
        {
          body: formData,
          method: "POST",
        },
        module.uploadStartError,
      );
      const completedJob = await waitForBulkJob(
        module,
        job,
        module.uploadFinishedError,
      );
      const { fileName, spreadsheetBlob } = await downloadCompletedJobFile(
        module,
        completedJob.jobId,
        completedJob.fileName || module.defaultUploadFileName,
      );

      triggerFileDownload(spreadsheetBlob, fileName);
      updateWorkflowState(module.id, { selectedFile: null });
      showToast(module.uploadSuccessMessage);
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : "Could not upload the workbook.";

      showToast(message, { duration: 5000, isError: true });
    } finally {
      updateWorkflowState(module.id, { isUploading: false });
    }
  }

  return (
    <div style={pageShellStyle}>
      <div style={pageInnerStyle}>
        <div style={moduleHeaderStyle}>
          <h1 style={pageHeadingStyle}>HUX Bulk Loader</h1>
          <p style={pageSubheadingStyle}>
            Download the current Shopify workbook, make your edits offline, and
            upload the revised spreadsheet when you are ready.
          </p>
        </div>

        <div style={dashboardGridStyle}>
        {BULK_WORKFLOW_MODULES.map((module) => {
          const state = workflowState[module.id];
          const actionButtonStyle = state.isDownloading
            ? { ...secondaryButtonStyle, cursor: "wait", opacity: 0.7 }
            : primaryButtonStyle;
          const uploadButtonStyle = state.isUploading || !state.selectedFile
            ? { ...secondaryButtonStyle, cursor: "not-allowed", opacity: 0.6 }
            : primaryButtonStyle;

          return (
            <section key={module.id} style={moduleCardStyle}>
              <div style={moduleHeaderStyle}>
                <p style={moduleKickerStyle}>{module.kicker}</p>
                <h2 style={moduleTitleStyle}>{module.title}</h2>
                <p style={moduleSummaryStyle}>{module.summary}</p>
              </div>

              <div style={moduleActionGridStyle}>
                <div style={moduleActionCardStyle}>
                  <div style={actionBodyStyle}>
                    <p style={moduleActionLabelStyle}>Download</p>
                    <h3 style={moduleActionTitleStyle}>Current workbook</h3>
                    <p style={bodyCopyStyle}>{module.downloadDescription}</p>
                    <div style={buttonRowStyle}>
                      <button
                        type="button"
                        style={actionButtonStyle}
                        onClick={() => handleDownload(module)}
                        disabled={state.isUploading || state.isDownloading}
                      >
                        {state.isDownloading
                          ? "Working..."
                          : module.downloadButtonLabel}
                      </button>
                    </div>
                  </div>
                </div>

                <div style={moduleActionCardStyle}>
                  <div style={actionBodyStyle}>
                    <p style={moduleActionLabelStyle}>Upload</p>
                    <h3 style={moduleActionTitleStyle}>Apply your edits</h3>
                    <p style={bodyCopyStyle}>{module.uploadDescription}</p>
                    <label style={inputLabelStyle}>
                      {module.dropZoneLabel}
                      <input
                        type="file"
                        accept=".xlsx"
                        onChange={(event) =>
                          handleFileSelection(module.id, event)
                        }
                        disabled={state.isUploading}
                        style={fileInputStyle}
                      />
                    </label>
                    {state.selectedFile ? (
                      <p style={selectedFileStyle}>
                        Selected file: {state.selectedFile.name}
                      </p>
                    ) : null}
                    <div style={buttonRowStyle}>
                      <button
                        type="button"
                        style={uploadButtonStyle}
                        onClick={() => handleUpload(module)}
                        disabled={
                          !state.selectedFile ||
                          state.isDownloading ||
                          state.isUploading
                        }
                      >
                        {state.isUploading
                          ? "Working..."
                          : module.uploadButtonLabel}
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            </section>
          );
        })}
        </div>
      </div>
    </div>
  );
}

export function ErrorBoundary() {
  return boundary.error(useRouteError());
}

export const headers: HeadersFunction = (headersArgs) => {
  return boundary.headers(headersArgs);
};
