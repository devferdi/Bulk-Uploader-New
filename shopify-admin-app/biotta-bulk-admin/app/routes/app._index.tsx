import { useRef, useState } from "react";
import type { ChangeEvent, CSSProperties } from "react";
import type { HeadersFunction, LoaderFunctionArgs } from "react-router";
import { useRouteError } from "react-router";
import { boundary } from "@shopify/shopify-app-react-router/server";

import { authenticateAdminWithTrace } from "../shopify.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  await authenticateAdminWithTrace(request, "routes/app._index");
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
      "Variants, images, metafields, and inventory in one XLSX export.",
    downloadDescription:
      "Create a fresh workbook from the live store.",
    uploadDescription:
      "Upload your edited workbook to apply product changes.",
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
      "Definitions, handles, statuses, and values in one spreadsheet.",
    downloadDescription:
      "Pull the latest metaobject records into XLSX.",
    uploadDescription:
      "Upload the workbook to create or update records.",
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
      "Manual and smart collections, rules, images, and metafields.",
    downloadDescription:
      "Export the current collection setup into XLSX.",
    uploadDescription:
      "Upload the workbook to sync collection changes.",
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
      "File URLs, current alt text, and new alt text updates.",
    downloadDescription:
      "Export your files into a workbook for review.",
    uploadDescription:
      "Upload the workbook to update Shopify file alt text.",
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
  display: "grid",
  gap: "1rem",
};

const moduleCardStyle: CSSProperties = {
  background: "#ffffff",
  border: "1px solid #e4e7ec",
  borderRadius: "24px",
  boxShadow: "0 1px 2px rgba(16, 24, 40, 0.04)",
  display: "grid",
  gap: "1rem",
  padding: "1.35rem",
};

const moduleHeaderStyle: CSSProperties = {
  display: "grid",
  gap: "0.45rem",
};

const moduleKickerStyle: CSSProperties = {
  color: "#667085",
  fontSize: "0.72rem",
  fontWeight: 700,
  letterSpacing: "0.08em",
  margin: 0,
  textTransform: "uppercase",
};

const moduleTitleStyle: CSSProperties = {
  color: "#101828",
  fontSize: "1.5rem",
  fontWeight: 700,
  lineHeight: 1.2,
  margin: 0,
};

const moduleSummaryStyle: CSSProperties = {
  color: "#475467",
  fontSize: "0.98rem",
  lineHeight: 1.55,
  margin: 0,
  maxWidth: "52rem",
};

const moduleActionGridStyle: CSSProperties = {
  display: "grid",
  gap: "1rem",
  gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
};

const moduleActionCardStyle: CSSProperties = {
  background: "#f8fafc",
  border: "1px solid #e4e7ec",
  borderRadius: "20px",
  display: "flex",
  flexDirection: "column",
  gap: "1rem",
  minHeight: "17rem",
  padding: "1.15rem",
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
  fontSize: "1.1rem",
  fontWeight: 700,
  lineHeight: 1.3,
  margin: 0,
};

const pageShellStyle: CSSProperties = {
  background: "#f3f4f6",
  fontFamily:
    'Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
  minHeight: "100vh",
  padding: "1rem",
};

const pageInnerStyle: CSSProperties = {
  display: "grid",
  gap: "1rem",
  margin: "0 auto",
  maxWidth: "1280px",
};

const actionBodyStyle: CSSProperties = {
  display: "grid",
  gap: "0.65rem",
};

const actionFooterStyle: CSSProperties = {
  display: "grid",
  gap: "0.85rem",
  marginTop: "auto",
};

const bodyCopyStyle: CSSProperties = {
  color: "#475467",
  fontSize: "0.95rem",
  lineHeight: 1.55,
  margin: 0,
};

const hiddenFileInputStyle: CSSProperties = {
  display: "none",
};

const filePickerStyle: CSSProperties = {
  display: "grid",
  gap: "0.7rem",
};

const filePickerRowStyle: CSSProperties = {
  display: "flex",
  flexWrap: "wrap",
  gap: "0.75rem",
  width: "100%",
};

const primaryButtonStyle: CSSProperties = {
  alignItems: "center",
  appearance: "none",
  background: "#111827",
  border: "1px solid #111827",
  borderRadius: "14px",
  color: "#ffffff",
  cursor: "pointer",
  display: "inline-flex",
  fontSize: "0.95rem",
  fontWeight: 700,
  justifyContent: "center",
  minHeight: "3rem",
  padding: "0.85rem 1.1rem",
  width: "100%",
};

const secondaryButtonStyle: CSSProperties = {
  alignItems: "center",
  appearance: "none",
  background: "#ffffff",
  border: "1px solid #d0d5dd",
  borderRadius: "14px",
  color: "#101828",
  cursor: "pointer",
  display: "inline-flex",
  fontSize: "0.95rem",
  fontWeight: 600,
  justifyContent: "center",
  minHeight: "3rem",
  padding: "0.85rem 1rem",
  width: "100%",
};

const fileChooseButtonStyle: CSSProperties = {
  ...secondaryButtonStyle,
  minWidth: "10.5rem",
  width: "auto",
};

const fileNameStyle: CSSProperties = {
  alignItems: "center",
  background: "#ffffff",
  border: "1px solid #d0d5dd",
  borderRadius: "14px",
  color: "#475467",
  display: "flex",
  flex: "1 1 14rem",
  fontSize: "0.92rem",
  minHeight: "3rem",
  margin: 0,
  overflow: "hidden",
  padding: "0.85rem 1rem",
  textOverflow: "ellipsis",
  whiteSpace: "nowrap",
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
  const fileInputRefs = useRef<
    Partial<Record<ModuleId, HTMLInputElement | null>>
  >({});

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

  function setFileInputRef(moduleId: ModuleId) {
    return (element: HTMLInputElement | null) => {
      fileInputRefs.current[moduleId] = element;
    };
  }

  function openFilePicker(moduleId: ModuleId) {
    const input = fileInputRefs.current[moduleId];

    if (!input) {
      return;
    }

    input.value = "";
    input.click();
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
      const input = fileInputRefs.current[module.id];
      if (input) {
        input.value = "";
      }
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
                  </div>
                  <div style={actionFooterStyle}>
                    <div>
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
                  </div>
                  <div style={actionFooterStyle}>
                    <div style={filePickerStyle}>
                      <input
                        ref={setFileInputRef(module.id)}
                        type="file"
                        accept=".xlsx"
                        onChange={(event) =>
                          handleFileSelection(module.id, event)
                        }
                        disabled={state.isUploading}
                        aria-label={module.dropZoneLabel}
                        style={hiddenFileInputStyle}
                      />
                      <div style={filePickerRowStyle}>
                        <button
                          type="button"
                          style={
                            state.isUploading
                              ? {
                                  ...fileChooseButtonStyle,
                                  cursor: "wait",
                                  opacity: 0.7,
                                }
                              : fileChooseButtonStyle
                          }
                          onClick={() => openFilePicker(module.id)}
                          disabled={state.isUploading || state.isDownloading}
                        >
                          Choose workbook
                        </button>
                        <p
                          style={{
                            ...fileNameStyle,
                            color: state.selectedFile ? "#101828" : "#667085",
                          }}
                          title={state.selectedFile?.name ?? "No workbook selected"}
                        >
                          {state.selectedFile?.name ?? "No workbook selected"}
                        </p>
                      </div>
                    </div>
                    <div>
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
