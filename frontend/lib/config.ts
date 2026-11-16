export const MAX_FILES = 20;
export const MAX_FILE_BYTES = 150 * 1024 * 1024;
export const MAX_TOTAL_BYTES = 300 * 1024 * 1024;
export const JOB_TTL_SECONDS = 25 * 60 * 60;
export const RESULT_TTL_MS = 24 * 60 * 60 * 1000;
export const SIGNED_UPLOAD_TTL_MS = 60 * 60 * 1000;
export const SIGNED_WORKER_TTL_MS = 60 * 60 * 1000;
export const SIGNED_DOWNLOAD_TTL_MS = 15 * 60 * 1000;
export const QUEUE_TOPIC = "ocr2word-convert";
export const CLEANUP_SET_KEY = "ocr2word:cleanup:index";
export const OCR_CHUNK_TRIGGER_PAGES = 9;
export const OCR_CHUNK_TARGET_SCAN_PAGES = 6;
export const OCR_CHUNK_MAX_TOTAL_PAGES = 40;
export const DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
export const XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export const PDF_EXTENSIONS = new Set(["pdf"]);
export const IMAGE_EXTENSIONS = new Set(["png", "jpg", "jpeg", "tif", "tiff", "bmp", "webp"]);
export const EXCEL_EXTENSIONS = new Set(["xlsx", "xlsm", "xls"]);
export const ALLOWED_EXTENSIONS = new Set([
  ...PDF_EXTENSIONS,
  ...IMAGE_EXTENSIONS,
  ...EXCEL_EXTENSIONS,
]);

const MIME_TYPES: Record<string, string> = {
  pdf: "application/pdf",
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  tif: "image/tiff",
  tiff: "image/tiff",
  bmp: "image/bmp",
  webp: "image/webp",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
  xls: "application/vnd.ms-excel",
};

export function extensionOf(filename: string): string {
  const dot = filename.lastIndexOf(".");
  return dot >= 0 ? filename.slice(dot + 1).toLowerCase() : "";
}

export function mimeTypeFor(filename: string): string {
  return MIME_TYPES[extensionOf(filename)] ?? "application/octet-stream";
}

export function safeDownloadStem(filename: string): string {
  const extension = extensionOf(filename);
  const withoutExtension = extension ? filename.slice(0, -(extension.length + 1)) : filename;
  const cleaned = withoutExtension
    .normalize("NFC")
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/^[ .-]+|[ .-]+$/g, "")
    .slice(0, 100);
  return cleaned || "tai-lieu";
}

export function outputDetails(
  files: Array<{ name: string }>,
  outputFormat: "docx" | "xlsx" = "docx",
) {
  if (files.length === 1) {
    if (outputFormat === "xlsx") {
      return {
        name: `${safeDownloadStem(files[0].name)}-editable.xlsx`,
        extension: "xlsx",
        contentType: XLSX_CONTENT_TYPE,
      };
    }
    const suffix = "-editable.docx";
    return {
      name: `${safeDownloadStem(files[0].name)}${suffix}`,
      extension: "docx",
      contentType: DOCX_CONTENT_TYPE,
    };
  }
  return {
    name: "ocr2word-ket-qua.zip",
    extension: "zip",
    contentType: "application/zip",
  };
}

export function internalSecret(): string {
  const value = process.env.OCR2WORD_INTERNAL_SECRET?.trim();
  if (value) return value;
  if (!process.env.VERCEL) return "ocr2word-local-development-only";
  throw new Error("Thiếu biến môi trường OCR2WORD_INTERNAL_SECRET.");
}
