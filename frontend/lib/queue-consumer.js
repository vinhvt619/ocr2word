import { randomUUID } from "node:crypto";

import { registerDevConsumer } from "@vercel/queue";

import { errorMessage, normalizeOptions, validJobId } from "./api.js";
import { createWorkerReadUrl, createWorkerWriteUrl, deleteBlobs } from "./blob.js";
import {
  DOCX_CONTENT_TYPE,
  MAX_TOTAL_BYTES,
  OCR_CHUNK_MAX_TOTAL_PAGES,
  OCR_CHUNK_TARGET_SCAN_PAGES,
  OCR_CHUNK_TRIGGER_PAGES,
  QUEUE_TOPIC,
  internalSecret,
} from "./config.js";
import { conversionQueue } from "./queue-client.js";
import {
  acquireJobLock,
  addCleanupPathnames,
  getChunk,
  getChunks,
  getJob,
  releaseJobLock,
  saveChunk,
  updateJob,
} from "./store.js";

/** @typedef {import("./types.js").ConvertQueueMessage} ConvertQueueMessage */
/** @typedef {import("./types.js").JobRecord} JobRecord */
/** @typedef {import("./types.js").PdfChunkRecord} PdfChunkRecord */
/** @typedef {{ start_page: number, end_page: number }} PageRange */

class NonRetryableJobError extends Error {}
class JobBusyError extends Error {}

class WorkerResponseError extends Error {
  /**
   * @param {string} message
   * @param {number} status
   */
  constructor(message, status) {
    super(message);
    /** @type {number} */
    this.status = status;
  }
}

/**
 * @param {unknown} item
 * @returns {item is Record<string, unknown>}
 */
function isRecordObject(item) {
  return Boolean(item && typeof item === "object");
}

/**
 * @param {Record<string, unknown>} payload
 * @param {number} status
 * @returns {string}
 */
function workerErrorDetail(payload, status) {
  if (typeof payload.detail === "string") return payload.detail;
  if (Array.isArray(payload.detail)) {
    const messages = payload.detail.flatMap((item) => {
      if (!item || typeof item !== "object") return [];
      const issue = /** @type {Record<string, unknown>} */ (item);
      const location = Array.isArray(issue.loc) ? issue.loc.map(String).join(".") : "payload";
      const message = typeof issue.msg === "string" ? issue.msg : "Dữ liệu không hợp lệ";
      return [`${location}: ${message}`];
    });
    if (messages.length) return `Container OCR từ chối dữ liệu: ${messages.join("; ")}`;
  }
  return `Container OCR trả về HTTP ${status}.`;
}

/**
 * @returns {string}
 */
function workerBaseUrl() {
  const workerBase = process.env.OCR2WORD_WORKER_URL?.trim();
  if (!workerBase) throw new Error("Thiếu service binding OCR2WORD_WORKER_URL.");
  return workerBase.replace(/\/$/, "");
}

/**
 * @param {string} pathname
 * @param {Record<string, unknown>} body
 * @returns {Promise<Record<string, unknown>>}
 */
async function callWorker(pathname, body) {
  const secret = internalSecret();
  const response = await fetch(new URL(pathname, `${workerBaseUrl()}/`), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${secret}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
    signal: AbortSignal.timeout(270_000),
  });
  const payload = /** @type {Record<string, unknown>} */ (
    await response.json().catch(() => ({}))
  );
  if (!response.ok) {
    throw new WorkerResponseError(workerErrorDetail(payload, response.status), response.status);
  }
  return payload;
}

/**
 * @param {unknown} value
 * @param {number} [fallback=0]
 * @returns {number}
 */
function finiteInteger(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : fallback;
}

/**
 * @param {Record<string, unknown>} payload
 * @returns {Array<Record<string, unknown>>}
 */
function reportItems(payload) {
  return Array.isArray(payload.report)
    ? payload.report.filter(isRecordObject)
    : [];
}

/**
 * @param {JobRecord} job
 * @param {string} downloadUrl
 */
function sourceFile(job, downloadUrl) {
  const source = job.files[0];
  return {
    name: source.name,
    size: source.size,
    content_type: source.content_type,
    download_url: downloadUrl,
  };
}

/**
 * @param {ConvertQueueMessage} message
 * @param {JobRecord} [job]
 * @returns {string}
 */
export function stageLockScope(message, job) {
  if (job?.options?.output_format === "xlsx") {
    return "job";
  }
  return message.action === "chunk" ? `chunk-${message.chunk_index ?? "invalid"}` : (message.action || "convert");
}

/**
 * @param {string} jobId
 * @param {number} index
 * @returns {string}
 */
function chunkPathname(jobId, index) {
  return `ocr2word/${jobId}/intermediate/chunk-${String(index + 1).padStart(4, "0")}.docx`;
}

/**
 * @param {number} pageCount
 * @param {number[]} ocrPages
 * @param {number} [maxScanPages]
 * @param {number} [maxTotalPages]
 * @returns {PageRange[]}
 */
export function planPdfChunks(
  pageCount,
  ocrPages,
  maxScanPages = OCR_CHUNK_TARGET_SCAN_PAGES,
  maxTotalPages = OCR_CHUNK_MAX_TOTAL_PAGES,
) {
  if (!Number.isInteger(pageCount) || pageCount < 1) return [];
  const scans = new Set(ocrPages.filter((page) => Number.isInteger(page) && page >= 1 && page <= pageCount));
  /** @type {PageRange[]} */
  const chunks = [];
  let startPage = 1;
  let scanCount = 0;
  for (let page = 1; page <= pageCount; page += 1) {
    const addsScan = scans.has(page) ? 1 : 0;
    const wouldExceedScans = page > startPage && scanCount + addsScan > maxScanPages;
    const wouldExceedPages = page > startPage && page - startPage >= maxTotalPages;
    if (wouldExceedScans || wouldExceedPages) {
      chunks.push({ start_page: startPage, end_page: page - 1 });
      startPage = page;
      scanCount = 0;
    }
    scanCount += addsScan;
  }
  chunks.push({ start_page: startPage, end_page: pageCount });
  return chunks;
}

/**
 * @param {ConvertQueueMessage} message
 * @param {string} idempotencyKey
 * @returns {Promise<void>}
 */
async function enqueue(message, idempotencyKey) {
  if (process.env.NODE_ENV === "development") ensureLocalConversionConsumer();
  await conversionQueue.send(QUEUE_TOPIC, message, {
    idempotencyKey,
    retentionSeconds: 24 * 60 * 60,
  });
}

/**
 * @param {JobRecord} job
 * @returns {Promise<void>}
 */
async function cleanupFinishedJob(job) {
  const pathnames = [...job.files.map((file) => file.pathname)];
  if (job.chunk_total) {
    const chunks = await getChunks(job.job_id, job.chunk_total);
    pathnames.push(...chunks.map((chunk) => chunk.pathname));
  }
  try {
    await deleteBlobs([...new Set(pathnames)]);
  } catch (error) {
    console.warn("Could not delete source/intermediate blobs after conversion", error);
  }
}

/**
 * @param {JobRecord} job
 * @param {Array<Record<string, unknown>>} report
 * @param {string} [message]
 * @returns {Promise<void>}
 */
async function completeJob(
  job,
  report,
  message = "Đã chuyển đổi xong. Kết quả sẵn sàng để tải xuống.",
) {
  await updateJob(job.job_id, {
    status: "complete",
    percent: 100,
    message,
    current_file: null,
    completed_files: job.total_files,
    report,
    error: null,
  });
  await cleanupFinishedJob({ ...job, report });
}

/**
 * @param {JobRecord} job
 * @returns {Promise<void>}
 */
async function processAnalyze(job) {
  if (job.options.output_format === "xlsx") {
    await processLegacy(job);
    return;
  }
  if (job.files.length !== 1) throw new NonRetryableJobError("Luồng phân tích chỉ nhận một file PDF.");
  if (job.pipeline === "chunked_pdf" && job.chunk_total) {
    const chunks = await getChunks(job.job_id, job.chunk_total);
    await Promise.all(chunks.filter((chunk) => chunk.status !== "complete").map((chunk) => enqueue(
      { job_id: job.job_id, action: "chunk", chunk_index: chunk.index },
      `${job.job_id}:chunk:${chunk.index}`,
    )));
    await maybeQueueMerge(job);
    return;
  }
  if (job.pipeline === "legacy" && job.source_page_count) {
    await enqueue({ job_id: job.job_id, action: "convert" }, `${job.job_id}:convert`);
    return;
  }
  await updateJob(job.job_id, {
    status: "processing",
    percent: 22,
    message: "Đang kiểm tra số trang và nhận diện trang scan…",
    error: null,
  });
  const downloadUrl = await createWorkerReadUrl(job.files[0].pathname);
  const payload = await callWorker("/internal/analyze-pdf", {
    job_id: job.job_id,
    file: sourceFile(job, downloadUrl),
    options: job.options,
  });
  const pageCount = finiteInteger(payload.page_count);
  const ocrPages = Array.isArray(payload.ocr_pages)
    ? payload.ocr_pages.map((page) => finiteInteger(page)).filter((page) => page > 0)
    : [];
  if (pageCount < 1) throw new Error("Container OCR không trả về số trang PDF hợp lệ.");

  if (ocrPages.length < OCR_CHUNK_TRIGGER_PAGES) {
    await updateJob(job.job_id, {
      pipeline: "legacy",
      source_page_count: pageCount,
      ocr_page_count: ocrPages.length,
      percent: 25,
      message: ocrPages.length
        ? `PDF có ${ocrPages.length} trang scan; có thể xử lý an toàn trong một lượt.`
        : "PDF có lớp chữ; đang dùng luồng chuyển đổi nhanh giữ bố cục.",
    });
    await enqueue({ job_id: job.job_id, action: "convert" }, `${job.job_id}:convert`);
    return;
  }

  const ranges = planPdfChunks(pageCount, ocrPages);
  /** @type {PdfChunkRecord[]} */
  const chunks = ranges.map((range, index) => ({
    job_id: job.job_id,
    index,
    start_page: range.start_page,
    end_page: range.end_page,
    pathname: chunkPathname(job.job_id, index),
    status: "queued",
    size: 0,
    report: [],
    error: null,
  }));
  await Promise.all(chunks.map(saveChunk));
  await addCleanupPathnames(job.job_id, chunks.map((chunk) => chunk.pathname));
  await updateJob(job.job_id, {
    pipeline: "chunked_pdf",
    source_page_count: pageCount,
    ocr_page_count: ocrPages.length,
    chunk_total: chunks.length,
    chunk_completed: 0,
    status: "queued",
    percent: 25,
    message: `PDF scan nặng đã được chia thành ${chunks.length} phần để tránh timeout Vercel.`,
  });
  await Promise.all(chunks.map((chunk) => enqueue(
    { job_id: job.job_id, action: "chunk", chunk_index: chunk.index },
    `${job.job_id}:chunk:${chunk.index}`,
  )));
}

/**
 * @param {JobRecord} job
 * @returns {Promise<void>}
 */
async function processLegacy(job) {
  await updateJob(job.job_id, {
    pipeline: "legacy",
    status: "processing",
    percent: 26,
    message: "Container OCR đã bắt đầu xử lý tài liệu…",
    error: null,
  });
  const inputUrls = await Promise.all(job.files.map((file) => createWorkerReadUrl(file.pathname)));
  const outputUrl = await createWorkerWriteUrl(job.output_pathname, job.output_content_type, MAX_TOTAL_BYTES);
  const payload = await callWorker("/internal/convert", {
    job_id: job.job_id,
    files: job.files.map((file, index) => ({
      name: file.name,
      size: file.size,
      content_type: file.content_type,
      download_url: inputUrls[index],
    })),
    options: job.options,
    output: {
      name: job.output_name,
      content_type: job.output_content_type,
      upload_url: outputUrl,
    },
    notify_progress: true,
  });
  const report = reportItems(payload);
  await completeJob(job, report);
}

/**
 * @param {JobRecord} job
 * @returns {Promise<void>}
 */
async function maybeQueueMerge(job) {
  const total = job.chunk_total ?? 0;
  if (!total) return;
  const chunks = await getChunks(job.job_id, total);
  if (chunks.length === total && chunks.every((chunk) => chunk.status === "complete")) {
    await enqueue({ job_id: job.job_id, action: "merge" }, `${job.job_id}:merge`);
  }
}

/**
 * @param {JobRecord} job
 * @param {number} chunkIndex
 * @returns {Promise<void>}
 */
async function processChunk(job, chunkIndex) {
  if (job.pipeline !== "chunked_pdf" || !job.chunk_total) {
    throw new NonRetryableJobError("Job không có kế hoạch chia PDF hợp lệ.");
  }
  if (!Number.isInteger(chunkIndex) || chunkIndex < 0 || chunkIndex >= job.chunk_total) {
    throw new NonRetryableJobError("Chỉ số phần PDF không hợp lệ.");
  }
  const chunk = await getChunk(job.job_id, chunkIndex);
  if (!chunk) throw new NonRetryableJobError("Không tìm thấy trạng thái phần PDF.");
  if (chunk.status === "complete") {
    await maybeQueueMerge(job);
    return;
  }
  await saveChunk({ ...chunk, status: "processing", error: null });
  const sourceUrl = await createWorkerReadUrl(job.files[0].pathname);
  const outputUrl = await createWorkerWriteUrl(chunk.pathname, DOCX_CONTENT_TYPE, MAX_TOTAL_BYTES);
  const payload = await callWorker("/internal/convert", {
    job_id: job.job_id,
    files: [sourceFile(job, sourceUrl)],
    options: job.options,
    output: {
      name: `part-${chunk.index + 1}.docx`,
      content_type: DOCX_CONTENT_TYPE,
      upload_url: outputUrl,
    },
    page_range: {
      start_page: chunk.start_page,
      end_page: chunk.end_page,
    },
    notify_progress: false,
  });
  await saveChunk({
    ...chunk,
    status: "complete",
    size: finiteInteger(payload.output_size),
    report: reportItems(payload),
    error: null,
  });
  const chunks = await getChunks(job.job_id, job.chunk_total);
  const completed = chunks.filter((item) => item.status === "complete").length;
  await updateJob(job.job_id, {
    status: "processing",
    percent: 25 + Math.round((completed / job.chunk_total) * 55),
    chunk_completed: completed,
    message: `Đã OCR ${completed}/${job.chunk_total} phần PDF.`,
    current_file: job.files[0].name,
    error: null,
  });
  if (completed === job.chunk_total) await maybeQueueMerge(job);
}

/**
 * @param {JobRecord} job
 * @param {PdfChunkRecord[]} chunks
 * @returns {Array<Record<string, unknown>>}
 */
function combinedChunkReport(job, chunks) {
  const warnings = chunks.flatMap((chunk) => chunk.report.flatMap((item) => (
    Array.isArray(item.warnings) ? item.warnings.map(String) : []
  )));
  const tableCount = chunks.reduce((sum, chunk) => sum + chunk.report.reduce(
    (partSum, item) => partSum + finiteInteger(item.tables),
    0,
  ), 0);
  return [{
    file: job.files[0].name,
    status: "success",
    output: job.output_name,
    warnings: [
      `Đã chia ${job.source_page_count ?? 0} trang thành ${chunks.length} phần OCR và ghép lại tự động.`,
      ...new Set(warnings),
    ],
    pages: job.source_page_count ?? null,
    tables: tableCount,
  }];
}

/**
 * @param {JobRecord} job
 * @returns {Promise<void>}
 */
async function processMerge(job) {
  if (job.pipeline !== "chunked_pdf" || !job.chunk_total) {
    throw new NonRetryableJobError("Job không có các phần PDF để ghép.");
  }
  const chunks = (await getChunks(job.job_id, job.chunk_total)).sort((a, b) => a.index - b.index);
  if (chunks.length !== job.chunk_total || chunks.some((chunk) => chunk.status !== "complete" || chunk.size < 1)) {
    throw new Error("Một hoặc nhiều phần PDF chưa xử lý xong.");
  }
  await updateJob(job.job_id, {
    status: "packaging",
    percent: 82,
    message: `Đang ghép ${chunks.length} phần thành một file Word…`,
    current_file: job.files[0].name,
    error: null,
  });
  const partUrls = await Promise.all(chunks.map((chunk) => createWorkerReadUrl(chunk.pathname)));
  const outputUrl = await createWorkerWriteUrl(job.output_pathname, DOCX_CONTENT_TYPE, MAX_TOTAL_BYTES);
  await callWorker("/internal/merge-docx", {
    job_id: job.job_id,
    parts: chunks.map((chunk, index) => ({
      name: `part-${chunk.index + 1}.docx`,
      size: chunk.size,
      content_type: DOCX_CONTENT_TYPE,
      download_url: partUrls[index],
    })),
    output: {
      name: job.output_name,
      content_type: DOCX_CONTENT_TYPE,
      upload_url: outputUrl,
    },
  });
  const report = combinedChunkReport(job, chunks);
  await completeJob(job, report);
}

/**
 * @param {ConvertQueueMessage} message
 * @param {JobRecord} job
 * @returns {Promise<void>}
 */
export async function runStage(message, job) {
  if (job.options.output_format === "xlsx") {
    if (message.action === "chunk" || message.action === "merge") {
      throw new NonRetryableJobError(`Tác vụ '${message.action}' không được hỗ trợ cho định dạng xuất Excel (.xlsx).`);
    }
  }
  // Rebuild the worker options from the current schema for jobs created before this release.
  const engine = job.options.math_engine;
  job = {
    ...job,
    options: normalizeOptions({
      ...job.options,
      math_engine: engine === "local_hybrid" || engine === "rules" ? engine : undefined,
    }),
  };
  switch (message.action || "convert") {
    case "analyze":
      return processAnalyze(job);
    case "convert":
      return processLegacy(job);
    case "chunk":
      return processChunk(job, message.chunk_index ?? -1);
    case "merge":
      return processMerge(job);
    default:
      throw new NonRetryableJobError("Hành động hàng đợi không hợp lệ.");
  }
}

/**
 * @param {ConvertQueueMessage} message
 * @param {number} deliveryCount
 * @returns {Promise<void>}
 */
export async function processConversionJob(message, deliveryCount) {
  if (!message || !validJobId(message.job_id)) throw new NonRetryableJobError("Thông điệp hàng đợi không hợp lệ.");
  const job = await getJob(message.job_id);
  if (!job || job.status === "complete") return;
  const lockToken = randomUUID();
  const scope = stageLockScope(message, job);
  if (!await acquireJobLock(job.job_id, lockToken, scope)) {
    throw new JobBusyError("Giai đoạn này đang được một worker khác xử lý.");
  }
  try {
    await runStage(message, job);
  } catch (error) {
    if (error instanceof NonRetryableJobError) {
      await updateJob(job.job_id, { status: "error", message: error.message, error: error.message });
      throw error;
    }
    const detail = errorMessage(error, "Container OCR tạm thời không phản hồi.");
    const nonRetryableWorkerError = error instanceof WorkerResponseError
      && error.status >= 400 && error.status < 500;
    const chunkIndex = typeof message.chunk_index === "number" && Number.isInteger(message.chunk_index)
      ? message.chunk_index
      : null;
    if (nonRetryableWorkerError || deliveryCount >= 3) {
      if (message.action === "chunk" && chunkIndex !== null) {
        const chunk = await getChunk(job.job_id, chunkIndex);
        if (chunk) await saveChunk({ ...chunk, status: "error", error: detail });
      }
      await updateJob(job.job_id, {
        status: "error",
        message: nonRetryableWorkerError ? detail : `${detail} Vercel đã thử lại 3 lần.`,
        error: detail,
      });
      throw new NonRetryableJobError(detail);
    }
    if (message.action === "chunk" && chunkIndex !== null) {
      const chunk = await getChunk(job.job_id, chunkIndex);
      if (chunk) await saveChunk({ ...chunk, status: "queued", error: null });
    }
    await updateJob(job.job_id, {
      status: "retrying",
      message: `${detail} Vercel sẽ tự động thử lại riêng giai đoạn này.`,
      error: null,
    });
    throw error;
  } finally {
    await releaseJobLock(job.job_id, lockToken, scope).catch((lockError) => {
      console.warn(`Could not release ${scope} lock for job ${job.job_id}`, lockError);
    });
  }
}

/**
 * @param {unknown} error
 * @param {{ deliveryCount: number }} metadata
 * @returns {{ acknowledge: true } | { afterSeconds: number }}
 */
function retryDirective(error, metadata) {
  if (error instanceof NonRetryableJobError) return { acknowledge: true };
  if (error instanceof JobBusyError) return { afterSeconds: 20 };
  return { afterSeconds: Math.min(120, 15 * 2 ** Math.max(0, metadata.deliveryCount - 1)) };
}

export const conversionCallbackOptions = {
  visibilityTimeoutSeconds: 5 * 60,
  retry: retryDirective,
};

let localConsumerRegistered = false;

/**
 * @returns {void}
 */
export function ensureLocalConversionConsumer() {
  if (process.env.NODE_ENV !== "development" || localConsumerRegistered) return;
  registerDevConsumer({
    topic: QUEUE_TOPIC,
    client: conversionQueue,
    consumerGroup: "ocr2word-local-convert",
    handler: async (message, metadata) => processConversionJob(
      /** @type {ConvertQueueMessage} */ (message),
      metadata.deliveryCount,
    ),
    ...conversionCallbackOptions,
  });
  localConsumerRegistered = true;
}
