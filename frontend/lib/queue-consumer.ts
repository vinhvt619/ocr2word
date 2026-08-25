import { randomUUID } from "node:crypto";

import { registerDevConsumer } from "@vercel/queue";

import { errorMessage, validJobId } from "./api";
import { createWorkerReadUrl, createWorkerWriteUrl, deleteBlobs } from "./blob";
import {
  DOCX_CONTENT_TYPE,
  MAX_TOTAL_BYTES,
  OCR_CHUNK_MAX_TOTAL_PAGES,
  OCR_CHUNK_TARGET_SCAN_PAGES,
  OCR_CHUNK_TRIGGER_PAGES,
  QUEUE_TOPIC,
  internalSecret,
} from "./config";
import { conversionQueue } from "./queue-client";
import {
  acquireJobLock,
  addCleanupPathnames,
  getChunk,
  getChunks,
  getJob,
  releaseJobLock,
  saveChunk,
  updateJob,
} from "./store";
import type { ConvertQueueMessage, JobRecord, PdfChunkRecord } from "./types";

class NonRetryableJobError extends Error {}
class JobBusyError extends Error {}

class WorkerResponseError extends Error {
  constructor(message: string, readonly status: number) {
    super(message);
  }
}

interface PageRange {
  start_page: number;
  end_page: number;
}

function workerErrorDetail(payload: Record<string, unknown>, status: number): string {
  if (typeof payload.detail === "string") return payload.detail;
  if (Array.isArray(payload.detail)) {
    const messages = payload.detail.flatMap((item) => {
      if (!item || typeof item !== "object") return [];
      const issue = item as Record<string, unknown>;
      const location = Array.isArray(issue.loc) ? issue.loc.map(String).join(".") : "payload";
      const message = typeof issue.msg === "string" ? issue.msg : "Dữ liệu không hợp lệ";
      return [`${location}: ${message}`];
    });
    if (messages.length) return `Container OCR từ chối dữ liệu: ${messages.join("; ")}`;
  }
  return `Container OCR trả về HTTP ${status}.`;
}

function workerBaseUrl(): string {
  const workerBase = process.env.OCR2WORD_WORKER_URL?.trim();
  if (!workerBase) throw new Error("Thiếu service binding OCR2WORD_WORKER_URL.");
  return workerBase.replace(/\/$/, "");
}

async function callWorker(pathname: string, body: Record<string, unknown>): Promise<Record<string, unknown>> {
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
  const payload = await response.json().catch(() => ({})) as Record<string, unknown>;
  if (!response.ok) {
    throw new WorkerResponseError(workerErrorDetail(payload, response.status), response.status);
  }
  return payload;
}

function finiteInteger(value: unknown, fallback = 0): number {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : fallback;
}

function reportItems(payload: Record<string, unknown>): Array<Record<string, unknown>> {
  return Array.isArray(payload.report)
    ? payload.report.filter((item): item is Record<string, unknown> => Boolean(item && typeof item === "object"))
    : [];
}

function sourceFile(job: JobRecord, downloadUrl: string) {
  const source = job.files[0];
  return {
    name: source.name,
    size: source.size,
    content_type: source.content_type,
    download_url: downloadUrl,
  };
}

function stageLockScope(message: ConvertQueueMessage): string {
  return message.action === "chunk" ? `chunk-${message.chunk_index ?? "invalid"}` : (message.action || "convert");
}

function chunkPathname(jobId: string, index: number): string {
  return `ocr2word/${jobId}/intermediate/chunk-${String(index + 1).padStart(4, "0")}.docx`;
}

function aiInputPathname(jobId: string): string {
  return `ocr2word/${jobId}/intermediate/before-ai.docx`;
}

export function planPdfChunks(
  pageCount: number,
  ocrPages: number[],
  maxScanPages = OCR_CHUNK_TARGET_SCAN_PAGES,
  maxTotalPages = OCR_CHUNK_MAX_TOTAL_PAGES,
): PageRange[] {
  if (!Number.isInteger(pageCount) || pageCount < 1) return [];
  const scans = new Set(ocrPages.filter((page) => Number.isInteger(page) && page >= 1 && page <= pageCount));
  const chunks: PageRange[] = [];
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

async function enqueue(message: ConvertQueueMessage, idempotencyKey: string): Promise<void> {
  if (process.env.NODE_ENV === "development") ensureLocalConversionConsumer();
  await conversionQueue.send(QUEUE_TOPIC, message, {
    idempotencyKey,
    retentionSeconds: 24 * 60 * 60,
  });
}

function needsSeparateAi(job: JobRecord): boolean {
  return job.options.ai_assist && job.files.length === 1 && job.output_content_type === DOCX_CONTENT_TYPE;
}

async function cleanupFinishedJob(job: JobRecord): Promise<void> {
  const pathnames = [...job.files.map((file) => file.pathname)];
  if (job.ai_input_pathname) pathnames.push(job.ai_input_pathname);
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

async function completeJob(
  job: JobRecord,
  report: Array<Record<string, unknown>>,
  message = "Đã chuyển đổi xong. Kết quả sẵn sàng để tải xuống.",
): Promise<void> {
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

async function queueAiStage(
  job: JobRecord,
  inputPathname: string,
  inputSize: number,
  report: Array<Record<string, unknown>>,
): Promise<void> {
  if (inputSize < 1) throw new Error("Container OCR không trả về kích thước Word trung gian hợp lệ.");
  await addCleanupPathnames(job.job_id, [inputPathname]);
  await updateJob(job.job_id, {
    status: "queued",
    percent: 88,
    message: "Đã tạo Word; đang chờ AI hiệu đính phần chữ…",
    current_file: job.files[0]?.name ?? null,
    report,
    ai_input_pathname: inputPathname,
    ai_input_size: inputSize,
    error: null,
  });
  await enqueue({ job_id: job.job_id, action: "ai" }, `${job.job_id}:ai`);
}

async function processAnalyze(job: JobRecord): Promise<void> {
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
    options: { ...job.options, ai_assist: false },
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
  const chunks: PdfChunkRecord[] = ranges.map((range, index) => ({
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

async function processLegacy(job: JobRecord): Promise<void> {
  const separateAi = needsSeparateAi(job);
  const targetPathname = separateAi ? aiInputPathname(job.job_id) : job.output_pathname;
  if (separateAi) await addCleanupPathnames(job.job_id, [targetPathname]);
  await updateJob(job.job_id, {
    pipeline: "legacy",
    status: "processing",
    percent: 26,
    message: "Container OCR đã bắt đầu xử lý tài liệu…",
    error: null,
  });
  const inputUrls = await Promise.all(job.files.map((file) => createWorkerReadUrl(file.pathname)));
  const outputUrl = await createWorkerWriteUrl(targetPathname, job.output_content_type, MAX_TOTAL_BYTES);
  const payload = await callWorker("/internal/convert", {
    job_id: job.job_id,
    files: job.files.map((file, index) => ({
      name: file.name,
      size: file.size,
      content_type: file.content_type,
      download_url: inputUrls[index],
    })),
    options: separateAi ? { ...job.options, ai_assist: false } : job.options,
    output: {
      name: separateAi ? "before-ai.docx" : job.output_name,
      content_type: job.output_content_type,
      upload_url: outputUrl,
    },
    notify_progress: true,
  });
  const report = reportItems(payload);
  if (separateAi) {
    await queueAiStage(job, targetPathname, finiteInteger(payload.output_size), report);
  } else {
    await completeJob(job, report);
  }
}

async function maybeQueueMerge(job: JobRecord): Promise<void> {
  const total = job.chunk_total ?? 0;
  if (!total) return;
  const chunks = await getChunks(job.job_id, total);
  if (chunks.length === total && chunks.every((chunk) => chunk.status === "complete")) {
    await enqueue({ job_id: job.job_id, action: "merge" }, `${job.job_id}:merge`);
  }
}

async function processChunk(job: JobRecord, chunkIndex: number): Promise<void> {
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
    options: { ...job.options, ai_assist: false },
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

function combinedChunkReport(job: JobRecord, chunks: PdfChunkRecord[]): Array<Record<string, unknown>> {
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
    ai: null,
  }];
}

async function processMerge(job: JobRecord): Promise<void> {
  if (job.pipeline !== "chunked_pdf" || !job.chunk_total) {
    throw new NonRetryableJobError("Job không có các phần PDF để ghép.");
  }
  if (needsSeparateAi(job) && job.ai_input_pathname && job.ai_input_size) {
    await enqueue({ job_id: job.job_id, action: "ai" }, `${job.job_id}:ai`);
    return;
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
  const separateAi = needsSeparateAi(job);
  const targetPathname = separateAi ? aiInputPathname(job.job_id) : job.output_pathname;
  if (separateAi) await addCleanupPathnames(job.job_id, [targetPathname]);
  const partUrls = await Promise.all(chunks.map((chunk) => createWorkerReadUrl(chunk.pathname)));
  const outputUrl = await createWorkerWriteUrl(targetPathname, DOCX_CONTENT_TYPE, MAX_TOTAL_BYTES);
  const payload = await callWorker("/internal/merge-docx", {
    job_id: job.job_id,
    parts: chunks.map((chunk, index) => ({
      name: `part-${chunk.index + 1}.docx`,
      size: chunk.size,
      content_type: DOCX_CONTENT_TYPE,
      download_url: partUrls[index],
    })),
    output: {
      name: separateAi ? "before-ai.docx" : job.output_name,
      content_type: DOCX_CONTENT_TYPE,
      upload_url: outputUrl,
    },
  });
  const report = combinedChunkReport(job, chunks);
  if (separateAi) {
    await queueAiStage(job, targetPathname, finiteInteger(payload.output_size), report);
  } else {
    await completeJob(job, report);
  }
}

async function processAi(job: JobRecord): Promise<void> {
  if (!job.ai_input_pathname || !job.ai_input_size) {
    throw new NonRetryableJobError("Không tìm thấy bản Word trung gian cho bước AI.");
  }
  await updateJob(job.job_id, {
    status: "processing",
    percent: 89,
    message: "AI đang hiệu đính lỗi dấu và từ bị dính; định dạng Word được giữ nguyên…",
    current_file: job.files[0]?.name ?? null,
    error: null,
  });
  const inputUrl = await createWorkerReadUrl(job.ai_input_pathname);
  const outputUrl = await createWorkerWriteUrl(job.output_pathname, DOCX_CONTENT_TYPE, MAX_TOTAL_BYTES);
  const payload = await callWorker("/internal/ai-assist", {
    job_id: job.job_id,
    file: {
      name: "before-ai.docx",
      size: job.ai_input_size,
      content_type: DOCX_CONTENT_TYPE,
      download_url: inputUrl,
    },
    output: {
      name: job.output_name,
      content_type: DOCX_CONTENT_TYPE,
      upload_url: outputUrl,
    },
  });
  const warnings = Array.isArray(payload.warnings) ? payload.warnings.map(String) : [];
  const report = job.report.length ? job.report.map((item, index) => index === 0 ? {
    ...item,
    warnings: [...(Array.isArray(item.warnings) ? item.warnings : []), ...warnings],
    ai: payload.ai && typeof payload.ai === "object" ? payload.ai : null,
  } : item) : [{
    file: job.files[0]?.name ?? "tài liệu",
    status: "success",
    output: job.output_name,
    warnings,
    ai: payload.ai && typeof payload.ai === "object" ? payload.ai : null,
  }];
  const fallback = payload.fallback === true;
  await completeJob(
    job,
    report,
    fallback
      ? "Đã tạo Word; AI tạm thời không khả dụng nên nội dung OCR được giữ nguyên."
      : "Đã chuyển đổi và hiệu đính AI xong. Kết quả sẵn sàng để tải xuống.",
  );
}

async function runStage(message: ConvertQueueMessage, job: JobRecord): Promise<void> {
  switch (message.action || "convert") {
    case "analyze":
      return processAnalyze(job);
    case "convert":
      return processLegacy(job);
    case "chunk":
      return processChunk(job, message.chunk_index ?? -1);
    case "merge":
      return processMerge(job);
    case "ai":
      return processAi(job);
    default:
      throw new NonRetryableJobError("Hành động hàng đợi không hợp lệ.");
  }
}

export async function processConversionJob(message: ConvertQueueMessage, deliveryCount: number): Promise<void> {
  if (!message || !validJobId(message.job_id)) throw new NonRetryableJobError("Thông điệp hàng đợi không hợp lệ.");
  const job = await getJob(message.job_id);
  if (!job || job.status === "complete") return;
  const lockToken = randomUUID();
  const scope = stageLockScope(message);
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
    if (nonRetryableWorkerError || deliveryCount >= 3) {
      if (message.action === "chunk" && Number.isInteger(message.chunk_index)) {
        const chunk = await getChunk(job.job_id, message.chunk_index!);
        if (chunk) await saveChunk({ ...chunk, status: "error", error: detail });
      }
      await updateJob(job.job_id, {
        status: "error",
        message: nonRetryableWorkerError ? detail : `${detail} Vercel đã thử lại 3 lần.`,
        error: detail,
      });
      throw new NonRetryableJobError(detail);
    }
    if (message.action === "chunk" && Number.isInteger(message.chunk_index)) {
      const chunk = await getChunk(job.job_id, message.chunk_index!);
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

function retryDirective(error: unknown, metadata: { deliveryCount: number }) {
  if (error instanceof NonRetryableJobError) return { acknowledge: true } as const;
  if (error instanceof JobBusyError) return { afterSeconds: 20 } as const;
  return { afterSeconds: Math.min(120, 15 * 2 ** Math.max(0, metadata.deliveryCount - 1)) } as const;
}

export const conversionCallbackOptions = {
  visibilityTimeoutSeconds: 5 * 60,
  retry: retryDirective,
};

let localConsumerRegistered = false;

export function ensureLocalConversionConsumer(): void {
  if (process.env.NODE_ENV !== "development" || localConsumerRegistered) return;
  registerDevConsumer({
    topic: QUEUE_TOPIC,
    client: conversionQueue,
    consumerGroup: "ocr2word-local-convert",
    handler: async (message, metadata) => processConversionJob(message as ConvertQueueMessage, metadata.deliveryCount),
    ...conversionCallbackOptions,
  });
  localConsumerRegistered = true;
}
