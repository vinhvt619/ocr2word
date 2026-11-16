import type { NextApiRequest, NextApiResponse } from "next";

import { errorMessage, validJobId } from "../../../../lib/api";
import { extensionOf, PDF_EXTENSIONS, QUEUE_TOPIC } from "../../../../lib/config";
import { conversionQueue } from "../../../../lib/queue-client";
import { getJob, updateJob } from "../../../../lib/store";
import type { ConvertQueueMessage } from "../../../../lib/types";

export default async function handler(request: NextApiRequest, response: NextApiResponse) {
  if (request.method !== "POST") {
    response.setHeader("Allow", "POST");
    return response.status(405).json({ detail: "Phương thức không được hỗ trợ." });
  }
  const jobId = Array.isArray(request.query.jobId) ? request.query.jobId[0] : request.query.jobId;
  if (!validJobId(jobId)) return response.status(404).json({ detail: "Không tìm thấy lượt chuyển đổi." });
  const job = await getJob(jobId);
  if (!job) return response.status(404).json({ detail: "Lượt chuyển đổi đã hết hạn hoặc không tồn tại." });
  if (job.status !== "uploading") return response.status(200).json({ job_id: jobId, status: job.status });
  try {
    const shouldAnalyze = job.files.length === 1
      && PDF_EXTENSIONS.has(extensionOf(job.files[0].name))
      && job.options.ocr_mode !== "text"
      && job.options.output_format !== "xlsx";
    const action: ConvertQueueMessage["action"] = shouldAnalyze ? "analyze" : "convert";
    await updateJob(jobId, {
      status: "queued",
      percent: 21,
      message: shouldAnalyze
        ? "Tài liệu đã tải lên; đang phân tích PDF để chọn luồng tối ưu…"
        : "Tài liệu đã tải lên; đang chờ container OCR…",
      error: null,
      pipeline: shouldAnalyze ? "analyzing_pdf" : "legacy",
    });
    const message: ConvertQueueMessage = { job_id: jobId, action };
    if (process.env.NODE_ENV === "development") {
      const { ensureLocalConversionConsumer } = await import("../../../../lib/queue-consumer");
      ensureLocalConversionConsumer();
    }
    await conversionQueue.send(QUEUE_TOPIC, message, {
      idempotencyKey: `${jobId}:${action}`,
      retentionSeconds: 24 * 60 * 60,
    });
    return response.status(202).json({ job_id: jobId, status: "queued" });
  } catch (error) {
    const detail = errorMessage(error, "Không thể đưa tài liệu vào hàng đợi Vercel.");
    await updateJob(jobId, { status: "error", message: detail, error: detail });
    return response.status(503).json({ detail });
  }
}
