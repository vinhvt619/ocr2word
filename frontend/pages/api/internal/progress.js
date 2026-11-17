import { internalRequestAuthorized, requestBody, validJobId } from "../../../lib/api.js";
import { getJob, updateJob } from "../../../lib/store.js";

/** @typedef {import("next").NextApiRequest} NextApiRequest */
/** @typedef {import("next").NextApiResponse} NextApiResponse */
/** @typedef {import("../../../lib/types.js").JobStatus} JobStatus */

/**
 * @param {NextApiRequest} request
 * @param {NextApiResponse} response
 */
export default async function handler(request, response) {
  if (request.method !== "POST") {
    response.setHeader("Allow", "POST");
    return response.status(405).json({ detail: "Phương thức không được hỗ trợ." });
  }
  if (!internalRequestAuthorized(request)) return response.status(401).json({ detail: "Không được phép." });
  const body = requestBody(request);
  if (!validJobId(body.job_id)) return response.status(404).json({ detail: "Không tìm thấy lượt chuyển đổi." });
  const job = await getJob(body.job_id);
  if (!job) return response.status(404).json({ detail: "Không tìm thấy lượt chuyển đổi." });
  if (["complete", "error"].includes(job.status)) return response.status(202).json({ ok: true });

  const requestedPercent = Math.round(Number(body.percent));
  const percent = Number.isFinite(requestedPercent)
    ? Math.max(job.percent, Math.min(99, Math.max(0, requestedPercent)))
    : job.percent;
  /** @type {Set<JobStatus>} */
  const allowedStatuses = new Set(["processing", "packaging"]);
  const requestedStatus = typeof body.status === "string"
    ? /** @type {JobStatus} */ (body.status)
    : "processing";
  const status = allowedStatuses.has(requestedStatus) ? requestedStatus : "processing";
  const completedFiles = Math.max(
    job.completed_files,
    Math.min(job.total_files, Math.max(0, Math.round(Number(body.completed_files ?? job.completed_files)))),
  );
  await updateJob(job.job_id, {
    status,
    percent,
    message: typeof body.message === "string" ? body.message.slice(0, 500) : job.message,
    current_file: typeof body.current_file === "string" ? body.current_file.slice(0, 255) : null,
    completed_files: completedFiles,
  });
  return response.status(202).json({ ok: true });
}
