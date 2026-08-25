import type { NextApiRequest, NextApiResponse } from "next";

import { internalRequestAuthorized, requestBody, validJobId } from "../../../lib/api";
import { getJob, updateJob } from "../../../lib/store";
import type { JobStatus } from "../../../lib/types";

export default async function handler(request: NextApiRequest, response: NextApiResponse) {
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
  const allowedStatuses = new Set<JobStatus>(["processing", "packaging"]);
  const requestedStatus = typeof body.status === "string" ? body.status as JobStatus : "processing";
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
