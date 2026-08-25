import type { NextApiRequest, NextApiResponse } from "next";

import { validJobId } from "../../../../lib/api";
import { createDownloadUrl } from "../../../../lib/blob";
import { getJob } from "../../../../lib/store";

export default async function handler(request: NextApiRequest, response: NextApiResponse) {
  if (request.method !== "GET") {
    response.setHeader("Allow", "GET");
    return response.status(405).json({ detail: "Phương thức không được hỗ trợ." });
  }
  const jobId = Array.isArray(request.query.jobId) ? request.query.jobId[0] : request.query.jobId;
  if (!validJobId(jobId)) return response.status(404).json({ detail: "Không tìm thấy lượt chuyển đổi." });
  const job = await getJob(jobId);
  if (!job) return response.status(404).json({ detail: "Lượt chuyển đổi đã hết hạn hoặc không tồn tại." });
  const result: Record<string, unknown> = { ...job, files: undefined, options: undefined, output_pathname: undefined };
  if (job.status === "complete") {
    result.download_url = await createDownloadUrl(job.output_pathname);
    result.download_name = job.output_name;
  }
  response.setHeader("Cache-Control", "no-store");
  return response.status(200).json(result);
}
