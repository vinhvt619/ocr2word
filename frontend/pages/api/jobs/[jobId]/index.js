import { validJobId } from "../../../../lib/api.js";
import { createDownloadUrl } from "../../../../lib/blob.js";
import { getJob } from "../../../../lib/store.js";

/** @typedef {import("next").NextApiRequest} NextApiRequest */
/** @typedef {import("next").NextApiResponse} NextApiResponse */

/**
 * @param {NextApiRequest} request
 * @param {NextApiResponse} response
 */
export default async function handler(request, response) {
  if (request.method !== "GET") {
    response.setHeader("Allow", "GET");
    return response.status(405).json({ detail: "Phương thức không được hỗ trợ." });
  }
  const jobId = Array.isArray(request.query.jobId) ? request.query.jobId[0] : request.query.jobId;
  if (!validJobId(jobId)) return response.status(404).json({ detail: "Không tìm thấy lượt chuyển đổi." });
  const job = await getJob(jobId);
  if (!job) return response.status(404).json({ detail: "Lượt chuyển đổi đã hết hạn hoặc không tồn tại." });
  /** @type {Record<string, unknown>} */
  const result = { ...job, files: undefined, options: undefined, output_pathname: undefined };
  if (job.status === "complete") {
    result.download_url = await createDownloadUrl(job.output_pathname);
    result.download_name = job.output_name;
  }
  response.setHeader("Cache-Control", "no-store");
  return response.status(200).json(result);
}
