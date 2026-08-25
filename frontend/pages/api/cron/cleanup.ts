import type { NextApiRequest, NextApiResponse } from "next";

import { cronRequestAuthorized } from "../../../lib/api";
import { deleteBlobs } from "../../../lib/blob";
import { getCleanup, listCleanupIds, removeCleanup } from "../../../lib/store";

export default async function handler(request: NextApiRequest, response: NextApiResponse) {
  if (request.method !== "GET") {
    response.setHeader("Allow", "GET");
    return response.status(405).json({ detail: "Phương thức không được hỗ trợ." });
  }
  if (!cronRequestAuthorized(request)) return response.status(401).json({ detail: "Không được phép." });
  const now = Date.now();
  let removed = 0;
  const ids = (await listCleanupIds()).slice(0, 100);
  for (const jobId of ids) {
    const cleanup = await getCleanup(jobId);
    if (cleanup && cleanup.expires_at_ms > now) continue;
    if (cleanup) {
      try {
        await deleteBlobs(cleanup.pathnames);
      } catch (error) {
        console.warn(`Could not clean blobs for job ${jobId}`, error);
        continue;
      }
    }
    await removeCleanup(jobId);
    removed += 1;
  }
  return response.status(200).json({ ok: true, scanned: ids.length, removed });
}
