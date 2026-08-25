import type { NextApiRequest, NextApiResponse } from "next";

import { ALLOWED_EXTENSIONS, MAX_FILES, MAX_FILE_BYTES, MAX_TOTAL_BYTES } from "../../lib/config";
import { configuredQueueRegion } from "../../lib/queue-client";

const unavailableOcr = {
  available: false,
  installed: false,
  command: null,
  version: null,
  languages: [],
  message: "Không kết nối được container OCR trên Vercel.",
};

const unavailableAi = {
  available: false,
  installed: false,
  server_running: false,
  model_available: false,
  model: process.env.OCR2WORD_AI_MODEL || "gemini-3.7-flash",
  provider: "Google Gemini API",
  endpoint: "https://generativelanguage.googleapis.com/v1beta/openai",
  requires_api_key: true,
  message: "Chưa có Gemini API key. Hãy chạy .\\setup-gemini.ps1 một lần.",
};

export default async function handler(request: NextApiRequest, response: NextApiResponse) {
  if (request.method !== "GET") {
    response.setHeader("Allow", "GET");
    return response.status(405).json({ detail: "Phương thức không được hỗ trợ." });
  }

  const configuration = {
    blob: Boolean(process.env.BLOB_READ_WRITE_TOKEN || (process.env.BLOB_STORE_ID && process.env.VERCEL_OIDC_TOKEN)),
    redis: Boolean(
      (process.env.UPSTASH_REDIS_REST_URL && process.env.UPSTASH_REDIS_REST_TOKEN)
      || (process.env.KV_REST_API_URL && process.env.KV_REST_API_TOKEN),
    ),
    secret: Boolean(process.env.OCR2WORD_INTERNAL_SECRET),
    worker: Boolean(process.env.OCR2WORD_WORKER_URL),
  };
  let workerStatus: Record<string, unknown> = {};
  if (configuration.worker) {
    try {
      const workerUrl = new URL("/api/status", `${process.env.OCR2WORD_WORKER_URL!.replace(/\/$/, "")}/`);
      const workerResponse = await fetch(workerUrl, { cache: "no-store", signal: AbortSignal.timeout(5000) });
      if (workerResponse.ok) workerStatus = await workerResponse.json() as Record<string, unknown>;
    } catch {
      // Return a useful status payload even while the worker is cold or unavailable.
    }
  }
  const workerOcr = workerStatus.ocr && typeof workerStatus.ocr === "object"
    ? workerStatus.ocr as { available?: boolean }
    : null;
  const ready = Object.values(configuration).every(Boolean) && workerOcr?.available === true;
  const missing = Object.entries(configuration).filter(([, value]) => !value).map(([key]) => key);

  return response.status(200).json({
    app: "OCR2Word",
    version: "2.0.0",
    ocr: workerStatus.ocr ?? unavailableOcr,
    ai: workerStatus.ai ?? unavailableAi,
    formats: Array.from(ALLOWED_EXTENSIONS).map((extension) => `.${extension}`).sort(),
    limits: {
      max_files: MAX_FILES,
      max_file_mb: MAX_FILE_BYTES / (1024 * 1024),
      max_total_mb: MAX_TOTAL_BYTES / (1024 * 1024),
    },
    deployment: {
      mode: "vercel",
      available: ready,
      direct_upload: true,
      async_jobs: true,
      local_parity: process.env.OCR2WORD_LOCAL_PARITY === "1",
      message: ready
        ? "Vercel Blob, Queue, Redis và container OCR đã sẵn sàng."
        : `Thiếu cấu hình Vercel: ${missing.join(", ") || "container OCR chưa phản hồi"}.`,
    },
    runtime: {
      controller: {
        node: process.versions.node,
        environment: process.env.VERCEL_ENV || process.env.NODE_ENV || "unknown",
        region: process.env.VERCEL_REGION || "dev1",
      },
      queue: {
        region: configuredQueueRegion(),
        local_callback: process.env.NODE_ENV === "development",
      },
      worker: workerStatus.runtime ?? null,
    },
  });
}
