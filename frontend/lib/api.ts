import { timingSafeEqual } from "node:crypto";

import type { NextApiRequest } from "next";

import { internalSecret } from "./config";
import type { JobOptions } from "./types";

const JOB_ID_PATTERN = /^[a-f0-9]{32}$/;

export function validJobId(value: unknown): value is string {
  return typeof value === "string" && JOB_ID_PATTERN.test(value);
}

export function requestBody(request: NextApiRequest): Record<string, unknown> {
  if (request.body && typeof request.body === "object" && !Array.isArray(request.body)) {
    return request.body as Record<string, unknown>;
  }
  if (typeof request.body === "string") {
    try {
      const parsed = JSON.parse(request.body) as unknown;
      if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
        return parsed as Record<string, unknown>;
      }
    } catch {
      // The caller receives a normal validation error below.
    }
  }
  return {};
}

export function normalizeOptions(value: unknown): JobOptions {
  const source = value && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown>
    : {};
  const language = typeof source.language === "string" ? source.language : "vie";
  const ocrMode = source.ocr_mode;
  const dpi = Number(source.dpi ?? 300);
  if (!/^[a-z0-9_+.-]{2,80}$/i.test(language)) throw new Error("Ngôn ngữ OCR không hợp lệ.");
  if (!["auto", "force", "text"].includes(String(ocrMode ?? "auto"))) {
    throw new Error("Chế độ OCR không hợp lệ.");
  }
  if (!Number.isInteger(dpi) || dpi < 150 || dpi > 600) throw new Error("DPI phải từ 150 đến 600.");
  const booleanValue = (key: string, fallback: boolean) => (
    typeof source[key] === "boolean" ? source[key] as boolean : fallback
  );
  const rawOutputFormat = source.output_format ? String(source.output_format).toLowerCase() : "docx";
  if (!["docx", "xlsx"].includes(rawOutputFormat)) {
    throw new Error("Định dạng xuất không hợp lệ (chỉ hỗ trợ docx hoặc xlsx).");
  }
  const outputFormat = rawOutputFormat as "docx" | "xlsx";
  const mathOcr = booleanValue("math_ocr", false);
  const rawMathEngine = source.math_engine ? String(source.math_engine) : undefined;
  if (rawMathEngine && !["rules", "local_hybrid"].includes(rawMathEngine)) {
    throw new Error("Công cụ nhận diện toán không hợp lệ.");
  }

  if (outputFormat === "xlsx") {
    if (mathOcr) {
      throw new Error("Chuyển đổi sang Excel (.xlsx) không hỗ trợ tính năng nhận diện công thức toán.");
    }
  }

  return {
    output_format: outputFormat,
    language,
    ocr_mode: String(ocrMode ?? "auto") as JobOptions["ocr_mode"],
    dpi,
    page_breaks: booleanValue("page_breaks", true),
    preserve_images: booleanValue("preserve_images", true),
    preserve_layout: booleanValue("preserve_layout", true),
    detect_tables: booleanValue("detect_tables", true),
    math_ocr: mathOcr,
    ...(rawMathEngine ? { math_engine: rawMathEngine as "rules" | "local_hybrid" } : {}),
  };
}

function secureEqual(left: string, right: string): boolean {
  const leftBuffer = Buffer.from(left);
  const rightBuffer = Buffer.from(right);
  return leftBuffer.length === rightBuffer.length && timingSafeEqual(leftBuffer, rightBuffer);
}

export function internalRequestAuthorized(request: NextApiRequest): boolean {
  const authorization = request.headers.authorization ?? "";
  return secureEqual(authorization, `Bearer ${internalSecret()}`);
}

export function cronRequestAuthorized(request: NextApiRequest): boolean {
  const secret = process.env.CRON_SECRET?.trim();
  if (!secret) return !process.env.VERCEL;
  return secureEqual(request.headers.authorization ?? "", `Bearer ${secret}`);
}

export function errorMessage(error: unknown, fallback = "Đã xảy ra lỗi không mong đợi."): string {
  return error instanceof Error && error.message ? error.message : fallback;
}
