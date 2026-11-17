import { timingSafeEqual } from "node:crypto";

import { internalSecret } from "./config.js";

/** @typedef {import("next").NextApiRequest} NextApiRequest */
/** @typedef {import("./types.js").JobOptions} JobOptions */

const JOB_ID_PATTERN = /^[a-f0-9]{32}$/;

/**
 * @param {unknown} value
 * @returns {value is string}
 */
export function validJobId(value) {
  return typeof value === "string" && JOB_ID_PATTERN.test(value);
}

/**
 * @param {NextApiRequest} request
 * @returns {Record<string, unknown>}
 */
export function requestBody(request) {
  if (request.body && typeof request.body === "object" && !Array.isArray(request.body)) {
    return /** @type {Record<string, unknown>} */ (request.body);
  }
  if (typeof request.body === "string") {
    try {
      const parsed = JSON.parse(request.body);
      if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
        return /** @type {Record<string, unknown>} */ (parsed);
      }
    } catch {
      // The caller receives a normal validation error below.
    }
  }
  return {};
}

/**
 * @param {unknown} value
 * @returns {JobOptions}
 */
export function normalizeOptions(value) {
  const source = value && typeof value === "object" && !Array.isArray(value)
    ? /** @type {Record<string, unknown>} */ (value)
    : {};
  const language = typeof source.language === "string" ? source.language : "vie";
  const ocrMode = source.ocr_mode;
  const dpi = Number(source.dpi ?? 300);
  if (!/^[a-z0-9_+.-]{2,80}$/i.test(language)) throw new Error("Ngôn ngữ OCR không hợp lệ.");
  if (!["auto", "force", "text"].includes(String(ocrMode ?? "auto"))) {
    throw new Error("Chế độ OCR không hợp lệ.");
  }
  if (!Number.isInteger(dpi) || dpi < 150 || dpi > 600) throw new Error("DPI phải từ 150 đến 600.");

  /**
   * @param {string} key
   * @param {boolean} fallback
   * @returns {boolean}
   */
  function booleanValue(key, fallback) {
    return typeof source[key] === "boolean" ? /** @type {boolean} */ (source[key]) : fallback;
  }

  const rawOutputFormat = source.output_format ? String(source.output_format).toLowerCase() : "docx";
  if (!["docx", "xlsx"].includes(rawOutputFormat)) {
    throw new Error("Định dạng xuất không hợp lệ (chỉ hỗ trợ docx hoặc xlsx).");
  }
  const outputFormat = /** @type {"docx" | "xlsx"} */ (rawOutputFormat);
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
    ocr_mode: /** @type {JobOptions["ocr_mode"]} */ (String(ocrMode ?? "auto")),
    dpi,
    page_breaks: booleanValue("page_breaks", true),
    preserve_images: booleanValue("preserve_images", true),
    preserve_layout: booleanValue("preserve_layout", true),
    detect_tables: booleanValue("detect_tables", true),
    math_ocr: mathOcr,
    ...(rawMathEngine ? { math_engine: /** @type {"rules" | "local_hybrid"} */ (rawMathEngine) } : {}),
  };
}

/**
 * @param {string} left
 * @param {string} right
 * @returns {boolean}
 */
function secureEqual(left, right) {
  const leftBuffer = Buffer.from(left);
  const rightBuffer = Buffer.from(right);
  return leftBuffer.length === rightBuffer.length && timingSafeEqual(leftBuffer, rightBuffer);
}

/**
 * @param {NextApiRequest} request
 * @returns {boolean}
 */
export function internalRequestAuthorized(request) {
  const authorization = request.headers.authorization ?? "";
  return secureEqual(authorization, `Bearer ${internalSecret()}`);
}

/**
 * @param {NextApiRequest} request
 * @returns {boolean}
 */
export function cronRequestAuthorized(request) {
  const secret = process.env.CRON_SECRET?.trim();
  if (!secret) return !process.env.VERCEL;
  return secureEqual(request.headers.authorization ?? "", `Bearer ${secret}`);
}

/**
 * @param {unknown} error
 * @param {string} [fallback]
 * @returns {string}
 */
export function errorMessage(error, fallback = "Đã xảy ra lỗi không mong đợi.") {
  return error instanceof Error && error.message ? error.message : fallback;
}
