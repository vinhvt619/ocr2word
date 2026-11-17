import {
  del,
  issueSignedToken,
  presignUrl,
} from "@vercel/blob";

import {
  SIGNED_DOWNLOAD_TTL_MS,
  SIGNED_UPLOAD_TTL_MS,
  SIGNED_WORKER_TTL_MS,
} from "./config.js";

/**
 * @param {string} pathname
 * @param {"get" | "head" | "put"} operation
 * @param {number} validForMs
 * @param {{ contentType?: string, maximumSizeInBytes?: number, allowOverwrite?: boolean }} [options]
 * @returns {Promise<string>}
 */
async function signedUrl(pathname, operation, validForMs, options = {}) {
  const validUntil = Date.now() + validForMs;
  const token = await issueSignedToken({
    pathname,
    operations: [operation],
    validUntil,
    allowedContentTypes: options.contentType ? [options.contentType] : undefined,
    maximumSizeInBytes: options.maximumSizeInBytes,
  });
  const result = operation === "put"
    ? await presignUrl(token, {
        access: "private",
        operation,
        pathname,
        validUntil,
        allowedContentTypes: options.contentType ? [options.contentType] : undefined,
        maximumSizeInBytes: options.maximumSizeInBytes,
        addRandomSuffix: false,
        allowOverwrite: options.allowOverwrite ?? false,
      })
    : await presignUrl(token, {
        access: "private",
        operation,
        pathname,
        validUntil,
      });
  return result.presignedUrl;
}

/**
 * @param {string} pathname
 * @param {string} contentType
 * @param {number} maximumSizeInBytes
 * @returns {Promise<string>}
 */
export function createBrowserUploadUrl(pathname, contentType, maximumSizeInBytes) {
  return signedUrl(pathname, "put", SIGNED_UPLOAD_TTL_MS, { contentType, maximumSizeInBytes });
}

/**
 * @param {string} pathname
 * @returns {Promise<string>}
 */
export function createWorkerReadUrl(pathname) {
  return signedUrl(pathname, "get", SIGNED_WORKER_TTL_MS);
}

/**
 * @param {string} pathname
 * @param {string} contentType
 * @param {number} maximumSizeInBytes
 * @returns {Promise<string>}
 */
export function createWorkerWriteUrl(pathname, contentType, maximumSizeInBytes) {
  return signedUrl(pathname, "put", SIGNED_WORKER_TTL_MS, {
    contentType,
    maximumSizeInBytes,
    allowOverwrite: true,
  });
}

/**
 * @param {string} pathname
 * @returns {Promise<string>}
 */
export function createDownloadUrl(pathname) {
  return signedUrl(pathname, "get", SIGNED_DOWNLOAD_TTL_MS);
}

/**
 * @param {string[]} pathnames
 * @returns {Promise<void>}
 */
export async function deleteBlobs(pathnames) {
  if (!pathnames.length) return;
  await del(pathnames);
}
