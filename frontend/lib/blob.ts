import {
  del,
  issueSignedToken,
  presignUrl,
} from "@vercel/blob";

import {
  SIGNED_DOWNLOAD_TTL_MS,
  SIGNED_UPLOAD_TTL_MS,
  SIGNED_WORKER_TTL_MS,
} from "./config";

async function signedUrl(
  pathname: string,
  operation: "get" | "head" | "put",
  validForMs: number,
  options: { contentType?: string; maximumSizeInBytes?: number; allowOverwrite?: boolean } = {},
): Promise<string> {
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

export function createBrowserUploadUrl(
  pathname: string,
  contentType: string,
  maximumSizeInBytes: number,
): Promise<string> {
  return signedUrl(pathname, "put", SIGNED_UPLOAD_TTL_MS, { contentType, maximumSizeInBytes });
}

export function createWorkerReadUrl(pathname: string): Promise<string> {
  return signedUrl(pathname, "get", SIGNED_WORKER_TTL_MS);
}

export function createWorkerWriteUrl(
  pathname: string,
  contentType: string,
  maximumSizeInBytes: number,
): Promise<string> {
  return signedUrl(pathname, "put", SIGNED_WORKER_TTL_MS, {
    contentType,
    maximumSizeInBytes,
    allowOverwrite: true,
  });
}

export function createDownloadUrl(pathname: string): Promise<string> {
  return signedUrl(pathname, "get", SIGNED_DOWNLOAD_TTL_MS);
}

export async function deleteBlobs(pathnames: string[]): Promise<void> {
  if (!pathnames.length) return;
  await del(pathnames);
}
