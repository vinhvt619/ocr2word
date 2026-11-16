import { randomUUID } from "node:crypto";

import type { NextApiRequest, NextApiResponse } from "next";

import { errorMessage, normalizeOptions, requestBody } from "../../../lib/api";
import { createBrowserUploadUrl } from "../../../lib/blob";
import {
  ALLOWED_EXTENSIONS,
  EXCEL_EXTENSIONS,
  MAX_FILES,
  MAX_FILE_BYTES,
  MAX_TOTAL_BYTES,
  RESULT_TTL_MS,
  extensionOf,
  mimeTypeFor,
  outputDetails,
} from "../../../lib/config";
import { registerCleanup, saveJob } from "../../../lib/store";
import type { JobFile, JobRecord } from "../../../lib/types";

interface IncomingFile {
  name: string;
  size: number;
}

function validateFiles(value: unknown): IncomingFile[] {
  if (!Array.isArray(value) || value.length === 0) throw new Error("Hãy chọn ít nhất một file.");
  if (value.length > MAX_FILES) throw new Error(`Mỗi lượt chỉ nhận tối đa ${MAX_FILES} file.`);
  const files = value.map((item) => {
    if (!item || typeof item !== "object" || Array.isArray(item)) throw new Error("Danh sách file không hợp lệ.");
    const candidate = item as Record<string, unknown>;
    const name = typeof candidate.name === "string" ? candidate.name.normalize("NFC").trim() : "";
    const size = Number(candidate.size);
    const extension = extensionOf(name);
    if (!name || name.length > 255 || !ALLOWED_EXTENSIONS.has(extension)) {
      throw new Error(`'${name || "file"}' không thuộc định dạng được hỗ trợ.`);
    }
    if (!Number.isInteger(size) || size <= 0) throw new Error(`'${name}' là file rỗng hoặc có kích thước không hợp lệ.`);
    if (size > MAX_FILE_BYTES) throw new Error(`'${name}' lớn hơn giới hạn 150 MB.`);
    return { name, size };
  });
  const total = files.reduce((sum, file) => sum + file.size, 0);
  if (total > MAX_TOTAL_BYTES) throw new Error("Tổng dung lượng tải lên lớn hơn giới hạn 300 MB.");
  return files;
}

export default async function handler(request: NextApiRequest, response: NextApiResponse) {
  if (request.method !== "POST") {
    response.setHeader("Allow", "POST");
    return response.status(405).json({ detail: "Phương thức không được hỗ trợ." });
  }
  try {
    const body = requestBody(request);
    const incomingFiles = validateFiles(body.files);
    const options = normalizeOptions(body.options);
    if (options.output_format === "xlsx") {
      for (const file of incomingFiles) {
        if (EXCEL_EXTENSIONS.has(extensionOf(file.name))) {
          throw new Error("Chuyển đổi sang Excel (.xlsx) chỉ nhận file PDF hoặc hình ảnh (không nhận file Excel làm đầu vào).");
        }
      }
    }
    const jobId = randomUUID().replaceAll("-", "");
    const files: JobFile[] = incomingFiles.map((file, index) => {
      const extension = extensionOf(file.name);
      return {
        name: file.name,
        size: file.size,
        content_type: mimeTypeFor(file.name),
        pathname: `ocr2word/${jobId}/input/${String(index + 1).padStart(2, "0")}.${extension}`,
      };
    });
    const uploads = await Promise.all(files.map(async (file, index) => ({
      index,
      content_type: file.content_type,
      url: await createBrowserUploadUrl(file.pathname, file.content_type, file.size),
    })));
    const output = outputDetails(files, options.output_format);
    const now = new Date();
    const expiresAt = new Date(now.getTime() + RESULT_TTL_MS);
    const outputPathname = `ocr2word/${jobId}/output/result.${output.extension}`;
    const job: JobRecord = {
      job_id: jobId,
      status: "uploading",
      percent: 0,
      message: "Đang tải tài liệu trực tiếp lên Vercel Blob…",
      current_file: null,
      completed_files: 0,
      total_files: files.length,
      created_at: now.toISOString(),
      updated_at: now.toISOString(),
      expires_at: expiresAt.toISOString(),
      files,
      options,
      output_pathname: outputPathname,
      output_name: output.name,
      output_content_type: output.contentType,
      report: [],
      error: null,
    };
    await Promise.all([
      saveJob(job),
      registerCleanup(jobId, {
        expires_at_ms: expiresAt.getTime(),
        pathnames: [...files.map((file) => file.pathname), outputPathname],
      }),
    ]);
    response.setHeader("Cache-Control", "no-store");
    return response.status(201).json({ job_id: jobId, uploads });
  } catch (error) {
    return response.status(422).json({ detail: errorMessage(error, "Không thể chuẩn bị lượt tải lên.") });
  }
}
