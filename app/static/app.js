const state = {
  files: [],
  status: null,
  resultUrl: null,
  resultName: null,
  resultJobId: null,
  resultIsObjectUrl: false,
  busy: false,
  progressValue: 0,
};

const elements = {
  fileInput: document.querySelector("#fileInput"),
  dropzone: document.querySelector("#dropzone"),
  filePanel: document.querySelector("#filePanel"),
  fileList: document.querySelector("#fileList"),
  fileCount: document.querySelector("#fileCount"),
  clearFiles: document.querySelector("#clearFiles"),
  addMore: document.querySelector("#addMore"),
  convertButton: document.querySelector("#convertButton"),
  progressPanel: document.querySelector("#progressPanel"),
  progressBar: document.querySelector("#progressBar"),
  progressPercent: document.querySelector("#progressPercent"),
  progressStage: document.querySelector("#progressStage"),
  progressDetail: document.querySelector("#progressDetail"),
  inlineNotice: document.querySelector("#inlineNotice"),
  resultPanel: document.querySelector("#resultPanel"),
  resultTitle: document.querySelector("#resultTitle"),
  resultMessage: document.querySelector("#resultMessage"),
  resultDetails: document.querySelector("#resultDetails"),
  downloadAgain: document.querySelector("#downloadAgain"),
  engineStatus: document.querySelector("#engineStatus"),
  ocrInfo: document.querySelector("#ocrInfo"),
  forceOcr: document.querySelector("#forceOcr"),
  language: document.querySelector("#language"),
  dpi: document.querySelector("#dpi"),
  preserveImages: document.querySelector("#preserveImages"),
  preserveLayout: document.querySelector("#preserveLayout"),
  detectTables: document.querySelector("#detectTables"),
  pageBreaks: document.querySelector("#pageBreaks"),
  aiAssist: document.querySelector("#aiAssist"),
  aiConfig: document.querySelector("#aiConfig"),
  aiInfo: document.querySelector("#aiInfo"),
  aiSetupHint: document.querySelector("#aiSetupHint"),
  aiModelBadge: document.querySelector("#aiModelBadge"),
  aiSwitchDescription: document.querySelector("#aiSwitchDescription"),
  aiProviderTitle: document.querySelector("#aiProviderTitle"),
  aiProviderDescription: document.querySelector("#aiProviderDescription"),
  aiSetupCopy: document.querySelector("#aiSetupCopy"),
  aiSetupCommand: document.querySelector("#aiSetupCommand"),
  aiPrivacy: document.querySelector("#aiPrivacy"),
  privacyBadge: document.querySelector("#privacyBadge"),
  privacyFooter: document.querySelector("#privacyFooter"),
};

const imageExtensions = new Set(["png", "jpg", "jpeg", "tif", "tiff", "bmp", "webp"]);
const allowedExtensions = new Set(["pdf", ...imageExtensions, "xlsx", "xlsm", "xls"]);
const previewUrls = new Map();

function extensionOf(filename) {
  return filename.includes(".") ? filename.split(".").pop().toLowerCase() : "";
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  const units = ["KB", "MB", "GB"];
  let value = bytes / 1024;
  let unit = units[0];
  for (let index = 1; value >= 1024 && index < units.length; index += 1) {
    value /= 1024;
    unit = units[index];
  }
  return `${value >= 10 ? value.toFixed(1) : value.toFixed(2)} ${unit}`;
}

function fileKey(file) {
  return `${file.name}:${file.size}:${file.lastModified}`;
}

function previewUrlFor(file) {
  const key = fileKey(file);
  if (!previewUrls.has(key)) previewUrls.set(key, URL.createObjectURL(file));
  return previewUrls.get(key);
}

function releaseUnusedPreviewUrls() {
  const activeKeys = new Set(state.files.map(fileKey));
  previewUrls.forEach((url, key) => {
    if (activeKeys.has(key)) return;
    URL.revokeObjectURL(url);
    previewUrls.delete(key);
  });
}

function configurePdfPreviewLink(link, file, previewUrl) {
  link.href = previewUrl;
  link.target = "_blank";
  link.rel = "noopener";
  link.title = `Mở ${file.name} trong tab mới để đối chiếu`;
  link.setAttribute("aria-label", `Mở file PDF ${file.name} trong tab mới để đối chiếu`);
}

function showNotice(message = "") {
  elements.inlineNotice.textContent = message;
  elements.inlineNotice.hidden = !message;
}

function fileKind(extension) {
  if (extension === "pdf") return "pdf";
  if (imageExtensions.has(extension)) return "image";
  return "excel";
}

function hasOcrInput() {
  return state.files.some((file) => {
    const extension = extensionOf(file.name);
    return extension === "pdf" || imageExtensions.has(extension);
  });
}

function syncAiPanel() {
  const enabled = elements.aiAssist.checked;
  const aiStatus = state.status?.ai;
  const unavailable = Boolean(aiStatus && !aiStatus.available);
  elements.aiConfig.hidden = !enabled && !unavailable;
  if (enabled && unavailable) {
    elements.aiAssist.checked = false;
    elements.aiConfig.hidden = false;
    showNotice(aiStatus.message || "Tính năng AI chưa sẵn sàng.");
  }
}

function updateBusyControls() {
  const deploymentUnavailable = state.status?.deployment?.mode === "vercel"
    && state.status.deployment.available === false;
  elements.fileInput.disabled = state.busy;
  elements.clearFiles.disabled = state.busy;
  elements.addMore.disabled = state.busy;
  elements.convertButton.disabled = state.busy || state.files.length === 0 || deploymentUnavailable;
  elements.aiAssist.disabled = state.busy || state.status?.ai?.available === false;
  elements.dropzone.setAttribute("aria-disabled", String(state.busy));
  elements.fileList.querySelectorAll(".remove-file").forEach((button) => {
    button.disabled = state.busy;
  });
}

function renderFiles() {
  releaseUnusedPreviewUrls();
  elements.fileList.replaceChildren();
  state.files.forEach((file) => {
    const extension = extensionOf(file.name);
    const kind = fileKind(extension);
    const previewUrl = kind === "pdf" ? previewUrlFor(file) : null;
    const item = document.createElement("li");
    item.className = "file-item";

    const type = document.createElement(previewUrl ? "a" : "span");
    type.className = `file-type ${kind === "pdf" ? "" : kind}`;
    type.textContent = extension.toUpperCase();
    if (previewUrl) {
      type.classList.add("file-preview-link");
      configurePdfPreviewLink(type, file, previewUrl);
    }

    const copy = document.createElement("span");
    copy.className = "file-copy";
    const name = document.createElement(previewUrl ? "a" : "strong");
    name.textContent = file.name;
    name.title = file.name;
    if (previewUrl) {
      name.className = "file-name-preview";
      configurePdfPreviewLink(name, file, previewUrl);
    }
    const meta = document.createElement("span");
    const kindLabel = kind === "pdf" ? "Tài liệu PDF" : (kind === "image" ? "Hình ảnh OCR" : "Bảng tính Excel");
    meta.textContent = `${formatBytes(file.size)} • ${kindLabel}${previewUrl ? " • Nhấn để mở đối chiếu ↗" : ""}`;
    copy.append(name, meta);

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "remove-file";
    remove.setAttribute("aria-label", `Bỏ file ${file.name}`);
    remove.innerHTML = '<svg viewBox="0 0 20 20" aria-hidden="true"><path d="m6 6 8 8M14 6l-8 8" /></svg>';
    remove.addEventListener("click", () => {
      if (state.busy) return;
      state.files = state.files.filter((candidate) => fileKey(candidate) !== fileKey(file));
      resetProgress();
      renderFiles();
    });

    item.append(type, copy, remove);
    elements.fileList.append(item);
  });

  elements.fileCount.textContent = String(state.files.length);
  elements.filePanel.hidden = state.files.length === 0;
  elements.dropzone.hidden = state.files.length > 0;
  updateBusyControls();
  elements.resultPanel.hidden = true;
  if (state.status?.deployment?.mode === "vercel" && state.status.deployment.available === false) {
    showNotice(state.status.deployment.message || "Dự án Vercel chưa cấu hình xong.");
  } else if (hasOcrInput() && state.status && !state.status.ocr.available) {
    showNotice("Lưu ý: Tesseract chưa sẵn sàng. PDF có lớp chữ và Excel vẫn xử lý được; PDF scan hoặc hình ảnh cần cài OCR hay chạy bằng Docker.");
  } else {
    showNotice("");
  }
}

function addFiles(fileList) {
  if (state.busy) return;
  const existing = new Set(state.files.map(fileKey));
  const errors = [];
  const maxFiles = state.status?.limits?.max_files ?? 20;
  const maxFileBytes = (state.status?.limits?.max_file_mb ?? 150) * 1024 * 1024;
  const maxTotalBytes = (state.status?.limits?.max_total_mb ?? 300) * 1024 * 1024;

  for (const file of fileList) {
    const extension = extensionOf(file.name);
    if (!allowedExtensions.has(extension)) {
      errors.push(`${file.name}: định dạng không hỗ trợ`);
      continue;
    }
    if (file.size > maxFileBytes) {
      errors.push(`${file.name}: lớn hơn ${formatBytes(maxFileBytes)}`);
      continue;
    }
    if (existing.has(fileKey(file))) continue;
    if (state.files.length >= maxFiles) {
      errors.push(`Chỉ được chọn tối đa ${maxFiles} file`);
      break;
    }
    state.files.push(file);
    existing.add(fileKey(file));
  }

  const total = state.files.reduce((sum, file) => sum + file.size, 0);
  if (total > maxTotalBytes) {
    const addedKeys = new Set(Array.from(fileList, fileKey));
    state.files = state.files.filter((file) => !addedKeys.has(fileKey(file)));
    errors.push(`Tổng dung lượng vượt ${formatBytes(maxTotalBytes)}`);
  }
  resetProgress();
  renderFiles();
  if (errors.length) showNotice(errors.join("\n"));
  elements.fileInput.value = "";
}

function decodeReport(value) {
  if (!value) return [];
  try {
    const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
    const binary = atob(normalized);
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    return JSON.parse(new TextDecoder("utf-8").decode(bytes));
  } catch (_error) {
    return [];
  }
}

function filenameFromHeaders(getHeader) {
  const disposition = getHeader("content-disposition") || "";
  const utf8 = disposition.match(/filename\*=utf-8''([^;]+)/i);
  if (utf8) {
    try { return decodeURIComponent(utf8[1]); } catch (_error) { /* continue */ }
  }
  const basic = disposition.match(/filename="?([^";]+)"?/i);
  return basic ? basic[1] : (state.files.length > 1 ? "ocr2word-ket-qua.zip" : "tai-lieu-editable.docx");
}

function downloadBlob(blob, filename) {
  if (state.resultUrl && state.resultIsObjectUrl) URL.revokeObjectURL(state.resultUrl);
  state.resultUrl = URL.createObjectURL(blob);
  state.resultName = filename;
  state.resultJobId = null;
  state.resultIsObjectUrl = true;
  const link = document.createElement("a");
  link.href = state.resultUrl;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
}

function downloadRemote(url, filename, jobId) {
  if (state.resultUrl && state.resultIsObjectUrl) URL.revokeObjectURL(state.resultUrl);
  state.resultUrl = url;
  state.resultName = filename;
  state.resultJobId = jobId;
  state.resultIsObjectUrl = false;
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.rel = "noopener";
  document.body.append(link);
  link.click();
  link.remove();
}

function renderReport(report, filename) {
  const successful = report.filter((item) => item.status === "success");
  const failed = report.filter((item) => item.status === "error");
  const warnings = successful.flatMap((item) => item.warnings || []);
  const detectedTables = successful.reduce((sum, item) => sum + (item.tables || 0), 0);
  const aiProcessed = successful.reduce((sum, item) => sum + (item.ai?.processed_paragraphs || 0), 0);
  const aiChanged = successful.reduce((sum, item) => sum + (item.ai?.changed_paragraphs || 0), 0);
  elements.resultTitle.textContent = failed.length ? "Đã hoàn tất một phần" : "Đã chuyển đổi xong";
  elements.resultMessage.textContent = `${filename} đã được tải xuống.`;
  elements.resultDetails.replaceChildren();
  if (successful.length > 1) {
    const detail = document.createElement("li");
    detail.textContent = `${successful.length} file Word đã được đóng gói trong ZIP.`;
    elements.resultDetails.append(detail);
  }
  if (detectedTables > 0) {
    const detail = document.createElement("li");
    detail.textContent = `${detectedTables} bảng trong ảnh/PDF scan đã được tạo thành bảng Word có thể chỉnh sửa.`;
    elements.resultDetails.append(detail);
  }
  if (aiProcessed > 0) {
    const detail = document.createElement("li");
    detail.textContent = `AI đã kiểm tra ${aiProcessed} đoạn và sửa ${aiChanged} đoạn chữ; bảng và kiểu chữ được giữ nguyên.`;
    elements.resultDetails.append(detail);
  }
  for (const warning of warnings.slice(0, 3)) {
    const detail = document.createElement("li");
    detail.textContent = warning;
    elements.resultDetails.append(detail);
  }
  for (const item of failed.slice(0, 3)) {
    const detail = document.createElement("li");
    detail.textContent = `${item.file}: ${item.message}`;
    elements.resultDetails.append(detail);
  }
  elements.resultPanel.hidden = false;
}

function resetProgress() {
  state.progressValue = 0;
  elements.progressPanel.hidden = true;
  elements.progressPanel.classList.remove("is-error");
  elements.progressBar.setAttribute("aria-valuenow", "0");
  elements.progressBar.querySelector("span").style.width = "0%";
  elements.progressPercent.textContent = "0%";
}

function setProgress(value, stage, detail = "", { error = false } = {}) {
  const requested = Number.isFinite(Number(value)) ? Math.max(0, Math.min(100, Math.round(Number(value)))) : 0;
  state.progressValue = Math.max(state.progressValue, requested);
  elements.progressPanel.hidden = false;
  elements.progressPanel.classList.toggle("is-error", error);
  elements.progressStage.textContent = stage || "Đang xử lý tài liệu…";
  const deploymentMode = state.status?.deployment?.mode;
  const defaultDetail = deploymentMode === "vercel"
    ? "OCR2Word đang xử lý trong container Vercel."
    : (deploymentMode === "modal"
      ? "OCR2Word đang xử lý trong container Modal theo yêu cầu."
      : "OCR2Word đang xử lý trên máy của bạn.");
  elements.progressDetail.textContent = detail || defaultDetail;
  elements.progressPercent.textContent = `${state.progressValue}%`;
  elements.progressBar.setAttribute("aria-valuenow", String(state.progressValue));
  elements.progressBar.querySelector("span").style.width = `${state.progressValue}%`;
}

function createJobId() {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID().replace(/-/g, "");
  return `job${Date.now().toString(36)}${Math.random().toString(36).slice(2, 12)}`;
}

function progressDetail(payload) {
  const details = [];
  if (payload.current_file) details.push(payload.current_file);
  if (payload.total_files > 1) {
    details.push(`${payload.completed_files || 0}/${payload.total_files} file đã xử lý`);
  }
  const deploymentMode = state.status?.deployment?.mode;
  const fallback = deploymentMode === "vercel"
    ? "OCR2Word đang xử lý trong container Vercel."
    : (deploymentMode === "modal"
      ? "OCR2Word đang xử lý trong container Modal theo yêu cầu."
      : "OCR2Word đang xử lý trên máy của bạn.");
  return details.join(" • ") || fallback;
}

function startProgressPolling(jobId) {
  let active = true;
  let timer = null;

  const tick = async () => {
    try {
      const response = await fetch(`/api/progress/${encodeURIComponent(jobId)}`, { cache: "no-store" });
      if (active && response.ok) {
        const payload = await response.json();
        const displayedPercent = payload.status === "complete" ? Math.min(99, payload.percent) : payload.percent;
        setProgress(
          displayedPercent,
          payload.message || "Đang chuyển đổi tài liệu…",
          progressDetail(payload),
          { error: payload.status === "error" },
        );
        if (["complete", "error"].includes(payload.status)) return;
      }
    } catch (_error) {
      // Upload/download continues even if an individual progress poll fails.
    }
    if (active) timer = window.setTimeout(tick, 450);
  };

  tick();
  return () => {
    active = false;
    if (timer !== null) window.clearTimeout(timer);
  };
}

function sendConversion(form) {
  return new Promise((resolve, reject) => {
    const request = new XMLHttpRequest();
    request.open("POST", "/api/convert");
    request.responseType = "blob";
    request.upload.addEventListener("progress", (event) => {
      if (!event.lengthComputable) return;
      const percent = Math.round((event.loaded / event.total) * 12);
      setProgress(percent, "Đang tải tài liệu lên…", `${formatBytes(event.loaded)} / ${formatBytes(event.total)}`);
    });
    request.addEventListener("load", async () => {
      if (request.status >= 200 && request.status < 300) {
        resolve({
          blob: request.response,
          getHeader: (name) => request.getResponseHeader(name),
        });
        return;
      }
      let message = `Không thể chuyển đổi (HTTP ${request.status}).`;
      try {
        const payload = JSON.parse(await request.response.text());
        if (typeof payload.detail === "string") message = payload.detail;
        else if (payload.detail) message = JSON.stringify(payload.detail);
      } catch (_error) { /* use fallback */ }
      reject(new Error(message));
    });
    request.addEventListener("error", () => reject(new Error("Mất kết nối khi chuyển đổi tài liệu.")));
    request.addEventListener("abort", () => reject(new Error("Yêu cầu chuyển đổi đã bị hủy.")));
    request.send(form);
  });
}

function conversionOptions(useAi) {
  return {
    language: elements.language.value,
    ocr_mode: document.querySelector('input[name="ocrMode"]:checked').value,
    dpi: Number(elements.dpi.value),
    preserve_images: elements.preserveImages.checked,
    preserve_layout: elements.preserveLayout.checked,
    detect_tables: elements.detectTables.checked,
    page_breaks: elements.pageBreaks.checked,
    ai_assist: useAi,
  };
}

async function responseError(response, fallback) {
  try {
    const payload = await response.json();
    if (typeof payload.detail === "string") return payload.detail;
  } catch (_error) { /* use fallback */ }
  return fallback;
}

function uploadToVercelBlob(file, upload, completedBytes, totalBytes) {
  return new Promise((resolve, reject) => {
    const request = new XMLHttpRequest();
    request.open("PUT", upload.url);
    request.setRequestHeader("Content-Type", upload.content_type);
    request.upload.addEventListener("progress", (event) => {
      const loaded = event.lengthComputable ? event.loaded : 0;
      const overallBytes = Math.min(totalBytes, completedBytes + loaded);
      const percent = totalBytes ? Math.round((overallBytes / totalBytes) * 20) : 0;
      setProgress(
        percent,
        "Đang tải trực tiếp lên Vercel Blob…",
        `${file.name} • ${formatBytes(overallBytes)} / ${formatBytes(totalBytes)}`,
      );
    });
    request.addEventListener("load", () => {
      if (request.status >= 200 && request.status < 300) resolve();
      else reject(new Error(`Vercel Blob không nhận được '${file.name}' (HTTP ${request.status}).`));
    });
    request.addEventListener("error", () => reject(new Error(`Mất kết nối khi tải '${file.name}' lên Vercel Blob.`)));
    request.addEventListener("abort", () => reject(new Error(`Đã hủy tải '${file.name}'.`)));
    request.send(file);
  });
}

async function waitForVercelJob(jobId) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < 35 * 60 * 1000) {
    const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}`, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(await responseError(response, "Không đọc được trạng thái chuyển đổi trên Vercel."));
    }
    const payload = await response.json();
    setProgress(
      payload.status === "complete" ? 100 : payload.percent,
      payload.message || "Đang chuyển đổi tài liệu…",
      progressDetail(payload),
      { error: payload.status === "error" },
    );
    if (payload.status === "complete") return payload;
    if (payload.status === "error") throw new Error(payload.error || payload.message || "Chuyển đổi không thành công.");
    await new Promise((resolve) => window.setTimeout(resolve, 800));
  }
  throw new Error("Lượt chuyển đổi vượt quá thời gian chờ 35 phút.");
}

async function sendVercelConversion(useAi) {
  const createResponse = await fetch("/api/jobs", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      files: state.files.map((file) => ({ name: file.name, size: file.size })),
      options: conversionOptions(useAi),
    }),
  });
  if (!createResponse.ok) {
    throw new Error(await responseError(createResponse, "Không thể tạo lượt chuyển đổi trên Vercel."));
  }
  const created = await createResponse.json();
  const jobId = created.job_id;
  state.resultJobId = jobId;
  const totalBytes = state.files.reduce((sum, file) => sum + file.size, 0);
  let completedBytes = 0;
  for (const upload of created.uploads) {
    const file = state.files[upload.index];
    if (!file) throw new Error("Vercel trả về danh sách tải lên không hợp lệ.");
    await uploadToVercelBlob(file, upload, completedBytes, totalBytes);
    completedBytes += file.size;
  }
  setProgress(20, "Đã tải tài liệu lên Blob", "Đang đưa công việc vào Vercel Queue…");
  const startResponse = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/start`, { method: "POST" });
  if (!startResponse.ok) {
    throw new Error(await responseError(startResponse, "Không thể bắt đầu container OCR."));
  }
  return waitForVercelJob(jobId);
}

function createModalJob(form) {
  return new Promise((resolve, reject) => {
    const request = new XMLHttpRequest();
    request.open("POST", "/api/modal/jobs");
    request.upload.addEventListener("progress", (event) => {
      if (!event.lengthComputable) return;
      const percent = Math.round((event.loaded / event.total) * 20);
      setProgress(
        percent,
        "Đang tải tài liệu lên Modal…",
        `${formatBytes(event.loaded)} / ${formatBytes(event.total)}`,
      );
    });
    request.addEventListener("load", () => {
      let payload = null;
      try { payload = JSON.parse(request.responseText); } catch (_error) { /* handled below */ }
      if (request.status >= 200 && request.status < 300 && payload?.job_id) {
        resolve(payload);
        return;
      }
      const detail = typeof payload?.detail === "string" ? payload.detail : null;
      reject(new Error(detail || `Không thể tạo lượt chuyển đổi trên Modal (HTTP ${request.status}).`));
    });
    request.addEventListener("error", () => reject(new Error("Mất kết nối khi tải tài liệu lên Modal.")));
    request.addEventListener("abort", () => reject(new Error("Đã hủy tải tài liệu lên Modal.")));
    request.send(form);
  });
}

async function waitForModalJob(jobId) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < 40 * 60 * 1000) {
    const response = await fetch(`/api/modal/jobs/${encodeURIComponent(jobId)}`, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(await responseError(response, "Không đọc được trạng thái chuyển đổi trên Modal."));
    }
    const payload = await response.json();
    setProgress(
      payload.status === "complete" ? 100 : payload.percent,
      payload.message || "Đang chuyển đổi tài liệu…",
      progressDetail(payload),
      { error: payload.status === "error" },
    );
    if (payload.status === "complete") return payload;
    if (payload.status === "error") {
      throw new Error(payload.error || payload.message || "Chuyển đổi không thành công.");
    }
    await new Promise((resolve) => window.setTimeout(resolve, 800));
  }
  throw new Error("Lượt chuyển đổi vượt quá thời gian chờ 40 phút.");
}

async function sendModalConversion(useAi) {
  const form = new FormData();
  state.files.forEach((file) => form.append("files", file, file.name));
  form.append("options", JSON.stringify(conversionOptions(useAi)));
  const created = await createModalJob(form);
  state.resultJobId = created.job_id;
  setProgress(20, "Đã tải tài liệu lên Modal", "Đang khởi động container OCR…");
  return waitForModalJob(created.job_id);
}

async function convertFiles() {
  if (!state.files.length || state.busy) return;
  const useAi = elements.aiAssist.checked;
  if (useAi && !state.status?.ai?.available) {
    showNotice(state.status?.ai?.message || "AI chưa sẵn sàng.");
    elements.aiConfig.hidden = false;
    return;
  }
  if (state.status?.deployment?.mode === "vercel" && !state.status.deployment.available) {
    showNotice(state.status.deployment.message || "Dự án Vercel chưa cấu hình xong.");
    return;
  }
  state.busy = true;
  resetProgress();
  showNotice("");
  elements.resultPanel.hidden = true;
  elements.convertButton.classList.add("is-busy");
  updateBusyControls();
  setProgress(0, "Đang chuẩn bị…", `${state.files.length} tài liệu chờ xử lý`);

  let stopPolling = () => {};

  try {
    if (state.status?.deployment?.mode === "modal") {
      const result = await sendModalConversion(useAi);
      const filename = result.download_name || (state.files.length > 1 ? "ocr2word-ket-qua.zip" : "tai-lieu-editable.docx");
      setProgress(100, "Đã chuyển đổi xong", "Kết quả được giữ tối đa 7 ngày trên Modal.");
      downloadRemote(result.download_url, filename, result.job_id);
      renderReport(Array.isArray(result.report) ? result.report : [], filename);
    } else if (state.status?.deployment?.mode === "vercel") {
      const result = await sendVercelConversion(useAi);
      const filename = result.download_name || (state.files.length > 1 ? "ocr2word-ket-qua.zip" : "tai-lieu-editable.docx");
      setProgress(100, "Đã chuyển đổi xong", "Kết quả đã sẵn sàng để tải xuống.");
      downloadRemote(result.download_url, filename, result.job_id);
      renderReport(Array.isArray(result.report) ? result.report : [], filename);
    } else {
      const form = new FormData();
      state.files.forEach((file) => form.append("files", file, file.name));
      const options = conversionOptions(useAi);
      for (const [key, value] of Object.entries(options)) form.append(key, String(value));
      const jobId = createJobId();
      form.append("job_id", jobId);
      stopPolling = startProgressPolling(jobId);
      const response = await sendConversion(form);
      const report = decodeReport(response.getHeader("x-ocr2word-report"));
      const filename = filenameFromHeaders(response.getHeader);
      setProgress(100, "Đã chuyển đổi xong", "Kết quả đã sẵn sàng để tải xuống.");
      downloadBlob(response.blob, filename);
      renderReport(report, filename);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : "Không thể chuyển đổi tài liệu.";
    setProgress(state.progressValue, "Chuyển đổi chưa hoàn tất", message, { error: true });
    showNotice(message);
  } finally {
    stopPolling();
    state.busy = false;
    elements.convertButton.classList.remove("is-busy");
    updateBusyControls();
  }
}

async function loadStatus() {
  try {
    const response = await fetch("/api/status", { cache: "no-store" });
    if (!response.ok) throw new Error("status unavailable");
    state.status = await response.json();
    const { ocr, ai } = state.status;
    const onVercel = state.status?.deployment?.mode === "vercel";
    const onModal = state.status?.deployment?.mode === "modal";
    elements.engineStatus.classList.remove("is-checking", "is-warning");
    elements.ocrInfo.classList.remove("is-warning");
    if (ocr.available) {
      elements.engineStatus.querySelector("span:last-child").textContent = "OCR sẵn sàng";
      elements.ocrInfo.querySelector("p").textContent = `${ocr.message} Ngôn ngữ: ${ocr.languages.join(", ")}.`;
    } else {
      elements.engineStatus.classList.add("is-warning");
      elements.engineStatus.querySelector("span:last-child").textContent = "OCR chưa cài";
      elements.ocrInfo.classList.add("is-warning");
      elements.ocrInfo.querySelector("p").textContent = ocr.message;
    }
    if (ai) {
      elements.aiModelBadge.textContent = ai.model;
      elements.aiInfo.textContent = ai.message;
      elements.aiInfo.classList.toggle("is-warning", !ai.available);
      elements.aiSetupHint.hidden = ai.available;
      elements.aiSwitchDescription.textContent = "Google Gemini sửa dấu và từ bị dính/tách";
      elements.aiProviderTitle.textContent = ai.provider || "Google Gemini API";
      elements.aiProviderDescription.textContent = "Dùng Free Tier, chỉ gọi khi bạn bật AI";
      elements.aiSetupCopy.textContent = "Tạo API key tại Google AI Studio rồi chạy một lần:";
      elements.aiSetupCommand.textContent = onModal ? ".\\setup-modal.ps1" : ".\\setup-gemini.ps1";
      if (onVercel) {
        elements.aiPrivacy.textContent = "File được lưu tạm trong Vercel Blob riêng tư. Khi bật AI, chỉ phần chữ OCR được gửi tới Google Gemini để hiệu đính; file gốc tự xóa sau khi xử lý và kết quả hết hạn sau 24 giờ.";
        elements.privacyBadge.lastChild.textContent = " Blob riêng tư";
        elements.privacyFooter.innerHTML = "<strong>OCR2Word</strong> • File nguồn tự xóa sau chuyển đổi; kết quả trong Blob riêng tư hết hạn sau 24 giờ.";
      } else if (onModal) {
        elements.aiPrivacy.textContent = "File được truyền mã hóa vào tác vụ Modal và tự hết hạn tối đa sau 7 ngày. Khi bật AI, chỉ phần chữ OCR được gửi tới Google Gemini để hiệu đính.";
        elements.privacyBadge.lastChild.textContent = " Tác vụ riêng tư";
        elements.privacyFooter.innerHTML = "<strong>OCR2Word</strong> • Container chỉ bật khi sử dụng; file và kết quả Modal tự hết hạn tối đa sau 7 ngày.";
      } else {
        elements.aiPrivacy.textContent = "Khi bật AI, chỉ phần chữ OCR được gửi tới Google Gemini để hiệu đính; file tài liệu vẫn được xử lý trên máy của bạn.";
        elements.privacyBadge.lastChild.textContent = " Xử lý tài liệu trên máy";
        elements.privacyFooter.innerHTML = "<strong>OCR2Word</strong> • OCR chạy cục bộ; chỉ chữ OCR được gửi tới Gemini khi bạn bật AI.";
      }
      if (!ai.available) {
        elements.aiAssist.checked = false;
        elements.aiAssist.disabled = true;
      }
    }
    if (onVercel && state.status.deployment.available === false) {
      showNotice(state.status.deployment.message || "Dự án Vercel chưa cấu hình xong.");
    }
    syncAiPanel();
    renderFiles();
  } catch (_error) {
    elements.engineStatus.classList.remove("is-checking");
    elements.engineStatus.classList.add("is-warning");
    elements.engineStatus.querySelector("span:last-child").textContent = "Không đọc được trạng thái";
    elements.ocrInfo.classList.add("is-warning");
    elements.ocrInfo.querySelector("p").textContent = "Không thể kết nối đến dịch vụ chuyển đổi.";
    elements.aiInfo.textContent = "Không đọc được trạng thái AI.";
    elements.aiInfo.classList.add("is-warning");
  }
}

elements.dropzone.addEventListener("click", () => {
  if (!state.busy) elements.fileInput.click();
});
elements.dropzone.addEventListener("keydown", (event) => {
  if (!state.busy && (event.key === "Enter" || event.key === " ")) {
    event.preventDefault();
    elements.fileInput.click();
  }
});
elements.fileInput.addEventListener("change", () => addFiles(elements.fileInput.files));
elements.addMore.addEventListener("click", () => {
  if (!state.busy) elements.fileInput.click();
});
elements.clearFiles.addEventListener("click", () => {
  if (state.busy) return;
  state.files = [];
  resetProgress();
  renderFiles();
});
elements.convertButton.addEventListener("click", convertFiles);
elements.aiAssist.addEventListener("change", () => {
  showNotice("");
  syncAiPanel();
});
elements.downloadAgain.addEventListener("click", async () => {
  if (!state.resultName) return;
  if (state.resultJobId) {
    try {
      const jobPath = state.status?.deployment?.mode === "modal"
        ? `/api/modal/jobs/${encodeURIComponent(state.resultJobId)}`
        : `/api/jobs/${encodeURIComponent(state.resultJobId)}`;
      const response = await fetch(jobPath, { cache: "no-store" });
      if (!response.ok) throw new Error(await responseError(response, "Liên kết tải đã hết hạn."));
      const payload = await response.json();
      state.resultUrl = payload.download_url;
    } catch (error) {
      showNotice(error instanceof Error ? error.message : "Không thể tạo lại liên kết tải xuống.");
      return;
    }
  }
  if (!state.resultUrl) return;
  const link = document.createElement("a");
  link.href = state.resultUrl;
  link.download = state.resultName;
  document.body.append(link);
  link.click();
  link.remove();
});

for (const eventName of ["dragenter", "dragover"]) {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.add("is-dragging");
  });
}
for (const eventName of ["dragleave", "drop"]) {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.remove("is-dragging");
  });
}
elements.dropzone.addEventListener("drop", (event) => {
  if (!state.busy) addFiles(event.dataTransfer.files);
});

window.addEventListener("beforeunload", () => {
  if (state.resultUrl && state.resultIsObjectUrl) URL.revokeObjectURL(state.resultUrl);
  previewUrls.forEach((url) => URL.revokeObjectURL(url));
  previewUrls.clear();
});

loadStatus();
