const state = {
  files: [],
  status: null,
  resultUrl: null,
  resultName: null,
  resultJobId: null,
  resultIsObjectUrl: false,
  busy: false,
  outputFormat: "docx",
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
  outputDocx: document.querySelector("#outputDocx"),
  outputXlsx: document.querySelector("#outputXlsx"),
  excelModeNotice: document.querySelector("#excelModeNotice"),
  language: document.querySelector("#language"),
  dpi: document.querySelector("#dpi"),
  preserveImages: document.querySelector("#preserveImages"),
  preserveLayout: document.querySelector("#preserveLayout"),
  detectTables: document.querySelector("#detectTables"),
  pageBreaks: document.querySelector("#pageBreaks"),
  mathOcr: document.querySelector("#mathOcr"),
  mathEngineRow: document.querySelector("#mathEngineRow"),
  mathEngine: document.querySelector("#mathEngine"),
  mathOcrHelp: document.querySelector("#mathOcrHelp"),
  privacyBadge: document.querySelector("#privacyBadge"),
  privacyFooter: document.querySelector("#privacyFooter"),
  savedIndicator: document.querySelector("#savedIndicator"),
  advancedActiveCount: document.querySelector("#advancedActiveCount"),
  convertMoreBtn: document.querySelector("#convertMoreBtn"),
  resetSettingsBtn: document.querySelector("#resetSettingsBtn"),
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

function selectedMathEngine() {
  return elements.mathEngine?.value || "local_hybrid";
}

function updateMathEngineHelp() {
  if (elements.mathEngineRow) {
    elements.mathEngineRow.hidden = !elements.mathOcr.checked;
  }
  if (!elements.mathOcrHelp) return;
  const engine = selectedMathEngine();
  let message = "Mặc định dùng quy tắc cấu trúc PDF kết hợp mô hình chạy trên máy chủ. Vùng công thức ngoài phạm vi hỗ trợ được giữ thành ảnh kèm cảnh báo; vui lòng đối chiếu lại ký hiệu sau khi chuyển đổi.";
  if (engine === "rules") {
    message = "Chỉ phân tích cấu trúc chữ, baseline và nét vẽ của PDF có lớp chữ số (không nạp model, không dùng API). Không hỗ trợ ảnh hoặc PDF scan.";
  } else {
    const localInfo = state.status?.math_ocr?.engines?.local_hybrid;
    if (localInfo && localInfo.available === false) {
      message += " (Lưu ý: Mô hình công thức cục bộ chưa sẵn sàng trên máy chủ này; với PDF số bạn có thể chọn 'Chỉ theo quy tắc PDF'.)";
    }
  }
  elements.mathOcrHelp.textContent = message;
}

function updateBusyControls() {
  const deploymentUnavailable = state.status?.deployment?.mode === "vercel"
    && state.status.deployment.available === false;
  const isXlsx = state.outputFormat === "xlsx";
  elements.fileInput.disabled = state.busy;
  elements.clearFiles.disabled = state.busy;
  elements.addMore.disabled = state.busy;
  if (elements.outputDocx) elements.outputDocx.disabled = state.busy;
  if (elements.outputXlsx) elements.outputXlsx.disabled = state.busy;
  elements.convertButton.disabled = state.busy || state.files.length === 0 || deploymentUnavailable;

  if (isXlsx) {
    elements.mathOcr.disabled = true;
    if (elements.mathEngine) elements.mathEngine.disabled = true;
    elements.preserveImages.disabled = true;
    elements.preserveLayout.disabled = true;
    elements.pageBreaks.disabled = true;
  } else {
    elements.mathOcr.disabled = state.busy;
    if (elements.mathEngine) elements.mathEngine.disabled = state.busy;
    elements.preserveImages.disabled = state.busy;
    elements.preserveLayout.disabled = state.busy;
    elements.pageBreaks.disabled = state.busy;
  }
  elements.dropzone.setAttribute("aria-disabled", String(state.busy));
  elements.fileList.querySelectorAll(".remove-file").forEach((button) => {
    button.disabled = state.busy;
  });
}

function updateOutputFormatUI() {
  const isXlsx = state.outputFormat === "xlsx";
  if (elements.outputDocx) {
    elements.outputDocx.checked = !isXlsx;
    elements.outputDocx.disabled = state.busy;
  }
  if (elements.outputXlsx) {
    elements.outputXlsx.checked = isXlsx;
    elements.outputXlsx.disabled = state.busy;
  }
  if (elements.excelModeNotice) elements.excelModeNotice.hidden = !isXlsx;

  if (isXlsx) {
    elements.mathOcr.checked = false;
    if (elements.mathEngineRow) elements.mathEngineRow.hidden = true;
    if (elements.downloadAgain) {
      elements.downloadAgain.innerHTML = '<svg viewBox="0 0 20 20" aria-hidden="true" width="16" height="16"><path d="M3 17a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-4a1 1 0 1 0-2 0v3H5v-3a1 1 0 1 0-2 0v4zm7-4a1 1 0 0 0 .707-.293l4-4a1 1 0 1 0-1.414-1.414L11 9.586V3a1 1 0 1 0-2 0v6.586L6.707 7.293a1 1 0 0 0-1.414 1.414l4 4A1 1 0 0 0 10 13z"/></svg> Tải lại file Excel (.xlsx)';
    }
    if (elements.resultMessage) {
      elements.resultMessage.textContent = "File Excel đã được tạo thành công và tự động tải về.";
    }
    // Revalidate pre-selected files when switching Word -> Excel
    const excelFiles = state.files.filter((f) => ["xlsx", "xlsm", "xls"].includes(extensionOf(f.name)));
    if (excelFiles.length > 0) {
      state.files = state.files.filter((f) => !["xlsx", "xlsm", "xls"].includes(extensionOf(f.name)));
      showNotice(`Đã loại bỏ ${excelFiles.length} file Excel vì định dạng xuất Excel chỉ nhận file PDF hoặc hình ảnh.`);
      renderFiles();
    }
  } else {
    if (elements.mathOcr.checked && elements.mathEngineRow) {
      elements.mathEngineRow.hidden = false;
    }
    if (elements.downloadAgain) {
      elements.downloadAgain.innerHTML = '<svg viewBox="0 0 20 20" aria-hidden="true" width="16" height="16"><path d="M3 17a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-4a1 1 0 1 0-2 0v3H5v-3a1 1 0 1 0-2 0v4zm7-4a1 1 0 0 0 .707-.293l4-4a1 1 0 1 0-1.414-1.414L11 9.586V3a1 1 0 1 0-2 0v6.586L6.707 7.293a1 1 0 0 0-1.414 1.414l4 4A1 1 0 0 0 10 13z"/></svg> Tải lại file Word (.docx)';
    }
    if (elements.resultMessage) {
      elements.resultMessage.textContent = "File Word đã được tạo thành công và tự động tải về.";
    }
  }

  // Update main action button text and progress text for selected format
  if (elements.convertButton) {
    const labelSpan = elements.convertButton.querySelector(".button-label > span:not(.kbd-hint)");
    if (labelSpan) {
      labelSpan.textContent = isXlsx ? "Bắt đầu chuyển sang Excel (.xlsx)" : "Bắt đầu chuyển sang Word (.docx)";
    }
    const progressSpan = elements.convertButton.querySelector(".button-progress");
    if (progressSpan) {
      progressSpan.innerHTML = isXlsx
        ? "<i></i> Đang nhận diện và tạo bảng tính Excel..."
        : "<i></i> Đang nhận diện và tạo file Word...";
    }
  }

  // Update text mode help description for selected format
  const textModeRadio = document.querySelector('input[name="ocrMode"][value="text"]');
  const textModeSmall = textModeRadio ? textModeRadio.closest(".radio-card")?.querySelector("small") : null;
  if (textModeSmall) {
    textModeSmall.textContent = isXlsx
      ? "Chỉ trích xuất lớp chữ có sẵn; không đọc vùng ảnh quét/scan"
      : "Giữ trang scan như ảnh; chỉ sửa được lớp chữ sẵn có";
  }
  updateBusyControls();
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
    if (state.outputFormat === "xlsx" && ["xlsx", "xlsm", "xls"].includes(extension)) {
      errors.push(`${file.name}: Chuyển đổi sang Excel (.xlsx) chỉ nhận PDF hoặc ảnh.`);
      continue;
    }
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
  const fallbackSingle = state.outputFormat === "xlsx" ? "tai-lieu-editable.xlsx" : "tai-lieu-editable.docx";
  return basic ? basic[1] : (state.files.length > 1 ? "ocr2word-ket-qua.zip" : fallbackSingle);
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
  const isXlsx = state.outputFormat === "xlsx";
  const successful = report.filter((item) => item.status === "success");
  const failed = report.filter((item) => item.status === "error");
  const warnings = successful.flatMap((item) => item.warnings || []);
  const detectedTables = successful.reduce((sum, item) => sum + (item.tables || 0), 0);
  const sheetCount = successful.reduce((sum, item) => sum + (item.sheets || 0), 0);
  const mathResults = successful.filter((item) => item.math);
  const equationCount = mathResults.reduce((sum, item) => sum + (item.math.equations || 0), 0);
  const fallbackPages = mathResults.reduce((sum, item) => sum + (item.math.fallback_pages || 0), 0);
  const fallbackRegions = mathResults.reduce((sum, item) => sum + (item.math.fallback_regions || 0), 0);
  const recognizedByRules = mathResults.reduce((sum, item) => sum + (item.math.recognized_by_rules || 0), 0);
  const recognizedByLocal = mathResults.reduce((sum, item) => sum + (item.math.recognized_by_local || 0), 0);
  const mathIncomplete = !isXlsx && mathResults.some((item) => ["partial", "unavailable"].includes(item.math.status));
  const noEquations = !isXlsx && mathIncomplete && equationCount === 0;
  const mathNotDetected = !isXlsx && mathResults.length > 0 && !mathIncomplete && equationCount === 0;
  const mathNotEnabled = !isXlsx && !mathResults.length && warnings.some((warning) => warning.includes("Phát hiện dấu hiệu ký hiệu hoặc font toán học"));

  const hasWarning = failed.length > 0 || (!isXlsx && (mathIncomplete || mathNotDetected || mathNotEnabled));
  elements.resultPanel.classList.toggle("is-error", noEquations || (failed.length > 0 && successful.length === 0));
  elements.resultPanel.classList.toggle("is-warning", !noEquations && hasWarning);
  elements.resultPanel.querySelector(".result-icon path").setAttribute("d", (noEquations || (!isXlsx && (mathIncomplete || mathNotDetected || mathNotEnabled)))
    ? "M12 7v6m0 4h.01M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20"
    : "m5 12 4 4L19 6");
  elements.resultTitle.textContent = noEquations ? "Chưa tạo được Equation"
    : mathIncomplete ? "Công thức mới chuyển đổi được một phần"
      : mathNotDetected ? "Không tìm thấy công thức để tạo Equation"
      : mathNotEnabled ? "Đã chuyển chữ/bố cục, chưa nhận diện công thức"
        : failed.length ? "Đã hoàn tất một phần" : "Đã chuyển đổi xong";
  elements.resultMessage.textContent = noEquations
    ? "File giữ nội dung nguồn đã tải xuống để đối chiếu; công thức trong file này chưa chỉnh sửa được."
    : `${filename} đã được tải xuống.`;
  if (noEquations) {
    showNotice("Nhận diện công thức chưa thành công với chế độ đã chọn. Xem chi tiết bên dưới.");
    setProgress(100, "Chưa tạo được Equation", "Xử lý đã kết thúc, nhưng chưa có công thức Word chỉnh sửa được.", { error: true });
  } else if (mathNotEnabled) {
    showNotice("Tài liệu có dấu hiệu công thức toán. Chế độ chuyển chữ/bố cục không tạo Equation; bật 'Nhận diện công thức toán' nếu cần sửa công thức.");
  }
  elements.resultDetails.replaceChildren();
  if (!isXlsx && mathResults.length) {
    const detail = document.createElement("li");
    const breakdown = [];
    if (recognizedByRules > 0) breakdown.push(`${recognizedByRules} từ quy tắc PDF`);
    if (recognizedByLocal > 0) breakdown.push(`${recognizedByLocal} từ mô hình cục bộ`);
    const breakdownText = breakdown.length ? ` (${breakdown.join(", ")})` : "";
    detail.textContent = `${equationCount} Equation chỉnh sửa được${breakdownText}; ${fallbackPages} trang và ${fallbackRegions} vùng còn giữ dạng ảnh.`;
    elements.resultDetails.append(detail);
  }
  if (successful.length > 1) {
    const detail = document.createElement("li");
    detail.textContent = `${successful.length} file ${isXlsx ? "Excel" : "Word"} đã được đóng gói trong ZIP.`;
    elements.resultDetails.append(detail);
  }
  if (detectedTables > 0 || (isXlsx && sheetCount > 0)) {
    const detail = document.createElement("li");
    if (isXlsx) {
      detail.textContent = sheetCount > 0
        ? `${sheetCount} sheet đã được tạo trong bảng tính Excel có thể chỉnh sửa (${detectedTables} bảng dữ liệu).`
        : `${detectedTables} bảng đã được trích xuất thành các sheet Excel có thể chỉnh sửa.`;
    } else {
      detail.textContent = `${detectedTables} bảng trong ảnh/PDF scan đã được tạo thành bảng Word có thể chỉnh sửa.`;
    }
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

function conversionOptions() {
  const isXlsx = state.outputFormat === "xlsx";
  const opts = {
    output_format: state.outputFormat || "docx",
    language: elements.language.value,
    ocr_mode: document.querySelector('input[name="ocrMode"]:checked').value,
    dpi: Number(elements.dpi.value),
    preserve_images: isXlsx ? false : elements.preserveImages.checked,
    preserve_layout: isXlsx ? false : elements.preserveLayout.checked,
    detect_tables: elements.detectTables.checked,
    page_breaks: isXlsx ? false : elements.pageBreaks.checked,
  };
  if (!isXlsx) {
    opts.math_ocr = elements.mathOcr.checked;
    opts.math_engine = selectedMathEngine();
  }
  return opts;
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

async function sendVercelConversion() {
  const createResponse = await fetch("/api/jobs", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      files: state.files.map((file) => ({ name: file.name, size: file.size })),
      options: conversionOptions(),
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

async function sendModalConversion() {
  const form = new FormData();
  state.files.forEach((file) => form.append("files", file, file.name));
  form.append("options", JSON.stringify(conversionOptions()));
  const created = await createModalJob(form);
  state.resultJobId = created.job_id;
  setProgress(20, "Đã tải tài liệu lên Modal", "Đang khởi động container OCR…");
  return waitForModalJob(created.job_id);
}

async function convertFiles() {
  if (!state.files.length || state.busy) return;
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
      const result = await sendModalConversion();
      const fallbackSingle = state.outputFormat === "xlsx" ? "tai-lieu-editable.xlsx" : "tai-lieu-editable.docx";
      const filename = result.download_name || (state.files.length > 1 ? "ocr2word-ket-qua.zip" : fallbackSingle);
      setProgress(100, "Đã chuyển đổi xong", "Kết quả được giữ tối đa 7 ngày trên Modal.");
      downloadRemote(result.download_url, filename, result.job_id);
      renderReport(Array.isArray(result.report) ? result.report : [], filename);
    } else if (state.status?.deployment?.mode === "vercel") {
      const result = await sendVercelConversion();
      const fallbackSingle = state.outputFormat === "xlsx" ? "tai-lieu-editable.xlsx" : "tai-lieu-editable.docx";
      const filename = result.download_name || (state.files.length > 1 ? "ocr2word-ket-qua.zip" : fallbackSingle);
      setProgress(100, "Đã chuyển đổi xong", "Kết quả đã sẵn sàng để tải xuống.");
      downloadRemote(result.download_url, filename, result.job_id);
      renderReport(Array.isArray(result.report) ? result.report : [], filename);
    } else {
      const form = new FormData();
      state.files.forEach((file) => form.append("files", file, file.name));
      const options = conversionOptions();
      for (const [key, value] of Object.entries(options)) {
        if (value !== undefined && value !== null) {
          form.append(key, String(value));
        }
      }
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
    const { ocr } = state.status;
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
    if (onVercel) {
      elements.privacyBadge.lastChild.textContent = " Blob riêng tư";
      elements.privacyFooter.innerHTML = "<strong>OCR2Word</strong> • File nguồn tự xóa sau chuyển đổi thành công. Kết quả hết hạn sau 24 giờ.";
    } else if (onModal) {
      elements.privacyBadge.lastChild.textContent = " Xử lý trên Modal";
      elements.privacyFooter.innerHTML = "<strong>OCR2Word</strong> • File và kết quả trong tác vụ Modal hết hạn sau tối đa 7 ngày.";
    }
    if (onVercel && state.status.deployment.available === false) {
      showNotice(state.status.deployment.message || "Dự án Vercel chưa cấu hình xong.");
    }
    updateMathEngineHelp();
    renderFiles();
  } catch (_error) {
    elements.engineStatus.classList.remove("is-checking");
    elements.engineStatus.classList.add("is-warning");
    elements.engineStatus.querySelector("span:last-child").textContent = "Không đọc được trạng thái";
    elements.ocrInfo.classList.add("is-warning");
    elements.ocrInfo.querySelector("p").textContent = "Không thể kết nối đến dịch vụ chuyển đổi.";
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
document.querySelectorAll('input[name="outputFormat"]').forEach((radio) => {
  radio.addEventListener("change", (e) => {
    if (state.busy) {
      updateOutputFormatUI();
      return;
    }
    state.outputFormat = e.target.value;
    updateOutputFormatUI();
    saveUserSettings();
  });
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
elements.mathOcr.addEventListener("change", () => {
  showNotice("");
  updateMathEngineHelp();
});
if (elements.mathEngine) {
  elements.mathEngine.addEventListener("change", () => {
    showNotice("");
    updateMathEngineHelp();
  });
}
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

// ==========================================
// Settings Persistence & UI Interaction
// ==========================================

const SETTINGS_STORAGE_KEY = "ocr2word_user_settings_v2";

let savedNoticeTimer = null;
function showSavedNotice() {
  const indicator = elements.savedIndicator || document.querySelector("#savedIndicator");
  if (!indicator) return;
  indicator.classList.add("is-visible");
  clearTimeout(savedNoticeTimer);
  savedNoticeTimer = setTimeout(() => {
    indicator.classList.remove("is-visible");
  }, 2200);
}

function updateAdvancedBadge() {
  const badge = elements.advancedActiveCount || document.querySelector("#advancedActiveCount");
  if (!badge) return;
  const items = [elements.preserveLayout, elements.preserveImages, elements.detectTables, elements.pageBreaks];
  const activeCount = items.filter((item) => item && item.checked).length;
  badge.textContent = `${activeCount}/4 đang bật`;
  badge.classList.toggle("is-partial", activeCount > 0 && activeCount < 4);
  badge.classList.toggle("is-none", activeCount === 0);
}

function updateActivePresetHighlight() {
  const currentOcr = document.querySelector('input[name="ocrMode"]:checked')?.value || "auto";
  const mathOn = Boolean(elements.mathOcr?.checked);

  document.querySelectorAll(".preset-pill[data-preset]").forEach((pill) => {
    const preset = pill.dataset.preset;
    let isActive = false;
    if (preset === "math") {
      isActive = mathOn && currentOcr === "auto";
    } else if (preset === "scan") {
      isActive = !mathOn && currentOcr === "force";
    } else if (preset === "tables") {
      isActive = !mathOn && currentOcr === "auto" && Boolean(elements.detectTables?.checked);
    }
    pill.classList.toggle("is-active", isActive);
  });
}

function saveUserSettings() {
  try {
    const selectedOcr = document.querySelector('input[name="ocrMode"]:checked');
    const data = {
      ocrMode: selectedOcr ? selectedOcr.value : "auto",
      outputFormat: state.outputFormat || "docx",
      language: elements.language?.value || "vie",
      dpi: elements.dpi?.value || "300",
      preserveImages: elements.preserveImages ? elements.preserveImages.checked : true,
      preserveLayout: elements.preserveLayout ? elements.preserveLayout.checked : true,
      detectTables: elements.detectTables ? elements.detectTables.checked : true,
      pageBreaks: elements.pageBreaks ? elements.pageBreaks.checked : true,
      mathOcr: elements.mathOcr ? elements.mathOcr.checked : false,
      mathEngine: selectedMathEngine(),
    };
    localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(data));
    showSavedNotice();
    updateAdvancedBadge();
    updateActivePresetHighlight();
  } catch (_error) {
    /* localStorage inaccessible */
  }
}

function restoreUserSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!raw) {
      updateAdvancedBadge();
      updateActivePresetHighlight();
      return;
    }
    const saved = JSON.parse(raw);
    if (saved.outputFormat) {
      state.outputFormat = saved.outputFormat;
    }
    if (saved.ocrMode) {
      const radio = document.querySelector(`input[name="ocrMode"][value="${saved.ocrMode}"]`);
      if (radio) radio.checked = true;
    }
    if (saved.language && elements.language) elements.language.value = saved.language;
    if (saved.dpi && elements.dpi) elements.dpi.value = String(saved.dpi);
    if (typeof saved.preserveImages === "boolean" && elements.preserveImages) elements.preserveImages.checked = saved.preserveImages;
    if (typeof saved.preserveLayout === "boolean" && elements.preserveLayout) elements.preserveLayout.checked = saved.preserveLayout;
    if (typeof saved.detectTables === "boolean" && elements.detectTables) elements.detectTables.checked = saved.detectTables;
    if (typeof saved.pageBreaks === "boolean" && elements.pageBreaks) elements.pageBreaks.checked = saved.pageBreaks;
    if (typeof saved.mathOcr === "boolean" && elements.mathOcr) elements.mathOcr.checked = saved.mathOcr;
    if (elements.mathEngine && ["rules", "local_hybrid"].includes(saved.mathEngine)) {
      elements.mathEngine.value = saved.mathEngine;
    }

    updateOutputFormatUI();
    updateMathEngineHelp();
    updateAdvancedBadge();
    updateActivePresetHighlight();
  } catch (_error) {
    /* json parse error */
  }
}

function applyPreset(presetName) {
  if (state.busy) return;
  if (presetName === "math") {
    const radio = document.querySelector('input[name="ocrMode"][value="auto"]');
    if (radio) radio.checked = true;
    if (elements.mathOcr) elements.mathOcr.checked = true;
    if (elements.mathEngine) elements.mathEngine.value = "local_hybrid";
    if (elements.preserveLayout) elements.preserveLayout.checked = true;
    if (elements.preserveImages) elements.preserveImages.checked = true;
    if (elements.detectTables) elements.detectTables.checked = true;
    if (elements.pageBreaks) elements.pageBreaks.checked = true;
    if (elements.dpi) elements.dpi.value = "300";
  } else if (presetName === "scan") {
    const radio = document.querySelector('input[name="ocrMode"][value="force"]');
    if (radio) radio.checked = true;
    if (elements.mathOcr) elements.mathOcr.checked = false;
    if (elements.language) elements.language.value = "vie";
    if (elements.preserveLayout) elements.preserveLayout.checked = true;
    if (elements.preserveImages) elements.preserveImages.checked = true;
    if (elements.detectTables) elements.detectTables.checked = true;
    if (elements.pageBreaks) elements.pageBreaks.checked = true;
    if (elements.dpi) elements.dpi.value = "300";
  } else if (presetName === "tables") {
    const radio = document.querySelector('input[name="ocrMode"][value="auto"]');
    if (radio) radio.checked = true;
    if (elements.mathOcr) elements.mathOcr.checked = false;
    if (elements.preserveLayout) elements.preserveLayout.checked = true;
    if (elements.preserveImages) elements.preserveImages.checked = true;
    if (elements.detectTables) elements.detectTables.checked = true;
    if (elements.pageBreaks) elements.pageBreaks.checked = true;
    if (elements.dpi) elements.dpi.value = "300";
  } else if (presetName === "default") {
    const radio = document.querySelector('input[name="ocrMode"][value="auto"]');
    if (radio) radio.checked = true;
    if (elements.language) elements.language.value = "vie";
    if (elements.dpi) elements.dpi.value = "300";
    if (elements.preserveLayout) elements.preserveLayout.checked = true;
    if (elements.preserveImages) elements.preserveImages.checked = true;
    if (elements.detectTables) elements.detectTables.checked = true;
    if (elements.pageBreaks) elements.pageBreaks.checked = true;
    if (elements.mathOcr) elements.mathOcr.checked = false;
    if (elements.mathEngine) elements.mathEngine.value = "local_hybrid";
  }

  updateMathEngineHelp();
  saveUserSettings();
  showNotice(`Đã áp dụng mẫu: ${presetName === 'math' ? 'Toán học & Đồ án' : presetName === 'scan' ? 'Sách & Scan' : presetName === 'tables' ? 'Bảng dữ liệu' : 'Mặc định'}`);
  setTimeout(() => showNotice(""), 3000);
}

// Attach change listeners for automatic setting persistence
const settingWatchList = [
  elements.language,
  elements.dpi,
  elements.preserveLayout,
  elements.preserveImages,
  elements.detectTables,
  elements.pageBreaks,
  elements.mathOcr,
  elements.mathEngine,
];
settingWatchList.forEach((input) => {
  if (input) {
    input.addEventListener("change", () => {
      saveUserSettings();
    });
  }
});

document.querySelectorAll('input[name="ocrMode"]').forEach((radio) => {
  radio.addEventListener("change", () => {
    saveUserSettings();
  });
});

// Presets buttons event listeners
document.querySelectorAll(".preset-pill[data-preset]").forEach((btn) => {
  btn.addEventListener("click", () => {
    applyPreset(btn.dataset.preset);
  });
});

const resetSettingsBtn = elements.resetSettingsBtn || document.querySelector("#resetSettingsBtn");
if (resetSettingsBtn) {
  resetSettingsBtn.addEventListener("click", () => {
    applyPreset("default");
  });
}

// Convert more files button
const convertMoreBtn = elements.convertMoreBtn || document.querySelector("#convertMoreBtn");
if (convertMoreBtn) {
  convertMoreBtn.addEventListener("click", () => {
    if (state.busy) return;
    state.files = [];
    resetProgress();
    renderFiles();
    elements.fileInput.click();
  });
}

// Global keyboard shortcut: Ctrl + Enter / Cmd + Enter
document.addEventListener("keydown", (event) => {
  if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
    if (state.busy) return;
    event.preventDefault();
    if (state.files.length > 0) {
      convertFiles();
    } else {
      showNotice("Vui lòng chọn hoặc kéo thả ít nhất 1 tài liệu trước khi bấm chuyển đổi.");
      elements.fileInput.click();
    }
  }
});

window.addEventListener("beforeunload", () => {
  if (state.resultUrl && state.resultIsObjectUrl) URL.revokeObjectURL(state.resultUrl);
  previewUrls.forEach((url) => URL.revokeObjectURL(url));
  previewUrls.clear();
});

// Restore saved settings on initial load
restoreUserSettings();
loadStatus().then(() => {
  // Re-verify active preset highlight and badge after status sync
  updateAdvancedBadge();
  updateActivePresetHighlight();
});
