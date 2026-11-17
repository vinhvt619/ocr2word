from __future__ import annotations

import re
import shutil
import tempfile
import unicodedata
import zipfile
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any, Callable, Literal, Mapping, Sequence

from PIL import Image as PillowImage
from PIL import UnidentifiedImageError
from pydantic import BaseModel, ConfigDict, Field

from app.converters import (
    ConversionError,
    ConversionOptions,
    ConversionResult,
    convert_excel,
    convert_image,
    convert_pdf,
    validate_pdf_signature,
)


PDF_EXTENSIONS = {".pdf"}
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
ALLOWED_EXTENSIONS = PDF_EXTENSIONS | EXCEL_EXTENSIONS | IMAGE_EXTENSIONS

# Modal Function arguments currently have a 100 MB transport ceiling. Keep a
# buffer for serialization metadata instead of accepting a request that cannot
# be handed to the OCR worker after the browser has finished uploading it.
MAX_FILE_BYTES = 80 * 1024 * 1024
MAX_TOTAL_BYTES = 90 * 1024 * 1024
MAX_OUTPUT_BYTES = 90 * 1024 * 1024
MAX_FILES = 20

ProgressCallback = Callable[[int, str, str | None, int, str], None]


class ServerlessConversionOptions(BaseModel):
    model_config = ConfigDict(extra="forbid")

    output_format: Literal["docx", "xlsx"] = "docx"
    language: Literal["vie", "eng", "vie+eng", "eng+vie"] = "vie"
    ocr_mode: Literal["auto", "force", "text"] = "auto"
    dpi: Literal[200, 300, 400] = 300
    page_breaks: bool = True
    preserve_images: bool = True
    preserve_layout: bool = True
    detect_tables: bool = True
    math_ocr: bool = False
    math_engine: Literal["rules", "local_hybrid"] | None = None


@dataclass(frozen=True)
class ServerlessJobOutput:
    filename: str
    media_type: str
    content: bytes
    report: list[dict[str, object]]

    def to_modal_payload(self) -> dict[str, object]:
        return {
            "filename": self.filename,
            "media_type": self.media_type,
            "content": self.content,
            "report": self.report,
        }


def _report_progress(
    callback: ProgressCallback | None,
    percent: int | float,
    message: str,
    *,
    current_file: str | None = None,
    completed_files: int = 0,
    status: str = "processing",
) -> None:
    if callback is None:
        return
    callback(
        max(0, min(100, round(float(percent)))),
        message,
        current_file,
        completed_files,
        status,
    )


def _safe_stem(filename: str) -> str:
    stem = Path(filename).stem.strip() or "tai-lieu"
    stem = unicodedata.normalize("NFC", stem)
    stem = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "-", stem)
    stem = re.sub(r"\s+", " ", stem).strip(" .-")
    return stem[:100] or "tai-lieu"


def _unique_name(desired: str, used_names: set[str]) -> str:
    path = Path(desired)
    candidate = desired
    counter = 2
    while candidate.casefold() in used_names:
        candidate = f"{path.stem}-{counter}{path.suffix}"
        counter += 1
    used_names.add(candidate.casefold())
    return candidate


def _validate_signature(path: Path, suffix: str) -> None:
    with path.open("rb") as stream:
        signature = stream.read(8)
    if suffix == ".pdf":
        validate_pdf_signature(path, signature)
        return
    if suffix in {".xlsx", ".xlsm"} and not signature.startswith(b"PK"):
        raise ConversionError("File Excel hiện đại không phải gói XLSX/XLSM hợp lệ.")
    if suffix == ".xls" and not signature.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        raise ConversionError("File .xls không đúng định dạng Excel 97-2003.")
    if suffix in IMAGE_EXTENSIONS:
        try:
            with PillowImage.open(path) as image:
                image.verify()
        except (UnidentifiedImageError, OSError, SyntaxError, ValueError) as exc:
            raise ConversionError("File hình ảnh không hợp lệ hoặc đã bị hỏng.") from exc


def _convert_one(input_path: Path, output_path: Path, options: ConversionOptions) -> ConversionResult:
    if options.output_format == "xlsx":
        from app.converters import convert_to_excel
        return convert_to_excel(input_path, output_path, options)
    suffix = input_path.suffix.lower()
    if suffix in PDF_EXTENSIONS:
        return convert_pdf(input_path, output_path, options)
    if suffix in IMAGE_EXTENSIONS:
        return convert_image(input_path, output_path, options)
    return convert_excel(input_path, output_path, options)


def _package_results(
    successes: list[tuple[Path, str]],
    report: list[dict[str, object]],
    work_dir: Path,
) -> tuple[Path, str, str]:
    if len(report) == 1 and len(successes) == 1:
        media_type = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if successes[0][0].suffix.lower() == ".xlsx"
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        return (
            successes[0][0],
            successes[0][1],
            media_type,
        )

    output_name = "ocr2word-ket-qua.zip"
    zip_path = work_dir / output_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for output_path, filename in successes:
            archive.write(output_path, filename)
        if any(item["status"] == "error" or item.get("warnings") for item in report):
            lines = ["BÁO CÁO CHUYỂN ĐỔI OCR2WORD", ""]
            for item in report:
                if item["status"] == "success":
                    lines.append(f"[THÀNH CÔNG] {item['file']} -> {item['output']}")
                    for warning in item.get("warnings", []):
                        lines.append(f"  - Lưu ý: {warning}")
                else:
                    lines.append(f"[LỖI] {item['file']}: {item['message']}")
            archive.writestr("bao-cao-chuyen-doi.txt", "\n".join(lines).encode("utf-8"))
    return zip_path, output_name, "application/zip"


def run_serverless_conversion(
    files: Sequence[Mapping[str, Any]],
    option_payload: Mapping[str, Any],
    *,
    progress_callback: ProgressCallback | None = None,
) -> ServerlessJobOutput:
    """Convert uploaded bytes inside one disposable serverless worker.

    Inputs and the returned DOCX/ZIP are intentionally held only in the
    Function invocation. Modal encrypts those payloads and expires them after
    at most seven days, so no persistent Blob/database is needed.
    """

    if not files:
        raise ConversionError("Hãy chọn ít nhất một file.")
    if len(files) > MAX_FILES:
        raise ConversionError(f"Mỗi lượt chỉ nhận tối đa {MAX_FILES} file.")
    try:
        option_data = ServerlessConversionOptions.model_validate(dict(option_payload))
    except Exception as exc:
        raise ConversionError(f"Tùy chọn chuyển đổi không hợp lệ: {exc}") from exc

    work_dir = Path(tempfile.mkdtemp(prefix="ocr2word-modal-"))
    input_dir = work_dir / "input"
    output_dir = work_dir / "output"
    input_dir.mkdir()
    output_dir.mkdir()

    try:
        saved_inputs: list[tuple[Path, str]] = []
        used_input_names: set[str] = set()
        total_size = 0
        _report_progress(progress_callback, 21, "Container OCR đang nhận tài liệu…", status="uploading")
        for index, item in enumerate(files):
            original_name = str(item.get("name") or "tai-lieu")
            content = item.get("content")
            if not isinstance(content, bytes):
                raise ConversionError(f"'{original_name}' không có dữ liệu nhị phân hợp lệ.")
            if not content:
                raise ConversionError(f"'{original_name}' là file rỗng.")
            if len(content) > MAX_FILE_BYTES:
                raise ConversionError(f"'{original_name}' lớn hơn giới hạn 80 MB của Modal.")
            total_size += len(content)
            if total_size > MAX_TOTAL_BYTES:
                raise ConversionError("Tổng dung lượng tải lên lớn hơn giới hạn 90 MB của Modal.")

            suffix = Path(original_name).suffix.lower()
            if option_data.output_format == "xlsx" and suffix in EXCEL_EXTENSIONS:
                raise ConversionError(
                    "Chuyển đổi sang Excel (.xlsx) chỉ nhận file PDF hoặc hình ảnh (không nhận file Excel làm đầu vào)."
                )
            if suffix not in ALLOWED_EXTENSIONS:
                raise ConversionError(
                    f"'{original_name}' không được hỗ trợ. Chỉ nhận PDF, hình ảnh và Excel."
                )
            internal_name = _unique_name(f"{_safe_stem(original_name)}{suffix}", used_input_names)
            input_path = input_dir / internal_name
            input_path.write_bytes(content)
            _validate_signature(input_path, suffix)
            saved_inputs.append((input_path, original_name))
            _report_progress(
                progress_callback,
                21 + ((index + 1) / len(files)) * 9,
                f"Đã nhận file {index + 1}/{len(files)}.",
                current_file=original_name,
                completed_files=0,
                status="uploading",
            )

        if option_data.output_format == "xlsx":
            if option_data.math_ocr:
                raise ConversionError("Chuyển đổi sang Excel (.xlsx) không hỗ trợ tính năng nhận diện công thức toán.")

        options = ConversionOptions(
            output_format=option_data.output_format,
            language=option_data.language,
            ocr_mode=option_data.ocr_mode,
            dpi=option_data.dpi,
            page_breaks=option_data.page_breaks,
            preserve_images=option_data.preserve_images,
            preserve_layout=option_data.preserve_layout,
            detect_tables=option_data.detect_tables,
            math_ocr=option_data.math_ocr,
            math_engine=option_data.math_engine,
        )
        options.validate()

        successes: list[tuple[Path, str]] = []
        report: list[dict[str, object]] = []
        used_output_names: set[str] = set()
        for file_index, (input_path, original_name) in enumerate(saved_inputs):
            if option_data.output_format == "xlsx":
                output_suffix = "-editable.xlsx"
            else:
                output_suffix = "-editable.docx"
            output_name = _unique_name(
                f"{_safe_stem(original_name)}{output_suffix}",
                used_output_names,
            )
            output_path = output_dir / output_name

            def file_progress(file_fraction: float, message: str) -> None:
                overall = 30 + ((file_index + file_fraction) / len(saved_inputs)) * 64
                _report_progress(
                    progress_callback,
                    overall,
                    message,
                    current_file=original_name,
                    completed_files=file_index,
                    status="converting",
                )

            def converter_progress(local_percent: int, message: str) -> None:
                file_progress(local_percent / 100, message)

            try:
                result = _convert_one(
                    input_path,
                    output_path,
                    replace(options, progress_callback=converter_progress),
                )
                successes.append((result.output_path, output_name))
                report.append(
                    {
                        "file": original_name,
                        "status": "success",
                        "output": output_name,
                        "warnings": list(dict.fromkeys(result.warnings)),
                        "pages": result.page_count,
                        "sheets": result.sheet_count,
                        "tables": result.table_count,
                        "math": result.math_report(),
                    }
                )
            except ConversionError as exc:
                report.append({"file": original_name, "status": "error", "message": str(exc)})
            except Exception as exc:
                report.append(
                    {
                        "file": original_name,
                        "status": "error",
                        "message": f"Lỗi không mong đợi khi chuyển đổi: {exc}",
                    }
                )

            _report_progress(
                progress_callback,
                30 + ((file_index + 1) / len(saved_inputs)) * 64,
                f"Đã xử lý file {file_index + 1}/{len(saved_inputs)}.",
                current_file=original_name,
                completed_files=file_index + 1,
                status="converting",
            )

        if not successes:
            messages = [str(item.get("message", "Không thể chuyển đổi.")) for item in report]
            raise ConversionError("\n".join(messages))

        _report_progress(
            progress_callback,
            96,
            "Đang chuẩn bị file tải xuống…",
            completed_files=len(successes),
            status="packaging",
        )
        output_path, output_name, media_type = _package_results(successes, report, work_dir)
        output_content = output_path.read_bytes()
        if len(output_content) > MAX_OUTPUT_BYTES:
            raise ConversionError(
                "File kết quả lớn hơn giới hạn truyền 90 MB của Modal. "
                "Hãy chia tài liệu thành các phần nhỏ hơn rồi chuyển đổi lại."
            )
        _report_progress(
            progress_callback,
            100,
            "Đã chuyển đổi xong. Kết quả sẵn sàng để tải xuống.",
            completed_files=len(successes),
            status="complete",
        )
        return ServerlessJobOutput(
            filename=output_name,
            media_type=media_type,
            content=output_content,
            report=report,
        )
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
