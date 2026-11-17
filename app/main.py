from __future__ import annotations

import base64
import hmac
import json
import logging
import os
import platform
import re
import shutil
import tempfile
import time
import unicodedata
import urllib.parse
import zipfile
from dataclasses import replace
from pathlib import Path
from typing import Annotated, Callable

import httpx
from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from PIL import Image as PillowImage
from PIL import UnidentifiedImageError
from pydantic import BaseModel, ConfigDict, Field, field_validator
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool

from app import __version__
from app.converters import (
    ConversionError,
    ConversionOptions,
    analyze_pdf,
    convert_excel,
    convert_image,
    convert_pdf,
    extract_pdf_pages,
    validate_pdf_signature,
)
from app.converters.math_pipeline import get_math_engines_status
from app.services.ocr import get_ocr_engine_info
from app.services.docx_merge import merge_docx_parts
from app.services.progress import ProgressRegistry


LOGGER = logging.getLogger("ocr2word")
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)
PROJECT_ROOT = Path(__file__).resolve().parents[1]
STATIC_DIR = Path(__file__).resolve().parent / "static"
PDF_EXTENSIONS = {".pdf"}
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
ALLOWED_EXTENSIONS = PDF_EXTENSIONS | EXCEL_EXTENSIONS | IMAGE_EXTENSIONS
MAX_FILE_BYTES = 150 * 1024 * 1024
MAX_TOTAL_BYTES = 300 * 1024 * 1024
MAX_FILES = 20
READ_CHUNK_SIZE = 1024 * 1024
JOB_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{8,80}$")
PROGRESS = ProgressRegistry()
VERCEL_BLOB_HOST = "blob.vercel-storage.com"


app = FastAPI(
    title="OCR2Word",
    description="Chuyển PDF, hình ảnh và Excel sang Word có thể chỉnh sửa.",
    version=__version__,
    docs_url="/api/docs",
    redoc_url=None,
)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


def _validate_vercel_blob_url(value: str, *, allow_presigned_upload: bool = False) -> str:
    parsed = urllib.parse.urlsplit(value)
    hostname = (parsed.hostname or "").casefold()
    is_blob_host = hostname == VERCEL_BLOB_HOST or hostname.endswith(f".{VERCEL_BLOB_HOST}")
    is_presigned_upload = (
        allow_presigned_upload
        and hostname == "vercel.com"
        and parsed.path == "/api/blob/"
    )
    if (
        len(value) > 8_000
        or parsed.scheme != "https"
        or not hostname
        or not (is_blob_host or is_presigned_upload)
        or parsed.username
        or parsed.password
        or parsed.fragment
    ):
        raise ValueError("URL Vercel Blob không hợp lệ")
    return value


def _validate_vercel_blob_upload_url(value: str) -> str:
    return _validate_vercel_blob_url(value, allow_presigned_upload=True)


class _VercelInputFile(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str = Field(min_length=1, max_length=255)
    size: int = Field(ge=1, le=MAX_FILE_BYTES)
    content_type: str = Field(min_length=1, max_length=200)
    download_url: str

    _download_url_validator = field_validator("download_url")(_validate_vercel_blob_url)


class _VercelConversionOptions(BaseModel):
    model_config = ConfigDict(extra="forbid")

    output_format: str = Field(default="docx", pattern=r"^(docx|xlsx)$")
    language: str = Field(default="vie", min_length=2, max_length=80, pattern=r"^[A-Za-z0-9_+.-]+$")
    ocr_mode: str = Field(default="auto", pattern=r"^(auto|force|text)$")
    dpi: int = Field(default=300, ge=150, le=600)
    page_breaks: bool = True
    preserve_images: bool = True
    preserve_layout: bool = True
    detect_tables: bool = True
    math_ocr: bool = False
    math_engine: str | None = Field(default=None, pattern=r"^(rules|local_hybrid)$")


class _VercelOutput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str = Field(min_length=1, max_length=255)
    content_type: str = Field(min_length=1, max_length=200)
    upload_url: str

    _upload_url_validator = field_validator("upload_url")(_validate_vercel_blob_upload_url)


class _VercelPageRange(BaseModel):
    model_config = ConfigDict(extra="forbid")

    start_page: int = Field(ge=1, le=10_000)
    end_page: int = Field(ge=1, le=10_000)


class _VercelConversionJob(BaseModel):
    model_config = ConfigDict(extra="forbid")

    job_id: str = Field(pattern=r"^[a-f0-9]{32}$")
    files: list[_VercelInputFile] = Field(min_length=1, max_length=MAX_FILES)
    options: _VercelConversionOptions
    output: _VercelOutput
    page_range: _VercelPageRange | None = None
    notify_progress: bool = True


class _VercelPdfAnalysisJob(BaseModel):
    model_config = ConfigDict(extra="forbid")

    job_id: str = Field(pattern=r"^[a-f0-9]{32}$")
    file: _VercelInputFile
    options: _VercelConversionOptions


class _VercelMergeJob(BaseModel):
    model_config = ConfigDict(extra="forbid")

    job_id: str = Field(pattern=r"^[a-f0-9]{32}$")
    parts: list[_VercelInputFile] = Field(min_length=1, max_length=200)
    output: _VercelOutput


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


async def _save_upload(
    upload: UploadFile,
    destination: Path,
    remaining_total: int,
    progress_callback: Callable[[int], None] | None = None,
) -> int:
    written = 0
    with destination.open("wb") as output:
        while chunk := await upload.read(READ_CHUNK_SIZE):
            written += len(chunk)
            if written > MAX_FILE_BYTES:
                raise ConversionError(f"'{upload.filename}' lớn hơn giới hạn 150 MB.")
            if written > remaining_total:
                raise ConversionError("Tổng dung lượng tải lên lớn hơn giới hạn 300 MB.")
            output.write(chunk)
            if progress_callback:
                progress_callback(written)
    if written == 0:
        raise ConversionError(f"'{upload.filename}' là file rỗng.")
    return written


def _encode_warnings(payload: list[dict[str, object]]) -> str:
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("ascii")


def _convert_one(input_path: Path, output_path: Path, options: ConversionOptions):
    if options.output_format == "xlsx":
        from .converters import convert_to_excel
        return convert_to_excel(input_path, output_path, options)
    suffix = input_path.suffix.lower()
    if suffix in PDF_EXTENSIONS:
        return convert_pdf(input_path, output_path, options)
    if suffix in IMAGE_EXTENSIONS:
        return convert_image(input_path, output_path, options)
    return convert_excel(input_path, output_path, options)


def _internal_secret() -> str:
    configured = os.environ.get("OCR2WORD_INTERNAL_SECRET", "").strip()
    if configured:
        return configured
    if not os.environ.get("VERCEL"):
        return "ocr2word-local-development-only"
    raise ConversionError("Container OCR thiếu biến môi trường OCR2WORD_INTERNAL_SECRET.")


def _authorize_internal_request(authorization: str | None) -> None:
    expected = f"Bearer {_internal_secret()}"
    if not authorization or not hmac.compare_digest(authorization, expected):
        raise HTTPException(status_code=401, detail="Không được phép gọi worker OCR.")


class _VercelProgressReporter:
    def __init__(self, job_id: str, *, enabled: bool = True) -> None:
        self.job_id = job_id
        self.control_url = os.environ.get("OCR2WORD_CONTROL_URL", "").strip() if enabled else ""
        self.last_percent = -1
        self.last_sent_at = 0.0

    def update(
        self,
        percent: int | float,
        message: str,
        *,
        status: str = "processing",
        current_file: str | None = None,
        completed_files: int = 0,
        force: bool = False,
    ) -> None:
        if not self.control_url:
            return
        rounded = max(0, min(99, round(float(percent))))
        now = time.monotonic()
        if not force and rounded <= self.last_percent and now - self.last_sent_at < 0.8:
            return
        if not force and rounded < self.last_percent + 1 and now - self.last_sent_at < 0.8:
            return
        endpoint = urllib.parse.urljoin(self.control_url.rstrip("/") + "/", "api/internal/progress")
        try:
            response = httpx.post(
                endpoint,
                json={
                    "job_id": self.job_id,
                    "status": status,
                    "percent": rounded,
                    "message": message,
                    "current_file": current_file,
                    "completed_files": completed_files,
                },
                headers={"Authorization": f"Bearer {_internal_secret()}"},
                timeout=3,
            )
            response.raise_for_status()
            self.last_percent = max(self.last_percent, rounded)
            self.last_sent_at = now
        except (httpx.HTTPError, ConversionError):
            LOGGER.debug("Could not persist progress for Vercel job %s", self.job_id, exc_info=True)


def _download_vercel_input(
    source: _VercelInputFile,
    destination: Path,
    *,
    total_written_before: int,
    reporter: _VercelProgressReporter,
    file_index: int,
    total_files: int,
) -> int:
    written = 0
    try:
        timeout = httpx.Timeout(connect=20, read=300, write=60, pool=20)
        with httpx.Client(timeout=timeout, follow_redirects=False) as client:
            with client.stream("GET", source.download_url) as response:
                response.raise_for_status()
                with destination.open("wb") as output:
                    for chunk in response.iter_bytes(READ_CHUNK_SIZE):
                        if not chunk:
                            continue
                        written += len(chunk)
                        if written > MAX_FILE_BYTES or total_written_before + written > MAX_TOTAL_BYTES:
                            raise ConversionError("File tải từ Vercel Blob vượt giới hạn dung lượng.")
                        output.write(chunk)
                        fraction = (file_index + min(1.0, written / max(1, source.size))) / total_files
                        reporter.update(
                            22 + fraction * 8,
                            f"Container đang nhận file {file_index + 1}/{total_files} từ Vercel Blob…",
                            current_file=source.name,
                            completed_files=file_index,
                        )
    except ConversionError:
        raise
    except httpx.HTTPError as exc:
        raise ConversionError(f"Không tải được '{source.name}' từ Vercel Blob.") from exc
    if written != source.size:
        raise ConversionError(f"Kích thước '{source.name}' trên Blob không khớp với lúc tải lên.")
    return written


def _upload_vercel_result(path: Path, destination: _VercelOutput) -> None:
    try:
        timeout = httpx.Timeout(connect=20, read=300, write=300, pool=20)
        with path.open("rb") as stream, httpx.Client(timeout=timeout, follow_redirects=False) as client:
            response = client.put(
                destination.upload_url,
                content=stream,
                headers={
                    "Content-Type": destination.content_type,
                    "Content-Length": str(path.stat().st_size),
                },
            )
            response.raise_for_status()
    except (OSError, httpx.HTTPError) as exc:
        raise ConversionError("Không thể lưu kết quả vào Vercel Blob.") from exc


def _package_worker_results(
    successes: list[tuple[Path, str]],
    report: list[dict[str, object]],
    work_dir: Path,
) -> Path:
    if len(report) == 1 and len(successes) == 1:
        return successes[0][0]
    zip_path = work_dir / "ocr2word-ket-qua.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for output_path, output_name in successes:
            archive.write(output_path, output_name)
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
    return zip_path


DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ZIP_MIME = "application/zip"


def _validate_worker_output_contract(payload: _VercelConversionJob) -> None:
    if not payload.files:
        raise ConversionError("Hãy chọn ít nhất một file.")
    if sum(item.size for item in payload.files) > MAX_TOTAL_BYTES:
        raise ConversionError("Tổng dung lượng file vượt giới hạn 300 MB.")

    fmt = (getattr(payload.options, "output_format", None) or "docx").lower()
    if fmt not in ("docx", "xlsx"):
        raise ConversionError(f"Định dạng xuất '{fmt}' không được hỗ trợ.")

    file_count = len(payload.files)
    out_ext = Path(payload.output.name).suffix.lower()
    ct = payload.output.content_type.strip().lower()

    if file_count > 1:
        if out_ext != ".zip":
            raise ConversionError(
                f"Xử lý nhiều file ({file_count} files) yêu cầu tên file kết quả có đuôi '.zip', nhưng nhận được '{out_ext}' ({payload.output.name})."
            )
        if ct != ZIP_MIME:
            raise ConversionError(
                f"Xử lý nhiều file ({file_count} files) yêu cầu Content-Type '{ZIP_MIME}', nhưng nhận được '{payload.output.content_type}'."
            )
    else:
        if fmt == "xlsx":
            expected_ext = ".xlsx"
            expected_mime = XLSX_MIME
        else:
            expected_ext = ".docx"
            expected_mime = DOCX_MIME

        if out_ext != expected_ext:
            raise ConversionError(
                f"Định dạng xuất '{fmt}' yêu cầu tên file kết quả có đuôi '{expected_ext}', nhưng nhận được '{out_ext}' ({payload.output.name})."
            )
        if ct != expected_mime:
            raise ConversionError(
                f"Định dạng xuất '{fmt}' yêu cầu Content-Type '{expected_mime}', nhưng nhận được '{payload.output.content_type}'."
            )

    # Incompatible options for XLSX
    if fmt == "xlsx":
        if payload.options.math_ocr:
            raise ConversionError("Chuyển đổi sang Excel (.xlsx) không hỗ trợ tính năng nhận diện công thức toán.")
        if payload.page_range is not None:
            raise ConversionError("Chuyển đổi sang Excel (.xlsx) không hỗ trợ chia nhỏ trang.")

    # Page range constraints
    if payload.page_range is not None:
        if file_count != 1:
            raise ConversionError("Chỉ có thể chia trang khi xử lý đúng một file PDF.")


def _run_vercel_conversion(payload: _VercelConversionJob) -> dict[str, object]:
    _validate_worker_output_contract(payload)
    reporter = _VercelProgressReporter(payload.job_id, enabled=payload.notify_progress)
    work_dir = Path(tempfile.mkdtemp(prefix="ocr2word-vercel-"))
    input_dir = work_dir / "input"
    output_dir = work_dir / "output"
    input_dir.mkdir()
    output_dir.mkdir()
    try:
        saved_inputs: list[tuple[Path, str]] = []
        used_input_names: set[str] = set()
        total_size = 0
        reporter.update(22, "Container OCR đang nhận tài liệu từ Vercel Blob…", force=True)
        for index, source in enumerate(payload.files):
            suffix = Path(source.name).suffix.lower()
            if suffix not in ALLOWED_EXTENSIONS:
                raise ConversionError(f"'{source.name}' không thuộc định dạng được hỗ trợ.")
            internal_name = _unique_name(f"{_safe_stem(source.name)}{suffix}", used_input_names)
            input_path = input_dir / internal_name
            total_size += _download_vercel_input(
                source,
                input_path,
                total_written_before=total_size,
                reporter=reporter,
                file_index=index,
                total_files=len(payload.files),
            )
            _validate_signature(input_path, suffix)
            saved_inputs.append((input_path, source.name))

        if payload.page_range is not None:
            if payload.options.output_format == "xlsx":
                raise ConversionError("Chuyển đổi sang Excel (.xlsx) không hỗ trợ chia nhỏ trang.")
            if len(saved_inputs) != 1 or saved_inputs[0][0].suffix.lower() != ".pdf":
                raise ConversionError("Chỉ có thể chia trang khi xử lý đúng một file PDF.")
            source_path, original_name = saved_inputs[0]
            subset_path = input_dir / (
                f"chunk-{payload.page_range.start_page:04d}-{payload.page_range.end_page:04d}.pdf"
            )
            extract_pdf_pages(
                source_path,
                subset_path,
                start_page=payload.page_range.start_page,
                end_page=payload.page_range.end_page,
            )
            saved_inputs = [(subset_path, original_name)]

        option_data = payload.options
        if option_data.output_format == "xlsx":
            if option_data.math_ocr:
                raise ConversionError("Chuyển đổi sang Excel (.xlsx) không hỗ trợ tính năng nhận diện công thức toán.")
            for input_path, original_name in saved_inputs:
                if input_path.suffix.lower() in EXCEL_EXTENSIONS:
                    raise ConversionError(
                        "Chuyển đổi sang Excel (.xlsx) chỉ nhận file PDF hoặc hình ảnh (không nhận file Excel làm đầu vào)."
                    )

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
            output_name = _unique_name(f"{_safe_stem(original_name)}{output_suffix}", used_output_names)
            output_path = output_dir / output_name

            def file_progress(file_fraction: float, message: str) -> None:
                overall = 31 + ((file_index + file_fraction) / len(saved_inputs)) * 61
                reporter.update(
                    overall,
                    message,
                    current_file=original_name,
                    completed_files=file_index,
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
                        "warnings": result.warnings,
                        "pages": result.page_count,
                        "sheets": result.sheet_count,
                        "tables": result.table_count,
                        "math": result.math_report(),
                    }
                )
            except ConversionError as exc:
                report.append({"file": original_name, "status": "error", "message": str(exc)})
            except Exception as exc:
                LOGGER.exception("Unexpected Vercel conversion failure for %s", original_name)
                report.append(
                    {
                        "file": original_name,
                        "status": "error",
                        "message": f"Lỗi không mong đợi khi chuyển đổi: {exc}",
                    }
                )
            reporter.update(
                31 + ((file_index + 1) / len(saved_inputs)) * 61,
                f"Đã xử lý file {file_index + 1}/{len(saved_inputs)}.",
                current_file=original_name,
                completed_files=file_index + 1,
                force=True,
            )

        if not successes:
            raise ConversionError("\n".join(str(item.get("message", "Không thể chuyển đổi.")) for item in report))
        packaging_message = "Đang đóng gói kết quả Excel…" if option_data.output_format == "xlsx" else "Đang đóng gói kết quả Word…"
        reporter.update(95, packaging_message, status="packaging", force=True)
        result_path = _package_worker_results(successes, report, work_dir)
        reporter.update(98, "Đang lưu kết quả vào Vercel Blob…", status="packaging", force=True)
        _upload_vercel_result(result_path, payload.output)
        return {
            "report": report,
            "output_name": payload.output.name,
            "output_size": result_path.stat().st_size,
        }
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def _conversion_options(option_data: _VercelConversionOptions) -> ConversionOptions:
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
    return options


def _run_vercel_pdf_analysis(payload: _VercelPdfAnalysisJob) -> dict[str, object]:
    if Path(payload.file.name).suffix.lower() != ".pdf":
        raise ConversionError("Chỉ có thể phân tích trước file PDF.")
    work_dir = Path(tempfile.mkdtemp(prefix="ocr2word-analyze-"))
    input_path = work_dir / "source.pdf"
    try:
        _download_vercel_input(
            payload.file,
            input_path,
            total_written_before=0,
            reporter=_VercelProgressReporter(payload.job_id, enabled=False),
            file_index=0,
            total_files=1,
        )
        _validate_signature(input_path, ".pdf")
        analysis = analyze_pdf(input_path, _conversion_options(payload.options))
        return analysis.to_dict()
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def _validate_docx(path: Path) -> None:
    if not zipfile.is_zipfile(path):
        raise ConversionError(f"'{path.name}' không phải file Word DOCX hợp lệ.")
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
    except (OSError, zipfile.BadZipFile) as exc:
        raise ConversionError(f"Không thể đọc phần Word '{path.name}'.") from exc
    if "[Content_Types].xml" not in names or "word/document.xml" not in names:
        raise ConversionError(f"'{path.name}' thiếu cấu trúc Word bắt buộc.")


def _run_vercel_merge(payload: _VercelMergeJob) -> dict[str, object]:
    work_dir = Path(tempfile.mkdtemp(prefix="ocr2word-merge-"))
    input_dir = work_dir / "parts"
    input_dir.mkdir()
    output_path = work_dir / "merged.docx"
    try:
        part_paths: list[Path] = []
        total_size = 0
        reporter = _VercelProgressReporter(payload.job_id, enabled=False)
        for index, part in enumerate(payload.parts):
            part_path = input_dir / f"part-{index + 1:04d}.docx"
            total_size += _download_vercel_input(
                part,
                part_path,
                total_written_before=total_size,
                reporter=reporter,
                file_index=index,
                total_files=len(payload.parts),
            )
            _validate_docx(part_path)
            part_paths.append(part_path)
        merge_docx_parts(part_paths, output_path)
        _validate_docx(output_path)
        _upload_vercel_result(output_path, payload.output)
        return {
            "output_name": payload.output.name,
            "output_size": output_path.stat().st_size,
            "part_count": len(part_paths),
        }
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.get("/", include_in_schema=False)
async def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html", media_type="text/html; charset=utf-8")


@app.get("/api/status")
async def status() -> dict[str, object]:
    ocr = get_ocr_engine_info()
    return {
        "app": "OCR2Word",
        "version": __version__,
        "ocr": ocr.to_dict(),
        "math_ocr": get_math_engines_status(),
        "formats": sorted(ALLOWED_EXTENSIONS),
        "limits": {
            "max_files": MAX_FILES,
            "max_file_mb": MAX_FILE_BYTES // (1024 * 1024),
            "max_total_mb": MAX_TOTAL_BYTES // (1024 * 1024),
        },
        "runtime": {
            "python": platform.python_version(),
            "system": platform.system(),
            "machine": platform.machine(),
            "container": Path("/.dockerenv").exists(),
            "vercel": bool(os.environ.get("VERCEL")),
        },
    }


@app.post("/internal/convert", include_in_schema=False)
async def vercel_internal_convert(
    payload: _VercelConversionJob,
    authorization: Annotated[str | None, Header()] = None,
) -> dict[str, object]:
    _authorize_internal_request(authorization)
    try:
        return await run_in_threadpool(_run_vercel_conversion, payload)
    except ConversionError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except HTTPException:
        raise
    except Exception as exc:
        LOGGER.exception("Vercel worker failed for job %s", payload.job_id)
        raise HTTPException(status_code=500, detail="Container OCR không thể xử lý tài liệu.") from exc


@app.post("/internal/analyze-pdf", include_in_schema=False)
async def vercel_internal_analyze_pdf(
    payload: _VercelPdfAnalysisJob,
    authorization: Annotated[str | None, Header()] = None,
) -> dict[str, object]:
    _authorize_internal_request(authorization)
    try:
        return await run_in_threadpool(_run_vercel_pdf_analysis, payload)
    except ConversionError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:
        LOGGER.exception("Vercel PDF analysis failed for job %s", payload.job_id)
        raise HTTPException(status_code=500, detail="Container OCR không thể phân tích PDF.") from exc


@app.post("/internal/merge-docx", include_in_schema=False)
async def vercel_internal_merge_docx(
    payload: _VercelMergeJob,
    authorization: Annotated[str | None, Header()] = None,
) -> dict[str, object]:
    _authorize_internal_request(authorization)
    try:
        return await run_in_threadpool(_run_vercel_merge, payload)
    except ConversionError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:
        LOGGER.exception("Vercel DOCX merge failed for job %s", payload.job_id)
        raise HTTPException(status_code=500, detail="Container OCR không thể ghép các phần Word.") from exc


@app.get("/api/progress/{job_id}")
async def conversion_progress(job_id: str) -> dict[str, object]:
    if not JOB_ID_PATTERN.fullmatch(job_id):
        raise HTTPException(status_code=404, detail="Không tìm thấy tiến độ chuyển đổi.")
    snapshot = PROGRESS.get(job_id)
    if snapshot is None:
        raise HTTPException(status_code=404, detail="Không tìm thấy tiến độ chuyển đổi.")
    return snapshot


@app.post("/api/convert")
async def convert(
    files: Annotated[list[UploadFile], File(description="PDF, hình ảnh hoặc Excel cần chuyển đổi")],
    output_format: Annotated[str, Form()] = "docx",
    language: Annotated[str, Form()] = "vie",
    ocr_mode: Annotated[str, Form()] = "auto",
    dpi: Annotated[int, Form()] = 300,
    page_breaks: Annotated[bool, Form()] = True,
    preserve_images: Annotated[bool, Form()] = True,
    preserve_layout: Annotated[bool, Form()] = True,
    detect_tables: Annotated[bool, Form()] = True,
    math_ocr: Annotated[bool, Form()] = False,
    math_engine: Annotated[str | None, Form()] = None,
    job_id: Annotated[str | None, Form()] = None,
) -> FileResponse:
    if not files:
        raise HTTPException(status_code=400, detail="Hãy chọn ít nhất một file.")
    if len(files) > MAX_FILES:
        raise HTTPException(status_code=400, detail=f"Mỗi lượt chỉ nhận tối đa {MAX_FILES} file.")
    job_id = job_id.strip() if job_id else None
    if job_id and not JOB_ID_PATTERN.fullmatch(job_id):
        raise HTTPException(status_code=400, detail="Mã theo dõi tiến độ không hợp lệ.")

    if output_format not in {"docx", "xlsx"}:
        raise HTTPException(status_code=400, detail="Định dạng xuất không hợp lệ (chỉ hỗ trợ docx hoặc xlsx).")
    if output_format == "xlsx":
        if math_ocr:
            raise HTTPException(status_code=400, detail="Chuyển đổi sang Excel (.xlsx) không hỗ trợ tính năng nhận diện công thức toán.")

    normalized_math_engine = math_engine.strip() if math_engine and math_engine.strip() else None
    options = ConversionOptions(
        output_format=output_format,
        language=language,
        ocr_mode=ocr_mode,
        dpi=dpi,
        page_breaks=page_breaks,
        preserve_images=preserve_images,
        preserve_layout=preserve_layout,
        detect_tables=detect_tables,
        math_ocr=math_ocr,
        math_engine=normalized_math_engine,
    )
    try:
        options.validate()
    except ConversionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if job_id:
        PROGRESS.start(job_id, total_files=len(files))

    work_dir = Path(tempfile.mkdtemp(prefix="ocr2word-"))
    input_dir = work_dir / "input"
    output_dir = work_dir / "output"
    input_dir.mkdir()
    output_dir.mkdir()
    total_size = 0
    used_input_names: set[str] = set()
    used_output_names: set[str] = set()
    saved_inputs: list[tuple[Path, str]] = []

    try:
        size_hints = [int(getattr(upload, "size", 0) or 0) for upload in files]
        expected_upload_bytes = sum(size_hints) if all(size > 0 for size in size_hints) else 0
        for upload_index, upload in enumerate(files):
            original_name = upload.filename or "tai-lieu"
            suffix = Path(original_name).suffix.lower()
            if output_format == "xlsx" and suffix in EXCEL_EXTENSIONS:
                raise HTTPException(
                    status_code=400,
                    detail="Chuyển đổi sang Excel (.xlsx) chỉ nhận file PDF hoặc hình ảnh (không nhận file Excel làm đầu vào).",
                )
            if suffix not in ALLOWED_EXTENSIONS:
                raise ConversionError(
                    f"'{original_name}' không được hỗ trợ. Chỉ nhận PDF, hình ảnh PNG/JPG/TIFF/BMP/WEBP và Excel."
                )
            internal_name = _unique_name(f"{_safe_stem(original_name)}{suffix}", used_input_names)
            input_path = input_dir / internal_name
            uploaded_before = total_size

            def upload_progress(written: int, *, index: int = upload_index, name: str = original_name) -> None:
                if expected_upload_bytes:
                    fraction = (uploaded_before + written) / expected_upload_bytes
                else:
                    current_hint = max(1, size_hints[index] or MAX_FILE_BYTES)
                    fraction = (index + min(1.0, written / current_hint)) / len(files)
                PROGRESS.update(
                    job_id,
                    status="uploading",
                    percent=min(14, fraction * 14),
                    message=f"Đang nhận file {index + 1}/{len(files)}…",
                    current_file=name,
                )

            total_size += await _save_upload(
                upload,
                input_path,
                MAX_TOTAL_BYTES - total_size,
                upload_progress,
            )
            _validate_signature(input_path, suffix)
            saved_inputs.append((input_path, original_name))
            await upload.close()
            PROGRESS.update(
                job_id,
                percent=((upload_index + 1) / len(files)) * 14,
                message=f"Đã nhận file {upload_index + 1}/{len(files)}.",
                current_file=original_name,
            )

        successes: list[tuple[Path, str]] = []
        report: list[dict[str, object]] = []
        PROGRESS.update(job_id, status="converting", percent=15, message="Bắt đầu chuyển đổi tài liệu…")
        for file_index, (input_path, original_name) in enumerate(saved_inputs):
            if output_format == "xlsx":
                output_suffix = "-editable.xlsx"
            else:
                output_suffix = "-editable.docx"
            output_name = _unique_name(f"{_safe_stem(original_name)}{output_suffix}", used_output_names)
            output_path = output_dir / output_name

            def file_progress(
                file_fraction: float,
                message: str,
                *,
                index: int = file_index,
                name: str = original_name,
            ) -> None:
                overall = 15 + ((index + file_fraction) / len(saved_inputs)) * 77
                PROGRESS.update(
                    job_id,
                    status="converting",
                    percent=min(92, overall),
                    message=message,
                    current_file=name,
                    completed_files=index,
                )

            def converter_progress(local_percent: int, message: str) -> None:
                file_progress(local_percent / 100, message)

            file_options = replace(options, progress_callback=converter_progress)
            converter_progress(0, f"Đang chuyển đổi file {file_index + 1}/{len(saved_inputs)}…")
            try:
                result = await run_in_threadpool(_convert_one, input_path, output_path, file_options)
                successes.append((result.output_path, output_name))
                report.append(
                    {
                        "file": original_name,
                        "status": "success",
                        "output": output_name,
                        "warnings": result.warnings,
                        "pages": result.page_count,
                        "sheets": result.sheet_count,
                        "tables": result.table_count,
                        "math": result.math_report(),
                    }
                )
            except ConversionError as exc:
                report.append({"file": original_name, "status": "error", "message": str(exc)})
            except Exception as exc:  # avoid leaking stack traces to the browser
                LOGGER.exception("Unexpected conversion failure for %s", original_name)
                report.append(
                    {
                        "file": original_name,
                        "status": "error",
                        "message": f"Lỗi không mong đợi khi chuyển đổi: {exc}",
                    }
                )
            PROGRESS.update(
                job_id,
                status="converting",
                percent=15 + ((file_index + 1) / len(saved_inputs)) * 77,
                message=f"Đã xử lý file {file_index + 1}/{len(saved_inputs)}.",
                current_file=original_name,
                completed_files=file_index + 1,
            )

        if not successes:
            messages = [str(item.get("message", "Không thể chuyển đổi.")) for item in report]
            raise ConversionError("\n".join(messages))

        headers = {
            "X-OCR2Word-Report": _encode_warnings(report),
            "Cache-Control": "no-store",
        }
        PROGRESS.update(job_id, status="packaging", percent=95, message="Đang chuẩn bị file tải xuống…")
        if len(files) == 1 and len(successes) == 1:
            output_path, output_name = successes[0]
            PROGRESS.update(
                job_id,
                status="complete",
                percent=100,
                message="Đã chuyển đổi xong. Đang gửi file về trình duyệt…",
                current_file=None,
            )
            media_type = (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                if output_format == "xlsx"
                else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            return FileResponse(
                output_path,
                media_type=media_type,
                filename=output_name,
                headers=headers,
                background=BackgroundTask(shutil.rmtree, work_dir, ignore_errors=True),
            )

        zip_path = work_dir / "ocr2word-ket-qua.zip"
        PROGRESS.update(job_id, status="packaging", percent=97, message="Đang đóng gói file kết quả…")
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for output_path, output_name in successes:
                archive.write(output_path, output_name)
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
        PROGRESS.update(
            job_id,
            status="complete",
            percent=100,
            message="Đã chuyển đổi xong. Đang gửi file ZIP về trình duyệt…",
            current_file=None,
        )
        return FileResponse(
            zip_path,
            media_type="application/zip",
            filename="ocr2word-ket-qua.zip",
            headers=headers,
            background=BackgroundTask(shutil.rmtree, work_dir, ignore_errors=True),
        )
    except ConversionError as exc:
        shutil.rmtree(work_dir, ignore_errors=True)
        PROGRESS.update(job_id, status="error", message=str(exc), current_file=None)
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except HTTPException as exc:
        shutil.rmtree(work_dir, ignore_errors=True)
        PROGRESS.update(job_id, status="error", message=str(exc.detail), current_file=None)
        raise
    except Exception as exc:
        shutil.rmtree(work_dir, ignore_errors=True)
        LOGGER.exception("Upload or conversion request failed")
        PROGRESS.update(
            job_id,
            status="error",
            message="Không thể xử lý yêu cầu chuyển đổi.",
            current_file=None,
        )
        raise HTTPException(status_code=500, detail="Không thể xử lý yêu cầu chuyển đổi.") from exc


def run() -> None:
    import uvicorn

    uvicorn.run("app.main:app", host="127.0.0.1", port=8000, reload=False)


if __name__ == "__main__":
    run()
