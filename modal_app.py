from __future__ import annotations

import base64
import hmac
import json
import os
import platform
import re
import secrets
import time
import urllib.parse
from pathlib import Path
from typing import Any

import modal
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse, Response
from fastapi.staticfiles import StaticFiles


PROJECT_ROOT = Path(__file__).resolve().parent
STATIC_DIR = PROJECT_ROOT / "app" / "static"
REQUIREMENTS_PATH = PROJECT_ROOT / "requirements.txt"
MATH_LOCAL_REQUIREMENTS_PATH = PROJECT_ROOT / "requirements-math-local.txt"
APP_NAME = "ocr2word"
SECRET_NAME = "ocr2word-secrets"
JOB_STORE_NAME = "ocr2word-jobs"
JOB_TTL_SECONDS = 7 * 24 * 60 * 60
JOB_ID_PATTERN = re.compile(r"^[a-f0-9]{32}$")
PDF_EXTENSIONS = {".pdf"}
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
ALLOWED_EXTENSIONS = PDF_EXTENSIONS | EXCEL_EXTENSIONS | IMAGE_EXTENSIONS
MAX_FILES = 20
MAX_FILE_BYTES = 80 * 1024 * 1024
MAX_TOTAL_BYTES = 90 * 1024 * 1024
READ_CHUNK_SIZE = 1024 * 1024


app = modal.App(APP_NAME)
runtime_secret = modal.Secret.from_name(SECRET_NAME)
jobs = modal.Dict.from_name(JOB_STORE_NAME, create_if_missing=True)

worker_image = (
    modal.Image.debian_slim(python_version="3.12")
    .apt_install(
        "tesseract-ocr",
        "tesseract-ocr-eng",
        "tesseract-ocr-vie",
        "libglib2.0-0",
        "libgomp1",
    )
    .pip_install_from_requirements(REQUIREMENTS_PATH)
    .run_commands(
        "pip install --no-cache-dir --index-url https://download.pytorch.org/whl/cpu 'torch>=2.2,<3' 'torchvision>=0.17,<1'",
        "pip install --no-cache-dir --no-deps pix2tex==0.1.4",
    )
    .pip_install_from_requirements(MATH_LOCAL_REQUIREMENTS_PATH)
    .run_commands(
        # Fail the image build on missing/incompatible inference imports rather
        # than discovering a workstation-only dependency at the first upload.
        "python -c \"from transformers import PreTrainedTokenizerFast; from pix2tex.models import get_model; from pix2tex.utils import parse_args, post_process, token2str\""
    )
    .env(
        {
            "NO_ALBUMENTATIONS_UPDATE": "1",
            "OCR2WORD_MATH_MODEL_DIR": "/opt/ocr2word-models/pix2tex",
        }
    )
    # The following build command imports app to download verified weights.
    # Its sources must exist at build time, not only at container startup.
    .add_local_python_source("app", copy=True)
    .run_commands(
        # Modal copies Python sources under /root. Unlike function execution,
        # a plain image-builder Python process does not add it to sys.path.
        "PYTHONPATH=/root python -c \"from pathlib import Path; from app.converters.math_local import download_and_verify_local_math_checkpoints; download_and_verify_local_math_checkpoints(Path('/opt/ocr2word-models/pix2tex'))\""
    )
)

web_image = (
    modal.Image.debian_slim(python_version="3.12")
    .uv_pip_install(
        "fastapi>=0.115,<1",
        "python-multipart>=0.0.20,<1",
    )
    .add_local_python_source("app")
    .add_local_dir(STATIC_DIR, remote_path="/assets")
)


def _read_job(job_id: str) -> dict[str, Any] | None:
    value = jobs.get(job_id)
    return dict(value) if isinstance(value, dict) else None


def _update_job(job_id: str, **changes: Any) -> dict[str, Any]:
    state = _read_job(job_id) or {
        "job_id": job_id,
        "created_at": time.time(),
        "expires_at": time.time() + JOB_TTL_SECONDS,
    }
    if "percent" in changes:
        requested = max(0, min(100, round(float(changes["percent"]))))
        changes["percent"] = max(int(state.get("percent", 0)), requested)
    state.update(changes)
    state["updated_at"] = time.time()
    jobs.put(job_id, state)
    return state


@app.function(
    image=worker_image,
    secrets=[runtime_secret],
    cpu=(1.0, 1.0),
    memory=(4096, 6144),
    timeout=35 * 60,
    retries=1,
    min_containers=0,
    max_containers=1,
    buffer_containers=0,
    scaledown_window=60,
)
def ocr_job(
    job_id: str,
    files: list[dict[str, object]],
    options: dict[str, object],
) -> dict[str, object]:
    from app.converters import ConversionError
    from app.services.serverless_jobs import run_serverless_conversion

    last_percent = -1
    last_update_at = 0.0

    def progress(
        percent: int,
        message: str,
        current_file: str | None,
        completed_files: int,
        status: str,
    ) -> None:
        nonlocal last_percent, last_update_at
        now = time.monotonic()
        # Keep the public state at 99% until Modal has persisted the Function
        # result. The web poller promotes it to complete after get(timeout=0)
        # can read those bytes.
        effective_status = "packaging" if status == "complete" else status
        effective_percent = min(99, percent) if status == "complete" else percent
        force = effective_status in {"error", "packaging"} or effective_percent >= 99
        if not force and effective_percent <= last_percent and now - last_update_at < 0.8:
            return
        if not force and effective_percent < last_percent + 1 and now - last_update_at < 0.8:
            return
        _update_job(
            job_id,
            status=effective_status,
            percent=effective_percent,
            message=message,
            current_file=current_file,
            completed_files=completed_files,
        )
        last_percent = max(last_percent, effective_percent)
        last_update_at = now

    try:
        _update_job(
            job_id,
            status="processing",
            percent=20,
            message="Modal đang khởi động container OCR…",
            started_at=time.time(),
        )
        output = run_serverless_conversion(
            files,
            options,
            progress_callback=progress,
        )
        payload = output.to_modal_payload()
        _update_job(
            job_id,
            status="packaging",
            percent=99,
            message="Đang lưu kết quả an toàn trên Modal…",
            current_file=None,
            completed_files=len(files),
        )
        return payload
    except ConversionError as exc:
        _update_job(
            job_id,
            status="error",
            message=str(exc),
            error=str(exc),
            current_file=None,
            finished_at=time.time(),
        )
        raise
    except Exception as exc:
        message = "Container Modal không thể xử lý tài liệu."
        _update_job(
            job_id,
            status="error",
            message=message,
            error=message,
            current_file=None,
            finished_at=time.time(),
        )
        raise RuntimeError(message) from exc


def _basic_auth_response():
    return PlainTextResponse(
        "Cần đăng nhập để sử dụng OCR2Word.",
        status_code=401,
        headers={
            "WWW-Authenticate": 'Basic realm="OCR2Word gia dinh", charset="UTF-8"',
            "Cache-Control": "no-store",
        },
    )


def _authorized(authorization: str | None) -> bool:
    username = os.environ.get("OCR2WORD_BASIC_USER", "ocr2word")
    password = os.environ.get("OCR2WORD_BASIC_PASSWORD", "")
    if not password:
        return True
    if not authorization:
        return False
    token = base64.b64encode(f"{username}:{password}".encode("utf-8")).decode("ascii")
    return hmac.compare_digest(authorization, f"Basic {token}")


def create_web_app():
    assets_dir = Path("/assets") if Path("/assets").is_dir() else STATIC_DIR
    web_app = FastAPI(
        title="OCR2Word Modal",
        docs_url=None,
        redoc_url=None,
        openapi_url=None,
    )

    @web_app.middleware("http")
    async def protect_family_site(request: Request, call_next):
        if request.url.path != "/healthz" and not _authorized(request.headers.get("authorization")):
            return _basic_auth_response()
        response = await call_next(request)
        response.headers.setdefault("X-Content-Type-Options", "nosniff")
        response.headers.setdefault("Referrer-Policy", "same-origin")
        if request.url.path.startswith("/api/"):
            response.headers.setdefault("Cache-Control", "no-store")
        return response

    @web_app.get("/healthz", include_in_schema=False)
    async def healthz() -> dict[str, bool]:
        return {"ok": True}

    @web_app.get("/", include_in_schema=False)
    async def index() -> FileResponse:
        return FileResponse(assets_dir / "index.html", media_type="text/html; charset=utf-8")

    @web_app.get("/api/status")
    async def status() -> dict[str, object]:
        from app import __version__

        return {
            "app": "OCR2Word",
            "version": __version__,
            "ocr": {
                "available": True,
                "engine": "Tesseract OCR",
                "message": "Tesseract OCR chạy trong container Modal theo yêu cầu.",
                "languages": ["vie", "eng"],
            },
            "math_ocr": {
                "default_when_enabled": "local_hybrid",
                "engines": {
                    "local_hybrid": {
                        "available": True,
                        "configured": True,
                        "weights_ready": True,
                        "uses_local_model": True,
                        "model_name": "pix2tex",
                        "revision": "v0.0.1",
                        "reason": "Mô hình công thức cục bộ đã đóng gói sẵn trong Modal worker container.",
                    },
                    "rules": {
                        "available": True,
                        "configured": True,
                        "weights_ready": True,
                        "uses_local_model": False,
                    },
                },
            },
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
                "container": True,
                "modal": True,
                "vercel": False,
            },
            "deployment": {
                "mode": "modal",
                "available": True,
                "message": "Modal tự bật container khi có lượt chuyển đổi và tự tắt khi rảnh.",
                "result_ttl_hours": 168,
            },
        }

    @web_app.post("/api/modal/jobs", status_code=202)
    async def create_job(
        files: list[UploadFile] = File(description="PDF, hình ảnh hoặc Excel"),
        options: str = Form(default="{}"),
    ) -> dict[str, object]:
        if not files:
            raise HTTPException(status_code=400, detail="Hãy chọn ít nhất một file.")
        if len(files) > MAX_FILES:
            raise HTTPException(status_code=400, detail=f"Mỗi lượt chỉ nhận tối đa {MAX_FILES} file.")
        try:
            option_payload = json.loads(options)
            if not isinstance(option_payload, dict):
                raise ValueError
        except (json.JSONDecodeError, ValueError) as exc:
            raise HTTPException(status_code=400, detail="Tùy chọn chuyển đổi không hợp lệ.") from exc

        output_format = str(option_payload.get("output_format", "docx")).lower()
        if output_format not in {"docx", "xlsx"}:
            raise HTTPException(status_code=400, detail="Định dạng xuất không hợp lệ (chỉ hỗ trợ docx hoặc xlsx).")
        if output_format == "xlsx":
            if option_payload.get("math_ocr"):
                raise HTTPException(status_code=400, detail="Chuyển đổi sang Excel (.xlsx) không hỗ trợ tính năng nhận diện công thức toán.")

        uploaded: list[dict[str, object]] = []
        total_size = 0
        for upload in files:
            original_name = upload.filename or "tai-lieu"
            suffix = Path(original_name).suffix.lower()
            if output_format == "xlsx" and suffix in EXCEL_EXTENSIONS:
                raise HTTPException(
                    status_code=400,
                    detail="Chuyển đổi sang Excel (.xlsx) chỉ nhận file PDF hoặc hình ảnh (không nhận file Excel làm đầu vào).",
                )
            if suffix not in ALLOWED_EXTENSIONS:
                raise HTTPException(
                    status_code=400,
                    detail=f"'{original_name}' không thuộc định dạng được hỗ trợ.",
                )
            content = bytearray()
            while chunk := await upload.read(READ_CHUNK_SIZE):
                content.extend(chunk)
                if len(content) > MAX_FILE_BYTES:
                    raise HTTPException(
                        status_code=413,
                        detail=f"'{original_name}' lớn hơn giới hạn 80 MB của Modal.",
                    )
                if total_size + len(content) > MAX_TOTAL_BYTES:
                    raise HTTPException(
                        status_code=413,
                        detail="Tổng dung lượng tải lên lớn hơn giới hạn 90 MB của Modal.",
                    )
            await upload.close()
            if not content:
                raise HTTPException(status_code=400, detail=f"'{original_name}' là file rỗng.")
            total_size += len(content)
            uploaded.append({"name": original_name, "content": bytes(content)})

        job_id = secrets.token_hex(16)
        now = time.time()
        await jobs.put.aio(
            job_id,
            {
                "job_id": job_id,
                "status": "queued",
                "percent": 20,
                "message": "Đã tải tài liệu lên. Đang xếp hàng xử lý…",
                "current_file": None,
                "completed_files": 0,
                "total_files": len(uploaded),
                "created_at": now,
                "updated_at": now,
                "expires_at": now + JOB_TTL_SECONDS,
            },
        )
        try:
            call = await ocr_job.spawn.aio(job_id, uploaded, option_payload)
            await jobs.put.aio(f"call:{job_id}", call.object_id)
        except Exception as exc:
            await jobs.put.aio(
                job_id,
                {
                    "job_id": job_id,
                    "status": "error",
                    "percent": 20,
                    "message": "Không thể đưa tài liệu vào hàng đợi Modal.",
                    "error": "Không thể đưa tài liệu vào hàng đợi Modal.",
                    "current_file": None,
                    "completed_files": 0,
                    "total_files": len(uploaded),
                    "created_at": now,
                    "updated_at": time.time(),
                    "expires_at": now + JOB_TTL_SECONDS,
                },
            )
            raise HTTPException(
                status_code=503,
                detail="Không thể đưa tài liệu vào hàng đợi Modal.",
            ) from exc
        return {"job_id": job_id, "status": "queued"}

    async def get_job_or_404(job_id: str) -> dict[str, Any]:
        if not JOB_ID_PATTERN.fullmatch(job_id):
            raise HTTPException(status_code=404, detail="Không tìm thấy lượt chuyển đổi.")
        state = await jobs.get.aio(job_id)
        if not isinstance(state, dict):
            raise HTTPException(status_code=404, detail="Không tìm thấy lượt chuyển đổi.")
        if time.time() > float(state.get("expires_at", 0)):
            await jobs.pop.aio(job_id, None)
            await jobs.pop.aio(f"call:{job_id}", None)
            raise HTTPException(status_code=410, detail="Kết quả đã hết hạn sau 7 ngày.")
        return dict(state)

    async def finalize_job_if_ready(job_id: str, state: dict[str, Any]) -> dict[str, Any]:
        if state.get("status") in {"complete", "error"}:
            return state
        call_id = await jobs.get.aio(f"call:{job_id}")
        if not isinstance(call_id, str):
            return state
        function_call = modal.FunctionCall.from_id(call_id)
        try:
            result = await function_call.get.aio(timeout=0)
        except TimeoutError:
            return state
        except modal.exception.OutputExpiredError:
            failed = {
                **state,
                "status": "error",
                "message": "Kết quả đã hết hạn trên Modal.",
                "error": "Kết quả đã hết hạn trên Modal.",
                "updated_at": time.time(),
                "finished_at": time.time(),
            }
            await jobs.put.aio(job_id, failed)
            return failed
        except Exception:
            refreshed = await jobs.get.aio(job_id)
            if isinstance(refreshed, dict) and refreshed.get("status") == "error":
                return dict(refreshed)
            failed = {
                **state,
                "status": "error",
                "message": "Container Modal không thể xử lý tài liệu.",
                "error": "Container Modal không thể xử lý tài liệu.",
                "updated_at": time.time(),
                "finished_at": time.time(),
            }
            await jobs.put.aio(job_id, failed)
            return failed
        if not isinstance(result, dict) or not isinstance(result.get("content"), bytes):
            failed = {
                **state,
                "status": "error",
                "message": "Kết quả Modal không hợp lệ.",
                "error": "Kết quả Modal không hợp lệ.",
                "updated_at": time.time(),
                "finished_at": time.time(),
            }
            await jobs.put.aio(job_id, failed)
            return failed
        completed = {
            **state,
            "status": "complete",
            "percent": 100,
            "message": "Đã chuyển đổi xong. Kết quả sẵn sàng để tải xuống.",
            "current_file": None,
            "output_name": str(result.get("filename") or "tai-lieu-editable.docx"),
            "media_type": str(
                result.get("media_type")
                or "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
            "report": result.get("report") if isinstance(result.get("report"), list) else [],
            "updated_at": time.time(),
            "finished_at": time.time(),
        }
        await jobs.put.aio(job_id, completed)
        return completed

    @web_app.get("/api/modal/jobs/{job_id}")
    async def job_status(job_id: str, request: Request) -> dict[str, object]:
        state = await get_job_or_404(job_id)
        state = await finalize_job_if_ready(job_id, state)
        public_state = {
            key: value
            for key, value in state.items()
            if key not in {"created_at", "updated_at", "expires_at"}
        }
        if state.get("status") == "complete":
            public_state["download_url"] = str(
                request.url_for("download_job", job_id=job_id)
            )
            public_state["download_name"] = state.get("output_name")
        return public_state

    @web_app.get("/api/modal/jobs/{job_id}/download", name="download_job")
    async def download_job(job_id: str) -> Response:
        state = await get_job_or_404(job_id)
        state = await finalize_job_if_ready(job_id, state)
        if state.get("status") == "error":
            raise HTTPException(
                status_code=422,
                detail=str(state.get("error") or state.get("message") or "Chuyển đổi không thành công."),
            )
        if state.get("status") != "complete":
            raise HTTPException(status_code=202, detail="Tài liệu vẫn đang được xử lý.")
        call_id = await jobs.get.aio(f"call:{job_id}")
        if not isinstance(call_id, str):
            raise HTTPException(status_code=410, detail="Kết quả không còn khả dụng.")
        function_call = modal.FunctionCall.from_id(call_id)
        try:
            result = await function_call.get.aio(timeout=0)
        except modal.exception.OutputExpiredError as exc:
            raise HTTPException(status_code=410, detail="Kết quả đã hết hạn trên Modal.") from exc
        except TimeoutError as exc:
            raise HTTPException(status_code=202, detail="Tài liệu vẫn đang được xử lý.") from exc
        except Exception as exc:
            raise HTTPException(status_code=422, detail="Không thể tải kết quả chuyển đổi.") from exc
        if not isinstance(result, dict) or not isinstance(result.get("content"), bytes):
            raise HTTPException(status_code=500, detail="Kết quả Modal không hợp lệ.")
        filename = str(result.get("filename") or "tai-lieu-editable.docx")
        media_type = str(
            result.get("media_type")
            or "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        ascii_name = re.sub(r"[^A-Za-z0-9._-]", "-", filename) or "ocr2word-result"
        encoded_name = urllib.parse.quote(filename, safe="")
        return Response(
            content=result["content"],
            media_type=media_type,
            headers={
                "Content-Disposition": (
                    f'attachment; filename="{ascii_name}"; filename*=UTF-8\'\'{encoded_name}'
                ),
                "Cache-Control": "no-store",
            },
        )

    web_app.mount("/static", StaticFiles(directory=assets_dir), name="static")
    return web_app


@app.function(
    image=web_image,
    secrets=[runtime_secret],
    cpu=(0.25, 0.5),
    memory=(1024, 2048),
    timeout=5 * 60,
    min_containers=0,
    max_containers=1,
    buffer_containers=0,
    scaledown_window=60,
)
@modal.asgi_app()
@modal.concurrent(max_inputs=20)
def web():
    return create_web_app()


@app.local_entrypoint()
def smoke_test(
    file: str,
    output_dir: str = "output/modal-smoke",
    math_ocr: bool = False,
    math_engine: str = "local_hybrid",
) -> None:
    source = Path(file).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if source.stat().st_size > MAX_FILE_BYTES:
        raise ValueError("File kiểm thử lớn hơn giới hạn 80 MB của Modal.")
    job_id = secrets.token_hex(16)
    now = time.time()
    jobs.put(
        job_id,
        {
            "job_id": job_id,
            "status": "queued",
            "percent": 0,
            "message": "Đang chạy smoke test…",
            "current_file": source.name,
            "completed_files": 0,
            "total_files": 1,
            "created_at": now,
            "updated_at": now,
            "expires_at": now + JOB_TTL_SECONDS,
        },
    )
    result = ocr_job.remote(
        job_id,
        [{"name": source.name, "content": source.read_bytes()}],
        {
            "language": "vie+eng",
            "ocr_mode": "auto",
            "dpi": 300,
            "page_breaks": True,
            "preserve_images": True,
            "preserve_layout": True,
            "detect_tables": True,
            "math_ocr": math_ocr,
            "math_engine": math_engine,
        },
    )
    target_dir = Path(output_dir).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / str(result["filename"])
    target.write_bytes(result["content"])
    print(target)
