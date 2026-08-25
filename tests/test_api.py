from __future__ import annotations

import base64
import io
import json
from pathlib import Path

import httpx
import pytest
from openpyxl import Workbook

from app.main import app
from app.services.ai_assist import AiAssistResult, AiEngineInfo


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.mark.anyio
async def test_home_page_and_status() -> None:
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        home = await client.get("/")
        assert home.status_code == 200
        assert "OCR2Word" in home.text
        assert "Hỗ trợ với AI" in home.text

        status = await client.get("/api/status")
        assert status.status_code == 200
        payload = status.json()
        assert payload["app"] == "OCR2Word"
        assert ".pdf" in payload["formats"]
        assert ".png" in payload["formats"]
        assert ".xlsx" in payload["formats"]
        assert "ocr" in payload
        assert "ai" in payload
        assert payload["runtime"]["python"]
        assert payload["runtime"]["system"]
        assert payload["ai"]["model"]
        assert "api_key" not in payload["ai"]


@pytest.mark.anyio
async def test_rejects_unsupported_upload() -> None:
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.post(
            "/api/convert",
            files={"files": ("notes.txt", b"hello", "text/plain")},
        )
        assert response.status_code == 422


@pytest.mark.anyio
async def test_excel_conversion_reports_complete_progress() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet.append(["Name", "Value"])
    worksheet.append(["Alpha", 42])
    payload = io.BytesIO()
    workbook.save(payload)
    payload.seek(0)
    job_id = "testexcelprogress2026"

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.post(
            "/api/convert",
            files={
                "files": (
                    "sample.xlsx",
                    payload.getvalue(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
            data={"job_id": job_id},
        )
        assert response.status_code == 200
        assert response.headers["content-type"].startswith(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        progress = await client.get(f"/api/progress/{job_id}")
        assert progress.status_code == 200
        snapshot = progress.json()
        assert snapshot["status"] == "complete"
        assert snapshot["percent"] == 100
        assert snapshot["completed_files"] == 1


@pytest.mark.anyio
async def test_ai_option_falls_back_when_gemini_key_is_missing(monkeypatch) -> None:
    workbook = Workbook()
    payload = io.BytesIO()
    workbook.save(payload)
    monkeypatch.setattr(
        "app.main.get_ai_engine_info",
        lambda: AiEngineInfo(
            available=False,
            installed=False,
            server_running=False,
            model_available=False,
            model="gemini-3.7-flash",
            provider="Google Gemini API",
            endpoint="https://generativelanguage.googleapis.com/v1beta/openai",
            requires_api_key=True,
            message="Chưa có Gemini API key. Hãy chạy .\\setup-gemini.ps1 một lần.",
        ),
    )

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.post(
            "/api/convert",
            files={
                "files": (
                    "sample.xlsx",
                    payload.getvalue(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
            data={"ai_assist": "true"},
        )

    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    encoded_report = response.headers["x-ocr2word-report"]
    report = json.loads(base64.urlsafe_b64decode(encoded_report).decode("utf-8"))
    assert report[0]["status"] == "success"
    assert "setup-gemini.ps1" in " ".join(report[0]["warnings"])
    assert "chưa hiệu đính" in " ".join(report[0]["warnings"])


@pytest.mark.anyio
async def test_ai_option_runs_after_conversion_and_is_reported(monkeypatch) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["Nội dung"])
    worksheet.append(["Sốhóa đơn"])
    payload = io.BytesIO()
    workbook.save(payload)
    calls: list[dict[str, object]] = []

    monkeypatch.setattr(
        "app.main.get_ai_engine_info",
        lambda: AiEngineInfo(
            True,
            True,
            True,
            True,
            "gemini-3.7-flash",
            "Google Gemini API",
            "https://generativelanguage.googleapis.com/v1beta/openai",
            True,
            "Sẵn sàng",
        ),
    )

    def fake_ai(path, *, model, progress_callback):
        calls.append({"path": path, "model": model})
        progress_callback(100, "AI hoàn tất")
        return AiAssistResult(
            model=model,
            processed_paragraphs=2,
            changed_paragraphs=1,
            rejected_changes=0,
            skipped_paragraphs=0,
            batch_count=1,
        )

    monkeypatch.setattr("app.main.apply_ai_assistance_to_docx", fake_ai)

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.post(
            "/api/convert",
            files={
                "files": (
                    "sample.xlsx",
                    payload.getvalue(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
            data={
                "ai_assist": "true",
                "job_id": "testaiassist2026",
            },
        )

    assert response.status_code == 200
    assert "editable-ai.docx" in response.headers["content-disposition"]
    assert len(calls) == 1
    assert calls[0]["model"] == "gemini-3.7-flash"
    assert Path(calls[0]["path"]).suffix == ".docx"
    encoded_report = response.headers["x-ocr2word-report"]
    report = json.loads(base64.urlsafe_b64decode(encoded_report).decode("utf-8"))
    assert report[0]["ai"]["processed_paragraphs"] == 2
    assert report[0]["ai"]["changed_paragraphs"] == 1


@pytest.mark.anyio
async def test_vercel_worker_requires_secret_and_accepts_private_blob_urls(monkeypatch) -> None:
    monkeypatch.setenv("OCR2WORD_INTERNAL_SECRET", "test-worker-secret")
    monkeypatch.setattr(
        "app.main._run_vercel_conversion",
        lambda payload: {"report": [], "output_name": payload.output.name},
    )
    body = {
        "job_id": "a" * 32,
        "files": [
            {
                "name": "sample.pdf",
                "size": 8,
                "content_type": "application/pdf",
                "download_url": "https://store.private.blob.vercel-storage.com/input.pdf?token=test",
            }
        ],
        "options": {
            "language": "vie",
            "ocr_mode": "auto",
            "dpi": 300,
            "page_breaks": True,
            "preserve_images": True,
            "preserve_layout": True,
            "detect_tables": True,
            "ai_assist": False,
        },
        "output": {
            "name": "sample-editable.docx",
            "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "upload_url": "https://vercel.com/api/blob/?pathname=result.docx&token=test",
        },
    }
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        unauthorized = await client.post("/internal/convert", json=body)
        accepted = await client.post(
            "/internal/convert",
            json=body,
            headers={"Authorization": "Bearer test-worker-secret"},
        )

    assert unauthorized.status_code == 401
    assert accepted.status_code == 200
    assert accepted.json()["output_name"] == "sample-editable.docx"
