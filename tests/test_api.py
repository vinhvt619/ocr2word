from __future__ import annotations

import base64
import io
import json
from pathlib import Path

import httpx
import pytest
from openpyxl import Workbook

from app.main import app


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
        assert "Nhận diện công thức toán học" in home.text

        status = await client.get("/api/status")
        assert status.status_code == 200
        payload = status.json()
        assert payload["app"] == "OCR2Word"
        assert ".pdf" in payload["formats"]
        assert ".png" in payload["formats"]
        assert ".xlsx" in payload["formats"]
        assert "ocr" in payload
        assert payload["math_ocr"]["default_when_enabled"] == "local_hybrid"
        assert payload["runtime"]["python"]
        assert payload["runtime"]["system"]


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
