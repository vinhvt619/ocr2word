from __future__ import annotations

import base64
import io
import json
from pathlib import Path

import httpx
import openpyxl
import pymupdf
import pytest
from docx import Document
from openpyxl import Workbook

import app.main as app_main
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


def _sample_pdf_bytes() -> bytes:
    """Create a real one-page digital PDF containing text and a 2x2 table."""
    pdf = pymupdf.open()
    try:
        page = pdf.new_page(width=600, height=800)
        page.insert_text((50, 40), "BAO CAO 2026", fontsize=14)
        page.draw_rect(pymupdf.Rect(50, 70, 350, 170))
        page.draw_line((50, 120), (350, 120))
        page.draw_line((200, 70), (200, 170))
        page.insert_text((70, 100), "Ma SP")
        page.insert_text((220, 100), "So luong")
        page.insert_text((70, 150), "SP01")
        page.insert_text((220, 150), "25")
        return pdf.tobytes()
    finally:
        pdf.close()


def _docx_text(payload: bytes) -> str:
    doc = Document(io.BytesIO(payload))
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    return "\n".join(parts)


@pytest.mark.anyio
async def test_prefixed_real_pdf_converts_to_xlsx_docx_and_analyzes(monkeypatch) -> None:
    raw_pdf = _sample_pdf_bytes()
    prefixed_pdf = b"\xef\xbb\xbf\n" + raw_pdf

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        xlsx_response = await client.post(
            "/api/convert",
            files={"files": ("prefixed.pdf", prefixed_pdf, "application/pdf")},
            data={"output_format": "xlsx"},
        )
        assert xlsx_response.status_code == 200
        assert xlsx_response.headers["content-type"].startswith(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_response.content))
        assert "Bảng 1" in wb.sheetnames
        ws = wb["Bảng 1"]
        assert ws.cell(row=1, column=1).value == "Ma SP"
        assert ws.cell(row=2, column=1).value == "SP01"
        assert ws.cell(row=2, column=2).value == 25

        docx_response = await client.post(
            "/api/convert",
            files={"files": ("prefixed.pdf", prefixed_pdf, "application/pdf")},
            data={"output_format": "docx"},
        )
        assert docx_response.status_code == 200
        assert docx_response.headers["content-type"].startswith(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        docx_content = _docx_text(docx_response.content)
        assert "BAO CAO 2026" in docx_content
        assert "SP01" in docx_content

        normal_response = await client.post(
            "/api/convert",
            files={"files": ("normal.pdf", raw_pdf, "application/pdf")},
            data={"output_format": "docx"},
        )
        assert normal_response.status_code == 200
        assert "BAO CAO 2026" in _docx_text(normal_response.content)

        monkeypatch.setenv("OCR2WORD_INTERNAL_SECRET", "test-analyze-secret")
        monkeypatch.setattr(
            app_main,
            "_download_vercel_input",
            lambda source, destination, **_kwargs: (destination.write_bytes(prefixed_pdf), len(prefixed_pdf))[1],
        )
        analysis_response = await client.post(
            "/internal/analyze-pdf",
            json={
                "job_id": "b" * 32,
                "file": {
                    "name": "prefixed.pdf",
                    "size": len(prefixed_pdf),
                    "content_type": "application/pdf",
                    "download_url": "https://store.private.blob.vercel-storage.com/prefixed.pdf?token=test",
                },
                "options": {
                    "language": "vie",
                    "ocr_mode": "auto",
                    "dpi": 300,
                    "page_breaks": True,
                    "preserve_images": True,
                    "preserve_layout": True,
                    "detect_tables": True,
                },
            },
            headers={"Authorization": "Bearer test-analyze-secret"},
        )
        assert analysis_response.status_code == 200
        assert analysis_response.json()["page_count"] == 1


@pytest.mark.anyio
@pytest.mark.parametrize(
    "payload",
    [
        b"<!doctype html><html><body>Error 404</body></html>",
        b"<html><body>Error %PDF-1.7 not downloadable</body></html>",
        (
            b"<html><body>%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
            b"2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n</body></html>"
        ),
        b"\xef\xbb\xbf\n%PDF-1.7\n1 0 obj\n<< /Type /Catalog >>",
        b" " * 1050 + _sample_pdf_bytes(),
    ],
)
async def test_rejects_invalid_and_spoofed_pdf_uploads(monkeypatch, payload: bytes) -> None:
    called = False

    def fail_if_converted(*_args):
        nonlocal called
        called = True
        raise AssertionError("converter must not run for invalid PDFs")

    monkeypatch.setattr(app_main, "_convert_one", fail_if_converted)
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.post(
            "/api/convert",
            files={"files": ("spoofed.pdf", payload, "application/pdf")},
        )

    assert response.status_code == 422
    assert "File có đuôi .pdf nhưng nội dung không phải PDF hợp lệ." in response.json().get("detail", "")
    assert called is False


def _background_cv_pdf_bytes(*, with_bezier_and_pua: bool = False) -> bytes:
    from PIL import Image, ImageDraw

    bg = Image.new("RGB", (800, 1100), (242, 246, 252))
    ImageDraw.Draw(bg).rectangle([0, 0, 240, 1100], fill=(30, 58, 138))
    buf = io.BytesIO()
    bg.save(buf, format="PNG")
    bg_bytes = buf.getvalue()

    pdf = pymupdf.open()
    try:
        page = pdf.new_page(width=595, height=842)
        page.insert_image(pymupdf.Rect(2, 2, 593, 840), stream=bg_bytes)
        for index in range(4):
            page.insert_text(
                (40, 55 + index * 18),
                "hidden opacity zero ocr layer " * 2,
                fontsize=9,
                fill_opacity=0.0,
            )
        lines = [
            "NGUYEN VAN A - SENIOR SOFTWARE ENGINEER",
            "Email: nguyenvana@example.com | Phone: 0901234567 | Ho Chi Minh City",
            "SUMMARY: Full-stack engineer specializing in Python, FastAPI, React, and Cloud.",
            "EXPERIENCE: Senior Backend Developer at TechCorp Vietnam (2023 - Present)",
            "Architected high-throughput document conversion pipelines with PyMuPDF and OpenCV.",
            "PROJECTS: OCR2Word - Intelligent PDF and Image to Word/Excel Converter (2025)",
        ]
        for index, line in enumerate(lines):
            y_pos = 100 + index * 14 if (with_bezier_and_pua and index in {0, 1}) else 100 + index * 110
            page.insert_text((40, y_pos), line, fontsize=10.5)
        if with_bezier_and_pua:
            writer = pymupdf.TextWriter(page.rect)
            writer.append(
                (40, 740),
                "Developed scalable microservices \u0192 maintained high availability across regional clusters daily.",
                font=pymupdf.Font("cjk"),
                fontsize=10.5,
            )
            writer.write_text(page)
            for xref in range(1, pdf.xref_length()):
                if pdf.xref_is_stream(xref):
                    stream = pdf.xref_stream(xref)
                    if b"<0136> <0192>" in stream:
                        pdf.update_stream(xref, stream.replace(b"<0136> <0192>", b"<0136> <e001>"))
            shape = page.new_shape()
            shape.draw_line((60, 105.0), (75, 105.0))
            shape.draw_bezier((75, 105.0), (78, 111.0), (62, 111.0), (60, 105.0))
            shape.finish(fill=(0.2, 0.3, 0.5), color=None, closePath=True)
            shape.commit()
        return pdf.tobytes()
    finally:
        pdf.close()


@pytest.mark.anyio
@pytest.mark.parametrize("prefix", [b"", b"\xef\xbb\xbf\n"])
async def test_api_convert_auto_background_cv_pdf_preserves_text_and_drawings(
    monkeypatch: pytest.MonkeyPatch,
    prefix: bytes,
) -> None:
    def fail_if_ocr_called(*_args, **_kwargs):
        raise AssertionError("raster OCR must not be called for designed background-image CV PDFs")

    monkeypatch.setattr("app.converters.pdf._write_ocr_page", fail_if_ocr_called)
    payload = prefix + _background_cv_pdf_bytes()

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.post(
            "/api/convert",
            files={"files": ("cv.pdf", payload, "application/pdf")},
            data={"ocr_mode": "auto"},
        )

    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    doc = Document(io.BytesIO(response.content))
    extracted = _docx_text(response.content)
    assert "NGUYEN VAN A" in extracted
    assert "OCR2Word" in extracted
    assert len(doc.element.xpath(".//w:drawing")) >= 1


@pytest.mark.anyio
async def test_api_convert_auto_local_hybrid_background_cv_with_bezier_and_isolated_replacement_emits_no_equations(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def fail_if_ocr_called(*_args, **_kwargs):
        raise AssertionError("raster OCR must not be called for designed background-image CV PDFs")

    def fail_if_local_model_loaded():
        raise AssertionError("local_hybrid model must not load on non-math CV prose")

    monkeypatch.setattr("app.converters.pdf._write_ocr_page", fail_if_ocr_called)
    monkeypatch.setattr("app.converters.math_local.ensure_local_math_model_ready", fail_if_local_model_loaded)
    payload = _background_cv_pdf_bytes(with_bezier_and_pua=True)

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.post(
            "/api/convert",
            files={"files": ("cv-local-hybrid.pdf", payload, "application/pdf")},
            data={
                "ocr_mode": "auto",
                "math_ocr": "true",
                "math_engine": "local_hybrid",
            },
        )

    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    doc = Document(io.BytesIO(response.content))
    extracted = _docx_text(response.content)
    assert "NGUYEN VAN A" in extracted
    assert "OCR2Word" in extracted
    assert len(doc.element.xpath(".//m:oMath")) == 0
    assert len(doc.element.xpath(".//w:drawing")) >= 1
