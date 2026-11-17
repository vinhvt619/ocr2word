from __future__ import annotations

import io
import zipfile
from pathlib import Path

import openpyxl
import pymupdf
import pytest
from docx import Document

import app.main as app_main
from app.converters import ConversionError, ConversionResult
from app.services import serverless_jobs


DEFAULT_OPTIONS = {
    "language": "vie+eng",
    "ocr_mode": "auto",
    "dpi": 300,
    "page_breaks": True,
    "preserve_images": True,
    "preserve_layout": True,
    "detect_tables": True,
}


def _fake_convert(input_path: Path, output_path: Path, options) -> ConversionResult:
    options.report_progress(50, f"Đang xử lý {input_path.name}")
    document = Document()
    document.add_paragraph(f"Nội dung từ {input_path.name}")
    document.save(output_path)
    options.report_progress(100, "Hoàn tất")
    return ConversionResult(output_path=output_path, page_count=1, table_count=0)


def test_single_file_returns_editable_docx_and_progress(monkeypatch):
    monkeypatch.setattr(serverless_jobs, "_convert_one", _fake_convert)
    progress: list[tuple[int, str]] = []

    result = serverless_jobs.run_serverless_conversion(
        [{"name": "Tài liệu.pdf", "content": b"%PDF-1.7\nmock"}],
        DEFAULT_OPTIONS,
        progress_callback=lambda percent, message, *_args: progress.append((percent, message)),
    )

    assert result.filename == "Tài liệu-editable.docx"
    assert result.media_type.endswith("wordprocessingml.document")
    assert result.content.startswith(b"PK")
    assert result.report[0]["status"] == "success"
    assert progress[-1][0] == 100


def test_multiple_files_are_packaged_with_unique_names(monkeypatch):
    monkeypatch.setattr(serverless_jobs, "_convert_one", _fake_convert)

    result = serverless_jobs.run_serverless_conversion(
        [
            {"name": "same.pdf", "content": b"%PDF-1.7\nfirst"},
            {"name": "same.PDF", "content": b"%PDF-1.7\nsecond"},
        ],
        DEFAULT_OPTIONS,
    )

    assert result.filename == "ocr2word-ket-qua.zip"
    assert result.media_type == "application/zip"
    with zipfile.ZipFile(io.BytesIO(result.content)) as archive:
        names = archive.namelist()
    assert names == ["same-editable.docx", "same-editable-2.docx"]
    assert len(result.report) == 2


@pytest.mark.parametrize(
    ("payload", "message"),
    [
        ({"language": "fra"}, "Tùy chọn chuyển đổi không hợp lệ"),
        ({"dpi": 250}, "Tùy chọn chuyển đổi không hợp lệ"),
        ({"ocr_mode": "unknown"}, "Tùy chọn chuyển đổi không hợp lệ"),
    ],
)
def test_rejects_invalid_options(payload, message):
    with pytest.raises(ConversionError, match=message):
        serverless_jobs.run_serverless_conversion(
            [{"name": "sample.pdf", "content": b"%PDF-1.7\nmock"}],
            {**DEFAULT_OPTIONS, **payload},
        )


def test_rejects_spoofed_pdf_before_conversion(monkeypatch):
    called = False

    def should_not_run(*_args):
        nonlocal called
        called = True
        raise AssertionError("converter must not run")

    monkeypatch.setattr(serverless_jobs, "_convert_one", should_not_run)
    with pytest.raises(ConversionError, match="không phải PDF hợp lệ"):
        serverless_jobs.run_serverless_conversion(
            [{"name": "sample.pdf", "content": b"not-a-pdf"}],
            DEFAULT_OPTIONS,
        )
    assert called is False


def test_enforces_combined_modal_transport_limit(monkeypatch):
    monkeypatch.setattr(serverless_jobs, "MAX_TOTAL_BYTES", 20)
    with pytest.raises(ConversionError, match="Tổng dung lượng"):
        serverless_jobs.run_serverless_conversion(
            [
                {"name": "one.pdf", "content": b"%PDF-1.7\n12"},
                {"name": "two.pdf", "content": b"%PDF-1.7\n34"},
            ],
            DEFAULT_OPTIONS,
        )


def test_rejects_result_larger_than_modal_transport_limit(monkeypatch):
    monkeypatch.setattr(serverless_jobs, "_convert_one", _fake_convert)
    monkeypatch.setattr(serverless_jobs, "MAX_OUTPUT_BYTES", 10)

    with pytest.raises(ConversionError, match="File kết quả lớn hơn"):
        serverless_jobs.run_serverless_conversion(
            [{"name": "sample.pdf", "content": b"%PDF-1.7\nmock"}],
            DEFAULT_OPTIONS,
        )


def _sample_pdf_bytes() -> bytes:
    pdf = pymupdf.open()
    try:
        page = pdf.new_page(width=600, height=800)
        page.insert_text((50, 40), "MODAL PREFIXED PDF", fontsize=14)
        page.draw_rect(pymupdf.Rect(50, 70, 350, 170))
        page.draw_line((50, 120), (350, 120))
        page.draw_line((200, 70), (200, 170))
        page.insert_text((70, 100), "Ma SP")
        page.insert_text((220, 100), "So luong")
        page.insert_text((70, 150), "M01")
        page.insert_text((220, 150), "7")
        return pdf.tobytes()
    finally:
        pdf.close()


def test_run_serverless_conversion_accepts_prefixed_real_pdf() -> None:
    prefixed_pdf = b"\xef\xbb\xbf\n" + _sample_pdf_bytes()

    docx_result = serverless_jobs.run_serverless_conversion(
        [{"name": "prefixed.pdf", "content": prefixed_pdf}],
        DEFAULT_OPTIONS,
    )
    assert docx_result.filename == "prefixed-editable.docx"
    doc = Document(io.BytesIO(docx_result.content))
    text = "\n".join(
        [p.text for p in doc.paragraphs]
        + [cell.text for table in doc.tables for row in table.rows for cell in row.cells]
    )
    assert "MODAL PREFIXED PDF" in text
    assert "M01" in text

    xlsx_result = serverless_jobs.run_serverless_conversion(
        [{"name": "prefixed.pdf", "content": prefixed_pdf}],
        {**DEFAULT_OPTIONS, "output_format": "xlsx"},
    )
    assert xlsx_result.filename == "prefixed-editable.xlsx"
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_result.content))
    assert "Bảng 1" in wb.sheetnames
    assert wb["Bảng 1"].cell(row=2, column=1).value == "M01"
    assert wb["Bảng 1"].cell(row=2, column=2).value == 7


@pytest.mark.parametrize(
    "payload",
    [
        b"<html><body>%PDF-1.7 error page</body></html>",
        (
            b"<html><body>%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
            b"2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n</body></html>"
        ),
    ],
)
def test_rejects_html_with_pdf_marker_before_conversion(monkeypatch, payload: bytes) -> None:
    called = False

    def should_not_run(*_args):
        nonlocal called
        called = True
        raise AssertionError("converter must not run")

    monkeypatch.setattr(serverless_jobs, "_convert_one", should_not_run)
    with pytest.raises(ConversionError, match="không phải PDF hợp lệ"):
        serverless_jobs.run_serverless_conversion(
            [{"name": "spoofed.pdf", "content": payload}],
            DEFAULT_OPTIONS,
        )
    assert called is False


def test_prefixed_protected_pdf_matches_normal_protected_pdf_behavior(tmp_path: Path) -> None:
    enc_doc = pymupdf.open()
    try:
        enc_doc.new_page().insert_text((72, 72), "Protected")
        normal_bytes = enc_doc.tobytes(
            encryption=pymupdf.PDF_ENCRYPT_AES_256,
            user_pw="secret",
            owner_pw="owner",
        )
    finally:
        enc_doc.close()
    prefixed_bytes = b"\xef\xbb\xbf\n" + normal_bytes

    normal_path = tmp_path / "normal-protected.pdf"
    prefixed_path = tmp_path / "prefixed-protected.pdf"
    normal_path.write_bytes(normal_bytes)
    prefixed_path.write_bytes(prefixed_bytes)

    # Both validation wrappers accept password-protected PDFs so the downstream
    # converter handles password/encryption failure uniformly.
    app_main._validate_signature(normal_path, ".pdf")
    app_main._validate_signature(prefixed_path, ".pdf")
    serverless_jobs._validate_signature(normal_path, ".pdf")
    serverless_jobs._validate_signature(prefixed_path, ".pdf")

    with pytest.raises(ConversionError) as normal_exc:
        serverless_jobs.run_serverless_conversion(
            [{"name": "normal-protected.pdf", "content": normal_bytes}],
            DEFAULT_OPTIONS,
        )
    with pytest.raises(ConversionError) as prefixed_exc:
        serverless_jobs.run_serverless_conversion(
            [{"name": "prefixed-protected.pdf", "content": prefixed_bytes}],
            DEFAULT_OPTIONS,
        )

    assert "không phải PDF hợp lệ" not in str(normal_exc.value)
    assert "không phải PDF hợp lệ" not in str(prefixed_exc.value)
    assert str(prefixed_exc.value) == str(normal_exc.value)
