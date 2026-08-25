from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest
from docx import Document

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
    "ai_assist": False,
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
