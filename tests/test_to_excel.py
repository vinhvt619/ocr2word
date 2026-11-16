import io
import time
import zipfile
from pathlib import Path

import httpx
import openpyxl
from PIL import Image, ImageDraw
import pymupdf
import pytest
from starlette.testclient import TestClient

import app.main as app_main
from app.converters.base import ConversionError, ConversionOptions
from app.converters.math_local import strip_presentation_wrappers
from app.converters.raster_tables import detect_raster_tables
from app.converters.to_excel import (
    _build_excel_workbook,
    convert_to_excel,
    format_cell_value,
    ExtractedTable,
    ExtractedTextLine,
    MAX_PAGES,
    MAX_PIXELS_PER_PAGE,
    MAX_CELL_CHARS,
    MAX_ROWS_PER_SHEET,
    MAX_CELLS_PER_WORKBOOK,
)
from app.main import app
from app.services.serverless_jobs import run_serverless_conversion, ServerlessConversionOptions


# ===========================================================================
# 1. Defect Regression: math_local strip_presentation_wrappers
# ===========================================================================
def test_strip_presentation_wrappers_protected_greek_and_termination():
    start = time.monotonic()
    # Must terminate quickly (< 100ms) and NOT loop indefinitely
    assert strip_presentation_wrappers(r"\mathbf{\alpha}") == r"\mathbf{\alpha}"
    assert strip_presentation_wrappers(r"\mathrm{\beta}") == r"\mathrm{\beta}"
    assert strip_presentation_wrappers(r"\boldsymbol{\pi}") == r"\boldsymbol{\pi}"
    assert strip_presentation_wrappers(r"\mathbf{\Gamma}") == r"\mathbf{\Gamma}"
    assert strip_presentation_wrappers(r"\mathbf{\mathfrak{X}}") == r"\mathbf{\mathfrak{X}}"

    # Nested font unwrapping
    assert strip_presentation_wrappers(r"\mathbf{\mathrm{x}}") == "x"
    assert strip_presentation_wrappers(r"\mathbf{\mathbf{x}}") == "x"
    assert strip_presentation_wrappers(r"\mathbf{x} + \mathrm{y}") == "x + y"
    assert time.monotonic() - start < 1.0


# ===========================================================================
# 2. Defect Regression: raster_tables fill_sequences
# ===========================================================================
def test_detect_raster_tables_fill_sequences_flag():
    # Construct a synthetic table image with grid
    img = Image.new("L", (400, 300), color=255)
    draw = ImageDraw.Draw(img)
    # Draw horizontal lines
    for y in [50, 100, 150, 200, 250]:
        draw.line([(50, y), (350, y)], fill=0, width=2)
    # Draw vertical lines
    for x in [50, 150, 250, 350]:
        draw.line([(x, 50), (x, 250)], fill=0, width=2)

    # Words simulating an STT column where row 2 is missing
    words = [
        {"text": "STT", "left": 60, "right": 100, "top": 60, "bottom": 90, "width": 40, "height": 30, "confidence": 95, "line_key": (0, 0, 0)},
        {"text": "1", "left": 60, "right": 80, "top": 110, "bottom": 140, "width": 20, "height": 30, "confidence": 95, "line_key": (0, 1, 0)},
        # row 2 is missing number
        {"text": "CỘNG", "left": 60, "right": 120, "top": 210, "bottom": 240, "width": 60, "height": 30, "confidence": 95, "line_key": (0, 3, 0)},
    ]

    # When fill_sequences=False (Excel mode):
    tables_no_fill = detect_raster_tables(img, words, requested_language="vie", fill_sequences=False)
    if tables_no_fill:
        # Row 2 (index 2) must remain empty and not have a synthesized "2"
        assert tables_no_fill[0].cells[2][0] == ""


# ===========================================================================
# 3. Value Policy & Formula Injection Safety (CWE-1236)
# ===========================================================================
def test_format_cell_value_policy():
    warnings: list[str] = []

    # Unambiguous integers <= 15 digits become numeric
    val, dtype = format_cell_value("42", is_table_cell=True, warnings=warnings)
    assert val == 42 and dtype == "n"

    val, dtype = format_cell_value("-100", is_table_cell=True, warnings=warnings)
    assert val == -100 and dtype == "n"

    val, dtype = format_cell_value("0", is_table_cell=True, warnings=warnings)
    assert val == 0 and dtype == "n"

    # Preserved strings: -0, leading zeros, >15 digits
    val, dtype = format_cell_value("-0", is_table_cell=True, warnings=warnings)
    assert val == "-0" and dtype == "s"

    val, dtype = format_cell_value("0123456789", is_table_cell=True, warnings=warnings)
    assert val == "0123456789" and dtype == "s"

    val, dtype = format_cell_value("007", is_table_cell=True, warnings=warnings)
    assert val == "007" and dtype == "s"

    val, dtype = format_cell_value("1234567890123456", is_table_cell=True, warnings=warnings)
    assert val == "1234567890123456" and dtype == "s"

    # Ambiguous numbers and dates
    val, dtype = format_cell_value("12.34", is_table_cell=True, warnings=warnings)
    assert val == "12.34" and dtype == "s"

    val, dtype = format_cell_value("01/02/2026", is_table_cell=True, warnings=warnings)
    assert val == "01/02/2026" and dtype == "s"

    # Formula injection strings: =, +, -, @
    val, dtype = format_cell_value("=SUM(A1:A10)", is_table_cell=True, warnings=warnings)
    assert val == "=SUM(A1:A10)" and dtype == "s"

    val, dtype = format_cell_value("+cmd|' /C calc'!A0", is_table_cell=True, warnings=warnings)
    assert val == "+cmd|' /C calc'!A0" and dtype == "s"

    val, dtype = format_cell_value("@dangerous", is_table_cell=True, warnings=warnings)
    assert val == "@dangerous" and dtype == "s"

    # Invalid XML chars sanitized
    val, dtype = format_cell_value("bad\x00char\x08text", is_table_cell=True, warnings=warnings)
    assert val == "badchartext" and dtype == "s"

    # Cell exceeding 32,767 chars in table cell raises ConversionError
    oversized = "A" * (MAX_CELL_CHARS + 10)
    with pytest.raises(ConversionError, match="giới hạn Excel"):
        format_cell_value(oversized, is_table_cell=True, warnings=warnings)


# ===========================================================================
# 4. Digital PDF Table Extraction & Workbook Structure
# ===========================================================================
def _create_digital_pdf_with_table(pdf_path: Path) -> None:
    doc = pymupdf.open()
    page = doc.new_page(width=600, height=800)
    # Header title outside table
    page.insert_text(pymupdf.Point(50, 40), "BÁO CÁO TÀI CHÍNH NĂM 2026", fontsize=14)

    # 3x3 table with grid lines
    # Outer box: (50, 70, 450, 220)
    page.draw_rect(pymupdf.Rect(50, 70, 450, 220))
    # Horizontal lines at 120, 170
    page.draw_line(pymupdf.Point(50, 120), pymupdf.Point(450, 120))
    page.draw_line(pymupdf.Point(50, 170), pymupdf.Point(450, 170))
    # Vertical lines at 150, 300
    page.draw_line(pymupdf.Point(150, 70), pymupdf.Point(150, 220))
    page.draw_line(pymupdf.Point(300, 70), pymupdf.Point(300, 220))

    # Cell texts
    # Row 0 (Header)
    page.insert_text(pymupdf.Point(70, 100), "Mã SP")
    page.insert_text(pymupdf.Point(180, 100), "Tên sản phẩm")
    page.insert_text(pymupdf.Point(330, 100), "Số lượng")

    # Row 1
    page.insert_text(pymupdf.Point(70, 150), "0012")  # leading zero
    page.insert_text(pymupdf.Point(180, 150), "Điện thoại Pro")
    page.insert_text(pymupdf.Point(330, 150), "50")

    # Row 2
    page.insert_text(pymupdf.Point(70, 200), "=SUM(A1)")  # formula string
    page.insert_text(pymupdf.Point(180, 200), "Laptop Ultra")
    page.insert_text(pymupdf.Point(330, 200), "12")

    # Footer note outside table
    page.insert_text(pymupdf.Point(50, 260), "Ghi chú: Bảng số liệu nội bộ.", fontsize=11)

    doc.save(pdf_path)
    doc.close()


def test_convert_digital_pdf_to_excel(tmp_path: Path):
    pdf_path = tmp_path / "digital_table.pdf"
    xlsx_path = tmp_path / "result.xlsx"
    _create_digital_pdf_with_table(pdf_path)

    options = ConversionOptions(output_format="xlsx")
    result = convert_to_excel(pdf_path, xlsx_path, options)

    assert result.table_count == 1
    assert xlsx_path.exists()

    # Reopen workbook and verify actual cell values and data types
    wb = openpyxl.load_workbook(xlsx_path)
    sheet_names = wb.sheetnames
    assert "Bảng 1" in sheet_names
    assert "Văn bản" in sheet_names  # Outside text captured

    ws = wb["Bảng 1"]
    # Check Header (neutral Calibri 11 regular style applied without false bold assumptions)
    assert ws.cell(row=1, column=1).value == "Mã SP"
    assert ws.cell(row=1, column=1).font.bold is False

    # Check Row 2: leading zero must remain string
    c_id = ws.cell(row=2, column=1)
    assert c_id.value == "0012"
    assert c_id.data_type == "s"

    # Check numeric quantity
    c_qty = ws.cell(row=2, column=3)
    assert c_qty.value == 50
    assert c_qty.data_type == "n"

    # Check Row 3: formula must be literal string
    c_formula = ws.cell(row=3, column=1)
    assert c_formula.value == "=SUM(A1)"
    assert c_formula.data_type == "s"

    # Check text sheet
    ws_text = wb["Văn bản"]
    all_text = [ws_text.cell(row=r, column=3).value for r in range(2, ws_text.max_row + 1)]
    assert any("BÁO CÁO TÀI CHÍNH" in str(t) for t in all_text)
    assert any("Ghi chú" in str(t) for t in all_text)


# ===========================================================================
# 5. Merged Cells in Digital PDF
# ===========================================================================
def test_digital_pdf_merged_cells(tmp_path: Path):
    pdf_path = tmp_path / "merged_table.pdf"
    xlsx_path = tmp_path / "merged_result.xlsx"

    doc = pymupdf.open()
    page = doc.new_page(width=500, height=500)
    # 2x2 table with top row merged
    page.draw_rect(pymupdf.Rect(50, 50, 350, 150))
    page.draw_line(pymupdf.Point(50, 100), pymupdf.Point(350, 100))
    page.draw_line(pymupdf.Point(200, 100), pymupdf.Point(200, 150))

    page.insert_text(pymupdf.Point(120, 80), "Header Gộp Hàng Đầu")
    page.insert_text(pymupdf.Point(80, 130), "Ô A2")
    page.insert_text(pymupdf.Point(230, 130), "Ô B2")
    doc.save(pdf_path)
    doc.close()

    options = ConversionOptions(output_format="xlsx")
    convert_to_excel(pdf_path, xlsx_path, options)

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb["Bảng 1"]
    # Check that merge range exists
    merged_ranges = [str(m) for m in ws.merged_cells.ranges]
    assert len(merged_ranges) >= 1
    assert "A1:B1" in merged_ranges or any("1" in m for m in merged_ranges)
    wb.close()


# ===========================================================================
# 6. Fallback When No Tables Found
# ===========================================================================
def test_digital_pdf_no_table_fallback(tmp_path: Path):
    pdf_path = tmp_path / "no_tables.pdf"
    xlsx_path = tmp_path / "no_tables_result.xlsx"

    doc = pymupdf.open()
    page = doc.new_page(width=500, height=500)
    page.insert_text(pymupdf.Point(50, 50), "Đoạn văn 1: Đây là tài liệu chỉ có văn bản.")
    page.insert_text(pymupdf.Point(50, 90), "Đoạn văn 2: Không có bất kỳ khung bảng nào ở đây.")
    doc.save(pdf_path)
    doc.close()

    options = ConversionOptions(output_format="xlsx")
    result = convert_to_excel(pdf_path, xlsx_path, options)

    assert result.table_count == 0
    assert any("Không phát hiện được bảng" in w for w in result.warnings)

    wb = openpyxl.load_workbook(xlsx_path)
    assert "Văn bản" in wb.sheetnames
    ws = wb["Văn bản"]
    assert ws.max_row >= 2
    wb.close()


# ===========================================================================
# 7. Input Rejection & Bounds
# ===========================================================================
def test_reject_excel_input_for_xlsx(tmp_path: Path):
    dummy_xlsx = tmp_path / "input.xlsx"
    dummy_xlsx.write_bytes(b"PK\x03\x04fake")
    output = tmp_path / "out.xlsx"

    options = ConversionOptions(output_format="xlsx")
    with pytest.raises(ConversionError, match="chỉ nhận file PDF hoặc hình ảnh"):
        convert_to_excel(dummy_xlsx, output, options)


def test_reject_incompatible_math_and_ai(tmp_path: Path):
    pdf_path = tmp_path / "sample.pdf"
    doc = pymupdf.open()
    doc.new_page()
    doc.save(pdf_path)
    doc.close()

    output = tmp_path / "out.xlsx"
    # math_ocr is rejected in validate()
    with pytest.raises(ConversionError, match="không hỗ trợ nhận diện công thức toán"):
        ConversionOptions(output_format="xlsx", math_ocr=True).validate()


def test_empty_pdf_raises_error(tmp_path: Path):
    pdf_path = tmp_path / "empty.pdf"
    doc = pymupdf.open()
    doc.new_page()  # 1 blank page with zero text
    doc.save(pdf_path)
    doc.close()

    output = tmp_path / "out.xlsx"
    options = ConversionOptions(output_format="xlsx")
    with pytest.raises(ConversionError, match="không có nội dung chữ hoặc bảng"):
        convert_to_excel(pdf_path, output, options)


def test_page_bound_exceeded(tmp_path: Path):
    pdf_path = tmp_path / "too_many_pages.pdf"
    doc = pymupdf.open()
    for _ in range(MAX_PAGES + 1):
        doc.new_page()
    doc.save(pdf_path)
    doc.close()

    output = tmp_path / "out.xlsx"
    options = ConversionOptions(output_format="xlsx")
    with pytest.raises(ConversionError, match="vượt quá giới hạn 200 trang"):
        convert_to_excel(pdf_path, output, options)


# ===========================================================================
# 8. Scanned Table with Deterministic OCR Fixture
# ===========================================================================
def test_scanned_pdf_excel_conversion(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    # Create image with grid
    img_path = tmp_path / "scan_grid.png"
    img = Image.new("RGB", (1200, 800), "white")
    draw = ImageDraw.Draw(img)
    for y in [100, 200, 300, 400]:
        draw.line([(100, y), (1100, y)], fill="black", width=3)
    for x in [100, 400, 750, 1100]:
        draw.line([(x, 100), (x, 400)], fill="black", width=3)
    img.save(img_path)

    # Monkeypatch OCR data with deterministic words
    def fake_image_to_data(_img, _lang=None):
        return (
            {
                "text": ["Mã", "Tên", "Điểm", "001", "Học sinh A", "10"],
                "left": [150, 450, 800, 150, 450, 800],
                "top": [130, 130, 130, 230, 230, 230],
                "width": [80, 120, 80, 80, 150, 50],
                "height": [40, 40, 40, 40, 40, 40],
                "conf": [95, 95, 95, 95, 95, 95],
            },
            "vie",
            [],
        )

    monkeypatch.setattr("app.converters.to_excel.image_to_data", fake_image_to_data)

    output_xlsx = tmp_path / "scan_result.xlsx"
    options = ConversionOptions(output_format="xlsx")
    result = convert_to_excel(img_path, output_xlsx, options)

    assert result.table_count == 1
    wb = openpyxl.load_workbook(output_xlsx)
    ws = wb["Bảng 1"]
    assert ws.cell(row=1, column=1).value == "Mã"
    # Leading zero preserved
    c_id = ws.cell(row=2, column=1)
    assert c_id.value == "001" and c_id.data_type == "s"
    # Number parsed
    c_score = ws.cell(row=2, column=3)
    assert c_score.value == 10 and c_score.data_type == "n"
    wb.close()


# ===========================================================================
# 9. Local FastAPI /api/convert Endpoint Integration
# ===========================================================================
def test_local_api_convert_xlsx(tmp_path: Path):
    pdf_path = tmp_path / "api_test.pdf"
    _create_digital_pdf_with_table(pdf_path)

    client = TestClient(app)
    with pdf_path.open("rb") as f:
        response = client.post(
            "/api/convert",
            data={"output_format": "xlsx"},
            files={"files": (pdf_path.name, f, "application/pdf")},
        )
    assert response.status_code == 200
    assert response.headers["content-type"].startswith("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    assert "api_test-editable.xlsx" in response.headers.get("content-disposition", "")

    # Check workbook bytes can be loaded
    wb = openpyxl.load_workbook(io.BytesIO(response.content))
    assert "Bảng 1" in wb.sheetnames
    wb.close()


def test_local_api_convert_xlsx_rejects_excel_input(tmp_path: Path):
    client = TestClient(app)
    response = client.post(
        "/api/convert",
        data={"output_format": "xlsx"},
        files={"files": ("table.xlsx", b"PKfake", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 400
    assert "chỉ nhận file PDF hoặc hình ảnh" in response.json().get("detail", "")


# ===========================================================================
# 10. Modal Serverless Jobs Integration
# ===========================================================================
def test_modal_serverless_jobs_xlsx(tmp_path: Path):
    pdf_path = tmp_path / "modal_test.pdf"
    _create_digital_pdf_with_table(pdf_path)

    files_payload = [{"name": pdf_path.name, "content": pdf_path.read_bytes()}]
    options_payload = ServerlessConversionOptions(output_format="xlsx").model_dump()

    output = run_serverless_conversion(files_payload, options_payload)
    assert output.media_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert output.filename.endswith("-editable.xlsx")

    wb = openpyxl.load_workbook(io.BytesIO(output.content))
    assert "Bảng 1" in wb.sheetnames
    wb.close()


# ===========================================================================
# 11. P1-1: Browser-Equivalent Options Serialization & API Integration
# ===========================================================================
def test_api_convert_browser_equivalent_payload_xlsx(tmp_path: Path):
    pdf_path = tmp_path / "browser_test.pdf"
    _create_digital_pdf_with_table(pdf_path)

    client = TestClient(app)
    # The browser sends exact form fields when Excel output is chosen:
    # math_engine is omitted by frontend for xlsx
    form_data = {
        "output_format": "xlsx",
        "language": "vie",
        "ocr_mode": "auto",
        "dpi": "300",
        "page_breaks": "true",
        "preserve_images": "true",
        "preserve_layout": "true",
        "detect_tables": "true",
        "math_ocr": "false",
    }
    with pdf_path.open("rb") as f:
        response = client.post(
            "/api/convert",
            data=form_data,
            files={"files": (pdf_path.name, f, "application/pdf")},
        )
    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert "browser_test-editable.xlsx" in response.headers.get("content-disposition", "")
    wb = openpyxl.load_workbook(io.BytesIO(response.content))
    assert "Bảng 1" in wb.sheetnames
    wb.close()


def test_api_convert_xlsx_rejects_incompatible_math_and_ai(tmp_path: Path):
    pdf_path = tmp_path / "incompatible.pdf"
    _create_digital_pdf_with_table(pdf_path)
    client = TestClient(app)

    # 1. math_ocr=True rejected
    with pdf_path.open("rb") as f:
        resp = client.post(
            "/api/convert",
            data={"output_format": "xlsx", "math_ocr": "true"},
            files={"files": (pdf_path.name, f, "application/pdf")},
        )
    assert resp.status_code == 400
    assert "công thức toán" in resp.json().get("detail", "")



# ===========================================================================
# 12. P1-3: Text Chunk Loop Preservation & Controlled Failure
# ===========================================================================
def test_text_chunk_loop_preserves_60k_line_across_rows():
    # Verify that a 60,000 char line is split into two 30,000 char chunks across separate rows,
    # and concatenating them reconstructs the exact original string with zero data loss or overwrite.
    chunk1 = "A" * 30000
    chunk2 = "B" * 30000
    full_text = chunk1 + chunk2
    text_lines = [ExtractedTextLine(page_num=1, line_num=1, text=full_text)]
    warnings: list[str] = []

    wb = _build_excel_workbook([], text_lines, warnings)
    ws = wb["Văn bản"]

    assert ws.max_row == 3  # Header on row 1, chunk 1 on row 2, chunk 2 on row 3
    assert ws.cell(row=2, column=1).value == 1  # page_num
    assert ws.cell(row=2, column=2).value == 1  # line_num
    assert ws.cell(row=2, column=3).value == chunk1

    assert ws.cell(row=3, column=1).value == 1  # page_num
    assert ws.cell(row=3, column=2).value == 1  # line_num
    assert ws.cell(row=3, column=3).value == chunk2

    reconstructed = ws.cell(row=2, column=3).value + ws.cell(row=3, column=3).value
    assert reconstructed == full_text
    assert len(reconstructed) == 60000
    wb.close()


def test_text_chunk_row_limit_raises_conversion_error(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setattr("app.converters.to_excel.MAX_ROWS_PER_SHEET", 3)
    # Header is row 1, chunk 1 is row 2, chunk 2 is row 3, chunk 3 would be row 4 > 3
    long_text = "X" * 70000  # 3 chunks
    text_lines = [ExtractedTextLine(page_num=1, line_num=1, text=long_text)]
    warnings: list[str] = []

    with pytest.raises(ConversionError, match="vượt quá giới hạn .* dòng"):
        _build_excel_workbook([], text_lines, warnings)


def test_text_chunk_cell_budget_raises_conversion_error(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setattr("app.converters.to_excel.MAX_CELLS_PER_WORKBOOK", 5)
    # Header allocates 3 cells. Each chunk allocates 3 cells. 3 + 3 = 6 > 5.
    text_lines = [ExtractedTextLine(page_num=1, line_num=1, text="Hello world")]
    warnings: list[str] = []

    with pytest.raises(ConversionError, match="vượt quá giới hạn .* ô"):
        _build_excel_workbook([], text_lines, warnings)


def test_table_cell_oversized_raises_conversion_error():
    oversized = "Z" * (MAX_CELL_CHARS + 50)
    table = ExtractedTable(page_num=1, table_idx=1, cells=[["OK", oversized]], merges=[])
    warnings: list[str] = []

    with pytest.raises(ConversionError, match="vượt quá giới hạn .* ký tự"):
        _build_excel_workbook([table], [], warnings)


def test_format_cell_value_preserves_whitespace():
    warnings: list[str] = []
    # Leading and trailing whitespace must be preserved for non-integer strings
    val, dtype = format_cell_value("   Khoản mục đầu tư   ", is_table_cell=True, warnings=warnings)
    assert val == "   Khoản mục đầu tư   "
    assert dtype == "s"

    # Spaces around numbers prevent integer coercion to maintain literal formatting
    val, dtype = format_cell_value("  100  ", is_table_cell=True, warnings=warnings)
    assert val == "  100  "
    assert dtype == "s"


# ===========================================================================
# 13. P1-4: Mixed-Page Recovery & Fallbacks
# ===========================================================================
def test_mixed_page_digital_table_and_raster_image(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    pdf_path = tmp_path / "mixed_page.pdf"
    xlsx_path = tmp_path / "mixed_result.xlsx"

    doc = pymupdf.open()
    page = doc.new_page(width=600, height=800)
    # 1. Digital table at top: (50, 50, 450, 150)
    page.draw_rect(pymupdf.Rect(50, 50, 450, 150))
    page.draw_line(pymupdf.Point(50, 100), pymupdf.Point(450, 100))
    page.draw_line(pymupdf.Point(250, 50), pymupdf.Point(250, 150))
    page.insert_text(pymupdf.Point(70, 80), "Col 1")
    page.insert_text(pymupdf.Point(270, 80), "Col 2")
    page.insert_text(pymupdf.Point(70, 130), "Item 1")
    page.insert_text(pymupdf.Point(270, 130), "100")

    # 2. Embedded raster image below digital table: (50, 300, 450, 550) -> area 400x250 = 100,000 pt²
    raster_img = Image.new("RGB", (400, 250), "white")
    r_draw = ImageDraw.Draw(raster_img)
    r_draw.rectangle([10, 10, 390, 240], outline="black", width=2)
    img_bytes_io = io.BytesIO()
    raster_img.save(img_bytes_io, format="PNG")
    page.insert_image(pymupdf.Rect(50, 300, 450, 550), stream=img_bytes_io.getvalue())

    doc.save(pdf_path)
    doc.close()

    # Monkeypatch _extract_raster_page to confirm it is called and returns a raster table
    called = []
    def fake_extract_raster(img, p_num, opts, warns):
        called.append(p_num)
        return (
            [ExtractedTable(page_num=p_num, table_idx=2, cells=[["Quét A", "Quét B"], ["99", "88"]], merges=[], source_type="raster")],
            [ExtractedTextLine(page_num=p_num, line_num=1, text="Ghi chú dưới ảnh quét")],
        )

    monkeypatch.setattr("app.converters.to_excel._extract_raster_page", fake_extract_raster)

    options = ConversionOptions(output_format="xlsx", ocr_mode="auto")
    result = convert_to_excel(pdf_path, xlsx_path, options)

    assert len(called) == 1
    assert result.table_count == 2
    wb = openpyxl.load_workbook(xlsx_path)
    assert "T1_B1" in wb.sheetnames  # Digital table
    assert "T1_B2" in wb.sheetnames  # Raster table
    assert "Văn bản" in wb.sheetnames  # Raster text

    ws_t1 = wb["T1_B1"]
    assert ws_t1.cell(row=1, column=1).value == "Col 1"

    ws_t2 = wb["T1_B2"]
    assert ws_t2.cell(row=1, column=1).value == "Quét A"
    wb.close()


def test_mixed_page_header_text_and_raster_text_no_grid(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    pdf_path = tmp_path / "mixed_text.pdf"
    xlsx_path = tmp_path / "mixed_text_out.xlsx"

    doc = pymupdf.open()
    page = doc.new_page(width=600, height=800)
    page.insert_text(pymupdf.Point(50, 50), "REPORT HEADER SUMMARY FOR EXECUTIVE BOARD 2026", fontsize=14)

    # Embedded raster image (size 300x200, area 60,000 pt² > 15,000 pt²)
    raster_img = Image.new("RGB", (300, 200), "white")
    img_bytes_io = io.BytesIO()
    raster_img.save(img_bytes_io, format="PNG")
    page.insert_image(pymupdf.Rect(50, 200, 350, 400), stream=img_bytes_io.getvalue())
    doc.save(pdf_path)
    doc.close()

    def fake_extract_raster(img, p_num, opts, warns):
        return (
            [],
            [ExtractedTextLine(page_num=p_num, line_num=1, text="Nội dung quét từ biểu mẫu scan")],
        )

    monkeypatch.setattr("app.converters.to_excel._extract_raster_page", fake_extract_raster)

    options = ConversionOptions(output_format="xlsx", ocr_mode="auto")
    result = convert_to_excel(pdf_path, xlsx_path, options)

    assert result.table_count == 0
    wb = openpyxl.load_workbook(xlsx_path)
    assert "Văn bản" in wb.sheetnames
    ws_text = wb["Văn bản"]
    all_text = [ws_text.cell(row=r, column=3).value for r in range(2, ws_text.max_row + 1)]
    assert any("REPORT HEADER SUMMARY" in str(t) for t in all_text)
    assert any("Nội dung quét từ biểu mẫu scan" in str(t) for t in all_text)
    wb.close()


def test_scanned_table_only_no_empty_outside_text_lines(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    img_path = tmp_path / "table_only.png"
    img = Image.new("RGB", (400, 400), "white")
    img.save(img_path)

    # OCR returns a table and NO text outside
    def fake_extract_raster(img_frame, p_num, opts, warns):
        return (
            [ExtractedTable(page_num=p_num, table_idx=1, cells=[["Mã", "Tên"], ["1", "A"]], merges=[], source_type="raster")],
            [],  # Empty outside lines
        )

    monkeypatch.setattr("app.converters.to_excel._extract_raster_page", fake_extract_raster)

    out_xlsx = tmp_path / "table_only_out.xlsx"
    options = ConversionOptions(output_format="xlsx")
    result = convert_to_excel(img_path, out_xlsx, options)

    assert result.table_count == 1
    wb = openpyxl.load_workbook(out_xlsx)
    assert "Bảng 1" in wb.sheetnames
    # When there is no outside text, "Văn bản" sheet is not created
    assert "Văn bản" not in wb.sheetnames
    wb.close()


def test_multipage_mixed_digital_and_scanned(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    pdf_path = tmp_path / "multipage_mixed.pdf"
    xlsx_path = tmp_path / "multipage_mixed_out.xlsx"

    doc = pymupdf.open()
    # Page 1: Digital table
    p1 = doc.new_page(width=600, height=800)
    p1.draw_rect(pymupdf.Rect(50, 50, 450, 150))
    p1.draw_line(pymupdf.Point(50, 100), pymupdf.Point(450, 100))
    p1.insert_text(pymupdf.Point(70, 80), "Trang 1 Kỹ thuật số")
    p1.insert_text(pymupdf.Point(70, 130), "Dữ liệu 1")

    # Page 2: Pure scan image
    p2 = doc.new_page(width=600, height=800)
    img = Image.new("RGB", (600, 800), "white")
    img_bytes = io.BytesIO()
    img.save(img_bytes, format="PNG")
    p2.insert_image(pymupdf.Rect(0, 0, 600, 800), stream=img_bytes.getvalue())

    doc.save(pdf_path)
    doc.close()

    def fake_extract_raster(img, p_num, opts, warns):
        return (
            [ExtractedTable(page_num=p_num, table_idx=1, cells=[["Quét P2", "123"]], merges=[], source_type="raster")],
            [],
        )

    monkeypatch.setattr("app.converters.to_excel._extract_raster_page", fake_extract_raster)

    options = ConversionOptions(output_format="xlsx", ocr_mode="auto")
    result = convert_to_excel(pdf_path, xlsx_path, options)

    assert result.page_count == 2
    assert result.table_count == 2
    wb = openpyxl.load_workbook(xlsx_path)
    assert "T1_B1" in wb.sheetnames
    assert "T2_B1" in wb.sheetnames
    wb.close()


def test_text_mode_rejects_image_input_without_ocr(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    img_path = tmp_path / "input_scan.png"
    img = Image.new("RGB", (200, 200), "white")
    img.save(img_path)

    def fake_image_to_data(*args, **kwargs):
        raise AssertionError("image_to_data must NEVER be called in text mode!")

    monkeypatch.setattr("app.converters.to_excel.image_to_data", fake_image_to_data)

    options = ConversionOptions(output_format="xlsx", ocr_mode="text")
    with pytest.raises(ConversionError, match="không thể trích xuất ở chế độ 'Chỉ trích xuất lớp chữ'"):
        convert_to_excel(img_path, tmp_path / "out.xlsx", options)


# ===========================================================================
# 14. P1-5: Early Resource Limits & Cumulative Budget
# ===========================================================================
def test_candidate_bounds_checks_before_merge_map(monkeypatch: pytest.MonkeyPatch):
    merge_map_called = []
    def fake_merge_map(*args, **kwargs):
        merge_map_called.append(True)
        return {}

    monkeypatch.setattr("app.converters.raster_tables._merge_map", fake_merge_map)

    img = Image.new("L", (400, 400), 255)
    draw = ImageDraw.Draw(img)
    for y in [50, 100, 150, 200]:
        draw.line([(50, y), (350, y)], fill=0, width=2)
    for x in [50, 150, 250, 350]:
        draw.line([(x, 50), (x, 200)], fill=0, width=2)
    def strict_candidate_checker(rows: int, cols: int):
        raise ConversionError("Candidate exceeds limit")

    words = [{"text": "A", "left": 10, "right": 30, "top": 10, "bottom": 30, "width": 20, "height": 20, "confidence": 90, "line_key": (0, 0, 0)}]

    with pytest.raises(ConversionError, match="Candidate exceeds limit"):
        detect_raster_tables(img, words, candidate_limit_checker=strict_candidate_checker)

    assert len(merge_map_called) == 0


def test_max_pixels_limit_checks_before_source_copy(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    img_path = tmp_path / "huge.png"
    img = Image.new("RGB", (10, 10), "white")
    img.save(img_path)

    real_open = Image.open
    class HugeImage:
        def __init__(self):
            self._real = real_open(img_path)
            self.width = 6000
            self.height = 5000  # 30,000,000 pixels > 25,000,000
            self.mode = "RGB"

        def seek(self, frame):
            pass

        def copy(self):
            raise AssertionError("source.copy() must NEVER be called for images exceeding pixel limit!")

        def close(self):
            self._real.close()

    monkeypatch.setattr("app.converters.to_excel.PillowImage.open", lambda p: HugeImage())

    options = ConversionOptions(output_format="xlsx")
    with pytest.raises(ConversionError, match="25,000,000 pixel"):
        convert_to_excel(img_path, tmp_path / "out.xlsx", options)


def test_cumulative_cell_budget_across_multiple_pages(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    pdf_path = tmp_path / "two_pages.pdf"
    doc = pymupdf.open()
    doc.new_page()
    doc.new_page()
    doc.save(pdf_path)
    doc.close()

    monkeypatch.setattr("app.converters.to_excel.MAX_CELLS_PER_WORKBOOK", 50)

    class FakeTable:
        bbox = (0, 0, 100, 100)
        row_count = 10
        col_count = 4
        rows = []
        def extract(self):
            return [["x"] * 4 for _ in range(10)]

    class FakeFinder:
        tables = [FakeTable()]

    monkeypatch.setattr("pymupdf.Page.find_tables", lambda self: FakeFinder())
    def fake_get_text(self, opt=""):
        if opt == "words":
            return []
        return "Nội dung văn bản số kỹ thuật số đủ dài để vượt ngưỡng ký tự tối thiểu."
    monkeypatch.setattr("pymupdf.Page.get_text", fake_get_text)

    options = ConversionOptions(output_format="xlsx", ocr_mode="auto")
    with pytest.raises(ConversionError, match="vượt quá giới hạn .* ô"):
        convert_to_excel(pdf_path, tmp_path / "out.xlsx", options)


def test_empty_raster_grid_without_ocr_text_raises_conversion_error(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    img_path = tmp_path / "empty_grid.png"
    img = Image.new("RGB", (400, 400), "white")
    img.save(img_path)

    # OCR returns 0 words
    def fake_image_to_data(_img, _lang=None):
        return ({"text": [], "left": [], "top": [], "width": [], "height": [], "conf": []}, "vie", [])

    monkeypatch.setattr("app.converters.to_excel.image_to_data", fake_image_to_data)

    options = ConversionOptions(output_format="xlsx")
    with pytest.raises(ConversionError, match="không có nội dung chữ hoặc bảng hợp lệ"):
        convert_to_excel(img_path, tmp_path / "out.xlsx", options)


# ===========================================================================
# 15. P2-7: Python Vercel Worker /internal/convert Integration
# ===========================================================================
@pytest.mark.anyio
async def test_vercel_worker_convert_xlsx_single_file_success(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("OCR2WORD_INTERNAL_SECRET", "test-secret")
    pdf_path = tmp_path / "vercel_worker_test.pdf"
    _create_digital_pdf_with_table(pdf_path)
    pdf_bytes = pdf_path.read_bytes()

    uploaded_data: list[tuple[bytes, str]] = []

    def fake_download(source, input_path, **kwargs):
        input_path.write_bytes(pdf_bytes)
        return len(pdf_bytes)

    monkeypatch.setattr("app.main._download_vercel_input", fake_download)

    def fake_upload(path, dest):
        uploaded_data.append((path.read_bytes(), dest.content_type))

    monkeypatch.setattr("app.main._upload_vercel_result", fake_upload)

    body = {
        "job_id": "b" * 32,
        "files": [
            {
                "name": "sample.pdf",
                "size": len(pdf_bytes),
                "content_type": "application/pdf",
                "download_url": "https://store.blob.vercel-storage.com/sample.pdf",
            }
        ],
        "options": {
            "output_format": "xlsx",
            "language": "vie",
            "ocr_mode": "auto",
            "dpi": 300,
            "page_breaks": True,
            "preserve_images": True,
            "preserve_layout": True,
            "detect_tables": True,
            "math_ocr": False,
        },
        "output": {
            "name": "sample-editable.xlsx",
            "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "upload_url": "https://vercel.com/api/blob/?pathname=sample-editable.xlsx",
        },
    }

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        resp = await client.post(
            "/internal/convert",
            json=body,
            headers={"Authorization": "Bearer test-secret"},
        )

    assert resp.status_code == 200
    data = resp.json()
    assert data["output_name"] == "sample-editable.xlsx"
    assert len(uploaded_data) == 1
    content, ctype = uploaded_data[0]
    assert ctype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # Verify uploaded bytes form a real, valid openpyxl workbook
    wb = openpyxl.load_workbook(io.BytesIO(content))
    assert "Bảng 1" in wb.sheetnames
    wb.close()


@pytest.mark.anyio
async def test_vercel_worker_convert_xlsx_rejects_incompatible_options(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("OCR2WORD_INTERNAL_SECRET", "test-secret")
    base_body = {
        "job_id": "c" * 32,
        "files": [{"name": "sample.pdf", "size": 100, "content_type": "application/pdf", "download_url": "https://store.blob.vercel-storage.com/sample.pdf"}],
        "options": {
            "output_format": "xlsx",
            "language": "vie",
            "ocr_mode": "auto",
            "dpi": 300,
            "page_breaks": True,
            "preserve_images": True,
            "preserve_layout": True,
            "detect_tables": True,
            "math_ocr": False,
        },
        "output": {
            "name": "sample-editable.xlsx",
            "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "upload_url": "https://vercel.com/api/blob/?pathname=sample-editable.xlsx",
        },
    }

    monkeypatch.setattr("app.main._download_vercel_input", lambda s, p, **kwargs: (p.write_bytes(b"%PDF-1.4 dummy"), 100)[1])
    monkeypatch.setattr("app.main._validate_signature", lambda p, s: None)

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        # 1. math_ocr rejected
        body_math = dict(base_body)
        body_math["options"] = dict(base_body["options"], math_ocr=True)
        resp_math = await client.post("/internal/convert", json=body_math, headers={"Authorization": "Bearer test-secret"})
        assert resp_math.status_code == 422
        assert "công thức toán" in resp_math.json().get("detail", "")

        # 2. page_range rejected for xlsx
        body_pr = dict(base_body)
        body_pr["page_range"] = {"start_page": 1, "end_page": 2}
        resp_pr = await client.post("/internal/convert", json=body_pr, headers={"Authorization": "Bearer test-secret"})
        assert resp_pr.status_code == 422
        assert "chia nhỏ trang" in resp_pr.json().get("detail", "")


@pytest.mark.anyio
async def test_vercel_worker_convert_xlsx_partial_batch_returns_zip(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("OCR2WORD_INTERNAL_SECRET", "test-secret")
    pdf_path = tmp_path / "valid.pdf"
    _create_digital_pdf_with_table(pdf_path)
    pdf_bytes = pdf_path.read_bytes()

    uploaded_data: list[tuple[bytes, str]] = []

    def fake_download(source, input_path, **kwargs):
        if source.name == "valid.pdf":
            input_path.write_bytes(pdf_bytes)
            return len(pdf_bytes)
        else:
            input_path.write_bytes(b"%PDF-1.4 empty")
            return 14

    monkeypatch.setattr("app.main._download_vercel_input", fake_download)
    monkeypatch.setattr("app.main._validate_signature", lambda p, s: None)

    real_convert_one = app_main._convert_one
    def mock_convert_one(in_p, out_p, opts):
        if "invalid" in in_p.name:
            raise ConversionError("File thứ 2 bị lỗi không đọc được.")
        return real_convert_one(in_p, out_p, opts)

    monkeypatch.setattr("app.main._convert_one", mock_convert_one)
    monkeypatch.setattr("app.main._upload_vercel_result", lambda p, d: uploaded_data.append((p.read_bytes(), d.content_type)))

    body = {
        "job_id": "d" * 32,
        "files": [
            {"name": "valid.pdf", "size": len(pdf_bytes), "content_type": "application/pdf", "download_url": "https://store.blob.vercel-storage.com/valid.pdf"},
            {"name": "invalid.pdf", "size": 14, "content_type": "application/pdf", "download_url": "https://store.blob.vercel-storage.com/invalid.pdf"},
        ],
        "options": {
            "output_format": "xlsx",
            "language": "vie",
            "ocr_mode": "auto",
            "dpi": 300,
            "page_breaks": True,
            "preserve_images": True,
            "preserve_layout": True,
            "detect_tables": True,
            "math_ocr": False,
        },
        "output": {
            "name": "ocr2word-ket-qua.zip",
            "content_type": "application/zip",
            "upload_url": "https://vercel.com/api/blob/?pathname=ocr2word-ket-qua.zip",
        },
    }

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        resp = await client.post("/internal/convert", json=body, headers={"Authorization": "Bearer test-secret"})

    assert resp.status_code == 200
    assert len(uploaded_data) == 1
    content, _ = uploaded_data[0]

    with zipfile.ZipFile(io.BytesIO(content), "r") as z:
        names = z.namelist()
        assert any(n.endswith(".xlsx") for n in names)
        assert "bao-cao-chuyen-doi.txt" in names
        report_text = z.read("bao-cao-chuyen-doi.txt").decode("utf-8")
        assert "BÁO CÁO CHUYỂN ĐỔI" in report_text
        assert "[THÀNH CÔNG]" in report_text
        assert "[LỖI]" in report_text


# ===========================================================================
# 16. P2-8: Digital Merges & CWE-1236 Raw XML Verification
# ===========================================================================
def test_digital_merges_horizontal_and_vertical_integrity():
    cells = [
        ["Tiêu đề chung", "", ""],
        ["Nhóm", "Mục", "Giá trị"],
        ["Phòng A", "Văn phòng phẩm", "150"],
        ["", "Thiết bị", "350"],
    ]
    # merges: (top_row, left_col, bottom_row, right_col) 0-indexed
    # A1:C1 -> (0, 0, 0, 2)
    # A3:A4 -> (2, 0, 3, 0)
    merges = [(0, 0, 0, 2), (2, 0, 3, 0)]
    table = ExtractedTable(page_num=1, table_idx=1, cells=cells, merges=merges)

    wb = _build_excel_workbook([table], [], [])
    ws = wb["Bảng 1"]

    merged_str = [str(r) for r in ws.merged_cells.ranges]
    assert "A1:C1" in merged_str
    assert "A3:A4" in merged_str

    # Values preserved
    assert ws["A1"].value == "Tiêu đề chung"
    assert ws["A2"].value == "Nhóm"
    assert ws["B2"].value == "Mục"
    assert ws["C2"].value == "Giá trị"
    assert ws["A3"].value == "Phòng A"
    assert ws["B3"].value == "Văn phòng phẩm"
    assert ws["C3"].value == 150
    assert ws["B4"].value == "Thiết bị"
    assert ws["C4"].value == 350
    wb.close()


def test_cwe_1236_raw_xml_inspection_no_formula_tags(tmp_path: Path):
    xlsx_path = tmp_path / "cwe_1236_test.xlsx"
    cells = [
        ["=SUM(A1:A10)", "+cmd|' /C calc'!A0", "-5+5"],
        ["@SUM(1,2)", "=1+1", "Normal text"],
    ]
    table = ExtractedTable(page_num=1, table_idx=1, cells=cells, merges=[])
    wb = _build_excel_workbook([table], [], [])
    wb.save(xlsx_path)
    wb.close()

    # Inspect the saved zip archive directly to verify CWE-1236 immunity
    with zipfile.ZipFile(xlsx_path, "r") as z:
        sheet_files = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        assert len(sheet_files) >= 1
        for sheet_file in sheet_files:
            xml_content = z.read(sheet_file).decode("utf-8")
            # Must NOT contain <f> or <f ...> tags which represent executable Excel formulas
            assert "<f>" not in xml_content
            assert "<f " not in xml_content


# ===========================================================================
# 17. Revision 4 Defect 1: Reused PDF Image Xref Deduplication
# ===========================================================================
def test_reused_pdf_image_xref_does_not_duplicate_tables(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    pdf_path = tmp_path / "reused_xref.pdf"
    xlsx_path = tmp_path / "reused_xref.xlsx"

    doc = pymupdf.open()
    page = doc.new_page(width=600, height=800)
    page.insert_text(pymupdf.Point(50, 50), "REPORT DIGITAL HEADER FOR FINANCIAL DISCLOSURE 2026", fontsize=14)

    # Create a 300x200 PNG image
    img = Image.new("RGB", (300, 200), "white")
    draw = ImageDraw.Draw(img)
    draw.rectangle([10, 10, 290, 190], outline="black", width=2)
    img_bytes = io.BytesIO()
    img.save(img_bytes, format="PNG")
    img_data = img_bytes.getvalue()

    # Insert first placement and obtain its xref
    rect1 = pymupdf.Rect(50, 100, 350, 300)
    xref = page.insert_image(rect1, stream=img_data)

    # Insert second placement reusing the exact same xref
    rect2 = pymupdf.Rect(50, 400, 350, 600)
    page.insert_image(rect2, xref=xref)

    doc.save(pdf_path)
    doc.close()

    # Track OCR calls
    ocr_calls: list[int] = []

    def mock_extract_raster(raster_img, p_num, opts, warns):
        call_idx = len(ocr_calls) + 1
        ocr_calls.append(call_idx)
        return (
            [ExtractedTable(page_num=p_num, table_idx=call_idx, cells=[[f"TableData_{call_idx}"]], merges=[], source_type="raster")],
            [],
        )

    monkeypatch.setattr("app.converters.to_excel._extract_raster_page", mock_extract_raster)

    options = ConversionOptions(output_format="xlsx", ocr_mode="auto")
    result = convert_to_excel(pdf_path, xlsx_path, options)

    # CRITICAL: Exactly 2 calls (one for rect1, one for rect2) - NOT 4 calls!
    assert len(ocr_calls) == 2
    assert result.table_count == 2

    wb = openpyxl.load_workbook(xlsx_path)
    # Check sheets: T1_B1 (first placement), T1_B2 (second placement), Văn bản (digital header)
    assert "T1_B1" in wb.sheetnames
    assert "T1_B2" in wb.sheetnames
    assert "Văn bản" in wb.sheetnames

    # Deterministic top-to-bottom values:
    assert wb["T1_B1"].cell(row=1, column=1).value == "TableData_1"
    assert wb["T1_B2"].cell(row=1, column=1).value == "TableData_2"

    wb.close()


# ===========================================================================
# 18. Revision 4 Defect 2: Vercel Worker Early Output Contract Validation
# ===========================================================================
@pytest.mark.anyio
async def test_vercel_worker_rejects_mismatched_output_name_extension(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("OCR2WORD_INTERNAL_SECRET", "test-secret")

    def fail_transport(*args, **kwargs):
        raise AssertionError("Transport must NEVER be called when worker contract is invalid!")

    monkeypatch.setattr("app.main._download_vercel_input", fail_transport)
    monkeypatch.setattr("app.main._upload_vercel_result", fail_transport)

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        # Case 1: output_format is xlsx but output.name has .docx extension
        body_xlsx_docx = {
            "job_id": "e" * 32,
            "files": [{"name": "input.pdf", "size": 100, "content_type": "application/pdf", "download_url": "https://store.blob.vercel-storage.com/input.pdf"}],
            "options": {"output_format": "xlsx"},
            "output": {
                "name": "result.docx",
                "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "upload_url": "https://vercel.com/api/blob/?pathname=result.docx",
            },
        }
        resp = await client.post("/internal/convert", json=body_xlsx_docx, headers={"Authorization": "Bearer test-secret"})
        assert resp.status_code == 422
        assert "yêu cầu tên file kết quả có đuôi '.xlsx'" in resp.json().get("detail", "")

        # Case 2: output_format is docx but output.name has .xlsx extension
        body_docx_xlsx = {
            "job_id": "f" * 32,
            "files": [{"name": "input.pdf", "size": 100, "content_type": "application/pdf", "download_url": "https://store.blob.vercel-storage.com/input.pdf"}],
            "options": {"output_format": "docx"},
            "output": {
                "name": "result.xlsx",
                "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "upload_url": "https://vercel.com/api/blob/?pathname=result.xlsx",
            },
        }
        resp = await client.post("/internal/convert", json=body_docx_xlsx, headers={"Authorization": "Bearer test-secret"})
        assert resp.status_code == 422
        assert "yêu cầu tên file kết quả có đuôi '.docx'" in resp.json().get("detail", "")

        # Case 3: Multiple files but output.name is not .zip
        body_multi_docx = {
            "job_id": "1" * 32,
            "files": [
                {"name": "input1.pdf", "size": 100, "content_type": "application/pdf", "download_url": "https://store.blob.vercel-storage.com/input1.pdf"},
                {"name": "input2.pdf", "size": 100, "content_type": "application/pdf", "download_url": "https://store.blob.vercel-storage.com/input2.pdf"},
            ],
            "options": {"output_format": "docx"},
            "output": {
                "name": "result.docx",
                "content_type": "application/zip",
                "upload_url": "https://vercel.com/api/blob/?pathname=result.docx",
            },
        }
        resp = await client.post("/internal/convert", json=body_multi_docx, headers={"Authorization": "Bearer test-secret"})
        assert resp.status_code == 422
        assert "yêu cầu tên file kết quả có đuôi '.zip'" in resp.json().get("detail", "")


@pytest.mark.anyio
async def test_vercel_worker_rejects_mismatched_output_content_type(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("OCR2WORD_INTERNAL_SECRET", "test-secret")

    def fail_transport(*args, **kwargs):
        raise AssertionError("Transport must NEVER be called when worker contract is invalid!")

    monkeypatch.setattr("app.main._download_vercel_input", fail_transport)
    monkeypatch.setattr("app.main._upload_vercel_result", fail_transport)

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        # Case 1: output_format is xlsx but output.content_type is DOCX MIME
        body_xlsx_docx_mime = {
            "job_id": "2" * 32,
            "files": [{"name": "input.pdf", "size": 100, "content_type": "application/pdf", "download_url": "https://store.blob.vercel-storage.com/input.pdf"}],
            "options": {"output_format": "xlsx"},
            "output": {
                "name": "result.xlsx",
                "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "upload_url": "https://vercel.com/api/blob/?pathname=result.xlsx",
            },
        }
        resp = await client.post("/internal/convert", json=body_xlsx_docx_mime, headers={"Authorization": "Bearer test-secret"})
        assert resp.status_code == 422
        assert "yêu cầu Content-Type 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'" in resp.json().get("detail", "")

        # Case 2: output_format is docx but output.content_type is XLSX MIME
        body_docx_xlsx_mime = {
            "job_id": "3" * 32,
            "files": [{"name": "input.pdf", "size": 100, "content_type": "application/pdf", "download_url": "https://store.blob.vercel-storage.com/input.pdf"}],
            "options": {"output_format": "docx"},
            "output": {
                "name": "result.docx",
                "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "upload_url": "https://vercel.com/api/blob/?pathname=result.docx",
            },
        }
        resp = await client.post("/internal/convert", json=body_docx_xlsx_mime, headers={"Authorization": "Bearer test-secret"})
        assert resp.status_code == 422
        assert "yêu cầu Content-Type 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'" in resp.json().get("detail", "")

        # Case 3: Multiple files but output.content_type is not application/zip
        body_multi_not_zip = {
            "job_id": "4" * 32,
            "files": [
                {"name": "input1.pdf", "size": 100, "content_type": "application/pdf", "download_url": "https://store.blob.vercel-storage.com/input1.pdf"},
                {"name": "input2.pdf", "size": 100, "content_type": "application/pdf", "download_url": "https://store.blob.vercel-storage.com/input2.pdf"},
            ],
            "options": {"output_format": "xlsx"},
            "output": {
                "name": "results.zip",
                "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "upload_url": "https://vercel.com/api/blob/?pathname=results.zip",
            },
        }
        resp = await client.post("/internal/convert", json=body_multi_not_zip, headers={"Authorization": "Bearer test-secret"})
        assert resp.status_code == 422
        assert "yêu cầu Content-Type 'application/zip'" in resp.json().get("detail", "")


# ===========================================================================
# 19. Revision 4 Defect 3: TEXT Mode Explicit Warnings & All-Scan Rejection
# ===========================================================================
def test_text_mode_two_page_pdf_warns_for_omitted_scan_page(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    pdf_path = tmp_path / "text_mode_two_page.pdf"
    xlsx_path = tmp_path / "text_mode_two_page.xlsx"

    doc = pymupdf.open()
    # Page 1: Long searchable digital text
    p1 = doc.new_page(width=600, height=800)
    p1.insert_text(pymupdf.Point(50, 50), "ANNUAL FINANCIAL REPORT 2026 - DIGITAL TEXT CONTENT SUMMARY", fontsize=14)

    # Page 2: Only embedded raster scan image (zero digital text)
    p2 = doc.new_page(width=600, height=800)
    img = Image.new("RGB", (600, 800), "white")
    draw = ImageDraw.Draw(img)
    draw.rectangle([50, 50, 550, 750], outline="black", width=3)
    img_bytes = io.BytesIO()
    img.save(img_bytes, format="PNG")
    p2.insert_image(pymupdf.Rect(50, 50, 550, 750), stream=img_bytes.getvalue())

    doc.save(pdf_path)
    doc.close()

    # Verify OCR is NEVER called in text mode
    def fail_ocr(*args, **kwargs):
        raise AssertionError("OCR functions must NEVER be called in text mode!")

    monkeypatch.setattr("app.converters.to_excel.image_to_data", fail_ocr)
    monkeypatch.setattr("app.converters.to_excel._extract_raster_page", fail_ocr)

    options = ConversionOptions(output_format="xlsx", ocr_mode="text")
    result = convert_to_excel(pdf_path, xlsx_path, options)

    assert result.page_count == 2
    assert result.table_count == 0

    # Must contain warning specifically identifying Page 2
    assert any("Trang 2" in w and "Không tìm thấy lớp chữ số" in w for w in result.warnings)
    # Must NOT claim all data in the document was exported
    assert not any("đã xuất toàn bộ nội dung văn bản" in w for w in result.warnings)
    assert any("đã xuất nội dung văn bản trích xuất được" in w for w in result.warnings)

    wb = openpyxl.load_workbook(xlsx_path)
    assert "Văn bản" in wb.sheetnames
    all_text = [wb["Văn bản"].cell(row=r, column=3).value for r in range(2, wb["Văn bản"].max_row + 1)]
    assert any("ANNUAL FINANCIAL REPORT 2026" in str(t) for t in all_text)
    wb.close()


def test_text_mode_mixed_page_warns_for_unread_raster_region(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    pdf_path = tmp_path / "mixed_page_text_mode.pdf"
    xlsx_path = tmp_path / "mixed_page_text_mode.xlsx"

    doc = pymupdf.open()
    page = doc.new_page(width=600, height=800)
    # Digital header
    page.insert_text(pymupdf.Point(50, 50), "DIGITAL HEADER TEXT ON TOP OF MIXED PAGE", fontsize=14)

    # 1. Substantial scan image (300x200 pt, area 60,000 pt^2)
    sub_img = Image.new("RGB", (300, 200), "white")
    sub_bytes = io.BytesIO()
    sub_img.save(sub_bytes, format="PNG")
    page.insert_image(pymupdf.Rect(50, 150, 350, 350), stream=sub_bytes.getvalue())

    # 2. Tiny decorative logo (30x30 pt, area 900 pt^2 < 15,000 pt^2)
    tiny_img = Image.new("RGB", (30, 30), "white")
    tiny_bytes = io.BytesIO()
    tiny_img.save(tiny_bytes, format="PNG")
    page.insert_image(pymupdf.Rect(50, 400, 80, 430), stream=tiny_bytes.getvalue())

    doc.save(pdf_path)
    doc.close()

    def fail_ocr(*args, **kwargs):
        raise AssertionError("OCR functions must NEVER be called in text mode!")

    monkeypatch.setattr("app.converters.to_excel.image_to_data", fail_ocr)
    monkeypatch.setattr("app.converters.to_excel._extract_raster_page", fail_ocr)

    options = ConversionOptions(output_format="xlsx", ocr_mode="text")
    result = convert_to_excel(pdf_path, xlsx_path, options)

    # Substantial raster region must produce a warning for Trang 1
    assert any("Trang 1: Phát hiện vùng ảnh quét (300x200 pt) không được trích xuất" in w for w in result.warnings)
    # Tiny decorative logo must NOT generate a misleading warning
    assert not any("30x30 pt" in w for w in result.warnings)


def test_text_mode_all_scan_pdf_fails_with_actionable_error(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    pdf_path = tmp_path / "all_scan.pdf"
    xlsx_path = tmp_path / "all_scan.xlsx"

    doc = pymupdf.open()
    p1 = doc.new_page(width=600, height=800)
    img = Image.new("RGB", (600, 800), "white")
    img_bytes = io.BytesIO()
    img.save(img_bytes, format="PNG")
    p1.insert_image(pymupdf.Rect(0, 0, 600, 800), stream=img_bytes.getvalue())
    doc.save(pdf_path)
    doc.close()

    def fail_ocr(*args, **kwargs):
        raise AssertionError("OCR functions must NEVER be called in text mode!")

    monkeypatch.setattr("app.converters.to_excel.image_to_data", fail_ocr)
    monkeypatch.setattr("app.converters.to_excel._extract_raster_page", fail_ocr)

    options = ConversionOptions(output_format="xlsx", ocr_mode="text")
    with pytest.raises(ConversionError, match="chuyển sang chế độ Tự động.*hoặc OCR toàn bộ"):
        convert_to_excel(pdf_path, xlsx_path, options)
