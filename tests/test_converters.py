from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from app.converters import (
    ConversionOptions,
    analyze_pdf,
    convert_excel,
    convert_image,
    convert_pdf,
    extract_pdf_pages,
)
from app.converters.layout_pdf import _repair_tabbed_image_headers, repair_raw_text_blocks
from app.converters.ocr_layout import (
    _font_size_for_line,
    annotate_ocr_styles,
    merge_ocr_lines,
    normalize_document_ocr_text,
    normalize_ocr_text,
    ocr_words_to_lines,
    write_ocr_lines,
)
from app.services.ocr import _hocr_to_data, get_ocr_engine_info


def _all_docx_text(path: Path) -> str:
    document = Document(path)
    parts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    return "\n".join(parts)


def test_convert_digital_pdf_to_editable_docx(tmp_path: Path) -> None:
    source = tmp_path / "sample.pdf"
    output = tmp_path / "sample-editable.docx"
    pdf = canvas.Canvas(str(source), pagesize=letter)
    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(72, 720, "Editable PDF sample")
    pdf.setFont("Helvetica", 11)
    pdf.drawString(72, 690, "Invoice number: INV-2026-081")
    pdf.showPage()
    pdf.drawString(72, 720, "Second page content")
    pdf.save()

    result = convert_pdf(
        source,
        output,
        ConversionOptions(ocr_mode="auto", preserve_images=False, detect_tables=False),
    )

    assert result.page_count == 2
    assert output.exists() and output.stat().st_size > 1_000
    text = _all_docx_text(output)
    assert "Editable PDF sample" in text
    assert "INV-2026-081" in text
    assert "Second page content" in text
    converted = Document(output)
    separator = next(
        paragraph
        for paragraph in converted.paragraphs
        if 'w:type="page"' in paragraph._p.xml
    )
    assert separator.paragraph_format.line_spacing.pt == 1.0
    assert separator.paragraph_format.keep_with_next is False


def test_analyze_and_extract_pdf_page_range(tmp_path: Path) -> None:
    import pymupdf
    from PIL import Image, ImageDraw
    from reportlab.lib.utils import ImageReader

    source = tmp_path / "mixed.pdf"
    subset = tmp_path / "mixed-pages-2-3.pdf"
    scan_image = Image.new("RGB", (850, 1100), "white")
    drawing = ImageDraw.Draw(scan_image)
    drawing.text((80, 100), "SCANNED PAGE TWO", fill="black")
    image_buffer = __import__("io").BytesIO()
    scan_image.save(image_buffer, format="PNG")
    image_buffer.seek(0)

    pdf = canvas.Canvas(str(source), pagesize=letter)
    digital_text = "This digital page contains enough editable text for reliable detection. " * 2
    pdf.drawString(50, 720, digital_text)
    pdf.showPage()
    pdf.drawImage(ImageReader(image_buffer), 0, 0, width=letter[0], height=letter[1])
    # Simulate a scanned PDF that already has a long but unreliable hidden OCR
    # layer. Visual image coverage must still route this page through OCR.
    pdf.setFillColorRGB(1, 1, 1)
    pdf.drawString(50, 700, "Unreliable hidden OCR text should not make this page digital. " * 2)
    pdf.setFillColorRGB(0, 0, 0)
    pdf.showPage()
    pdf.drawString(50, 720, digital_text)
    pdf.save()

    analysis = analyze_pdf(source, ConversionOptions(ocr_mode="auto"))
    assert analysis.page_count == 3
    assert analysis.ocr_pages == (2,)
    assert analysis.image_backed_pages == (2,)

    assert extract_pdf_pages(source, subset, start_page=2, end_page=3) == 2
    with pymupdf.open(subset) as document:
        assert document.page_count == 2
        assert "digital page" in document[1].get_text().lower()


def test_auto_mixed_pdf_preserves_layout_instead_of_rebuilding_every_page(
    tmp_path: Path,
    monkeypatch,
) -> None:
    from PIL import Image, ImageDraw
    from reportlab.lib.utils import ImageReader

    source = tmp_path / "mixed-layout.pdf"
    output = tmp_path / "mixed-layout.docx"
    scan_image = Image.new("RGB", (850, 1100), "white")
    ImageDraw.Draw(scan_image).text((80, 100), "SIGNED SCANNED FORM", fill="black")
    image_buffer = __import__("io").BytesIO()
    scan_image.save(image_buffer, format="PNG")
    image_buffer.seek(0)

    pdf = canvas.Canvas(str(source), pagesize=letter)
    pdf.drawString(50, 720, "Editable digital cover text with stable PDF positioning. " * 2)
    pdf.showPage()
    pdf.drawImage(ImageReader(image_buffer), 0, 0, width=letter[0], height=letter[1])
    pdf.setFillColorRGB(1, 1, 1)
    pdf.drawString(50, 700, "Bad hidden OCR layer that must not replace the page image. " * 2)
    pdf.save()

    calls: list[tuple[Path, Path]] = []

    def fake_layout_converter(input_path, output_path, **_kwargs):
        calls.append((input_path, output_path))
        document = Document()
        document.add_paragraph("LAYOUT PRESERVED")
        document.save(output_path)

    monkeypatch.setattr("app.converters.pdf.layout_converter_available", lambda: True)
    monkeypatch.setattr("app.converters.pdf.convert_pdf_with_layout", fake_layout_converter)

    result = convert_pdf(
        source,
        output,
        ConversionOptions(
            ocr_mode="auto",
            preserve_layout=True,
            preserve_images=True,
            page_breaks=True,
        ),
    )

    assert calls == [(source, output)]
    assert output.exists()
    assert "LAYOUT PRESERVED" in _all_docx_text(output)
    assert any("giữ nguyên 1 trang scan" in warning for warning in result.warnings)


def test_repairs_missing_spaces_from_glyph_gaps() -> None:
    blocks = [
        {
            "type": 0,
            "lines": [
                {
                    "dir": (1.0, 0.0),
                    "spans": [
                        {
                            "font": "VnNimbusRomanNo9L",
                            "size": 12.0,
                            "chars": [
                                {"c": "đ", "origin": (0, 12), "bbox": (0, 0, 6, 12)},
                                {"c": "ồ", "origin": (6, 12), "bbox": (6, 0, 12, 12)},
                                {"c": "á", "origin": (15, 12), "bbox": (15, 0, 21, 12)},
                                {"c": "n", "origin": (21, 12), "bbox": (21, 0, 27, 12)},
                            ],
                        },
                        {
                            "font": "VnNimbusRomanNo9L",
                            "size": 12.0,
                            "chars": [{"c": " ", "origin": (27, 12), "bbox": (27, 0, 30, 12)}],
                        },
                        {
                            "font": "VnNimbusRomanNo9L",
                            "size": 12.0,
                            "chars": [
                                {"c": "t", "origin": (30, 12), "bbox": (30, 0, 34, 12)},
                                {"c": "ố", "origin": (34, 12), "bbox": (34, 0, 40, 12)},
                                {"c": "t", "origin": (40, 12), "bbox": (40, 0, 44, 12)},
                            ],
                        },
                    ],
                }
            ],
        }
    ]

    repair_raw_text_blocks(blocks)

    spans = blocks[0]["lines"][0]["spans"]
    text = "".join(character["c"] for span in spans for character in span["chars"])
    assert text.startswith("đồ án tốt")
    assert all(span["font"] == "Times New Roman" for span in spans)


def test_repairs_tabbed_image_header_without_touching_plain_tabs() -> None:
    from docx.oxml import OxmlElement

    document = Document()
    header = document.add_paragraph()
    header.add_run()._r.append(OxmlElement("w:drawing"))
    header.add_run("\t")
    header.add_run("Header line one")
    header.add_run("\t")
    header.add_run("Header line two")
    plain = document.add_paragraph("Left\tRight")

    assert _repair_tabbed_image_headers(document)
    assert header.text.endswith("Header line one")
    assert document.paragraphs[1].text == "Header line two"
    assert plain.text == "Left\tRight"


def test_layout_pdf_mode_keeps_editable_text_and_repairs_gap(tmp_path: Path) -> None:
    source = tmp_path / "layout.pdf"
    output = tmp_path / "layout-editable.docx"
    pdf = canvas.Canvas(str(source), pagesize=letter)
    pdf.setStrokeColorRGB(0.05, 0.2, 0.8)
    pdf.setLineWidth(3)
    pdf.rect(36, 36, 540, 720)
    pdf.setFillColorRGB(0, 0, 0)
    pdf.setFont("Times-Roman", 12)
    x = 72
    pdf.drawString(x, 700, "Hello")
    x += pdf.stringWidth("Hello", "Times-Roman", 12) + 4
    pdf.drawString(x, 700, "World")
    pdf.save()

    result = convert_pdf(
        source,
        output,
        ConversionOptions(
            ocr_mode="text",
            preserve_layout=True,
            preserve_images=True,
            page_breaks=True,
            detect_tables=False,
        ),
    )

    assert result.page_count == 1
    assert output.exists() and output.stat().st_size > 1_000
    assert "Hello World" in _all_docx_text(output)
    document = Document(output)
    assert all(section.bottom_margin.pt <= 24.0 for section in document.sections)


def test_convert_xlsx_with_styles_and_merged_cells(tmp_path: Path) -> None:
    source = tmp_path / "doanh-thu.xlsx"
    output = tmp_path / "doanh-thu-editable.docx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Doanh thu"
    sheet.append(["Tháng", "Doanh thu", "Tăng trưởng"])
    sheet.append(["Tháng 1", 1_250_000, 0.125])
    sheet.append(["Tháng 2", 1_480_000, 0.184])
    sheet.append(["Tổng cộng", "=SUM(B2:B3)", None])
    sheet.merge_cells("B4:C4")
    sheet["A1"].font = Font(bold=True, color="FFFFFF")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1E63E9")
    sheet["B2"].number_format = '#,##0 "VND"'
    sheet["C2"].number_format = "0.0%"
    sheet["B3"].number_format = '#,##0 "VND"'
    sheet["C3"].number_format = "0.0%"
    workbook.save(source)

    result = convert_excel(source, output, ConversionOptions())

    assert result.sheet_count == 1
    assert output.exists() and output.stat().st_size > 1_000
    document = Document(output)
    assert any(paragraph.text == "Doanh thu" for paragraph in document.paragraphs)
    assert document.tables
    text = _all_docx_text(output)
    assert "Tháng 1" in text
    assert "1,250,000" in text
    assert "12.5%" in text
    assert "=SUM(B2:B3)" in text
    table_xml = document.tables[0]._tbl.xml
    assert "w:tblW" in table_xml
    assert "w:tblGrid" in table_xml
    assert "w:tblHeader" in table_xml


def test_convert_image_to_editable_docx_with_explicit_word_spacing(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PIL import Image

    source = tmp_path / "photo.png"
    output = tmp_path / "photo-editable.docx"
    Image.new("RGB", (1200, 700), "white").save(source)

    def fake_image_to_data(_image, _language):
        return (
            {
                "text": ["HELLO", "WORLD", "!"],
                "conf": ["98", "97", "96"],
                "block_num": [1, 1, 1],
                "par_num": [1, 1, 1],
                "line_num": [1, 1, 1],
                "left": [120, 360, 590],
                "top": [180, 180, 180],
                "width": [190, 190, 20],
                "height": [72, 72, 72],
            },
            "eng",
            [],
        )

    monkeypatch.setattr("app.converters.raster_tables.image_to_data", fake_image_to_data)
    progress: list[int] = []
    result = convert_image(
        source,
        output,
        ConversionOptions(
            language="eng",
            progress_callback=lambda percent, _message: progress.append(percent),
        ),
    )

    assert result.page_count == 1
    assert output.exists() and output.stat().st_size > 1_000
    assert "HELLO WORLD!" in _all_docx_text(output)
    document = Document(output)
    assert document.styles["Normal"].font.name == "Times New Roman"
    assert document.paragraphs[0].runs[0].font.name == "Times New Roman"
    assert progress[0] >= 0
    assert progress[-1] == 100


def test_normalizes_common_vietnamese_ocr_noise_and_joined_words() -> None:
    cleaned = normalize_ocr_text(
        "GIÁY ĐÈ NGHỊ THANH TOÁN | srr | Sốhóa đơn | Ngàyhóa đơn | Nội đung"
    )

    assert "GIẤY ĐỀ NGHỊ THANH TOÁN" in cleaned
    assert "STT" in cleaned
    assert "Số hóa đơn" in cleaned
    assert "Ngày hóa đơn" in cleaned
    assert "Nội dung" in cleaned
    assert normalize_ocr_text("LÀ\nrenewable energy") == "DVA\nrenewable energy"
    assert normalize_ocr_text("Transe: 1⁄1") == "Trang: 1⁄1"
    assert normalize_ocr_text("(l, họ tên)") == "(ký, họ tên)"
    assert normalize_ocr_text("Số tiên băng chữ") == "Số tiền bằng chữ"
    assert normalize_ocr_text("+ Tÿ giá ngoại t§ (vàng bạc)") == "+ Tỷ giá ngoại tệ (vàng bạc)"
    assert normalize_ocr_text("TGƠĐ") == "TGĐ"
    assert normalize_ocr_text("THƯ CAM KÉT") == "THƯ CAM KẾT"
    assert normalize_ocr_text("Qúy khách hàng") == "Quý khách hàng"
    assert normalize_ocr_text("hiện nay. để đảm bảo") == "hiện nay, để đảm bảo"


def test_hocr_parser_keeps_tesseract_font_size() -> None:
    payload = b"""<?xml version='1.0' encoding='UTF-8'?>
    <html xmlns='http://www.w3.org/1999/xhtml'><body>
      <div class='ocr_page'><span class='ocr_line'>
        <span class='ocrx_word' title='bbox 10 20 70 55; x_wconf 97; x_fsize 12'>Hello</span>
        <span class='ocrx_word' title='bbox 82 20 150 55; x_wconf 96; x_fsize 12'>World</span>
      </span></div>
    </body></html>"""

    data = _hocr_to_data(payload)

    assert data["text"] == ["Hello", "World"]
    assert data["font_size"] == [12.0, 12.0]
    assert data["left"] == [10, 82]
    assert data["width"] == [60, 68]


def test_pixel_hocr_font_below_32_is_converted_to_points() -> None:
    line = {
        "height": 22,
        "median_height": 22,
        "words": [
            {"font_size": 29, "height": 21},
            {"font_size": 29, "height": 22},
        ],
    }

    assert _font_size_for_line(line, 72 / 200) == pytest.approx(9.1872)

    document = Document()
    rendered_line = {
        **line,
        "text": "Nội dung OCR",
        "left": 100,
        "top": 100,
        "right": 700,
        "bottom": 122,
        "confidence": 99,
        "word_count": 3,
    }
    rendered_line["words"] = [
        {
            "text": "Nội dung OCR",
            "font_size": 29,
            "height": 22,
            "left": 100,
            "top": 100,
            "right": 700,
            "bottom": 122,
        }
    ]
    write_ocr_lines(
        document,
        [rendered_line],
        image_width=1000,
        content_width_inches=6.5,
        source_dpi=200,
    )

    assert document.paragraphs[0].runs[0].font.size.pt == pytest.approx(9.0)


def test_ocr_fragments_on_same_baseline_become_one_editable_line() -> None:
    def fragment(text: str, left: int, right: int, top: int = 100) -> dict[str, object]:
        word = {
            "text": text,
            "confidence": 96,
            "left": left,
            "top": top,
            "right": right,
            "bottom": top + 30,
            "width": right - left,
            "height": 30,
            "font_size": 30,
        }
        return {
            "text": text,
            "left": left,
            "top": top,
            "right": right,
            "bottom": top + 30,
            "height": 30,
            "median_height": 30,
            "confidence": 96,
            "word_count": 1,
            "words": [word],
        }

    merged = merge_ocr_lines(
        [fragment("KTT", 100, 170, 104), fragment("Kế toán trưởng", 350, 620)],
        [],
    )

    assert len(merged) == 1
    assert merged[0]["text"] == "KTT Kế toán trưởng"


def test_sparse_ocr_duplicate_abbreviation_is_removed() -> None:
    def line(text: str, left: int, top: int, right: int) -> dict[str, object]:
        return {
            "text": text,
            "left": left,
            "top": top,
            "right": right,
            "bottom": top + 30,
            "height": 30,
            "median_height": 30,
            "confidence": 95,
            "word_count": len(text.split()),
            "words": [],
        }

    merged = merge_ocr_lines(
        [line("KTT: Kế toán trưởng", 100, 100, 500), line("KTIT", 100, 108, 180)],
        [],
    )

    assert [item["text"] for item in merged] == ["KTT: Kế toán trưởng"]


def test_short_backslash_ocr_artifact_is_removed() -> None:
    artifact = {
        "text": "\\ei",
        "left": 900,
        "top": 100,
        "right": 940,
        "bottom": 130,
        "height": 30,
        "median_height": 30,
        "confidence": 96,
        "word_count": 1,
        "words": [],
    }

    assert merge_ocr_lines([artifact], []) == []


def test_scanned_text_style_detection_reaches_docx_runs() -> None:
    from PIL import Image, ImageDraw, ImageFont

    regular_path = Path(r"C:\Windows\Fonts\times.ttf")
    bold_path = Path(r"C:\Windows\Fonts\timesbd.ttf")
    italic_path = Path(r"C:\Windows\Fonts\timesi.ttf")
    if not all(path.is_file() for path in (regular_path, bold_path, italic_path)):
        pytest.skip("Times New Roman test fonts are unavailable")

    image = Image.new("RGB", (1900, 760), "white")
    draw = ImageDraw.Draw(image)
    fonts = {
        "regular": ImageFont.truetype(str(regular_path), 64),
        "bold": ImageFont.truetype(str(bold_path), 64),
        "italic": ImageFont.truetype(str(italic_path), 64),
    }
    specifications = [
        ("regular words provide a stable normal style baseline", "regular", 50, False),
        ("strong phrase stays visibly bold", "bold", 220, False),
        ("italic sentence leans clearly right", "italic", 390, False),
        ("underlined label", "regular", 560, True),
    ]
    words: list[dict[str, object]] = []
    for line_number, (text, style, top, underline) in enumerate(specifications, 1):
        x = 90
        first_left: int | None = None
        last_right = x
        line_bottom = top
        for token in text.split():
            bbox = draw.textbbox((x, top), token, font=fonts[style])
            draw.text((x, top), token, font=fonts[style], fill="black")
            first_left = bbox[0] if first_left is None else first_left
            last_right = bbox[2]
            line_bottom = max(line_bottom, bbox[3])
            words.append(
                {
                    "text": token,
                    "confidence": 99,
                    "left": bbox[0],
                    "top": bbox[1],
                    "right": bbox[2],
                    "bottom": bbox[3],
                    "width": bbox[2] - bbox[0],
                    "height": bbox[3] - bbox[1],
                    "font_size": 16,
                    "line_key": (1, 1, line_number),
                }
            )
            x = bbox[2] + 28
        if underline and first_left is not None:
            draw.line((first_left, line_bottom + 5, last_right, line_bottom + 5), fill="black", width=4)

    lines = ocr_words_to_lines(words)
    styled = annotate_ocr_styles(image, lines)
    assert not any(word.get("bold") for word in styled[0]["words"])
    assert all(word.get("bold") for word in styled[1]["words"])
    assert styled[2]["italic"] is True
    assert all(word.get("underline") for word in styled[3]["words"])

    document = Document()
    write_ocr_lines(
        document,
        lines,
        image_width=image.width,
        content_width_inches=6.5,
        source_dpi=300,
        image=image,
    )
    paragraphs = {paragraph.text: paragraph for paragraph in document.paragraphs}
    assert all(run.bold for run in paragraphs[specifications[1][0]].runs if run.text.strip())
    assert all(run.italic for run in paragraphs[specifications[2][0]].runs if run.text.strip())
    assert all(run.underline for run in paragraphs[specifications[3][0]].runs if run.text.strip())


def test_wrapped_ocr_lines_become_real_justified_paragraphs() -> None:
    def line(text: str, left: int, top: int, right: int) -> dict[str, object]:
        word = {
            "text": text,
            "confidence": 99,
            "left": left,
            "top": top,
            "right": right,
            "bottom": top + 30,
            "width": right - left,
            "height": 30,
            "font_size": 12,
        }
        return {
            "text": text,
            "left": left,
            "top": top,
            "right": right,
            "bottom": top + 30,
            "height": 30,
            "median_height": 30,
            "confidence": 99,
            "word_count": len(text.split()),
            "words": [word],
        }

    lines = [
        line("First paragraph begins with an indented visual line", 160, 100, 900),
        line("and continues at the common body margin", 100, 140, 900),
        line("before ending on its final visual line.", 100, 180, 760),
        line("Second paragraph starts after a larger gap", 160, 250, 900),
        line("and also continues at the body margin.", 100, 290, 780),
    ]
    document = Document()

    write_ocr_lines(
        document,
        lines,
        image_width=1000,
        content_width_inches=8.5,
        source_dpi=100,
    )

    assert len(document.paragraphs) == 2
    assert document.paragraphs[0].alignment == 3
    assert document.paragraphs[1].alignment == 3
    assert document.paragraphs[0].paragraph_format.first_line_indent is not None
    assert "continues at the common body margin" in document.paragraphs[0].text


def test_restores_obvious_missing_stt_numbers() -> None:
    from app.converters.raster_tables import _fill_missing_stt_sequence

    cells = [
        ["STT", "Nội dung"],
        ["", "Dòng 1"],
        ["", "Dòng 2"],
        ["3", "Dòng 3"],
        ["CỘNG", ""],
    ]

    _fill_missing_stt_sequence(cells)

    assert [row[0] for row in cells] == ["STT", "1", "2", "3", "CỘNG"]


def test_restores_step_numbers_for_merged_flowchart_rows() -> None:
    from app.converters.raster_tables import _fill_missing_step_sequence

    cells = [
        ["Bước", "Trách nhiệm"],
        ["", "CBCNV"],
        ["", "KTT"],
        ["", ""],
        ["", ""],
        ["", "TGĐ"],
    ]
    merges = [(2, 0, 4, 0)]

    _fill_missing_step_sequence(cells, merges)

    assert [row[0] for row in cells] == ["Bước", "1", "2", "", "", "3"]


def test_final_ocr_cleanup_preserves_font_and_removes_short_stamp_noise() -> None:
    document = Document()
    corrected = document.add_paragraph()
    corrected_run = corrected.add_run("Số tiên băng chữ")
    corrected_run.font.name = "Times New Roman"
    document.add_paragraph("TẢ")
    table = document.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "TGƠĐ"

    normalized, removed = normalize_document_ocr_text(document)

    assert normalized >= 2
    assert removed == 1
    assert [paragraph.text for paragraph in document.paragraphs] == ["Số tiền bằng chữ"]
    assert document.paragraphs[0].runs[0].font.name == "Times New Roman"
    assert table.cell(0, 0).text == "TGĐ"


def test_sparse_ocr_completes_a_line_without_losing_better_accents() -> None:
    primary = {
        "text": "Họ và tên người đề nghị/đề xuất thanh",
        "left": 100,
        "top": 200,
        "right": 720,
        "bottom": 245,
        "height": 45,
        "confidence": 95,
        "word_count": 8,
    }
    sparse = {
        "text": "Họ và tên người đề nghị/đề xuât thanh toán",
        "left": 100,
        "top": 200,
        "right": 810,
        "bottom": 245,
        "height": 45,
        "confidence": 92,
        "word_count": 9,
    }
    noise = {
        "text": "Z",
        "left": 900,
        "top": 300,
        "right": 920,
        "bottom": 340,
        "height": 40,
        "confidence": 96,
        "word_count": 1,
    }

    merged = merge_ocr_lines([primary], [sparse, noise])

    assert [line["text"] for line in merged] == [
        "Họ và tên người đề nghị/đề xuất thanh toán"
    ]


def test_sparse_ocr_does_not_match_a_long_line_to_a_small_stamp() -> None:
    stamp = {
        "text": "7)",
        "left": 720,
        "top": 100,
        "right": 810,
        "bottom": 130,
        "height": 30,
        "confidence": 82,
        "word_count": 1,
    }
    primary = {
        "text": "+ Tới cả hóa đơn đêu gửi vào địa chỉ",
        "left": 200,
        "top": 128,
        "right": 920,
        "bottom": 180,
        "height": 52,
        "confidence": 88,
        "word_count": 10,
    }
    sparse = {
        "text": "+ Tất cả hóa đơn đều gửi vào địa ch",
        "left": 200,
        "top": 100,
        "right": 900,
        "bottom": 180,
        "height": 80,
        "confidence": 93,
        "word_count": 10,
    }

    merged = merge_ocr_lines([stamp, primary], [sparse])

    assert [line["text"] for line in merged] == [
        "+ Tất cả hóa đơn đều gửi vào địa chỉ"
    ]


def test_sparse_ocr_drops_a_short_label_already_in_the_primary_line() -> None:
    primary = {
        "text": "Hình thức thanh toán: ☐ Tiền mặt ☐ Chuyển khoản",
        "left": 100,
        "top": 200,
        "right": 950,
        "bottom": 250,
        "height": 50,
        "confidence": 95,
        "word_count": 9,
    }
    duplicate = {
        "text": "Tiền mặt",
        "left": 430,
        "top": 202,
        "right": 560,
        "bottom": 250,
        "height": 48,
        "confidence": 97,
        "word_count": 2,
    }

    merged = merge_ocr_lines([primary], [duplicate])

    assert [line["text"] for line in merged] == [primary["text"]]


def test_convert_ruled_table_image_to_editable_word_table(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PIL import Image, ImageDraw

    source = tmp_path / "table-scan.png"
    output = tmp_path / "table-scan-editable.docx"
    image = Image.new("RGB", (1200, 800), "white")
    draw = ImageDraw.Draw(image)
    x_lines = [80, 330, 700, 1120]
    y_lines = [80, 240, 430, 680]
    for x in x_lines:
        draw.line((x, y_lines[0], x, y_lines[-1]), fill="black", width=5)
    for y in y_lines:
        draw.line((x_lines[0], y, x_lines[-1], y), fill="black", width=5)
    image.save(source)

    values = [
        ("MUC", 135, 145),
        ("NOI DUNG", 425, 145),
        ("GHI CHU", 810, 145),
        ("1", 170, 315),
        ("Tam ung", 430, 315),
        ("Da duyet", 805, 315),
        ("2", 170, 515),
        ("Thanh toan", 430, 515),
        ("Hoan tat", 805, 515),
    ]

    def fake_image_to_data(_image, _language):
        return (
            {
                "text": [value[0] for value in values],
                "conf": ["98"] * len(values),
                "block_num": list(range(1, len(values) + 1)),
                "par_num": [1] * len(values),
                "line_num": [1] * len(values),
                "left": [value[1] for value in values],
                "top": [value[2] for value in values],
                "width": [95] * len(values),
                "height": [42] * len(values),
            },
            "vie+eng",
            [],
        )

    monkeypatch.setattr("app.converters.raster_tables.image_to_data", fake_image_to_data)
    monkeypatch.setattr(
        "app.converters.raster_tables.image_region_to_data",
        lambda *_args, **_kwargs: ({"text": [], "conf": []}, "vie", []),
    )
    result = convert_image(source, output, ConversionOptions(detect_tables=True))

    assert result.table_count == 1
    document = Document(output)
    assert len(document.tables) == 1
    table = document.tables[0]
    assert len(table.rows) == 3
    assert len(table.columns) == 3
    assert table.cell(0, 0).text == "MUC"
    assert table.cell(1, 1).text == "Tam ung"
    assert table.cell(2, 2).text == "Hoan tat"
    assert table.cell(1, 1).paragraphs[0].runs[0].font.name == "Times New Roman"
    assert "w:tblW" in table._tbl.xml
    assert "w:tblGrid" in table._tbl.xml


def test_low_confidence_ocr_autorotates_only_when_quality_improves(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PIL import Image

    from app.converters.raster_tables import _autorotate_low_confidence_image

    low_quality = {
        "text": [f"bad-{index}" for index in range(20)],
        "conf": ["50"] * 20,
    }
    high_quality = {
        "text": [f"word-{index}" for index in range(20)],
        "conf": ["92"] * 20,
    }
    monkeypatch.setattr("app.converters.raster_tables.detect_image_rotation", lambda _image: 270)
    monkeypatch.setattr(
        "app.converters.raster_tables.image_to_data",
        lambda _image, _language: (high_quality, "vie+eng", []),
    )

    original = Image.new("RGB", (600, 900), "white")
    rotated, data, warnings = _autorotate_low_confidence_image(original, low_quality, "vie+eng")

    assert rotated.size == (900, 600)
    assert data is high_quality
    assert warnings and "90°" in warnings[0]


@pytest.mark.skipif(not get_ocr_engine_info().available, reason="Tesseract OCR is not installed")
def test_image_ocr_with_tesseract(tmp_path: Path) -> None:
    from PIL import Image, ImageDraw, ImageFont

    source = tmp_path / "camera-document.png"
    output = tmp_path / "camera-document-editable.docx"
    image = Image.new("RGB", (1800, 700), "white")
    draw = ImageDraw.Draw(image)
    font = None
    for candidate in ["arial.ttf", "DejaVuSans.ttf"]:
        try:
            font = ImageFont.truetype(candidate, 76)
            break
        except OSError:
            continue
    if font is None:
        font = ImageFont.load_default()
    draw.text((100, 240), "IMAGE DOCUMENT 2026", fill="black", font=font)
    image.save(source)

    result = convert_image(source, output, ConversionOptions(language="eng"))

    assert result.page_count == 1
    assert "IMAGE DOCUMENT" in _all_docx_text(output).upper()


@pytest.mark.skipif(not get_ocr_engine_info().available, reason="Tesseract OCR is not installed")
def test_scanned_pdf_ocr(tmp_path: Path) -> None:
    from PIL import Image, ImageDraw, ImageFont

    image_path = tmp_path / "scan.png"
    image = Image.new("RGB", (1800, 800), "white")
    draw = ImageDraw.Draw(image)
    font = None
    for candidate in ["arial.ttf", "DejaVuSans.ttf"]:
        try:
            font = ImageFont.truetype(candidate, 72)
            break
        except OSError:
            continue
    if font is None:
        font = ImageFont.load_default()
    draw.text((120, 220), "SCANNED DOCUMENT 2026", fill="black", font=font)
    image.save(image_path)

    source = tmp_path / "scan.pdf"
    output = tmp_path / "scan-editable.docx"
    pdf = canvas.Canvas(str(source), pagesize=letter)
    pdf.drawImage(str(image_path), 45, 250, width=520, height=230)
    pdf.save()

    convert_pdf(source, output, ConversionOptions(ocr_mode="force", language="eng", dpi=300))
    assert "SCANNED DOCUMENT" in _all_docx_text(output).upper()


def test_auto_mode_converts_scanned_pdf_grid_to_word_table(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PIL import Image, ImageDraw

    image_path = tmp_path / "scanned-grid.png"
    scan = Image.new("RGB", (1200, 800), "white")
    draw = ImageDraw.Draw(scan)
    x_lines = [80, 330, 700, 1120]
    y_lines = [80, 240, 430, 680]
    for x in x_lines:
        draw.line((x, y_lines[0], x, y_lines[-1]), fill="black", width=5)
    for y in y_lines:
        draw.line((x_lines[0], y, x_lines[-1], y), fill="black", width=5)
    scan.save(image_path)

    source = tmp_path / "scanned-grid.pdf"
    output = tmp_path / "scanned-grid-editable.docx"
    pdf = canvas.Canvas(str(source), pagesize=letter)
    pdf.drawImage(str(image_path), 36, 144, width=540, height=360)
    pdf.save()

    # The embedded image is rasterized at 200 DPI. These boxes sit in the
    # detected cells after the PDF-to-image coordinate transform.
    values = [
        ("STT", 260, 950),
        ("NOI DUNG", 650, 950),
        ("KET QUA", 1210, 950),
        ("1", 270, 1180),
        ("Tam ung", 660, 1180),
        ("Dong y", 1210, 1180),
        ("2", 270, 1450),
        ("Thanh toan", 660, 1450),
        ("Hoan tat", 1210, 1450),
    ]

    def fake_image_to_data(_image, _language):
        return (
            {
                "text": [value[0] for value in values],
                "conf": ["98"] * len(values),
                "block_num": list(range(1, len(values) + 1)),
                "par_num": [1] * len(values),
                "line_num": [1] * len(values),
                "left": [value[1] for value in values],
                "top": [value[2] for value in values],
                "width": [110] * len(values),
                "height": [45] * len(values),
            },
            "vie+eng",
            [],
        )

    monkeypatch.setattr("app.converters.raster_tables.image_to_data", fake_image_to_data)
    monkeypatch.setattr(
        "app.converters.raster_tables.image_region_to_data",
        lambda *_args, **_kwargs: ({"text": [], "conf": []}, "vie", []),
    )
    result = convert_pdf(
        source,
        output,
        ConversionOptions(
            ocr_mode="auto",
            dpi=200,
            preserve_layout=True,
            preserve_images=True,
            detect_tables=True,
        ),
    )

    assert result.table_count == 1
    document = Document(output)
    assert len(document.tables) == 1
    assert document.tables[0].cell(1, 1).text == "Tam ung"
    assert "Tam ung" in _all_docx_text(output)
