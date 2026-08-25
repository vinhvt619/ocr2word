from __future__ import annotations

import io
import statistics
import warnings as python_warnings
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Inches, RGBColor
from PIL import Image

from .base import (
    LANDSCAPE_CONTENT_WIDTH_DXA,
    PORTRAIT_CONTENT_WIDTH_DXA,
    ConversionError,
    ConversionOptions,
    ConversionResult,
    MissingDependencyError,
    add_compact_page_break,
    clean_cell_text,
    create_document,
    make_word_table,
    set_run_font,
    widths_from_text,
)
from .layout_pdf import convert_pdf_with_layout, layout_converter_available, repair_raw_text_blocks
from .ocr_layout import normalize_document_ocr_text
from .raster_tables import write_editable_ocr_page

try:
    import pymupdf
except ImportError:  # pragma: no cover - exercised in dependency diagnostics
    pymupdf = None


@dataclass(frozen=True, slots=True)
class PdfAnalysis:
    page_count: int
    ocr_pages: tuple[int, ...]
    image_backed_pages: tuple[int, ...]

    @property
    def ocr_page_count(self) -> int:
        return len(self.ocr_pages)

    def to_dict(self) -> dict[str, object]:
        payload = asdict(self)
        payload["ocr_pages"] = list(self.ocr_pages)
        payload["image_backed_pages"] = list(self.image_backed_pages)
        payload["ocr_page_count"] = self.ocr_page_count
        return payload


def _rect_overlap_ratio(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    left, top = max(a[0], b[0]), max(a[1], b[1])
    right, bottom = min(a[2], b[2]), min(a[3], b[3])
    if right <= left or bottom <= top:
        return 0.0
    intersection = (right - left) * (bottom - top)
    area = max(1.0, (a[2] - a[0]) * (a[3] - a[1]))
    return intersection / area


def _font_name(pdf_font: str) -> str:
    name = (pdf_font or "").split("+")[-1]
    lowered = name.lower()
    if "times" in lowered or "nimbusroman" in lowered or "nimbusrom" in lowered:
        return "Times New Roman"
    if "courier" in lowered or "mono" in lowered or "nimbusmono" in lowered:
        return "Courier New"
    if "arial" in lowered or "helvetica" in lowered or "nimbussans" in lowered:
        return "Arial"
    if "calibri" in lowered:
        return "Calibri"
    return "Calibri"


def _span_text(span: dict[str, Any]) -> str:
    characters = span.get("chars")
    if characters:
        return "".join(str(character.get("c", "")) for character in characters)
    return str(span.get("text", ""))


def _extract_tables(page) -> list[dict[str, Any]]:
    if not hasattr(page, "find_tables"):
        return []
    try:
        with python_warnings.catch_warnings():
            python_warnings.simplefilter("ignore")
            finder = page.find_tables()
    except Exception:
        return []
    items: list[dict[str, Any]] = []
    for table in getattr(finder, "tables", []):
        try:
            matrix = table.extract()
        except Exception:
            continue
        if not matrix or not any(any(clean_cell_text(cell) for cell in row) for row in matrix):
            continue
        items.append({"kind": "table", "bbox": tuple(table.bbox), "rows": matrix})
    return items


def _base_font_size(text_dict: dict[str, Any]) -> float:
    sizes: list[float] = []
    for block in text_dict.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                if _span_text(span).strip():
                    sizes.append(float(span.get("size", 11)))
    return statistics.median(sizes) if sizes else 11.0


def _block_text(block: dict[str, Any]) -> str:
    lines: list[str] = []
    for line in block.get("lines", []):
        text = "".join(_span_text(span) for span in line.get("spans", []))
        if text.strip():
            lines.append(text.strip())
    return "\n".join(lines)


def _heading_style(block: dict[str, Any], base_size: float) -> str | None:
    text = _block_text(block)
    if not text or len(text) > 180 or text.count("\n") > 2:
        return None
    spans = [
        span
        for line in block.get("lines", [])
        for span in line.get("spans", [])
        if _span_text(span).strip()
    ]
    if not spans:
        return None
    max_size = max(float(span.get("size", base_size)) for span in spans)
    mostly_bold = sum(bool(int(span.get("flags", 0)) & 16) for span in spans) >= max(1, len(spans) // 2)
    if max_size >= base_size * 1.55:
        return "Heading 1"
    if max_size >= base_size * 1.28:
        return "Heading 2"
    if mostly_bold and len(text) <= 100:
        return "Heading 3"
    return None


def _write_text_block(doc, block: dict[str, Any], base_size: float, page_width: float) -> None:
    style = _heading_style(block, base_size)
    paragraph = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    lines = [
        line
        for line in block.get("lines", [])
        if any(_span_text(span).strip() for span in line.get("spans", []))
    ]
    for line_index, line in enumerate(lines):
        for span in line.get("spans", []):
            text = _span_text(span).replace("\x00", "")
            if not text:
                continue
            run = paragraph.add_run(text)
            flags = int(span.get("flags", 0))
            run.bold = bool(flags & 16)
            run.italic = bool(flags & 2)
            set_run_font(run, _font_name(str(span.get("font", ""))), float(span.get("size", 11)))
            try:
                color = int(span.get("color", 0)) & 0xFFFFFF
                red, green, blue = (color >> 16) & 255, (color >> 8) & 255, color & 255
                if red * 0.299 + green * 0.587 + blue * 0.114 > 238:
                    color = 0x000000
                run.font.color.rgb = RGBColor.from_string(f"{color:06X}")
            except (TypeError, ValueError):
                pass
        if line_index < len(lines) - 1:
            paragraph.add_run().add_break(WD_BREAK.LINE)

    bbox = tuple(block.get("bbox", (0, 0, page_width, 0)))
    block_width = max(0.0, bbox[2] - bbox[0])
    center_delta = abs(((bbox[0] + bbox[2]) / 2) - (page_width / 2))
    if block_width < page_width * 0.85 and center_delta < page_width * 0.07:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif bbox[2] > page_width * 0.92 and block_width < page_width * 0.6:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def _image_bytes_for_docx(block: dict[str, Any]) -> io.BytesIO | None:
    raw = block.get("image")
    if not raw:
        return None
    try:
        image = Image.open(io.BytesIO(raw))
        image.load()
        if image.mode not in {"RGB", "RGBA", "L"}:
            image = image.convert("RGB")
        stream = io.BytesIO()
        image.save(stream, format="PNG")
        stream.seek(0)
        return stream
    except Exception:
        return None


def _write_image_block(doc, block: dict[str, Any], content_width_inches: float) -> None:
    stream = _image_bytes_for_docx(block)
    if stream is None:
        return
    bbox = tuple(block.get("bbox", (0, 0, content_width_inches * 72, 0)))
    width_inches = max(0.5, min(content_width_inches, (bbox[2] - bbox[0]) / 72))
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        paragraph.add_run().add_picture(stream, width=Inches(width_inches))
    except Exception:
        paragraph._element.getparent().remove(paragraph._element)


def _write_pdf_table(doc, rows: list[list[Any]], content_width_dxa: int) -> None:
    normalized = [[clean_cell_text(value) for value in row] for row in rows]
    col_count = max((len(row) for row in normalized), default=1)
    normalized = [row + [""] * (col_count - len(row)) for row in normalized]
    widths = widths_from_text(normalized, content_width_dxa)
    table = make_word_table(doc, len(normalized), col_count, widths, header_rows=1)
    for row_index, row in enumerate(normalized):
        for col_index, value in enumerate(row):
            cell = table.cell(row_index, col_index)
            cell.text = value
            paragraph = cell.paragraphs[0]
            if row_index == 0:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                set_run_font(run, "Calibri", 9.5)
                run.bold = row_index == 0


def _digital_page_items(page, options: ConversionOptions) -> tuple[list[dict[str, Any]], float]:
    text_dict = page.get_text("rawdict", sort=True)
    repair_raw_text_blocks(text_dict.get("blocks", []))
    table_items = _extract_tables(page) if options.detect_tables else []
    table_bboxes = [item["bbox"] for item in table_items]
    items: list[dict[str, Any]] = list(table_items)
    page_area = max(1.0, page.rect.width * page.rect.height)
    for block in text_dict.get("blocks", []):
        bbox = tuple(block.get("bbox", (0, 0, 0, 0)))
        if any(_rect_overlap_ratio(bbox, table_bbox) > 0.28 for table_bbox in table_bboxes):
            continue
        if block.get("type") == 0 and _block_text(block):
            items.append({"kind": "text", "bbox": bbox, "block": block})
        elif block.get("type") == 1 and options.preserve_images:
            block_area = max(0.0, (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]))
            if block_area >= 2400 and block_area / page_area < 0.72:
                items.append({"kind": "image", "bbox": bbox, "block": block})
    items.sort(key=lambda item: (round(item["bbox"][1], 1), round(item["bbox"][0], 1)))
    return items, _base_font_size(text_dict)


def _largest_image_ratio(page) -> float:
    page_area = max(1.0, page.rect.width * page.rect.height)
    largest_image_ratio = 0.0
    try:
        for image in page.get_image_info():
            bbox = tuple(image.get("bbox", (0, 0, 0, 0)))
            image_area = max(0.0, (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]))
            largest_image_ratio = max(largest_image_ratio, image_area / page_area)
    except Exception:
        return 0.0
    return largest_image_ratio


def _should_ocr_page(page, compact_text: str, options: ConversionOptions) -> bool:
    if options.ocr_mode == "force":
        return True
    if options.ocr_mode == "text":
        return False
    largest_image_ratio = _largest_image_ratio(page)
    # A scan can already contain a poor hidden OCR layer. Treat the page as a
    # scan based on its visual structure instead of trusting that text layer;
    # otherwise garbled hidden text replaces the original page image.
    if largest_image_ratio >= 0.70:
        return True
    if not compact_text:
        return True
    if len(compact_text) >= options.min_text_chars:
        return False
    return largest_image_ratio >= 0.35


def analyze_pdf(input_path: Path, options: ConversionOptions) -> PdfAnalysis:
    """Inspect a PDF cheaply so the serverless queue can route heavy scan jobs."""

    options.validate()
    if pymupdf is None:
        raise MissingDependencyError("Thiếu PyMuPDF. Hãy chạy setup.ps1 hoặc cài các gói trong requirements.txt.")
    try:
        pdf = pymupdf.open(input_path)
    except Exception as exc:
        raise ConversionError(f"Không thể mở PDF: {exc}") from exc
    try:
        if pdf.page_count == 0:
            raise ConversionError("PDF không có trang nào.")
        ocr_pages: list[int] = []
        image_backed_pages: list[int] = []
        for page_index, page in enumerate(pdf):
            compact_text = "".join(page.get_text("text", sort=True).split())
            if _should_ocr_page(page, compact_text, options):
                ocr_pages.append(page_index + 1)
            if _largest_image_ratio(page) >= 0.70:
                image_backed_pages.append(page_index + 1)
        return PdfAnalysis(
            page_count=pdf.page_count,
            ocr_pages=tuple(ocr_pages),
            image_backed_pages=tuple(image_backed_pages),
        )
    finally:
        pdf.close()


def extract_pdf_pages(
    input_path: Path,
    output_path: Path,
    *,
    start_page: int,
    end_page: int,
) -> int:
    """Copy a one-based inclusive page range to a compact temporary PDF."""

    if pymupdf is None:
        raise MissingDependencyError("Thiếu PyMuPDF. Hãy chạy setup.ps1 hoặc cài các gói trong requirements.txt.")
    try:
        source = pymupdf.open(input_path)
    except Exception as exc:
        raise ConversionError(f"Không thể mở PDF: {exc}") from exc
    try:
        if start_page < 1 or end_page < start_page or end_page > source.page_count:
            raise ConversionError(
                f"Khoảng trang PDF không hợp lệ: {start_page}-{end_page}/{source.page_count}."
            )
        subset = pymupdf.open()
        try:
            subset.insert_pdf(source, from_page=start_page - 1, to_page=end_page - 1)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            subset.save(output_path, garbage=4, deflate=True)
        finally:
            subset.close()
        return end_page - start_page + 1
    finally:
        source.close()


def _write_ocr_page(
    doc,
    page,
    options: ConversionOptions,
    content_width_inches: float,
    content_width_dxa: int,
) -> tuple[int, int, list[str]]:
    pixmap = page.get_pixmap(dpi=options.dpi, colorspace=pymupdf.csRGB, alpha=False, annots=True)
    image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
    image.info["dpi"] = (options.dpi, options.dpi)
    return write_editable_ocr_page(
        doc,
        image,
        options,
        content_width_inches=content_width_inches,
        content_width_dxa=content_width_dxa,
        source_dpi=options.dpi,
    )


def convert_pdf(input_path: Path, output_path: Path, options: ConversionOptions) -> ConversionResult:
    options.validate()
    options.report_progress(2, "Đang mở và phân tích PDF…")
    if pymupdf is None:
        raise MissingDependencyError("Thiếu PyMuPDF. Hãy chạy setup.ps1 hoặc cài các gói trong requirements.txt.")
    try:
        pdf = pymupdf.open(input_path)
    except Exception as exc:
        raise ConversionError(f"Không thể mở PDF: {exc}") from exc
    if pdf.page_count == 0:
        pdf.close()
        raise ConversionError("PDF không có trang nào.")

    page_count = pdf.page_count
    result_warnings: list[str] = []
    options.report_progress(5, f"Đã đọc {page_count} trang PDF.")

    use_layout_converter = (
        options.preserve_layout
        and options.preserve_images
        and options.page_breaks
        and options.ocr_mode != "force"
        and layout_converter_available()
    )
    if use_layout_converter:
        scan_pages: list[int] = []
        if options.ocr_mode == "auto":
            for page_index, page in enumerate(pdf):
                raw_text = "".join(page.get_text("text", sort=True).split())
                if _should_ocr_page(page, raw_text, options):
                    scan_pages.append(page_index + 1)
                options.report_progress(
                    5 + ((page_index + 1) / page_count) * 15,
                    f"Đang phân tích trang {page_index + 1}/{page_count}…",
                )

        # A mixed PDF already has reliable positioned text on most pages.
        # Rebuilding the whole file just because a few pages are scans destroys
        # covers, TOCs, logos and page furniture. In automatic mode, keep the
        # mixed document on the high-fidelity layout path; scanned pages remain
        # exact page images while the real PDF text stays editable. Users who
        # explicitly choose force OCR still get an all-editable reconstruction.
        preserve_mixed_layout = bool(scan_pages) and len(scan_pages) < page_count
        if not scan_pages or preserve_mixed_layout:
            if preserve_mixed_layout:
                result_warnings.append(
                    f"Đã giữ nguyên {len(scan_pages)} trang scan/phức tạp dưới dạng ảnh để không mất "
                    "bìa, nền, logo hoặc chữ ký; các trang PDF có lớp chữ vẫn chỉnh sửa được. "
                    "Muốn nhận dạng lại mọi trang, hãy chọn 'Ưu tiên chỉnh sửa'."
                )
            pdf.close()
            try:
                options.report_progress(24, "Đang giữ bìa, hình ảnh và dựng lại bố cục PDF trong Word…")
                convert_pdf_with_layout(
                    input_path,
                    output_path,
                    detect_tables=options.detect_tables,
                )
                options.report_progress(96, "Đang hoàn thiện file Word…")
                options.report_progress(100, "Đã tạo xong file Word theo bố cục PDF.")
                return ConversionResult(
                    output_path=output_path,
                    warnings=result_warnings,
                    page_count=page_count,
                )
            except ConversionError as exc:
                result_warnings.append(
                    f"Không thể dùng chế độ giữ bố cục ({exc}); đã chuyển sang chế độ dựng nội dung."
                )
                try:
                    pdf = pymupdf.open(input_path)
                except Exception as reopen_exc:
                    raise ConversionError(f"Không thể mở lại PDF: {reopen_exc}") from reopen_exc
                options.report_progress(8, "Đang chuyển sang chế độ dựng nội dung PDF…")
        else:
            result_warnings.append(
                f"PDF gồm {len(scan_pages)} trang scan; đang OCR thành chữ và bảng Word có thể chỉnh sửa. "
                "Nền, dấu mộc và chữ ký trên ảnh có thể không được giữ nguyên."
            )
            options.report_progress(8, "Đang OCR chữ và bảng trên các trang scan…")

    first_page = pdf[0]
    page_width_inches = max(4.0, float(first_page.rect.width) / 72.0)
    page_height_inches = max(4.0, float(first_page.rect.height) / 72.0)
    landscape = page_width_inches > page_height_inches
    page_margin_inches = 0.75
    content_width_inches = max(3.0, page_width_inches - page_margin_inches * 2)
    content_width_dxa = round(content_width_inches * 1440)
    doc = create_document(
        input_path.name,
        landscape=landscape,
        body_font="Times New Roman",
        page_width_inches=page_width_inches,
        page_height_inches=page_height_inches,
        margin_inches=page_margin_inches,
    )
    detected_tables = 0
    ocr_page_count = 0
    try:
        for page_index, page in enumerate(pdf):
            options.report_progress(
                8 + (page_index / page_count) * 82,
                f"Đang xử lý trang {page_index + 1}/{page_count}…",
            )
            raw_text = "".join(page.get_text("text", sort=True).split())
            needs_ocr = _should_ocr_page(page, raw_text, options)
            if needs_ocr:
                ocr_page_count += 1
                line_count, table_count, ocr_warnings = _write_ocr_page(
                    doc,
                    page,
                    options,
                    content_width_inches,
                    content_width_dxa,
                )
                detected_tables += table_count
                result_warnings.extend(ocr_warnings)
                if line_count == 0 and table_count == 0:
                    result_warnings.append(f"Trang {page_index + 1}: OCR không tìm thấy văn bản đủ tin cậy.")
            else:
                items, base_size = _digital_page_items(page, options)
                if not items and options.ocr_mode == "text":
                    result_warnings.append(f"Trang {page_index + 1}: không có lớp văn bản để trích xuất.")
                for item in items:
                    if item["kind"] == "text":
                        _write_text_block(doc, item["block"], base_size, page.rect.width)
                    elif item["kind"] == "table":
                        _write_pdf_table(doc, item["rows"], content_width_dxa)
                    elif item["kind"] == "image":
                        _write_image_block(doc, item["block"], content_width_inches)
            if options.page_breaks and page_index < pdf.page_count - 1:
                add_compact_page_break(doc)
            options.report_progress(
                8 + ((page_index + 1) / page_count) * 82,
                f"Đã xử lý trang {page_index + 1}/{page_count}.",
            )
    finally:
        pdf.close()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    options.report_progress(95, "Đang ghi file Word…")
    if ocr_page_count:
        normalize_document_ocr_text(doc)
    try:
        doc.save(output_path)
    except Exception as exc:
        raise ConversionError(f"Không thể tạo file Word: {exc}") from exc
    options.report_progress(100, "Đã tạo xong file Word từ PDF.")
    return ConversionResult(
        output_path=output_path,
        warnings=list(dict.fromkeys(result_warnings)),
        page_count=page_count,
        table_count=detected_tables,
    )
