from __future__ import annotations

from dataclasses import dataclass, field
import io
import math
from pathlib import Path
import re
from typing import Any, Sequence

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PillowImage
from PIL import ImageOps, UnidentifiedImageError
import pymupdf

from .base import ConversionError, ConversionOptions, ConversionResult
from .ocr_layout import merge_ocr_lines, ocr_data_to_words, ocr_words_to_lines
from . import raster_tables
from app.services.ocr import detect_image_rotation, image_to_data

# ---------------------------------------------------------------------------
# Resource bounds and limits (as settled in plan)
# ---------------------------------------------------------------------------
MAX_PAGES = 200
MAX_PIXELS_PER_PAGE = 25_000_000
MAX_ROWS_PER_SHEET = 50_000
MAX_COLS_PER_SHEET = 256
MAX_CELLS_PER_WORKBOOK = 500_000
MAX_CELL_CHARS = 32_767

XML_ILLEGAL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f]")
INTEGER_RE = re.compile(r"^-?[1-9]\d{0,14}$")
EXCEL_SHEET_INVALID_CHARS_RE = re.compile(r"[\\/*?:\[\]]")

PDF_EXTENSIONS = {".pdf"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
DISALLOWED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


@dataclass(slots=True)
class ExtractedTable:
    page_num: int
    table_idx: int
    cells: list[list[str]]
    merges: list[tuple[int, int, int, int]]  # (top_row, left_col, bottom_row, right_col) 0-indexed
    source_type: str = "digital"  # "digital" or "raster"


@dataclass(slots=True)
class ExtractedTextLine:
    page_num: int
    line_num: int
    text: str


def _clean_sheet_name(desired: str, existing_names: set[str]) -> str:
    cleaned = EXCEL_SHEET_INVALID_CHARS_RE.sub("-", desired).strip()
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .-")
    if not cleaned:
        cleaned = "Sheet"
    base = cleaned[:31]
    candidate = base
    counter = 2
    lower_existing = {n.casefold() for n in existing_names}
    while candidate.casefold() in lower_existing:
        suffix = f"_{counter}"
        avail = 31 - len(suffix)
        candidate = f"{base[:avail]}{suffix}"
        counter += 1
    existing_names.add(candidate)
    return candidate


def format_cell_value(
    raw: Any,
    is_table_cell: bool = True,
    warnings: list[str] | None = None,
) -> tuple[Any, str]:
    """Format raw string according to the conservative value policy.

    Returns:
        (value, data_type) where data_type is 'n' (numeric) or 's' (literal string).
    """
    if raw is None:
        return "", "s"
    text = str(raw)
    if XML_ILLEGAL_CHARS_RE.search(text):
        text = XML_ILLEGAL_CHARS_RE.sub("", text)
        if warnings is not None and "Đã loại bỏ một số ký tự điều khiển không hợp lệ trong ô." not in warnings:
            warnings.append("Đã loại bỏ một số ký tự điều khiển không hợp lệ trong ô.")

    if not text:
        return "", "s"

    if len(text) > MAX_CELL_CHARS:
        raise ConversionError(
            f"Ô dữ liệu vượt quá giới hạn độ dài {len(text):,} ký tự "
            f"(giới hạn Excel là {MAX_CELL_CHARS:,} ký tự)."
        )

    # Explicit 0
    if text == "0":
        return 0, "n"
    # Preserved -0
    if text == "-0":
        return "-0", "s"

    # Only unambiguous plain signed/unsigned integers <= 15 digits with no leading zero
    if INTEGER_RE.match(text):
        try:
            return int(text), "n"
        except ValueError:
            return text, "s"

    # All formula-like strings (=, +, -, @) and ambiguous numbers/decimals/dates remain literal string
    return text, "s"


def _extract_digital_merges(tab: pymupdf.Table) -> list[tuple[int, int, int, int]]:
    """Determine reliable merged cell coordinates from a PyMuPDF Table object."""
    if not tab.rows:
        return []

    x_points: set[float] = set()
    y_points: set[float] = set()
    for row in tab.rows:
        for cell_bbox in row.cells:
            if cell_bbox is not None:
                x_points.add(round(cell_bbox[0], 1))
                x_points.add(round(cell_bbox[2], 1))
                y_points.add(round(cell_bbox[1], 1))
                y_points.add(round(cell_bbox[3], 1))

    if not x_points or not y_points:
        return []

    sorted_x = sorted(x_points)
    sorted_y = sorted(y_points)

    def find_col(val: float) -> int:
        return min(range(len(sorted_x)), key=lambda i: abs(sorted_x[i] - val))

    def find_row(val: float) -> int:
        return min(range(len(sorted_y)), key=lambda i: abs(sorted_y[i] - val))

    merges: list[tuple[int, int, int, int]] = []
    for r_idx, row in enumerate(tab.rows):
        for c_idx, cell_bbox in enumerate(row.cells):
            if cell_bbox is not None:
                left_col = find_col(cell_bbox[0])
                right_col = find_col(cell_bbox[2]) - 1
                top_row = find_row(cell_bbox[1])
                bottom_row = find_row(cell_bbox[3]) - 1
                if 0 <= top_row <= bottom_row < len(tab.rows) and 0 <= left_col <= right_col < tab.col_count:
                    if right_col > left_col or bottom_row > top_row:
                        merges.append((top_row, left_col, bottom_row, right_col))
    return merges


def _rgb_image(image: PillowImage.Image) -> PillowImage.Image:
    image = ImageOps.exif_transpose(image)
    if image.mode == "RGBA":
        background = PillowImage.new("RGB", image.size, "white")
        background.paste(image, mask=image.getchannel("A"))
        return background
    if image.mode not in {"RGB", "L"}:
        return image.convert("RGB")
    return image


def _check_page_pixels(width_pixels: int, height_pixels: int, page_label: str) -> None:
    total_pixels = width_pixels * height_pixels
    if total_pixels > MAX_PIXELS_PER_PAGE:
        raise ConversionError(
            f"{page_label} có kích thước {width_pixels}x{height_pixels} ({total_pixels:,} pixel), "
            f"vượt quá giới hạn an toàn {MAX_PIXELS_PER_PAGE:,} pixel của hệ thống. "
            "Vui lòng giảm DPI hoặc giảm kích thước ảnh."
        )


def _check_table_candidate_bounds(rows: int, cols: int, page_label: str) -> None:
    if rows > MAX_ROWS_PER_SHEET or cols > MAX_COLS_PER_SHEET or (rows * cols) > MAX_CELLS_PER_WORKBOOK:
        raise ConversionError(
            f"{page_label} phát hiện bảng có kích thước {rows}x{cols} ({rows * cols:,} ô), "
            f"vượt quá giới hạn bộ nhớ {MAX_CELLS_PER_WORKBOOK:,} ô. "
            "Vui lòng chia nhỏ tài liệu."
        )


def _extract_raster_page(
    image: PillowImage.Image,
    page_num: int,
    options: ConversionOptions,
    warnings: list[str],
) -> tuple[list[ExtractedTable], list[ExtractedTextLine]]:
    """Perform OCR and table detection on a raster image/page with fill_sequences=False."""
    _check_page_pixels(image.width, image.height, f"Trang/ảnh {page_num}")

    data, _language, ocr_warnings = image_to_data(image, options.language)
    warnings.extend(ocr_warnings)

    image, data, rotation_warnings = raster_tables._autorotate_low_confidence_image(
        image,
        data,
        options.language,
    )
    warnings.extend(rotation_warnings)

    words = ocr_data_to_words(data)
    detected_tables: list[raster_tables.RasterTable] = []

    if options.detect_tables:
        if raster_tables.cv2 is None or raster_tables.np is None:
            warnings.append("Thiếu OpenCV nên chưa thể nhận diện lưới bảng trong ảnh.")
        else:
            try:
                # Call detect_raster_tables with fill_sequences=False and early limit checking
                detected_tables = raster_tables.detect_raster_tables(
                    image,
                    words,
                    requested_language=options.language,
                    warnings=warnings,
                    fill_sequences=False,
                    candidate_limit_checker=lambda r, c: _check_table_candidate_bounds(r, c, f"Trang {page_num}"),
                )
            except ConversionError:
                raise
            except Exception:
                warnings.append(f"Trang {page_num}: Không thể phân tích lưới bảng; đã chuyển sang trích xuất văn bản.")

    # Validate candidate bounds
    for tab in detected_tables:
        _check_table_candidate_bounds(tab.row_count, tab.column_count, f"Trang {page_num}")

    extracted_tables: list[ExtractedTable] = []
    for idx, tab in enumerate(detected_tables, start=1):
        extracted_tables.append(
            ExtractedTable(
                page_num=page_num,
                table_idx=idx,
                cells=tab.cells,
                merges=tab.merges,
                source_type="raster",
            )
        )

    # Extract words outside detected tables
    def inside_any_table(word: dict[str, Any]) -> bool:
        cx = (word["left"] + word["right"]) / 2
        cy = (word["top"] + word["bottom"]) / 2
        return any(
            t.bbox[0] <= cx <= t.bbox[2] and t.bbox[1] <= cy <= t.bbox[3]
            for t in detected_tables
        )

    outside_words = [w for w in words if not inside_any_table(w)]
    lines = ocr_words_to_lines(outside_words)
    extracted_lines: list[ExtractedTextLine] = []
    for l_idx, line in enumerate(lines, start=1):
        line_text = " ".join(w["text"] for w in line.get("words", []) if w.get("text")).strip()
        if line_text:
            extracted_lines.append(ExtractedTextLine(page_num=page_num, line_num=l_idx, text=line_text))

    return extracted_tables, extracted_lines


def _build_excel_workbook(
    tables: list[ExtractedTable],
    text_lines: list[ExtractedTextLine],
    warnings: list[str],
) -> openpyxl.Workbook:
    """Build formatted Excel workbook following layout, safety, and bound policies."""
    wb = openpyxl.Workbook()
    # Remove initial default sheet once custom sheets are created
    default_sheet = wb.active
    existing_sheet_names: set[str] = set()
    total_allocated_cells = 0

    header_font = Font(name="Calibri", size=11, bold=True)
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    regular_font = Font(name="Calibri", size=11)
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )
    center_align = Alignment(vertical="center", wrap_text=True)

    # 1. Populate tables into worksheets
    single_table = len(tables) == 1
    for tab in tables:
        if single_table:
            sheet_title = _clean_sheet_name("Bảng 1", existing_sheet_names)
        else:
            sheet_title = _clean_sheet_name(f"T{tab.page_num}_B{tab.table_idx}", existing_sheet_names)

        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        num_rows = len(tab.cells)
        num_cols = max((len(row) for row in tab.cells), default=0)

        if num_rows > MAX_ROWS_PER_SHEET:
            raise ConversionError(
                f"Bảng {sheet_title} có {num_rows:,} dòng, vượt quá giới hạn Excel {MAX_ROWS_PER_SHEET:,} dòng."
            )
        if num_cols > MAX_COLS_PER_SHEET:
            raise ConversionError(
                f"Bảng {sheet_title} có {num_cols:,} cột, vượt quá giới hạn Excel {MAX_COLS_PER_SHEET:,} cột."
            )

        col_max_widths: dict[int, int] = {}

        for r_idx, row in enumerate(tab.cells, start=1):
            for c_idx, cell_value in enumerate(row, start=1):
                total_allocated_cells += 1
                if total_allocated_cells > MAX_CELLS_PER_WORKBOOK:
                    raise ConversionError(
                        f"Tổng số ô bảng tính ({total_allocated_cells:,}) vượt quá giới hạn {MAX_CELLS_PER_WORKBOOK:,} ô. "
                        "Vui lòng chia nhỏ tài liệu."
                    )
                val, dtype = format_cell_value(cell_value, is_table_cell=True, warnings=warnings)
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = val
                cell.data_type = dtype
                cell.alignment = center_align
                cell.border = thin_border
                cell.font = regular_font

                str_len = len(str(val))
                col_max_widths[c_idx] = max(col_max_widths.get(c_idx, 0), str_len)

        # Apply reliable merges
        for top, left, bottom, right in tab.merges:
            if 0 <= top <= bottom < num_rows and 0 <= left <= right < num_cols:
                # 1-indexed for openpyxl
                ws.merge_cells(
                    start_row=top + 1,
                    start_column=left + 1,
                    end_row=bottom + 1,
                    end_column=right + 1,
                )

        # Set column widths bounded between 10 and 50
        for col_idx, max_w in col_max_widths.items():
            letter = get_column_letter(col_idx)
            ws.column_dimensions[letter].width = max(10, min(50, max_w + 3))

    # 2. Populate text sheet if text lines exist
    if text_lines:
        text_sheet_name = _clean_sheet_name("Văn bản", existing_sheet_names)
        ws_text = wb.create_sheet(title=text_sheet_name)
        ws_text.views.sheetView[0].showGridLines = True

        # Header
        headers = ["Trang", "Dòng", "Nội dung"]
        for c_idx, h in enumerate(headers, start=1):
            cell = ws_text.cell(row=1, column=c_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = center_align

        total_allocated_cells += 3
        max_text_width = 15
        current_row = 2
        for item in text_lines:
            line_str = item.text
            chunks = [line_str[i : i + 30000] for i in range(0, max(1, len(line_str)), 30000)]
            for chunk in chunks:
                if current_row > MAX_ROWS_PER_SHEET:
                    raise ConversionError(
                        f"Sheet '{text_sheet_name}' vượt quá giới hạn {MAX_ROWS_PER_SHEET:,} dòng. "
                        "Vui lòng chia nhỏ tài liệu."
                    )
                total_allocated_cells += 3
                if total_allocated_cells > MAX_CELLS_PER_WORKBOOK:
                    raise ConversionError(
                        f"Tổng số ô bảng tính ({total_allocated_cells:,}) vượt quá giới hạn {MAX_CELLS_PER_WORKBOOK:,} ô. "
                        "Vui lòng chia nhỏ tài liệu."
                    )

                ws_text.cell(row=current_row, column=1, value=item.page_num).alignment = center_align
                ws_text.cell(row=current_row, column=2, value=item.line_num).alignment = center_align

                val, dtype = format_cell_value(chunk, is_table_cell=False, warnings=warnings)
                c_content = ws_text.cell(row=current_row, column=3)
                c_content.value = val
                c_content.data_type = dtype
                c_content.alignment = center_align
                max_text_width = max(max_text_width, min(80, len(str(val))))
                current_row += 1

        ws_text.column_dimensions["A"].width = 10
        ws_text.column_dimensions["B"].width = 10
        ws_text.column_dimensions["C"].width = min(80, max_text_width + 4)

    # 3. Clean up the default initial sheet
    if default_sheet is not None and default_sheet in wb.worksheets and len(wb.worksheets) > 1:
        wb.remove(default_sheet)

    return wb


def convert_to_excel(
    input_path: Path,
    output_path: Path,
    options: ConversionOptions,
) -> ConversionResult:
    """Convert a PDF or scanned image document directly to an Excel (.xlsx) workbook.

    Raises:
        ConversionError: If input format is invalid, limits exceeded, or no content found.
    """
    options.validate()
    if options.output_format != "xlsx":
        raise ConversionError(f"Định dạng xuất không khớp: yêu cầu xlsx nhưng nhận {options.output_format}.")

    suffix = input_path.suffix.lower()
    if suffix in DISALLOWED_EXCEL_EXTENSIONS:
        raise ConversionError(
            "Chuyển đổi sang Excel (.xlsx) chỉ nhận file PDF hoặc hình ảnh (không nhận file Excel làm đầu vào)."
        )
    if suffix not in PDF_EXTENSIONS and suffix not in IMAGE_EXTENSIONS:
        raise ConversionError(
            f"'{input_path.name}' không được hỗ trợ. Chuyển đổi sang Excel chỉ nhận PDF và hình ảnh."
        )

    options.report_progress(5, "Đang mở tài liệu…")
    warnings: list[str] = []
    extracted_tables: list[ExtractedTable] = []
    extracted_lines: list[ExtractedTextLine] = []
    cumulative_cells = 0
    total_source_pages = 0

    # -----------------------------------------------------------------------
    # Case A: PDF Document
    # -----------------------------------------------------------------------
    if suffix in PDF_EXTENSIONS:
        try:
            doc = pymupdf.open(input_path)
        except Exception as exc:
            raise ConversionError(f"Không thể mở file PDF: {exc}") from exc

        try:
            total_pages = len(doc)
            total_source_pages = total_pages
            if total_pages == 0:
                raise ConversionError("File PDF không có trang nào.")
            if total_pages > MAX_PAGES:
                raise ConversionError(
                    f"Tài liệu PDF có {total_pages} trang, vượt quá giới hạn {MAX_PAGES} trang cho xuất Excel. "
                    "Vui lòng chia nhỏ file."
                )

            options.report_progress(10, f"Đang xử lý PDF {total_pages} trang…")

            for page_index in range(total_pages):
                page_num = page_index + 1
                page = doc[page_index]
                page_progress = 10 + int((page_index / total_pages) * 80)
                options.report_progress(page_progress, f"Đang trích xuất trang {page_num}/{total_pages}…")

                digital_text = page.get_text().strip()
                images = page.get_images()

                tab_finder = None
                if options.detect_tables:
                    try:
                        tab_finder = page.find_tables()
                    except ConversionError:
                        raise
                    except Exception:
                        tab_finder = None

                has_digital_tables = bool(tab_finder and tab_finder.tables)
                if options.ocr_mode == "text":
                    has_digital_content = len(digital_text) > 0 or has_digital_tables
                else:
                    has_digital_content = len(digital_text) >= options.min_text_chars or has_digital_tables

                # If page has completely 0 text and 0 images, skip as empty page with documented warning
                if len(digital_text) == 0 and len(images) == 0:
                    warnings.append(f"Trang {page_num}: Trang rỗng không có nội dung chữ hoặc ảnh; đã bỏ qua.")
                    continue

                candidate_rects: list[pymupdf.Rect] = []
                if images:
                    seen_box_keys: set[tuple[float, float, float, float]] = set()
                    unique_xrefs = sorted({img_info[0] for img_info in images if img_info[0] > 0})
                    for xref in unique_xrefs:
                        try:
                            rects = page.get_image_rects(xref)
                        except Exception:
                            rects = []
                        for rect in rects:
                            box_key = (round(rect.x0, 1), round(rect.y0, 1), round(rect.x1, 1), round(rect.y1, 1))
                            if box_key not in seen_box_keys:
                                seen_box_keys.add(box_key)
                                candidate_rects.append(rect)
                    candidate_rects.sort(key=lambda r: (round(r.y0, 1), round(r.x0, 1)))

                # Text mode: check if page has no digital content but has images
                if options.ocr_mode == "text" and not has_digital_content:
                    if candidate_rects or images:
                        warnings.append(
                            f"Trang {page_num}: Không tìm thấy lớp chữ số (digital text); các vùng ảnh quét trên trang không được OCR ở chế độ 'Chỉ trích xuất lớp chữ'. "
                            "Vui lòng dùng chế độ Tự động hoặc OCR toàn bộ để đọc ảnh scan."
                        )
                    continue

                if options.ocr_mode == "force":
                    use_ocr = True
                elif options.ocr_mode == "text":
                    use_ocr = False
                else:  # auto
                    use_ocr = not has_digital_content

                page_tables: list[ExtractedTable] = []
                page_lines: list[ExtractedTextLine] = []

                if not use_ocr:
                    # 1. Digital table extraction
                    digital_table_bboxes: list[tuple[float, float, float, float]] = []
                    if options.detect_tables:
                        try:
                            if tab_finder is None:
                                tab_finder = page.find_tables()
                            if tab_finder.tables:
                                for t_idx, tab in enumerate(tab_finder.tables, start=1):
                                    _check_table_candidate_bounds(
                                        tab.row_count, tab.col_count, f"Trang {page_num}"
                                    )
                                    merges = _extract_digital_merges(tab)
                                    raw_cells = [
                                        [str(c if c is not None else "") for c in row]
                                        for row in tab.extract()
                                    ]
                                    page_tables.append(
                                        ExtractedTable(
                                            page_num=page_num,
                                            table_idx=t_idx,
                                            cells=raw_cells,
                                            merges=merges,
                                            source_type="digital",
                                        )
                                    )
                                    digital_table_bboxes.append(tab.bbox)
                        except ConversionError:
                            raise
                        except Exception:
                            warnings.append(f"Trang {page_num}: Không thể phân tích bảng số; giữ nguyên nội dung chữ.")

                    # 2. Extract non-table digital text with per-word geometry
                    digital_words = page.get_text("words")
                    outside_words: list[tuple] = []
                    for w in digital_words:
                        x0, y0, x1, y1 = w[:4]
                        cx = (x0 + x1) / 2
                        cy = (y0 + y1) / 2
                        in_tab = any(
                            tx0 <= cx <= tx1 and ty0 <= cy <= ty1
                            for tx0, ty0, tx1, ty1 in digital_table_bboxes
                        )
                        if not in_tab:
                            outside_words.append(w)

                    lines_map: dict[tuple[int, int], list[tuple]] = {}
                    for w in outside_words:
                        lines_map.setdefault((w[5], w[6]), []).append(w)

                    for key in sorted(lines_map.keys()):
                        line_words = lines_map[key]
                        line_words.sort(key=lambda item: item[0])
                        line_text = " ".join(item[4] for item in line_words).strip()
                        if line_text:
                            page_lines.append(
                                ExtractedTextLine(
                                    page_num=page_num,
                                    line_num=len(page_lines) + 1,
                                    text=line_text,
                                )
                            )

                    # 3. Inspect substantial uncovered raster regions
                    if candidate_rects:
                        for rect in candidate_rects:
                            rw = abs(rect.x1 - rect.x0)
                            rh = abs(rect.y1 - rect.y0)
                            area = rw * rh
                            # Substantial image: >= 15,000 pt^2 and min side >= 50 pt
                            if area >= 15000 and min(rw, rh) >= 50:
                                covered = any(
                                    tx0 <= rect.x0 and ty0 <= rect.y0 and tx1 >= rect.x1 and ty1 >= rect.y1
                                    for tx0, ty0, tx1, ty1 in digital_table_bboxes
                                )
                                if not covered:
                                    if options.ocr_mode == "text":
                                        warnings.append(
                                            f"Trang {page_num}: Phát hiện vùng ảnh quét ({int(rw)}x{int(rh)} pt) không được trích xuất do đang ở chế độ 'Chỉ trích xuất lớp chữ' (text). "
                                            "Vui lòng dùng chế độ Tự động để OCR vùng ảnh này."
                                        )
                                    else:
                                        try:
                                            rect_w = math.ceil(rw * (options.dpi / 72.0))
                                            rect_h = math.ceil(rh * (options.dpi / 72.0))
                                            _check_page_pixels(rect_w, rect_h, f"Trang {page_num} (vùng ảnh)")
                                            pix = page.get_pixmap(dpi=options.dpi, clip=rect)
                                            raster_img = PillowImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
                                            r_tables, r_lines = _extract_raster_page(raster_img, page_num, options, warnings)
                                            for r_t in r_tables:
                                                r_t.table_idx = len(page_tables) + 1
                                                page_tables.append(r_t)
                                            for r_l in r_lines:
                                                r_l.line_num = len(page_lines) + 1
                                                page_lines.append(r_l)
                                        except ConversionError:
                                            raise
                                        except Exception as exc:
                                            warnings.append(f"Trang {page_num}: Không thể OCR vùng ảnh quét: {exc}")

                else:
                    # Scan extraction
                    rect = page.rect
                    render_w = math.ceil(rect.width * (options.dpi / 72.0))
                    render_h = math.ceil(rect.height * (options.dpi / 72.0))
                    _check_page_pixels(render_w, render_h, f"Trang {page_num}")
                    pix = page.get_pixmap(dpi=options.dpi)
                    img = PillowImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    page_tables, page_lines = _extract_raster_page(img, page_num, options, warnings)

                # Enforce cumulative cell budget before accumulating page results
                new_table_cells = sum(
                    len(t.cells) * max((len(r) for r in t.cells), default=0)
                    for t in page_tables
                )
                new_line_cells = sum(
                    max(1, math.ceil(len(l.text) / 30000)) * 3
                    for l in page_lines
                )
                if cumulative_cells + new_table_cells + new_line_cells > MAX_CELLS_PER_WORKBOOK:
                    raise ConversionError(
                        f"Tổng số ô bảng tính ({cumulative_cells + new_table_cells + new_line_cells:,}) vượt quá giới hạn {MAX_CELLS_PER_WORKBOOK:,} ô. "
                        "Vui lòng chia nhỏ tài liệu."
                    )
                cumulative_cells += new_table_cells + new_line_cells
                extracted_tables.extend(page_tables)
                extracted_lines.extend(page_lines)

        finally:
            doc.close()

    # -----------------------------------------------------------------------
    # Case B: Raster Image Document
    # -----------------------------------------------------------------------
    elif suffix in IMAGE_EXTENSIONS:
        if options.ocr_mode == "text":
            raise ConversionError(
                "Hình ảnh chỉ chứa dữ liệu điểm ảnh (raster) nên không thể trích xuất ở chế độ 'Chỉ trích xuất lớp chữ' (text). "
                "Vui lòng chọn chế độ Tự động hoặc OCR toàn bộ."
            )
        try:
            source = PillowImage.open(input_path)
        except Exception as exc:
            raise ConversionError(f"Không thể mở file ảnh: {exc}") from exc

        try:
            frame_count = getattr(source, "n_frames", 1)
            total_source_pages = frame_count
            if frame_count > MAX_PAGES:
                raise ConversionError(
                    f"File ảnh có {frame_count} khung hình, vượt quá giới hạn {MAX_PAGES} trang. "
                    "Vui lòng chia nhỏ file."
                )

            options.report_progress(10, f"Đang xử lý ảnh {frame_count} khung…")

            for frame_index in range(frame_count):
                page_num = frame_index + 1
                source.seek(frame_index)
                _check_page_pixels(source.width, source.height, f"Khung ảnh {page_num}")
                frame = _rgb_image(source.copy())
                frame_progress = 10 + int((frame_index / frame_count) * 80)
                options.report_progress(frame_progress, f"Đang OCR ảnh/khung {page_num}/{frame_count}…")

                r_tables, r_lines = _extract_raster_page(frame, page_num, options, warnings)

                new_table_cells = sum(
                    len(t.cells) * max((len(r) for r in t.cells), default=0)
                    for t in r_tables
                )
                new_line_cells = sum(
                    max(1, math.ceil(len(l.text) / 30000)) * 3
                    for l in r_lines
                )
                if cumulative_cells + new_table_cells + new_line_cells > MAX_CELLS_PER_WORKBOOK:
                    raise ConversionError(
                        f"Tổng số ô bảng tính ({cumulative_cells + new_table_cells + new_line_cells:,}) vượt quá giới hạn {MAX_CELLS_PER_WORKBOOK:,} ô. "
                        "Vui lòng chia nhỏ tài liệu."
                    )
                cumulative_cells += new_table_cells + new_line_cells
                extracted_tables.extend(r_tables)
                extracted_lines.extend(r_lines)
        finally:
            source.close()

    # -----------------------------------------------------------------------
    # Validation of Recoverable Content
    # -----------------------------------------------------------------------
    has_useful_content = any(
        any(c.strip() for c in row)
        for t in extracted_tables
        for row in t.cells
    ) or any(l.text.strip() for l in extracted_lines)

    if not has_useful_content:
        if options.ocr_mode == "text":
            raise ConversionError(
                "Tài liệu không có lớp chữ số (digital text) để trích xuất ở chế độ 'Chỉ trích xuất lớp chữ' (text). "
                "Vui lòng chuyển sang chế độ Tự động (auto) hoặc OCR toàn bộ (force) để nhận diện nội dung quét."
            )
        raise ConversionError(
            "Tài liệu không có nội dung chữ hoặc bảng hợp lệ để chuyển đổi sang Excel."
        )

    if not extracted_tables:
        if any("không được trích xuất" in w or "không được OCR" in w or "đã bỏ qua" in w for w in warnings):
            warnings.append(
                "Không phát hiện được bảng trong tài liệu; đã xuất nội dung văn bản trích xuất được vào bảng tính."
            )
        else:
            warnings.append(
                "Không phát hiện được bảng trong tài liệu; đã xuất toàn bộ nội dung văn bản vào bảng tính."
            )

    # -----------------------------------------------------------------------
    # Build & Save Workbook
    # -----------------------------------------------------------------------
    options.report_progress(92, "Đang đóng gói bảng tính Excel…")
    wb = _build_excel_workbook(extracted_tables, extracted_lines, warnings)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(output_path)
    except Exception as exc:
        raise ConversionError(f"Không thể ghi file Excel: {exc}") from exc
    finally:
        wb.close()

    options.report_progress(100, "Đã xuất xong file Excel.")
    return ConversionResult(
        output_path=output_path,
        warnings=warnings,
        page_count=max(1, total_source_pages),
        sheet_count=len(wb.sheetnames),
        table_count=len(extracted_tables),
    )
