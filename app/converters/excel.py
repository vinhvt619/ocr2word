from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Any, Iterable

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor

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
    set_cell_shading,
    set_run_font,
)

try:
    import openpyxl
    from openpyxl.styles.numbers import is_date_format
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    openpyxl = None
    is_date_format = None
    get_column_letter = None


MAX_EXCEL_ROWS = 10_000
MAX_EXCEL_COLUMNS = 256
MAX_EXCEL_CELLS = 750_000


def _actual_bounds(*worksheets) -> tuple[int, int, int, int] | None:
    coordinates: list[tuple[int, int]] = []
    for worksheet in worksheets:
        cells = getattr(worksheet, "_cells", {}).values()
        for cell in cells:
            if getattr(cell, "value", None) is not None:
                coordinates.append((cell.row, cell.column))
    if not coordinates:
        return None
    rows, columns = zip(*coordinates)
    return min(rows), max(rows), min(columns), max(columns)


def _format_number(value: float | int, number_format: str) -> str:
    fmt = number_format or ""
    if "%" in fmt:
        decimals = 0
        if "." in fmt:
            decimals = len(fmt.split(".", 1)[1].split("%", 1)[0])
        return f"{value * 100:,.{decimals}f}%"
    if any(symbol in fmt for symbol in ("$", "₫", "€", "£")):
        symbol = next((symbol for symbol in ("₫", "$", "€", "£") if symbol in fmt), "")
        decimals = 2 if ".00" in fmt else 0
        return f"{value:,.{decimals}f} {symbol}".strip()
    if isinstance(value, int) or (isinstance(value, float) and value.is_integer()):
        return f"{int(value):,}" if "," in fmt else str(int(value))
    decimals = 2
    if "." in fmt:
        tail = fmt.split(".", 1)[1]
        count = sum(char in {"0", "#"} for char in tail)
        if count:
            decimals = min(count, 8)
    return f"{value:,.{decimals}f}" if "," in fmt else f"{value:.{decimals}f}"


def _cell_text(formula_cell, value_cell) -> str:
    value = value_cell.value
    if value is None:
        value = formula_cell.value
    if value is None:
        return ""
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        if isinstance(value, dt.datetime):
            return value.strftime("%d/%m/%Y %H:%M")
        if isinstance(value, dt.date):
            return value.strftime("%d/%m/%Y")
        return value.strftime("%H:%M:%S")
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        fmt = getattr(formula_cell, "number_format", "") or ""
        if is_date_format and is_date_format(fmt):
            return str(value)
        return _format_number(value, fmt)
    return clean_cell_text(value)


def _rgb_from_openpyxl(color) -> str | None:
    if color is None or getattr(color, "type", None) != "rgb":
        return None
    rgb = getattr(color, "rgb", None)
    if not rgb:
        return None
    return str(rgb)[-6:].upper()


def _apply_excel_style(word_cell, excel_cell, *, displayed_value: object, is_header: bool) -> None:
    paragraph = word_cell.paragraphs[0]
    horizontal = getattr(excel_cell.alignment, "horizontal", None)
    if horizontal in {"center", "centerContinuous"}:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif horizontal == "right":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif horizontal == "justify":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif isinstance(displayed_value, (int, float)) and not isinstance(displayed_value, bool):
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    fill_type = getattr(excel_cell.fill, "fill_type", None)
    fill = _rgb_from_openpyxl(getattr(excel_cell.fill, "fgColor", None))
    if fill_type == "solid" and fill and fill not in {"000000", "FFFFFF"}:
        set_cell_shading(word_cell, fill)

    for run in paragraph.runs:
        font = excel_cell.font
        set_run_font(run, getattr(font, "name", None) or "Calibri", getattr(font, "sz", None) or 9.5)
        run.bold = bool(getattr(font, "bold", False)) or is_header
        run.italic = bool(getattr(font, "italic", False))
        color = _rgb_from_openpyxl(getattr(font, "color", None))
        if color:
            run.font.color.rgb = RGBColor.from_string(color)


def _column_widths(worksheet, columns: list[int], content_width_dxa: int) -> list[int]:
    weights: list[float] = []
    for column in columns:
        letter = get_column_letter(column)
        width = worksheet.column_dimensions[letter].width
        weights.append(max(4.0, min(35.0, float(width or 10))))
    total = sum(weights) or 1
    widths = [max(420, round(content_width_dxa * weight / total)) for weight in weights]
    widths[-1] += content_width_dxa - sum(widths)
    return widths


def _row_groups(min_row: int, max_row: int, max_rows_per_table: int) -> Iterable[list[int]]:
    body_start = min_row + 1
    if body_start > max_row:
        yield [min_row]
        return
    for start in range(body_start, max_row + 1, max_rows_per_table - 1):
        yield [min_row, *range(start, min(max_row + 1, start + max_rows_per_table - 1))]


def _merge_contained_cells(table, worksheet, row_indices: list[int], columns: list[int]) -> None:
    row_positions = {row: index for index, row in enumerate(row_indices)}
    col_positions = {column: index for index, column in enumerate(columns)}
    for merged_range in worksheet.merged_cells.ranges:
        rows = list(range(merged_range.min_row, merged_range.max_row + 1))
        cols = list(range(merged_range.min_col, merged_range.max_col + 1))
        if not all(row in row_positions for row in rows) or not all(col in col_positions for col in cols):
            continue
        top, bottom = row_positions[rows[0]], row_positions[rows[-1]]
        left, right = col_positions[cols[0]], col_positions[cols[-1]]
        if top == bottom and left == right:
            continue
        try:
            merged_cell = table.cell(top, left).merge(table.cell(bottom, right))
            for paragraph in list(merged_cell.paragraphs)[1:]:
                if not paragraph.text.strip():
                    paragraph._element.getparent().remove(paragraph._element)
        except (IndexError, ValueError):
            continue


def _write_openpyxl_sheet(
    doc,
    formula_ws,
    value_ws,
    bounds: tuple[int, int, int, int],
    options: ConversionOptions,
    content_width_dxa: int,
) -> None:
    min_row, max_row, min_col, max_col = bounds
    doc.add_paragraph(formula_ws.title, style="Heading 1")
    max_cols = options.max_excel_columns_per_table
    for col_start in range(min_col, max_col + 1, max_cols):
        columns = list(range(col_start, min(max_col + 1, col_start + max_cols)))
        if max_col - min_col + 1 > max_cols:
            start_label = get_column_letter(columns[0])
            end_label = get_column_letter(columns[-1])
            doc.add_paragraph(f"Cột {start_label}-{end_label}", style="Heading 3")
        widths = _column_widths(formula_ws, columns, content_width_dxa)
        for row_indices in _row_groups(min_row, max_row, options.max_excel_rows_per_table):
            table = make_word_table(doc, len(row_indices), len(columns), widths, header_rows=1)
            for table_row, source_row in enumerate(row_indices):
                for table_col, source_col in enumerate(columns):
                    formula_cell = formula_ws.cell(source_row, source_col)
                    value_cell = value_ws.cell(source_row, source_col)
                    word_cell = table.cell(table_row, table_col)
                    word_cell.text = _cell_text(formula_cell, value_cell)
                    displayed_value = value_cell.value if value_cell.value is not None else formula_cell.value
                    _apply_excel_style(
                        word_cell,
                        formula_cell,
                        displayed_value=displayed_value,
                        is_header=table_row == 0,
                    )
            _merge_contained_cells(table, formula_ws, row_indices, columns)
            doc.add_paragraph().paragraph_format.space_after = 0


def _convert_xlsx(input_path: Path, output_path: Path, options: ConversionOptions) -> ConversionResult:
    if openpyxl is None:
        raise MissingDependencyError("Thiếu openpyxl để đọc file Excel.")
    options.report_progress(3, "Đang mở workbook Excel…")
    try:
        formula_book = openpyxl.load_workbook(input_path, data_only=False, read_only=False)
        value_book = openpyxl.load_workbook(input_path, data_only=True, read_only=False)
    except Exception as exc:
        raise ConversionError(f"Không thể mở file Excel: {exc}") from exc

    sheets: list[tuple[Any, Any, tuple[int, int, int, int] | None]] = []
    landscape = False
    for formula_ws in formula_book.worksheets:
        value_ws = value_book[formula_ws.title]
        bounds = _actual_bounds(formula_ws, value_ws)
        sheets.append((formula_ws, value_ws, bounds))
        if bounds and bounds[3] - bounds[2] + 1 > 6:
            landscape = True

    doc = create_document(input_path.name, landscape=landscape)
    content_width = LANDSCAPE_CONTENT_WIDTH_DXA if landscape else PORTRAIT_CONTENT_WIDTH_DXA
    warnings: list[str] = []
    nonempty_count = 0
    for index, (formula_ws, value_ws, bounds) in enumerate(sheets):
        options.report_progress(
            10 + (index / max(1, len(sheets))) * 78,
            f"Đang xử lý sheet {index + 1}/{len(sheets)}: {formula_ws.title}",
        )
        if bounds is None:
            warnings.append(f"Sheet '{formula_ws.title}' trống nên đã được bỏ qua.")
            options.report_progress(
                10 + ((index + 1) / max(1, len(sheets))) * 78,
                f"Đã bỏ qua sheet trống: {formula_ws.title}",
            )
            continue
        min_row, max_row, min_col, max_col = bounds
        rows = max_row - min_row + 1
        cols = max_col - min_col + 1
        if rows > MAX_EXCEL_ROWS or cols > MAX_EXCEL_COLUMNS or rows * cols > MAX_EXCEL_CELLS:
            formula_book.close()
            value_book.close()
            raise ConversionError(
                f"Sheet '{formula_ws.title}' quá lớn để đưa vào Word ({rows:,} dòng x {cols:,} cột)."
            )
        if nonempty_count and options.page_breaks:
            add_compact_page_break(doc)
        _write_openpyxl_sheet(doc, formula_ws, value_ws, bounds, options, content_width)
        nonempty_count += 1
        if getattr(formula_ws, "_charts", None):
            warnings.append(f"Sheet '{formula_ws.title}': biểu đồ không được đưa vào Word.")
        if getattr(formula_ws, "_images", None):
            warnings.append(f"Sheet '{formula_ws.title}': hình nhúng không được đưa vào Word.")
        options.report_progress(
            10 + ((index + 1) / max(1, len(sheets))) * 78,
            f"Đã xử lý sheet {index + 1}/{len(sheets)}.",
        )

    formula_book.close()
    value_book.close()
    if nonempty_count == 0:
        doc.add_paragraph("Workbook không có dữ liệu.")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    options.report_progress(94, "Đang ghi bảng Excel vào file Word…")
    doc.save(output_path)
    options.report_progress(100, "Đã tạo xong file Word từ Excel.")
    return ConversionResult(output_path, list(dict.fromkeys(warnings)), sheet_count=nonempty_count)


def _convert_xls(input_path: Path, output_path: Path, options: ConversionOptions) -> ConversionResult:
    try:
        import xlrd
    except ImportError as exc:
        raise MissingDependencyError("Thiếu xlrd để đọc định dạng Excel .xls cũ.") from exc
    try:
        options.report_progress(3, "Đang mở workbook Excel 97-2003…")
        workbook = xlrd.open_workbook(input_path, formatting_info=False)
    except Exception as exc:
        raise ConversionError(f"Không thể mở file .xls: {exc}") from exc
    landscape = any(sheet.ncols > 6 for sheet in workbook.sheets())
    content_width = LANDSCAPE_CONTENT_WIDTH_DXA if landscape else PORTRAIT_CONTENT_WIDTH_DXA
    doc = create_document(input_path.name, landscape=landscape)
    converted = 0
    sheets = workbook.sheets()
    for sheet_index, sheet in enumerate(sheets):
        options.report_progress(
            10 + (sheet_index / max(1, len(sheets))) * 78,
            f"Đang xử lý sheet {sheet_index + 1}/{len(sheets)}: {sheet.name}",
        )
        if not sheet.nrows or not sheet.ncols:
            continue
        if sheet.nrows > MAX_EXCEL_ROWS or sheet.ncols > MAX_EXCEL_COLUMNS:
            raise ConversionError(f"Sheet '{sheet.name}' quá lớn để đưa vào Word.")
        if converted and options.page_breaks:
            add_compact_page_break(doc)
        doc.add_paragraph(sheet.name, style="Heading 1")
        for col_start in range(0, sheet.ncols, options.max_excel_columns_per_table):
            columns = list(range(col_start, min(sheet.ncols, col_start + options.max_excel_columns_per_table)))
            if sheet.ncols > options.max_excel_columns_per_table:
                doc.add_paragraph(
                    f"Cột {get_column_letter(columns[0] + 1)}-{get_column_letter(columns[-1] + 1)}",
                    style="Heading 3",
                )
            widths = [content_width // len(columns)] * len(columns)
            widths[-1] += content_width - sum(widths)
            for row_indices in _row_groups(0, sheet.nrows - 1, options.max_excel_rows_per_table):
                table = make_word_table(doc, len(row_indices), len(columns), widths, header_rows=1)
                for target_row, source_row in enumerate(row_indices):
                    for target_col, source_col in enumerate(columns):
                        value = sheet.cell_value(source_row, source_col)
                        table.cell(target_row, target_col).text = clean_cell_text(value)
                        for run in table.cell(target_row, target_col).paragraphs[0].runs:
                            set_run_font(run, "Calibri", 9.5)
                            run.bold = target_row == 0
        converted += 1
        options.report_progress(
            10 + ((sheet_index + 1) / max(1, len(sheets))) * 78,
            f"Đã xử lý sheet {sheet_index + 1}/{len(sheets)}.",
        )
    if not converted:
        doc.add_paragraph("Workbook không có dữ liệu.")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    options.report_progress(94, "Đang ghi bảng Excel vào file Word…")
    doc.save(output_path)
    options.report_progress(100, "Đã tạo xong file Word từ Excel.")
    return ConversionResult(output_path, sheet_count=converted)


def convert_excel(input_path: Path, output_path: Path, options: ConversionOptions) -> ConversionResult:
    options.validate()
    suffix = input_path.suffix.lower()
    if suffix == ".xls":
        return _convert_xls(input_path, output_path, options)
    if suffix in {".xlsx", ".xlsm"}:
        return _convert_xlsx(input_path, output_path, options)
    raise ConversionError(f"Định dạng Excel không được hỗ trợ: {suffix}")
