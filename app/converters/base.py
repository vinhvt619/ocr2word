from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Literal, Sequence

from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.table import Table, _Cell


PORTRAIT_CONTENT_WIDTH_DXA = 9360
LANDSCAPE_CONTENT_WIDTH_DXA = 12960
TABLE_INDENT_DXA = 120
TABLE_CELL_MARGIN_DXA = {"top": 80, "start": 120, "bottom": 80, "end": 120}


class ConversionError(RuntimeError):
    """A user-facing conversion error."""


class MissingDependencyError(ConversionError):
    """Raised when an optional conversion dependency is unavailable."""


class MissingOcrEngineError(ConversionError):
    """Raised when OCR is required but Tesseract is unavailable."""


MathEngineName = Literal["rules", "local_hybrid"]
ALLOWED_MATH_ENGINES: tuple[str, ...] = ("rules", "local_hybrid")


@dataclass(slots=True)
class ConversionOptions:
    output_format: Literal["docx", "xlsx"] = "docx"
    language: str = "vie"
    ocr_mode: Literal["auto", "force", "text"] = "auto"
    dpi: int = 300
    page_breaks: bool = True
    preserve_images: bool = True
    preserve_layout: bool = True
    detect_tables: bool = True
    math_ocr: bool = False
    math_engine: str | None = None
    min_text_chars: int = 40
    max_excel_columns_per_table: int = 10
    max_excel_rows_per_table: int = 180
    progress_callback: Callable[[int, str], None] | None = field(
        default=None,
        repr=False,
        compare=False,
    )

    def resolved_math_engine(self) -> MathEngineName:
        if self.math_engine is None:
            return "local_hybrid"
        if self.math_engine not in ALLOWED_MATH_ENGINES:
            raise ConversionError("Chế độ nhận diện công thức không hợp lệ.")
        return self.math_engine  # type: ignore[return-value]

    def validate(self) -> None:
        if self.output_format not in {"docx", "xlsx"}:
            raise ConversionError("Định dạng xuất không hợp lệ (chỉ hỗ trợ docx hoặc xlsx).")
        if self.output_format == "xlsx" and self.math_ocr:
            raise ConversionError("Chuyển đổi sang Excel (.xlsx) không hỗ trợ nhận diện công thức toán.")
        if self.ocr_mode not in {"auto", "force", "text"}:
            raise ConversionError("Chế độ OCR không hợp lệ.")
        if self.dpi not in {200, 300, 400}:
            raise ConversionError("DPI phải là 200, 300 hoặc 400.")
        allowed_languages = {"vie", "eng", "vie+eng", "eng+vie"}
        if self.language not in allowed_languages:
            raise ConversionError("Ngôn ngữ OCR không hợp lệ.")
        if not 4 <= self.max_excel_columns_per_table <= 20:
            raise ConversionError("Số cột tối đa mỗi bảng Word không hợp lệ.")
        if self.math_engine is not None and self.math_engine not in ALLOWED_MATH_ENGINES:
            raise ConversionError("Chế độ nhận diện công thức không hợp lệ.")

    def report_progress(self, percent: int | float, message: str) -> None:
        """Report converter-local progress without coupling backends to the web app."""

        if self.progress_callback is None:
            return
        value = max(0, min(100, round(float(percent))))
        try:
            self.progress_callback(value, message)
        except Exception:
            # Progress reporting must never make an otherwise valid conversion fail.
            return


@dataclass(slots=True)
class ConversionResult:
    output_path: Path
    warnings: list[str] = field(default_factory=list)
    page_count: int | None = None
    sheet_count: int | None = None
    table_count: int | None = None
    equation_count: int | None = None
    math_fallback_pages: int = 0
    math_fallback_regions: int = 0
    math_status: Literal["complete", "partial", "unavailable", "no_math"] | None = None
    math_engine: str | None = None
    math_recognized_by_rules: int = 0
    math_recognized_by_local: int = 0
    math_unresolved_regions: int = 0
    fragment_text_count: int = 0

    def math_report(self) -> dict[str, object] | None:
        if self.math_status is None:
            return None
        payload: dict[str, object] = {
            "status": self.math_status,
            "equations": self.equation_count or 0,
            "fallback_pages": self.math_fallback_pages,
            "fallback_regions": self.math_fallback_regions,
        }
        if self.fragment_text_count > 0:
            payload["fragment_text_count"] = self.fragment_text_count
        if (
            self.math_engine is not None
            or self.math_recognized_by_rules
            or self.math_recognized_by_local
            or self.math_unresolved_regions
        ):
            payload["engine"] = self.math_engine or "local_hybrid"
            payload["recognized_by_rules"] = self.math_recognized_by_rules
            payload["recognized_by_local"] = self.math_recognized_by_local
            payload["unresolved_regions"] = self.math_unresolved_regions
            payload["fragment_text_count"] = self.fragment_text_count
        return payload


def _set_style_font(style, name: str, size: float, color: str | None = None) -> None:
    style.font.name = name
    style.font.size = Pt(size)
    style.element.rPr.rFonts.set(qn("w:ascii"), name)
    style.element.rPr.rFonts.set(qn("w:hAnsi"), name)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), name)
    if color:
        style.font.color.rgb = RGBColor.from_string(color)


def create_document(
    source_name: str,
    *,
    landscape: bool = False,
    body_font: str = "Calibri",
    page_width_inches: float | None = None,
    page_height_inches: float | None = None,
    margin_inches: float = 1.0,
) -> DocumentObject:
    """Create a DOCX using the compact_reference_guide design preset."""

    doc = Document()
    section = doc.sections[0]
    if page_width_inches and page_height_inches:
        section.orientation = (
            WD_ORIENT.LANDSCAPE
            if page_width_inches > page_height_inches
            else WD_ORIENT.PORTRAIT
        )
        section.page_width = Inches(page_width_inches)
        section.page_height = Inches(page_height_inches)
    elif landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11)
        section.page_height = Inches(8.5)
    else:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Inches(8.5)
        section.page_height = Inches(11)

    safe_margin = max(0.20, min(float(margin_inches), 1.5))
    section.top_margin = Inches(safe_margin)
    section.right_margin = Inches(safe_margin)
    section.bottom_margin = Inches(safe_margin)
    section.left_margin = Inches(safe_margin)
    section.header_distance = min(Inches(safe_margin * 0.8), Inches(0.492))
    section.footer_distance = min(Inches(safe_margin * 0.8), Inches(0.492))

    normal = doc.styles["Normal"]
    _set_style_font(normal, body_font, 11)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.25

    heading_tokens = {
        "Heading 1": (16, "2E74B5", 18, 10),
        "Heading 2": (13, "2E74B5", 14, 7),
        "Heading 3": (12, "1F4D78", 10, 5),
    }
    for style_name, (size, color, before, after) in heading_tokens.items():
        style = doc.styles[style_name]
        _set_style_font(style, body_font, size, color)
        style.font.bold = True
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    if "OCR Caption" not in doc.styles:
        caption = doc.styles.add_style("OCR Caption", WD_STYLE_TYPE.PARAGRAPH)
    else:
        caption = doc.styles["OCR Caption"]
    _set_style_font(caption, body_font, 9, "5B6573")
    caption.font.italic = True
    caption.paragraph_format.space_before = Pt(4)
    caption.paragraph_format.space_after = Pt(4)

    props = doc.core_properties
    props.title = Path(source_name).stem
    props.subject = "Tài liệu có thể chỉnh sửa được tạo bởi OCR2Word"
    props.author = "OCR2Word"
    props.keywords = "OCR, PDF, Excel, Word, editable"
    return doc


def add_compact_page_break(document: DocumentObject) -> None:
    """Insert a hard page boundary without creating a nearly empty page.

    ``python-docx`` uses the document's normal paragraph style for
    ``add_page_break()``.  When the preceding OCR content finishes close to
    the bottom margin, Word can move that entire paragraph to the following
    page and then apply its break, leaving a blank page in between.  A
    one-point, non-keeping separator still provides a deterministic source
    page boundary while fitting in the smallest remaining line area.
    """

    # If the last paragraph in the document is already empty (no text, drawing, or math),
    # attach the page break directly to it instead of creating a new paragraph.
    if document.paragraphs:
        last_p = document.paragraphs[-1]
        has_text = bool(last_p.text.strip())
        has_drawings = bool(last_p._element.xpath(".//a:blip") or last_p._element.xpath(".//w:drawing"))
        has_math = bool(last_p._element.xpath(".//m:oMath"))
        if not (has_text or has_drawings or has_math):
            paragraph_format = last_p.paragraph_format
            paragraph_format.space_before = Pt(0)
            paragraph_format.space_after = Pt(0)
            paragraph_format.line_spacing = Pt(1)
            paragraph_format.keep_together = False
            paragraph_format.keep_with_next = False
            paragraph_format.widow_control = False
            run = last_p.add_run()
            run.font.size = Pt(1)
            run.add_break(WD_BREAK.PAGE)
            return

    paragraph = document.add_paragraph()
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(0)
    paragraph_format.line_spacing = Pt(1)
    paragraph_format.keep_together = False
    paragraph_format.keep_with_next = False
    paragraph_format.widow_control = False
    run = paragraph.add_run()
    run.font.size = Pt(1)
    run.add_break(WD_BREAK.PAGE)


def set_run_font(run, name: str = "Calibri", size: float | None = None) -> None:
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), name)
    if size is not None:
        run.font.size = Pt(max(6, min(float(size), 42)))


def set_cell_shading(cell: _Cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = tc_pr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        tc_pr.append(shading)
    shading.set(qn("w:fill"), fill[-6:].upper())
    shading.set(qn("w:val"), "clear")


def set_cell_margins(cell: _Cell, **kwargs: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin in ("top", "start", "bottom", "end"):
        value = kwargs.get(margin)
        if value is None:
            continue
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(int(value)))
        node.set(qn("w:type"), "dxa")


def _set_cell_width(cell: _Cell, width_dxa: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(width_dxa)))
    tc_w.set(qn("w:type"), "dxa")


def configure_table_geometry(
    table: Table,
    widths_dxa: Sequence[int],
    *,
    indent_dxa: int = TABLE_INDENT_DXA,
) -> None:
    """Apply deterministic fixed-width Word table geometry."""

    widths = [max(240, int(width)) for width in widths_dxa]
    total_width = sum(widths)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False

    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(total_width))
    tbl_w.set(qn("w:type"), "dxa")

    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), str(indent_dxa))
    tbl_ind.set(qn("w:type"), "dxa")

    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        grid_col = OxmlElement("w:gridCol")
        grid_col.set(qn("w:w"), str(width))
        grid.append(grid_col)

    for row in table.rows:
        for index, cell in enumerate(row.cells[: len(widths)]):
            _set_cell_width(cell, widths[index])
            cell.width = Inches(widths[index] / 1440)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell, **TABLE_CELL_MARGIN_DXA)


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = tr_pr.find(qn("w:tblHeader"))
    if tbl_header is None:
        tbl_header = OxmlElement("w:tblHeader")
        tr_pr.append(tbl_header)
    tbl_header.set(qn("w:val"), "true")


def make_word_table(
    doc: DocumentObject,
    rows: int,
    cols: int,
    widths_dxa: Sequence[int],
    *,
    header_rows: int = 1,
) -> Table:
    table = doc.add_table(rows=max(1, rows), cols=max(1, cols))
    table.style = "Table Grid"
    configure_table_geometry(table, widths_dxa)
    for header_index in range(min(header_rows, len(table.rows))):
        set_repeat_table_header(table.rows[header_index])
        for cell in table.rows[header_index].cells:
            set_cell_shading(cell, "E8EEF5")
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
    return table


def widths_from_text(
    rows: Sequence[Sequence[object]],
    content_width_dxa: int,
    *,
    min_weight: float = 4,
    max_weight: float = 36,
) -> list[int]:
    if not rows:
        return [content_width_dxa]
    col_count = max((len(row) for row in rows), default=1)
    weights: list[float] = []
    for col in range(col_count):
        samples = [str(row[col] or "") for row in rows[:80] if col < len(row)]
        longest = max((max((len(line) for line in value.splitlines()), default=0) for value in samples), default=0)
        weights.append(max(min_weight, min(max_weight, longest + 2)))
    weight_sum = sum(weights) or 1
    widths = [max(360, round(content_width_dxa * weight / weight_sum)) for weight in weights]
    delta = content_width_dxa - sum(widths)
    widths[-1] += delta
    return widths


def clean_cell_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\x00", "").strip()
