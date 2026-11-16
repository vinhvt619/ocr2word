"""Shared formula page structures and editable Word equation rendering."""
from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Literal

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt
from PIL import Image, ImageOps
from pydantic import BaseModel, ConfigDict, Field, model_validator

from . import office_math
from .base import (
    ConversionError,
    create_document,
    make_word_table,
    set_run_font,
)

MAX_LATEX_LENGTH = office_math.MAX_LATEX_LENGTH
MAX_PAGE_EDGE = 3200


class MathRecognitionError(ConversionError):
    pass


def _is_valid_bbox(bbox: list[float] | None) -> bool:
    if not bbox or len(bbox) != 4:
        return False
    left, top, right, bottom = bbox
    return bool(0 <= left < right <= 1000 and 0 <= top < bottom <= 1000)


class _Span(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["text", "math"]
    content: str = Field(max_length=20_000)
    bold: bool = False
    italic: bool = False
    bbox: list[float] | None = Field(default=None, min_length=4, max_length=4)

    @model_validator(mode="after")
    def check_bbox(self):
        if self.bbox is not None and not _is_valid_bbox(self.bbox):
            raise ValueError("Invalid span region")
        return self


class _Cell(BaseModel):
    model_config = ConfigDict(extra="forbid")
    spans: list[_Span] = Field(max_length=200)


class _Row(BaseModel):
    model_config = ConfigDict(extra="forbid")
    cells: list[_Cell] = Field(min_length=1, max_length=30)


class _Block(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["paragraph", "equation", "table", "figure"]
    spans: list[_Span] = Field(max_length=500)
    rows: list[_Row] = Field(max_length=200)
    # [left, top, right, bottom], normalized to 0..1000 of the rendered page.
    bbox: list[float] = Field(min_length=4, max_length=4)
    alignment: Literal["left", "center", "right", "justify"]
    font_size: float = Field(ge=6, le=36)

    @model_validator(mode="after")
    def check_contents(self):
        if not _is_valid_bbox(self.bbox):
            raise ValueError("Invalid page region")
        if self.kind == "table" and not self.rows:
            raise ValueError("Empty table")
        if self.kind in {"paragraph", "equation"} and not any(s.content.strip() for s in self.spans):
            raise ValueError("Empty text block")
        if self.kind == "equation" and not any(s.kind == "math" for s in self.spans):
            raise ValueError("Equation is missing math")
        return self


class _Page(BaseModel):
    model_config = ConfigDict(extra="forbid")
    has_math: bool
    blocks: list[_Block] = Field(max_length=500)
    uncertain: bool

    @model_validator(mode="after")
    def check_math(self):
        if self.has_math:
            spans = [s for b in self.blocks for s in b.spans]
            spans.extend(s for b in self.blocks for r in b.rows for c in r.cells for s in c.spans)
            if not any(s.kind == "math" and s.content.strip() for s in spans):
                raise ValueError("Math page is missing equations")
        elif self.blocks:
            raise ValueError("Non-math pages must not be reconstructed")
        return self


_BALANCED_MATH_RE = re.compile(
    r"(\\\[.+?\\\]|\\\(.+?\\\)|\$\$[^$]+?\$\$|(?<!\\)\$[^$\n]+?(?<!\\)\$)",
    re.DOTALL,
)
_LIKELY_LATEX_RE = re.compile(r"(\\[a-zA-Z]+|[_^{}])")
_CURRENCY_PROSE_RE = re.compile(r"^\s*\d+(?:[.,]\d+)?\s+[a-zA-ZÀ-ỹ]")


def latex_to_omml(latex: str):
    """Convert a bounded LaTeX expression to an actual m:oMath OOXML node."""
    source = office_math.strip_outer_math_delimiters(latex)
    if not source or len(source) > MAX_LATEX_LENGTH or source == "__UNRESOLVED_MATH__":
        raise MathRecognitionError("Công thức trống hoặc quá dài.")
    if office_math._FORBIDDEN_LATEX_COMMANDS.search(source):
        raise MathRecognitionError("Công thức chứa lệnh LaTeX không được hỗ trợ.")
    try:
        return office_math.latex_to_omml_element(source)
    except Exception as exc:
        raise MathRecognitionError("Chưa thể dựng công thức này thành Equation Word.") from exc


def prepare_math_image(image: Image.Image) -> Image.Image:
    """Normalize orientation, alpha channel, color mode, and max edge before formula recognition."""
    normalized = ImageOps.exif_transpose(image)
    if normalized.mode in ("RGBA", "LA") or (normalized.mode == "P" and "transparency" in normalized.info):
        rgba = normalized.convert("RGBA")
        background = Image.new("RGB", rgba.size, "white")
        background.paste(rgba, mask=rgba.getchannel("A"))
        normalized = background
    elif normalized.mode != "RGB":
        normalized = normalized.convert("RGB")

    max_edge = max(normalized.width, normalized.height)
    if max_edge > MAX_PAGE_EDGE:
        scale = MAX_PAGE_EDGE / max_edge
        new_size = (max(1, round(normalized.width * scale)), max(1, round(normalized.height * scale)))
        normalized = normalized.resize(new_size, Image.Resampling.LANCZOS)
    return normalized


def _crop(image: Image.Image, bbox: list[float]) -> Image.Image:
    left, top, right, bottom = bbox
    return image.crop((
        int(left * image.width / 1000),
        int(top * image.height / 1000),
        max(1, round(right * image.width / 1000)),
        max(1, round(bottom * image.height / 1000)),
    ))


def _add_picture(paragraph, image: Image.Image, width: float) -> None:
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    buffer.seek(0)
    paragraph.add_run().add_picture(buffer, width=Inches(max(0.1, width)))


def _rescue_text_span_delimiters(span: _Span) -> list[_Span]:
    """Split a text span only when it contains explicit, valid LaTeX math delimiters."""
    if span.kind != "text" or not span.content:
        return [span]
    text = span.content
    if "\\" not in text and "$" not in text:
        return [span]
    # If single $ is present, require an even count of unescaped $ before attempting rescue.
    unescaped_dollars = len(re.findall(r"(?<!\\)\$", text))
    if unescaped_dollars % 2 != 0 and "\\(" not in text and "\\[" not in text:
        return [span]

    parts: list[_Span] = []
    last_end = 0
    for match in _BALANCED_MATH_RE.finditer(text):
        token = match.group(0)
        if token.startswith("$") and not token.startswith("$$"):
            inner = token[1:-1]
            if not inner or inner != inner.strip():
                continue
            if not _LIKELY_LATEX_RE.search(inner) or _CURRENCY_PROSE_RE.match(inner):
                continue
        else:
            inner = office_math.strip_outer_math_delimiters(token)
            if not inner:
                continue
        try:
            latex_to_omml(inner)
        except MathRecognitionError:
            continue
        if match.start() > last_end:
            parts.append(_Span(
                kind="text",
                content=text[last_end:match.start()],
                bold=span.bold,
                italic=span.italic,
            ))
        parts.append(_Span(kind="math", content=inner))
        last_end = match.end()

    if not parts:
        return [span]
    if last_end < len(text):
        parts.append(_Span(
            kind="text",
            content=text[last_end:],
            bold=span.bold,
            italic=span.italic,
        ))
    return parts


def _expand_spans(spans: list[_Span]) -> list[_Span]:
    expanded: list[_Span] = []
    for span in spans:
        expanded.extend(_rescue_text_span_delimiters(span))
    return expanded


def _write_spans(paragraph, spans: list[_Span], font_size: float, *, display: bool = False, alignment: str = "center") -> int:
    count = 0
    expanded = _expand_spans(spans)
    if display:
        omath_nodes = []
        for span in expanded:
            if span.kind == "math":
                omath_nodes.append(latex_to_omml(span.content))
                count += 1
            elif span.content.strip():
                text = span.content
                if not text.startswith((" ", "\u2003")):
                    text = f"   {text}"
                run = office_math.make_math_run(text, style="p", normal_text=True)
                if omath_nodes:
                    omath_nodes[-1].append(run)
                else:
                    omath = OxmlElement("m:oMath")
                    omath.append(run)
                    omath_nodes.append(omath)
        if omath_nodes:
            paragraph._p.append(office_math.wrap_display_omath(omath_nodes, alignment=alignment))
        return count

    for span in expanded:
        if span.kind == "math":
            paragraph._p.append(latex_to_omml(span.content))
            count += 1
        else:
            run = paragraph.add_run(span.content)
            set_run_font(run, "Times New Roman", font_size)
            run.bold, run.italic = span.bold, span.italic
    return count


def _new_document(source_name: str, width: float, height: float):
    doc = create_document(
        source_name,
        body_font="Times New Roman",
        page_width_inches=width,
        page_height_inches=height,
        margin_inches=0.30,
    )
    sec = doc.sections[0]
    sec.top_margin = Pt(18)
    sec.bottom_margin = Pt(14)
    sec.left_margin = Pt(24)
    sec.right_margin = Pt(24)
    sec.header_distance = Pt(10)
    sec.footer_distance = Pt(10)
    return doc


def _prepare_inline_items(spans: list[_Span]) -> tuple[list[tuple], int, int, bool]:
    """Prepare inline spans before writing to a paragraph or table cell.

    Returns (prepared_items, equation_count, inline_fallback_count, block_must_fallback).
    """
    prepared: list[tuple] = []
    equations = 0
    inline_fallbacks = 0
    for span in _expand_spans(spans):
        if span.kind == "math":
            try:
                omath = latex_to_omml(span.content)
                prepared.append(("math", omath))
                equations += 1
            except MathRecognitionError:
                if _is_valid_bbox(span.bbox):
                    prepared.append(("crop", span.bbox))
                    inline_fallbacks += 1
                else:
                    return [], 0, 0, True
        else:
            prepared.append(("text", span))
    return prepared, equations, inline_fallbacks, False


def _commit_inline_items(paragraph, prepared: list[tuple], image: Image.Image, usable_width: float, font_size: float) -> None:
    for kind, payload in prepared:
        if kind == "math":
            paragraph._p.append(payload)
        elif kind == "crop":
            bbox = payload
            crop = _crop(image, bbox)
            pic_width = min(usable_width, max(0.25, usable_width * (bbox[2] - bbox[0]) / 1000))
            _add_picture(paragraph, crop, pic_width)
        else:
            span: _Span = payload
            run = paragraph.add_run(span.content)
            set_run_font(run, "Times New Roman", font_size)
            run.bold, run.italic = span.bold, span.italic


def _write_page_detailed(
    page: _Page,
    image: Image.Image,
    source_name: str,
    width: float,
    height: float,
    *,
    page_label: str = "Trang 1",
) -> tuple:
    document = _new_document(source_name, width, height)
    usable = max(1.0, width - (48 / 72))
    equation_count = table_count = fallback_regions = 0
    page_warnings: list[str] = []
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }

    fig_count = sum(1 for b in page.blocks if b.kind == "figure")
    text_blocks = sum(1 for b in page.blocks if b.kind in {"paragraph", "equation"})
    effective_blocks = sum(len(b.rows) if b.kind == "table" else (2 if b.kind == "figure" else 1) for b in page.blocks)

    if effective_blocks > 22 or len(page.blocks) > 20:
        eq_space_after = Pt(0.1)
        prose_space_after = Pt(0.2)
        fig_space_after = Pt(1.0)
        p_line_spacing = 0.96
        fig_max_ratio = 0.10 if text_blocks > 25 else (0.14 if text_blocks > 15 else 0.20)
        fig_max_h = 1.0 if text_blocks > 25 else (1.3 if text_blocks > 15 else (1.8 if fig_count >= 2 else 2.0))
    elif effective_blocks > 14 or len(page.blocks) > 13:
        eq_space_after = Pt(0.4)
        prose_space_after = Pt(0.6)
        fig_space_after = Pt(1.5)
        p_line_spacing = 0.98
        fig_max_ratio = 0.18 if text_blocks > 15 else 0.24
        fig_max_h = 1.5 if text_blocks > 15 else (2.0 if fig_count >= 2 else 2.4)
    elif effective_blocks > 8:
        eq_space_after = Pt(0.8)
        prose_space_after = Pt(1.0)
        fig_space_after = Pt(2.0)
        p_line_spacing = 1.0
        fig_max_ratio = 0.28
        fig_max_h = 2.4 if fig_count >= 2 else 2.8
    else:
        eq_space_after = Pt(1.2)
        prose_space_after = Pt(1.5)
        fig_space_after = Pt(2.2)
        p_line_spacing = 1.05
        fig_max_ratio = 0.35
        fig_max_h = 3.5

    for block_index, block in enumerate(page.blocks, 1):
        if block.kind == "table":
            columns = max(len(row.cells) for row in block.rows)
            prepared_grid: list[list[list[tuple]]] = []
            table_eqs = 0
            table_fallbacks = 0
            table_aborted = False
            for row in block.rows:
                prepared_row: list[list[tuple]] = []
                for cell in row.cells:
                    items, cell_eqs, cell_fb, must_abort = _prepare_inline_items(cell.spans)
                    if must_abort:
                        table_aborted = True
                        break
                    prepared_row.append(items)
                    table_eqs += cell_eqs
                    table_fallbacks += cell_fb
                if table_aborted:
                    break
                prepared_grid.append(prepared_row)

            if table_aborted:
                paragraph = document.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(2.5)
                paragraph.paragraph_format.line_spacing = 1.05
                crop = _crop(image, block.bbox)
                _add_picture(paragraph, crop, min(usable, usable * (block.bbox[2] - block.bbox[0]) / 1000))
                fallback_regions += 1
                page_warnings.append(
                    f"{page_label} (vùng {block_index}): Chưa dựng được công thức trong bảng; "
                    "đã giữ ảnh vùng bảng (chưa chỉnh sửa được), các phần khác trên trang vẫn chỉnh sửa được."
                )
                continue

            table = make_word_table(
                document,
                len(block.rows),
                columns,
                [round(usable * 1440 / columns)] * columns,
                header_rows=0,
            )
            cell_usable = usable / max(1, columns)
            table_font_size = min(block.font_size, 9.0) if effective_blocks > 18 else (min(block.font_size, 9.5) if effective_blocks > 12 else block.font_size)
            for row_index, prepared_row in enumerate(prepared_grid):
                anchor_col_index = 0
                for col_index, items in enumerate(prepared_row):
                    if (
                        col_index > 0
                        and len(items) == 1
                        and items[0][0] == "text"
                        and getattr(items[0][1], "content", None) == "__DOCX_CELL_MERGED_LEFT__"
                    ):
                        table.cell(row_index, anchor_col_index).merge(table.cell(row_index, col_index))
                        continue
                    anchor_col_index = col_index
                    cell_p = table.cell(row_index, col_index).paragraphs[0]
                    cell_p.paragraph_format.space_before = Pt(0)
                    cell_p.paragraph_format.space_after = Pt(0)
                    cell_p.paragraph_format.line_spacing = 1.0
                    _commit_inline_items(
                        cell_p,
                        items,
                        image,
                        cell_usable,
                        table_font_size,
                    )
            table_count += 1
            equation_count += table_eqs
            if table_fallbacks:
                fallback_regions += table_fallbacks
                page_warnings.append(
                    f"{page_label} (vùng {block_index}): Đã giữ ảnh {table_fallbacks} công thức trong bảng "
                    "chưa chỉnh sửa được; các ô bảng khác vẫn chỉnh sửa được."
                )
            continue

        paragraph = document.add_paragraph()
        paragraph.alignment = align_map[block.alignment]
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.line_spacing = p_line_spacing
        if block.kind == "equation":
            paragraph.paragraph_format.space_after = eq_space_after
        elif block.kind == "figure":
            paragraph.paragraph_format.space_after = fig_space_after
        else:
            paragraph.paragraph_format.space_after = prose_space_after

        if block.kind == "figure":
            crop = _crop(image, block.bbox)
            crop_w_norm = block.bbox[2] - block.bbox[0]
            crop_h_norm = block.bbox[3] - block.bbox[1]
            target_w = min(usable, max(0.5, usable * crop_w_norm / 1000))
            if crop_w_norm > 0 and crop_h_norm > 0:
                calc_h = target_w * (crop_h_norm / crop_w_norm)
                max_allowed_h = min(fig_max_h, (height - 1.2) * fig_max_ratio)
                if calc_h > max_allowed_h:
                    target_w = target_w * (max_allowed_h / calc_h)
            _add_picture(paragraph, crop, target_w)
            continue

        if effective_blocks > 18:
            effective_font_size = min(block.font_size, 10.5)
        elif effective_blocks > 12:
            effective_font_size = min(block.font_size, 11.5)
        else:
            effective_font_size = block.font_size

        if block.kind == "equation":
            try:
                equation_count += _write_spans(
                    paragraph,
                    block.spans,
                    effective_font_size,
                    display=True,
                    alignment=block.alignment,
                )
            except MathRecognitionError:
                crop = _crop(image, block.bbox)
                _add_picture(paragraph, crop, min(usable, usable * (block.bbox[2] - block.bbox[0]) / 1000))
                fallback_regions += 1
                page_warnings.append(
                    f"{page_label} (vùng {block_index}): Chưa dựng được công thức riêng dòng; "
                    "đã giữ ảnh vùng công thức (chưa chỉnh sửa được), các phần khác trên trang vẫn chỉnh sửa được."
                )
            continue

        # block.kind == "paragraph"
        items, para_eqs, para_fallbacks, must_abort = _prepare_inline_items(block.spans)
        if must_abort:
            crop = _crop(image, block.bbox)
            _add_picture(paragraph, crop, min(usable, usable * (block.bbox[2] - block.bbox[0]) / 1000))
            fallback_regions += 1
            page_warnings.append(
                f"{page_label} (vùng {block_index}): Chưa dựng được công thức trong đoạn; "
                "đã giữ ảnh vùng đoạn văn (chưa chỉnh sửa được), các phần khác trên trang vẫn chỉnh sửa được."
            )
            continue

        _commit_inline_items(paragraph, items, image, usable, effective_font_size)
        equation_count += para_eqs
        if para_fallbacks:
            fallback_regions += para_fallbacks
            page_warnings.append(
                f"{page_label} (vùng {block_index}): Đã giữ ảnh {para_fallbacks} công thức inline "
                "chưa chỉnh sửa được; văn bản xung quanh vẫn chỉnh sửa được."
            )

    # Strip any trailing empty paragraphs
    while document.paragraphs:
        last_p = document.paragraphs[-1]
        has_text = bool(last_p.text.strip())
        has_drawings = bool(last_p._element.xpath(".//a:blip") or last_p._element.xpath(".//w:drawing"))
        has_math = bool(last_p._element.xpath(".//m:oMath"))
        if not (has_text or has_drawings or has_math):
            p_elem = last_p._element
            p_elem.getparent().remove(p_elem)
        else:
            break

    return document, equation_count, table_count, fallback_regions, page_warnings


def _source_image_page(image: Image.Image, source_name: str, width: float, height: float, output: Path):
    document = _new_document(source_name, width, height)
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(0)
    usable_width = min(width - 1.2, (height - 1.3) * image.width / max(1, image.height))
    _add_picture(paragraph, image, usable_width)
    document.save(output)


def _math_status(equations: int, fallback_pages: int, fallback_regions: int, fragment_count: int = 0) -> str:
    if fallback_pages or fallback_regions or fragment_count:
        return "partial" if equations else "unavailable"
    return "complete" if equations else "no_math"
