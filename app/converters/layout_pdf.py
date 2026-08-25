from __future__ import annotations

import statistics
import threading
import unicodedata
from copy import deepcopy
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.text.paragraph import Paragraph

from .base import ConversionError, MissingDependencyError

try:
    from pdf2docx import Converter as PdfLayoutConverter
    from pdf2docx.page.RawPageFactory import RawPageFactory
    from pdf2docx.page.RawPageFitz import RawPageFitz
except ImportError:  # pragma: no cover - exercised in dependency diagnostics
    PdfLayoutConverter = None
    RawPageFactory = None
    RawPageFitz = object


_CONVERSION_LOCK = threading.RLock()


def _font_alias(pdf_font: str) -> str:
    """Map common PDF/TeX font aliases to fonts normally available in Word."""

    name = (pdf_font or "").split("+")[-1]
    lowered = name.casefold()
    if "nimbusroman" in lowered or "nimbusrom" in lowered or "times" in lowered:
        return "Times New Roman"
    if "nimbussans" in lowered or "helvetica" in lowered or "arial" in lowered:
        return "Arial"
    if "nimbusmono" in lowered or "courier" in lowered or "mono" in lowered:
        return "Courier New"
    if "calibri" in lowered:
        return "Calibri"
    return name or "Times New Roman"


def _is_combining(character: str) -> bool:
    return bool(character) and unicodedata.category(character[0]).startswith("M")


def _space_character(previous: dict[str, Any], current: dict[str, Any]) -> dict[str, Any]:
    previous_bbox = tuple(previous.get("bbox", (0, 0, 0, 0)))
    current_bbox = tuple(current.get("bbox", (0, 0, 0, 0)))
    baseline = current.get("origin") or previous.get("origin") or (previous_bbox[2], previous_bbox[3])
    return {
        "c": " ",
        "origin": (previous_bbox[2], baseline[1]),
        "bbox": (
            previous_bbox[2],
            min(previous_bbox[1], current_bbox[1]),
            current_bbox[0],
            max(previous_bbox[3], current_bbox[3]),
        ),
    }


def _needs_inferred_space(
    previous: dict[str, Any] | None,
    current: dict[str, Any],
    threshold: float,
) -> bool:
    if previous is None:
        return False
    previous_text = str(previous.get("c", ""))
    current_text = str(current.get("c", ""))
    if not previous_text or not current_text or previous_text.isspace() or current_text.isspace():
        return False
    if _is_combining(current_text):
        return False
    previous_bbox = tuple(previous.get("bbox", (0, 0, 0, 0)))
    current_bbox = tuple(current.get("bbox", (0, 0, 0, 0)))
    gap = float(current_bbox[0]) - float(previous_bbox[2])
    return gap >= threshold


def _repair_line(line: dict[str, Any]) -> None:
    """Restore visually present spaces that are absent from the PDF text stream."""

    direction = tuple(line.get("dir", (1.0, 0.0)))
    if len(direction) >= 2 and (abs(float(direction[0]) - 1.0) > 0.05 or abs(float(direction[1])) > 0.05):
        return

    raw_spans = [span for span in line.get("spans", []) if span.get("chars")]
    spans: list[dict[str, Any]] = []
    pending_leading_spaces: list[dict[str, Any]] = []
    for span in raw_spans:
        characters = list(span.get("chars", []))
        span_text = "".join(str(character.get("c", "")) for character in characters)
        if span_text and span_text.isspace():
            # pdf2docx drops whitespace-only spans. Merge them into a neighboring
            # text span before its object model is created.
            if spans:
                spans[-1].setdefault("chars", []).extend(characters)
            else:
                pending_leading_spaces.extend(characters)
            continue
        if pending_leading_spaces:
            span["chars"] = pending_leading_spaces + characters
            pending_leading_spaces = []
        spans.append(span)
    if pending_leading_spaces and spans:
        spans[-1].setdefault("chars", []).extend(pending_leading_spaces)
    line["spans"] = spans
    if not spans:
        return

    explicit_space_widths: list[float] = []
    font_sizes: list[float] = []
    for span in spans:
        span["font"] = _font_alias(str(span.get("font", "")))
        font_sizes.append(float(span.get("size", 11.0) or 11.0))
        for character in span.get("chars", []):
            if str(character.get("c", "")) == " ":
                bbox = tuple(character.get("bbox", (0, 0, 0, 0)))
                width = float(bbox[2]) - float(bbox[0])
                if width > 0:
                    explicit_space_widths.append(width)

    median_font_size = statistics.median(font_sizes) if font_sizes else 11.0
    estimated_space_width = (
        min(statistics.median(explicit_space_widths), median_font_size * 0.33)
        if explicit_space_widths
        else median_font_size * 0.27
    )
    threshold = max(0.8, estimated_space_width * 0.62, median_font_size * 0.14)

    previous: dict[str, Any] | None = None
    for span in spans:
        repaired: list[dict[str, Any]] = []
        for character in list(span.get("chars", [])):
            if _needs_inferred_space(previous, character, threshold):
                repaired.append(_space_character(previous, character))
            repaired.append(character)
            previous = character
        span["chars"] = repaired


def repair_raw_text_blocks(blocks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    for block in blocks:
        if block.get("type") != 0:
            continue
        lines = block.get("lines", [])
        for line in lines:
            _repair_line(line)
        # pdf2docx may merge adjacent PDF lines into one Word paragraph. Keep a
        # word boundary at every source line transition; a trailing space is
        # harmless when the line remains a hard break.
        for line in lines:
            spans = line.get("spans", [])
            if not spans:
                continue
            characters = spans[-1].get("chars", [])
            if not characters or str(characters[-1].get("c", "")).isspace():
                continue
            last = characters[-1]
            bbox = tuple(last.get("bbox", (0, 0, 0, 0)))
            font_size = float(spans[-1].get("size", 11.0) or 11.0)
            characters.append(
                {
                    "c": " ",
                    "origin": (bbox[2], (last.get("origin") or (bbox[2], bbox[3]))[1]),
                    "bbox": (bbox[2], bbox[1], bbox[2] + max(1.0, font_size * 0.25), bbox[3]),
                }
            )
    return blocks


class SpacingAwareRawPageFitz(RawPageFitz):
    """pdf2docx page adapter with Vietnamese spacing and Word font repairs."""

    def _preprocess_text(self, **settings):
        blocks = super()._preprocess_text(**settings)
        return repair_raw_text_blocks(blocks)


def layout_converter_available() -> bool:
    return PdfLayoutConverter is not None and RawPageFactory is not None


def _repair_tabbed_image_headers(document) -> bool:
    """Split a merged, tabbed logo header into two positioned paragraphs."""

    changed = False
    for paragraph in list(document.paragraphs):
        if not paragraph._p.xpath(".//w:drawing") or paragraph._p.xpath(".//w:br"):
            continue
        seen_text = False
        separator = None
        for run in list(paragraph.runs):
            if seen_text and run._r.xpath(".//w:tab") and not run.text.replace("\t", "").strip():
                separator = run
                break
            if run.text.replace("\t", "").strip():
                seen_text = True
        if separator is None:
            continue

        tab_stops = list(paragraph.paragraph_format.tab_stops)
        second_line_indent = tab_stops[0].position if tab_stops else paragraph.paragraph_format.left_indent
        new_element = OxmlElement("w:p")
        if paragraph._p.pPr is not None:
            new_element.append(deepcopy(paragraph._p.pPr))
        paragraph._p.addnext(new_element)
        second_line = Paragraph(new_element, paragraph._parent)
        second_line.paragraph_format.tab_stops.clear_all()
        second_line.paragraph_format.left_indent = second_line_indent or Pt(0)
        second_line.paragraph_format.right_indent = Pt(0)
        second_line.paragraph_format.first_line_indent = Pt(0)
        second_line.paragraph_format.space_before = Pt(0)
        second_line.paragraph_format.space_after = Pt(0)

        siblings = list(paragraph._p)
        separator_index = siblings.index(separator._r)
        for element in siblings[separator_index + 1 :]:
            new_element.append(element)
        paragraph._p.remove(separator._r)
        changed = True
    return changed


def _normalize_layout_document(output_path: Path) -> None:
    """Prevent a page footer from spilling onto a new Word page.

    pdf2docx creates one section per PDF page. Large source bottom margins can
    leave too little room for its absolutely positioned page content, so Word
    may move the final footer paragraph to an otherwise blank next page.
    """

    document = Document(output_path)
    changed = _repair_tabbed_image_headers(document)
    for section in document.sections:
        bottom_margin = section.bottom_margin
        if bottom_margin is not None and bottom_margin.pt > 24.0:
            section.bottom_margin = Pt(24.0)
            changed = True
    if changed:
        document.save(output_path)


def convert_pdf_with_layout(
    input_path: Path,
    output_path: Path,
    *,
    detect_tables: bool = True,
    start: int = 0,
    end: int | None = None,
) -> None:
    if not layout_converter_available():
        raise MissingDependencyError(
            "Thiếu pdf2docx để giữ bố cục PDF. Hãy chạy setup.ps1 hoặc cài các gói trong requirements.txt."
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    converter = None
    with _CONVERSION_LOCK:
        previous_backend = RawPageFactory.MAP.get("PYMUPDF")
        RawPageFactory.MAP["PYMUPDF"] = SpacingAwareRawPageFitz
        try:
            converter = PdfLayoutConverter(str(input_path))
            converter.convert(
                str(output_path),
                start=start,
                end=end,
                multi_processing=False,
                parse_lattice_table=detect_tables,
                # Stream-table guessing often mistakes TOCs and justified prose for tables.
                parse_stream_table=False,
                delete_end_line_hyphen=False,
            )
            _normalize_layout_document(output_path)
        except Exception as exc:
            raise ConversionError(f"Không thể dựng Word theo bố cục PDF: {exc}") from exc
        finally:
            if converter is not None:
                converter.close()
            if previous_backend is not None:
                RawPageFactory.MAP["PYMUPDF"] = previous_backend
