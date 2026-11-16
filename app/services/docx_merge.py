from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import nsmap, qn
from docx.shared import Pt

from app.converters import ConversionError
from app.converters.base import add_compact_page_break

CORE_PROPERTIES_NAMESPACE = (
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
)

try:
    from docxcompose.composer import Composer
except ImportError:  # pragma: no cover - exercised by deployment diagnostics
    Composer = None
finally:
    # docxcompose 1.4 mutates python-docx's process-wide ``cp`` namespace to
    # the *custom* properties URI during import.  Later assignments such as
    # ``document.core_properties.keywords = ...`` then serialize an invalid
    # second ``cp:keywords`` element and Microsoft Word reports the DOCX as
    # corrupted.  Composer uses its own namespace map internally, so restoring
    # python-docx's core-properties mapping here is both safe and necessary.
    nsmap["cp"] = CORE_PROPERTIES_NAMESPACE


def _apply_page_break_to_part(part_doc: Document) -> bool:
    """Apply page_break_before to the first content element of part_doc.

    Returns True if successfully attached, or False if a standalone break is needed.
    """
    body = part_doc.element.body
    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            pPr = child.get_or_add_pPr()
            if pPr.find(qn("w:pageBreakBefore")) is None:
                pPr.append(OxmlElement("w:pageBreakBefore"))
            return True
        elif tag == "tbl":
            tc_p = child.xpath(".//w:tc[1]//w:p[1]")
            if tc_p:
                pPr = tc_p[0].get_or_add_pPr()
                if pPr.find(qn("w:pageBreakBefore")) is None:
                    pPr.append(OxmlElement("w:pageBreakBefore"))
                return True
    return False


def _normalize_section_margins(doc: Document) -> None:
    for sec in doc.sections:
        sec.top_margin = min(sec.top_margin, Pt(18))
        sec.bottom_margin = min(sec.bottom_margin, Pt(14))
        sec.left_margin = min(sec.left_margin, Pt(24))
        sec.right_margin = min(sec.right_margin, Pt(24))
        sec.header_distance = min(sec.header_distance, Pt(10))
        sec.footer_distance = min(sec.footer_distance, Pt(10))


def merge_docx_parts(parts: list[Path], output_path: Path, *, page_breaks: bool = True) -> None:
    """Merge ordered OCR chunks while preserving styles, tables and relationships."""

    if not parts:
        raise ConversionError("Không có phần Word nào để ghép.")
    missing = [str(path) for path in parts if not path.is_file()]
    if missing:
        raise ConversionError(f"Thiếu phần Word trung gian: {', '.join(missing)}")
    if Composer is None:
        raise ConversionError("Container thiếu docxcompose để ghép các phần Word.")
    try:
        master = Document(parts[0])
        _normalize_section_margins(master)
        composer = Composer(master)
        for part_path in parts[1:]:
            part_doc = Document(part_path)
            # Remove any leading empty paragraphs in the incoming part document
            while part_doc.paragraphs:
                first_p = part_doc.paragraphs[0]
                has_text = bool(first_p.text.strip())
                has_drawings = bool(first_p._element.xpath(".//a:blip") or first_p._element.xpath(".//w:drawing"))
                has_math = bool(first_p._element.xpath(".//m:oMath"))
                if not (has_text or has_drawings or has_math):
                    p_elem = first_p._element
                    p_elem.getparent().remove(p_elem)
                else:
                    break

            # Each OCR chunk ends at a real source page.
            # Using page_break_before on the first incoming element avoids creating
            # an extra empty paragraph at the bottom of the previous page, which Word
            # would otherwise push onto a new stranded page when the page is full.
            if page_breaks:
                add_compact_page_break(master)
            composer.append(part_doc)
        _normalize_section_margins(master)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        composer.save(output_path)
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"Không thể ghép các phần Word: {exc}") from exc
