from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml.ns import nsmap

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


def merge_docx_parts(parts: list[Path], output_path: Path) -> None:
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
        composer = Composer(master)
        for part_path in parts[1:]:
            # Each OCR chunk ends at a real source page. Insert a deterministic
            # boundary so content from the next chunk never flows onto it.
            add_compact_page_break(master)
            composer.append(Document(part_path))
        output_path.parent.mkdir(parents=True, exist_ok=True)
        composer.save(output_path)
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"Không thể ghép các phần Word: {exc}") from exc
