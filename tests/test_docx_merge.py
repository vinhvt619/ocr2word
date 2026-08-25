from __future__ import annotations

from pathlib import Path
import zipfile

from docx import Document
from lxml import etree

from app.services.docx_merge import merge_docx_parts


def test_docxcompose_import_keeps_core_property_namespace_valid(tmp_path: Path) -> None:
    output = tmp_path / "core-properties.docx"
    document = Document()
    document.core_properties.title = "OCR chunk"
    document.core_properties.keywords = "OCR, PDF, Word"
    document.save(output)

    with zipfile.ZipFile(output) as archive:
        core = etree.fromstring(archive.read("docProps/core.xml"))

    core_namespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    custom_namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    assert len(core.findall(f"{{{core_namespace}}}keywords")) == 1
    assert not core.findall(f"{{{custom_namespace}}}keywords")


def test_merge_docx_parts_preserves_order_formatting_and_tables(tmp_path: Path) -> None:
    first = tmp_path / "part-1.docx"
    second = tmp_path / "part-2.docx"
    output = tmp_path / "merged.docx"

    first_document = Document()
    first_title = first_document.add_paragraph().add_run("PHẦN MỘT")
    first_title.bold = True
    first_table = first_document.add_table(rows=1, cols=2)
    first_table.cell(0, 0).text = "Mã"
    first_table.cell(0, 1).text = "001"
    first_document.save(first)

    second_document = Document()
    second_title = second_document.add_paragraph().add_run("PHẦN HAI")
    second_title.italic = True
    second_table = second_document.add_table(rows=1, cols=2)
    second_table.cell(0, 0).text = "Nội dung"
    second_table.cell(0, 1).text = "Có thể sửa"
    second_document.save(second)

    merge_docx_parts([first, second], output)

    merged = Document(output)
    text = "\n".join(paragraph.text for paragraph in merged.paragraphs)
    assert text.index("PHẦN MỘT") < text.index("PHẦN HAI")
    assert next(paragraph for paragraph in merged.paragraphs if paragraph.text == "PHẦN MỘT").runs[0].bold
    assert next(paragraph for paragraph in merged.paragraphs if paragraph.text == "PHẦN HAI").runs[0].italic
    assert len(merged.tables) == 2
    assert merged.tables[0].cell(0, 1).text == "001"
    assert merged.tables[1].cell(0, 1).text == "Có thể sửa"
    assert 'w:type="page"' in merged._element.body.xml
    separator = next(
        paragraph
        for paragraph in merged.paragraphs
        if 'w:type="page"' in paragraph._p.xml
    )
    assert separator.paragraph_format.line_spacing.pt == 1.0
    assert separator.paragraph_format.keep_with_next is False
