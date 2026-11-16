"""Recognize equations with PDF structure rules and an optional local model."""
from __future__ import annotations

import logging
import re
import shutil
import tempfile
from dataclasses import replace
from pathlib import Path
from typing import Any, Literal

from docx import Document
from PIL import Image, UnidentifiedImageError

from .base import (
    LANDSCAPE_CONTENT_WIDTH_DXA,
    PORTRAIT_CONTENT_WIDTH_DXA,
    ConversionError,
    ConversionOptions,
    ConversionResult,
    create_document,
)
from .math_pdf import (
    MAX_PAGE_EDGE,
    MathRecognitionError,
    _Block,
    _Cell,
    _crop,
    _math_status,
    _Page,
    _Row,
    _source_image_page,
    _Span,
    _write_page_detailed,
    latex_to_omml,
    prepare_math_image,
)
from .math_regions import (
    detect_visual_fraction_regions,
    is_fragment_text,
    is_likely_formula_text,
    is_single_formula_image,
    normalize_bbox_to_1000,
    rect_overlap_ratio,
    union_bbox,
    vertical_overlap_ratio,
)
from .math_rules import extract_digital_math_page

logger = logging.getLogger(__name__)


def get_math_engines_status() -> dict[str, dict[str, object]]:
    """Return availability and metadata for all math engines without importing torch."""
    from . import math_local

    local_info = math_local.get_local_math_engine_info()
    return {
        "default_when_enabled": "local_hybrid",
        "engines": {
            "local_hybrid": {
                "available": local_info.available,
                "configured": local_info.package_installed,
                "weights_ready": local_info.weights_available,
                "uses_local_model": True,
                "model_name": local_info.model_name,
                "revision": math_local.PIX2TEX_RELEASE_TAG,
                "reason": local_info.reason,
                "label": "Cục bộ (Không dùng API)",
                "description": "Kết hợp quy tắc PDF số với mô hình nhận diện công thức chạy trên máy chủ.",
            },
            "rules": {
                "available": True,
                "configured": True,
                "weights_ready": True,
                "uses_local_model": False,
                "label": "Chỉ theo quy tắc PDF (Nhanh, không tải model)",
                "description": "Chỉ phân tích cấu trúc chữ và nét vẽ của PDF số; không hỗ trợ ảnh hoặc bản scan.",
            },
        },
    }


def _count_docx_equations(docx_path: Path) -> int:
    """Count actual top-level m:oMath elements present in the saved DOCX package."""
    try:
        doc = Document(docx_path)
        return len(doc.element.xpath(".//m:oMath[not(ancestor::m:oMath)]"))
    except Exception:
        return 0


def _engine_summary_warning(
    engine: str,
    equations: int,
    fallback_regions: int,
    source_kind: str = "PDF",
    *,
    fallback_pages: int = 0,
    recognized_by_rules: int = 0,
    recognized_by_local: int = 0,
) -> str:
    engine_note = (
        "Chế độ theo quy tắc PDF."
        if engine == "rules"
        else "Công thức được nhận diện bằng chế độ cục bộ. Hãy đối chiếu lại ký hiệu và số liệu."
    )
    if not equations and (fallback_pages or fallback_regions):
        return (
            "Chưa tạo được Equation nào. Đây là bản giữ nội dung nguồn để đối chiếu, "
            f"không phải bản công thức chỉnh sửa được ({fallback_pages} trang và "
            f"{fallback_regions} vùng chưa nhận diện thành công). {engine_note}"
        )
    detail_parts = []
    if recognized_by_rules:
        detail_parts.append(f"{recognized_by_rules} theo quy tắc PDF")
    if recognized_by_local:
        detail_parts.append(f"{recognized_by_local} bằng mô hình cục bộ")
    detail_suffix = f" ({', '.join(detail_parts)})" if detail_parts else ""
    if fallback_pages:
        return (
            f"Đã tạo {equations} Equation{detail_suffix}; còn {fallback_pages} trang và {fallback_regions} vùng "
            f"giữ dạng ảnh, chưa chỉnh sửa được. {engine_note}"
        )
    if fallback_regions > 0:
        return (
            f"Đã tạo {equations} công thức Equation có thể chỉnh sửa{detail_suffix}; "
            f"{fallback_regions} vùng công thức ngoài phạm vi đã giữ dạng ảnh (chưa chỉnh sửa được). "
            f"{engine_note}"
        )
    return (
        f"Đã tạo {equations} công thức Equation có thể chỉnh sửa{detail_suffix}. "
        f"Các trang có công thức được dựng lại từ {source_kind}. {engine_note}"
    )


def _upgrade_unresolved_spans_with_local(page_ir: _Page, page_image: Image.Image) -> tuple[int, int]:
    """Attempt to resolve unresolved math spans in a digital page IR using the local pix2tex model.

    Returns (upgraded_by_local_count, remaining_unresolved_count).
    """
    from . import math_local

    upgraded = 0
    remaining = 0

    def process_span(span: _Span, fallback_bbox: list[float]) -> None:
        nonlocal upgraded, remaining
        if span.kind != "math":
            return
        needs_local = span.content == "__UNRESOLVED_MATH__"
        if not needs_local:
            try:
                latex_to_omml(span.content)
            except MathRecognitionError:
                needs_local = True
        if not needs_local:
            return
        target_bbox = span.bbox if span.bbox is not None else fallback_bbox
        try:
            crop_img = _crop(page_image, target_bbox)
            recognized_latex = math_local.recognize_formula_crop(crop_img)
            latex_to_omml(recognized_latex)
            span.content = recognized_latex
            upgraded += 1
        except Exception as exc:
            logger.debug("Local math fallback could not resolve region: %s", type(exc).__name__)
            span.content = "__UNRESOLVED_MATH__"
            if span.bbox is None:
                span.bbox = list(fallback_bbox)
            remaining += 1

    for block in page_ir.blocks:
        for span in block.spans:
            process_span(span, block.bbox)
        for row in block.rows:
            for cell in row.cells:
                for span in cell.spans:
                    process_span(span, block.bbox)

    return upgraded, remaining


def _is_plain_prose_ocr(lines: list[dict[str, Any]], visual_fraction_boxes: list[tuple[int, int, int, int]]) -> bool:
    """Return True if OCR lines clearly represent non-mathematical prose and there are no visual fractions."""
    if visual_fraction_boxes:
        return False
    if not lines:
        return False
    full_text = " ".join(str(line.get("text", "")) for line in lines).strip()
    if not full_text:
        return False
    if is_likely_formula_text(full_text) or any(is_likely_formula_text(str(l.get("text", ""))) for l in lines):
        return False
    # Check for math operators or explicit equation patterns
    if any(sym in full_text for sym in ("=", "≠", "≤", "≥", "≈", "±", "√", "∫", "∑", "∏", "^", "_")):
        return False
    if re.search(r"\b[a-zA-Z]\s*[\+\-\*/]\s*[a-zA-Z0-9]", full_text):
        return False
    multi_letter_words = re.findall(r"[a-zA-ZÀ-ỹ]{2,}", full_text)
    return len(multi_letter_words) >= 2


def _build_local_raster_page_ir(
    frame: Image.Image,
    options: ConversionOptions,
) -> tuple[_Page, int, int, list[str]]:
    """Segment a raster image or scanned page into prose, tables, and formula crops, running local pix2tex on formulas.

    Returns (page_ir, recognized_by_local, unresolved_regions, warnings).
    """
    from app.services.ocr import image_region_to_data, image_to_data
    from . import math_local
    from .ocr_layout import (
        merge_ocr_lines,
        normalize_ocr_text,
        ocr_data_to_words,
        ocr_words_to_lines,
    )
    from .raster_tables import (
        RasterTable,
        _autorotate_low_confidence_image,
        detect_raster_tables,
    )

    warnings: list[str] = []
    ocr_available = True
    words: list[dict[str, Any]] = []
    sparse_words: list[dict[str, Any]] = []
    tables: list[RasterTable] = []

    try:
        data, _lang, ocr_warnings = image_to_data(frame, options.language)
        frame, data, rot_warnings = _autorotate_low_confidence_image(frame, data, options.language)
        warnings.extend(ocr_warnings)
        warnings.extend(rot_warnings)
        words = ocr_data_to_words(data)
        if options.detect_tables:
            try:
                tables = detect_raster_tables(
                    frame,
                    words,
                    requested_language=options.language,
                    warnings=warnings,
                )
            except Exception:
                pass
        try:
            sparse_data, _slang, sparse_warnings = image_region_to_data(frame, options.language, psm=11)
            sparse_words = ocr_data_to_words(sparse_data, confidence_threshold=40)
            warnings.extend(sparse_warnings)
        except ConversionError:
            pass
    except ConversionError:
        ocr_available = False

    table_bboxes = [t.bbox for t in tables]
    visual_fraction_boxes = detect_visual_fraction_regions(frame, excluded_bboxes=table_bboxes)

    def inside_table(word: dict[str, Any]) -> bool:
        cx = (word["left"] + word["right"]) * 0.5
        cy = (word["top"] + word["bottom"]) * 0.5
        return any(
            left <= cx <= right and top <= cy <= bottom
            for left, top, right, bottom in table_bboxes
        )

    outside_lines = (
        merge_ocr_lines(
            ocr_words_to_lines(w for w in words if not inside_table(w)),
            ocr_words_to_lines(w for w in sparse_words if not inside_table(w)),
        )
        if ocr_available
        else []
    )

    # 1. Single formula snippet fast-path (when image is a compact formula crop, not plain prose or a table)
    if not tables and is_single_formula_image(frame) and not _is_plain_prose_ocr(outside_lines, visual_fraction_boxes):
        norm_full = [0.0, 0.0, 1000.0, 1000.0]
        try:
            latex = math_local.recognize_formula_crop(frame)
            latex_to_omml(latex)
            return (
                _Page(
                    has_math=True,
                    uncertain=False,
                    blocks=[
                        _Block(
                            kind="equation",
                            spans=[_Span(kind="math", content=latex, bbox=norm_full)],
                            rows=[],
                            bbox=norm_full,
                            alignment="center",
                            font_size=12.0,
                        )
                    ],
                ),
                1,
                0,
                warnings,
            )
        except Exception as exc:
            if visual_fraction_boxes or any(is_likely_formula_text(str(l.get("text", ""))) for l in outside_lines) or not outside_lines:
                logger.warning("Local formula recognition failed on single formula image (%s)", type(exc).__name__)
                return (
                    _Page(
                        has_math=True,
                        uncertain=False,
                        blocks=[
                            _Block(
                                kind="equation",
                                spans=[_Span(kind="math", content="__UNRESOLVED_MATH__", bbox=norm_full)],
                                rows=[],
                                bbox=norm_full,
                                alignment="center",
                                font_size=12.0,
                            )
                        ],
                    ),
                    0,
                    1,
                    warnings,
                )

    # 2. Multi-block / full page segmentation
    formula_boxes: list[tuple[int, int, int, int]] = list(visual_fraction_boxes)
    absorbed_line_indices: set[int] = set()

    for l_idx, line in enumerate(outside_lines):
        lbox = (int(line["left"]), int(line["top"]), int(line["right"]), int(line["bottom"]))
        ltext = normalize_ocr_text(str(line.get("text", ""))).strip()
        if not ltext:
            absorbed_line_indices.add(l_idx)
            continue

        # Check if this OCR line overlaps or sits on the same formula band as a detected visual fraction box
        merged_with_vbox = False
        for v_idx, vbox in enumerate(formula_boxes):
            # 1. Existing area overlap
            if rect_overlap_ratio(lbox, vbox) > 0.15:
                merged_with_vbox = True
            # 2. Horizontal same-band overlap
            elif (
                vertical_overlap_ratio(lbox, vbox) > 0.35
                and max(0, lbox[0] - vbox[2], vbox[0] - lbox[2]) <= max(48, int((vbox[3] - vbox[1]) * 0.9))
                and (is_likely_formula_text(ltext) or is_fragment_text(ltext))
            ):
                merged_with_vbox = True
            # 3. Vertical stacking: Numerators and denominators directly above or below fraction bar
            else:
                h_overlap = min(lbox[2], vbox[2]) - max(lbox[0], vbox[0])
                l_w = max(1, lbox[2] - lbox[0])
                v_w = max(1, vbox[2] - vbox[0])
                v_gap = max(0, vbox[1] - lbox[3], lbox[1] - vbox[3])
                if v_gap <= 28 and h_overlap > 0 and (h_overlap / l_w >= 0.25 or h_overlap / v_w >= 0.25):
                    if is_fragment_text(ltext) or is_likely_formula_text(ltext) or len(ltext) <= 18:
                        merged_with_vbox = True
            # 4. Adjacent subscript/superscript (e.g. '2', '3', 'i', 'n')
            if not merged_with_vbox and len(ltext) <= 4:
                h_gap = max(0, lbox[0] - vbox[2], vbox[0] - lbox[2])
                v_dist = max(0, lbox[1] - vbox[3], vbox[1] - lbox[3])
                if h_gap <= 20 and v_dist <= 16:
                    merged_with_vbox = True

            if merged_with_vbox:
                ub = union_bbox([vbox, lbox])
                formula_boxes[v_idx] = (int(ub[0]), int(ub[1]), int(ub[2]), int(ub[3]))
                absorbed_line_indices.add(l_idx)
                break
        if merged_with_vbox:
            continue

        if "{" in ltext or is_likely_formula_text(ltext):
            pad_x = 8
            pad_y = 6
            formula_boxes.append(
                (
                    max(0, lbox[0] - pad_x),
                    max(0, lbox[1] - pad_y),
                    min(frame.width, lbox[2] + pad_x),
                    min(frame.height, lbox[3] + pad_y),
                )
            )
            absorbed_line_indices.add(l_idx)

    # Merge any overlapping formula boxes after line absorption
    if formula_boxes:
        formula_boxes.sort(key=lambda b: (b[1], b[0]))
        consolidated: list[tuple[int, int, int, int]] = []
        for box in formula_boxes:
            if consolidated and (
                rect_overlap_ratio(box, consolidated[-1]) > 0.15
                or (
                    vertical_overlap_ratio(box, consolidated[-1]) > 0.45
                    and max(0, box[0] - consolidated[-1][2], consolidated[-1][0] - box[2]) <= 32
                )
            ):
                ub = union_bbox([consolidated[-1], box])
                consolidated[-1] = (int(ub[0]), int(ub[1]), int(ub[2]), int(ub[3]))
            else:
                consolidated.append(box)
        formula_boxes = consolidated

    # Check if any table cells contain formulas
    has_table_math = any(
        is_likely_formula_text(cell_text)
        for tbl in tables
        for row in tbl.cells
        for cell_text in row
        if cell_text
    )

    if not formula_boxes and not has_table_math:
        cand_lines = [l for l in outside_lines if is_likely_formula_text(str(l.get("text", "")))]
        if cand_lines:
            for cl in cand_lines:
                cl_box = (int(cl["left"]), int(cl["top"]), int(cl["right"]), int(cl["bottom"]))
                formula_boxes.append(
                    (
                        max(0, cl_box[0] - 6),
                        max(0, cl_box[1] - 4),
                        min(frame.width, cl_box[2] + 6),
                        min(frame.height, cl_box[3] + 4),
                    )
                )
        else:
            return _Page(has_math=False, blocks=[], uncertain=False), 0, 0, warnings

    dated_blocks: list[tuple[int, int, _Block]] = []
    recognized_local = 0
    unresolved_count = 0

    for l_idx, line in enumerate(outside_lines):
        if l_idx in absorbed_line_indices:
            continue
        lbox = (float(line["left"]), float(line["top"]), float(line["right"]), float(line["bottom"]))
        if any(rect_overlap_ratio(lbox, fb) > 0.25 for fb in formula_boxes):
            continue
        ltext = normalize_ocr_text(str(line.get("text", ""))).strip()
        if not ltext:
            continue
        norm_bbox = normalize_bbox_to_1000(lbox, frame.width, frame.height)
        dated_blocks.append(
            (
                int(line["top"]),
                int(line["left"]),
                _Block(
                    kind="paragraph",
                    spans=[_Span(kind="text", content=ltext)],
                    rows=[],
                    bbox=norm_bbox,
                    alignment="left",
                    font_size=11.5,
                ),
            )
        )

    for fx0, fy0, fx1, fy1 in formula_boxes:
        norm_bbox = normalize_bbox_to_1000((fx0, fy0, fx1, fy1), frame.width, frame.height)
        crop_img = frame.crop((fx0, fy0, max(fx0 + 1, fx1), max(fy0 + 1, fy1)))
        try:
            latex = math_local.recognize_formula_crop(crop_img)
            latex_to_omml(latex)
            span = _Span(kind="math", content=latex, bbox=norm_bbox)
            recognized_local += 1
        except Exception as exc:
            logger.debug("Local formula crop failed (%s)", type(exc).__name__)
            span = _Span(kind="math", content="__UNRESOLVED_MATH__", bbox=norm_bbox)
            unresolved_count += 1
        dated_blocks.append(
            (
                fy0,
                fx0,
                _Block(
                    kind="equation",
                    spans=[span],
                    rows=[],
                    bbox=norm_bbox,
                    alignment="center",
                    font_size=11.5,
                ),
            )
        )

    for tbl in tables:
        tbbox = tbl.bbox
        norm_tbbox = normalize_bbox_to_1000(tbbox, frame.width, frame.height)
        rows: list[_Row] = []
        for r_idx, row in enumerate(tbl.cells):
            cells: list[_Cell] = []
            for c_idx, cell_val in enumerate(row):
                cleaned = normalize_ocr_text(cell_val or "").strip()
                if not cleaned:
                    cells.append(_Cell(spans=[_Span(kind="text", content="")]))
                    continue
                if is_likely_formula_text(cleaned):
                    # Missing/invalid cell geometry must not masquerade as plain
                    # text success. Without a safe crop the writer preserves
                    # the enclosing table and marks it unresolved.
                    cell_norm_bbox = None
                    if r_idx >= len(tbl.y_lines) - 1 or c_idx >= len(tbl.x_lines) - 1:
                        cells.append(_Cell(spans=[_Span(kind="math", content="__UNRESOLVED_MATH__")]))
                        unresolved_count += 1
                        warnings.append(f"Ô bảng hàng {r_idx + 1}, cột {c_idx + 1}: thiếu tọa độ công thức; cần đối chiếu ảnh bảng.")
                        continue
                    cx0, cy0 = tbl.x_lines[c_idx], tbl.y_lines[r_idx]
                    cx1, cy1 = tbl.x_lines[c_idx + 1], tbl.y_lines[r_idx + 1]
                    cell_norm_bbox = normalize_bbox_to_1000((cx0, cy0, cx1, cy1), frame.width, frame.height)
                    cell_crop = frame.crop((cx0, cy0, max(cx0 + 1, cx1), max(cy0 + 1, cy1)))
                    try:
                        latex = math_local.recognize_formula_crop(cell_crop)
                        latex_to_omml(latex)
                        cells.append(_Cell(spans=[_Span(kind="math", content=latex, bbox=cell_norm_bbox)]))
                        recognized_local += 1
                        continue
                    except Exception:
                        cells.append(_Cell(spans=[_Span(kind="math", content="__UNRESOLVED_MATH__", bbox=cell_norm_bbox)]))
                        unresolved_count += 1
                        warnings.append(f"Ô bảng hàng {r_idx + 1}, cột {c_idx + 1}: chưa tạo được Equation; đã giữ ảnh vùng công thức.")
                        continue
                cells.append(_Cell(spans=[_Span(kind="text", content=cleaned, bold=(r_idx == 0))]))
            rows.append(_Row(cells=cells))
        dated_blocks.append(
            (
                tbbox[1],
                tbbox[0],
                _Block(
                    kind="table",
                    spans=[],
                    rows=rows,
                    bbox=norm_tbbox,
                    alignment="left",
                    font_size=10.0,
                ),
            )
        )

    dated_blocks.sort(key=lambda item: (item[0], item[1]))
    blocks = [item[2] for item in dated_blocks]
    if not blocks or (recognized_local == 0 and unresolved_count == 0):
        return _Page(has_math=False, blocks=[], uncertain=False), 0, 0, warnings

    return _Page(has_math=True, blocks=blocks, uncertain=False), recognized_local, unresolved_count, warnings


def convert_pdf_with_math(
    input_path: Path,
    output_path: Path,
    options: ConversionOptions,
) -> ConversionResult:
    """Convert a PDF with mathematical formula recognition using the selected `math_engine`."""
    options.validate()
    engine = options.resolved_math_engine()
    import pymupdf
    from app.services.docx_merge import merge_docx_parts
    from .pdf import _should_ocr_page, convert_pdf

    try:
        pdf = pymupdf.open(input_path)
    except Exception as exc:
        raise ConversionError("Không thể mở PDF để nhận diện công thức.") from exc

    warnings: list[str] = []
    equations = tables = fallback_regions = fallback_pages = 0
    recognized_by_rules = recognized_by_local = unresolved_regions = 0
    scan_pages_seen = 0

    with pdf, tempfile.TemporaryDirectory(prefix="ocr2word-math-local-") as directory:
        if not len(pdf):
            raise ConversionError("PDF không có trang nào.")
        parts: list[Path] = []
        total_pages = len(pdf)

        for index, source_page in enumerate(pdf):
            page_label = f"Trang {index + 1}"
            options.report_progress(
                5 + (index / total_pages) * 86,
                f"Đang phân tích cấu trúc và công thức trang {index + 1}/{total_pages}…",
            )
            rect = source_page.rect
            scale = min(options.dpi / 72, MAX_PAGE_EDGE / max(rect.width, rect.height))
            pixmap = source_page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), colorspace=pymupdf.csRGB, alpha=False)
            image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
            part = Path(directory) / f"page-{index:05d}.docx"
            parts.append(part)

            raw_text = "".join(source_page.get_text("text", sort=True).split())
            is_scan = _should_ocr_page(source_page, raw_text, options)

            if not is_scan:
                digital_res = extract_digital_math_page(
                    source_page,
                    detect_tables=options.detect_tables,
                    preserve_images=options.preserve_images,
                )
                if digital_res.has_math:
                    page_ir = digital_res.page_model
                    page_local_upgraded = 0
                    page_unres = len(digital_res.unresolved_regions)
                    if engine == "local_hybrid" and page_unres > 0:
                        from . import math_local

                        math_local.ensure_local_math_model_ready()
                        page_local_upgraded, page_unres = _upgrade_unresolved_spans_with_local(page_ir, image)

                    document, count, table_count, fb_count, page_warnings = _write_page_detailed(
                        page_ir,
                        image,
                        input_path.name,
                        rect.width / 72,
                        rect.height / 72,
                        page_label=page_label,
                    )
                    document.save(part)
                    actual_local = min(count, page_local_upgraded)
                    actual_rules = max(0, count - actual_local)
                    equations += count
                    recognized_by_rules += actual_rules
                    recognized_by_local += actual_local
                    tables += table_count
                    fallback_regions += fb_count
                    unresolved_regions += max(page_unres, fb_count)
                    warnings.extend(page_warnings)
                    continue

            if is_scan:
                scan_pages_seen += 1
                if engine == "rules":
                    _source_image_page(image, input_path.name, rect.width / 72, rect.height / 72, part)
                    fallback_pages += 1
                    unresolved_regions += 1
                    warnings.append(
                        f"{page_label}: là trang scan/ảnh không có lớp chữ số; chế độ 'Theo quy tắc PDF' "
                        "không chạy mô hình nhận diện toán. Đã giữ ảnh trang gốc."
                    )
                    continue

                # engine == "local_hybrid" on a scanned page
                from . import math_local

                math_local.ensure_local_math_model_ready()
                page_ir, local_cnt, unres_cnt, scan_warnings = _build_local_raster_page_ir(image, options)
                warnings.extend(f"{page_label}: {w}" for w in scan_warnings)
                if page_ir.has_math:
                    if options.preserve_images:
                        for im in source_page.get_images():
                            xref = im[0]
                            for r in source_page.get_image_rects(xref):
                                norm_box = normalize_bbox_to_1000((r.x0, r.y0, r.x1, r.y1), rect.width, rect.height)
                                if not any(rect_overlap_ratio(norm_box, b.bbox) > 0.6 for b in page_ir.blocks if b.bbox):
                                    page_ir.blocks.append(_Block(
                                        kind="figure",
                                        spans=[],
                                        rows=[],
                                        bbox=norm_box,
                                        alignment="center",
                                        font_size=11.0,
                                    ))
                    document, count, table_count, fb_count, page_warnings = _write_page_detailed(
                        page_ir,
                        image,
                        input_path.name,
                        rect.width / 72,
                        rect.height / 72,
                        page_label=page_label,
                    )
                    document.save(part)
                    equations += count
                    recognized_by_local += count
                    tables += table_count
                    fallback_regions += fb_count
                    unresolved_regions += max(unres_cnt, fb_count)
                    warnings.extend(page_warnings)
                    continue

            # Page without math: keep regular PDF converter path
            subset = Path(directory) / f"page-{index:05d}.pdf"
            with pymupdf.open() as single:
                single.insert_pdf(pdf, from_page=index, to_page=index)
                single.save(subset)
            try:
                result = convert_pdf(subset, part, replace(options, math_ocr=False, progress_callback=None))
                warnings.extend(f"{page_label}: {warning}" for warning in result.warnings)
                tables += result.table_count or 0
                if options.preserve_images and is_scan and source_page.get_images():
                    try:
                        part_doc = Document(part)
                        if not part_doc.inline_shapes:
                            for im in source_page.get_images():
                                xref = im[0]
                                for r in source_page.get_image_rects(xref):
                                    norm_box = normalize_bbox_to_1000((r.x0, r.y0, r.x1, r.y1), rect.width, rect.height)
                                    crop = _crop(image, norm_box)
                                    p = part_doc.add_paragraph()
                                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    _add_picture(p, crop, min(rect.width / 72 - 0.9, (rect.width / 72 - 0.9) * (norm_box[2] - norm_box[0]) / 1000))
                            part_doc.save(part)
                    except Exception:
                        pass
            except ConversionError:
                _source_image_page(image, input_path.name, rect.width / 72, rect.height / 72, part)
                fallback_pages += 1
                warnings.append(f"{page_label}: đã giữ ảnh gốc vì OCR chữ chưa sẵn sàng.")

        if equations == 0 and (fallback_pages > 0 or fallback_regions > 0 or unresolved_regions > 0):
            if engine == "rules" and scan_pages_seen > 0:
                raise ConversionError(
                    "Chế độ 'Theo quy tắc PDF' chỉ hỗ trợ PDF có lớp chữ số; tài liệu scan/ảnh "
                    "cần chọn 'Cục bộ (Không dùng API)' để nhận diện công thức."
                )
            raise ConversionError(
                "Chưa tạo được Equation: tài liệu có vùng công thức ngoài phạm vi hỗ trợ hoặc lỗi mã hóa font. "
                "Đã dừng để không xuất file chỉ chứa ảnh thay cho công thức chỉnh sửa được."
            )

        page_count = len(pdf)
        options.report_progress(94, "Đang ghép các trang và Equation vào Word…")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if len(parts) == 1:
            shutil.copyfile(parts[0], output_path)
        else:
            merge_docx_parts(parts, output_path, page_breaks=options.page_breaks)

    verified_equations = _count_docx_equations(output_path)
    if verified_equations > 0:
        equations = verified_equations

    # Scan final DOCX for formula fragment paragraphs (Phase B.4)
    fragment_paras_found = 0
    try:
        final_doc = Document(output_path)
        for p in final_doc.paragraphs:
            ptxt = p.text.strip()
            # If paragraph contains an oMath, it's an equation, not a text fragment
            if p._element.xpath(".//m:oMath"):
                continue
            if is_fragment_text(ptxt):
                fragment_paras_found += 1
    except Exception:
        pass

    if fragment_paras_found > 0:
        warnings.append(
            f"Phát hiện {fragment_paras_found} đoạn văn có dấu hiệu là mảnh ký tự/chữ số của công thức; "
            "đã đánh dấu partial để người dùng đối chiếu."
        )

    status = _math_status(equations, fallback_pages, fallback_regions, fragment_count=fragment_paras_found)

    warnings.insert(
        0,
        _engine_summary_warning(
            engine,
            equations,
            fallback_regions,
            source_kind="PDF",
            fallback_pages=fallback_pages,
            recognized_by_rules=recognized_by_rules,
            recognized_by_local=recognized_by_local,
        ),
    )
    options.report_progress(100, "Đã tạo Word với công thức toán.")
    return ConversionResult(
        output_path=output_path,
        warnings=list(dict.fromkeys(warnings)),
        page_count=page_count,
        table_count=tables,
        equation_count=equations,
        math_fallback_pages=fallback_pages,
        math_fallback_regions=fallback_regions,
        math_status=status,
        math_engine=engine,
        math_recognized_by_rules=recognized_by_rules,
        math_recognized_by_local=recognized_by_local,
        math_unresolved_regions=unresolved_regions,
        fragment_text_count=fragment_paras_found,
    )


def convert_image_with_math(
    input_path: Path,
    output_path: Path,
    options: ConversionOptions,
) -> ConversionResult:
    """Convert an image (including multi-frame TIFF) with math recognition using the selected `math_engine`."""
    options.validate()
    engine = options.resolved_math_engine()
    if engine == "rules":
        raise ConversionError(
            "Chế độ 'Theo quy tắc PDF' chỉ hỗ trợ PDF có lớp chữ số; với ảnh hoặc PDF scan, "
            "hãy chọn 'Cục bộ (Không dùng API)'."
        )

    from app.services.docx_merge import merge_docx_parts
    from . import math_local
    from .ocr_layout import normalize_document_ocr_text
    from .raster_tables import write_editable_ocr_page

    math_local.ensure_local_math_model_ready()
    options.report_progress(2, "Đang đọc hình ảnh…")
    try:
        source = Image.open(input_path)
    except (UnidentifiedImageError, OSError) as exc:
        raise ConversionError(f"Không thể mở hình ảnh: {exc}") from exc

    warnings: list[str] = []
    equations = tables = fallback_regions = fallback_pages = 0
    recognized_by_local = unresolved_regions = 0

    try:
        frame_count = max(1, int(getattr(source, "n_frames", 1)))
        with tempfile.TemporaryDirectory(prefix="ocr2word-math-local-img-") as directory:
            parts: list[Path] = []
            for frame_index in range(frame_count):
                source.seek(frame_index)
                frame = prepare_math_image(source.copy())
                landscape = frame.width > frame.height
                width_inches = 11.69 if landscape else 8.27
                height_inches = 8.27 if landscape else 11.69
                label = "Ảnh" if frame_count == 1 else f"Khung ảnh {frame_index + 1}"
                options.report_progress(
                    5 + (frame_index / frame_count) * 86,
                    f"Đang nhận diện chữ và công thức ảnh {frame_index + 1}/{frame_count}…",
                )
                part = Path(directory) / f"frame-{frame_index:05d}.docx"
                parts.append(part)

                page_ir, local_cnt, unres_cnt, frame_ocr_warnings = _build_local_raster_page_ir(frame, options)
                warnings.extend(f"{label}: {w}" for w in frame_ocr_warnings)

                if page_ir.has_math:
                    document, count, table_count, fb_count, page_warnings = _write_page_detailed(
                        page_ir,
                        frame,
                        input_path.name,
                        width_inches,
                        height_inches,
                        page_label=label,
                    )
                    document.save(part)
                    equations += count
                    recognized_by_local += count
                    tables += table_count
                    fallback_regions += fb_count
                    unresolved_regions += max(unres_cnt, fb_count)
                    warnings.extend(page_warnings)
                    continue

                # Frame without math: use standard OCR with math_ocr=False
                fallback_options = replace(options, math_ocr=False, progress_callback=None)
                content_width_dxa = LANDSCAPE_CONTENT_WIDTH_DXA if landscape else PORTRAIT_CONTENT_WIDTH_DXA
                content_width_inches = content_width_dxa / 1440
                frame_doc = create_document(
                    input_path.name,
                    landscape=landscape,
                    body_font="Times New Roman",
                )
                try:
                    line_count, table_count, frame_warnings = write_editable_ocr_page(
                        frame_doc,
                        frame,
                        fallback_options,
                        content_width_inches=content_width_inches,
                        content_width_dxa=content_width_dxa,
                    )
                    if line_count == 0 and table_count == 0:
                        _source_image_page(frame, input_path.name, width_inches, height_inches, part)
                        fallback_pages += 1
                        warnings.append(f"{label}: OCR không tìm thấy văn bản đủ tin cậy; đã giữ ảnh gốc.")
                    else:
                        normalize_document_ocr_text(frame_doc)
                        frame_doc.save(part)
                        tables += table_count
                        warnings.extend(f"{label}: {w}" for w in frame_warnings)
                except ConversionError:
                    _source_image_page(frame, input_path.name, width_inches, height_inches, part)
                    fallback_pages += 1
                    warnings.append(f"{label}: đã giữ ảnh gốc vì OCR chữ chưa sẵn sàng.")

            if equations == 0 and (fallback_pages > 0 or fallback_regions > 0 or unresolved_regions > 0):
                raise ConversionError(
                    "Chưa tạo được Equation từ hình ảnh này; mô hình cục bộ không nhận diện được cấu trúc công thức hợp lệ."
                )

            options.report_progress(94, "Đang tạo file Word từ hình ảnh…")
            output_path.parent.mkdir(parents=True, exist_ok=True)
            if len(parts) == 1:
                shutil.copyfile(parts[0], output_path)
            else:
                merge_docx_parts(parts, output_path, page_breaks=options.page_breaks)
    finally:
        source.close()

    verified_equations = _count_docx_equations(output_path)
    if verified_equations > 0:
        equations = verified_equations

    fragment_paras_found = 0
    try:
        final_doc = Document(output_path)
        for p in final_doc.paragraphs:
            ptxt = p.text.strip()
            if p._element.xpath(".//m:oMath"):
                continue
            if is_fragment_text(ptxt):
                fragment_paras_found += 1
    except Exception:
        pass

    if fragment_paras_found > 0:
        warnings.append(
            f"Phát hiện {fragment_paras_found} đoạn văn có dấu hiệu là mảnh ký tự/chữ số của công thức; "
            "đã đánh dấu partial để người dùng đối chiếu."
        )

    status = _math_status(equations, fallback_pages, fallback_regions, fragment_count=fragment_paras_found)

    warnings.insert(
        0,
        _engine_summary_warning(
            engine,
            equations,
            fallback_regions,
            source_kind="ảnh gốc",
            fallback_pages=fallback_pages,
            recognized_by_rules=0,
            recognized_by_local=recognized_by_local,
        ),
    )
    options.report_progress(100, "Đã tạo xong file Word từ hình ảnh.")
    return ConversionResult(
        output_path=output_path,
        warnings=list(dict.fromkeys(warnings)),
        page_count=frame_count,
        table_count=tables,
        equation_count=equations,
        math_fallback_pages=fallback_pages,
        math_fallback_regions=fallback_regions,
        math_status=status,
        math_engine=engine,
        math_recognized_by_rules=0,
        math_recognized_by_local=recognized_by_local,
        math_unresolved_regions=unresolved_regions,
        fragment_text_count=fragment_paras_found,
    )
