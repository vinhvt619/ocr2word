from __future__ import annotations

from bisect import bisect_right
from dataclasses import dataclass
from statistics import median
from typing import Any, Iterable

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from PIL import Image as PillowImage

from app.services.ocr import detect_image_rotation, image_region_to_data, image_to_data

from .base import ConversionError, ConversionOptions, make_word_table, set_cell_margins, set_run_font
from .ocr_layout import (
    merge_ocr_lines,
    normalize_ocr_text,
    ocr_data_to_words,
    ocr_words_to_lines,
    ocr_words_to_text,
    write_ocr_lines,
)

try:
    import cv2
    import numpy as np
except ImportError:  # pragma: no cover - covered by dependency diagnostics
    cv2 = None
    np = None


@dataclass(slots=True)
class RasterTable:
    bbox: tuple[int, int, int, int]
    x_lines: list[int]
    y_lines: list[int]
    cells: list[list[str]]
    merges: list[tuple[int, int, int, int]]

    @property
    def row_count(self) -> int:
        return len(self.y_lines) - 1

    @property
    def column_count(self) -> int:
        return len(self.x_lines) - 1


class _DisjointSet:
    def __init__(self, size: int) -> None:
        self.parent = list(range(size))

    def find(self, value: int) -> int:
        while self.parent[value] != value:
            self.parent[value] = self.parent[self.parent[value]]
            value = self.parent[value]
        return value

    def union(self, left: int, right: int) -> None:
        left_root = self.find(left)
        right_root = self.find(right)
        if left_root != right_root:
            self.parent[right_root] = left_root


def _run_centers(mask) -> list[int]:
    indices = np.flatnonzero(mask)
    if not len(indices):
        return []
    centers: list[int] = []
    start = previous = int(indices[0])
    for raw in indices[1:]:
        value = int(raw)
        if value > previous + 1:
            centers.append(round((start + previous) / 2))
            start = value
        previous = value
    centers.append(round((start + previous) / 2))
    return centers


def _dedupe_positions(values: Iterable[int], tolerance: int) -> list[int]:
    clusters: list[list[int]] = []
    for value in sorted(int(item) for item in values):
        if not clusters or value - clusters[-1][-1] > tolerance:
            clusters.append([value])
        else:
            clusters[-1].append(value)
    return [round(sum(cluster) / len(cluster)) for cluster in clusters]


def _profile_lines(mask, bbox: tuple[int, int, int, int], *, vertical: bool) -> list[int]:
    left, top, right, bottom = bbox
    crop = mask[top : bottom + 1, left : right + 1]
    if crop.size == 0:
        return []
    profile = np.count_nonzero(crop, axis=0 if vertical else 1)
    span = crop.shape[0] if vertical else crop.shape[1]
    offset = left if vertical else top
    tolerance = max(3, round((right - left if vertical else bottom - top) * 0.003))
    best: list[int] = []
    # Prefer strong full-width/full-height rules. Only relax the threshold when
    # a low-contrast scan does not expose the three lines needed for a grid.
    # This avoids treating small boxes inside a flowchart as extra table columns.
    for coverage in (0.16, 0.10, 0.07):
        threshold = max(7, round(span * coverage))
        centers = _run_centers(profile >= threshold)
        positions = _dedupe_positions((offset + center for center in centers), tolerance)
        if len(positions) > len(best):
            best = positions
        if len(positions) >= 3:
            return positions
    return best


def _split_line_clusters(lines: list[int], page_extent: int) -> list[list[int]]:
    if len(lines) < 4:
        return [lines]
    gaps = [right - left for left, right in zip(lines, lines[1:])]
    typical = float(np.median(gaps)) if gaps else 0.0
    threshold = max(page_extent * 0.12, typical * 3.6)
    clusters: list[list[int]] = [[lines[0]]]
    for index, gap in enumerate(gaps):
        if gap > threshold:
            clusters.append([])
        clusters[-1].append(lines[index + 1])
    valid = [cluster for cluster in clusters if len(cluster) >= 3]
    return valid or [lines]


def _line_present(mask, *, x0: int, y0: int, x1: int, y1: int, vertical: bool) -> bool:
    if vertical:
        top, bottom = sorted((y0, y1))
        padding = min(4, max(0, (bottom - top) // 8))
        radius = max(7, round(mask.shape[1] * 0.004))
        band = mask[
            top + padding : bottom - padding + 1,
            max(0, x0 - radius) : min(mask.shape[1], x0 + radius + 1),
        ]
        if band.size == 0:
            return False
        coverage = np.count_nonzero(np.any(band > 0, axis=1)) / max(1, band.shape[0])
    else:
        left, right = sorted((x0, x1))
        padding = min(4, max(0, (right - left) // 8))
        radius = max(7, round(mask.shape[0] * 0.004))
        band = mask[
            max(0, y0 - radius) : min(mask.shape[0], y0 + radius + 1),
            left + padding : right - padding + 1,
        ]
        if band.size == 0:
            return False
        coverage = np.count_nonzero(np.any(band > 0, axis=0)) / max(1, band.shape[1])
    return coverage >= 0.24


def _vertical_support_span(
    mask,
    x: int,
    candidate_top: int,
    candidate_bottom: int,
) -> tuple[int, int] | None:
    radius = max(5, round(mask.shape[1] * 0.0025))
    band = mask[:, max(0, x - radius) : min(mask.shape[1], x + radius + 1)]
    if band.size == 0:
        return None
    indices = np.flatnonzero(np.any(band > 0, axis=1))
    if not len(indices):
        return None
    gap_tolerance = max(5, round(mask.shape[0] * 0.004))
    spans: list[tuple[int, int]] = []
    start = previous = int(indices[0])
    for raw in indices[1:]:
        value = int(raw)
        if value > previous + gap_tolerance:
            spans.append((start, previous))
            start = value
        previous = value
    spans.append((start, previous))

    def overlap(span: tuple[int, int]) -> int:
        return max(0, min(span[1], candidate_bottom) - max(span[0], candidate_top))

    best = max(spans, key=lambda span: (overlap(span), span[1] - span[0]))
    if overlap(best) < max(12, round((candidate_bottom - candidate_top) * 0.55)):
        return None
    return best


def _expand_partial_table_grid(
    x_lines: list[int],
    y_lines: list[int],
    vertical_mask,
    horizontal_mask,
) -> tuple[list[int], list[int]]:
    """Recover the top of a table when a flowchart splits its grid component."""

    if len(x_lines) < 3 or len(y_lines) < 3:
        return x_lines, y_lines
    left_span = _vertical_support_span(vertical_mask, x_lines[0], y_lines[0], y_lines[-1])
    right_span = _vertical_support_span(vertical_mask, x_lines[-1], y_lines[0], y_lines[-1])
    if left_span is None or right_span is None:
        return x_lines, y_lines
    expanded_top = max(left_span[0], right_span[0])
    expanded_bottom = min(left_span[1], right_span[1])
    original_height = max(1, y_lines[-1] - y_lines[0])
    if (
        expanded_bottom <= expanded_top
        or y_lines[0] - expanded_top < max(12, round(original_height * 0.12))
    ):
        return x_lines, y_lines

    bbox = (x_lines[0], expanded_top, x_lines[-1], expanded_bottom)
    expanded_y = _profile_lines(horizontal_mask, bbox, vertical=False)
    expanded_x = _profile_lines(vertical_mask, bbox, vertical=True)
    if len(expanded_y) < len(y_lines) + 1 or len(expanded_x) < 3:
        return x_lines, y_lines

    height = max(1, expanded_bottom - expanded_top + 1)
    radius = max(5, round(vertical_mask.shape[1] * 0.0025))
    strong_x: list[int] = []
    for x in expanded_x:
        band = vertical_mask[
            expanded_top : expanded_bottom + 1,
            max(0, x - radius) : min(vertical_mask.shape[1], x + radius + 1),
        ]
        coverage = np.count_nonzero(np.any(band > 0, axis=1)) / height
        if coverage >= 0.48:
            strong_x.append(x)
    if len(strong_x) < 3:
        return x_lines, y_lines
    return strong_x, expanded_y


def _merge_map(
    x_lines: list[int],
    y_lines: list[int],
    vertical_mask,
    horizontal_mask,
) -> tuple[list[tuple[int, int, int, int]], dict[tuple[int, int], tuple[int, int]]]:
    rows = len(y_lines) - 1
    columns = len(x_lines) - 1
    groups = _DisjointSet(rows * columns)

    def cell_id(row: int, column: int) -> int:
        return row * columns + column

    for row in range(rows):
        for boundary in range(1, columns):
            if not _line_present(
                vertical_mask,
                x0=x_lines[boundary],
                y0=y_lines[row],
                x1=x_lines[boundary],
                y1=y_lines[row + 1],
                vertical=True,
            ):
                groups.union(cell_id(row, boundary - 1), cell_id(row, boundary))

    for boundary in range(1, rows):
        for column in range(columns):
            if not _line_present(
                horizontal_mask,
                x0=x_lines[column],
                y0=y_lines[boundary],
                x1=x_lines[column + 1],
                y1=y_lines[boundary],
                vertical=False,
            ):
                groups.union(cell_id(boundary - 1, column), cell_id(boundary, column))

    members: dict[int, set[tuple[int, int]]] = {}
    for row in range(rows):
        for column in range(columns):
            members.setdefault(groups.find(cell_id(row, column)), set()).add((row, column))

    merges: list[tuple[int, int, int, int]] = []
    canonical: dict[tuple[int, int], tuple[int, int]] = {}
    for cells in members.values():
        min_row = min(row for row, _column in cells)
        max_row = max(row for row, _column in cells)
        min_col = min(column for _row, column in cells)
        max_col = max(column for _row, column in cells)
        rectangle = {
            (row, column)
            for row in range(min_row, max_row + 1)
            for column in range(min_col, max_col + 1)
        }
        anchor = (min_row, min_col)
        if cells == rectangle:
            for cell in cells:
                canonical[cell] = anchor
            if len(cells) > 1:
                merges.append((min_row, min_col, max_row, max_col))
        else:
            for cell in cells:
                canonical[cell] = cell
    return merges, canonical


def _contains(outer: tuple[int, int, int, int], inner: tuple[int, int, int, int]) -> bool:
    return (
        inner[0] >= outer[0]
        and inner[1] >= outer[1]
        and inner[2] <= outer[2]
        and inner[3] <= outer[3]
    )


def _table_candidates(image) -> tuple[list[tuple[list[int], list[int]]], Any, Any]:
    if cv2 is None or np is None:
        return [], None, None

    source_image = image.convert("RGB")
    source_width, source_height = source_image.size
    source_extent = max(source_width, source_height)
    analysis_scale = 1.0 if source_extent <= 2600 else 1450 / source_extent
    if analysis_scale < 1.0:
        analysis_image = source_image.resize(
            (
                max(1, round(source_width * analysis_scale)),
                max(1, round(source_height * analysis_scale)),
            ),
            PillowImage.Resampling.LANCZOS,
        )
    else:
        analysis_image = source_image

    rgb = np.asarray(analysis_image)
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    block_size = max(31, (min(gray.shape) // 45) | 1)
    binary = cv2.adaptiveThreshold(
        gray,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV,
        block_size,
        15,
    )
    height, width = gray.shape
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(18, width // 32), 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(18, height // 32)))
    horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel)
    vertical = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel)
    horizontal = cv2.morphologyEx(horizontal, cv2.MORPH_CLOSE, np.ones((1, 3), np.uint8))
    vertical = cv2.morphologyEx(vertical, cv2.MORPH_CLOSE, np.ones((3, 1), np.uint8))
    grid = cv2.bitwise_or(horizontal, vertical)

    count, _labels, stats, _centroids = cv2.connectedComponentsWithStats(grid, 8)
    components: list[tuple[int, int, int, int]] = []
    for index in range(1, count):
        left, top, component_width, component_height, _area = [int(value) for value in stats[index]]
        if component_width < width * 0.20 or component_height < max(35, height * 0.018):
            continue
        components.append(
            (
                max(0, left - 3),
                max(0, top - 3),
                min(width - 1, left + component_width + 2),
                min(height - 1, top + component_height + 2),
            )
        )

    raw_candidates: list[tuple[list[int], list[int]]] = []
    for bbox in sorted(components, key=lambda item: (item[2] - item[0]) * (item[3] - item[1]), reverse=True):
        y_lines = _profile_lines(horizontal, bbox, vertical=False)
        for y_cluster in _split_line_clusters(y_lines, height):
            if len(y_cluster) < 3:
                continue
            tightened = (bbox[0], y_cluster[0], bbox[2], y_cluster[-1])
            x_lines = _profile_lines(vertical, tightened, vertical=True)
            x_lines, y_cluster = _expand_partial_table_grid(
                x_lines,
                y_cluster,
                vertical,
                horizontal,
            )
            if not 3 <= len(x_lines) <= 21:
                continue
            if not 3 <= len(y_cluster) <= 81:
                continue
            if (len(x_lines) - 1) * (len(y_cluster) - 1) > 1200:
                continue
            raw_candidates.append((x_lines, y_cluster))

    accepted: list[tuple[list[int], list[int]]] = []
    accepted_boxes: list[tuple[int, int, int, int]] = []
    for x_lines, y_lines in sorted(
        raw_candidates,
        key=lambda item: (item[0][-1] - item[0][0]) * (item[1][-1] - item[1][0]),
        reverse=True,
    ):
        bbox = (x_lines[0], y_lines[0], x_lines[-1], y_lines[-1])
        if any(_contains(existing, bbox) for existing in accepted_boxes):
            continue
        accepted.append((x_lines, y_lines))
        accepted_boxes.append(bbox)
    accepted.sort(key=lambda item: (item[1][0], item[0][0]))
    if analysis_scale < 1.0:
        inverse_scale = 1.0 / analysis_scale
        accepted = [
            (
                _dedupe_positions((round(value * inverse_scale) for value in x_lines), 2),
                _dedupe_positions((round(value * inverse_scale) for value in y_lines), 2),
            )
            for x_lines, y_lines in accepted
        ]
        vertical = cv2.resize(
            vertical,
            (source_width, source_height),
            interpolation=cv2.INTER_NEAREST,
        )
        horizontal = cv2.resize(
            horizontal,
            (source_width, source_height),
            interpolation=cv2.INTER_NEAREST,
        )
    return accepted, vertical, horizontal


def _table_region_words(
    image,
    bbox: tuple[int, int, int, int],
    vertical_mask,
    horizontal_mask,
    requested_language: str,
) -> tuple[list[dict[str, Any]], list[str]]:
    left, top, right, bottom = bbox
    rgb = np.asarray(image.convert("RGB"))[top : bottom + 1, left : right + 1].copy()
    if rgb.size == 0:
        return [], []

    line_mask = cv2.bitwise_or(vertical_mask, horizontal_mask)[top : bottom + 1, left : right + 1]
    padding = max(3, round(max(image.size) / 1000))
    line_mask = cv2.dilate(line_mask, np.ones((padding, padding), np.uint8))
    rgb[line_mask > 0] = 255
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    gray = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(gray)
    cleaned = PillowImage.fromarray(cv2.cvtColor(gray, cv2.COLOR_GRAY2RGB))
    try:
        data, _language, warnings = image_region_to_data(
            cleaned,
            requested_language,
            psm=11,
        )
    except ConversionError:
        return [], []

    region_words = ocr_data_to_words(data, confidence_threshold=35)
    for word in region_words:
        word["left"] += left
        word["right"] += left
        word["top"] += top
        word["bottom"] += top
    return region_words, warnings


def _group_table_words(
    words: Iterable[dict[str, Any]],
    x_lines: list[int],
    y_lines: list[int],
    canonical: dict[tuple[int, int], tuple[int, int]],
) -> dict[tuple[int, int], list[dict[str, Any]]]:
    rows = len(y_lines) - 1
    columns = len(x_lines) - 1
    grouped: dict[tuple[int, int], list[dict[str, Any]]] = {}
    for word in words:
        center_x = (word["left"] + word["right"]) / 2
        center_y = (word["top"] + word["bottom"]) / 2
        if not (x_lines[0] <= center_x <= x_lines[-1] and y_lines[0] <= center_y <= y_lines[-1]):
            continue
        column = bisect_right(x_lines, center_x) - 1
        row = bisect_right(y_lines, center_y) - 1
        if row == rows:
            row -= 1
        if column == columns:
            column -= 1
        if 0 <= row < rows and 0 <= column < columns:
            grouped.setdefault(canonical[(row, column)], []).append(word)
    return grouped


def _cell_candidate(words: list[dict[str, Any]]) -> tuple[str, float, int]:
    if not words:
        return "", 0.0, 0
    text = normalize_ocr_text(ocr_words_to_text(words))
    useful = sum(character.isalnum() for character in text)
    confidence = sum(float(word.get("confidence", 0)) for word in words) / max(1, len(words))
    return text, confidence, useful


def _best_cell_text(
    page_words: list[dict[str, Any]],
    clean_words: list[dict[str, Any]],
) -> str:
    page_text, page_confidence, page_useful = _cell_candidate(page_words)
    clean_text, clean_confidence, clean_useful = _cell_candidate(clean_words)
    if not clean_text:
        return page_text
    if not page_text:
        return clean_text
    if clean_useful >= max(1, round(page_useful * 0.65)) and clean_confidence >= page_confidence - 12:
        return clean_text
    page_score = page_confidence + min(24, page_useful * 0.25)
    clean_score = clean_confidence + min(24, clean_useful * 0.25)
    return clean_text if clean_score >= page_score else page_text


def _fill_missing_stt_sequence(cells: list[list[str]]) -> None:
    """Restore obvious 1..N row labels when OCR only sees part of an STT column."""

    if not 4 <= len(cells) <= 14 or not cells or not cells[0]:
        return
    if normalize_ocr_text(cells[0][0]).upper() != "STT":
        return
    if not any("CỘNG" in normalize_ocr_text(value).upper() for value in cells[-1]):
        return
    observed: dict[int, int] = {}
    for row_index in range(1, len(cells) - 1):
        value = normalize_ocr_text(cells[row_index][0])
        if value.isdigit():
            observed[row_index] = int(value)
    if not observed or not any(row_index == value for row_index, value in observed.items()):
        return
    for row_index in range(1, len(cells) - 1):
        value = normalize_ocr_text(cells[row_index][0])
        if not value or (len(value) <= 2 and not value.isdigit()):
            cells[row_index][0] = str(row_index)


def _fill_missing_step_sequence(
    cells: list[list[str]],
    merges: list[tuple[int, int, int, int]],
) -> None:
    if len(cells) < 3 or not cells[0]:
        return
    if normalize_ocr_text(cells[0][0]).casefold() != "bước":
        return
    covered_rows: set[int] = set()
    for top, left, bottom, right in merges:
        if left <= 0 <= right:
            covered_rows.update(range(top + 1, bottom + 1))
    anchors = [row for row in range(1, len(cells)) if row not in covered_rows]
    if not 2 <= len(anchors) <= 20:
        return
    for step, row_index in enumerate(anchors, start=1):
        if not normalize_ocr_text(cells[row_index][0]):
            cells[row_index][0] = str(step)


def detect_raster_tables(
    image,
    words: list[dict[str, Any]],
    requested_language: str | None = None,
    warnings: list[str] | None = None,
) -> list[RasterTable]:
    candidates, vertical_mask, horizontal_mask = _table_candidates(image)
    if not candidates:
        return []

    tables: list[RasterTable] = []
    for x_lines, y_lines in candidates:
        rows = len(y_lines) - 1
        columns = len(x_lines) - 1
        merges, canonical = _merge_map(x_lines, y_lines, vertical_mask, horizontal_mask)
        page_grouped = _group_table_words(words, x_lines, y_lines, canonical)
        clean_grouped: dict[tuple[int, int], list[dict[str, Any]]] = {}
        if requested_language:
            cleaned_words, region_warnings = _table_region_words(
                image,
                (x_lines[0], y_lines[0], x_lines[-1], y_lines[-1]),
                vertical_mask,
                horizontal_mask,
                requested_language,
            )
            clean_grouped = _group_table_words(cleaned_words, x_lines, y_lines, canonical)
            if warnings is not None:
                warnings.extend(region_warnings)

        cells = [["" for _column in range(columns)] for _row in range(rows)]
        for anchor in set(page_grouped) | set(clean_grouped):
            row, column = anchor
            cells[row][column] = _best_cell_text(
                page_grouped.get(anchor, []),
                clean_grouped.get(anchor, []),
            )
        _fill_missing_stt_sequence(cells)
        _fill_missing_step_sequence(cells, merges)
        tables.append(
            RasterTable(
                bbox=(x_lines[0], y_lines[0], x_lines[-1], y_lines[-1]),
                x_lines=x_lines,
                y_lines=y_lines,
                cells=cells,
                merges=merges,
            )
        )
    return tables


def _write_raster_table(doc, raster_table: RasterTable, content_width_dxa: int) -> None:
    source_widths = [
        raster_table.x_lines[index + 1] - raster_table.x_lines[index]
        for index in range(raster_table.column_count)
    ]
    total_source_width = max(1, sum(source_widths))
    widths = [max(360, round(content_width_dxa * width / total_source_width)) for width in source_widths]
    widths[-1] += content_width_dxa - sum(widths)
    table = make_word_table(
        doc,
        raster_table.row_count,
        raster_table.column_count,
        widths,
        header_rows=0,
    )
    compact = raster_table.column_count >= 9 or raster_table.row_count >= 6
    # Dense scan tables otherwise gain enough Word cell padding/font height to
    # spill a final row onto a nearly-empty page.  Nine points remains readable
    # at A4 width and more closely matches common 11-12 px source forms.
    font_size = 9.0 if raster_table.column_count <= 6 else 8.5 if raster_table.column_count <= 10 else 7
    if compact:
        for row in table.rows:
            for cell in row.cells:
                set_cell_margins(cell, top=35, start=45, bottom=35, end=45)
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.line_spacing = 1.0

    for top, left, bottom, right in raster_table.merges:
        try:
            table.cell(top, left).merge(table.cell(bottom, right))
        except (IndexError, ValueError):
            continue

    for row_index, row in enumerate(raster_table.cells):
        for column_index, value in enumerate(row):
            if not value:
                continue
            word_cell = table.cell(row_index, column_index)
            word_cell.text = value
            paragraph = word_cell.paragraphs[0]
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            if row_index == 0 or len(value) <= 12:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                set_run_font(run, "Times New Roman", font_size)
                run.bold = row_index == 0
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = 0


def _ocr_quality(data: dict[str, list[Any]]) -> tuple[int, float]:
    confidences: list[float] = []
    for text, raw_confidence in zip(data.get("text", []), data.get("conf", [])):
        if not str(text).strip():
            continue
        try:
            confidence = float(raw_confidence)
        except (TypeError, ValueError):
            continue
        if confidence >= 0:
            confidences.append(confidence)
    return len(confidences), median(confidences) if confidences else 0.0


def _autorotate_low_confidence_image(
    image,
    data: dict[str, list[Any]],
    requested_language: str,
) -> tuple[Any, dict[str, list[Any]], list[str]]:
    word_count, original_quality = _ocr_quality(data)
    if word_count < 20 or original_quality >= 65:
        return image, data, []

    rotation = detect_image_rotation(image)
    if rotation not in {90, 180, 270}:
        return image, data, []
    rotated = image.rotate(rotation, expand=True, fillcolor="white")
    rotated_data, _language, warnings = image_to_data(rotated, requested_language)
    rotated_count, rotated_quality = _ocr_quality(rotated_data)
    enough_words = rotated_count >= max(10, round(word_count * 0.70))
    if not enough_words or rotated_quality < original_quality + 10:
        return image, data, []

    display_rotation = rotation if rotation <= 180 else rotation - 360
    warnings.append(
        f"Đã tự xoay trang scan {abs(display_rotation)}° để OCR chữ và bảng đúng chiều."
    )
    return rotated, rotated_data, warnings


def write_editable_ocr_page(
    doc,
    image,
    options: ConversionOptions,
    *,
    content_width_inches: float,
    content_width_dxa: int,
    source_dpi: float | None = None,
) -> tuple[int, int, list[str]]:
    """Write OCR text and detected raster grids as editable Word content."""

    data, _language, warnings = image_to_data(image, options.language)
    image, data, rotation_warnings = _autorotate_low_confidence_image(
        image,
        data,
        options.language,
    )
    warnings.extend(rotation_warnings)
    words = ocr_data_to_words(data)
    tables: list[RasterTable] = []
    if options.detect_tables:
        if cv2 is None or np is None:
            warnings.append("Thiếu OpenCV nên chưa thể nhận diện bảng trong ảnh.")
        else:
            try:
                tables = detect_raster_tables(
                    image,
                    words,
                    requested_language=options.language,
                    warnings=warnings,
                )
            except Exception:
                warnings.append("Không thể phân tích lưới bảng; đã giữ nội dung dưới dạng đoạn văn OCR.")

    sparse_words: list[dict[str, Any]] = []
    try:
        sparse_data, _sparse_language, sparse_warnings = image_region_to_data(
            image,
            options.language,
            psm=11,
        )
        sparse_words = ocr_data_to_words(sparse_data, confidence_threshold=40)
        warnings.extend(sparse_warnings)
    except ConversionError:
        pass

    if not tables:
        lines = merge_ocr_lines(
            ocr_words_to_lines(words),
            ocr_words_to_lines(sparse_words),
        )
        line_count = write_ocr_lines(
            doc,
            lines,
            image_width=image.width,
            content_width_inches=content_width_inches,
            source_dpi=source_dpi,
            image=image,
        )
        return line_count, 0, warnings

    def inside_table(word: dict[str, Any]) -> bool:
        center_x = (word["left"] + word["right"]) / 2
        center_y = (word["top"] + word["bottom"]) / 2
        return any(
            left <= center_x <= right and top <= center_y <= bottom
            for left, top, right, bottom in (table.bbox for table in tables)
        )

    outside_lines = merge_ocr_lines(
        ocr_words_to_lines(word for word in words if not inside_table(word)),
        ocr_words_to_lines(word for word in sparse_words if not inside_table(word)),
    )
    items: list[tuple[int, int, Any]] = [
        (line["top"], 1, line) for line in outside_lines
    ] + [
        (table.bbox[1], 0, table) for table in tables
    ]
    items.sort(key=lambda item: (item[0], item[1]))

    line_buffer: list[dict[str, Any]] = []

    def flush_lines() -> None:
        if not line_buffer:
            return
        write_ocr_lines(
            doc,
            line_buffer,
            image_width=image.width,
            content_width_inches=content_width_inches,
            source_dpi=source_dpi,
            image=image,
        )
        line_buffer.clear()

    for _top, kind, payload in items:
        if kind == 1:
            line_buffer.append(payload)
        else:
            flush_lines()
            _write_raster_table(doc, payload, content_width_dxa)
    flush_lines()
    return len(outside_lines), len(tables), warnings
