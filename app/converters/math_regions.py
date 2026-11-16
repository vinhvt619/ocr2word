"""Shared region intermediate representation and geometric classifiers for math OCR.

Provides:
- `MathRegion`: structured representation of an inline, display, or table-cell
  formula region with status (`recognized`, `unsupported`, `failed`), reason
  code, LaTeX payload, and engine source (`rules`, `local`).
- Verified PDF table extraction vs. horizontal fraction-bar / underline
  classification so thin fraction bars or clip paths are never mistaken for
  table grids, and table borders/underlines are never mistaken for fractions.
- Classical raster/scan formula region segmentation (OpenCV + OCR line heuristics)
  for local math inference on scanned PDFs and images.
"""
from __future__ import annotations

import re
import warnings as python_warnings
from dataclasses import dataclass
from typing import Any, Iterable, Literal

import cv2
import numpy as np
from PIL import Image

from .base import clean_cell_text

MathRegionKind = Literal["inline", "display", "table_cell"]
MathRegionStatus = Literal["recognized", "unsupported", "failed"]
MathRegionSource = Literal["rules", "local"]


@dataclass(slots=True)
class MathRegion:
    bbox: tuple[float, float, float, float]
    kind: MathRegionKind = "display"
    status: MathRegionStatus = "recognized"
    reason: str | None = None
    latex: str | None = None
    source: MathRegionSource | None = None
    confidence: float = 0.0
    trailing_text: str | None = None


@dataclass(frozen=True, slots=True)
class HorizontalBar:
    bbox: tuple[float, float, float, float]

    @property
    def x0(self) -> float:
        return self.bbox[0]

    @property
    def y0(self) -> float:
        return self.bbox[1]

    @property
    def x1(self) -> float:
        return self.bbox[2]

    @property
    def y1(self) -> float:
        return self.bbox[3]

    @property
    def width(self) -> float:
        return max(0.0, self.bbox[2] - self.bbox[0])

    @property
    def height(self) -> float:
        return max(0.0, self.bbox[3] - self.bbox[1])

    @property
    def y_mid(self) -> float:
        return (self.bbox[1] + self.bbox[3]) * 0.5


def rect_overlap_ratio(
    a: tuple[float, float, float, float],
    b: tuple[float, float, float, float],
) -> float:
    left, top = max(a[0], b[0]), max(a[1], b[1])
    right, bottom = min(a[2], b[2]), min(a[3], b[3])
    if right <= left or bottom <= top:
        return 0.0
    intersection = (right - left) * (bottom - top)
    area = max(1.0, (a[2] - a[0]) * (a[3] - a[1]))
    return intersection / area


def horizontal_overlap_ratio(
    a: tuple[float, float, float, float],
    b: tuple[float, float, float, float],
) -> float:
    left = max(a[0], b[0])
    right = min(a[2], b[2])
    if right <= left:
        return 0.0
    width = min(max(1.0, a[2] - a[0]), max(1.0, b[2] - b[0]))
    return (right - left) / width


def vertical_overlap_ratio(
    a: tuple[float, float, float, float],
    b: tuple[float, float, float, float],
) -> float:
    top = max(a[1], b[1])
    bottom = min(a[3], b[3])
    if bottom <= top:
        return 0.0
    height = min(max(1.0, a[3] - a[1]), max(1.0, b[3] - b[1]))
    return (bottom - top) / height


def union_bbox(bboxes: Iterable[tuple[float, float, float, float]]) -> tuple[float, float, float, float]:
    boxes = list(bboxes)
    if not boxes:
        return (0.0, 0.0, 0.0, 0.0)
    return (
        min(b[0] for b in boxes),
        min(b[1] for b in boxes),
        max(b[2] for b in boxes),
        max(b[3] for b in boxes),
    )


def extract_pdf_figure_regions(page, text_dict, table_bboxes):
    """Preserve visible vector figures as well as embedded raster images.

    PDF rawdict type=1 contains *only* raster images. Line-art diagrams must
    come from drawing clusters. Thin fraction bars/underlines and verified
    table grids are not figures. Labels inside a figure remain in its crop,
    so callers must exclude these boxes from editable text extraction.
    """
    width, height = float(page.rect.width), float(page.rect.height)
    figures = []
    for block in text_dict.get("blocks", []):
        if block.get("type") == 1:
            box = tuple(float(v) for v in block["bbox"])
            area = max(0, box[2] - box[0]) * max(0, box[3] - box[1])
            if area >= 64 and area < width * height * 0.72:
                if not any(rect_overlap_ratio(box, tb) > 0.7 for tb in table_bboxes):
                    figures.append(box)
    drawings = [d for d in page.get_drawings()
                if d.get("type") != "clip" and (d.get("color") is not None or d.get("fill") is not None)
                # PDF producers often paint white rectangles behind every
                # text run. Clustering these would rasterize entire prose
                # sections and silently remove editable math.
                and not (d.get("color") is None and d.get("fill") is not None
                         and all(float(c) >= 0.97 for c in d["fill"]))]
    for rect in page.cluster_drawings(drawings=drawings):
        box = tuple(rect)
        if rect.width < 40 or rect.height < 40 or rect.get_area() >= width * height * 0.72:
            continue
        if any(rect_overlap_ratio(box, tb) > 0.7 for tb in table_bboxes):
            continue
        figures.append(box)

    # Join annotations with their photo, and side-by-side figures occupying
    # the same band. One source crop preserves their internal positioning.
    merged = []
    for box in figures:
        pending = box
        changed = True
        while changed:
            changed = False
            remaining = []
            for other in merged:
                gap = max(0, pending[0] - other[2], other[0] - pending[2])
                overlap = max(rect_overlap_ratio(pending, other), rect_overlap_ratio(other, pending))
                if overlap > 0.1 or (vertical_overlap_ratio(pending, other) > 0.5 and gap < width * 0.12):
                    pending = union_bbox([pending, other])
                    changed = True
                else:
                    remaining.append(other)
            merged = remaining
        merged.append(pending)

    result = []
    for box in merged:
        padded = (max(0, box[0] - 8), max(0, box[1] - 12), min(width, box[2] + 8), min(height, box[3] + 12))
        # Do not chop captions/labels at the boundary, but do not rasterize
        # unrelated prose merely because it is near a drawing.
        labels = []
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                text = "".join(c.get("c", "") for s in line.get("spans", []) for c in s.get("chars", []))
                lb = tuple(line["bbox"])
                if rect_overlap_ratio(lb, padded) > 0.3 and (
                    text.strip().startswith(("Hình", "Figure")) or len(text.strip()) <= 12
                ):
                    labels.append(lb)
        final = union_bbox([padded, *labels])
        if not any(rect_overlap_ratio(final, tb) > 0.2 for tb in table_bboxes):
            result.append(final)
    return result


def normalize_bbox_to_1000(
    bbox: tuple[float, float, float, float],
    page_width: float,
    page_height: float,
) -> list[float]:
    pw = max(1.0, float(page_width))
    ph = max(1.0, float(page_height))
    left = max(0.0, min(999.0, bbox[0] * 1000.0 / pw))
    top = max(0.0, min(999.0, bbox[1] * 1000.0 / ph))
    right = max(left + 1.0, min(1000.0, bbox[2] * 1000.0 / pw))
    bottom = max(top + 1.0, min(1000.0, bbox[3] * 1000.0 / ph))
    return [round(left, 2), round(top, 2), round(right, 2), round(bottom, 2)]


def bbox_1000_to_pixels(
    bbox_1000: list[float] | tuple[float, float, float, float],
    img_width: int,
    img_height: int,
    *,
    pad_px: int = 6,
) -> tuple[int, int, int, int]:
    left, top, right, bottom = bbox_1000
    x0 = max(0, int(left * img_width / 1000.0) - pad_px)
    y0 = max(0, int(top * img_height / 1000.0) - pad_px)
    x1 = min(img_width, max(x0 + 1, round(right * img_width / 1000.0) + pad_px))
    y1 = min(img_height, max(y0 + 1, round(bottom * img_height / 1000.0) + pad_px))
    return (x0, y0, x1, y1)


def _visible_drawing_segments(page) -> tuple[list[tuple[float, float, float, float]], list[tuple[float, float, float, float]]]:
    """Return (horizontal_segments, vertical_segments) from visible PDF vector drawings."""
    h_segs: list[tuple[float, float, float, float]] = []
    v_segs: list[tuple[float, float, float, float]] = []
    try:
        drawings = page.get_drawings()
    except Exception:
        return h_segs, v_segs

    for d in drawings:
        dtype = d.get("type")
        if dtype is None or dtype == "clip":
            continue
        if d.get("color") is None and d.get("fill") is None:
            continue
        added_from_items = False
        for item in d.get("items", ()):
            if not item:
                continue
            op = item[0]
            if op == "l" and len(item) >= 3:
                p1, p2 = item[1], item[2]
                x0, x1 = min(float(p1.x), float(p2.x)), max(float(p1.x), float(p2.x))
                y0, y1 = min(float(p1.y), float(p2.y)), max(float(p1.y), float(p2.y))
                if (x1 - x0) >= 2.5 and (y1 - y0) <= 2.2:
                    h_segs.append((x0, y0, x1, y1))
                    added_from_items = True
                elif (y1 - y0) >= 2.5 and (x1 - x0) <= 2.2:
                    v_segs.append((x0, y0, x1, y1))
                    added_from_items = True
            elif op == "re" and len(item) >= 2:
                r = item[1]
                rx0, ry0, rx1, ry1 = float(r.x0), float(r.y0), float(r.x1), float(r.y1)
                rw, rh = rx1 - rx0, ry1 - ry0
                if rw >= 2.5 and rh <= 3.6:
                    h_segs.append((rx0, ry0, rx1, ry1))
                    added_from_items = True
                elif rh >= 2.5 and rw <= 3.6:
                    v_segs.append((rx0, ry0, rx1, ry1))
                    added_from_items = True
                elif rw >= 12.0 and rh >= 8.0 and d.get("color") is not None:
                    h_segs.append((rx0, ry0, rx1, ry0 + 1.0))
                    h_segs.append((rx0, ry1 - 1.0, rx1, ry1))
                    v_segs.append((rx0, ry0, rx0 + 1.0, ry1))
                    v_segs.append((rx1 - 1.0, ry0, rx1, ry1))
                    added_from_items = True

        if not added_from_items:
            rect = d.get("rect")
            if rect is not None:
                x0, y0, x1, y1 = float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1)
                w = x1 - x0
                h = y1 - y0
                if w >= 2.5 and h <= 3.6:
                    h_segs.append((x0, y0, x1, y1))
                elif h >= 2.5 and w <= 3.6:
                    v_segs.append((x0, y0, x1, y1))
                elif w >= 12.0 and h >= 8.0 and d.get("color") is not None:
                    h_segs.append((x0, y0, x1, y0 + 1.0))
                    h_segs.append((x0, y1 - 1.0, x1, y1))
                    v_segs.append((x0, y0, x0 + 1.0, y1))
                    v_segs.append((x1 - 1.0, y0, x1, y1))

    return h_segs, v_segs


def _cluster_positions(values: list[float], tolerance: float = 3.0) -> int:
    if not values:
        return 0
    sorted_vals = sorted(values)
    clusters = 1
    anchor = sorted_vals[0]
    for v in sorted_vals[1:]:
        if abs(v - anchor) > tolerance:
            clusters += 1
            anchor = v
    return clusters


def extract_verified_pdf_tables(page) -> list[dict[str, Any]]:
    """Extract PDF tables while rejecting false-positive tables caused by fraction bars or clip boxes."""
    if not hasattr(page, "find_tables"):
        return []
    try:
        with python_warnings.catch_warnings():
            python_warnings.simplefilter("ignore")
            finder = page.find_tables()
    except Exception:
        return []

    h_segs, v_segs = _visible_drawing_segments(page)
    verified: list[dict[str, Any]] = []
    for table in getattr(finder, "tables", []):
        try:
            matrix = table.extract()
        except Exception:
            continue
        if not matrix:
            continue
        non_empty_rows = [row for row in matrix if any(clean_cell_text(cell) for cell in row)]
        col_count = max((len(row) for row in matrix), default=0)
        if len(non_empty_rows) < 2 or col_count < 2:
            continue
        tbbox = tuple(float(v) for v in table.bbox)
        tw = tbbox[2] - tbbox[0]
        th = tbbox[3] - tbbox[1]
        if tw < 75.0 or th < 22.0:
            continue

        # Check visible horizontal and vertical border lines inside/around tbbox.
        # Note: table borders may be drawn as per-cell segments, so we group segments by y-level / x-level
        # and sum their horizontal / vertical span within the table bbox.
        h_by_y: list[tuple[float, float]] = []
        for x0, y0, x1, y1 in h_segs:
            if y1 < tbbox[1] - 4.0 or y0 > tbbox[3] + 4.0:
                continue
            ox0 = max(x0, tbbox[0] - 4.0)
            ox1 = min(x1, tbbox[2] + 4.0)
            if ox1 - ox0 >= 12.0:
                h_by_y.append(((y0 + y1) * 0.5, ox1 - ox0))

        v_by_x: list[tuple[float, float]] = []
        for x0, y0, x1, y1 in v_segs:
            if x1 < tbbox[0] - 4.0 or x0 > tbbox[2] + 4.0:
                continue
            oy0 = max(y0, tbbox[1] - 4.0)
            oy1 = min(y1, tbbox[3] + 4.0)
            if oy1 - oy0 >= 10.0:
                v_by_x.append(((x0 + x1) * 0.5, oy1 - oy0))

        # Aggregate span coverage per y-cluster and x-cluster.
        h_levels: list[float] = []
        for y_val, span_w in sorted(h_by_y, key=lambda item: item[0]):
            if not h_levels or abs(y_val - h_levels[-1]) > 3.5:
                h_levels.append(y_val)
        v_levels: list[float] = []
        for x_val, span_h in sorted(v_by_x, key=lambda item: item[0]):
            if not v_levels or abs(x_val - v_levels[-1]) > 3.5:
                v_levels.append(x_val)

        # Require both horizontal and vertical border structure (at least 2 horizontal levels and 2 vertical levels).
        if len(h_levels) < 2 or len(v_levels) < 2:
            continue

        verified.append({"kind": "table", "bbox": tbbox, "rows": matrix})

    return verified


def extract_candidate_horizontal_bars(
    page,
    excluded_table_bboxes: Iterable[tuple[float, float, float, float]] = (),
) -> list[HorizontalBar]:
    """Extract thin horizontal segments from page drawings that are outside verified tables."""
    excluded = list(excluded_table_bboxes)
    page_width = float(page.rect.width)
    h_segs, _ = _visible_drawing_segments(page)

    filtered: list[tuple[float, float, float, float]] = []
    for seg in h_segs:
        w = seg[2] - seg[0]
        h = seg[3] - seg[1]
        if w < 2.5 or w > page_width * 0.78 or h > 3.6:
            continue
        if any(rect_overlap_ratio(seg, tb) > 0.25 for tb in excluded):
            continue
        filtered.append(seg)

    if not filtered:
        return []

    # Sort by y_mid then x0 and merge collinear overlapping/adjacent segments.
    filtered.sort(key=lambda s: (round((s[1] + s[3]) * 0.5, 1), s[0]))
    merged: list[HorizontalBar] = []
    for seg in filtered:
        bar = HorizontalBar(seg)
        merged_into_existing = False
        for idx in range(len(merged) - 1, -1, -1):
            prev = merged[idx]
            if abs(bar.y_mid - prev.y_mid) <= 1.2 and bar.x0 <= prev.x1 + 1.5 and bar.x1 >= prev.x0 - 1.5:
                merged[idx] = HorizontalBar(
                    (
                        min(prev.x0, bar.x0),
                        min(prev.y0, bar.y0),
                        max(prev.x1, bar.x1),
                        max(prev.y1, bar.y1),
                    )
                )
                merged_into_existing = True
                break
        if not merged_into_existing:
            merged.append(bar)
    return merged


def classify_fraction_bars(
    candidate_bars: Iterable[HorizontalBar],
    char_boxes: Iterable[dict[str, Any]],
) -> list[HorizontalBar]:
    """Filter horizontal bars to keep only true fraction bars (must have numerator above and denominator below)."""
    chars = [c for c in char_boxes if str(c.get("c", "")).strip()]
    fraction_bars: list[HorizontalBar] = []
    for bar in candidate_bars:
        # Underlines under headings are typically thick (>= 1.5pt) or wide (> 160pt).
        if bar.height >= 1.55 and bar.width >= 45.0:
            continue
        y_mid = bar.y_mid
        x_pad = max(1.5, min(3.5, bar.width * 0.10))
        above_chars = []
        below_chars = []
        for ch in chars:
            cb = ch["bbox"]
            cx_mid = (cb[0] + cb[2]) * 0.5
            if cx_mid < bar.x0 - x_pad or cx_mid > bar.x1 + x_pad:
                continue
            # Use baseline origin_y when available so tall CambriaMath ascenders do not skew glyph center.
            size = float(ch.get("size", 10.5) or 10.5)
            origin_y = ch.get("origin_y")
            vy_mid = (float(origin_y) - size * 0.32) if origin_y is not None else ((cb[1] + cb[3]) * 0.5)
            if 0.8 <= (y_mid - vy_mid) <= 14.5:
                above_chars.append(ch)
            elif 0.8 <= (vy_mid - y_mid) <= 14.5:
                below_chars.append(ch)
        if not above_chars or not below_chars:
            continue
        # Reject heading underlines where the bar is thick (>= 1.45pt) and the text above is Vietnamese prose.
        above_str = "".join(str(c.get("c", "")) for c in sorted(above_chars, key=lambda item: item["bbox"][0]))
        if bar.height >= 1.45 and bar.width >= 45.0 and any(ord(ch) > 127 and ch.isalpha() and not (0x0370 <= ord(ch) <= 0x03FF or 0x1D400 <= ord(ch) <= 0x1D7FF) for ch in above_str):
            if len(re.findall(r"[a-zA-ZÀ-ỹ]{2,}", above_str)) >= 1:
                continue
        fraction_bars.append(bar)
    return fraction_bars


# Heuristics for scan/image formula region detection
_DATE_OR_CODE_RE = re.compile(
    r"^\s*(?:\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?|\d+(?:[.,]\d+)?\s*(?:đ|vnd|usd|%|h|phút|giây|mm|cm|m|km|kg|kw))\s*$",
    re.IGNORECASE,
)
_EXPLICIT_MATH_TOKEN_RE = re.compile(
    r"(?:\\(?:frac|sqrt|sum|int|prod|lim|sin|cos|tan|cot|ln|log|alpha|beta|gamma|delta|omega|pi|varphi|rho|eta|le|ge|neq|approx|pm|times|cdot)"
    r"|[∫∬∭∮∑∏√∞≠≤≥≈≡±∓∂∇∈∉⊂⊆⊃⊇∪∩⇒⇔α-ωΑ-Ω]"
    r"|\b(?:sin|cos|tan|cot|arcsin|arccos|arctan|ln|log|lim|max|min|sup|inf|det)\s*[\(\[a-zA-Z0-9]"
    r"|[a-zA-Z0-9\)\]]\s*[\^_]\s*[\{\(a-zA-Z0-9\-\+])"
)
_EQUATION_REL_RE = re.compile(r"[=≠≤≥≈≡]|(?<![<>])(?:<=|>=|!=)(?![<>])")
_PROSE_WORD_RE = re.compile(
    r"\b(?:như|trong|được|không|của|những|người|trường|hợp|chương|phần|bảng|hình|trang|ngày|tháng|năm|theo|với|trên|dưới|khi|nếu|thì|hoặc|cần|phải|tính|chọn|kiểm|nghiệm|thiết|kế|động|cơ|trục|bánh|răng|xích|đai)\b",
    re.IGNORECASE,
)


def is_likely_formula_text(text: str) -> bool:
    """Return True if an OCR/text string looks like a mathematical formula rather than prose/date/currency."""
    cleaned = (text or "").strip()
    if not cleaned or len(cleaned) < 2:
        return False
    if _DATE_OR_CODE_RE.match(cleaned):
        return False
    if _EXPLICIT_MATH_TOKEN_RE.search(cleaned):
        return True
    if _EQUATION_REL_RE.search(cleaned):
        prose_words = len(_PROSE_WORD_RE.findall(cleaned))
        if prose_words <= 2 and len(cleaned) <= 140:
            return True
    # Short algebraic expression with operators like a/b + c/d or x^2 + y^2
    if re.search(r"\b[a-zA-Z]\s*[\+\-\*/=]\s*[a-zA-Z0-9]", cleaned):
        prose_words = len(_PROSE_WORD_RE.findall(cleaned))
        if prose_words == 0 and len(cleaned) <= 80:
            return True
    return False


def is_single_formula_image(image: Image.Image) -> bool:
    """Detect whether an input image is already a cropped formula snippet rather than a full document page."""
    w, h = image.size
    if w <= 0 or h <= 0:
        return False
    aspect = w / max(1.0, float(h))
    # Typical formula snippets are wider than tall or relatively compact in height.
    if h <= 420 and aspect >= 1.35:
        return True
    if h <= 260 and w <= 1800:
        return True
    return False


def detect_visual_fraction_regions(
    image: Image.Image,
    excluded_bboxes: Iterable[tuple[int, int, int, int]] = (),
) -> list[tuple[int, int, int, int]]:
    """Detect bounding boxes of stacked fraction formulas in a raster image using OpenCV morphology."""
    gray = np.array(image.convert("L"))
    if gray.size == 0:
        return []
    img_h, img_w = gray.shape[:2]
    excluded = [tuple(float(v) for v in b) for b in excluded_bboxes]
    # Binarize (ink = 255, background = 0)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    if np.count_nonzero(binary) == 0:
        return []

    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(binary, connectivity=8)
    if num_labels <= 1:
        return []

    bars: list[tuple[int, int, int, int]] = []
    glyphs: list[tuple[int, int, int, int]] = []
    for idx in range(1, num_labels):
        x, y, w, h, area = (
            int(stats[idx, cv2.CC_STAT_LEFT]),
            int(stats[idx, cv2.CC_STAT_TOP]),
            int(stats[idx, cv2.CC_STAT_WIDTH]),
            int(stats[idx, cv2.CC_STAT_HEIGHT]),
            int(stats[idx, cv2.CC_STAT_AREA]),
        )
        if area < 4:
            continue
        box = (x, y, x + w, y + h)
        if excluded and any(rect_overlap_ratio(box, tb) > 0.25 for tb in excluded):
            continue
        if w >= 14 and w <= int(img_w * 0.72) and h <= max(6, int(w * 0.14)):
            bars.append(box)
        elif h >= 4 and w >= 2 and h <= int(img_h * 0.22):
            glyphs.append(box)

    if not bars or not glyphs:
        return []

    max_near_gap = max(16, min(38, int(img_h * 0.024)))
    max_far_dist = max(34, min(78, int(img_h * 0.052)))
    formula_boxes: list[tuple[int, int, int, int]] = []
    for bx0, by0, bx1, by1 in bars:
        bw = bx1 - bx0
        bh = by1 - by0
        if bh >= 5 and bw >= max(110, int(img_w * 0.16)):
            # Likely a thick section heading underline rather than a fraction bar
            continue
        by_mid = (by0 + by1) * 0.5
        num_boxes = [
            g for g in glyphs
            if (g[0] + g[2]) * 0.5 >= bx0 - 6
            and (g[0] + g[2]) * 0.5 <= bx1 + 6
            and g[3] <= by1 + 3
            and (by_mid - g[3]) <= max_near_gap
            and (by_mid - g[1]) <= max_far_dist
        ]
        den_boxes = [
            g for g in glyphs
            if (g[0] + g[2]) * 0.5 >= bx0 - 6
            and (g[0] + g[2]) * 0.5 <= bx1 + 6
            and g[1] >= by0 - 3
            and (g[1] - by_mid) <= max_near_gap
            and (g[3] - by_mid) <= max_far_dist
        ]
        if not num_boxes or not den_boxes:
            continue
        # Ensure numerator and denominator are reasonably close to the bar
        min_up_gap = min(max(0.0, by0 - g[3]) for g in num_boxes)
        min_dn_gap = min(max(0.0, g[1] - by1) for g in den_boxes)
        if min_up_gap > max(14, int(img_h * 0.016)) or min_dn_gap > max(16, int(img_h * 0.018)):
            continue

        cluster = [(bx0, by0, bx1, by1)] + num_boxes + den_boxes
        fx0 = min(c[0] for c in cluster)
        fy0 = min(c[1] for c in cluster)
        fx1 = max(c[2] for c in cluster)
        fy1 = max(c[3] for c in cluster)
        fh = max(14, fy1 - fy0)

        # Expand horizontally along the same equation line (around by_mid) to capture LHS/RHS terms (e.g., `v = `).
        changed = True
        while changed:
            changed = False
            for gx0, gy0, gx1, gy1 in glyphs:
                if gx0 >= fx0 and gx1 <= fx1 and gy0 >= fy0 and gy1 <= fy1:
                    continue
                gy_mid = (gy0 + gy1) * 0.5
                if abs(gy_mid - by_mid) > max(14.0, fh * 0.42):
                    continue
                if gy0 < fy0 - 10 or gy1 > fy1 + 10:
                    continue
                horiz_gap = max(0, fx0 - gx1, gx0 - fx1)
                if horiz_gap <= max(24, int(fh * 0.65)):
                    fx0 = min(fx0, gx0)
                    fy0 = min(fy0, gy0)
                    fx1 = max(fx1, gx1)
                    fy1 = max(fy1, gy1)
                    changed = True

        pad_x = min(12, max(4, int(fh * 0.16)))
        pad_y = min(10, max(4, int(fh * 0.14)))
        formula_boxes.append(
            (
                max(0, fx0 - pad_x),
                max(0, fy0 - pad_y),
                min(img_w, fx1 + pad_x),
                min(img_h, fy1 + pad_y),
            )
        )

    # Merge overlapping formula boxes
    if not formula_boxes:
        return []
    formula_boxes.sort(key=lambda b: (b[1], b[0]))
    merged: list[tuple[int, int, int, int]] = []
    for box in formula_boxes:
        if merged and (
            rect_overlap_ratio(box, merged[-1]) > 0.15
            or (
                vertical_overlap_ratio(box, merged[-1]) > 0.45
                and max(0, box[0] - merged[-1][2], merged[-1][0] - box[2]) <= 24
            )
        ):
            prev = merged[-1]
            merged[-1] = (
                min(prev[0], box[0]),
                min(prev[1], box[1]),
                max(prev[2], box[2]),
                max(prev[3], box[3]),
            )
        else:
            merged.append(box)
    return merged


_VN_PROSE_PATTERN = re.compile(
    r"[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ"
    r"ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐa-zA-Z]{3,}"
)
_FRAGMENT_SYMBOLS_PATTERN = re.compile(
    r"^[\d\s+\-*/=><^_{}(),.·⁰¹²³⁴⁵⁶⁷⁸⁹\u2264\u2265\u221a\u03c0\u03b1\u03b2\u03b7\u03c3\u03c4\u03c9[\]a-zA-Z]{1,10}$"
)


def is_fragment_text(text: str) -> bool:
    """Check if a string looks like a broken formula fragment.

    E.g.: standalone digits ('5', '15000', '715'), dropped powers/indices ('3', '2'),
    or short operator/variable fragments ('TI', '0,2[τ]', 'Wj =', 'π dj').
    """
    s = text.strip()
    if not s:
        return False
    vn_words = _VN_PROSE_PATTERN.findall(s)
    if len(vn_words) >= 2:
        return False
    if re.fullmatch(r"^[\d\s.,]{1,12}$", s):
        return True
    if _FRAGMENT_SYMBOLS_PATTERN.fullmatch(s):
        return True
    return False
