"""Rule-based digital PDF mathematical expression parser.

Reconstructs editable Word equations (`m:oMath` / `m:oMathPara`) directly from
PDF character coordinates, baselines, font metadata, and vector drawings
(horizontal fraction bars and radical vinculums) without calling any ML model
or external API.
"""
from __future__ import annotations

import re
import statistics
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Literal

from .layout_pdf import repair_raw_text_blocks
from .math_pdf import _Block, _Cell, _Page, _Row, _Span, latex_to_omml
from .math_regions import (
    HorizontalBar,
    MathRegion,
    classify_fraction_bars,
    extract_candidate_horizontal_bars,
    extract_pdf_figure_regions,
    extract_verified_pdf_tables,
    horizontal_overlap_ratio,
    is_likely_formula_text,
    normalize_bbox_to_1000,
    rect_overlap_ratio,
    union_bbox,
    vertical_overlap_ratio,
)

# Adobe Symbol / Wingdings Private Use Area (U+F020..U+F0FF) mapping to LaTeX/Unicode
_SYMBOL_PUA_MAP: dict[str, str] = {
    "\uf020": " ",
    "\uf02b": "+",
    "\uf02d": "-",
    "\uf03d": "=",
    "\uf03c": "<",
    "\uf03e": ">",
    "\uf028": "(",
    "\uf029": ")",
    "\uf05b": "[",
    "\uf05d": "]",
    "\uf07b": r"\{",
    "\uf07d": r"\}",
    # Uppercase Greek in Symbol PUA
    "\uf044": r"\Delta",
    "\uf046": r"\Phi",
    "\uf047": r"\Gamma",
    "\uf04c": r"\Lambda",
    "\uf050": r"\Pi",
    "\uf051": r"\Theta",
    "\uf053": r"\Sigma",
    "\uf055": r"\Upsilon",
    "\uf057": r"\Omega",
    "\uf058": r"\Xi",
    "\uf059": r"\Psi",
    # Lowercase Greek in Symbol PUA
    "\uf061": r"\alpha",
    "\uf062": r"\beta",
    "\uf063": r"\chi",
    "\uf064": r"\delta",
    "\uf065": r"\varepsilon",
    "\uf066": r"\phi",
    "\uf067": r"\gamma",
    "\uf068": r"\eta",
    "\uf069": r"\iota",
    "\uf06a": r"\varphi",
    "\uf06b": r"\kappa",
    "\uf06c": r"\lambda",
    "\uf06d": r"\mu",
    "\uf06e": r"\nu",
    "\uf070": r"\pi",
    "\uf071": r"\theta",
    "\uf072": r"\rho",
    "\uf073": r"\sigma",
    "\uf074": r"\tau",
    "\uf075": r"\upsilon",
    "\uf076": r"\varpi",
    "\uf077": r"\omega",
    "\uf078": r"\xi",
    "\uf079": r"\psi",
    "\uf07a": r"\zeta",
    "\uf04a": r"\vartheta",
    "\uf056": r"\varsigma",
    # Math operators in Symbol PUA
    "\uf0a3": r"\le",
    "\uf0a5": r"\infty",
    "\uf0b0": r"^{\circ}",
    "\uf0b1": r"\pm",
    "\uf0b3": r"\ge",
    "\uf0b4": r"\times",
    "\uf0b5": r"\propto",
    "\uf0b6": r"\partial",
    "\uf0b7": "•",
    "\uf0b8": r"\div",
    "\uf0b9": r"\neq",
    "\uf0ba": r"\equiv",
    "\uf0bb": r"\approx",
    "\uf0c4": r"\otimes",
    "\uf0c5": r"\oplus",
    "\uf0c6": r"\emptyset",
    "\uf0c7": r"\cap",
    "\uf0c8": r"\cup",
    "\uf0cc": r"\subset",
    "\uf0cd": r"\subseteq",
    "\uf0ce": r"\in",
    "\uf0cf": r"\notin",
    "\uf0d1": r"\nabla",
    "\uf0d5": r"\prod",
    "\uf0d6": r"\sqrt",
    "\uf0d7": r"\cdot",
    "\uf0db": r"\Leftrightarrow",
    "\uf0dc": r"\Leftarrow",
    "\uf0de": r"\Rightarrow",
    "\uf0ac": r"\leftarrow",
    "\uf0ae": r"\rightarrow",
    "\uf0ab": r"\leftrightarrow",
    "\uf0e5": r"\sum",
    "\uf0f2": r"\int",
}

_UNICODE_MATH_MAP: dict[str, str] = {
    "α": r"\alpha",
    "β": r"\beta",
    "γ": r"\gamma",
    "δ": r"\delta",
    "ε": r"\varepsilon",
    "ϵ": r"\epsilon",
    "ζ": r"\zeta",
    "η": r"\eta",
    "θ": r"\theta",
    "ϑ": r"\vartheta",
    "ι": r"\iota",
    "κ": r"\kappa",
    "λ": r"\lambda",
    "μ": r"\mu",
    "µ": r"\mu",
    "ν": r"\nu",
    "ξ": r"\xi",
    "π": r"\pi",
    "ϖ": r"\varpi",
    "ρ": r"\rho",
    "ϱ": r"\varrho",
    "σ": r"\sigma",
    "ς": r"\varsigma",
    "τ": r"\tau",
    "υ": r"\upsilon",
    "φ": r"\varphi",
    "ϕ": r"\phi",
    "χ": r"\chi",
    "ψ": r"\psi",
    "ω": r"\omega",
    "Γ": r"\Gamma",
    "Δ": r"\Delta",
    "∆": r"\Delta",
    "Θ": r"\Theta",
    "Λ": r"\Lambda",
    "Ξ": r"\Xi",
    "Π": r"\Pi",
    "Σ": r"\Sigma",
    "Υ": r"\Upsilon",
    "Φ": r"\Phi",
    "Ψ": r"\Psi",
    "Ω": r"\Omega",
    "Ω": r"\Omega",
    "∑": r"\sum",
    "∏": r"\prod",
    "∐": r"\coprod",
    "∫": r"\int",
    "∬": r"\iint",
    "∭": r"\iiint",
    "∮": r"\oint",
    "√": r"\sqrt",
    "∛": r"\sqrt[3]",
    "∜": r"\sqrt[4]",
    "∞": r"\infty",
    "≠": r"\neq",
    "≤": r"\le",
    "⩽": r"\le",
    "≥": r"\ge",
    "⩾": r"\ge",
    "≈": r"\approx",
    "≡": r"\equiv",
    "±": r"\pm",
    "∓": r"\mp",
    "×": r"\times",
    "÷": r"\div",
    "⋅": r"\cdot",
    "·": r"\cdot",
    "∙": r"\cdot",
    "∂": r"\partial",
    "∇": r"\nabla",
    "∈": r"\in",
    "∉": r"\notin",
    "⊂": r"\subset",
    "⊆": r"\subseteq",
    "⊃": r"\supset",
    "⊇": r"\supseteq",
    "∪": r"\cup",
    "∩": r"\cap",
    "∅": r"\emptyset",
    "∀": r"\forall",
    "∃": r"\exists",
    "⇒": r"\Rightarrow",
    "⇐": r"\Leftarrow",
    "⇔": r"\Leftrightarrow",
    "→": r"\to",
    "←": r"\leftarrow",
    "↔": r"\leftrightarrow",
    "↦": r"\mapsto",
    "∝": r"\propto",
    "−": "-",
    "–": "-",
    "°": r"^{\circ}",
}

_MATH_FONT_KEYWORDS = (
    "cambriamath",
    "cmmi",
    "cmsy",
    "cmex",
    "msam",
    "msbm",
    "latinmodern-math",
    "stixmath",
    "asana-math",
    "xits-math",
    "mtmi",
    "mtsy",
    "mtex",
    "euclidmath",
    "symbol",
)

_NARY_LATEX_TOKENS = frozenset({
    r"\sum",
    r"\prod",
    r"\coprod",
    r"\int",
    r"\iint",
    r"\iiint",
    r"\oint",
    r"\lim",
    r"\max",
    r"\min",
    r"\sup",
    r"\inf",
})

_FUNCTION_WORDS = frozenset({
    "sin", "cos", "tan", "cot", "sec", "csc",
    "arcsin", "arccos", "arctan",
    "sinh", "cosh", "tanh", "coth",
    "ln", "log", "lg", "exp",
    "lim", "max", "min", "sup", "inf", "det", "gcd",
})

_VIETNAMESE_PROSE_WORDS = frozenset({
    "và", "hoặc", "của", "trong", "trên", "dưới", "theo", "với", "cho", "từ",
    "đến", "khi", "nếu", "thì", "được", "không", "những", "các", "một", "hai",
    "ba", "bốn", "năm", "bảng", "hình", "chương", "phần", "trang", "tính", "chọn",
    "kiểm", "nghiệm", "thiết", "kế", "công", "suất", "vận", "tốc", "trục", "bánh",
    "răng", "động", "cơ", "bồn", "trộn", "lực", "ma", "sát", "đường", "kính",
    "bán", "tải", "trọng", "thời", "gian", "quay", "vòng", "phút", "giây", "MỤC",
})


class UnsupportedMathStructureError(Exception):
    def __init__(self, reason: str, bbox: tuple[float, float, float, float]):
        super().__init__(reason)
        self.reason = reason
        self.bbox = bbox


@dataclass(slots=True)
class CharToken:
    c: str
    norm_c: str
    latex_token: str
    bbox: tuple[float, float, float, float]
    origin_x: float
    origin_y: float
    size: float
    font: str
    bold: bool = False
    italic: bool = False
    is_math_font: bool = False
    is_math_symbol: bool = False

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
    def cx_mid(self) -> float:
        return (self.bbox[0] + self.bbox[2]) * 0.5

    @property
    def cy_mid(self) -> float:
        return self.origin_y - self.size * 0.32


@dataclass(slots=True)
class MathAtom:
    latex: str
    bbox: tuple[float, float, float, float]
    origin_y: float
    size: float
    is_fraction: bool = False
    is_nary: bool = False
    is_radical: bool = False
    sub_latex: str | None = None
    sup_latex: str | None = None

    def render(self) -> str:
        base = self.latex
        if self.sub_latex is not None and self.sup_latex is not None:
            return f"{base}_{{{self.sub_latex}}}^{{{self.sup_latex}}}"
        if self.sub_latex is not None:
            return f"{base}_{{{self.sub_latex}}}"
        if self.sup_latex is not None:
            return f"{base}^{{{self.sup_latex}}}"
        return base


@dataclass(slots=True)
class DigitalMathPageResult:
    has_math: bool
    page_model: _Page
    recognized_by_rules: int = 0
    unresolved_regions: list[MathRegion] = field(default_factory=list)


def _is_math_alphanumeric(ch: str) -> bool:
    if not ch:
        return False
    cp = ord(ch[0])
    return 0x1D400 <= cp <= 0x1D7FF or 0x2100 <= cp <= 0x214F


_LATEX_TO_UNICODE_MAP: dict[str, str] = {v: k for k, v in _UNICODE_MATH_MAP.items()}
_LATEX_TO_UNICODE_MAP.update({
    r"\{": "{",
    r"\}": "}",
    r"^{\circ}": "°",
    r"\leftarrow": "←",
    r"\rightarrow": "→",
    r"\leftrightarrow": "↔",
})


def _normalize_char(raw_c: str, font_name: str) -> tuple[str, str, bool, bool]:
    """Return (norm_c, latex_token, is_math_font, is_math_symbol)."""
    lowered_font = (font_name or "").lower()
    is_math_font = any(k in lowered_font for k in _MATH_FONT_KEYWORDS)

    if raw_c in _SYMBOL_PUA_MAP:
        mapped = _SYMBOL_PUA_MAP[raw_c]
        if mapped == "•":
            return mapped, mapped, False, False
        unicode_char = _LATEX_TO_UNICODE_MAP.get(mapped, mapped)
        return unicode_char, mapped, True, True

    if raw_c == "\ufffd" or (raw_c and 0xE000 <= ord(raw_c[0]) <= 0xF8FF):
        # Preserve math-font flag from known math fonts, but do not treat an
        # isolated unknown/PUA glyph in ordinary prose as a math indicator.
        return "\ufffd", "\ufffd", is_math_font, False

    if _is_math_alphanumeric(raw_c):
        nfkc = unicodedata.normalize("NFKC", raw_c)
        if nfkc in _UNICODE_MATH_MAP:
            return nfkc, _UNICODE_MATH_MAP[nfkc], True, True
        return nfkc, nfkc, True, False

    nfkc = unicodedata.normalize("NFKC", raw_c)
    if nfkc in {"•", "◦", "▪", "▫", "\u2022", "\uf0b7"}:
        return "•", "•", False, False
    if nfkc in _UNICODE_MATH_MAP:
        mapped = _UNICODE_MATH_MAP[nfkc]
        is_sym = nfkc not in {"−", "–"}
        return ("-" if nfkc in {"−", "–"} else nfkc), mapped, is_math_font or is_sym, is_sym

    if nfkc == "%":
        return "%", r"\%", True, False

    if not nfkc.isalnum() and nfkc not in {"=", "+", "-", "(", ")", "[", "]", "{", "}", "|", "/", "<", ">", "%"}:
        is_math_font = False
    return nfkc, nfkc, is_math_font, False


def _extract_char_tokens(
    text_dict: dict[str, Any],
    excluded_bboxes: list[tuple[float, float, float, float]] | None = None,
    clip_bbox: tuple[float, float, float, float] | None = None,
) -> list[CharToken]:
    excluded = excluded_bboxes or []
    tokens: list[CharToken] = []
    for block in text_dict.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            direction = tuple(line.get("dir", (1.0, 0.0)))
            if len(direction) >= 2 and (abs(float(direction[0]) - 1.0) > 0.08 or abs(float(direction[1])) > 0.08):
                continue
            for span in line.get("spans", []):
                font_name = str(span.get("font", ""))
                size = float(span.get("size", 11.0) or 11.0)
                flags = int(span.get("flags", 0))
                bold = bool(flags & 16)
                italic = bool(flags & 2)
                span_origin = tuple(span.get("origin", (0.0, 0.0)))
                for ch_dict in span.get("chars", []):
                    raw_c = str(ch_dict.get("c", "")).replace("\x00", "")
                    if not raw_c:
                        continue
                    cbbox = tuple(float(v) for v in ch_dict.get("bbox", span.get("bbox", (0, 0, 0, 0))))
                    if clip_bbox is not None and rect_overlap_ratio(cbbox, clip_bbox) < 0.45:
                        continue
                    if any(rect_overlap_ratio(cbbox, ex) > 0.35 for ex in excluded):
                        continue
                    origin = tuple(ch_dict.get("origin", span_origin))
                    ox = float(origin[0]) if len(origin) >= 1 else cbbox[0]
                    oy = float(origin[1]) if len(origin) >= 2 else cbbox[3]
                    # Clamp inflated font ascender/descender boxes (e.g. CambriaMath) to visual glyph bounds
                    if (cbbox[3] - cbbox[1]) > size * 1.30:
                        vy0 = max(cbbox[1], oy - size * 0.88)
                        vy1 = min(cbbox[3], oy + size * 0.28)
                        if vy1 > vy0 + 1.0:
                            cbbox = (cbbox[0], vy0, cbbox[2], vy1)
                    norm_c, latex_token, is_mfont, is_msym = _normalize_char(raw_c, font_name)
                    tokens.append(
                        CharToken(
                            c=raw_c,
                            norm_c=norm_c,
                            latex_token=latex_token,
                            bbox=cbbox,
                            origin_x=ox,
                            origin_y=oy,
                            size=size,
                            font=font_name,
                            bold=bold,
                            italic=italic,
                            is_math_font=is_mfont,
                            is_math_symbol=is_msym,
                        )
                    )
    return tokens


def _format_script_content(latex_str: str) -> str:
    """Wrap multi-letter or Vietnamese script text in \\mathrm{...} when it is not already a LaTeX command."""
    s = latex_str.strip()
    if not s:
        return s
    if "\\" in s or "{" in s or "^" in s or "_" in s:
        return s
    # If script contains Vietnamese diacritics or 2+ letters (e.g. đc, ct, lv, sb, bx, br, ol, tv, xm, ms)
    letters = [c for c in s if c.isalpha()]
    has_non_ascii = any(ord(c) > 127 for c in letters)
    if has_non_ascii or len(letters) >= 2:
        return f"\\mathrm{{{s}}}"
    return s


def _tokens_to_flat_latex(tokens: list[CharToken]) -> str:
    """Convert a simple sequence of CharTokens (e.g. inside a script) to a LaTeX string."""
    if not tokens:
        return ""
    ordered = sorted(tokens, key=lambda t: t.x0)
    parts: list[str] = []
    prev: CharToken | None = None
    for tok in ordered:
        if tok.norm_c.isspace():
            continue
        if prev is not None:
            gap = tok.x0 - prev.x1
            if parts and parts[-1].startswith("\\") and parts[-1][-1].isalpha() and tok.latex_token[:1].isalpha():
                parts.append(" ")
            elif gap > max(2.2, tok.size * 0.35) and not parts[-1].endswith(" "):
                parts.append(" ")
        parts.append(tok.latex_token)
        prev = tok
    raw = "".join(parts).strip()
    # Replace known function words with \sin, \cos, etc.
    for fn in _FUNCTION_WORDS:
        raw = re.sub(rf"(?<!\\)\b{fn}\b", rf"\\{fn}", raw)
    return raw


def _parse_math_box(
    chars: list[CharToken],
    fraction_bars: list[HorizontalBar],
    radical_bars: list[HorizontalBar] | None = None,
    *,
    depth: int = 0,
) -> str:
    """Recursively convert a 2D box of CharTokens and HorizontalBars into structured LaTeX."""
    non_space = [c for c in chars if not c.norm_c.isspace()]
    if not non_space and not fraction_bars:
        return ""
    if depth > 6:
        raise UnsupportedMathStructureError(
            "deep_nesting",
            union_bbox([c.bbox for c in non_space] + [b.bbox for b in fraction_bars]),
        )
    if any(c.norm_c == "\ufffd" for c in non_space):
        raise UnsupportedMathStructureError(
            "corrupted_glyph_encoding",
            union_bbox([c.bbox for c in non_space] + [b.bbox for b in fraction_bars]),
        )
    if depth == 0:
        open_delims = sum(1 for c in non_space if c.norm_c in "([")
        close_delims = sum(1 for c in non_space if c.norm_c in ")]")
        if open_delims != close_delims:
            raise UnsupportedMathStructureError(
                "unbalanced_delimiters",
                union_bbox([c.bbox for c in non_space] + [b.bbox for b in fraction_bars]),
            )

    rad_bars = list(radical_bars or [])

    # 1. Identify top-level fraction bars (bars not nested inside another bar's numerator or denominator)
    top_bars: list[HorizontalBar] = []
    for i, bar in enumerate(fraction_bars):
        is_nested = False
        for j, other in enumerate(fraction_bars):
            if i == j:
                continue
            if (
                other.width >= bar.width + 6.0
                and bar.x0 >= other.x0 - 2.5
                and bar.x1 <= other.x1 + 2.5
                and 1.5 < abs(bar.y_mid - other.y_mid) <= 14.5
            ):
                is_nested = True
                break
        if not is_nested:
            top_bars.append(bar)
    top_bars.sort(key=lambda b: b.x0)

    consumed_char_ids: set[int] = set()
    consumed_bar_ids: set[int] = set()
    fraction_atoms: list[MathAtom] = []

    for bar in top_bars:
        consumed_bar_ids.add(id(bar))
        x_pad = max(1.5, min(3.2, bar.width * 0.08))

        # Nested bars belonging to numerator vs denominator of `bar`
        num_bars: list[HorizontalBar] = []
        den_bars: list[HorizontalBar] = []
        for other in fraction_bars:
            if id(other) == id(bar) or id(other) in consumed_bar_ids:
                continue
            ox_mid = (other.x0 + other.x1) * 0.5
            if bar.x0 - x_pad <= ox_mid <= bar.x1 + x_pad and abs(other.y_mid - bar.y_mid) <= 14.5:
                if other.y_mid < bar.y_mid - 1.0:
                    num_bars.append(other)
                    consumed_bar_ids.add(id(other))
                elif other.y_mid > bar.y_mid + 1.0:
                    den_bars.append(other)
                    consumed_bar_ids.add(id(other))

        # Tight vertical bounds around the bar using visual glyph center (extended if nested fraction bars are present)
        num_min_y = min([bar.y_mid - 16.0] + [nb.y_mid - 15.0 for nb in num_bars])
        den_max_y = max([bar.y_mid + 16.0] + [db.y_mid + 15.0 for db in den_bars])

        num_chars: list[CharToken] = []
        den_chars: list[CharToken] = []
        for ch in non_space:
            if id(ch) in consumed_char_ids:
                continue
            if ch.cx_mid < bar.x0 - x_pad or ch.cx_mid > bar.x1 + x_pad:
                continue
            if ch.cy_mid < bar.y_mid - 0.5 and ch.cy_mid >= num_min_y:
                num_chars.append(ch)
                consumed_char_ids.add(id(ch))
            elif ch.cy_mid > bar.y_mid + 0.5 and ch.cy_mid <= den_max_y:
                den_chars.append(ch)
                consumed_char_ids.add(id(ch))

        # Also pull in smaller script tokens immediately attached to a base char in num_chars or den_chars
        for ch in non_space:
            if id(ch) in consumed_char_ids or ch.size > 9.8:
                continue
            if ch.cx_mid < bar.x0 - x_pad or ch.cx_mid > bar.x1 + x_pad:
                continue
            if any(-1.0 <= ch.x0 - bc.x1 <= 9.5 and 1.0 <= abs(ch.origin_y - bc.origin_y) <= 8.0 for bc in num_chars):
                num_chars.append(ch)
                consumed_char_ids.add(id(ch))
            elif any(-1.0 <= ch.x0 - bc.x1 <= 9.5 and 1.0 <= abs(ch.origin_y - bc.origin_y) <= 8.0 for bc in den_chars):
                den_chars.append(ch)
                consumed_char_ids.add(id(ch))

        if (not num_chars and not num_bars) or (not den_chars and not den_bars):
            raise UnsupportedMathStructureError("unbalanced_fraction", bar.bbox)

        num_latex = _parse_math_box(num_chars, num_bars, rad_bars, depth=depth + 1)
        den_latex = _parse_math_box(den_chars, den_bars, rad_bars, depth=depth + 1)
        if not num_latex or not den_latex:
            raise UnsupportedMathStructureError("unbalanced_fraction", bar.bbox)

        fbbox = union_bbox(
            [bar.bbox]
            + [c.bbox for c in num_chars]
            + [c.bbox for c in den_chars]
            + [nb.bbox for nb in num_bars]
            + [db.bbox for db in den_bars]
        )
        avg_size = statistics.median([c.size for c in (num_chars + den_chars)] or [10.5])
        fraction_atoms.append(
            MathAtom(
                latex=f"\\frac{{{num_latex}}}{{{den_latex}}}",
                bbox=fbbox,
                origin_y=bar.y_mid + avg_size * 0.30,
                size=avg_size,
                is_fraction=True,
            )
        )

    remaining_chars = [c for c in non_space if id(c) not in consumed_char_ids]

    # Estimate main baseline of this box
    if fraction_atoms:
        main_origin_y = statistics.median([fa.origin_y for fa in fraction_atoms])
        main_size = statistics.median([fa.size for fa in fraction_atoms])
    else:
        max_char_size = max(c.size for c in remaining_chars)
        normal_Pool = [c for c in remaining_chars if c.size >= max_char_size * 0.85]
        main_origin_y = statistics.median([c.origin_y for c in (normal_Pool or remaining_chars)])
        main_size = statistics.median([c.size for c in (normal_Pool or remaining_chars)])

    # 2. Check for radicals (\sqrt) and group their radicands
    remaining_chars.sort(key=lambda c: c.x0)
    radical_atoms: list[MathAtom] = []
    rad_consumed_ids: set[int] = set()
    for idx, ch in enumerate(remaining_chars):
        if id(ch) in rad_consumed_ids:
            continue
        if ch.latex_token not in {r"\sqrt", r"\sqrt[3]", r"\sqrt[4]"}:
            continue
        rad_consumed_ids.add(id(ch))
        root_index_latex: str | None = None
        if ch.latex_token == r"\sqrt[3]":
            root_index_latex = "3"
        elif ch.latex_token == r"\sqrt[4]":
            root_index_latex = "4"
        else:
            # Look for a small index token immediately upper-left of the radical sign
            for prev_ch in remaining_chars[:idx]:
                if id(prev_ch) in rad_consumed_ids:
                    continue
                if (
                    prev_ch.size < main_size * 0.82
                    and ch.x0 - 10.0 <= prev_ch.x0 <= ch.x1 + 2.0
                    and prev_ch.cy_mid < ch.cy_mid
                ):
                    root_index_latex = prev_ch.latex_token
                    rad_consumed_ids.add(id(prev_ch))
                    break

        # Find vinculum bar near the top of the radical sign, if any
        vinculum: HorizontalBar | None = None
        for rb in rad_bars:
            if abs(rb.y_mid - ch.y0) <= 5.5 and rb.x0 <= ch.x1 + 6.0 and rb.x1 > ch.x1 + 2.0:
                vinculum = rb
                break

        radicand_chars: list[CharToken] = []
        radicand_fracs: list[MathAtom] = []
        if vinculum is not None:
            vx0, vx1 = ch.x1 - 2.0, vinculum.x1 + 2.0
            for rc in remaining_chars[idx + 1 :]:
                if id(rc) not in rad_consumed_ids and vx0 <= rc.cx_mid <= vx1:
                    radicand_chars.append(rc)
                    rad_consumed_ids.add(id(rc))
            for fa in list(fraction_atoms):
                fx_mid = (fa.bbox[0] + fa.bbox[2]) * 0.5
                if vx0 <= fx_mid <= vx1:
                    radicand_fracs.append(fa)
                    fraction_atoms.remove(fa)
        else:
            # Consume contiguous tokens immediately following \sqrt
            paren_depth = 0
            cur_x = ch.x1
            for rc in remaining_chars[idx + 1 :]:
                if id(rc) in rad_consumed_ids:
                    continue
                if rc.x0 - cur_x > main_size * 0.65 and paren_depth <= 0 and radicand_chars:
                    break
                if rc.norm_c in {"+", "-", "="} and paren_depth <= 0 and radicand_chars:
                    break
                radicand_chars.append(rc)
                rad_consumed_ids.add(id(rc))
                cur_x = rc.x1
                if rc.norm_c in {"(", "["}:
                    paren_depth += 1
                elif rc.norm_c in {")", "]"}:
                    paren_depth = max(0, paren_depth - 1)
                    if paren_depth == 0:
                        break

        if not radicand_chars and not radicand_fracs:
            raise UnsupportedMathStructureError("empty_radical", ch.bbox)

        if radicand_fracs and not radicand_chars and len(radicand_fracs) == 1:
            inner_latex = radicand_fracs[0].latex
            rbbox = union_bbox([ch.bbox, radicand_fracs[0].bbox])
        else:
            sub_atoms = [
                MathAtom(
                    latex=rc.latex_token,
                    bbox=rc.bbox,
                    origin_y=rc.origin_y,
                    size=rc.size,
                )
                for rc in radicand_chars
            ] + radicand_fracs
            sub_atoms.sort(key=lambda a: a.bbox[0])
            inner_latex = _assemble_atoms_with_scripts(sub_atoms, main_origin_y, main_size)
            rbbox = union_bbox([ch.bbox] + [a.bbox for a in sub_atoms])

        rad_cmd = f"\\sqrt[{root_index_latex}]{{{inner_latex}}}" if root_index_latex else f"\\sqrt{{{inner_latex}}}"
        radical_atoms.append(
            MathAtom(
                latex=rad_cmd,
                bbox=rbbox,
                origin_y=main_origin_y,
                size=main_size,
                is_radical=True,
            )
        )

    remaining_chars = [c for c in remaining_chars if id(c) not in rad_consumed_ids]

    # 3. Check for multi-row non-fraction vertical stacking (e.g. matrices or cases)
    normal_remaining = [
        c for c in remaining_chars
        if c.size >= main_size * 0.85 and c.latex_token not in _NARY_LATEX_TOKENS
    ]
    for i, c1 in enumerate(normal_remaining):
        for c2 in normal_remaining[i + 1 :]:
            if horizontal_overlap_ratio(c1.bbox, c2.bbox) > 0.45 and abs(c1.origin_y - c2.origin_y) > main_size * 0.95:
                raise UnsupportedMathStructureError(
                    "matrix_or_cases",
                    union_bbox([c.bbox for c in remaining_chars] + [fa.bbox for fa in fraction_atoms]),
                )

    # 4. Build initial atom list and attach scripts
    raw_atoms: list[MathAtom] = list(fraction_atoms) + list(radical_atoms)
    for ch in remaining_chars:
        raw_atoms.append(
            MathAtom(
                latex=ch.latex_token,
                bbox=ch.bbox,
                origin_y=ch.origin_y,
                size=ch.size,
                is_nary=ch.latex_token in _NARY_LATEX_TOKENS,
            )
        )
    raw_atoms.sort(key=lambda a: (round(a.bbox[0], 1), round(a.bbox[1], 1)))
    return _assemble_atoms_with_scripts(raw_atoms, main_origin_y, main_size)


def _classify_script_kind(
    atom: MathAtom,
    base_origin_y: float,
    base_size: float,
) -> Literal["normal", "sub", "sup"]:
    if atom.is_fraction or atom.is_radical or atom.is_nary:
        return "normal"
    dy = atom.origin_y - base_origin_y
    is_smaller = atom.size <= base_size * 0.86
    if (is_smaller and dy <= -base_size * 0.14) or dy <= -base_size * 0.28:
        return "sup"
    if (is_smaller and dy >= base_size * 0.11) or dy >= base_size * 0.24:
        return "sub"
    return "normal"


def _assemble_atoms_with_scripts(
    atoms: list[MathAtom],
    main_origin_y: float,
    main_size: float,
) -> str:
    if not atoms:
        return ""

    # First attach under/over limits that are horizontally overlapping an n-ary operator (\sum, \prod, \int, \lim)
    nary_indices = [i for i, a in enumerate(atoms) if a.is_nary]
    attached_to_nary: set[int] = set()
    for ni in nary_indices:
        nary_atom = atoms[ni]
        nx0, nx1 = nary_atom.bbox[0] - 6.0, nary_atom.bbox[2] + 10.0
        ny_mid = (nary_atom.bbox[1] + nary_atom.bbox[3]) * 0.5
        sub_parts: list[MathAtom] = []
        sup_parts: list[MathAtom] = []
        for idx, cand in enumerate(atoms):
            if idx == ni or idx in attached_to_nary or cand.is_fraction or cand.is_nary:
                continue
            cx_mid = (cand.bbox[0] + cand.bbox[2]) * 0.5
            if not (nx0 <= cx_mid <= nx1):
                continue
            cy_mid = (cand.bbox[1] + cand.bbox[3]) * 0.5
            if cand.size <= main_size * 0.88:
                if cy_mid > ny_mid + 1.5:
                    sub_parts.append(cand)
                    attached_to_nary.add(idx)
                elif cy_mid < ny_mid - 1.5:
                    sup_parts.append(cand)
                    attached_to_nary.add(idx)
        if sub_parts:
            sub_parts.sort(key=lambda a: a.bbox[0])
            nary_atom.sub_latex = "".join(a.latex for a in sub_parts)
        if sup_parts:
            sup_parts.sort(key=lambda a: a.bbox[0])
            nary_atom.sup_latex = "".join(a.latex for a in sup_parts)

    filtered_atoms = [a for i, a in enumerate(atoms) if i not in attached_to_nary]

    # Merge letters of known function names like l-i-m, s-i-n, c-o-s, l-o-g, l-n, m-a-x, m-i-n
    merged_atoms: list[MathAtom] = []
    i = 0
    while i < len(filtered_atoms):
        cur = filtered_atoms[i]
        if (
            not cur.is_fraction
            and not cur.is_radical
            and len(cur.latex) == 1
            and cur.latex.isalpha()
            and _classify_script_kind(cur, main_origin_y, main_size) == "normal"
        ):
            # Look ahead up to 6 single normal letters
            letters = [cur.latex]
            end_j = i
            while end_j + 1 < len(filtered_atoms) and len(letters) < 6:
                nxt = filtered_atoms[end_j + 1]
                if (
                    not nxt.is_fraction
                    and not nxt.is_radical
                    and len(nxt.latex) == 1
                    and nxt.latex.isalpha()
                    and _classify_script_kind(nxt, main_origin_y, main_size) == "normal"
                    and (nxt.bbox[0] - filtered_atoms[end_j].bbox[2]) <= main_size * 0.25
                ):
                    letters.append(nxt.latex)
                    end_j += 1
                else:
                    break
            word = "".join(letters)
            matched_fn = None
            for fn in sorted(_FUNCTION_WORDS, key=len, reverse=True):
                if word.startswith(fn):
                    matched_fn = fn
                    break
            if matched_fn is not None:
                fn_len = len(matched_fn)
                fn_atoms = filtered_atoms[i : i + fn_len]
                merged_atoms.append(
                    MathAtom(
                        latex=f"\\{matched_fn}",
                        bbox=union_bbox(a.bbox for a in fn_atoms),
                        origin_y=fn_atoms[0].origin_y,
                        size=fn_atoms[0].size,
                        is_nary=matched_fn in {"lim", "max", "min", "sup", "inf"},
                    )
                )
                i += fn_len
                continue
        merged_atoms.append(cur)
        i += 1

    # Group scripts (sub/sup) onto preceding base atom
    out_atoms: list[MathAtom] = []
    idx = 0
    cur_base_y = main_origin_y
    cur_base_size = main_size
    while idx < len(merged_atoms):
        atom = merged_atoms[idx]
        kind = _classify_script_kind(atom, cur_base_y, cur_base_size)
        if kind in {"sub", "sup"} and out_atoms:
            base_atom = out_atoms[-1]
            script_group: list[MathAtom] = [atom]
            script_y = atom.origin_y
            script_size = atom.size
            idx += 1
            while idx < len(merged_atoms):
                nxt = merged_atoms[idx]
                nxt_kind = _classify_script_kind(nxt, cur_base_y, cur_base_size)
                gap = nxt.bbox[0] - script_group[-1].bbox[2]
                if nxt_kind == kind and abs(nxt.origin_y - script_y) <= max(2.5, script_size * 0.35) and gap <= max(6.0, script_size * 0.85):
                    script_group.append(nxt)
                    idx += 1
                else:
                    break
            script_text = "".join(a.latex for a in script_group).strip()
            formatted_script = _format_script_content(script_text)
            if kind == "sub":
                base_atom.sub_latex = (
                    f"{base_atom.sub_latex}{formatted_script}" if base_atom.sub_latex else formatted_script
                )
            else:
                # Special case:^{\circ} token already has ^{\circ}
                if formatted_script == r"^{\circ}":
                    formatted_script = r"\circ"
                base_atom.sup_latex = (
                    f"{base_atom.sup_latex}{formatted_script}" if base_atom.sup_latex else formatted_script
                )
            base_atom.bbox = union_bbox([base_atom.bbox] + [a.bbox for a in script_group])
            continue

        out_atoms.append(atom)
        if not atom.is_fraction:
            cur_base_y = atom.origin_y
            cur_base_size = max(atom.size, main_size * 0.85)
        idx += 1

    # Render sequence with proper LaTeX spacing
    rendered_parts: list[str] = []
    for k, atom in enumerate(out_atoms):
        piece = atom.render()
        if rendered_parts:
            prev_piece = rendered_parts[-1]
            # Add space after a backslash command like \alpha, \pi, \sin before a letter/digit
            if re.search(r"\\[a-zA-Z]+$", prev_piece) and piece and (piece[0].isalnum() or piece[0] in {"(", "["}):
                rendered_parts.append(" ")
            elif prev_piece.endswith("}") and piece.startswith("\\"):
                rendered_parts.append(" ")
        rendered_parts.append(piece)

    latex = "".join(rendered_parts).strip()
    # Normalize standalone dots between numbers/variables in formulas:
    # Keep Vietnamese decimal comma as `{,}` or `,`, and convert multiplication `. ` between factors to `\cdot `
    latex = re.sub(r"(?<=[a-zA-Z0-9\}\)])\s*\.\s+(?=[a-zA-Z0-9\\\(\{])", r" \\cdot ", latex)
    return latex


@dataclass(slots=True)
class _VisualLine:
    chars: list[CharToken] = field(default_factory=list)
    bars: list[HorizontalBar] = field(default_factory=list)
    origin_y: float = 0.0
    base_size: float = 11.0

    @property
    def bbox(self) -> tuple[float, float, float, float]:
        return union_bbox([c.bbox for c in self.chars] + [b.bbox for b in self.bars])


def _cluster_visual_lines(
    chars: list[CharToken],
    fraction_bars: list[HorizontalBar],
) -> list[_VisualLine]:
    """Group page characters and fraction bars into visual lines / display equation bands."""
    if not chars and not fraction_bars:
        return []

    non_space_chars = [c for c in chars if not c.norm_c.isspace()]
    median_size = statistics.median([c.size for c in non_space_chars]) if non_space_chars else 11.0

    # Group overlapping or nested fraction bars into connected 2D fraction complexes.
    # Sort widest bars first so outer bars are registered before inner nested bars.
    bar_clusters: list[list[HorizontalBar]] = []
    for bar in sorted(fraction_bars, key=lambda b: (-b.width, b.y_mid, b.x0)):
        placed = False
        for cluster in bar_clusters:
            for existing in cluster:
                same_row = abs(bar.y_mid - existing.y_mid) <= 4.5
                nested = (
                    (
                        existing.width >= bar.width + 6.0
                        and bar.x0 >= existing.x0 - 2.5
                        and bar.x1 <= existing.x1 + 2.5
                    )
                    or (
                        bar.width >= existing.width + 6.0
                        and existing.x0 >= bar.x0 - 2.5
                        and existing.x1 <= bar.x1 + 2.5
                    )
                ) and abs(bar.y_mid - existing.y_mid) <= 14.5
                if same_row or nested:
                    cluster.append(bar)
                    placed = True
                    break
            if placed:
                break
        if not placed:
            bar_clusters.append([bar])

    # Identify characters that belong vertically & horizontally to the CLOSEST fraction bar cluster
    char_to_bar_cluster: dict[int, int] = {}
    char_to_bar_dist: dict[int, float] = {}
    cluster_y_ranges: list[tuple[float, float, float]] = []
    for c_idx, cluster in enumerate(bar_clusters):
        top_level = [
            b for b in cluster
            if not any(
                o.width >= b.width + 6.0 and b.x0 >= o.x0 - 2.5 and b.x1 <= o.x1 + 2.5 and 1.5 < abs(b.y_mid - o.y_mid) <= 14.5
                for o in cluster
            )
        ]
        anchor_y = statistics.median([b.y_mid for b in (top_level or cluster)])
        cluster_y_ranges.append((0.0, 0.0, anchor_y))
        for ch in chars:
            for b in cluster:
                x_pad = max(1.5, min(3.2, b.width * 0.08))
                if b.x0 - x_pad <= ch.cx_mid <= b.x1 + x_pad:
                    v_dist = abs(ch.cy_mid - b.y_mid)
                    if 0.5 <= v_dist <= 15.0:
                        prev_dist = char_to_bar_dist.get(id(ch), 999.0)
                        if v_dist < prev_dist:
                            char_to_bar_dist[id(ch)] = v_dist
                            char_to_bar_cluster[id(ch)] = c_idx

    # Second pass: attach smaller script chars immediately right-adjacent to a char already assigned to a bar cluster
    for ch in chars:
        if id(ch) in char_to_bar_cluster or ch.size > median_size * 0.86 or ch.norm_c.isspace():
            continue
        for c_idx, cluster in enumerate(bar_clusters):
            cluster_chars = [bc for bc in chars if char_to_bar_cluster.get(id(bc)) == c_idx]
            if any(-1.0 <= ch.x0 - bc.x1 <= 9.5 and 1.0 <= abs(ch.origin_y - bc.origin_y) <= 8.0 for bc in cluster_chars):
                char_to_bar_cluster[id(ch)] = c_idx
                break

    # Build initial lines from bar clusters
    lines: list[_VisualLine] = []
    for c_idx, cluster in enumerate(bar_clusters):
        _, _, anchor_y = cluster_y_ranges[c_idx]
        c_chars = [ch for ch in chars if char_to_bar_cluster.get(id(ch)) == c_idx]
        lines.append(
            _VisualLine(
                chars=c_chars,
                bars=list(cluster),
                origin_y=anchor_y + median_size * 0.30,
                base_size=median_size,
            )
        )

    # Separate remaining characters into normal-size and script-size
    unassigned = [ch for ch in chars if id(ch) not in char_to_bar_cluster]
    normal_chars = [ch for ch in unassigned if ch.size >= median_size * 0.85 or ch.norm_c.isspace()]
    script_chars = [ch for ch in unassigned if ch.size < median_size * 0.85 and not ch.norm_c.isspace()]

    # Assign normal_chars to existing lines or create new baseline lines
    for ch in sorted(normal_chars, key=lambda c: (round(c.origin_y, 1), c.x0)):
        best_line: _VisualLine | None = None
        best_dy = 999.0
        for line in lines:
            tol = 5.2 if line.bars else 3.4
            dy = abs(ch.origin_y - line.origin_y)
            if dy <= tol and dy < best_dy:
                best_dy = dy
                best_line = line
        if best_line is not None:
            best_line.chars.append(ch)
            if not best_line.bars and not ch.norm_c.isspace():
                normal_oys = [c.origin_y for c in best_line.chars if not c.norm_c.isspace() and c.size >= median_size * 0.85]
                if normal_oys:
                    best_line.origin_y = statistics.median(normal_oys)
        else:
            lines.append(
                _VisualLine(
                    chars=[ch],
                    bars=[],
                    origin_y=ch.origin_y,
                    base_size=max(ch.size, median_size * 0.85),
                )
            )

    # Attach script_chars to the closest line where a base token precedes or overlaps horizontally
    for sch in sorted(script_chars, key=lambda c: c.x0):
        best_line = None
        best_score = 999.0
        for line in lines:
            dy = abs(sch.origin_y - line.origin_y)
            if dy > line.base_size * 0.78:
                continue
            # Check horizontal proximity to any non-space char or bar in line
            min_dx = min(
                (
                    max(0.0, sch.x0 - c.x1, c.x0 - sch.x1)
                    for c in line.chars
                    if not c.norm_c.isspace()
                ),
                default=999.0,
            )
            if min_dx <= 16.0:
                score = dy + min_dx * 0.25
                if score < best_score:
                    best_score = score
                    best_line = line
        if best_line is not None:
            best_line.chars.append(sch)
        else:
            lines.append(_VisualLine(chars=[sch], bars=[], origin_y=sch.origin_y, base_size=median_size))

    # Filter out lines that contain only whitespace
    lines = [
        line for line in lines
        if line.bars or any(not c.norm_c.isspace() for c in line.chars)
    ]
    lines.sort(key=lambda l: (round(l.bbox[1], 1), round(l.bbox[0], 1)))
    return lines


def _word_has_vietnamese_diacritics(word: str) -> bool:
    for ch in word:
        if ord(ch) > 127 and ch.isalpha() and not _is_math_alphanumeric(ch):
            # Exclude Greek letters
            if not (0x0370 <= ord(ch) <= 0x03FF):
                return True
    return False


def _text_has_equation_sign(flat_text: str) -> bool:
    return bool(
        re.search(r"=\s*[\d\w\(\+\-\.\,]", flat_text)
        or re.search(r"[\d\w\)\.\,]\s*=", flat_text)
        or ("=" in flat_text and any(ch.isalpha() for ch in flat_text))
    )


def _segment_visual_line(
    line: _VisualLine,
    radical_bars: list[HorizontalBar],
    page_width: float,
    page_height: float,
) -> tuple[_Block | None, int, list[MathRegion]]:
    """Convert a _VisualLine into a _Block (`equation` or `paragraph`), returning (block, eq_count, unresolved)."""
    unresolved: list[MathRegion] = []
    non_space = [c for c in line.chars if not c.norm_c.isspace()]
    if not non_space and not line.bars:
        return None, 0, unresolved

    lbbox = line.bbox
    norm_bbox = normalize_bbox_to_1000(lbbox, page_width, page_height)
    font_size = max(6.0, min(36.0, round(line.base_size, 1)))

    # Quick check: does this line have ANY math indicators?
    has_bars = bool(line.bars)
    has_math_chars = any(c.is_math_font or c.is_math_symbol for c in non_space)
    has_scripts = any(
        _classify_script_kind(
            MathAtom(latex=c.latex_token, bbox=c.bbox, origin_y=c.origin_y, size=c.size),
            line.origin_y,
            line.base_size,
        ) in {"sub", "sup"}
        for c in non_space
    )
    flat_text = "".join(c.norm_c for c in sorted(line.chars, key=lambda c: c.x0)).strip()
    has_eq_sign = _text_has_equation_sign(flat_text)

    if not (has_bars or has_math_chars or has_scripts or has_eq_sign):
        # Pure prose line
        return (
            _Block(
                kind="paragraph",
                spans=[
                    _Span(
                        kind="text",
                        content=flat_text,
                        bold=sum(c.bold for c in non_space) > len(non_space) // 2,
                        italic=sum(c.italic for c in non_space) > len(non_space) // 2,
                    )
                ],
                rows=[],
                bbox=norm_bbox,
                alignment="left",
                font_size=font_size,
            ),
            0,
            unresolved,
        )

    # Check if the line starts with Vietnamese prose before a colon `:` or before an inline formula
    # e.g., "-Bán kính đặt lực ma sát: R_0 = 1/3 D = ..." or "+ Chọn khớp nối trục đàn hồi có: \eta_k = 0,99"
    # or "Trong đó: L_0 = 800 mm là chiều dài..."
    ordered_all = sorted(line.chars, key=lambda c: c.x0)

    # Find prose prefix / suffix vs math core.
    # Let's check if there is a colon `:` in the normal baseline outside fraction bars where the prefix before `:` is prose
    # AND the suffix after `:` actually contains math (e.g. `=` or fraction bar or math symbols/scripts), not pure prose!
    bar_x0 = min((b.x0 for b in line.bars), default=99999.0)
    colon_split_x: float | None = None
    for idx, ch in enumerate(ordered_all):
        if ch.norm_c == ":" and ch.x0 < bar_x0 - 2.0:
            prefix_str = "".join(c.norm_c for c in ordered_all[: idx + 1])
            suffix_non_space = [c for c in ordered_all[idx + 1 :] if not c.norm_c.isspace()]
            suffix_str = "".join(c.norm_c for c in ordered_all[idx + 1 :]).strip()
            suffix_has_math = (
                bool(line.bars)
                or "=" in suffix_str
                or any(c.is_math_font or c.is_math_symbol for c in suffix_non_space)
                or any(
                    _classify_script_kind(
                        MathAtom(latex=c.latex_token, bbox=c.bbox, origin_y=c.origin_y, size=c.size),
                        line.origin_y,
                        line.base_size,
                    ) in {"sub", "sup"}
                    for c in suffix_non_space
                )
            )
            if "=" not in prefix_str and suffix_has_math and (
                any(_word_has_vietnamese_diacritics(w) or w.lower() in _VIETNAMESE_PROSE_WORDS for w in re.findall(r"\w+", prefix_str))
                or len(re.findall(r"[a-zA-ZÀ-ỹ]{2,}", prefix_str)) >= 1
            ):
                colon_split_x = ch.x1
                break

    if colon_split_x is not None:
        prefix_chars = [c for c in line.chars if c.x1 <= colon_split_x + 0.5]
        suffix_chars = [c for c in line.chars if c.x1 > colon_split_x + 0.5]
        prefix_text = "".join(c.norm_c for c in sorted(prefix_chars, key=lambda c: c.x0))
        if not prefix_text.endswith(" "):
            prefix_text += " "

        # Parse suffix as inline math (plus optional trailing prose/units)
        spans: list[_Span] = [
            _Span(
                kind="text",
                content=prefix_text,
                bold=any(c.bold for c in prefix_chars if not c.norm_c.isspace()),
                italic=False,
            )
        ]
        math_spans, eq_cnt, unres = _build_math_and_trailing_spans(
            suffix_chars,
            line.bars,
            radical_bars,
            line.origin_y,
            line.base_size,
            page_width,
            page_height,
        )
        if eq_cnt == 0 and not unres:
            # Suffix had no actual math after all
            return (
                _Block(
                    kind="paragraph",
                    spans=[_Span(kind="text", content=flat_text)],
                    rows=[],
                    bbox=norm_bbox,
                    alignment="left",
                    font_size=font_size,
                ),
                0,
                [],
            )
        spans.extend(math_spans)
        unresolved.extend(unres)
        return (
            _Block(
                kind="paragraph",
                spans=spans,
                rows=[],
                bbox=norm_bbox,
                alignment="left",
                font_size=font_size,
            ),
            eq_cnt,
            unresolved,
        )

    # Check if the line is primarily Vietnamese prose with an inline symbol/formula embedded in it,
    # e.g. "Chọn đường kính D = 650 mm" or "Ta có công thức tính vận tốc v = L/t như sau"
    # Group normal baseline words outside fraction bars and see which segments are prose vs math.
    words_info = _find_prose_and_math_segments(line, page_width)
    if words_info["is_mixed_prose"]:
        spans, eq_cnt, unres = _build_mixed_prose_line_spans(
            line,
            radical_bars,
            page_width,
            page_height,
        )
        if spans:
            return (
                _Block(
                    kind="paragraph",
                    spans=spans,
                    rows=[],
                    bbox=norm_bbox,
                    alignment="left",
                    font_size=font_size,
                ),
                eq_cnt,
                unres,
            )

    # Standalone display equation line!
    math_spans, eq_cnt, unres = _build_math_and_trailing_spans(
        line.chars,
        line.bars,
        radical_bars,
        line.origin_y,
        line.base_size,
        page_width,
        page_height,
    )
    unresolved.extend(unres)
    if not math_spans:
        return None, 0, unresolved

    # Determine alignment (centered vs left-indented equation)
    block_w = lbbox[2] - lbbox[0]
    center_delta = abs(((lbbox[0] + lbbox[2]) * 0.5) - (page_width * 0.5))
    alignment: Literal["left", "center", "right", "justify"] = (
        "center" if (block_w < page_width * 0.82 and center_delta < page_width * 0.12) else "left"
    )
    kind: Literal["paragraph", "equation"] = "equation" if any(s.kind == "math" for s in math_spans) else "paragraph"
    return (
        _Block(
            kind=kind,
            spans=math_spans,
            rows=[],
            bbox=norm_bbox,
            alignment=alignment,
            font_size=font_size,
        ),
        eq_cnt,
        unresolved,
    )


def _find_prose_and_math_segments(line: _VisualLine, page_width: float) -> dict[str, Any]:
    """Check whether a line has Vietnamese prose words mixed with math rather than being a pure display equation."""
    # Exclude characters inside fraction bars or scripts
    bar_intervals = [(b.x0 - 2.0, b.x1 + 2.0) for b in line.bars]
    normal_chars = [
        c for c in sorted(line.chars, key=lambda ch: ch.x0)
        if not any(bx0 <= c.cx_mid <= bx1 for bx0, bx1 in bar_intervals)
        and _classify_script_kind(
            MathAtom(latex=c.latex_token, bbox=c.bbox, origin_y=c.origin_y, size=c.size),
            line.origin_y,
            line.base_size,
        ) == "normal"
    ]
    text = "".join(c.norm_c for c in normal_chars)
    # Strip trailing parenthesized unit like `(mm)`, `(m/phút)`, `(vg/ph)`, `(kW)`, `(1.2)`
    text_no_unit = re.sub(r"\(\s*[^()]{1,18}\s*\)\s*$", "", text).strip()
    words = re.findall(r"[a-zA-ZÀ-ỹ]{2,}", text_no_unit)
    prose_words = [
        w for w in words
        if w.lower() not in _FUNCTION_WORDS
        and (_word_has_vietnamese_diacritics(w) or w.lower() in _VIETNAMESE_PROSE_WORDS or len(w) >= 4)
    ]
    return {"is_mixed_prose": len(prose_words) >= 2}


def _build_mixed_prose_line_spans(
    line: _VisualLine,
    radical_bars: list[HorizontalBar],
    page_width: float,
    page_height: float,
) -> tuple[list[_Span], int, list[MathRegion]]:
    """Split a mixed prose+math paragraph line into text spans and complete inline math spans."""
    # Partition the horizontal extent of `line` into word/token chunks, keeping fraction bars and equations intact.
    ordered = sorted(line.chars, key=lambda c: c.x0)
    # Build horizontal x-intervals for all fraction bars
    math_x_intervals: list[ list[float] ] = [[b.x0 - 3.0, b.x1 + 3.0] for b in line.bars]

    # Also mark any char that is a script, math font, math symbol, or `=`, `+`, `<`, `>` between non-prose tokens
    for ch in ordered:
        if ch.norm_c.isspace():
            continue
        is_script = _classify_script_kind(
            MathAtom(latex=ch.latex_token, bbox=ch.bbox, origin_y=ch.origin_y, size=ch.size),
            line.origin_y,
            line.base_size,
        ) in {"sub", "sup"}
        if is_script or ch.is_math_font or ch.is_math_symbol or ch.norm_c in {"=", "±", "≤", "≥", "≠", "≈"}:
            math_x_intervals.append([ch.x0 - 2.0, ch.x1 + 2.0])

    if not math_x_intervals:
        flat = "".join(c.norm_c for c in ordered).strip()
        return ([_Span(kind="text", content=flat)] if flat else []), 0, []

    # Merge overlapping/nearby math x-intervals and expand across operators/numbers/variables that belong to the formula
    math_x_intervals.sort(key=lambda iv: iv[0])
    merged_iv: list[list[float]] = []
    for iv in math_x_intervals:
        if merged_iv and iv[0] <= merged_iv[-1][1] + 14.0:
            merged_iv[-1][1] = max(merged_iv[-1][1], iv[1])
        else:
            merged_iv.append(list(iv))

    # Expand each math interval left/right to absorb connected equation tokens (numbers, single-letter vars, operators, parens)
    # while stopping at multi-letter Vietnamese prose words!
    non_space = [c for c in ordered if not c.norm_c.isspace()]
    # Group contiguous non-space baseline characters into word chunks
    word_chunks: list[list[CharToken]] = []
    for ch in non_space:
        if word_chunks and (ch.x0 - word_chunks[-1][-1].x1) <= max(1.8, line.base_size * 0.20):
            word_chunks[-1].append(ch)
        else:
            word_chunks.append([ch])

    def _is_prose_chunk(chunk: list[CharToken]) -> bool:
        if any(
            _classify_script_kind(
                MathAtom(latex=c.latex_token, bbox=c.bbox, origin_y=c.origin_y, size=c.size),
                line.origin_y,
                line.base_size,
            ) in {"sub", "sup"}
            for c in chunk
        ):
            return False
        if any(c.is_math_font or c.is_math_symbol for c in chunk):
            return False
        s = "".join(c.norm_c for c in chunk).strip(" ,;:.!?-()")
        if not s or s.lower() in _FUNCTION_WORDS:
            return False
        if _word_has_vietnamese_diacritics(s) or s.lower() in _VIETNAMESE_PROSE_WORDS:
            return True
        if s.isalpha() and len(s) >= 3:
            return True
        return False

    for iv in merged_iv:
        changed = True
        while changed:
            changed = False
            for chunk in word_chunks:
                cx0, cx1 = chunk[0].x0, chunk[-1].x1
                if cx0 >= iv[0] and cx1 <= iv[1]:
                    continue
                if _is_prose_chunk(chunk):
                    continue
                gap = max(0.0, iv[0] - cx1, cx0 - iv[1])
                if gap <= 12.0:
                    iv[0] = min(iv[0], cx0)
                    iv[1] = max(iv[1], cx1)
                    changed = True

    # Re-merge intervals after expansion
    merged_iv.sort(key=lambda iv: iv[0])
    final_ivs: list[list[float]] = []
    for iv in merged_iv:
        if final_ivs and iv[0] <= final_ivs[-1][1] + 8.0:
            final_ivs[-1][1] = max(final_ivs[-1][1], iv[1])
        else:
            final_ivs.append(list(iv))

    spans: list[_Span] = []
    eq_count = 0
    unresolved: list[MathRegion] = []
    cur_x = -9999.0

    for iv0, iv1 in final_ivs:
        prose_part = [c for c in ordered if cur_x < c.cx_mid < iv0]
        if prose_part:
            ptext = "".join(c.norm_c for c in prose_part)
            if ptext:
                if spans and spans[-1].kind == "math" and not ptext[0].isspace() and ptext[0] not in ".,;:!?)":
                    ptext = " " + ptext
                if not ptext[-1].isspace() and ptext[-1] not in "([":
                    ptext = ptext + " "
                spans.append(_Span(kind="text", content=ptext))
        m_chars = [c for c in line.chars if iv0 <= c.cx_mid <= iv1]
        m_bars = [b for b in line.bars if iv0 - 2.0 <= (b.x0 + b.x1) * 0.5 <= iv1 + 2.0]
        m_spans, m_eqs, m_unres = _build_math_and_trailing_spans(
            m_chars,
            m_bars,
            radical_bars,
            line.origin_y,
            line.base_size,
            page_width,
            page_height,
        )
        spans.extend(m_spans)
        eq_count += m_eqs
        unresolved.extend(m_unres)
        cur_x = iv1

    tail_part = [c for c in ordered if c.cx_mid > cur_x]
    if tail_part:
        ttext = "".join(c.norm_c for c in tail_part)
        if ttext:
            if spans and spans[-1].kind == "math" and not ttext[0].isspace() and ttext[0] not in ".,;:!?)":
                ttext = " " + ttext
            spans.append(_Span(kind="text", content=ttext))

    return spans, eq_count, unresolved


def _build_math_and_trailing_spans(
    chars: list[CharToken],
    bars: list[HorizontalBar],
    radical_bars: list[HorizontalBar],
    line_origin_y: float,
    line_base_size: float,
    page_width: float,
    page_height: float,
) -> tuple[list[_Span], int, list[MathRegion]]:
    """Build math span (+ optional trailing equation number / unit text span) from a math region."""
    unresolved: list[MathRegion] = []
    non_space = [c for c in chars if not c.norm_c.isspace()]
    if not non_space and not bars:
        return [], 0, unresolved

    # Separate trailing Vietnamese prose or parenthesized equation number / unit at the far right of the line
    # e.g. `(1.2)`, `(m/phút)`, or `là tiết diện ngang của thùng` if it is outside all fraction bars.
    max_bar_x1 = max((b.x1 for b in bars), default=-9999.0)
    ordered = sorted(non_space, key=lambda c: c.x0)
    trailing_text = ""
    math_chars = list(chars)

    # First check if there is a trailing Vietnamese prose clause to the right of all fraction bars and `=` signs
    max_eq_x1 = max(
        [max_bar_x1] + [c.x1 for c in ordered if c.norm_c in {"=", "+", "-", "×", "÷"} or c.is_math_symbol],
        default=-9999.0,
    )
    tail_candidates = [c for c in ordered if c.x0 > max_bar_x1 + 1.0]
    if tail_candidates:
        # Group tail_candidates into words
        tail_words: list[list[CharToken]] = []
        for tc in tail_candidates:
            if tail_words and (tc.x0 - tail_words[-1][-1].x1) <= max(1.8, line_base_size * 0.22):
                tail_words[-1].append(tc)
            else:
                tail_words.append([tc])
        # Find the first word in tail_words (after max_eq_x1) that starts a Vietnamese prose phrase (no scripts)
        for w_idx, w_chars in enumerate(tail_words):
            if w_chars[0].x0 <= max_eq_x1 + 1.0:
                continue
            has_script_in_word = any(
                _classify_script_kind(
                    MathAtom(latex=c.latex_token, bbox=c.bbox, origin_y=c.origin_y, size=c.size),
                    line_origin_y,
                    line_base_size,
                ) in {"sub", "sup"}
                for c in w_chars
            )
            if has_script_in_word:
                continue
            w_str = "".join(c.norm_c for c in w_chars).strip(" ,;:.!?()")
            if (_word_has_vietnamese_diacritics(w_str) or w_str.lower() in _VIETNAMESE_PROSE_WORDS) and len(tail_words) - w_idx >= 2:
                split_x0 = w_chars[0].x0
                trailing_all = [c for c in chars if c.x0 >= split_x0 - 0.5]
                trailing_text = " " + "".join(c.norm_c for c in sorted(trailing_all, key=lambda c: c.x0)).strip()
                math_chars = [c for c in chars if c.x0 < split_x0 - 0.5]
                ordered = [c for c in ordered if c.x0 < split_x0 - 0.5]
                break

    # Check if the end of `ordered` forms a parenthesized group `( ... )` to the right of all fraction bars
    if not trailing_text and ordered and ordered[-1].norm_c == ")" and ordered[-1].x0 > max_bar_x1:
        open_idx: int | None = None
        depth = 0
        for idx in range(len(ordered) - 1, -1, -1):
            tok = ordered[idx]
            if tok.x0 <= max_bar_x1:
                break
            if tok.norm_c == ")":
                depth += 1
            elif tok.norm_c == "(":
                depth -= 1
                if depth == 0:
                    open_idx = idx
                    break
        if open_idx is not None and open_idx > 0:
            group_chars = ordered[open_idx:]
            inner_str = "".join(c.norm_c for c in group_chars[1:-1]).strip()
            gap_before = group_chars[0].x0 - ordered[open_idx - 1].x1
            # Treat as trailing text span if it is an equation number like `(1)` / `(2.1)` OR a unit with Vietnamese diacritics like `(m/phút)` / `(kết quả)`
            is_eq_num = bool(re.match(r"^\d+(?:[.,]\d+)*$", inner_str)) and gap_before >= 12.0
            is_vn_unit = _word_has_vietnamese_diacritics(inner_str)
            if is_eq_num or is_vn_unit:
                split_x0 = group_chars[0].x0
                trailing_all = [c for c in chars if c.x0 >= split_x0 - 0.5]
                trailing_text = " " + "".join(c.norm_c for c in sorted(trailing_all, key=lambda c: c.x0)).strip()
                math_chars = [c for c in chars if c.x0 < split_x0 - 0.5]

    region_bbox = union_bbox([c.bbox for c in math_chars if not c.norm_c.isspace()] + [b.bbox for b in bars])
    norm_bbox = normalize_bbox_to_1000(region_bbox, page_width, page_height)

    try:
        latex = _parse_math_box(math_chars, bars, radical_bars).lstrip("`").strip()
        if not latex or not any(ch.isalnum() or ch == "\\" for ch in latex):
            return ([_Span(kind="text", content=latex + trailing_text)] if (latex or trailing_text) else []), 0, unresolved
        # Verify that the generated LaTeX compiles to valid OMML
        latex_to_omml(latex)
    except Exception as exc:
        reason = exc.reason if isinstance(exc, UnsupportedMathStructureError) else "omml_compile_failed"
        unresolved.append(
            MathRegion(
                bbox=region_bbox,
                kind="display" if bars else "inline",
                status="unsupported",
                reason=reason,
                source="rules",
                trailing_text=trailing_text or None,
            )
        )
        spans: list[_Span] = [_Span(kind="math", content=r"\invalid", bbox=norm_bbox)]
        if trailing_text:
            spans.append(_Span(kind="text", content=trailing_text))
        return spans, 0, unresolved

    spans = [_Span(kind="math", content=latex, bbox=norm_bbox)]
    if trailing_text:
        spans.append(_Span(kind="text", content=trailing_text))
    return spans, 1, unresolved


def _build_table_block(
    table_info: dict[str, Any],
    page_width: float,
    page_height: float,
    base_size: float,
) -> tuple[_Block, int]:
    """Convert a verified PDF table dict into a _Block(kind='table'), converting simple cell formulas to math spans."""
    tbbox = table_info["bbox"]
    norm_bbox = normalize_bbox_to_1000(tbbox, page_width, page_height)
    raw_rows = table_info["rows"]
    rows_out: list[_Row] = []
    table_eqs = 0

    for raw_row in raw_rows:
        cells_out: list[_Cell] = []
        for col_idx, raw_cell in enumerate(raw_row):
            if raw_cell is None and col_idx > 0:
                cells_out.append(_Cell(spans=[_Span(kind="text", content="__DOCX_CELL_MERGED_LEFT__")]))
                continue
            text = "" if raw_cell is None else str(raw_cell).strip()
            if not text:
                cells_out.append(_Cell(spans=[_Span(kind="text", content="")]))
                continue
            # Normalize Symbol PUA and Unicode Math Alphanumeric characters inside table cells
            unicode_chars: list[str] = []
            latex_chars: list[str] = []
            has_math_sym = False
            for ch in text:
                if ch in _SYMBOL_PUA_MAP:
                    mapped = _SYMBOL_PUA_MAP[ch]
                    unicode_chars.append(_LATEX_TO_UNICODE_MAP.get(mapped, mapped))
                    latex_chars.append(mapped)
                    has_math_sym = True
                elif _is_math_alphanumeric(ch):
                    nfkc = unicodedata.normalize("NFKC", ch)
                    unicode_chars.append(nfkc)
                    latex_chars.append(_UNICODE_MATH_MAP.get(nfkc, nfkc))
                    has_math_sym = True
                elif ch in _UNICODE_MATH_MAP:
                    unicode_chars.append(ch)
                    latex_chars.append(_UNICODE_MATH_MAP[ch])
                    has_math_sym = True
                else:
                    unicode_chars.append(ch)
                    latex_chars.append(ch)
            unicode_str = "".join(unicode_chars)
            latex_str = "".join(latex_chars).replace("*", r" \cdot ")
            prose_words = [
                w for w in re.findall(r"[a-zA-ZÀ-ỹ]{2,}", unicode_str)
                if _word_has_vietnamese_diacritics(w) or w.lower() in _VIETNAMESE_PROSE_WORDS
            ]
            if (
                len(prose_words) == 0
                and len(unicode_str) <= 60
                and "\n" not in unicode_str
                and (has_math_sym or is_likely_formula_text(unicode_str))
            ):
                try:
                    latex_to_omml(latex_str)
                    cells_out.append(_Cell(spans=[_Span(kind="math", content=latex_str)]))
                    table_eqs += 1
                    continue
                except Exception:
                    pass
            cells_out.append(_Cell(spans=[_Span(kind="text", content=unicode_str)]))
        if cells_out:
            rows_out.append(_Row(cells=cells_out))

    return (
        _Block(
            kind="table",
            spans=[],
            rows=rows_out,
            bbox=norm_bbox,
            alignment="center",
            font_size=max(6.0, min(36.0, round(base_size * 0.9, 1))),
        ),
        table_eqs,
    )


def page_has_digital_math_candidates(page, *, detect_tables: bool = True) -> bool:
    """Return True if a digital PDF page has fraction bars, math fonts/symbols, or scripts."""
    text_dict = page.get_text("rawdict", sort=True)
    repair_raw_text_blocks(text_dict.get("blocks", []))
    tables = extract_verified_pdf_tables(page) if detect_tables else []
    table_bboxes = [t["bbox"] for t in tables]
    chars = _extract_char_tokens(text_dict, excluded_bboxes=table_bboxes)
    if not chars:
        return False
    cand_bars = extract_candidate_horizontal_bars(page, table_bboxes)
    frac_bars = classify_fraction_bars(
        cand_bars,
        [{"c": c.norm_c, "bbox": c.bbox, "origin_y": c.origin_y, "size": c.size} for c in chars],
    )
    if frac_bars:
        return True
    if any(c.is_math_font or c.is_math_symbol for c in chars if not c.norm_c.isspace()):
        return True
    # Check for clear subscript/superscript pairs or equation relations on a line
    lines = _cluster_visual_lines(chars, [])
    for line in lines:
        flat_text = "".join(c.norm_c for c in sorted(line.chars, key=lambda c: c.x0)).strip()
        if _text_has_equation_sign(flat_text):
            return True
        for c in line.chars:
            if not c.norm_c.isspace():
                if (
                    _classify_script_kind(
                        MathAtom(latex=c.latex_token, bbox=c.bbox, origin_y=c.origin_y, size=c.size),
                        line.origin_y,
                        line.base_size,
                    )
                    in {"sub", "sup"}
                ):
                    return True
    return False


def extract_digital_math_page(
    page,
    *,
    detect_tables: bool = True,
    preserve_images: bool = True,
) -> DigitalMathPageResult:
    """Reconstruct a digital PDF page with rule-based equation parsing."""
    page_width = float(page.rect.width)
    page_height = float(page.rect.height)
    text_dict = page.get_text("rawdict", sort=True)
    repair_raw_text_blocks(text_dict.get("blocks", []))

    verified_tables = extract_verified_pdf_tables(page) if detect_tables else []
    table_bboxes = [t["bbox"] for t in verified_tables]
    figure_bboxes = extract_pdf_figure_regions(page, text_dict, table_bboxes) if preserve_images else []

    chars = _extract_char_tokens(text_dict, excluded_bboxes=table_bboxes + figure_bboxes)
    non_space_chars = [c for c in chars if not c.norm_c.isspace()]
    base_size = statistics.median([c.size for c in non_space_chars]) if non_space_chars else 11.0

    cand_bars = extract_candidate_horizontal_bars(page, table_bboxes + figure_bboxes)
    frac_bars = classify_fraction_bars(
        cand_bars,
        [{"c": c.norm_c, "bbox": c.bbox, "origin_y": c.origin_y, "size": c.size} for c in non_space_chars],
    )
    frac_bar_ids = {id(b) for b in frac_bars}
    radical_bars = [b for b in cand_bars if id(b) not in frac_bar_ids]

    visual_lines = _cluster_visual_lines(chars, frac_bars)

    dated_items: list[tuple[float, float, _Block]] = []
    recognized_count = 0
    all_unresolved: list[MathRegion] = []

    for vline in visual_lines:
        block, eq_cnt, unres = _segment_visual_line(vline, radical_bars, page_width, page_height)
        if block is not None:
            lb = vline.bbox
            dated_items.append((round(lb[1], 1), round(lb[0], 1), block))
            recognized_count += eq_cnt
            all_unresolved.extend(unres)

    for tbl in verified_tables:
        tblock, teqs = _build_table_block(tbl, page_width, page_height, base_size)
        tb = tbl["bbox"]
        dated_items.append((round(tb[1], 1), round(tb[0], 1), tblock))
        recognized_count += teqs

    for ibbox in figure_bboxes:
        fblock = _Block(
            kind="figure", spans=[], rows=[],
            bbox=normalize_bbox_to_1000(ibbox, page_width, page_height),
            alignment="center", font_size=11.0,
        )
        dated_items.append((round(ibbox[1], 1), round(ibbox[0], 1), fblock))

    dated_items.sort(key=lambda item: (item[0], item[1]))
    blocks = [item[2] for item in dated_items]

    has_math = bool(recognized_count > 0 or all_unresolved)
    if not has_math:
        return DigitalMathPageResult(
            has_math=False,
            page_model=_Page(has_math=False, blocks=[], uncertain=False),
            recognized_by_rules=0,
            unresolved_regions=[],
        )

    return DigitalMathPageResult(
        has_math=True,
        page_model=_Page(has_math=True, blocks=blocks, uncertain=False),
        recognized_by_rules=recognized_count,
        unresolved_regions=all_unresolved,
    )
