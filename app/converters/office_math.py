"""Structured LaTeX -> MathML -> Office Math (OMML) adapter.

Builds native Word Equation (`m:oMath` / `m:oMathPara`) XML trees using `lxml` /
`python-docx` OxmlElement APIs so that fractions, scripts, radicals, accents,
bars, group characters, functions, limits, n-ary operators, delimiters,
matrices, cases, and aligned equations retain editable structure and accurate
argument boundaries.

Baseline tested versions:
- latex2mathml == 3.81.1
- mathml2omml == 0.0.2 (used only for OP_DICT operator metadata)
- python-docx == 1.2.0
- lxml == 6.1.3
"""
from __future__ import annotations

import re
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from latex2mathml.converter import convert as _latex_to_mathml
from mathml2omml.op_dict import OP_DICT

MAX_LATEX_LENGTH = 6000

_FORBIDDEN_LATEX_COMMANDS = re.compile(
    r"\\(?:input|include|write|openout|read|def|gdef|edef|xdef|newcommand|renewcommand"
    r"|providecommand|DeclareMathOperator|usepackage|RequirePackage|href|url|catcode"
    r"|csname|endcsname|immediate|write18|loop|repeat)\b"
)

# Commands that take under/over limits by default in display math unless \nolimits is given.
_DEFAULT_LIMIT_COMMANDS = re.compile(
    r"(\\(?:limsup|liminf|varlimsup|varliminf|injlim|projlim|varinjlim|varprojlim"
    r"|lim|max|min|sup|inf|det|gcd|Pr)(?![a-zA-Z]))"
    r"(?!\s*\\(?:limits|nolimits)(?![a-zA-Z]))(?=\s*[_^])"
)

# Environments that represent multi-line aligned equations (`m:eqArr`).
_ALIGNED_ENV_START = re.compile(r"\\begin\{(?:aligned|align\*?|split|gathered|gather\*?|eqnarray\*?)\}")
_ALIGNED_ENV_END = re.compile(r"\\end\{(?:aligned|align\*?|split|gathered|gather\*?|eqnarray\*?)\}")

# Known mathematical function names (emitted as <mi> or <mo> by latex2mathml).
_STANDARD_FUNCTIONS = frozenset({
    "arccos", "arcsin", "arctan", "arg", "cos", "cosh", "cot", "coth", "csc",
    "deg", "dim", "exp", "hom", "ker", "lg", "ln", "log", "sec", "sin", "sinh",
    "tan", "tanh", "arccot", "arcsec", "arccsc", "sech", "csch", "arsinh",
    "arcosh", "artanh",
})

_LIMIT_FUNCTIONS = frozenset({
    "lim", "sup", "inf", "max", "min", "det", "gcd", "Pr",
    "lim\u2006inf", "lim\u2006sup", "limsup", "liminf",
    "arg\u2009max", "arg\u2009min", "argmax", "argmin",
    "inj\u2006lim", "proj\u2006lim", "plim",
})

_NARY_CHARS = frozenset({
    "\u222b", "\u222c", "\u222d", "\u222e", "\u222f", "\u2230", "\u2231", "\u2232", "\u2233",
    "\u2211", "\u220f", "\u2210", "\u22c0", "\u22c1", "\u22c2", "\u22c3",
    "\u2a00", "\u2a01", "\u2a02", "\u2a04", "\u2a06",
})

_INTEGRAL_CHARS = frozenset({
    "\u222b", "\u222c", "\u222d", "\u222e", "\u222f", "\u2230", "\u2231", "\u2232", "\u2233",
})

# Top-level relation operators that terminate an n-ary body or unparenthesized function argument.
_RELATION_CHARS = frozenset({
    "=", "<", ">", "\u2264", "\u2265", "\u2260", "\u2248", "\u2261", "\u223c", "\u2243",
    "\u2245", "\u221d", "\u2192", "\u2190", "\u2194", "\u21d2", "\u21d0", "\u21d4", "\u21a6",
    "\u2208", "\u2209", "\u2282", "\u2286", "\u2283", "\u2287", "\u226a", "\u226b", "\u2a7d",
    "\u2a7e", "\u2254", "\u225c", "\u226e", "\u226f", "\u2270", "\u2271", "\u22a2", "\u22a8",
})

_ADDITIVE_CHARS = frozenset({
    "+", "-", "\u2212", "\u00b1", "\u2213",
})

_SEPARATOR_CHARS = frozenset({",", ";", ":"})

_OPEN_FENCES = frozenset({"(", "[", "{", "\u27e8", "\u2308", "\u230a", "\u2016", "\u2225"})
_CLOSE_FENCES = frozenset({")", "]", "}", "\u27e9", "\u2309", "\u230b", "\u2016", "\u2225"})

# Combining accent characters preferred by Word OMML <m:accPr><m:chr>.
_OVER_ACCENT_MAP = {
    "^": "\u0302",
    "\u005e": "\u0302",
    "\u02c6": "\u0302",
    "~": "\u0303",
    "\u007e": "\u0303",
    "\u02dc": "\u0303",
    "\u02d9": "\u0307",
    "\u00a8": "\u0308",
    "\u20db": "\u20db",
    "\u20dc": "\u20dc",
    "\u00b4": "\u0301",
    "`": "\u0300",
    "\u0060": "\u0300",
    "\u02d8": "\u0306",
    "\u02c7": "\u030c",
    "\u02da": "\u030a",
    "\u2192": "\u20d7",
    "\u2190": "\u20d6",
    "\u2194": "\u20e1",
}

_OVER_BAR_CHARS = frozenset({"\u00af", "\u2015", "\u0304", "\u0305", "_"})
_UNDER_BAR_CHARS = frozenset({"\u2015", "\u00af", "_", "\u0332"})

_OVER_GROUP_CHARS = frozenset({"\u23de", "\u23b4", "\u23dc"})
_UNDER_GROUP_CHARS = frozenset({"\u23df", "\u23b5", "\u23dd", "\u2192", "\u2190", "\u2194"})

_INVISIBLE_CHARS_RE = re.compile("[\u2061\u2062\u2063\u2064]")


def strip_outer_math_delimiters(latex: str) -> str:
    """Strip outer math delimiters only when they wrap the entire expression."""
    source = latex.strip()
    if not source:
        return ""
    for start, end in (("$$", "$$"), ("\\[", "\\]"), ("\\(", "\\)")):
        if source.startswith(start) and source.endswith(end) and len(source) >= len(start) + len(end):
            inner = source[len(start):-len(end)].strip()
            if start not in inner and end not in inner:
                return inner
            return source
    if source.startswith("$") and source.endswith("$") and not source.startswith("$$") and len(source) >= 2:
        inner = source[1:-1].strip()
        # Only strip if there is no unescaped $ inside inner.
        if not re.search(r"(?<!\\)\$", inner):
            return inner
    return source


def normalize_latex_for_mathml(latex: str) -> str:
    """Apply bounded, semantics-preserving normalizations before latex2mathml."""
    source = strip_outer_math_delimiters(latex)
    if not source or len(source) > MAX_LATEX_LENGTH:
        raise ValueError("Empty or excessively long LaTeX input")
    if _FORBIDDEN_LATEX_COMMANDS.search(source):
        raise ValueError("Forbidden LaTeX command")

    # Normalize aligned/gathered environments to split so latex2mathml parses & and \\ into mtable.
    start_matches = list(_ALIGNED_ENV_START.finditer(source))
    end_matches = list(_ALIGNED_ENV_END.finditer(source))
    if len(start_matches) != len(end_matches):
        raise ValueError("Unbalanced aligned environment")
    if start_matches:
        source = _ALIGNED_ENV_START.sub(r"\\begin{split}", source)
        source = _ALIGNED_ENV_END.sub(r"\\end{split}", source)

    # Preserve standard under/over limit semantics on \lim, \max, \min, \sup, \inf, etc.
    source = _DEFAULT_LIMIT_COMMANDS.sub(r"\1\\limits", source)
    return source


def _local_name(elem) -> str:
    tag = elem.tag
    if isinstance(tag, str) and "}" in tag:
        return tag.split("}", 1)[1]
    return str(tag)


def _elem_text(elem) -> str:
    return _INVISIBLE_CHARS_RE.sub("", "".join(elem.itertext()))


def _set_cambria_math(run_el):
    rpr = OxmlElement("w:rPr")
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:ascii"), "Cambria Math")
    rfonts.set(qn("w:hAnsi"), "Cambria Math")
    rpr.append(rfonts)
    math_rpr = run_el.find(qn("m:rPr"))
    run_el.insert(1 if math_rpr is not None else 0, rpr)


def make_math_run(text: str, *, style: str = "p", script: str | None = None, normal_text: bool = False):
    """Create a properly formatted <m:r> element with Cambria Math font properties."""
    run = OxmlElement("m:r")
    mrpr = OxmlElement("m:rPr")
    if normal_text:
        mrpr.append(OxmlElement("m:nor"))
    if script:
        scr_el = OxmlElement("m:scr")
        scr_el.set(qn("m:val"), script)
        mrpr.append(scr_el)
    if style:
        sty_el = OxmlElement("m:sty")
        sty_el.set(qn("m:val"), style)
        mrpr.append(sty_el)
    run.append(mrpr)
    _set_cambria_math(run)
    mt = OxmlElement("m:t")
    if text != text.strip():
        mt.set(qn("xml:space"), "preserve")
    mt.text = text
    run.append(mt)
    return run


def _script_and_style(attrib: dict, default_style: str) -> tuple[str | None, str]:
    variant = attrib.get("mathvariant")
    if not variant:
        return None, default_style
    script = None
    for candidate in ("script", "fraktur", "sans-serif", "monospace", "double-struck", "roman"):
        if candidate in variant:
            script = candidate
            break
    style_map = {"bold-italic": "bi", "bold": "b", "italic": "i", "normal": "p"}
    style = default_style
    for key, val in style_map.items():
        if key in variant:
            style = val
            break
    return script, style


def _op_metadata(mo_elem, recommended_form: str = "infix") -> dict:
    attrs = {
        "fence": False,
        "separator": False,
        "stretchy": False,
        "largeop": False,
        "movablelimits": False,
        "accent": False,
        "form": recommended_form,
    }
    content = _elem_text(mo_elem).strip()
    for form in (recommended_form, "infix", "prefix", "postfix"):
        dict_attrs = OP_DICT.get((content, form))
        if dict_attrs is not None:
            attrs.update(dict_attrs)
            attrs["form"] = form
            break
    for k, v in mo_elem.attrib.items():
        if v == "true":
            attrs[k] = True
        elif v == "false":
            attrs[k] = False
        else:
            attrs[k] = v
    return attrs


def _is_stretchy_fence(elem, recommended_form: str = "infix") -> bool:
    if _local_name(elem) != "mo":
        return False
    attrs = _op_metadata(elem, recommended_form)
    if attrs.get("fence") and attrs.get("stretchy"):
        return True
    text = _elem_text(elem).strip()
    if text in (_OPEN_FENCES | _CLOSE_FENCES | {"|"}) and ("minsize" in elem.attrib or "maxsize" in elem.attrib):
        return True
    return False


def _is_nary_mo(elem) -> bool:
    if _local_name(elem) != "mo":
        return False
    text = _elem_text(elem).strip()
    if text not in _NARY_CHARS:
        return False
    if elem.attrib.get("largeop") == "false":
        return False
    return True


def _is_nary_node(elem) -> bool:
    name = _local_name(elem)
    if name == "mo":
        return _is_nary_mo(elem)
    if name in ("msub", "msup", "msubsup", "munder", "mover", "munderover") and len(elem) >= 2:
        return _is_nary_mo(elem[0])
    return False


def _is_function_atom(elem) -> bool:
    name = _local_name(elem)
    text = _elem_text(elem).strip()
    if name == "mi" and len(text) > 1 and elem.attrib.get("mathvariant") is None:
        return text in _STANDARD_FUNCTIONS or text in _LIMIT_FUNCTIONS
    if name == "mo":
        if text in _LIMIT_FUNCTIONS or text in _STANDARD_FUNCTIONS:
            return True
        if elem.attrib.get("movablelimits") == "true" and text and text.isalpha():
            return True
    return False


def _is_function_node(elem) -> bool:
    name = _local_name(elem)
    if _is_function_atom(elem):
        return True
    if name in ("msub", "msup", "msubsup", "munder", "mover", "munderover") and len(elem) >= 2:
        return _is_function_atom(elem[0])
    return False


def _unwrap_single_mrow(elem):
    while _local_name(elem) in ("mrow", "mstyle") and len(elem) == 1:
        elem = elem[0]
    return elem


def _base_mo_text(elem) -> str | None:
    name = _local_name(elem)
    if name == "mo":
        return _elem_text(elem).strip()
    if name in ("msup", "msub", "msubsup") and len(elem) >= 1 and _local_name(elem[0]) == "mo":
        return _elem_text(elem[0]).strip()
    return None


def _fence_delta(elem) -> int:
    text = _base_mo_text(elem)
    if not text:
        return 0
    if text in _OPEN_FENCES:
        return 1
    if text in _CLOSE_FENCES:
        return -1
    return 0


def _is_top_level_relation(elem) -> bool:
    if _local_name(elem) != "mo":
        return False
    return _elem_text(elem).strip() in _RELATION_CHARS


def _is_Stopping_for_func_arg(elem) -> bool:
    if _local_name(elem) == "mo":
        text = _elem_text(elem).strip()
        if text in _RELATION_CHARS or text in _ADDITIVE_CHARS or text in _SEPARATOR_CHARS or text in _CLOSE_FENCES:
            return True
    if _is_nary_node(elem):
        return True
    return False


def _group_stretchy_fences(children: list) -> list:
    """Group matching stretchy prefix/postfix fence <mo> elements into synthetic ('fenced', open, close, seps, body) tuples."""
    non_space_indices = [idx for idx, ch in enumerate(children) if _local_name(ch) != "mspace"]
    if not non_space_indices:
        return list(children)
    first_ns = non_space_indices[0]
    last_ns = non_space_indices[-1]

    stack: list[list] = [[]]
    fence_meta: list[tuple[str, list[str]]] = []

    for idx, ch in enumerate(children):
        rec_form = "prefix" if idx == first_ns else ("postfix" if idx == last_ns else "infix")
        if _is_stretchy_fence(ch, rec_form):
            attrs = _op_metadata(ch, rec_form)
            form = attrs.get("form", rec_form)
            text = _elem_text(ch).strip()
            # Handle | or \| closing an already open | or \| fence
            if text in ("|", "\u2016", "\u2225") and fence_meta and fence_meta[-1][0] == text:
                form = "postfix"
            elif form == "infix" and ch.attrib.get("fence") == "true" and fence_meta:
                # \middle| inside \left...\right
                fence_meta[-1][1].append(text)
                stack[-1].append(("middle_sep", text))
                continue
            if form == "prefix" or (form == "infix" and not fence_meta):
                stack.append([])
                fence_meta.append((text, []))
                continue
            if form == "postfix" and fence_meta:
                body = stack.pop()
                open_chr, seps = fence_meta.pop()
                stack[-1].append(("fenced", open_chr, text, seps, body))
                continue
        stack[-1].append(ch)

    while len(stack) > 1:
        body = stack.pop()
        open_chr, seps = fence_meta.pop()
        stack[-1].append(("fenced", open_chr, "", seps, body))

    return stack[0]


def _build_fenced_node(open_chr: str, close_chr: str, seps: list[str], body: list, *, in_textstyle: bool) -> OxmlElement:
    d_el = OxmlElement("m:d")
    dpr = OxmlElement("m:dPr")
    beg = OxmlElement("m:begChr")
    beg.set(qn("m:val"), open_chr)
    dpr.append(beg)
    if seps:
        sep_el = OxmlElement("m:sepChr")
        sep_el.set(qn("m:val"), seps[0])
        dpr.append(sep_el)
    end = OxmlElement("m:endChr")
    end.set(qn("m:val"), close_chr)
    dpr.append(end)
    d_el.append(dpr)

    # Split body by ('middle_sep', ...) if present
    segments: list[list] = [[]]
    for item in body:
        if isinstance(item, tuple) and item[0] == "middle_sep":
            segments.append([])
        else:
            segments[-1].append(item)
    for seg in segments:
        e_el = OxmlElement("m:e")
        for node in _convert_sequence(seg, in_textstyle=in_textstyle):
            e_el.append(node)
        d_el.append(e_el)
    return d_el


def _build_nary_node(nary_elem, body_items: list, *, in_textstyle: bool) -> OxmlElement:
    name = _local_name(nary_elem)
    if name == "mo":
        op_el = nary_elem
        sub_mml = None
        sup_mml = None
        lim_loc = "undOvr" if _elem_text(op_el).strip() not in _INTEGRAL_CHARS and not in_textstyle else "subSup"
    else:
        op_el = nary_elem[0]
        if name in ("munder", "mover", "munderover"):
            lim_loc = "subSup" if in_textstyle else "undOvr"
        else:
            lim_loc = "subSup"
        if name in ("msub", "munder"):
            sub_mml, sup_mml = nary_elem[1], None
        elif name in ("msup", "mover"):
            sub_mml, sup_mml = None, nary_elem[1]
        else:
            sub_mml, sup_mml = nary_elem[1], nary_elem[2]

    chr_val = _elem_text(op_el).strip()
    nary = OxmlElement("m:nary")
    nary_pr = OxmlElement("m:naryPr")
    chr_el = OxmlElement("m:chr")
    chr_el.set(qn("m:val"), chr_val)
    nary_pr.append(chr_el)
    loc_el = OxmlElement("m:limLoc")
    loc_el.set(qn("m:val"), lim_loc)
    nary_pr.append(loc_el)
    if sub_mml is None:
        sub_hide = OxmlElement("m:subHide")
        sub_hide.set(qn("m:val"), "1")
        nary_pr.append(sub_hide)
    if sup_mml is None:
        sup_hide = OxmlElement("m:supHide")
        sup_hide.set(qn("m:val"), "1")
        nary_pr.append(sup_hide)
    nary.append(nary_pr)

    sub_el = OxmlElement("m:sub")
    if sub_mml is not None:
        for node in _convert_node_to_list(sub_mml, in_textstyle=in_textstyle):
            sub_el.append(node)
    nary.append(sub_el)

    sup_el = OxmlElement("m:sup")
    if sup_mml is not None:
        for node in _convert_node_to_list(sup_mml, in_textstyle=in_textstyle):
            sup_el.append(node)
    nary.append(sup_el)

    e_el = OxmlElement("m:e")
    for node in _convert_sequence(body_items, in_textstyle=in_textstyle):
        e_el.append(node)
    nary.append(e_el)
    return nary


def _build_fname_elements(func_elem, *, in_textstyle: bool) -> list[OxmlElement]:
    name = _local_name(func_elem)
    if name in ("mi", "mo"):
        text = _elem_text(func_elem).strip()
        return [make_math_run(text, style="p")]

    base_el = func_elem[0]
    base_text = _elem_text(base_el).strip()
    base_run = make_math_run(base_text, style="p")

    # Check if this should be a limit structure (m:limLow / m:limUpp) or script (m:sSub / m:sSup / m:sSubSup)
    is_movable_sub = (
        name == "msub"
        and _local_name(base_el) == "mo"
        and base_el.attrib.get("movablelimits") == "true"
        and not in_textstyle
    )
    if name == "munder" or is_movable_sub:
        lim_low = OxmlElement("m:limLow")
        e_el = OxmlElement("m:e")
        e_el.append(base_run)
        lim_el = OxmlElement("m:lim")
        for node in _convert_node_to_list(func_elem[1], in_textstyle=in_textstyle):
            lim_el.append(node)
        lim_low.append(e_el)
        lim_low.append(lim_el)
        return [lim_low]

    if name == "mover":
        lim_upp = OxmlElement("m:limUpp")
        e_el = OxmlElement("m:e")
        e_el.append(base_run)
        lim_el = OxmlElement("m:lim")
        for node in _convert_node_to_list(func_elem[1], in_textstyle=in_textstyle):
            lim_el.append(node)
        lim_upp.append(e_el)
        lim_upp.append(lim_el)
        return [lim_upp]

    if name == "munderover":
        lim_upp = OxmlElement("m:limUpp")
        e_upp = OxmlElement("m:e")
        e_upp.append(base_run)
        lim_u = OxmlElement("m:lim")
        for node in _convert_node_to_list(func_elem[2], in_textstyle=in_textstyle):
            lim_u.append(node)
        lim_upp.append(e_upp)
        lim_upp.append(lim_u)

        lim_low = OxmlElement("m:limLow")
        e_low = OxmlElement("m:e")
        e_low.append(lim_upp)
        lim_l = OxmlElement("m:lim")
        for node in _convert_node_to_list(func_elem[1], in_textstyle=in_textstyle):
            lim_l.append(node)
        lim_low.append(e_low)
        lim_low.append(lim_l)
        return [lim_low]

    if name == "msub":
        ssub = OxmlElement("m:sSub")
        e_el = OxmlElement("m:e")
        e_el.append(base_run)
        sub_el = OxmlElement("m:sub")
        for node in _convert_node_to_list(func_elem[1], in_textstyle=in_textstyle):
            sub_el.append(node)
        ssub.append(e_el)
        ssub.append(sub_el)
        return [ssub]

    if name == "msup":
        ssup = OxmlElement("m:sSup")
        e_el = OxmlElement("m:e")
        e_el.append(base_run)
        sup_el = OxmlElement("m:sup")
        for node in _convert_node_to_list(func_elem[1], in_textstyle=in_textstyle):
            sup_el.append(node)
        ssup.append(e_el)
        ssup.append(sup_el)
        return [ssup]

    if name == "msubsup":
        ssubsup = OxmlElement("m:sSubSup")
        e_el = OxmlElement("m:e")
        e_el.append(base_run)
        sub_el = OxmlElement("m:sub")
        for node in _convert_node_to_list(func_elem[1], in_textstyle=in_textstyle):
            sub_el.append(node)
        sup_el = OxmlElement("m:sup")
        for node in _convert_node_to_list(func_elem[2], in_textstyle=in_textstyle):
            sup_el.append(node)
        ssubsup.append(e_el)
        ssubsup.append(sub_el)
        ssubsup.append(sup_el)
        return [ssubsup]

    return [base_run]


def _consume_paren_group(items: list, start_idx: int) -> int:
    """Given items[start_idx] is an opening fence <mo>, return the exclusive end index after matching close fence."""
    depth = 0
    idx = start_idx
    while idx < len(items):
        item = items[idx]
        if not isinstance(item, tuple):
            delta = _fence_delta(item)
            depth += delta
            if depth <= 0 and idx > start_idx:
                return idx + 1
        idx += 1
    return len(items)


def _consume_function_arg(items: list, start_idx: int) -> int:
    """Determine the slice items[start_idx:end_idx] that forms the argument of a function."""
    idx = start_idx
    while idx < len(items) and not isinstance(items[idx], tuple) and _local_name(items[idx]) == "mspace":
        idx += 1
    if idx >= len(items):
        return start_idx

    first = items[idx]
    if isinstance(first, tuple) and first[0] == "fenced":
        return idx + 1

    if not isinstance(first, tuple):
        if _is_Stopping_for_func_arg(first):
            # Allow unary +/- immediately after function name, e.g. \sin(-x) or \sin -x
            if _local_name(first) == "mo" and _elem_text(first).strip() in ("+", "-", "\u2212"):
                idx += 1
                if idx >= len(items):
                    return start_idx
                first = items[idx]
            else:
                return start_idx

        # If argument starts with another function (e.g. \sin \cos x or \lim_{x\to 0} \sin x),
        # recursively include that function and its argument.
        if _is_function_node(first):
            sub_end = _consume_function_arg(items, idx + 1)
            return max(idx + 1, sub_end)

        # If argument starts with an explicit opening paren/bracket, consume up to matching close.
        if _fence_delta(first) > 0:
            return _consume_paren_group(items, idx)

    # Otherwise consume adjacent multiplicative atoms (stopping at +, -, =, comma, close fence, n-ary, etc.)
    # Also stop before differential d<var> if another atom was already consumed.
    atoms_consumed = 0
    while idx < len(items):
        cur = items[idx]
        if isinstance(cur, tuple):
            if cur[0] == "fenced":
                idx += 1
                atoms_consumed += 1
                continue
            break
        lname = _local_name(cur)
        if lname == "mspace":
            idx += 1
            continue
        if _is_Stopping_for_func_arg(cur) or _is_function_node(cur):
            break
        # If an atom is immediately followed by ( ... ), e.g. f(x) in \lim f(x), consume the ( ... ) as part of f(x).
        if _fence_delta(cur) > 0:
            if atoms_consumed > 0:
                idx = _consume_paren_group(items, idx)
            break
        # Stop before differential `dx` if we already consumed the main argument atom (e.g. \int \sin x\,dx)
        if (
            atoms_consumed > 0
            and lname == "mi"
            and _elem_text(cur).strip() == "d"
            and idx + 1 < len(items)
            and not isinstance(items[idx + 1], tuple)
            and _local_name(items[idx + 1]) == "mi"
        ):
            break
        idx += 1
        atoms_consumed += 1

    return idx


def _has_differential_in_subtree(elem) -> bool:
    """Return True if an element (e.g. <mfrac>dx/x</mfrac>) contains differential d<var>."""
    for parent in elem.iter():
        children = [ch for ch in parent if _local_name(ch) != "mspace"]
        for i, ch in enumerate(children[:-1]):
            if (
                _local_name(ch) == "mi"
                and _elem_text(ch).strip() == "d"
                and _local_name(children[i + 1]) in ("mi", "msub", "msup", "msubsup")
            ):
                return True
    return False


def _consume_nary_body(items: list, start_idx: int, *, is_integral: bool) -> int:
    """Determine the slice items[start_idx:end_idx] that forms the body <m:e> of an <m:nary> operator."""
    idx = start_idx
    fence_depth = 0
    saw_differential = False
    waiting_for_operand = True
    while idx < len(items):
        cur = items[idx]
        if isinstance(cur, tuple):
            if fence_depth == 0:
                waiting_for_operand = False
            idx += 1
            continue
        lname = _local_name(cur)
        if lname == "mspace":
            idx += 1
            continue
        if fence_depth == 0:
            if _is_top_level_relation(cur):
                break
            if lname == "mo" and _elem_text(cur).strip() in _SEPARATOR_CHARS:
                break
            if is_integral:
                if saw_differential and lname == "mo" and _elem_text(cur).strip() in _ADDITIVE_CHARS:
                    break
            else:
                if not waiting_for_operand and lname == "mo" and _elem_text(cur).strip() in _ADDITIVE_CHARS:
                    break

        delta = _fence_delta(cur)
        fence_depth = max(0, fence_depth + delta)
        if fence_depth == 0:
            if _is_nary_node(cur):
                op_el = cur if lname == "mo" else cur[0]
                if _elem_text(op_el).strip() in _INTEGRAL_CHARS:
                    idx = _consume_nary_body(items, idx + 1, is_integral=True)
                    waiting_for_operand = False
                    continue
                waiting_for_operand = True
            elif _is_function_node(cur):
                arg_end = _consume_function_arg(items, idx + 1)
                idx = max(idx + 1, arg_end)
                waiting_for_operand = False
                continue
            elif lname == "mo" and _elem_text(cur).strip() in _ADDITIVE_CHARS and waiting_for_operand:
                pass
            else:
                waiting_for_operand = False

        if is_integral and fence_depth == 0:
            if (
                lname == "mi"
                and _elem_text(cur).strip() == "d"
                and idx + 1 < len(items)
                and not isinstance(items[idx + 1], tuple)
                and _local_name(items[idx + 1]) in ("mi", "msub", "msup", "msubsup")
            ):
                saw_differential = True
            elif lname == "mfrac" and _has_differential_in_subtree(cur):
                saw_differential = True
        idx += 1
    return idx


def _convert_sequence(raw_children: list, *, in_textstyle: bool = False) -> list[OxmlElement]:
    items = _group_stretchy_fences(raw_children)
    result: list[OxmlElement] = []
    idx = 0
    while idx < len(items):
        item = items[idx]
        if isinstance(item, tuple):
            if item[0] == "fenced":
                _, open_chr, close_chr, seps, body = item
                result.append(_build_fenced_node(open_chr, close_chr, seps, body, in_textstyle=in_textstyle))
            idx += 1
            continue

        if _is_nary_node(item):
            op_el = item if _local_name(item) == "mo" else item[0]
            is_int = _elem_text(op_el).strip() in _INTEGRAL_CHARS
            body_end = _consume_nary_body(items, idx + 1, is_integral=is_int)
            body_slice = items[idx + 1:body_end]
            result.append(_build_nary_node(item, body_slice, in_textstyle=in_textstyle))
            idx = body_end
            continue

        if _is_function_node(item):
            arg_end = _consume_function_arg(items, idx + 1)
            fname_nodes = _build_fname_elements(item, in_textstyle=in_textstyle)
            if arg_end > idx + 1:
                func_el = OxmlElement("m:func")
                fname_el = OxmlElement("m:fName")
                for fn in fname_nodes:
                    fname_el.append(fn)
                e_el = OxmlElement("m:e")
                for an in _convert_sequence(items[idx + 1:arg_end], in_textstyle=in_textstyle):
                    e_el.append(an)
                func_el.append(fname_el)
                func_el.append(e_el)
                result.append(func_el)
                idx = arg_end
                continue
            result.extend(fname_nodes)
            idx += 1
            continue

        result.extend(_convert_node_to_list(item, in_textstyle=in_textstyle))
        idx += 1
    return result


def _convert_mover(elem, *, in_textstyle: bool) -> OxmlElement:
    base, over = elem[0], _unwrap_single_mrow(elem[1])
    if _local_name(over) == "mo":
        over_chr = _elem_text(over).strip()
        if over_chr in _OVER_BAR_CHARS:
            bar = OxmlElement("m:bar")
            bar_pr = OxmlElement("m:barPr")
            pos = OxmlElement("m:pos")
            pos.set(qn("m:val"), "top")
            bar_pr.append(pos)
            bar.append(bar_pr)
            e_el = OxmlElement("m:e")
            for node in _convert_node_to_list(base, in_textstyle=in_textstyle):
                e_el.append(node)
            bar.append(e_el)
            return bar

        if over_chr in _OVER_GROUP_CHARS:
            grp = OxmlElement("m:groupChr")
            grp_pr = OxmlElement("m:groupChrPr")
            chr_el = OxmlElement("m:chr")
            chr_el.set(qn("m:val"), over_chr)
            pos_el = OxmlElement("m:pos")
            pos_el.set(qn("m:val"), "top")
            vert_jc = OxmlElement("m:vertJc")
            vert_jc.set(qn("m:val"), "bot")
            grp_pr.append(chr_el)
            grp_pr.append(pos_el)
            grp_pr.append(vert_jc)
            grp.append(grp_pr)
            e_el = OxmlElement("m:e")
            for node in _convert_node_to_list(base, in_textstyle=in_textstyle):
                e_el.append(node)
            grp.append(e_el)
            return grp

        if over_chr in _OVER_ACCENT_MAP:
            acc = OxmlElement("m:acc")
            acc_pr = OxmlElement("m:accPr")
            chr_el = OxmlElement("m:chr")
            chr_el.set(qn("m:val"), _OVER_ACCENT_MAP[over_chr])
            acc_pr.append(chr_el)
            acc.append(acc_pr)
            e_el = OxmlElement("m:e")
            for node in _convert_node_to_list(base, in_textstyle=in_textstyle):
                e_el.append(node)
            acc.append(e_el)
            return acc

    lim_upp = OxmlElement("m:limUpp")
    e_el = OxmlElement("m:e")
    for node in _convert_node_to_list(base, in_textstyle=in_textstyle):
        e_el.append(node)
    lim_el = OxmlElement("m:lim")
    for node in _convert_node_to_list(elem[1], in_textstyle=in_textstyle):
        lim_el.append(node)
    lim_upp.append(e_el)
    lim_upp.append(lim_el)
    return lim_upp


def _convert_munder(elem, *, in_textstyle: bool) -> OxmlElement:
    base, under = elem[0], _unwrap_single_mrow(elem[1])
    if _local_name(under) == "mo":
        under_chr = _elem_text(under).strip()
        if under_chr in _UNDER_BAR_CHARS:
            bar = OxmlElement("m:bar")
            bar_pr = OxmlElement("m:barPr")
            pos = OxmlElement("m:pos")
            pos.set(qn("m:val"), "bot")
            bar_pr.append(pos)
            bar.append(bar_pr)
            e_el = OxmlElement("m:e")
            for node in _convert_node_to_list(base, in_textstyle=in_textstyle):
                e_el.append(node)
            bar.append(e_el)
            return bar

        if under_chr in _UNDER_GROUP_CHARS:
            grp = OxmlElement("m:groupChr")
            grp_pr = OxmlElement("m:groupChrPr")
            chr_el = OxmlElement("m:chr")
            chr_el.set(qn("m:val"), under_chr)
            pos_el = OxmlElement("m:pos")
            pos_el.set(qn("m:val"), "bot")
            vert_jc = OxmlElement("m:vertJc")
            vert_jc.set(qn("m:val"), "top")
            grp_pr.append(chr_el)
            grp_pr.append(pos_el)
            grp_pr.append(vert_jc)
            grp.append(grp_pr)
            e_el = OxmlElement("m:e")
            for node in _convert_node_to_list(base, in_textstyle=in_textstyle):
                e_el.append(node)
            grp.append(e_el)
            return grp

    lim_low = OxmlElement("m:limLow")
    e_el = OxmlElement("m:e")
    for node in _convert_node_to_list(base, in_textstyle=in_textstyle):
        e_el.append(node)
    lim_el = OxmlElement("m:lim")
    for node in _convert_node_to_list(elem[1], in_textstyle=in_textstyle):
        lim_el.append(node)
    lim_low.append(e_el)
    lim_low.append(lim_el)
    return lim_low


def _convert_mtable(elem, *, in_textstyle: bool) -> OxmlElement:
    rows = [ch for ch in elem if _local_name(ch) == "mtr"]
    # Aligned equation environments set displaystyle="true" on <mtable> in latex2mathml.
    if elem.attrib.get("displaystyle") == "true":
        eq_arr = OxmlElement("m:eqArr")
        for row in rows:
            e_el = OxmlElement("m:e")
            cells = [c for c in row if _local_name(c) == "mtd"]
            row_children = []
            for cell_idx, cell in enumerate(cells):
                if cell_idx >= 2 and cell_idx % 2 == 0:
                    # Add spacing between separate alignment column pairs
                    mspace = OxmlElement("mspace")
                    mspace.set("width", "1em")
                    row_children.append(mspace)
                for ch in cell:
                    # Skip empty dummy <mi/> inserted by latex2mathml before alignment points
                    if _local_name(ch) == "mi" and not _elem_text(ch).strip():
                        continue
                    row_children.append(ch)
            for node in _convert_sequence(row_children, in_textstyle=in_textstyle):
                e_el.append(node)
            eq_arr.append(e_el)
        return eq_arr

    matrix = OxmlElement("m:m")
    max_cols = max((len([c for c in r if _local_name(c) == "mtd"]) for r in rows), default=0)
    for row in rows:
        mr = OxmlElement("m:mr")
        cells = [c for c in row if _local_name(c) == "mtd"]
        for cell in cells:
            e_el = OxmlElement("m:e")
            for node in _convert_sequence(list(cell), in_textstyle=in_textstyle):
                e_el.append(node)
            mr.append(e_el)
        for _ in range(max_cols - len(cells)):
            mr.append(OxmlElement("m:e"))
        matrix.append(mr)
    return matrix


def _convert_node_to_list(elem, *, in_textstyle: bool = False) -> list[OxmlElement]:
    name = _local_name(elem)

    if name in ("math", "mrow", "mpadded", "semantics"):
        return _convert_sequence(list(elem), in_textstyle=in_textstyle)

    if name == "mstyle":
        child_textstyle = in_textstyle
        if elem.attrib.get("displaystyle") == "false":
            child_textstyle = True
        elif elem.attrib.get("displaystyle") == "true":
            child_textstyle = False
        return _convert_sequence(list(elem), in_textstyle=child_textstyle)

    if name in ("annotation", "annotation-xml", "none", "mprescripts"):
        return []

    if name == "mspace":
        width = elem.attrib.get("width", "")
        if width.startswith(("1em", "2em")):
            return [make_math_run(" ", style="p", normal_text=True)]
        return []

    if name == "mi":
        text = _elem_text(elem)
        if not text:
            return []
        default_sty = "i" if len(text.strip()) == 1 else "p"
        scr, sty = _script_and_style(elem.attrib, default_sty)
        return [make_math_run(text, style=sty, script=scr)]

    if name in ("mn", "mo"):
        text = _elem_text(elem)
        if not text:
            return []
        scr, sty = _script_and_style(elem.attrib, "p")
        return [make_math_run(text, style=sty, script=scr)]

    if name in ("mtext", "ms"):
        text = _elem_text(elem)
        if name == "ms":
            lquote = elem.attrib.get("lquote", '"')
            rquote = elem.attrib.get("rquote", '"')
            text = f"{lquote}{text}{rquote}"
        if not text:
            return []
        variant = elem.attrib.get("mathvariant")
        if variant is None or variant == "normal":
            return [make_math_run(text, style="p", normal_text=True)]
        scr, sty = _script_and_style(elem.attrib, "p")
        return [make_math_run(text, style=sty, script=scr)]

    if name == "mfrac":
        f_el = OxmlElement("m:f")
        thickness = (elem.attrib.get("linethickness") or "").strip()
        if thickness == "0" or re.match(r"^0(?:\.0+)?[a-zA-Z%]*$", thickness):
            fpr = OxmlElement("m:fPr")
            ftype = OxmlElement("m:type")
            ftype.set(qn("m:val"), "noBar")
            fpr.append(ftype)
            f_el.append(fpr)
        elif elem.attrib.get("bevelled") == "true":
            fpr = OxmlElement("m:fPr")
            ftype = OxmlElement("m:type")
            ftype.set(qn("m:val"), "skw")
            fpr.append(ftype)
            f_el.append(fpr)
        num_el = OxmlElement("m:num")
        for node in _convert_node_to_list(elem[0], in_textstyle=in_textstyle):
            num_el.append(node)
        den_el = OxmlElement("m:den")
        for node in _convert_node_to_list(elem[1], in_textstyle=in_textstyle):
            den_el.append(node)
        f_el.append(num_el)
        f_el.append(den_el)
        return [f_el]

    if name == "msqrt":
        rad = OxmlElement("m:rad")
        rad_pr = OxmlElement("m:radPr")
        deg_hide = OxmlElement("m:degHide")
        deg_hide.set(qn("m:val"), "1")
        rad_pr.append(deg_hide)
        rad.append(rad_pr)
        rad.append(OxmlElement("m:deg"))
        e_el = OxmlElement("m:e")
        for node in _convert_sequence(list(elem), in_textstyle=in_textstyle):
            e_el.append(node)
        rad.append(e_el)
        return [rad]

    if name == "mroot":
        rad = OxmlElement("m:rad")
        rad_pr = OxmlElement("m:radPr")
        deg_hide = OxmlElement("m:degHide")
        deg_hide.set(qn("m:val"), "0")
        rad_pr.append(deg_hide)
        rad.append(rad_pr)
        deg_el = OxmlElement("m:deg")
        for node in _convert_node_to_list(elem[1], in_textstyle=in_textstyle):
            deg_el.append(node)
        rad.append(deg_el)
        e_el = OxmlElement("m:e")
        for node in _convert_node_to_list(elem[0], in_textstyle=in_textstyle):
            e_el.append(node)
        rad.append(e_el)
        return [rad]

    if name == "msub":
        if _is_nary_node(elem):
            return [_build_nary_node(elem, [], in_textstyle=in_textstyle)]
        ssub = OxmlElement("m:sSub")
        e_el = OxmlElement("m:e")
        for node in _convert_node_to_list(elem[0], in_textstyle=in_textstyle):
            e_el.append(node)
        sub_el = OxmlElement("m:sub")
        for node in _convert_node_to_list(elem[1], in_textstyle=in_textstyle):
            sub_el.append(node)
        ssub.append(e_el)
        ssub.append(sub_el)
        return [ssub]

    if name == "msup":
        if _is_nary_node(elem):
            return [_build_nary_node(elem, [], in_textstyle=in_textstyle)]
        ssup = OxmlElement("m:sSup")
        e_el = OxmlElement("m:e")
        for node in _convert_node_to_list(elem[0], in_textstyle=in_textstyle):
            e_el.append(node)
        sup_el = OxmlElement("m:sup")
        for node in _convert_node_to_list(elem[1], in_textstyle=in_textstyle):
            sup_el.append(node)
        ssup.append(e_el)
        ssup.append(sup_el)
        return [ssup]

    if name == "msubsup":
        if _is_nary_node(elem):
            return [_build_nary_node(elem, [], in_textstyle=in_textstyle)]
        ssubsup = OxmlElement("m:sSubSup")
        e_el = OxmlElement("m:e")
        for node in _convert_node_to_list(elem[0], in_textstyle=in_textstyle):
            e_el.append(node)
        sub_el = OxmlElement("m:sub")
        for node in _convert_node_to_list(elem[1], in_textstyle=in_textstyle):
            sub_el.append(node)
        sup_el = OxmlElement("m:sup")
        for node in _convert_node_to_list(elem[2], in_textstyle=in_textstyle):
            sup_el.append(node)
        ssubsup.append(e_el)
        ssubsup.append(sub_el)
        ssubsup.append(sup_el)
        return [ssubsup]

    if name == "mover":
        if _is_nary_node(elem):
            return [_build_nary_node(elem, [], in_textstyle=in_textstyle)]
        return [_convert_mover(elem, in_textstyle=in_textstyle)]

    if name == "munder":
        if _is_nary_node(elem):
            return [_build_nary_node(elem, [], in_textstyle=in_textstyle)]
        return [_convert_munder(elem, in_textstyle=in_textstyle)]

    if name == "munderover":
        if _is_nary_node(elem):
            return [_build_nary_node(elem, [], in_textstyle=in_textstyle)]
        lim_upp = OxmlElement("m:limUpp")
        e_upp = OxmlElement("m:e")
        for node in _convert_node_to_list(elem[0], in_textstyle=in_textstyle):
            e_upp.append(node)
        lim_u = OxmlElement("m:lim")
        for node in _convert_node_to_list(elem[2], in_textstyle=in_textstyle):
            lim_u.append(node)
        lim_upp.append(e_upp)
        lim_upp.append(lim_u)

        lim_low = OxmlElement("m:limLow")
        e_low = OxmlElement("m:e")
        e_low.append(lim_upp)
        lim_l = OxmlElement("m:lim")
        for node in _convert_node_to_list(elem[1], in_textstyle=in_textstyle):
            lim_l.append(node)
        lim_low.append(e_low)
        lim_low.append(lim_l)
        return [lim_low]

    if name == "mtable":
        return [_convert_mtable(elem, in_textstyle=in_textstyle)]

    if name == "mphantom":
        phant = OxmlElement("m:phant")
        phant_pr = OxmlElement("m:phantPr")
        show = OxmlElement("m:show")
        show.set(qn("m:val"), "0")
        phant_pr.append(show)
        phant.append(phant_pr)
        e_el = OxmlElement("m:e")
        for node in _convert_sequence(list(elem), in_textstyle=in_textstyle):
            e_el.append(node)
        phant.append(e_el)
        return [phant]

    if name == "menclose":
        bbox = OxmlElement("m:borderBox")
        e_el = OxmlElement("m:e")
        for node in _convert_sequence(list(elem), in_textstyle=in_textstyle):
            e_el.append(node)
        bbox.append(e_el)
        return [bbox]

    if name == "mfenced":
        open_chr = elem.attrib.get("open", "(")
        close_chr = elem.attrib.get("close", ")")
        seps = list(elem.attrib.get("separators", ",").replace(" ", ""))
        d_el = OxmlElement("m:d")
        dpr = OxmlElement("m:dPr")
        beg = OxmlElement("m:begChr")
        beg.set(qn("m:val"), open_chr)
        dpr.append(beg)
        if seps:
            sep_el = OxmlElement("m:sepChr")
            sep_el.set(qn("m:val"), seps[0])
            dpr.append(sep_el)
        end = OxmlElement("m:endChr")
        end.set(qn("m:val"), close_chr)
        dpr.append(end)
        d_el.append(dpr)
        for ch in elem:
            e_el = OxmlElement("m:e")
            for node in _convert_node_to_list(ch, in_textstyle=in_textstyle):
                e_el.append(node)
            d_el.append(e_el)
        return [d_el]

    raise ValueError(f"Unsupported MathML element: {name}")


def latex_to_omml_element(latex: str) -> OxmlElement:
    """Convert a bounded LaTeX math string into an `<m:oMath>` OxmlElement."""
    source = normalize_latex_for_mathml(latex)
    mathml_str = _latex_to_mathml(source, display="block")
    mathml_root = parse_xml(mathml_str.encode("utf-8"))
    # Reject unparsed backslash commands left in MathML text (allowing literal \setminus if converted to unicode).
    for elem in mathml_root.iter():
        if elem.text and "\\" in elem.text and _local_name(elem) not in ("mtext", "ms"):
            raise ValueError(f"Unsupported LaTeX command in {elem.text!r}")

    omath = OxmlElement("m:oMath")
    children = _convert_node_to_list(mathml_root, in_textstyle=False)
    for child in children:
        omath.append(child)

    if not list(omath.iter(qn("m:t"))):
        raise ValueError("Empty Office Math expression")
    return omath


def wrap_display_omath(omath_nodes: list[OxmlElement], alignment: str = "center") -> OxmlElement:
    """Wrap one or more `<m:oMath>` elements into a display `<m:oMathPara>` element."""
    omath_para = OxmlElement("m:oMathPara")
    para_pr = OxmlElement("m:oMathParaPr")
    jc = OxmlElement("m:jc")
    jc_map = {
        "left": "left",
        "center": "center",
        "right": "right",
        "justify": "centerGroup",
    }
    jc.set(qn("m:val"), jc_map.get(alignment, "center"))
    para_pr.append(jc)
    omath_para.append(para_pr)
    for node in omath_nodes:
        omath_para.append(node)
    return omath_para
