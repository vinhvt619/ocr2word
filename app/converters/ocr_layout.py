from __future__ import annotations

from difflib import SequenceMatcher
from statistics import median
import re
import unicodedata
from typing import Any, Iterable

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

from app.services.ocr import image_to_data

from .base import ConversionOptions, set_run_font

try:
    import cv2
    import numpy as np
except ImportError:  # pragma: no cover - dependencies are diagnosed elsewhere
    cv2 = None
    np = None


_JOINED_WORD_FIXES = {
    "lầnbanhành": "Lần ban hành",
    "nhàcung": "Nhà cung",
    "mãsố": "Mã số",
    "ngàyhóa": "Ngày hóa",
    "nộiđung": "Nội dung",
    "nội đung": "Nội dung",
    "sốhóa": "Số hóa",
    "sốlượng": "Số lượng",
    "đơngiá": "Đơn giá",
    "thànhtiền": "Thành tiền",
    "tgđphụtrách": "TGĐ phụ trách",
    "môtả": "Mô tả",
}

_SHORT_LINE_ALLOWLIST = {
    "bp",
    "btc",
    "có",
    "do",
    "đã",
    "đề",
    "dva",
    "họ",
    "iv",
    "ký",
    "lý",
    "mã",
    "số",
    "stt",
    "ktt",
    "kttt",
    "ncc",
    "ptgđ",
    "tgđ",
    "tên",
    "từ",
    "tt",
    "tm",
    "unc",
    "và",
    "vat",
    "vt",
    "vnd",
}


def normalize_ocr_text(text: str) -> str:
    """Remove grid artifacts and fix common Vietnamese OCR joins."""

    value = unicodedata.normalize("NFC", str(text or ""))
    value = re.sub(r"\(\s*v\s*(?:[/\\]\s*v|a)\s*:", "(v/v:", value, flags=re.IGNORECASE)
    value = re.sub(r"\bTHƯ\s+CAM\s+KÉT\b", "THƯ CAM KẾT", value, flags=re.IGNORECASE)
    value = re.sub(r"\bQúy\b", "Quý", value)
    value = re.sub(r"\bLâp\b", "Lập", value)
    value = re.sub(r"\bđên\b", "đến", value, flags=re.IGNORECASE)
    value = re.sub(r"\bđông\s+hành\b", "đồng hành", value, flags=re.IGNORECASE)
    value = re.sub(r"\bsuôt\b", "suốt", value, flags=re.IGNORECASE)
    value = re.sub(r"\bthay\s+đồi\b", "thay đổi", value, flags=re.IGNORECASE)
    value = re.sub(r"\btiền\s+độ\b", "tiến độ", value, flags=re.IGNORECASE)
    value = re.sub(r"\bđề\s+(?=cùng|sớm)", "để ", value, flags=re.IGNORECASE)
    value = re.sub(r"\btriên\s+khai\b", "triển khai", value, flags=re.IGNORECASE)
    value = re.sub(r"\bT\s+hắng\b", "Thắng", value)
    value = re.sub(r"&(?=\S)", "& ", value)
    value = re.sub(r"\bhiện\s+nay\.\s+để\b", "hiện nay, để", value, flags=re.IGNORECASE)
    value = re.sub(
        r"\btiến\s+độ\.\s+(?=công\s+ty\b)",
        "tiến độ, ",
        value,
        flags=re.IGNORECASE,
    )
    value = re.sub(r"\bcám\s+ơn\b", "cảm ơn", value, flags=re.IGNORECASE)
    value = re.sub(r"^Trân\s+trọng\.$", "Trân trọng,", value, flags=re.IGNORECASE)
    value = re.sub(
        r"\bGI[ÁẦA]Y\s+(?:Đ[ÈỀ]\s+)?NGHỊ\s+THANH\s+TO[ÁA]N\b",
        "GIẤY ĐỀ NGHỊ THANH TOÁN",
        value,
        flags=re.IGNORECASE,
    )
    value = re.sub(r"\bB[ẰĂA]NG\s+K[ỆÊE]\b", "BẢNG KÊ", value, flags=re.IGNORECASE)
    value = re.sub(r"\bCÔ\s+PHÀN\b", "CỔ PHẦN", value, flags=re.IGNORECASE)
    value = re.sub(r"\bC[ỒỎÓ]\s+PH[ÀÁA]N\b", "CỔ PHẦN", value, flags=re.IGNORECASE)
    value = re.sub(
        r"\bCÔNG TY CÔ PH[ÁÀ]N\s+ÄNG TÀI TẠO\b",
        "CÔNG TY CỔ PHẦN NĂNG LƯỢNG TÁI TẠO",
        value,
        flags=re.IGNORECASE,
    )
    value = re.sub(r"\bNCOC\b", "NCC", value, flags=re.IGNORECASE)
    value = re.sub(r"\bLân ban hành\b", "Lần ban hành", value, flags=re.IGNORECASE)
    value = re.sub(r"\bSô tiền\b", "Số tiền", value, flags=re.IGNORECASE)
    value = re.sub(r"\bĐê nghị\b", "Đề nghị", value, flags=re.IGNORECASE)
    value = re.sub(r"\bKê toán\b", "Kế toán", value, flags=re.IGNORECASE)
    value = re.sub(r"\bTông Giám đôc\b", "Tổng Giám đốc", value, flags=re.IGNORECASE)
    value = re.sub(r"\bTrướng BP\b", "Trưởng BP", value, flags=re.IGNORECASE)
    value = re.sub(r"\b(?:LÀ|nvAe|DV;)(?=\s+renewable energy\b)", "DVA", value, flags=re.IGNORECASE)
    value = re.sub(r"\bTranse(?=:\s*1[⁄/]1\b)", "Trang", value, flags=re.IGNORECASE)
    value = re.sub(r"\bcá VAT\b", "cả VAT", value, flags=re.IGNORECASE)
    value = re.sub(r"\bSỐốTT\b", "STT", value, flags=re.IGNORECASE)
    value = re.sub(r"\b(\d+)năm\b", r"\1 năm", value, flags=re.IGNORECASE)
    value = re.sub(r"\bBênA\b", "Bên A", value, flags=re.IGNORECASE)
    value = re.sub(r"\bTổngGiámđốc\b", "Tổng Giám đốc", value, flags=re.IGNORECASE)
    value = re.sub(r"\b(?:TƠĐ|TGƠĐ)\b", "TGĐ", value, flags=re.IGNORECASE)
    value = re.sub(r"\bHKế toán\b", "Kế toán", value, flags=re.IGNORECASE)
    value = re.sub(r"\bphụ trách toán trưởng\b", "phụ trách Kế toán trưởng", value, flags=re.IGNORECASE)
    value = re.sub(r"\bVLI1\.", "VI.1.", value)
    value = re.sub(r"\bVL3\.", "VI.3.", value)
    value = re.sub(r"\bLƯU HÒ SƠ\b", "LƯU HỒ SƠ", value, flags=re.IGNORECASE)
    value = re.sub(r"\bĐịa điêm\b", "Địa điểm", value, flags=re.IGNORECASE)
    value = re.sub(r"\bT[ÿỷ]\s+giá\s+ngoại\s+t§", "Tỷ giá ngoại tệ", value, flags=re.IGNORECASE)
    value = re.sub(r"\bquy đồi\b", "quy đổi", value, flags=re.IGNORECASE)
    value = re.sub(r"\bđóng dầu\b", "đóng dấu", value, flags=re.IGNORECASE)
    value = re.sub(r"\bTới cả\b", "Tất cả", value, flags=re.IGNORECASE)
    value = re.sub(r"\btắt cả\b", "tất cả", value, flags=re.IGNORECASE)
    value = re.sub(r"\bđêu\b", "đều", value, flags=re.IGNORECASE)
    value = re.sub(r"\bđăng kợ\b", "đăng ký", value, flags=re.IGNORECASE)
    value = re.sub(r"\bđiện hử\b", "điện tử", value, flags=re.IGNORECASE)
    value = re.sub(r"\b200\.000/ND\b", "200.000 VND", value, flags=re.IGNORECASE)
    value = re.sub(r"\bđây đủ\b", "đầy đủ", value, flags=re.IGNORECASE)
    value = re.sub(r"\bCBCNYV\b", "CBCNV", value, flags=re.IGNORECASE)
    value = re.sub(r"\.\.\.\)\s+hợp\b", "...); trường hợp", value, flags=re.IGNORECASE)
    value = re.sub(r"\bphủi\b", "phải", value, flags=re.IGNORECASE)
    value = re.sub(r"\bghỉ\b", "ghi", value, flags=re.IGNORECASE)
    value = re.sub(r"\bPhiều\b", "Phiếu", value, flags=re.IGNORECASE)
    value = re.sub(r"\bsố kế toán\b", "sổ kế toán", value, flags=re.IGNORECASE)
    value = re.sub(r"\bSô\b", "Số", value)
    value = re.sub(r"\bRev\.0\]1\b", "Rev.01", value, flags=re.IGNORECASE)
    value = re.sub(r"\bảnh:\s*(?=\d{2}/\d{2}/\d{4}\b)", "Ngày: ", value, flags=re.IGNORECASE)
    value = re.sub(r"\bngay\b", "ngày", value)
    value = re.sub(r"\bnam\b", "năm", value)
    value = re.sub(r"\bcua\b", "của", value)
    value = re.sub(r"\brenewabie\b", "renewable", value, flags=re.IGNORECASE)
    value = re.sub(r"\bsrr\b", "STT", value, flags=re.IGNORECASE)
    value = value.replace("L]", "☐").replace("[]", "☐")
    value = re.sub(r"\bL1\b", "☐", value)
    value = re.sub(r"\bS[ôố]\s+ti[êe]n\s+b[ăa]ng\s+chữ\b", "Số tiền bằng chữ", value, flags=re.IGNORECASE)
    value = re.sub(r"\b1\s*\n\s*Đúng\b", "Đúng", value, flags=re.IGNORECASE)
    value = re.sub(r"\bthanh toán lân này\b", "thanh toán lần này", value, flags=re.IGNORECASE)
    value = re.sub(r"\([l1],\s*họ tên\)", "(ký, họ tên)", value, flags=re.IGNORECASE)
    value = re.sub(r"^họ tên\)", "(ký, họ tên)", value, flags=re.IGNORECASE)
    value = re.sub(r"(?<!,)\s+họ tên\)", " (ký, họ tên)", value, flags=re.IGNORECASE)
    for source, replacement in _JOINED_WORD_FIXES.items():
        value = re.sub(rf"\b{re.escape(source)}\b", replacement, value, flags=re.IGNORECASE)
    value = re.sub(r"(?<!\w)[|¦¬_=—~]{2,}(?!\w)", " ", value)
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r" *\n *", "\n", value)
    return value.strip()


def join_ocr_tokens(tokens: Iterable[str]) -> str:
    """Join OCR tokens with real word spacing and natural punctuation."""

    joined = " ".join(str(token).strip() for token in tokens if str(token).strip())
    joined = re.sub(r"\s+([,.;:!?%\)\]\}…])", r"\1", joined)
    joined = re.sub(r"([\(\[\{])\s+", r"\1", joined)
    return normalize_ocr_text(joined)


def ocr_data_to_words(
    data: dict[str, list[Any]],
    confidence_threshold: float = 40.0,
) -> list[dict[str, Any]]:
    words: list[dict[str, Any]] = []
    count = len(data.get("text", []))
    for index in range(count):
        text = normalize_ocr_text(str(data["text"][index]))
        try:
            confidence = float(data.get("conf", ["-1"] * count)[index])
        except (TypeError, ValueError):
            confidence = -1
        text = text.strip(" |¦¬_—=~")
        alphanumeric = sum(character.isalnum() for character in text)
        folded_text = unicodedata.normalize("NFD", text.casefold().replace("đ", "d"))
        folded_text = "".join(
            character for character in folded_text if unicodedata.category(character) != "Mn"
        )
        low_confidence_prefix = bool(
            re.fullmatch(r"\(?v(?:[/\\]?v|a)?:?", folded_text.strip())
        )
        minimum_confidence = 25.0 if low_confidence_prefix else confidence_threshold
        if (
            not text
            or confidence < minimum_confidence
            or (
                alphanumeric == 0
                and text not in {".", ",", ";", ":", "!", "?", "%", "+", "-", "/", "(", ")", "☐"}
            )
            or (len(text) >= 3 and alphanumeric / len(text) < 0.30)
        ):
            continue
        left = int(data.get("left", [0] * count)[index])
        top = int(data.get("top", [0] * count)[index])
        width = int(data.get("width", [0] * count)[index])
        height = int(data.get("height", [0] * count)[index])
        try:
            raw_font_size = data.get("font_size", [None] * count)[index]
            font_size = float(raw_font_size) if raw_font_size not in {None, ""} else None
        except (IndexError, TypeError, ValueError):
            font_size = None
        words.append(
            {
                "text": text,
                "confidence": confidence,
                "left": left,
                "top": top,
                "right": left + width,
                "bottom": top + height,
                "width": width,
                "height": height,
                "font_size": font_size,
                "line_key": (
                    int(data.get("block_num", [0] * count)[index]),
                    int(data.get("par_num", [0] * count)[index]),
                    int(data.get("line_num", [0] * count)[index]),
                ),
            }
        )
    return words


def ocr_words_to_lines(words: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[tuple[int, int, int], list[dict[str, Any]]] = {}
    for word in words:
        grouped.setdefault(tuple(word["line_key"]), []).append(word)

    lines: list[dict[str, Any]] = []
    for line_words in grouped.values():
        line_words.sort(key=lambda word: word["left"])
        confidences = [float(word.get("confidence", 0)) for word in line_words]
        lines.append(
            {
                "text": join_ocr_tokens(word["text"] for word in line_words),
                "left": min(word["left"] for word in line_words),
                "top": min(word["top"] for word in line_words),
                "right": max(word["right"] for word in line_words),
                "bottom": max(word["bottom"] for word in line_words),
                "height": max(word["height"] for word in line_words),
                "median_height": median(word["height"] for word in line_words),
                "confidence": sum(confidences) / max(1, len(confidences)),
                "word_count": len(line_words),
                "words": [dict(word) for word in line_words],
            }
        )
    lines.sort(key=lambda line: (line["top"], line["left"]))
    return lines


def _line_is_noise(line: dict[str, Any]) -> bool:
    text = normalize_ocr_text(str(line.get("text", "")))
    useful = sum(character.isalnum() for character in text)
    confidence = float(line.get("confidence", 0))
    compact = text.strip(" .,:;!?()[]{}-/")
    folded = compact.casefold()
    if not text or useful == 0:
        return True
    if useful <= 4 and ("\\" in text or "|" in text or "¦" in text):
        return True
    if useful == 1 and not compact.isdigit() and compact not in {"A", "B", "C", "D", "I", "V", "X"}:
        return True
    if re.fullmatch(r"\d{1,2}[.)]", text):
        return True
    if useful <= 3 and any(character.isdigit() for character in compact) and any(
        character.isalpha() for character in compact
    ):
        return True
    if (
        useful <= 3
        and compact.isupper()
        and folded not in _SHORT_LINE_ALLOWLIST
        and folded not in {"i", "ii", "iii", "iv", "v", "vi", "x"}
    ):
        return True
    if re.fullmatch(r"[A-Za-zÀ-ỹĐđ]", compact) and compact not in {"A", "B", "C", "D", "I", "V", "X"}:
        return True
    if useful <= 3 and confidence < 70 and folded not in _SHORT_LINE_ALLOWLIST:
        return True
    if useful < 8 and confidence < 50:
        return True
    return False


def _line_overlap(left: dict[str, Any], right: dict[str, Any]) -> float:
    horizontal = max(
        0,
        min(int(left["right"]), int(right["right"]))
        - max(int(left["left"]), int(right["left"])),
    )
    vertical = max(
        0,
        min(int(left["bottom"]), int(right["bottom"]))
        - max(int(left["top"]), int(right["top"])),
    )
    intersection = horizontal * vertical
    left_area = max(
        1,
        (int(left["right"]) - int(left["left"]))
        * (int(left["bottom"]) - int(left["top"])),
    )
    right_area = max(
        1,
        (int(right["right"]) - int(right["left"]))
        * (int(right["bottom"]) - int(right["top"])),
    )
    return intersection / max(1, left_area + right_area - intersection)


def _vertical_overlap(left: dict[str, Any], right: dict[str, Any]) -> float:
    overlap = max(
        0,
        min(int(left["bottom"]), int(right["bottom"]))
        - max(int(left["top"]), int(right["top"])),
    )
    minimum_height = max(
        1,
        min(
            int(left["bottom"]) - int(left["top"]),
            int(right["bottom"]) - int(right["top"]),
        ),
    )
    return overlap / minimum_height


def _line_score(line: dict[str, Any]) -> float:
    text = str(line.get("text", ""))
    useful = sum(character.isalnum() for character in text)
    words = max(1, int(line.get("word_count", len(text.split()))))
    confidence = float(line.get("confidence", 0))
    return confidence + min(32.0, useful * 0.45) + min(8.0, words * 0.8)


def _accent_key(text: str) -> str:
    decomposed = unicodedata.normalize("NFD", text.casefold().replace("đ", "d"))
    return "".join(
        character
        for character in decomposed
        if unicodedata.category(character) != "Mn" and character.isalnum()
    )


def _extend_line(
    shorter: dict[str, Any],
    longer: dict[str, Any],
) -> dict[str, Any] | None:
    shorter_words = str(shorter.get("text", "")).split()
    longer_words = str(longer.get("text", "")).split()
    if len(longer_words) <= len(shorter_words):
        return None
    if [_accent_key(word) for word in shorter_words] != [
        _accent_key(word) for word in longer_words[: len(shorter_words)]
    ]:
        return None
    result = dict(shorter)
    result["text"] = normalize_ocr_text(
        " ".join(shorter_words + longer_words[len(shorter_words) :])
    )
    result["right"] = max(int(shorter["right"]), int(longer["right"]))
    result["bottom"] = max(int(shorter["bottom"]), int(longer["bottom"]))
    result["height"] = max(int(shorter["height"]), int(longer["height"]))
    result["word_count"] = len(longer_words)
    if longer.get("words"):
        combined_words = [dict(word) for word in longer["words"]]
        for index, text in enumerate(shorter_words):
            if index < len(combined_words):
                combined_words[index]["text"] = text
        result["words"] = combined_words
    return result


def _merge_horizontal_line_fragments(lines: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Join OCR fragments that share a visual baseline.

    Tesseract frequently emits the abbreviation and its definition as two
    independent lines (for example ``KTT`` and ``Kế toán trưởng``), even
    though they occupy the same printed row.  Writing both as paragraphs
    doubles the vertical space and makes a scanned source page spill onto the
    next Word page.  Recombining only strongly-overlapping, reasonably-close
    fragments preserves reading order without flattening real columns.
    """

    if len(lines) < 2:
        return lines
    page_span = max(1, max(int(line["right"]) for line in lines) - min(int(line["left"]) for line in lines))
    maximum_gap = max(24.0, page_span * 0.24)
    pending = sorted((dict(line) for line in lines), key=lambda line: (int(line["top"]), int(line["left"])))
    result: list[dict[str, Any]] = []
    while pending:
        current = pending.pop(0)
        while True:
            match_index: int | None = None
            for index, candidate in enumerate(pending):
                if int(candidate["top"]) > int(current["bottom"]):
                    break
                left_fragment, right_fragment = sorted(
                    (current, candidate), key=lambda line: int(line["left"])
                )
                gap = int(right_fragment["left"]) - int(left_fragment["right"])
                height_limit = min(
                    max(
                        float(current.get("median_height", current.get("height", 1))) * 8,
                        maximum_gap,
                    ),
                    page_span * 0.45,
                )
                if (
                    gap >= 0
                    and gap <= height_limit
                    and _vertical_overlap(current, candidate) >= 0.62
                ):
                    match_index = index
                    break
            if match_index is None:
                break
            candidate = pending.pop(match_index)
            current_words = [dict(word) for word in current.get("words", [])]
            candidate_words = [dict(word) for word in candidate.get("words", [])]
            words = sorted(current_words + candidate_words, key=lambda word: int(word.get("left", 0)))
            current_count = max(1, int(current.get("word_count", len(current_words) or 1)))
            candidate_count = max(1, int(candidate.get("word_count", len(candidate_words) or 1)))
            current = {
                **current,
                "text": join_ocr_tokens(word["text"] for word in words)
                if words
                else normalize_ocr_text(f'{current.get("text", "")} {candidate.get("text", "")}'),
                "left": min(int(current["left"]), int(candidate["left"])),
                "top": min(int(current["top"]), int(candidate["top"])),
                "right": max(int(current["right"]), int(candidate["right"])),
                "bottom": max(int(current["bottom"]), int(candidate["bottom"])),
                "height": max(int(current.get("height", 0)), int(candidate.get("height", 0))),
                "median_height": median(
                    [
                        float(current.get("median_height", current.get("height", 1))),
                        float(candidate.get("median_height", candidate.get("height", 1))),
                    ]
                ),
                "confidence": (
                    float(current.get("confidence", 0)) * current_count
                    + float(candidate.get("confidence", 0)) * candidate_count
                )
                / (current_count + candidate_count),
                "word_count": current_count + candidate_count,
                "words": words,
            }
        result.append(current)
    result.sort(key=lambda line: (int(line["top"]), int(line["left"])))
    return result


def _remove_nearby_duplicate_labels(lines: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Drop a sparse-OCR abbreviation repeated beside an existing ``X: ...`` row."""

    result: list[dict[str, Any]] = []
    for index, line in enumerate(lines):
        text = str(line.get("text", "")).strip()
        candidate_key = _accent_key(text)
        letters = [character for character in text if character.isalpha()]
        is_short_label = bool(
            ":" not in text
            and " " not in text
            and 2 <= len(candidate_key) <= 6
            and letters
            and all(character == character.upper() for character in letters)
        )
        if not is_short_label:
            result.append(line)
            continue
        center = (int(line["top"]) + int(line["bottom"])) / 2
        duplicate = False
        for other_index, other in enumerate(lines):
            if other_index == index:
                continue
            other_text = str(other.get("text", "")).strip()
            if ":" not in other_text:
                continue
            label_text = other_text.split(":", 1)[0].strip()
            label_key = _accent_key(label_text)
            if not 2 <= len(label_key) <= 7:
                continue
            other_center = (int(other["top"]) + int(other["bottom"])) / 2
            nearby = abs(center - other_center) <= max(
                int(line["bottom"]) - int(line["top"]),
                int(other["bottom"]) - int(other["top"]),
            ) * 1.15
            similar = SequenceMatcher(None, candidate_key, label_key, autojunk=False).ratio() >= 0.66
            if nearby and similar:
                duplicate = True
                break
        if not duplicate:
            result.append(line)
    return result


def merge_ocr_lines(
    primary: Iterable[dict[str, Any]],
    supplemental: Iterable[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Merge normal and sparse-layout OCR lines, preferring the clearer candidate."""

    merged: list[dict[str, Any]] = []
    for line in primary:
        item = dict(line)
        item["text"] = normalize_ocr_text(str(item.get("text", "")))
        if not _line_is_noise(item):
            merged.append(item)
    for candidate in supplemental:
        candidate = dict(candidate)
        candidate["text"] = normalize_ocr_text(str(candidate.get("text", "")))
        if _line_is_noise(candidate):
            continue
        candidate_key = _accent_key(candidate["text"])
        if len(candidate_key) >= 4 and any(
            _vertical_overlap(existing, candidate) >= 0.55
            and candidate_key in _accent_key(str(existing.get("text", "")))
            for existing in merged
        ):
            # Sparse-layout OCR often emits an extra short label that already
            # exists inside a complete primary line (for example "Tiền mặt").
            continue
        matches = [
            (index, _line_overlap(existing, candidate))
            for index, existing in enumerate(merged)
        ]
        matches = [(index, overlap) for index, overlap in matches if overlap >= 0.18]
        if not matches:
            merged.append(candidate)
            continue
        index, _overlap = max(matches, key=lambda item: item[1])
        existing = merged[index]
        extension = _extend_line(existing, candidate)
        if extension is not None:
            merged[index] = extension
            continue
        existing_useful = sum(character.isalnum() for character in str(existing.get("text", "")))
        candidate_useful = sum(character.isalnum() for character in str(candidate.get("text", "")))
        candidate_is_more_complete = (
            candidate_useful >= existing_useful + 3
            and float(candidate.get("confidence", 0)) >= float(existing.get("confidence", 0)) - 8
        )
        candidate_is_shorter = candidate_useful < existing_useful
        materially_more_confident = (
            float(candidate.get("confidence", 0)) >= float(existing.get("confidence", 0)) + 8
        )
        if candidate_is_more_complete or (
            (not candidate_is_shorter or materially_more_confident)
            and _line_score(candidate) > _line_score(existing) + 2
        ):
            merged[index] = candidate
    merged = _remove_nearby_duplicate_labels(_merge_horizontal_line_fragments(merged))
    for index, line in enumerate(merged):
        if str(line.get("text", "")).casefold() != "nam" or index == 0:
            continue
        previous = str(merged[index - 1].get("text", "")).casefold()
        if "ngày" in previous or "tháng" in previous:
            line["text"] = "năm"
    return merged


def normalize_document_ocr_text(doc) -> tuple[int, int]:
    """Apply final OCR cleanup without replacing the document's font formatting."""

    normalized_count = 0
    removed_count = 0

    def replace_text(paragraph, value: str) -> None:
        nonlocal normalized_count
        if paragraph.text == value:
            return
        runs = paragraph.runs
        if runs:
            runs[0].text = value
            for run in runs[1:]:
                run.text = ""
        else:
            paragraph.add_run(value)
        normalized_count += 1

    for paragraph in list(doc.paragraphs):
        original = paragraph.text
        if not original.strip():
            continue
        normalized = normalize_ocr_text(original)
        if _line_is_noise(
            {
                "text": normalized,
                "confidence": 100,
                "word_count": max(1, len(normalized.split())),
            }
        ):
            element = paragraph._element
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)
                removed_count += 1
            continue
        replace_text(paragraph, normalized)

    seen: set[Any] = set()
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    key = paragraph._p
                    if key in seen:
                        continue
                    seen.add(key)
                    if paragraph.text.strip():
                        replace_text(paragraph, normalize_ocr_text(paragraph.text))
    return normalized_count, removed_count


def ocr_data_to_lines(
    data: dict[str, list[Any]],
    confidence_threshold: float = 40.0,
) -> list[dict[str, Any]]:
    """Group Tesseract word boxes into editable, explicitly spaced lines."""

    return ocr_words_to_lines(ocr_data_to_words(data, confidence_threshold))


def ocr_words_to_text(words: Iterable[dict[str, Any]]) -> str:
    lines = ocr_words_to_lines(words)
    return "\n".join(
        line["text"]
        for line in lines
        if line["text"]
        and not (
            sum(character.isalnum() for character in line["text"]) == 1
            and line["text"].strip().isalpha()
            and line["text"].strip() not in {"A", "B", "C", "D", "I", "V", "X"}
        )
    )


def _is_uppercase_heading(text: str) -> bool:
    letters = [character for character in text if character.isalpha()]
    return len(letters) >= 3 and all(character == character.upper() for character in letters)


def _font_size_for_line(line: dict[str, Any], points_per_pixel: float) -> float:
    raw_sizes = [
        float(word["font_size"])
        for word in line.get("words", [])
        if word.get("font_size") is not None and float(word["font_size"]) > 0
    ]
    word_heights = [float(word.get("height", 0)) for word in line.get("words", [])]
    pixel_height = median(word_heights) if word_heights else float(line.get("median_height", line["height"]))
    if raw_sizes:
        raw_size = float(median(raw_sizes))
        # hOCR reports point sizes when the raster carries DPI metadata and
        # pixel sizes otherwise.  Tesseract can emit pixel values below 32 as
        # well (especially for dense A4 scans), so the old ``raw_size > 32``
        # guard occasionally treated 25-31 px as 25-31 pt.  Compare against
        # the glyph box instead and apply a small x-height calibration; this
        # keeps scan text from overflowing into mostly-empty Word pages.
        if raw_size >= pixel_height * 0.65:
            raw_size *= points_per_pixel * 0.88
        return max(6.0, min(32.0, raw_size))
    return max(8.0, min(22.0, pixel_height * points_per_pixel * 1.12))


def _word_ink_features(gray, word: dict[str, Any]) -> tuple[float, bool]:
    if cv2 is None or np is None:
        return 0.0, False
    left = max(0, int(word["left"]))
    top = max(0, int(word["top"]))
    right = min(gray.shape[1], int(word["right"]))
    bottom = min(gray.shape[0], int(word["bottom"]))
    if right - left < 2 or bottom - top < 3:
        return 0.0, False

    pad = 2
    crop = gray[
        max(0, top - pad) : min(gray.shape[0], bottom + pad),
        max(0, left - pad) : min(gray.shape[1], right + pad),
    ]
    _threshold, binary = cv2.threshold(
        crop,
        0,
        255,
        cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU,
    )
    horizontal_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT,
        (max(8, int(crop.shape[1] * 0.55)), 1),
    )
    horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel)
    clean = cv2.subtract(binary, horizontal)
    foreground = clean > 0
    if not foreground.any():
        return 0.0, False
    eroded = cv2.erode(
        foreground.astype(np.uint8) * 255,
        np.ones((3, 3), np.uint8),
    )
    survival = float((eroded > 0).sum() / max(1, foreground.sum()))

    extension = max(8, int((bottom - top) * 0.35))
    underline_crop = gray[
        max(0, top - pad) : min(gray.shape[0], bottom + extension),
        max(0, left - pad) : min(gray.shape[1], right + pad),
    ]
    _threshold, underline_binary = cv2.threshold(
        underline_crop,
        0,
        255,
        cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU,
    )
    underline_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT,
        (max(6, int((right - left) * 0.45)), 1),
    )
    horizontal = cv2.morphologyEx(
        underline_binary,
        cv2.MORPH_OPEN,
        underline_kernel,
    )
    underline = False
    rows, _columns = np.where(horizontal > 0)
    if len(rows):
        for row_index in np.unique(rows):
            columns = np.where(horizontal[row_index] > 0)[0]
            if not len(columns):
                continue
            run_ratio = (int(columns.max()) - int(columns.min()) + 1) / max(
                1,
                underline_crop.shape[1],
            )
            row_ratio = int(row_index) / max(1, underline_crop.shape[0] - 1)
            if run_ratio >= 0.65 and row_ratio >= 0.64:
                underline = True
                break
    return survival, underline


def _italic_shear_score(gray, line: dict[str, Any]) -> tuple[float, float]:
    """Return the best de-slant shear and its projection gain.

    Italic glyphs have consistently slanted near-vertical strokes.  Shearing
    them back toward vertical makes the column projection substantially
    sharper; regular text normally peaks at a zero shear.
    """

    if cv2 is None or np is None or len(str(line.get("text", ""))) < 5:
        return 0.0, 1.0
    left = max(0, int(line["left"]) - 2)
    top = max(0, int(line["top"]) - 2)
    right = min(gray.shape[1], int(line["right"]) + 2)
    bottom = min(gray.shape[0], int(line["bottom"]) + 2)
    crop = gray[top:bottom, left:right]
    if crop.size == 0:
        return 0.0, 1.0
    _threshold, binary = cv2.threshold(
        crop,
        0,
        1,
        cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU,
    )
    component_count, labels, stats, _centroids = cv2.connectedComponentsWithStats(
        binary.astype(np.uint8),
        8,
    )
    clean = np.zeros_like(binary, dtype=np.uint8)
    minimum_height = max(4, int(binary.shape[0] * 0.20))
    minimum_width = max(8, int(binary.shape[0] * 0.35))
    for component_index in range(1, component_count):
        _x, _y, width, height, _area = stats[component_index]
        if height >= minimum_height or width >= minimum_width:
            clean[labels == component_index] = 1
    if int(clean.sum()) < 20:
        return 0.0, 1.0

    height, width = clean.shape
    padding = int(height * 0.60) + 4
    scores: list[tuple[float, float]] = []
    for shear in np.linspace(-0.45, 0.45, 37):
        transform = np.float32(
            [[1.0, float(shear), padding - float(shear) * height], [0.0, 1.0, 0.0]]
        )
        transformed = cv2.warpAffine(
            clean,
            transform,
            (width + 2 * padding, height),
            flags=cv2.INTER_NEAREST,
            borderValue=0,
        )
        projection = transformed.sum(axis=0).astype(float)
        score = float((projection * projection).sum() / max(1.0, projection.sum()))
        scores.append((score, float(shear)))
    best_score, best_shear = max(scores)
    zero_score = min(scores, key=lambda item: abs(item[1]))[0]
    return best_shear, best_score / max(0.0001, zero_score)


def annotate_ocr_styles(image, lines: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    """Estimate bold, italic and underline from the pixels behind OCR words."""

    items = [dict(line) for line in lines]
    if not items:
        return items
    if cv2 is None or np is None:
        for line in items:
            semantic_bold = _is_uppercase_heading(str(line.get("text", "")))
            for word in line.get("words", []):
                word["bold"] = semantic_bold
                word["italic"] = False
                word["underline"] = False
        return items

    if hasattr(image, "convert"):
        gray = np.asarray(image.convert("L"))
    else:
        raw = np.asarray(image)
        gray = cv2.cvtColor(raw, cv2.COLOR_BGR2GRAY) if raw.ndim == 3 else raw
    gray = np.ascontiguousarray(gray.astype(np.uint8))

    global_regular_scores: list[float] = []
    scores_by_size: dict[int, list[float]] = {}
    for line in items:
        shear, gain = _italic_shear_score(gray, line)
        line_italic = abs(shear) >= 0.14 and gain >= 1.12
        line["italic"] = line_italic
        uppercase = _is_uppercase_heading(str(line.get("text", "")))
        for word in line.get("words", []):
            survival, underline = _word_ink_features(gray, word)
            word["ink_survival"] = survival
            word["underline"] = underline
            word["italic"] = line_italic
            useful = sum(character.isalnum() for character in str(word.get("text", "")))
            if line_italic or underline or uppercase or useful < 2:
                continue
            global_regular_scores.append(survival)
            if word.get("font_size") is not None:
                key = int(round(float(word["font_size"])))
                scores_by_size.setdefault(key, []).append(survival)

    global_baseline = median(global_regular_scores) if global_regular_scores else 0.34
    for line in items:
        words = list(line.get("words", []))
        uppercase = _is_uppercase_heading(str(line.get("text", "")))
        scores = [float(word.get("ink_survival", 0.0)) for word in words]
        for index, word in enumerate(words):
            if uppercase:
                word["bold"] = True
                continue
            font_size = word.get("font_size")
            key = int(round(float(font_size))) if font_size is not None else 0
            peers = scores_by_size.get(key, [])
            baseline = median(peers) if len(peers) >= 8 else global_baseline
            threshold = baseline + max(0.055, baseline * 0.17)
            start = max(0, index - 1)
            end = min(len(scores), index + 2)
            local_average = sum(scores[start:end]) / max(1, end - start)
            score = scores[index]
            word["bold"] = bool(
                score >= threshold + 0.035
                or (score >= threshold - 0.018 and local_average >= threshold + 0.005)
            )
        line["bold"] = bool(words) and sum(bool(word.get("bold")) for word in words) >= max(
            1,
            len(words) // 2,
        )
        line["underline"] = bool(words) and sum(
            bool(word.get("underline")) for word in words
        ) >= max(1, len(words) // 2)
    return items


def _dominant_body_font(lines: list[dict[str, Any]], image_width: int) -> float:
    weighted: dict[float, int] = {}
    for line in lines:
        text = str(line.get("text", ""))
        width_ratio = (int(line["right"]) - int(line["left"])) / max(1, image_width)
        if (
            width_ratio < 0.34
            or line.get("italic")
            or _is_uppercase_heading(text)
            or len(text) < 20
        ):
            continue
        size = round(float(line["font_size"]) * 2) / 2
        weighted[size] = weighted.get(size, 0) + max(1, int(line.get("word_count", 1)))
    return max(weighted, key=weighted.get) if weighted else 11.0


def _is_body_line(line: dict[str, Any], image_width: int, body_font: float) -> bool:
    text = str(line.get("text", "")).strip()
    width_ratio = (int(line["right"]) - int(line["left"])) / max(1, image_width)
    underlined_words = sum(bool(word.get("underline")) for word in line.get("words", []))
    label = bool(re.match(r"^(?:kính\s+gửi|dear\b|to\s*:)", text, flags=re.IGNORECASE))
    return bool(
        len(text) >= 18
        and width_ratio >= 0.32
        and int(line["left"]) < image_width * 0.50
        and abs(float(line["font_size"]) - body_font) <= max(3.0, body_font * 0.25)
        and not line.get("italic")
        and not _is_uppercase_heading(text)
        and not underlined_words
        and not label
    )


def _group_ocr_lines(
    lines: list[dict[str, Any]],
    image_width: int,
) -> tuple[list[dict[str, Any]], int, float]:
    body_font = _dominant_body_font(lines, image_width)
    body_lines = [line for line in lines if _is_body_line(line, image_width, body_font)]
    base_candidates = sorted(
        int(line["left"])
        for line in body_lines
        if int(line["right"]) - int(line["left"]) >= image_width * 0.38
    )
    if base_candidates:
        body_left = base_candidates[max(0, int(len(base_candidates) * 0.25) - 1)]
    else:
        body_left = min((int(line["left"]) for line in lines), default=0)
    typical_height = median(
        [float(line.get("height", 0)) for line in body_lines]
        or [float(line.get("height", 1)) for line in lines]
    )
    maximum_intra_paragraph_gap = max(8.0, typical_height * 0.55)
    first_line_indent_threshold = max(24, int(image_width * 0.035))

    groups: list[dict[str, Any]] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        group_lines = [line]
        is_body = _is_body_line(line, image_width, body_font)
        if is_body:
            cursor = index + 1
            while cursor < len(lines):
                candidate = lines[cursor]
                if not _is_body_line(candidate, image_width, body_font):
                    break
                gap = int(candidate["top"]) - int(group_lines[-1]["bottom"])
                starts_new_paragraph = (
                    int(candidate["left"]) > body_left + first_line_indent_threshold
                )
                if gap > maximum_intra_paragraph_gap or starts_new_paragraph:
                    break
                group_lines.append(candidate)
                cursor += 1
            index = cursor
        else:
            index += 1
        groups.append(
            {
                "lines": group_lines,
                "kind": "body" if is_body else "line",
                "left": min(int(item["left"]) for item in group_lines),
                "top": min(int(item["top"]) for item in group_lines),
                "right": max(int(item["right"]) for item in group_lines),
                "bottom": max(int(item["bottom"]) for item in group_lines),
            }
        )
    return groups, body_left, body_font


def _line_words(line: dict[str, Any]) -> list[dict[str, Any]]:
    words = [dict(word) for word in line.get("words", []) if str(word.get("text", "")).strip()]
    if words:
        joined = join_ocr_tokens(word["text"] for word in words)
        expected = normalize_ocr_text(str(line.get("text", "")))
        if _accent_key(joined) == _accent_key(expected):
            return words
    return [
        {
            "text": normalize_ocr_text(str(line.get("text", ""))),
            "bold": bool(line.get("bold")),
            "italic": bool(line.get("italic")),
            "underline": bool(line.get("underline")),
        }
    ]


def _needs_word_space(previous: str, current: str) -> bool:
    if not previous or not current:
        return False
    if current[0] in ",.;:!?%)]}…":
        return False
    if previous[-1] in "([{“":
        return False
    return True


def _styled_segments(
    group_lines: list[dict[str, Any]],
) -> tuple[str, list[tuple[str, tuple[bool, bool, bool]]]]:
    raw_characters: list[str] = []
    raw_styles: list[tuple[bool, bool, bool]] = []
    previous = ""
    for line in group_lines:
        for word in _line_words(line):
            token = str(word.get("text", "")).strip()
            if not token:
                continue
            style = (
                bool(word.get("bold")),
                bool(word.get("italic", line.get("italic"))),
                bool(word.get("underline")),
            )
            if _needs_word_space(previous, token):
                raw_characters.append(" ")
                raw_styles.append(style)
            raw_characters.extend(token)
            raw_styles.extend([style] * len(token))
            previous = token
    raw_text = "".join(raw_characters)
    normalized = normalize_ocr_text(raw_text)
    if not normalized:
        return "", []

    normalized_styles: list[tuple[bool, bool, bool] | None] = [None] * len(normalized)
    matcher = SequenceMatcher(None, raw_text, normalized, autojunk=False)
    default_style = raw_styles[0] if raw_styles else (False, False, False)
    for operation, old_start, old_end, new_start, new_end in matcher.get_opcodes():
        if operation == "equal":
            for offset in range(new_end - new_start):
                normalized_styles[new_start + offset] = raw_styles[old_start + offset]
            continue
        source_styles = raw_styles[old_start:old_end]
        if source_styles:
            replacement_style = max(set(source_styles), key=source_styles.count)
        elif old_start > 0 and raw_styles:
            replacement_style = raw_styles[old_start - 1]
        elif old_start < len(raw_styles):
            replacement_style = raw_styles[old_start]
        else:
            replacement_style = default_style
        for new_index in range(new_start, new_end):
            normalized_styles[new_index] = replacement_style

    segments: list[tuple[str, tuple[bool, bool, bool]]] = []
    start = 0
    while start < len(normalized):
        style = normalized_styles[start] or default_style
        end = start + 1
        while end < len(normalized) and (normalized_styles[end] or default_style) == style:
            end += 1
        segments.append((normalized[start:end], style))
        start = end
    return normalized, segments


def write_ocr_lines(
    doc,
    lines: Iterable[dict[str, Any]],
    *,
    image_width: int,
    content_width_inches: float,
    source_dpi: float | None = None,
    image=None,
) -> int:
    line_items = list(lines)
    if not line_items:
        return 0

    image_width = max(1, int(image_width))
    points_per_pixel = (
        72.0 / source_dpi
        if source_dpi and source_dpi > 0
        else content_width_inches * 72.0 / image_width
    )
    page_width_inches = (
        image_width / source_dpi
        if source_dpi and source_dpi > 0
        else content_width_inches + 2.0
    )
    for line in line_items:
        line["font_size"] = _font_size_for_line(line, points_per_pixel)
    if image is not None:
        line_items = annotate_ocr_styles(image, line_items)

    # Adjacent centered italic lines are commonly a multi-line subtitle.  Use
    # one stable size without collapsing their deliberate vertical spacing.
    for index in range(1, len(line_items)):
        previous = line_items[index - 1]
        current = line_items[index]
        previous_center = (int(previous["left"]) + int(previous["right"])) / 2 / image_width
        current_center = (int(current["left"]) + int(current["right"])) / 2 / image_width
        gap = int(current["top"]) - int(previous["bottom"])
        if (
            previous.get("italic")
            and current.get("italic")
            and abs(previous_center - 0.5) < 0.09
            and abs(current_center - 0.5) < 0.09
            and gap <= max(float(previous["height"]), float(current["height"])) * 0.80
            and abs(float(previous["font_size"]) - float(current["font_size"])) <= 2.0
        ):
            stable_size = float(round(median([previous["font_size"], current["font_size"]])))
            previous["font_size"] = stable_size
            current["font_size"] = stable_size

    groups, body_left, _body_font = _group_ocr_lines(line_items, image_width)
    previous_bottom: int | None = None
    for group in groups:
        group_lines = group["lines"]
        text, segments = _styled_segments(group_lines)
        if not text:
            continue
        paragraph = doc.add_paragraph()
        # ``word["font_size"]`` is the raw hOCR value and can still be in
        # raster pixels.  Every line has already gone through
        # ``_font_size_for_line`` above, so only use those normalized point
        # values when writing Word runs.
        font_size = round(float(median(line["font_size"] for line in group_lines)) * 2) / 2
        font_size = max(8.0, min(32.0, font_size))
        for segment_text, (bold, italic, underline) in segments:
            run = paragraph.add_run(segment_text)
            set_run_font(run, "Times New Roman", font_size)
            run.bold = bold
            run.italic = italic
            run.underline = underline

        width_ratio = (int(group["right"]) - int(group["left"])) / image_width
        center = (int(group["left"]) + int(group["right"])) / 2 / image_width
        left_ratio = int(group["left"]) / image_width
        right_ratio = int(group["right"]) / image_width
        label = bool(re.match(r"^(?:kính\s+gửi|dear\b|to\s*:)", text, flags=re.IGNORECASE))
        all_italic = all(bool(line.get("italic")) for line in group_lines)
        if group["kind"] == "body":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            first_indent = max(0, int(group_lines[0]["left"]) - body_left)
            if first_indent >= image_width * 0.015:
                paragraph.paragraph_format.first_line_indent = Inches(
                    min(0.85, first_indent / image_width * page_width_inches)
                )
        elif (
            not label
            and abs(center - 0.5) < 0.075
            and (
                width_ratio <= 0.50
                or all_italic
                or _is_uppercase_heading(text)
                or bool(group_lines[0].get("underline"))
            )
        ):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif left_ratio >= 0.46 and right_ratio >= 0.78 and width_ratio <= 0.55:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            relative_left = max(0, int(group["left"]) - body_left)
            left_inches = relative_left / image_width * page_width_inches
            if left_inches >= 0.08:
                paragraph.paragraph_format.left_indent = Inches(
                    min(left_inches, max(0.0, content_width_inches - 1.25))
                )

        paragraph.paragraph_format.line_spacing = 1.15
        if previous_bottom is not None:
            gap_points = max(0.0, (int(group["top"]) - previous_bottom) * points_per_pixel)
            paragraph.paragraph_format.space_before = Pt(min(36.0, max(0.0, gap_points - 2.0)))
        paragraph.paragraph_format.space_after = Pt(0)
        previous_bottom = int(group["bottom"])
    return len(line_items)


def write_ocr_image(
    doc,
    image,
    options: ConversionOptions,
    *,
    content_width_inches: float,
    source_dpi: float | None = None,
) -> tuple[int, list[str]]:
    """OCR one raster image and write editable paragraphs with approximate layout."""

    data, _language, warnings = image_to_data(image, options.language)
    lines = ocr_data_to_lines(data)
    count = write_ocr_lines(
        doc,
        lines,
        image_width=image.width,
        content_width_inches=content_width_inches,
        source_dpi=source_dpi,
        image=image,
    )
    return count, warnings
