from __future__ import annotations

import os
import re
import shutil
import subprocess
from xml.etree import ElementTree
from dataclasses import asdict, dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any

from app.converters.base import MissingDependencyError, MissingOcrEngineError


_HOCR_BBOX_RE = re.compile(r"\bbbox\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)")
_HOCR_CONFIDENCE_RE = re.compile(r"\bx_wconf\s+(-?[\d.]+)")
_HOCR_FONT_SIZE_RE = re.compile(r"\bx_fsize\s+([\d.]+)")


def _class_names(element: ElementTree.Element) -> set[str]:
    return set(element.attrib.get("class", "").split())


def _hocr_to_data(payload: bytes) -> dict[str, list[Any]]:
    """Convert Tesseract hOCR into the dictionary shape used by the converters.

    hOCR exposes Tesseract's estimated point size (``x_fsize``), which the TSV
    output omits.  Keeping that value avoids deriving font size from the tallest
    accented glyph in a line -- the main cause of alternating large/small text
    in Vietnamese scans.
    """

    try:
        root = ElementTree.fromstring(payload)
    except ElementTree.ParseError as exc:
        raise MissingOcrEngineError(f"Tesseract trả về dữ liệu hOCR không hợp lệ: {exc}") from exc

    data: dict[str, list[Any]] = {
        "text": [],
        "conf": [],
        "block_num": [],
        "par_num": [],
        "line_num": [],
        "left": [],
        "top": [],
        "width": [],
        "height": [],
        "font_size": [],
    }
    line_number = 0
    for line in root.iter():
        if "ocr_line" not in _class_names(line):
            continue
        line_number += 1
        for word in line.iter():
            if "ocrx_word" not in _class_names(word):
                continue
            title = word.attrib.get("title", "")
            bbox_match = _HOCR_BBOX_RE.search(title)
            if bbox_match is None:
                continue
            left, top, right, bottom = (int(value) for value in bbox_match.groups())
            confidence_match = _HOCR_CONFIDENCE_RE.search(title)
            font_size_match = _HOCR_FONT_SIZE_RE.search(title)
            text = "".join(word.itertext()).strip()
            data["text"].append(text)
            data["conf"].append(confidence_match.group(1) if confidence_match else "-1")
            data["block_num"].append(1)
            data["par_num"].append(1)
            data["line_num"].append(line_number)
            data["left"].append(left)
            data["top"].append(top)
            data["width"].append(max(0, right - left))
            data["height"].append(max(0, bottom - top))
            data["font_size"].append(
                float(font_size_match.group(1)) if font_size_match else None
            )
    return data


@dataclass(slots=True)
class OcrEngineInfo:
    available: bool
    command: str | None
    languages: list[str]
    tessdata_dir: str | None
    message: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _candidate_commands() -> list[Path | str]:
    project_root = Path(__file__).resolve().parents[2]
    candidates: list[Path | str] = []
    if explicit := os.getenv("TESSERACT_CMD"):
        candidates.append(Path(explicit))
    if on_path := shutil.which("tesseract"):
        candidates.append(Path(on_path))
    candidates.extend(
        [
            project_root / "tools" / "tesseract" / "tesseract.exe",
            Path(os.getenv("PROGRAMFILES", r"C:\Program Files")) / "Tesseract-OCR" / "tesseract.exe",
            Path(os.getenv("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR" / "tesseract.exe",
            Path(os.getenv("PROGRAMFILES(X86)", r"C:\Program Files (x86)")) / "Tesseract-OCR" / "tesseract.exe",
        ]
    )
    return candidates


def _candidate_tessdata(command: str | None) -> list[Path]:
    project_root = Path(__file__).resolve().parents[2]
    candidates: list[Path] = []
    if explicit := os.getenv("OCR2WORD_TESSDATA_DIR"):
        candidates.append(Path(explicit))
    if explicit := os.getenv("TESSDATA_PREFIX"):
        prefix = Path(explicit)
        candidates.extend([prefix, prefix / "tessdata"])
    if local_app_data := os.getenv("LOCALAPPDATA"):
        candidates.append(Path(local_app_data) / "OCR2Word" / "tessdata")
    candidates.append(project_root / ".tessdata")
    if command:
        candidates.append(Path(command).resolve().parent / "tessdata")

    valid: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate.resolve()).casefold()
        if key in seen:
            continue
        seen.add(key)
        if candidate.is_dir() and any(candidate.glob("*.traineddata")):
            valid.append(candidate)
    return valid


def _list_languages(command: str, tessdata: Path | None) -> list[str]:
    env = os.environ.copy()
    if tessdata:
        env["TESSDATA_PREFIX"] = str(tessdata)
    try:
        completed = subprocess.run(
            [command, "--list-langs"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=12,
            check=False,
            env=env,
        )
    except (OSError, subprocess.SubprocessError):
        return []
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    return sorted(line for line in lines if not line.lower().startswith("list of available"))


@lru_cache(maxsize=1)
def get_ocr_engine_info() -> OcrEngineInfo:
    command: str | None = None
    for candidate in _candidate_commands():
        candidate_path = Path(candidate)
        if candidate_path.is_file():
            command = str(candidate_path)
            break
    if command is None:
        return OcrEngineInfo(
            available=False,
            command=None,
            languages=[],
            tessdata_dir=None,
            message="Chưa tìm thấy Tesseract. PDF có lớp chữ và Excel vẫn chuyển đổi được.",
        )

    try:
        import pytesseract  # noqa: F401
    except ImportError:
        return OcrEngineInfo(
            available=False,
            command=command,
            languages=[],
            tessdata_dir=None,
            message="Đã tìm thấy Tesseract nhưng thiếu gói Python pytesseract.",
        )

    tessdata: Path | None = None
    languages: list[str] = []
    for candidate in _candidate_tessdata(command):
        candidate_languages = _list_languages(command, candidate)
        if candidate_languages:
            tessdata = candidate
            languages = candidate_languages
            break
    if not languages:
        languages = _list_languages(command, None)
    if not languages:
        return OcrEngineInfo(
            available=False,
            command=command,
            languages=[],
            tessdata_dir=str(tessdata) if tessdata else None,
            message="Tesseract chưa có dữ liệu ngôn ngữ hợp lệ.",
        )
    if "vie" not in languages:
        message = "OCR sẵn sàng nhưng chưa có bộ dữ liệu tiếng Việt; ứng dụng sẽ dùng ngôn ngữ hiện có."
    else:
        vie_model = tessdata / "vie.traineddata" if tessdata else None
        high_quality = bool(vie_model and vie_model.is_file() and vie_model.stat().st_size >= 5_000_000)
        message = (
            "OCR tiếng Việt chất lượng cao đã sẵn sàng."
            if high_quality
            else "OCR tiếng Việt/Anh đã sẵn sàng (model nhanh)."
        )
    return OcrEngineInfo(
        available=True,
        command=command,
        languages=languages,
        tessdata_dir=str(tessdata) if tessdata else None,
        message=message,
    )


def _ocr_runtime(requested_language: str, *, prefer_vietnamese: bool):
    info = get_ocr_engine_info()
    if not info.command:
        raise MissingOcrEngineError(
            "Tài liệu này cần OCR nhưng chưa tìm thấy Tesseract. Hãy cài Tesseract hoặc chạy ứng dụng bằng Docker."
        )
    if not info.available:
        if "pytesseract" in info.message:
            raise MissingDependencyError(info.message)
        raise MissingOcrEngineError(info.message)

    requested = [part for part in requested_language.split("+") if part]
    selected = [part for part in requested if part in info.languages]
    warnings: list[str] = []
    if not selected:
        selected = ["eng"] if "eng" in info.languages else info.languages[:1]
    missing = [part for part in requested if part not in info.languages]
    if missing:
        warnings.append(
            f"Thiếu dữ liệu OCR cho {', '.join(missing)}; đã dùng {'+'.join(selected)}."
        )
    if prefer_vietnamese and requested and requested[0] == "vie" and "vie" in selected:
        # The Vietnamese model still recognizes Latin/English labels, while
        # combining it with English often strips Vietnamese diacritics.
        selected = ["vie"]

    try:
        import pytesseract
        from pytesseract import Output
    except ImportError as exc:
        raise MissingDependencyError("Thiếu gói pytesseract để nhận dạng PDF scan hoặc hình ảnh.") from exc

    pytesseract.pytesseract.tesseract_cmd = info.command
    if info.tessdata_dir:
        # pytesseract uses shlex(posix=False) on Windows and can pass quote
        # characters literally in --tessdata-dir. The environment variable is
        # inherited by Tesseract without that quoting ambiguity.
        os.environ["TESSDATA_PREFIX"] = info.tessdata_dir
    return pytesseract, Output, selected, warnings


def _image_to_data(
    image,
    requested_language: str,
    *,
    psm: int,
    prefer_vietnamese: bool,
) -> tuple[dict[str, list[Any]], str, list[str]]:
    pytesseract, _output, selected, warnings = _ocr_runtime(
        requested_language,
        prefer_vietnamese=prefer_vietnamese,
    )
    config = (
        f"--oem 1 --psm {int(psm)} "
        "-c preserve_interword_spaces=1 -c hocr_font_info=1"
    )
    try:
        hocr = pytesseract.image_to_pdf_or_hocr(
            image,
            lang="+".join(selected),
            config=config,
            extension="hocr",
        )
        data = _hocr_to_data(hocr)
    except pytesseract.TesseractError as exc:
        raise MissingOcrEngineError(f"Tesseract không thể OCR trang này: {exc}") from exc
    return data, "+".join(selected), warnings


def image_to_data(image, requested_language: str) -> tuple[dict[str, list[Any]], str, list[str]]:
    return _image_to_data(
        image,
        requested_language,
        psm=3,
        prefer_vietnamese=True,
    )


def image_region_to_data(
    image,
    requested_language: str,
    *,
    psm: int = 11,
) -> tuple[dict[str, list[Any]], str, list[str]]:
    """OCR a cleaned table/region with sparse-text segmentation."""

    return _image_to_data(
        image,
        requested_language,
        psm=psm,
        prefer_vietnamese=True,
    )


def detect_image_rotation(image) -> int:
    """Return the counter-clockwise PIL rotation needed to make text upright."""

    info = get_ocr_engine_info()
    if not info.available or not info.command or "osd" not in info.languages:
        return 0
    try:
        import pytesseract
        from pytesseract import Output
    except ImportError:
        return 0

    pytesseract.pytesseract.tesseract_cmd = info.command
    if info.tessdata_dir:
        os.environ["TESSDATA_PREFIX"] = info.tessdata_dir
    try:
        result = pytesseract.image_to_osd(
            image,
            config="--psm 0",
            output_type=Output.DICT,
        )
    except (pytesseract.TesseractError, OSError, ValueError):
        return 0

    try:
        clockwise = int(result.get("rotate", 0)) % 360
    except (TypeError, ValueError):
        return 0
    if clockwise not in {90, 180, 270}:
        return 0
    return (360 - clockwise) % 360
