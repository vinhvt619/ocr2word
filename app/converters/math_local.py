"""Local CPU mathematical formula OCR adapter using `pix2tex`.

Provides:
- Lightweight availability inspection (`get_local_math_engine_info`) without
  importing `torch` in the web process.
- Strict checkpoint existence and SHA-256 verification (`ensure_local_math_model_ready`)
  with zero runtime auto-downloads.
- Build-time / setup-time checkpoint fetcher (`download_and_verify_local_math_checkpoints`).
- Thread-safe lazy singleton CPU inference (`recognize_formula_crop`) with
  border-bleed cleanup, 4-sided padding, token/timeout limits, LaTeX post-processing,
  and OMML compilation validation.
"""
from __future__ import annotations

import atexit
import hashlib
import importlib.util
import logging
import math
import multiprocessing
import os
import re
import socket
import threading
import urllib.request
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

import cv2
import numpy as np
from PIL import Image

from . import office_math
from .base import ConversionError, MissingDependencyError

logger = logging.getLogger(__name__)

# Prevent albumentations from making network calls to PyPI on import
os.environ.setdefault("NO_ALBUMENTATIONS_UPDATE", "1")

PIX2TEX_PACKAGE_VERSION = "0.1.4"
PIX2TEX_RELEASE_TAG = "v0.0.1"
WEIGHTS_URL = f"https://github.com/lukas-blecher/LaTeX-OCR/releases/download/{PIX2TEX_RELEASE_TAG}/weights.pth"
RESIZER_URL = f"https://github.com/lukas-blecher/LaTeX-OCR/releases/download/{PIX2TEX_RELEASE_TAG}/image_resizer.pth"

WEIGHTS_SHA256 = "a63d9141c53d266cb682fb5a8bd83bd5cbe283145e0e78ebdc0f895195a1dfaa"
WEIGHTS_SIZE_BYTES = 102113875

RESIZER_SHA256 = "1c3820659985ad142b526490bb25c23d977176ac2073591b3bddada692718458"
RESIZER_SIZE_BYTES = 19441973

TOKENIZER_SHA256 = "1dc27b18d6a518d0d5ff3f4bb7bd98521fe80ad39e5b2a246d4109f1bb9d5019"
TOKENIZER_SIZE_BYTES = 24174

MAX_CROP_WIDTH = 1600
MAX_CROP_HEIGHT = 640
DEFAULT_MAX_TOKENS = 128
DEFAULT_TIMEOUT_SECONDS = 15.0
MODEL_LOAD_TIMEOUT_SECONDS = 60.0

_MODEL_LOCK = threading.RLock()
_MODEL_INSTANCE: _LocalMathProcess | None = None
_VERIFIED_CHECKPOINT_DIRS: set[str] = set()

_REPETITION_LOOP_RE = re.compile(r"(\\[a-zA-Z]+|[a-zA-Z0-9])\s*(?:\1\s*){5,}")

_SUSPICIOUS_MACRO_PATTERNS = (
    re.compile(r"\\operatorname\*\{(?!arg\s*max|arg\s*min|lim|max|min|sup|inf)[^}]+\}"),  # unknown operator hallucination
    re.compile(r"\\mathfrak\{"),                          # Fraktur hallucination on blurred crops
    re.compile(r"\\stackrel\{[^}]*\\stackrel"),           # nested stacked symbols
    re.compile(r"_\{[-+=*]\}"),                           # empty subscript with operator
    re.compile(r"\^\{[-+=*]\}"),                          # empty superscript with operator
)


def strip_presentation_wrappers(latex: str) -> str:
    """Remove purely presentation wrappers from pix2tex output.

    Allowed by user decision (mục 6.2 & mục 9):
    - 1x1 array with single cell: \\begin{array}{l}{X}\\end{array} -> X
    - Presentation sizing: \\displaystyle, \\scriptstyle, \\textstyle, \\scriptscriptstyle
    - Latin/digit font wrappers: \\mathbf{x}, \\mathrm{x}, \\boldsymbol{x}, {\\bf x} -> x

    FORBIDDEN:
    - Never touch \\mathfrak, \\mathcal, \\mathbb, Greek letters, or case.
    """
    s = (latex or "").strip()
    if not s:
        return ""

    # 1. Strip \\displaystyle, \\scriptstyle, \\textstyle, \\scriptscriptstyle
    s = re.sub(r"\\(?:displaystyle|scriptstyle|textstyle|scriptscriptstyle)\b\s*", "", s)

    # 2. Strip single-cell array: \\begin{array}{...} ... \\end{array}
    # Only if it does NOT contain row breaks (\\\\) or column separators (&)
    array_match = re.search(r"\\begin\{array\}\{[^{}]*\}(.*?)\\end\{array\}", s, flags=re.DOTALL)
    if array_match:
        inner = array_match.group(1).strip()
        if r"\\" not in inner and "&" not in inner:
            if inner.startswith("{") and inner.endswith("}"):
                inner = inner[1:-1].strip()
            s = s[: array_match.start()] + inner + s[array_match.end() :]

    # 3. Strip \\mathbf{...}, \\mathrm{...}, \\boldsymbol{...} for Latin letters, numbers, basic operators
    def _unwrap_font(m):
        content = m.group(1)
        if any(
            tok in content
            for tok in (
                r"\mathfrak", r"\mathcal", r"\mathbb",
                r"\alpha", r"\beta", r"\gamma", r"\delta", r"\epsilon", r"\varepsilon",
                r"\zeta", r"\eta", r"\theta", r"\vartheta", r"\iota", r"\kappa",
                r"\lambda", r"\mu", r"\nu", r"\xi", r"\pi", r"\rho", r"\sigma",
                r"\tau", r"\upsilon", r"\phi", r"\varphi", r"\chi", r"\psi", r"\omega",
                r"\Gamma", r"\Delta", r"\Theta", r"\Lambda", r"\Xi", r"\Pi", r"\Sigma",
                r"\Upsilon", r"\Phi", r"\Psi", r"\Omega",
            )
        ):
            return m.group(0)
        return content

    while True:
        prev_s = s
        for macro in (r"\mathbf", r"\mathrm", r"\boldsymbol"):
            pat = re.escape(macro) + r"\{([^{}]+)\}"
            while True:
                new_s = re.sub(pat, _unwrap_font, s)
                if new_s == s:
                    break
                s = new_s
        if s == prev_s:
            break

    # {\\bf ...}
    def _unwrap_bf(m):
        content = m.group(1)
        if any(tok in content for tok in (r"\mathfrak", r"\mathcal", r"\mathbb", r"\alpha", r"\beta", r"\gamma")):
            return m.group(0)
        return content

    s = re.sub(r"\{\\bf\s+([^{}]+)\}", _unwrap_bf, s)
    return s.strip()


def is_suspicious_latex(latex: str) -> tuple[bool, str]:
    """Check if predicted LaTeX contains suspicious hallucinated patterns."""
    if not latex:
        return True, "empty_latex"
    for pat in _SUSPICIOUS_MACRO_PATTERNS:
        if pat.search(latex):
            return True, f"suspicious_macro_{pat.pattern}"
    if latex.count("{") != latex.count("}"):
        return True, "unbalanced_braces"
    if latex.count("(") != latex.count(")"):
        return True, "unbalanced_parentheses"
    if re.search(r"\\(?:frac|dfrac|tfrac|cfrac)\s*\{\s*\}\s*\{", latex):
        return True, "empty_numerator"
    if re.search(r"\\(?:frac|dfrac|tfrac|cfrac)\s*\{[^{}]*\}\s*\{\s*\}", latex):
        return True, "empty_denominator"
    return False, ""


@dataclass(frozen=True, slots=True)
class LocalMathEngineInfo:
    available: bool
    package_installed: bool
    weights_available: bool
    model_name: str = f"pix2tex:{PIX2TEX_RELEASE_TAG}"
    reason: str | None = None

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


def _resolve_pix2tex_package_dir() -> Path | None:
    try:
        spec = importlib.util.find_spec("pix2tex")
    except (ImportError, ValueError):
        return None
    if spec is None or not spec.origin:
        return None
    return Path(spec.origin).resolve().parent


def resolve_local_math_paths() -> dict[str, Path] | None:
    """Return paths to config.yaml, tokenizer.json, weights.pth, and image_resizer.pth without importing torch."""
    pkg_dir = _resolve_pix2tex_package_dir()
    if pkg_dir is None:
        return None
    custom_dir = (
        os.environ.get("LOCAL_MATH_MODEL_DIR", "").strip()
        or os.environ.get("OCR2WORD_MATH_MODEL_DIR", "").strip()
    )
    ckpt_dir = Path(custom_dir).resolve() if custom_dir else (pkg_dir / "model" / "checkpoints")
    return {
        "ckpt_dir": ckpt_dir,
        "weights": ckpt_dir / "weights.pth",
        "resizer": ckpt_dir / "image_resizer.pth",
        "config": pkg_dir / "model" / "settings" / "config.yaml",
        "tokenizer": pkg_dir / "model" / "dataset" / "tokenizer.json",
    }


def get_local_math_engine_info() -> LocalMathEngineInfo:
    """Lightweight status probe that never imports `torch` or loads model weights."""
    required_modules = (
        "torch", "torchvision", "pix2tex", "transformers", "tokenizers",
        "timm", "einops", "munch", "entmax", "yaml",
    )
    missing_modules = []
    for module in required_modules:
        try:
            installed = importlib.util.find_spec(module) is not None
        except (ImportError, ValueError):
            installed = False
        if not installed:
            missing_modules.append(module)

    if missing_modules:
        return LocalMathEngineInfo(
            available=False,
            package_installed=False,
            weights_available=False,
            reason=f"Thiếu gói mô hình toán cục bộ: {', '.join(missing_modules)}.",
        )

    paths = resolve_local_math_paths()
    if paths is None:
        return LocalMathEngineInfo(
            available=False,
            package_installed=True,
            weights_available=False,
            reason="Không tìm thấy thư mục cấu hình pix2tex.",
        )

    weights = paths["weights"]
    resizer = paths["resizer"]
    tokenizer = paths["tokenizer"]
    config = paths["config"]

    if not (weights.is_file() and resizer.is_file() and tokenizer.is_file() and config.is_file()):
        return LocalMathEngineInfo(
            available=False,
            package_installed=True,
            weights_available=False,
            reason="Thiếu file trọng số mô hình công thức cục bộ (weights.pth / image_resizer.pth).",
        )

    if (
        weights.stat().st_size != WEIGHTS_SIZE_BYTES
        or resizer.stat().st_size != RESIZER_SIZE_BYTES
        or tokenizer.stat().st_size != TOKENIZER_SIZE_BYTES
    ):
        return LocalMathEngineInfo(
            available=False,
            package_installed=True,
            weights_available=False,
            reason="Kích thước file trọng số mô hình công thức cục bộ không khớp chuẩn.",
        )

    return LocalMathEngineInfo(
        available=True,
        package_installed=True,
        weights_available=True,
        reason=None,
    )


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def ensure_local_math_model_ready() -> dict[str, Path]:
    """Verify local math packages and SHA-256 checkpoint integrity before inference. Never auto-downloads."""
    info = get_local_math_engine_info()
    if not info.package_installed:
        raise MissingDependencyError(
            f"{info.reason} Cần cài đủ requirements-math-local.txt và PyTorch CPU. "
            "Hãy cài đặt hoặc chọn chế độ 'Theo quy tắc PDF'."
        )
    paths = resolve_local_math_paths()
    if paths is None or not info.weights_available:
        raise ConversionError(
            info.reason
            or "Mô hình công thức cục bộ chưa sẵn sàng. Hãy tải trước trọng số hoặc chọn 'Theo quy tắc PDF'."
        )

    ckpt_key = str(paths["ckpt_dir"])
    if ckpt_key not in _VERIFIED_CHECKPOINT_DIRS:
        w_hash = _sha256_file(paths["weights"])
        r_hash = _sha256_file(paths["resizer"])
        t_hash = _sha256_file(paths["tokenizer"])
        if w_hash != WEIGHTS_SHA256 or r_hash != RESIZER_SHA256 or t_hash != TOKENIZER_SHA256:
            raise ConversionError(
                "Mã băm SHA-256 của mô hình công thức cục bộ không khớp chuẩn an toàn. "
                "Đã từ chối nạp trọng số."
            )
        _VERIFIED_CHECKPOINT_DIRS.add(ckpt_key)
    return paths


def download_and_verify_local_math_checkpoints(target_dir: Path | None = None) -> Path:
    """Explicit build-time / setup helper to fetch and verify pix2tex v0.0.1 weights."""
    paths = resolve_local_math_paths()
    if paths is None:
        raise RuntimeError("pix2tex package is not installed; cannot resolve checkpoint directory.")
    ckpt_dir = target_dir.resolve() if target_dir is not None else paths["ckpt_dir"]
    ckpt_dir.mkdir(parents=True, exist_ok=True)

    specs = (
        (ckpt_dir / "weights.pth", WEIGHTS_URL, WEIGHTS_SHA256, WEIGHTS_SIZE_BYTES),
        (ckpt_dir / "image_resizer.pth", RESIZER_URL, RESIZER_SHA256, RESIZER_SIZE_BYTES),
    )
    for dest, url, expected_sha, expected_size in specs:
        if dest.is_file() and dest.stat().st_size == expected_size and _sha256_file(dest) == expected_sha:
            continue
        tmp_dest = dest.with_suffix(".pth.part")
        urllib.request.urlretrieve(url, tmp_dest)
        actual_size = tmp_dest.stat().st_size
        actual_sha = _sha256_file(tmp_dest)
        if actual_size != expected_size or actual_sha != expected_sha:
            tmp_dest.unlink(missing_ok=True)
            raise RuntimeError(
                f"Checksum mismatch for {dest.name}: expected {expected_sha} ({expected_size}B), "
                f"got {actual_sha} ({actual_size}B)."
            )
        tmp_dest.replace(dest)
    return ckpt_dir


def _minmax_size(
    img: Image.Image,
    max_dimensions: tuple[int, int] | list[int] | None = None,
    min_dimensions: tuple[int, int] | list[int] | None = None,
) -> Image.Image:
    if max_dimensions is not None:
        ratios = [a / b for a, b in zip(img.size, max_dimensions)]
        if any(r > 1 for r in ratios):
            size = np.array(img.size) // max(ratios)
            img = img.resize((max(1, int(size[0])), max(1, int(size[1]))), Image.Resampling.BILINEAR)
    if min_dimensions is not None:
        padded_size = [max(img_dim, min_dim) for img_dim, min_dim in zip(img.size, min_dimensions)]
        if padded_size != list(img.size):
            padded_im = Image.new("L", padded_size, 255)
            bbox = img.getbbox()
            if bbox:
                padded_im.paste(img, bbox)
            img = padded_im
    return img


def prepare_formula_crop(
    img: Image.Image,
    *,
    divable: int = 32,
    border: int = 0,
    strip_border_bleed: bool = False,
) -> Image.Image | None:
    """Normalize a formula crop to grayscale, remove adjacent-line border bleed, and pad cleanly."""
    if img.width <= 1 or img.height <= 1:
        return None
    if img.width > MAX_CROP_WIDTH or img.height > MAX_CROP_HEIGHT:
        scale = min(MAX_CROP_WIDTH / img.width, MAX_CROP_HEIGHT / img.height)
        img = img.resize(
            (max(2, round(img.width * scale)), max(2, round(img.height * scale))),
            Image.Resampling.LANCZOS,
        )

    threshold = 128
    data = np.array(img.convert("LA"))
    if data[..., -1].var() == 0:
        data = data[..., 0].astype(np.uint8)
    else:
        data = (255 - data[..., -1]).astype(np.uint8)
    dmin, dmax = int(data.min()), int(data.max())
    if dmax <= dmin:
        return None
    data = ((data.astype(np.float32) - dmin) / (dmax - dmin) * 255.0).astype(np.uint8)
    if data.mean() > threshold:
        gray = 255 * (data < threshold).astype(np.uint8)
    else:
        gray = 255 * (data > threshold).astype(np.uint8)
        data = 255 - data

    # Strip tiny cut-off slivers touching border from adjacent text lines without deleting exponents or dots
    if strip_border_bleed and gray.shape[0] >= 18:
        h_img, w_img = gray.shape[:2]
        num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(gray, connectivity=8)
        # When ambiguous, preserve source stroke ("khi mơ hồ giữ nét nguồn")
        for idx in range(1, num_labels):
            top = int(stats[idx, cv2.CC_STAT_TOP])
            comp_h = int(stats[idx, cv2.CC_STAT_HEIGHT])
            comp_w = int(stats[idx, cv2.CC_STAT_WIDTH])
            bottom = top + comp_h
            touches_edge = top == 0 or bottom >= h_img - 1
            is_thin_sliver = touches_edge and (comp_h <= 2 or (comp_h <= 3 and comp_w >= 12))
            if is_thin_sliver and num_labels > 2:
                gray[labels == idx] = 0
                data[labels == idx] = 255

    coords = cv2.findNonZero(gray)
    if coords is None:
        return None
    a, b, w, h = cv2.boundingRect(coords)
    if w <= 0 or h <= 0:
        return None
    rect = data[b : b + h, a : a + w]
    im = Image.fromarray(rect).convert("L")

    total_w = w + border * 2
    total_h = h + border * 2
    dims = []
    for x in (total_w, total_h):
        div, mod = divmod(x, divable)
        dims.append(divable * (div + (1 if mod > 0 else 0)))
    padded = Image.new("L", dims, 255)
    offset_x = border
    offset_y = border
    padded.paste(im, (offset_x, offset_y, offset_x + w, offset_y + h))
    return padded


def clean_predicted_latex(raw_latex: str) -> str:
    """Remove outer delimiters only; never guess away mathematical content.

    A matrix is not necessarily a fraction, an overscript may be meaningful,
    and floor delimiters/macros must survive even if the writer rejects them.
    Reject incomplete predictions instead of flattening them into valid-looking
    but different formulas. OMML support is checked separately by the caller.
    """
    s = office_math.strip_outer_math_delimiters(raw_latex or "").strip()
    if not s:
        return ""
    if _REPETITION_LOOP_RE.search(s):
        return ""

    if re.search(r"\\(?:frac|dfrac|tfrac|cfrac)\s*\{\s*\}", s):
        return ""
    if re.search(r"\\(?:frac|dfrac|tfrac|cfrac)\s*\{[^{}]*\}\s*\{\s*\}", s):
        return ""
    return s


class _LocalLatexOCR:
    """Thread-safe CPU wrapper around pix2tex encoder/decoder and image resizer."""

    def __init__(self, paths: dict[str, Path]) -> None:
        import torch
        import yaml
        from munch import Munch
        from timm.models.layers import StdConv2dSame
        from timm.models.resnetv2 import ResNetV2
        from transformers import PreTrainedTokenizerFast
        from pix2tex.models import get_model
        from pix2tex.utils import parse_args, post_process, token2str

        try:
            torch.set_num_threads(1)
            torch.set_num_interop_threads(1)
        except RuntimeError:
            pass
        self._torch = torch
        self._post_process = post_process
        self._token2str = token2str

        with open(paths["config"], "r", encoding="utf-8") as handle:
            params = yaml.safe_load(handle)
        self.args = parse_args(Munch(params), no_cuda=True)
        self.args.update(
            config=str(paths["config"]),
            checkpoint=str(paths["weights"]),
            tokenizer=str(paths["tokenizer"]),
            no_cuda=True,
            no_resize=False,
            wandb=False,
            device="cpu",
            temperature=0.25,
        )
        self.model = get_model(self.args)
        self.model.load_state_dict(
            torch.load(paths["weights"], map_location="cpu", weights_only=True)
        )
        self.model.eval()

        self.image_resizer = ResNetV2(
            layers=[2, 3, 3],
            num_classes=max(self.args.max_dimensions) // 32,
            global_pool="avg",
            in_chans=1,
            drop_rate=0.05,
            preact=True,
            stem_type="same",
            conv_layer=StdConv2dSame,
        ).to("cpu")
        self.image_resizer.load_state_dict(
            torch.load(paths["resizer"], map_location="cpu", weights_only=True)
        )
        self.image_resizer.eval()
        self.tokenizer = PreTrainedTokenizerFast(tokenizer_file=str(paths["tokenizer"]))

    def _to_tensor(self, img: Image.Image):
        gray = np.asarray(img.convert("L"), dtype=np.float32) / 255.0
        norm = (gray - 0.7931) / 0.1738
        return self._torch.from_numpy(norm).unsqueeze(0).unsqueeze(0)

    def _greedy_decode(self, context, max_tokens: int) -> tuple[Any, float]:
        """Deterministic greedy argmax decoding with average token log-probability."""
        import torch.nn.functional as F

        start_tokens = self._torch.LongTensor([self.args.bos_token])[:, None]
        decoder = self.model.decoder
        seq_len = min(max_tokens, int(self.args.max_seq_len))
        eos_token = self.args.eos_token

        was_training = decoder.net.training
        decoder.net.eval()

        out = start_tokens
        mask = self._torch.full_like(out, True, dtype=self._torch.bool, device=out.device)
        log_probs: list[float] = []

        for _ in range(seq_len):
            x = out[:, -decoder.max_seq_len:]
            mask_slice = mask[:, -decoder.max_seq_len:]
            logits = decoder.net(x, mask=mask_slice, context=context)[:, -1, :]
            log_sm = F.log_softmax(logits, dim=-1)
            next_token = self._torch.argmax(logits, dim=-1, keepdim=True)
            token_lp = log_sm.gather(dim=-1, index=next_token).item()
            log_probs.append(token_lp)

            out = self._torch.cat((out, next_token), dim=-1)
            mask = F.pad(mask, (0, 1), value=True)

            if eos_token is not None and (self._torch.cumsum(out == eos_token, 1)[:, -1] >= 1).all():
                break

        out = out[:, 1:]
        decoder.net.train(was_training)
        avg_lp = sum(log_probs) / max(1, len(log_probs)) if log_probs else 0.0
        return out, avg_lp

    def predict(self, crop_image: Image.Image, *, max_tokens: int = DEFAULT_MAX_TOKENS) -> str:
        crop_w, crop_h = crop_image.size
        # Minimum physical crop size check (reject tiny stray specks)
        if crop_w < 12 or crop_h < 8 or (crop_w * crop_h < 120):
            return ""
        # Vertical sliver rejection (narrow line border bleed or fragment)
        if crop_w < 32 and crop_h > 70:
            return ""

        padded = prepare_formula_crop(crop_image)
        if padded is None:
            return ""
        img = _minmax_size(padded, self.args.max_dimensions, self.args.min_dimensions)
        with self._torch.no_grad():
            input_image = img.convert("RGB").copy()
            r, w, h = 1.0, input_image.size[0], input_image.size[1]
            t = self._to_tensor(img)
            for _ in range(10):
                h = int(h * r)
                resized = input_image.resize(
                    (w, h),
                    Image.Resampling.BILINEAR if r > 1 else Image.Resampling.LANCZOS,
                )
                repadded = prepare_formula_crop(resized, strip_border_bleed=False)
                if repadded is None:
                    break
                img = _minmax_size(repadded, self.args.max_dimensions, self.args.min_dimensions)
                t = self._to_tensor(img)
                w = (self.image_resizer(t).argmax(-1).item() + 1) * 32
                if w == img.size[0]:
                    break
                r = w / max(1, img.size[0])

            context = self.model.encoder(t)
            dec, avg_lp = self._greedy_decode(context, max_tokens=max_tokens)
            # Low confidence gate: if average token log-prob < -1.80, reject to image
            if avg_lp < -1.80:
                return ""

            raw = self._post_process(self._token2str(dec, self.tokenizer)[0])
            cleaned = clean_predicted_latex(raw)
            normalized = strip_presentation_wrappers(cleaned)
            suspicious, _ = is_suspicious_latex(normalized)
            if suspicious:
                return ""

            # Aspect ratio gate on decoded text length:
            # If crop is very wide (>= 250px) but cleaned output is < 3 characters (e.g. just 'x' or '1')
            if crop_w >= 250 and len(normalized) < 3:
                return ""
            return normalized


def _deny_worker_network(*args, **kwargs):
    raise ConversionError("Mô hình cục bộ không được phép kết nối mạng khi nhận diện.")


def _local_math_worker(connection, paths: dict[str, Path]) -> None:
    """One private, reusable CPU process; no document/model data from a socket."""
    try:
        # Spawn does not inherit a caller's network mocks. Enforce the offline
        # contract inside this dedicated process, including model startup.
        os.environ["HF_HUB_OFFLINE"] = "1"
        os.environ["TRANSFORMERS_OFFLINE"] = "1"
        socket.create_connection = _deny_worker_network
        socket.socket.connect = _deny_worker_network
        socket.socket.connect_ex = _deny_worker_network
        engine = _LocalLatexOCR(paths)
        connection.send(("ready", None))
        while True:
            image, max_tokens = connection.recv()
            try:
                connection.send(("result", engine.predict(image, max_tokens=max_tokens)))
            except Exception as exc:
                # Never send raw provider/library messages containing source data.
                connection.send(("error", type(exc).__name__))
    except (EOFError, BrokenPipeError):
        pass
    except Exception as exc:
        try:
            connection.send(("error", type(exc).__name__))
        except (EOFError, BrokenPipeError, OSError):
            pass
    finally:
        connection.close()


class _LocalMathProcess:
    """Persistent spawn worker, terminated/reaped on timeout, then restartable.

    Unlike ThreadPoolExecutor shutdown, this can actually stop a stuck native
    inference. A successful crop reuses the loaded model for the next request.
    """

    def __init__(self) -> None:
        self._process = None
        self._connection = None

    def close(self) -> None:
        if self._connection is not None:
            self._connection.close()
            self._connection = None
        if self._process is not None:
            if self._process.is_alive():
                self._process.terminate()
                self._process.join(timeout=1)
            if self._process.is_alive():
                self._process.kill()
                self._process.join(timeout=1)
            if not self._process.is_alive():
                self._process.close()
                self._process = None

    def _receive(self, timeout: float, message: str):
        try:
            if not self._connection.poll(timeout):
                raise ConversionError(message)
            kind, payload = self._connection.recv()
        except (EOFError, OSError) as exc:
            self.close()
            raise ConversionError("Tiến trình nhận diện công thức cục bộ đã dừng bất thường.") from exc
        except ConversionError:
            self.close()
            raise
        if kind == "error":
            self.close()
            raise ConversionError("Mô hình công thức cục bộ không thể xử lý yêu cầu.")
        return kind, payload

    def start(self) -> None:
        if self._process is not None and self._process.is_alive():
            return
        self.close()
        paths = ensure_local_math_model_ready()
        context = multiprocessing.get_context("spawn")
        parent, child = context.Pipe()
        self._connection = parent
        self._process = context.Process(target=_local_math_worker, args=(child, paths), daemon=True)
        try:
            self._process.start()
        except Exception:
            parent.close()
            self._connection = self._process = None
            raise ConversionError("Không thể khởi động tiến trình nhận diện công thức cục bộ.") from None
        finally:
            child.close()
        kind, _ = self._receive(MODEL_LOAD_TIMEOUT_SECONDS, "Quá thời gian nạp mô hình công thức cục bộ.")
        if kind != "ready":
            self.close()
            raise ConversionError("Mô hình công thức cục bộ chưa sẵn sàng.")

    def predict(self, image: Image.Image, *, max_tokens: int, timeout_seconds: float) -> str:
        self.start()
        try:
            self._connection.send((image, max_tokens))
        except (OSError, EOFError) as exc:
            self.close()
            raise ConversionError("Không thể gửi vùng công thức tới mô hình cục bộ.") from exc
        kind, latex = self._receive(timeout_seconds, "Quá thời gian nhận diện vùng công thức cục bộ.")
        if kind != "result" or not isinstance(latex, str):
            self.close()
            raise ConversionError("Kết quả mô hình công thức cục bộ không hợp lệ.")
        return latex


def get_local_math_model() -> _LocalMathProcess:
    global _MODEL_INSTANCE
    with _MODEL_LOCK:
        if _MODEL_INSTANCE is None:
            _MODEL_INSTANCE = _LocalMathProcess()
        _MODEL_INSTANCE.start()
        return _MODEL_INSTANCE


_get_model_instance = get_local_math_model


@atexit.register
def _shutdown_local_math_model() -> None:
    if _MODEL_INSTANCE is not None:
        _MODEL_INSTANCE.close()


def recognize_formula_crop(
    image: Image.Image,
    *,
    timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS,
    max_tokens: int = DEFAULT_MAX_TOKENS,
) -> str:
    """Recognize a cropped formula image into validated LaTeX ready for OMML compilation."""
    if not math.isfinite(timeout_seconds) or timeout_seconds <= 0:
        raise ConversionError("Thời gian nhận diện công thức phải hữu hạn và lớn hơn 0.")
    if not 1 <= max_tokens <= 512:
        raise ConversionError("Giới hạn độ dài nhận diện công thức không hợp lệ.")
    with _MODEL_LOCK:
        engine = get_local_math_model()
        latex = engine.predict(image, max_tokens=max_tokens, timeout_seconds=timeout_seconds)

    if not latex:
        raise ConversionError("Mô hình cục bộ không nhận diện được công thức hợp lệ.")

    # Validate that the predicted LaTeX compiles to a valid OMML element
    try:
        office_math.latex_to_omml_element(latex)
    except Exception as exc:
        raise ConversionError("Chưa thể dựng công thức từ kết quả mô hình cục bộ.") from exc

    return latex
