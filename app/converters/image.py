from __future__ import annotations

from pathlib import Path

from PIL import Image as PillowImage
from PIL import ImageOps, UnidentifiedImageError

from .base import (
    LANDSCAPE_CONTENT_WIDTH_DXA,
    PORTRAIT_CONTENT_WIDTH_DXA,
    ConversionError,
    ConversionOptions,
    ConversionResult,
    add_compact_page_break,
    create_document,
)
from .ocr_layout import normalize_document_ocr_text
from . import raster_tables


def _rgb_image(image: PillowImage.Image) -> PillowImage.Image:
    image = ImageOps.exif_transpose(image)
    if image.mode == "RGBA":
        background = PillowImage.new("RGB", image.size, "white")
        background.paste(image, mask=image.getchannel("A"))
        return background
    if image.mode not in {"RGB", "L"}:
        return image.convert("RGB")
    return image


def convert_image(
    input_path: Path,
    output_path: Path,
    options: ConversionOptions,
) -> ConversionResult:
    options.validate()
    if options.math_ocr:
        from .math_pipeline import convert_image_with_math

        return convert_image_with_math(input_path, output_path, options)
    options.report_progress(2, "Đang đọc hình ảnh…")
    try:
        source = PillowImage.open(input_path)
    except (UnidentifiedImageError, OSError) as exc:
        raise ConversionError(f"Không thể mở hình ảnh: {exc}") from exc

    warnings: list[str] = []
    detected_tables = 0
    try:
        frame_count = max(1, int(getattr(source, "n_frames", 1)))
        source.seek(0)
        first_frame = _rgb_image(source.copy())
        landscape = first_frame.width > first_frame.height
        content_width_dxa = LANDSCAPE_CONTENT_WIDTH_DXA if landscape else PORTRAIT_CONTENT_WIDTH_DXA
        content_width_inches = content_width_dxa / 1440
        doc = create_document(
            input_path.name,
            landscape=landscape,
            body_font="Times New Roman",
        )
        options.report_progress(7, "Đang chuẩn bị OCR hình ảnh…")

        for frame_index in range(frame_count):
            source.seek(frame_index)
            frame = _rgb_image(source.copy())
            options.report_progress(
                8 + (frame_index / frame_count) * 82,
                f"Đang OCR ảnh {frame_index + 1}/{frame_count}…",
            )
            line_count, table_count, frame_warnings = raster_tables.write_editable_ocr_page(
                doc,
                frame,
                options,
                content_width_inches=content_width_inches,
                content_width_dxa=content_width_dxa,
            )
            detected_tables += table_count
            warnings.extend(frame_warnings)
            if line_count == 0 and table_count == 0:
                label = "Ảnh" if frame_count == 1 else f"Khung ảnh {frame_index + 1}"
                warnings.append(f"{label}: OCR không tìm thấy văn bản đủ tin cậy.")
            if options.page_breaks and frame_index < frame_count - 1:
                add_compact_page_break(doc)
            options.report_progress(
                8 + ((frame_index + 1) / frame_count) * 82,
                f"Đã OCR ảnh {frame_index + 1}/{frame_count}.",
            )
    finally:
        source.close()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    options.report_progress(95, "Đang tạo file Word từ hình ảnh…")
    normalize_document_ocr_text(doc)
    try:
        doc.save(output_path)
    except Exception as exc:
        raise ConversionError(f"Không thể tạo file Word: {exc}") from exc
    options.report_progress(100, "Đã tạo xong file Word từ hình ảnh.")
    return ConversionResult(
        output_path=output_path,
        warnings=list(dict.fromkeys(warnings)),
        page_count=frame_count,
        table_count=detected_tables,
    )
