"""Document conversion backends."""

from pathlib import Path

from .base import ConversionError, ConversionOptions, ConversionResult


def convert_excel(input_path: Path, output_path: Path, options: ConversionOptions) -> ConversionResult:
    # Keep backend imports lazy so app.services.ocr can import shared errors
    # without triggering the PDF -> OCR import cycle.
    from .excel import convert_excel as convert_excel_impl

    return convert_excel_impl(input_path, output_path, options)


def convert_pdf(input_path: Path, output_path: Path, options: ConversionOptions) -> ConversionResult:
    from .pdf import convert_pdf as convert_pdf_impl

    return convert_pdf_impl(input_path, output_path, options)


def analyze_pdf(input_path: Path, options: ConversionOptions):
    from .pdf import analyze_pdf as analyze_pdf_impl

    return analyze_pdf_impl(input_path, options)


def extract_pdf_pages(input_path: Path, output_path: Path, *, start_page: int, end_page: int) -> int:
    from .pdf import extract_pdf_pages as extract_pdf_pages_impl

    return extract_pdf_pages_impl(
        input_path,
        output_path,
        start_page=start_page,
        end_page=end_page,
    )


def convert_image(input_path: Path, output_path: Path, options: ConversionOptions) -> ConversionResult:
    from .image import convert_image as convert_image_impl

    return convert_image_impl(input_path, output_path, options)


def convert_to_excel(input_path: Path, output_path: Path, options: ConversionOptions) -> ConversionResult:
    from .to_excel import convert_to_excel as convert_to_excel_impl

    return convert_to_excel_impl(input_path, output_path, options)


def validate_pdf_signature(input_path: Path, signature: bytes | None = None) -> None:
    from .pdf import validate_pdf_signature as validate_pdf_signature_impl

    validate_pdf_signature_impl(input_path, signature)


__all__ = [
    "ConversionError",
    "ConversionOptions",
    "ConversionResult",
    "analyze_pdf",
    "convert_excel",
    "convert_image",
    "convert_pdf",
    "convert_to_excel",
    "extract_pdf_pages",
    "validate_pdf_signature",
]
