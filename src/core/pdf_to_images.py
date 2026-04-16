import fitz  # pymupdf
import os
from pathlib import Path


def pdf_to_images(pdf_path: str | Path, output_dir: str | Path = None, dpi: int = 300) -> list[Path]:
    """Convert each page of a PDF into a PNG image.

    Args:
        pdf_path: Path to the PDF file.
        output_dir: Directory to save images. Defaults to a folder named
                    '<pdf_stem>_images' next to the PDF.
        dpi: Resolution for rendering (300 recommended for OCR).

    Returns:
        List of Path objects pointing to the generated images.
    """
    pdf_path = Path(pdf_path)
    if output_dir is None:
        output_dir = pdf_path.parent / f"{pdf_path.stem}_images"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    image_paths: list[Path] = []

    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        out_path = output_dir / f"{pdf_path.stem}_page_{i + 1}.png"
        pix.save(str(out_path))
        image_paths.append(out_path)

    doc.close()
    return image_paths


def convert_all_pdfs(pdf_dir: str | Path, output_root: str | Path = None, dpi: int = 300) -> dict[str, list[Path]]:
    """Convert every PDF in a directory to images.

    Args:
        pdf_dir: Directory containing PDF files.
        output_root: Root directory for output. Defaults to '<pdf_dir>/images'.
        dpi: Resolution for rendering.

    Returns:
        Dict mapping PDF filename to its list of generated image paths.
    """
    pdf_dir = Path(pdf_dir)
    if output_root is None:
        output_root = pdf_dir / "images"
    output_root = Path(output_root)

    results: dict[str, list[Path]] = {}
    for pdf_file in sorted(pdf_dir.glob("*.pdf")):
        out_dir = output_root / pdf_file.stem
        results[pdf_file.name] = pdf_to_images(pdf_file, out_dir, dpi)

    return results


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Convert PDF(s) to images for OCR")
    parser.add_argument("--input", help="Path to a single PDF or a directory of PDFs")
    parser.add_argument("--output", "-o", default=None, help="Output directory")
    parser.add_argument("--dpi", type=int, default=300, help="Render DPI (default: 300)")
    args = parser.parse_args()

    input_path = Path(args.input)

    if input_path.is_file() and input_path.suffix.lower() == ".pdf":
        paths = pdf_to_images(input_path, args.output, args.dpi)
        print(f"Converted {input_path.name} -> {len(paths)} image(s)")
        for p in paths:
            print(f"  {p}")
    elif input_path.is_dir():
        all_results = convert_all_pdfs(input_path, args.output, args.dpi)
        for pdf_name, paths in all_results.items():
            print(f"{pdf_name} -> {len(paths)} image(s)")
    else:
        print(f"Error: '{input_path}' is not a valid PDF file or directory.")
