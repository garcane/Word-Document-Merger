#!/usr/bin/env python3
"""
OCR to Markdown Converter — Arabic Documents
Converts page images to Arabic Markdown using Tesseract (free, local, no API key needed).
Output .md files are ready to paste into ChatGPT or Gemini for translation.

Usage:
    python ocr_to_markdown.py ./images
    python ocr_to_markdown.py ./images --output ./markdown
    python ocr_to_markdown.py ./images --output ./markdown --pages 1 3 5

Dependencies:
    pip install pytesseract Pillow tqdm

    Plus Tesseract with Arabic language data:
        macOS:          brew install tesseract tesseract-lang
        Ubuntu/Debian:  sudo apt install tesseract-ocr tesseract-ocr-ara
        Windows:        https://github.com/UB-Mannheim/tesseract/wiki
"""

import argparse
import sys
from pathlib import Path
from tqdm import tqdm


# ─────────────────────────────────────────────
# Tesseract Check
# ─────────────────────────────────────────────

def check_dependencies():
    """Verify pytesseract and Arabic language data are available."""
    try:
        import pytesseract
    except ImportError:
        print("❌ pytesseract not installed. Run: pip install pytesseract")
        sys.exit(1)

    try:
        langs = pytesseract.get_languages()
        if "ara" not in langs:
            print("❌ Arabic language data not found in Tesseract.")
            print("   macOS:         brew install tesseract-lang")
            print("   Ubuntu/Debian: sudo apt install tesseract-ocr-ara")
            sys.exit(1)
    except Exception as e:
        print(f"❌ Could not connect to Tesseract: {e}")
        print("   Make sure Tesseract is installed and on your PATH.")
        sys.exit(1)


# ─────────────────────────────────────────────
# Image Loading
# ─────────────────────────────────────────────

def get_sorted_images(input_dir: Path, page_filter: list[int] | None) -> list[Path]:
    """Return sorted list of image paths, optionally filtered by page numbers."""
    images = sorted(
        [p for p in input_dir.iterdir() if p.suffix.lower() in (".png", ".jpg", ".jpeg")],
        key=lambda p: int("".join(filter(str.isdigit, p.stem)) or 0)
    )
    if page_filter:
        images = [img for img in images if any(str(n) in img.stem for n in page_filter)]
    return images


# ─────────────────────────────────────────────
# OCR
# ─────────────────────────────────────────────

def ocr_image_to_arabic_text(image_path: Path, psm: int = 3) -> str:
    """
    Run Tesseract OCR on an image with Arabic settings.
    Returns raw extracted Arabic text.
    """
    import pytesseract
    from PIL import Image

    img = Image.open(image_path)

    # PSM 3 = fully automatic page segmentation (best for mixed layouts)
    # PSM 6 = assume a single uniform block of text (use for dense prose)
    config = f"--psm {psm} --oem 1"

    text = pytesseract.image_to_string(img, lang="ara", config=config)
    return text.strip()


# ─────────────────────────────────────────────
# Markdown Formatting
# ─────────────────────────────────────────────

def format_as_markdown(page_name: str, arabic_text: str) -> str:
    """
    Wrap the raw OCR output in clean Markdown with RTL metadata.
    Preserves paragraph breaks; adds a page heading for navigation.
    """
    if not arabic_text:
        return f"# {page_name}\n\n> ⚠️ No text extracted from this page. The image may be blank or unreadable.\n"

    # Split on blank lines to preserve paragraph structure
    paragraphs = [p.strip() for p in arabic_text.split("\n\n") if p.strip()]
    body = "\n\n".join(paragraphs)

    return (
        f"# {page_name}\n\n"
        f"<!-- lang: ar | direction: rtl -->\n\n"
        f"{body}\n"
    )


# ─────────────────────────────────────────────
# Summary Report
# ─────────────────────────────────────────────

def save_summary_report(output_dir: Path, results: dict[str, bool]):
    """
    Save a brief processing summary so you know which pages need attention.
    """
    report_path = output_dir / "_ocr_summary.md"

    success = [p for p, ok in results.items() if ok]
    failed  = [p for p, ok in results.items() if not ok]

    lines = [
        "# OCR Processing Summary\n",
        f"- **Total pages:** {len(results)}",
        f"- **Succeeded:** {len(success)}",
        f"- **Failed / empty:** {len(failed)}\n",
    ]

    if failed:
        lines.append("## Pages Needing Attention\n")
        for page in failed:
            lines.append(f"- ⚠️ {page}")
        lines.append(
            "\nFor these pages, check the source image quality and consider "
            "increasing DPI when running pdf_converter.py (try --dpi 300).\n"
        )

    lines.append("## Next Step\n")
    lines.append(
        "Upload individual `.md` files to ChatGPT or Gemini and use a prompt such as:\n\n"
        "> *Translate the following Arabic text to English. "
        "Preserve the Markdown structure. "
        "For any domain-specific terms where multiple translations are plausible, "
        "note the alternatives in brackets.*\n"
    )

    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path


# ─────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Convert Arabic document images to Markdown using Tesseract (free, local).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process all pages in a folder
  python ocr_to_markdown.py ./images

  # Custom output directory
  python ocr_to_markdown.py ./images --output ./markdown

  # Process specific pages only
  python ocr_to_markdown.py ./images --pages 1 3 5

  # Use PSM 6 for dense single-column prose
  python ocr_to_markdown.py ./images --psm 6
        """
    )

    parser.add_argument("input_dir", type=str,
                        help="Directory of page images (output from pdf_converter.py)")
    parser.add_argument("-o", "--output", type=str, default=None,
                        help="Output directory for .md files (default: ./markdown)")
    parser.add_argument("--pages", type=int, nargs="+", default=None,
                        help="Only process these page numbers (e.g. --pages 1 3 5)")
    parser.add_argument("--psm", type=int, default=3, choices=[3, 4, 6],
                        help="Tesseract page segmentation mode: 3=auto, 4=single column, 6=single block (default: 3)")

    args = parser.parse_args()

    # ── Validate input
    input_dir = Path(args.input_dir)
    if not input_dir.exists():
        print(f"❌ Input directory not found: {input_dir}")
        sys.exit(1)

    output_dir = Path(args.output) if args.output else Path("./markdown")
    output_dir.mkdir(parents=True, exist_ok=True)

    # ── Check Tesseract
    check_dependencies()

    # ── Load images
    images = get_sorted_images(input_dir, args.pages)
    if not images:
        print(f"❌ No images found in: {input_dir}")
        sys.exit(1)

    # ── Print settings
    print("\n" + "=" * 50)
    print("  OCR to Markdown — Arabic Document Pipeline")
    print("=" * 50)
    print(f"  Input:    {input_dir}")
    print(f"  Output:   {output_dir}")
    print(f"  Pages:    {len(images)}")
    print(f"  PSM mode: {args.psm}")
    print("=" * 50 + "\n")

    results = {}

    for image_path in tqdm(images, desc="OCR progress", unit="page"):
        page_name = image_path.stem

        try:
            arabic_text = ocr_image_to_arabic_text(image_path, psm=args.psm)
            markdown = format_as_markdown(page_name, arabic_text)

            output_path = output_dir / f"{page_name}.md"
            output_path.write_text(markdown, encoding="utf-8")

            results[page_name] = bool(arabic_text)

        except Exception as e:
            tqdm.write(f"  ❌ Error on {page_name}: {e}")
            results[page_name] = False

    # ── Summary
    report_path = save_summary_report(output_dir, results)

    succeeded = sum(results.values())
    print(f"\n✅ Done. {succeeded}/{len(results)} pages extracted successfully.")
    print(f"📄 Summary saved to: {report_path}")
    print(f"\nOutput: {output_dir}/")
    print(f"  page_1.md, page_2.md ... — upload these to ChatGPT or Gemini to translate")


if __name__ == "__main__":
    main()
