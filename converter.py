#!/usr/bin/env python3
"""
ACSM to PDF Converter

Converts Adobe ACSM ebook tokens (PDF-sourced) to DRM-free PDF files
for personal offline reading.

The DRM removal process (adept_remove from libgourou) operates at the
encryption layer only — it decrypts the PDF without re-encoding, so
all images, paragraph structure, fonts, links, bookmarks, and
annotations are preserved exactly as in the original.

Prerequisites:
    libgourou (built from source)
    pip install PyMuPDF pypdf
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
LIBGOUROU_DIR = SCRIPT_DIR / "libgourou"
LIBGOUROU_BIN = LIBGOUROU_DIR / "utils"
ADEPT_DIR = Path.home() / ".config" / "adept"


def run(cmd, **kwargs):
    """Run a command and return the result."""
    defaults = {"capture_output": True, "text": True}
    defaults.update(kwargs)
    return subprocess.run(cmd, **defaults)


def find_tool(name):
    """Find a tool, checking local build directory first."""
    local = LIBGOUROU_BIN / name
    if local.exists() and os.access(local, os.X_OK):
        return str(local)
    system = shutil.which(name)
    if system:
        return system
    return None


# --- Conversion -----------------------------------------------------------


def detect_format(acsm_path):
    tree = ET.parse(acsm_path)
    root = tree.getroot()
    ns = {"adept": "http://ns.adobe.com/adept"}
    src_elem = root.find(".//adept:src", ns)
    if src_elem is not None and src_elem.text:
        src = src_elem.text.lower()
        if ".pdf" in src or "output=pdf" in src:
            return "pdf"
    fmt_elem = root.find(".//adept:metadata/adept:format", ns)
    if fmt_elem is not None and fmt_elem.text:
        if "pdf" in fmt_elem.text.lower():
            return "pdf"
    return "epub"


def register_device():
    device_file = ADEPT_DIR / "device.xml"
    if device_file.exists():
        print("[OK] Adobe device already registered.")
        return
    print("Registering Adobe device (anonymous)...")
    tool = find_tool("adept_activate")
    try:
        result = run([tool, "-a"], timeout=30)
    except subprocess.TimeoutExpired:
        raise RuntimeError("Device registration timed out (30s).")
    if result.returncode != 0:
        raise RuntimeError(f"Device registration failed: {result.stdout}\n{result.stderr}")
    print("[OK] Adobe device registered.")


def fulfill_acsm(acsm_path, output_path):
    print(f"Fulfilling ACSM: {acsm_path.name}")
    tool = find_tool("acsmdownloader")
    try:
        result = run([tool, "-f", str(acsm_path), "-o", str(output_path)], timeout=120)
    except subprocess.TimeoutExpired:
        raise RuntimeError("Download timed out (120s).")
    if result.returncode != 0:
        stderr = result.stderr or result.stdout or ""
        raise RuntimeError(f"ACSM download failed (exit {result.returncode}): {stderr[:500]}")
    if not output_path.exists():
        raise RuntimeError("Download completed but output file not found.")
    size_kb = output_path.stat().st_size / 1024
    print(f"[OK] Downloaded: {output_path.name} ({size_kb:.0f} KB)")


def remove_drm(input_path, output_path):
    """Remove Adobe DRM encryption from the PDF.

    This is a decryption-only operation — it does NOT re-encode, re-render,
    or modify the PDF content.  All images, text, paragraph structure, fonts,
    links, bookmarks, and annotations are preserved byte-for-byte.
    """
    print(f"Removing DRM: {input_path.name}")
    tool = find_tool("adept_remove")
    try:
        result = run([tool, "-f", str(input_path), "-o", str(output_path)], timeout=60)
    except subprocess.TimeoutExpired:
        raise RuntimeError("DRM removal timed out (60s).")
    if result.returncode != 0:
        raise RuntimeError(f"DRM removal failed: {(result.stderr or result.stdout)[:300]}")
    print(f"[OK] DRM removed: {output_path.name}")


# --- PDF Verification -----------------------------------------------------


class PDFCheckResult:
    def __init__(self):
        self.total_pages: int = 0
        self.pages_with_text: int = 0
        self.pages_image_only: list[int] = []
        self.sample_text: str = ""
        self.warnings: list[str] = []
        self.encrypted: bool = False
        self.has_fonts: bool = False
        self.has_bookmarks: bool = False
        self.link_count: int = 0

    @property
    def has_errors(self) -> bool:
        return self.encrypted

    @property
    def needs_ocr(self) -> bool:
        return len(self.pages_image_only) > 0

    @property
    def probably_image_only(self) -> bool:
        return (
            self.total_pages > 0
            and self.pages_with_text == 0
            and not self.has_fonts
        )

    @property
    def text_ratio(self) -> float:
        if self.total_pages == 0:
            return 0.0
        return self.pages_with_text / self.total_pages

    def summary(self) -> str:
        lines = [
            f"Total pages    : {self.total_pages}",
            f"Pages with text: {self.pages_with_text}",
            f"Image-only     : {len(self.pages_image_only)}",
            f"Text ratio     : {self.text_ratio:.0%}",
            f"Bookmarks      : {'Yes' if self.has_bookmarks else 'No'}",
            f"Links          : {self.link_count}",
        ]
        if self.encrypted:
            lines.append("! PDF is still encrypted!")
        if self.pages_image_only:
            pages_str = ", ".join(str(p) for p in self.pages_image_only[:10])
            if len(self.pages_image_only) > 10:
                pages_str += f" ... and {len(self.pages_image_only) - 10} more"
            lines.append(f"Image-only pages: {pages_str}")
        if self.warnings:
            lines.append("Warnings:")
            for w in self.warnings:
                lines.append(f"  {w}")
        return "\n".join(lines)


def _extract_text_pymupdf(pdf_path: Path, result: PDFCheckResult) -> bool:
    try:
        import fitz
    except ImportError:
        return False
    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        result.warnings.append(f"PyMuPDF cannot open PDF: {e}")
        return False
    if doc.is_encrypted:
        result.encrypted = True
        result.warnings.append("PDF is still encrypted after DRM removal.")
        doc.close()
        return True
    result.total_pages = len(doc)
    result.has_fonts = False

    # Check bookmarks / TOC
    toc = doc.get_toc()
    result.has_bookmarks = len(toc) > 0

    for i, page in enumerate(doc):
        try:
            text = page.get_text("text") or ""
            clean = text.strip()
            fonts = page.get_fonts()
            links = page.get_links()
            result.link_count += len(links)
            if fonts:
                result.has_fonts = True
            if len(clean) >= 5:
                result.pages_with_text += 1
                if not result.sample_text and len(clean) > 10:
                    result.sample_text = clean[:200]
            else:
                if fonts:
                    result.pages_with_text += 1
                    if not result.sample_text:
                        result.sample_text = "(text present but not extractable -- fonts embedded)"
                else:
                    result.pages_image_only.append(i + 1)
        except Exception:
            result.pages_image_only.append(i + 1)
    doc.close()
    return True


def _extract_text_pypdf(pdf_path: Path, result: PDFCheckResult) -> bool:
    try:
        from pypdf import PdfReader
    except ImportError:
        return False
    try:
        reader = PdfReader(pdf_path)
    except Exception as e:
        result.warnings.append(f"pypdf cannot open PDF: {e}")
        return False
    if reader.is_encrypted:
        result.encrypted = True
        result.warnings.append("PDF is still encrypted after DRM removal.")
        return True
    result.total_pages = len(reader.pages)
    for i, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
            clean = text.strip()
            if len(clean) >= 5:
                result.pages_with_text += 1
                if not result.sample_text and len(clean) > 10:
                    result.sample_text = clean[:200]
            else:
                result.pages_image_only.append(i + 1)
        except Exception:
            result.pages_image_only.append(i + 1)
    return True


def verify_pdf_readability(pdf_path: Path) -> PDFCheckResult:
    result = PDFCheckResult()
    if not pdf_path.exists():
        result.warnings.append(f"PDF file not found: {pdf_path}")
        return result
    if not _extract_text_pymupdf(pdf_path, result):
        if not _extract_text_pypdf(pdf_path, result):
            result.warnings.append(
                "Neither PyMuPDF nor pypdf is installed -- skipping text verification"
            )
    return result



# --- Pipeline -------------------------------------------------------------


def convert_pipeline(acsm_path, output_dir):
    """Generator yielding (step, message) tuples."""
    acsm_path = Path(acsm_path).resolve()
    if not acsm_path.exists():
        raise RuntimeError(f"File not found: {acsm_path}")
    if acsm_path.suffix != ".acsm":
        raise RuntimeError(f"Not an ACSM file: {acsm_path}")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = acsm_path.stem

    # Step 1: Check tools
    problems = []
    if not find_tool("acsmdownloader"):
        problems.append("acsmdownloader not found (run --setup)")
    if not find_tool("adept_activate"):
        problems.append("adept_activate not found (run --setup)")
    if not find_tool("adept_remove"):
        problems.append("adept_remove not found (run --setup)")
    if problems:
        raise RuntimeError("Missing components: " + "; ".join(set(problems)))
    yield (1, "All tools ready.")

    # Step 2: Detect format
    fmt = detect_format(acsm_path)
    if fmt != "pdf":
        raise RuntimeError(
            f"This ACSM file is for {fmt.upper()} format. "
            f"Only PDF-sourced ACSM files are supported."
        )
    yield (2, "Detected format: PDF")

    # Step 3: Register device
    register_device()
    yield (3, "Device registered.")

    # Step 4: Download
    drm_file = output_dir / f"{stem}_drm.pdf"
    fulfill_acsm(acsm_path, drm_file)
    yield (4, f"Downloaded: {drm_file.name}")

    # Step 5: Remove DRM (decryption only — preserves all PDF structure)
    output_file = output_dir / f"{stem}.pdf"
    remove_drm(drm_file, output_file)
    drm_file.unlink()
    yield (5, f"DRM removed: {output_file.name}")

    # Step 6: Verify readability and structure preservation
    print("Verifying PDF readability...")
    pdf_result = verify_pdf_readability(output_file)

    if pdf_result.encrypted:
        raise RuntimeError(
            "DRM removal incomplete: the PDF is still encrypted."
        )

    structure_parts = []
    if pdf_result.has_bookmarks:
        structure_parts.append("bookmarks intact")
    if pdf_result.link_count > 0:
        structure_parts.append(f"{pdf_result.link_count} links preserved")
    structure_info = (" — " + ", ".join(structure_parts)) if structure_parts else ""

    if pdf_result.probably_image_only:
        yield (6, (
            f"PDF scan: 0/{pdf_result.total_pages} pages have extractable text. "
            f"Image-only PDF detected{structure_info}."
        ))
    elif pdf_result.needs_ocr:
        img_count = len(pdf_result.pages_image_only)
        yield (6, (
            f"PDF scan: {pdf_result.pages_with_text}/{pdf_result.total_pages} pages "
            f"have text, {img_count} page(s) are image-only{structure_info}."
        ))
    else:
        yield (6, (
            f"PDF verified: {pdf_result.pages_with_text}/{pdf_result.total_pages} pages "
            f"have readable, selectable text{structure_info} — all OK."
        ))

    size_mb = output_file.stat().st_size / (1024 * 1024) if output_file.exists() else 0
    yield ("done", f"{output_file.name}|{size_mb:.1f} MB")


def do_convert(acsm_file, output_dir):
    try:
        for step, message in convert_pipeline(acsm_file, output_dir):
            if step == "done":
                parts = message.split("|")
                print(f"\n=== Done! ===\nFile: {parts[0]} ({parts[1]})")
            else:
                print(f"\n=== Step {step}: {message} ===")
    except RuntimeError as e:
        print(str(e))
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF-sourced ACSM ebook tokens to DRM-free PDF.",
    )
    parser.add_argument("acsm_file", nargs="?", help="Path to the .acsm file")
    parser.add_argument("-o", "--output-dir", default="output", help="Output directory")
    parser.add_argument("--verify-only", metavar="FILE", help="Audit an existing PDF")
    args = parser.parse_args()

    if args.verify_only:
        path = Path(args.verify_only)
        result = verify_pdf_readability(path)
        print(result.summary())
        sys.exit(1 if result.has_errors else 0)

    if not args.acsm_file:
        parser.print_help()
        sys.exit(1)
    do_convert(args.acsm_file, args.output_dir)


if __name__ == "__main__":
    main()
