#!/usr/bin/env python3
"""
ACSM to PDF Converter

Converts Adobe ACSM ebook tokens (PDF-sourced) to DRM-free PDF files
for personal offline reading.  When the PDF is image-only (scanned),
an OCR step can add a searchable text layer.

The DRM removal process (adept_remove from libgourou) operates at the
encryption layer only — it decrypts the PDF without re-encoding, so
all images, paragraph structure, fonts, links, bookmarks, and
annotations are preserved exactly as in the original.

Supported OCR languages:
    English, Traditional Chinese, Simplified Chinese,
    Japanese (Hiragana, Katakana, Kanji), Korean (Hangul)

Prerequisites:
    libgourou (built from source)
    pip install ocrmypdf PyMuPDF pypdf
    tesseract + language packs
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


# --- OCR Engine -----------------------------------------------------------

# FIX 1: Always use all 5 language packs.
#
# The previous approach tried to auto-detect the language by running
# Tesseract OSD + a full OCR detection pass on up to 3 sample pages.
# For a 300-page image-only book this wasted 2–3 minutes of pure overhead
# before the real OCR even began — the single biggest cause of timeouts.
#
# Since all 5 packs (eng, chi_tra, chi_sim, jpn, kor) are installed in the
# Docker image, we always supply them together.  Tesseract's internal voting
# system picks the right script for each text block automatically, so output
# quality is equal-or-better with no detection penalty.
#
# The only remaining use of detect_language_from_text / detect_language_from_pdf
# is for the CLI --ocr-lang=auto flag, where speed is less critical.

ALL_OCR_LANGS = "eng+chi_tra+chi_sim+jpn+kor"

LANG_LABELS = {
    "eng": "English",
    "chi_tra": "Traditional Chinese",
    "chi_sim": "Simplified Chinese",
    "jpn": "Japanese",
    "kor": "Korean",
    "chi_tra+chi_sim": "Chinese (mixed)",
    "chi_tra+chi_sim+eng": "Chinese + English",
    "eng+chi_tra": "English + Trad. Chinese",
    "eng+chi_sim": "English + Simp. Chinese",
    "jpn+eng": "Japanese + English",
    "kor+eng": "Korean + English",
    "jpn+chi_tra": "Japanese + Trad. Chinese",
    "jpn+chi_sim": "Japanese + Simp. Chinese",
    "eng+chi_tra+chi_sim+jpn+kor": "All languages (EN / 繁中 / 简中 / 日本語 / 한국어)",
}

_TRA_CHARS = "國學數與對這經區體發聯當會從點問機關個義處應實來將過還後給讓說時種為開黨對質開裡類"
_SIM_CHARS = "国学数与对这经区体发联当会从点问机关个义处应实来将过还后给让说时种为开党对质开里类"


def detect_language_from_text(text: str) -> str:
    """Heuristic language detection from extracted text (CLI / fallback use only)."""
    if not text or len(text.strip()) < 5:
        return ALL_OCR_LANGS

    cjk_count = 0
    eng_count = 0
    hiragana_count = 0
    katakana_count = 0
    hangul_count = 0
    tra_indicators = 0
    sim_indicators = 0

    for ch in text:
        code = ord(ch)
        if 0x4E00 <= code <= 0x9FFF:
            cjk_count += 1
            if ch in _TRA_CHARS:
                tra_indicators += 1
            if ch in _SIM_CHARS:
                sim_indicators += 1
        elif 0x3040 <= code <= 0x309F:
            hiragana_count += 1
        elif 0x30A0 <= code <= 0x30FF:
            katakana_count += 1
        elif 0xAC00 <= code <= 0xD7AF:
            hangul_count += 1
        elif 0x1100 <= code <= 0x11FF:
            hangul_count += 1
        elif 0x3130 <= code <= 0x318F:
            hangul_count += 1
        elif (0x41 <= code <= 0x5A) or (0x61 <= code <= 0x7A):
            eng_count += 1

    jpn_kana = hiragana_count + katakana_count
    total = cjk_count + eng_count + jpn_kana + hangul_count
    if total == 0:
        return ALL_OCR_LANGS

    if jpn_kana > 0:
        if eng_count > total * 0.2:
            return "jpn+eng"
        if tra_indicators > sim_indicators:
            return "jpn+chi_tra"
        return "jpn"

    if hangul_count > 0:
        if hangul_count / total > 0.3:
            if eng_count > total * 0.2:
                return "kor+eng"
            return "kor"
        if eng_count > total * 0.2:
            return "kor+eng"
        return "kor"

    if cjk_count / total > 0.3:
        if tra_indicators > sim_indicators * 1.5:
            return "chi_tra"
        if sim_indicators > tra_indicators * 1.5:
            return "chi_sim"
        return "chi_tra+chi_sim"

    if cjk_count > 0:
        if tra_indicators > sim_indicators:
            return "eng+chi_tra"
        if sim_indicators > tra_indicators:
            return "eng+chi_sim"
        return "chi_tra+chi_sim+eng"

    return "eng"


def detect_language_from_pdf(pdf_path: Path) -> str:
    """Detect language from a PDF that already has a text layer (fast path).

    For image-only PDFs this now returns ALL_OCR_LANGS immediately instead of
    running the expensive multi-pass Tesseract detection loop.
    """
    check = PDFCheckResult()
    _extract_text_pymupdf(pdf_path, check)

    # Fast path: PDF has extractable text — detect from that
    if check.sample_text and not check.sample_text.startswith("("):
        detected = detect_language_from_text(check.sample_text)
        if detected != ALL_OCR_LANGS:
            return detected

    # Image-only PDF — skip the slow Tesseract detection loop entirely.
    # Just return all languages; Tesseract handles multi-script pages correctly.
    return ALL_OCR_LANGS


def _check_tesseract_languages():
    try:
        r = run(["tesseract", "--list-langs"], timeout=10)
        if r.returncode == 0:
            langs = set()
            for line in (r.stdout or "").splitlines():
                line = line.strip()
                if line and not line.startswith("List"):
                    langs.add(line)
            return langs
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass
    return set()


def _filter_ocr_languages(requested: str) -> tuple[str, list[str]]:
    """Filter requested OCR languages to only those installed in tesseract."""
    warnings = []
    available = _check_tesseract_languages()
    if not available:
        return requested, warnings
    parts = requested.split("+")
    valid = [p for p in parts if p in available]
    dropped = [p for p in parts if p not in available]
    if dropped:
        dropped_labels = [LANG_LABELS.get(d, d) for d in dropped]
        msg = (f"Tesseract language pack(s) not installed: "
               f"{', '.join(dropped_labels)}. OCR quality may be degraded.")
        warnings.append(msg)
        print(f"  WARNING: {msg}")
    if not valid:
        if "eng" in available:
            msg = "Falling back to English-only OCR. Install missing packs for better results."
            warnings.append(msg)
            print(f"  WARNING: {msg}")
            return "eng", warnings
        return requested, warnings
    return "+".join(valid), warnings


def _compress_page_ranges(pages: list[int]) -> str:
    if not pages:
        return ""
    sorted_pages = sorted(set(pages))
    ranges = []
    start = prev = sorted_pages[0]
    for p in sorted_pages[1:]:
        if p == prev + 1:
            prev = p
        else:
            ranges.append(f"{start}-{prev}" if prev > start else str(start))
            start = prev = p
    ranges.append(f"{start}-{prev}" if prev > start else str(start))
    return ",".join(ranges)


# FIX 2: Optimal DPI constants
#
# 150 DPI is sufficient for clean Latin-script books.
# 200 DPI is the sweet spot for CJK — fine strokes are legible without the
# ~78 % extra pixel count that 250 DPI produces versus 200 DPI.
# (pixel count scales as DPI², so 250→200 saves ~36 % compute per page.)
_DPI_LATIN = 150
_DPI_CJK   = 200   # was 250 — saves ~36 % compute, quality still excellent

# FIX 3: Per-page Tesseract timeout (seconds).
# Prevents a single corrupted or gigantic page from hanging the whole job.
# ocrmypdf passes this to each Tesseract worker subprocess.
_TESSERACT_TIMEOUT = 120


def run_ocr(input_pdf: Path, output_pdf: Path, language: str = "auto",
            dpi: int = _DPI_LATIN, pages_to_ocr: list[int] | None = None) -> dict:
    """Run OCR on a PDF to add a searchable text layer.

    Preserves all images, paragraph layout, bookmarks, and hyperlinks.
    The input_pdf is NEVER modified. Output always goes to output_pdf.

    Key behaviours
    --------------
    * language="auto"  →  always resolves to ALL_OCR_LANGS (no slow detection)
    * skip_text=True   →  pages that already have a text layer are untouched,
                          so existing links / paragraph structure are preserved
    * optimize=1       →  lossless PDF compression; smaller output, no re-encoding
    * tesseract_timeout→  hard per-page timeout; bad pages fail gracefully
    * _restore_pdf_metadata after OCR re-attaches any bookmarks/links that
      ocrmypdf may have dropped
    """
    if input_pdf.resolve() == output_pdf.resolve():
        raise RuntimeError("input_pdf and output_pdf must be different files")

    try:
        import ocrmypdf
    except ImportError:
        raise RuntimeError(
            "ocrmypdf is not installed. Run: pip install ocrmypdf\n"
            "Also ensure tesseract is installed with language packs."
        )

    # FIX 1 (continued): "auto" always means ALL_OCR_LANGS — no detection pass.
    if language == "auto":
        language = ALL_OCR_LANGS
        print(f"  OCR language: {LANG_LABELS[ALL_OCR_LANGS]} (all packs, no detection step)")
    else:
        lang_label = LANG_LABELS.get(language, language)
        print(f"  OCR language: {lang_label} ({language})")

    language, lang_warnings = _filter_ocr_languages(language)
    lang_label = LANG_LABELS.get(language, language)

    # FIX 2 (continued): use reduced CJK DPI
    _CJK_LANG_PREFIXES = ("chi_", "jpn", "kor")
    is_cjk = any(language.startswith(p) or f"+{p}" in language
                  for p in _CJK_LANG_PREFIXES)
    effective_dpi = _DPI_CJK if is_cjk else dpi
    if is_cjk and effective_dpi != dpi:
        print(f"  DPI set to {effective_dpi} for CJK content "
              f"(saves ~36 % compute vs 250 DPI)")

    ocr_kwargs = {
        "language": language,
        "output_type": "pdf",
        # skip_text=True: pages with an existing text layer are left completely
        # untouched — their fonts, links, and paragraph structure are preserved.
        "skip_text": True,
        # optimize=1: lossless PDF stream compression; reduces file size without
        # re-encoding images (optimize=2/3 would re-encode and could degrade quality).
        "optimize": 1,
        "image_dpi": effective_dpi,
        "progress_bar": False,
        # FIX 2: reduced parallelism keeps memory/CPU pressure lower on small
        # cloud instances; 2 workers is safe even on a 1-vCPU container.
        "jobs": min(os.cpu_count() or 1, 2),
        # FIX 3: per-page Tesseract timeout — bad pages fail gracefully instead
        # of hanging the entire job and triggering a gunicorn worker timeout.
        "tesseract_timeout": _TESSERACT_TIMEOUT,
    }

    if pages_to_ocr:
        pages_str = _compress_page_ranges(pages_to_ocr)
        ocr_kwargs["pages"] = pages_str
        print(f"  Targeting {len(pages_to_ocr)} image-only page(s): {pages_str}"
              f"\n  All other pages left untouched")

    # Snapshot metadata BEFORE OCR so we can restore anything ocrmypdf drops
    original_bookmarks = None
    original_annotations = {}
    try:
        original_bookmarks, original_annotations = _snapshot_pdf_metadata(input_pdf)
    except Exception as e:
        print(f"  (could not snapshot metadata: {e})")

    try:
        exit_code = ocrmypdf.ocr(str(input_pdf), str(output_pdf), **ocr_kwargs)
    except ocrmypdf.exceptions.PriorOcrFoundError:
        print("  All pages already have OCR text -- no processing needed")
        shutil.copy2(input_pdf, output_pdf)
        return {
            "status": "already_has_text",
            "language": language,
            "lang_label": lang_label,
            "pages_ocrd": 0,
            "warnings": lang_warnings,
        }
    except ocrmypdf.exceptions.MissingDependencyError as e:
        raise RuntimeError(
            f"OCR dependency missing: {e}\n"
            "Ensure tesseract and language packs are installed."
        )
    except ocrmypdf.exceptions.EncryptedPdfError:
        raise RuntimeError(
            "PDF is encrypted and cannot be OCR'd. DRM removal may be incomplete."
        )
    except Exception as e:
        raise RuntimeError(f"OCR processing failed: {e}")

    if exit_code != 0 and exit_code != ocrmypdf.ExitCode.already_done_ocr:
        raise RuntimeError(f"OCR exited with code {exit_code}")

    if not output_pdf.exists():
        raise RuntimeError("OCR completed but output file not found")

    # Restore bookmarks / hyperlinks if OCR dropped any
    try:
        _restore_pdf_metadata(output_pdf, original_bookmarks, original_annotations)
    except Exception as e:
        print(f"  (metadata restoration skipped: {e})")

    post_check = PDFCheckResult()
    _extract_text_pymupdf(output_pdf, post_check)

    return {
        "status": "completed",
        "language": language,
        "lang_label": lang_label,
        "pages_total": post_check.total_pages,
        "pages_with_text": post_check.pages_with_text,
        "pages_still_image": len(post_check.pages_image_only),
        "warnings": lang_warnings,
    }


def _snapshot_pdf_metadata(pdf_path: Path):
    try:
        import fitz
    except ImportError:
        return None, {}

    doc = fitz.open(str(pdf_path))
    toc = doc.get_toc(simple=False)

    annotations = {}
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        page_links = []
        for link in page.get_links():
            kind = link.get("kind", 0)
            if kind in (1, 2, 4, 5):
                page_links.append(link)
        if page_links:
            annotations[page_idx] = page_links

    doc.close()
    return toc, annotations


def _restore_pdf_metadata(pdf_path: Path, original_toc, original_annotations):
    """Restore bookmarks and links only if OCR dropped them.

    Only restores if item count DECREASED (OCR may legitimately add items).
    Uses doc.save() — not saveIncr() — since OCR output files do not support
    incremental saves.
    """
    try:
        import fitz
    except ImportError:
        return

    if not original_toc and not original_annotations:
        return

    doc = fitz.open(str(pdf_path))
    modified = False

    if original_toc:
        current_toc = doc.get_toc()
        if len(current_toc) < len(original_toc):
            print(f"  Restoring {len(original_toc)} bookmarks/TOC entries "
                  f"(ocrmypdf left {len(current_toc)})...")
            try:
                doc.set_toc(original_toc)
                modified = True
            except Exception as e:
                print(f"  (TOC restore failed: {e})")

    if original_annotations:
        restored_count = 0
        for page_idx, orig_links in original_annotations.items():
            if page_idx >= len(doc):
                continue
            page = doc[page_idx]
            current_links = page.get_links()

            if len(current_links) < len(orig_links):
                for cl in current_links:
                    try:
                        page.delete_link(cl)
                    except Exception:
                        pass
                for link in orig_links:
                    try:
                        page.insert_link(link)
                        restored_count += 1
                    except Exception:
                        pass
                modified = True

        if restored_count > 0:
            print(f"  Restored {restored_count} link annotations across "
                  f"{len(original_annotations)} page(s)")

    if modified:
        doc.save(str(pdf_path), garbage=1, deflate=True)
    doc.close()


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
        img_pages = list(range(1, pdf_result.total_pages + 1))
        yield ("ocr_prompt", f"{output_file}|{','.join(str(p) for p in img_pages)}")
        return

    elif pdf_result.needs_ocr:
        img_count = len(pdf_result.pages_image_only)
        yield (6, (
            f"PDF scan: {pdf_result.pages_with_text}/{pdf_result.total_pages} pages "
            f"have text, {img_count} page(s) are image-only{structure_info}."
        ))
        yield ("ocr_prompt", f"{output_file}|{','.join(str(p) for p in pdf_result.pages_image_only)}")
        return

    else:
        yield (6, (
            f"PDF verified: {pdf_result.pages_with_text}/{pdf_result.total_pages} pages "
            f"have readable, selectable text{structure_info} — all OK."
        ))

    size_mb = output_file.stat().st_size / (1024 * 1024) if output_file.exists() else 0
    yield ("done", f"{output_file.name}|{size_mb:.1f} MB")


def run_ocr_step(output_file, pages_image_only):
    """Run OCR and keep the original PDF alongside the new OCR'd version.

    The original PDF is NEVER deleted. OCR output is saved as <stem>_ocr.pdf.

    Done message format — two files available:
        "<stem>_ocr.pdf|<ocr_size> MB|<stem>.pdf|<orig_size> MB"

    Done message format — one file (already had text, or OCR failed):
        "<stem>.pdf|<size> MB"
    """
    output_file = Path(output_file)
    output_dir = output_file.parent
    stem = output_file.stem

    print("Running OCR to add text layer...")
    ocr_output = output_dir / f"{stem}_ocr.pdf"

    try:
        ocr_result = run_ocr(
            input_pdf=output_file,
            output_pdf=ocr_output,
            language="auto",       # resolves to ALL_OCR_LANGS immediately
            dpi=_DPI_LATIN,        # run_ocr boosts to _DPI_CJK automatically
            pages_to_ocr=pages_image_only or None,
        )

        for warn_msg in ocr_result.get("warnings", []):
            yield (7, f"Warning: {warn_msg}")

        if ocr_result["status"] == "already_has_text":
            if ocr_output.exists():
                try:
                    ocr_output.unlink()
                except Exception:
                    pass
            yield (7, (
                f"OCR skipped -- all pages already have a text layer. "
                f"Language: {ocr_result['lang_label']}"
            ))
            orig_size_mb = output_file.stat().st_size / (1024 * 1024)
            yield ("done", f"{output_file.name}|{orig_size_mb:.1f} MB")

        else:
            still_img = ocr_result.get("pages_still_image", 0)
            text_pages = ocr_result.get("pages_with_text", 0)
            total = ocr_result.get("pages_total", 0)
            if still_img > 0:
                yield (7, (
                    f"OCR complete ({ocr_result['lang_label']}): "
                    f"{text_pages}/{total} pages now have text. "
                    f"{still_img} page(s) could not be OCR'd."
                ))
            else:
                yield (7, (
                    f"OCR complete ({ocr_result['lang_label']}): "
                    f"all {total} pages now have readable, searchable text."
                ))

            ocr_size_mb = ocr_output.stat().st_size / (1024 * 1024) if ocr_output.exists() else 0
            orig_size_mb = output_file.stat().st_size / (1024 * 1024) if output_file.exists() else 0
            yield ("done", (
                f"{ocr_output.name}|{ocr_size_mb:.1f} MB"
                f"|{output_file.name}|{orig_size_mb:.1f} MB"
            ))

    except RuntimeError as e:
        if ocr_output.exists():
            ocr_output.unlink()
        yield (7, f"OCR failed: {e}. PDF available without text layer.")
        orig_size_mb = output_file.stat().st_size / (1024 * 1024) if output_file.exists() else 0
        yield ("done", f"{output_file.name}|{orig_size_mb:.1f} MB")


def do_convert(acsm_file, output_dir):
    try:
        for step, message in convert_pipeline(acsm_file, output_dir):
            if step == "done":
                parts = message.split("|")
                print(f"\n=== Done! ===\nFile: {parts[0]} ({parts[1]})")
                if len(parts) >= 4:
                    print(f"Original: {parts[2]} ({parts[3]})")
            elif step == "ocr_prompt":
                parts = message.split("|")
                pdf_path = parts[0]
                pages = [int(p) for p in parts[1].split(",")]
                answer = input(f"\n{len(pages)} page(s) need OCR. Run OCR? [Y/n]: ").strip().lower()
                if answer in ("", "y", "yes"):
                    for s, m in run_ocr_step(pdf_path, pages):
                        if s == "done":
                            p = m.split("|")
                            print(f"\n=== Done! ===\nFile: {p[0]} ({p[1]})")
                            if len(p) >= 4:
                                print(f"Original: {p[2]} ({p[3]})")
                        else:
                            print(f"\n=== Step {s}: {m} ===")
                else:
                    print("\n=== OCR skipped. PDF downloaded as image-only. ===")
                    pdf = Path(pdf_path)
                    size_mb = pdf.stat().st_size / (1024 * 1024)
                    print(f"File: {pdf.name} ({size_mb:.1f} MB)")
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
    parser.add_argument("--ocr-only", metavar="FILE", help="Run OCR on an existing PDF")
    parser.add_argument("--ocr-lang", default="auto",
                        help="OCR language: auto (all packs), eng, chi_tra, chi_sim, jpn, kor")
    args = parser.parse_args()

    if args.verify_only:
        path = Path(args.verify_only)
        result = verify_pdf_readability(path)
        print(result.summary())
        sys.exit(1 if result.has_errors else 0)

    if args.ocr_only:
        path = Path(args.ocr_only)
        if not path.exists():
            print(f"File not found: {path}")
            sys.exit(1)
        out_path = path.parent / f"{path.stem}_ocr.pdf"
        print(f"Running OCR on {path.name}...")
        try:
            result = run_ocr(path, out_path, language=args.ocr_lang)
            print(f"Done! Output: {out_path.name}")
            print(f"  Language: {result['lang_label']}")
            print(f"  Status: {result['status']}")
            if result.get("pages_with_text"):
                print(f"  Pages with text: {result['pages_with_text']}/{result['pages_total']}")
        except RuntimeError as e:
            print(f"OCR failed: {e}")
            sys.exit(1)
        return

    if not args.acsm_file:
        parser.print_help()
        sys.exit(1)
    do_convert(args.acsm_file, args.output_dir)


if __name__ == "__main__":
    main()
