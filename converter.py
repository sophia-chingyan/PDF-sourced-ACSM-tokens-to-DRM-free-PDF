#!/usr/bin/env python3
"""
ACSM to EPUB/PDF Converter

Converts Adobe ACSM ebook tokens to DRM-free EPUB or PDF files
for personal offline reading.  When the PDF is image-only (scanned),
an OCR step automatically adds a searchable text layer.

Supported OCR languages:
    English, Traditional Chinese, Simplified Chinese,
    Japanese (Hiragana, Katakana, Kanji), Korean (Hangul)

Prerequisites (installed automatically by setup):
    brew install pugixml libzip openssl curl cmake
    brew install tesseract tesseract-lang   # for OCR
    pip install ocrmypdf PyMuPDF pypdf
    libgourou (built from source)

Usage:
    python3 converter.py --setup          # First-time setup
    python3 converter.py ebook.acsm       # Convert an ACSM file
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path, PurePosixPath
from urllib.parse import unquote, urlparse

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


# --- Setup ----------------------------------------------------------------


def setup_brew_deps():
    if not shutil.which("brew"):
        print("Homebrew is required. Install from https://brew.sh")
        sys.exit(1)
    deps = ["pugixml", "libzip", "openssl", "curl", "cmake"]
    print(f"Installing build dependencies: {', '.join(deps)}")
    result = run(["brew", "install"] + deps)
    if result.returncode != 0:
        print(f"brew install failed:\n{result.stderr}")
        sys.exit(1)
    print("[OK] Build dependencies installed.")


def _get_brew_prefixes():
    prefixes = {}
    for dep in ["pugixml", "libzip", "openssl", "curl"]:
        r = run(["brew", "--prefix", dep])
        prefixes[dep] = r.stdout.strip() if r.returncode == 0 else f"/opt/homebrew/opt/{dep}"
    return prefixes


def _patch_makefiles(brew_prefixes):
    include_flags = " ".join(f"-I{p}/include" for p in brew_prefixes.values())
    lib_flags = " ".join(f"-L{p}/lib" for p in brew_prefixes.values())
    root_mk = LIBGOUROU_DIR / "Makefile"
    content = root_mk.read_text()
    content = content.replace("$(AR) rcs --thin $@ $^", "libtool -static -o $@ $^")
    root_mk.write_text(content)
    utils_mk = LIBGOUROU_DIR / "utils" / "Makefile"
    content = utils_mk.read_text()
    content = content.replace(
        "CXXFLAGS=-Wall -fPIC -I$(ROOT)/include",
        f"CXXFLAGS=-Wall -fPIC -I$(ROOT)/include {include_flags}",
    )
    content = content.replace(
        "LDFLAGS += -L$(ROOT) -lcrypto",
        f"LDFLAGS += -L$(ROOT) {lib_flags} -lcrypto",
    )
    utils_mk.write_text(content)


def build_libgourou():
    if (LIBGOUROU_BIN / "acsmdownloader").exists():
        print("[OK] libgourou already built.")
        return
    repo_url = "https://forge.soutade.fr/soutade/libgourou.git"
    if not LIBGOUROU_DIR.exists():
        print("Cloning libgourou...")
        result = run(["git", "clone", "--recurse-submodules", repo_url, str(LIBGOUROU_DIR)])
        if result.returncode != 0:
            print(f"Clone failed:\n{result.stderr}")
            sys.exit(1)
    brew_prefixes = _get_brew_prefixes()
    include_flags = " ".join(f"-I{p}/include" for p in brew_prefixes.values())
    print("Patching Makefiles for macOS...")
    _patch_makefiles(brew_prefixes)
    print("Building libgourou...")
    env = os.environ.copy()
    env["CXXFLAGS"] = include_flags
    result = run(
        ["make", "BUILD_UTILS=1", "BUILD_STATIC=1", "BUILD_SHARED=0"],
        cwd=str(LIBGOUROU_DIR), env=env,
    )
    if result.returncode != 0:
        print(f"Build failed:\n{result.stdout}\n{result.stderr}")
        sys.exit(1)
    if not (LIBGOUROU_BIN / "acsmdownloader").exists():
        print("Build completed but binaries not found.")
        sys.exit(1)
    print("[OK] libgourou built successfully.")


def do_setup():
    print("=== Setting up ACSM Converter ===\n")
    setup_brew_deps()
    print()
    build_libgourou()
    print("\n=== Setup complete! ===")
    print("You can now convert ACSM files:")
    print("  python3 converter.py ebook.acsm")


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
    for i, page in enumerate(doc):
        try:
            text = page.get_text("text") or ""
            clean = text.strip()
            fonts = page.get_fonts()
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

# Character sets for CJK variant detection
_TRA_CHARS = "國學數與對這經區體發聯當會從點問機關個義處應實來將過還後給讓說時種為開黨對質開裡類"
_SIM_CHARS = "国学数与对这经区体发联当会从点问机关个义处应实来将过还后给让说时种为开党对质开里类"

# All supported languages (Tesseract lang codes)
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
    "eng+chi_tra+chi_sim+jpn+kor": "All (EN/ZH/JA/KO)",
}


def detect_language_from_text(text: str) -> str:
    """Detect script/language from sample text.

    Detects: English, Traditional Chinese, Simplified Chinese,
    Japanese (Hiragana, Katakana, Kanji), Korean (Hangul).

    Unicode ranges used:
      - CJK Unified Ideographs  U+4E00..U+9FFF  (shared by ZH/JA/KO)
      - Hiragana                U+3040..U+309F   (Japanese)
      - Katakana                U+30A0..U+30FF   (Japanese)
      - Hangul Syllables        U+AC00..U+D7AF   (Korean)
      - Hangul Jamo             U+1100..U+11FF   (Korean)
      - Hangul Compat. Jamo     U+3130..U+318F   (Korean)
      - Latin A-Z/a-z           U+0041..U+005A / U+0061..U+007A
    """
    if not text or len(text.strip()) < 5:
        return ALL_OCR_LANGS  # too little text -> use all

    cjk_count = 0       # Han ideographs (shared ZH/JA/KO)
    eng_count = 0        # Latin letters
    hiragana_count = 0   # Japanese hiragana
    katakana_count = 0   # Japanese katakana
    hangul_count = 0     # Korean hangul
    tra_indicators = 0   # Traditional Chinese character hints
    sim_indicators = 0   # Simplified Chinese character hints

    for ch in text:
        code = ord(ch)
        # CJK Unified Ideographs (shared by Chinese, Japanese kanji, Korean hanja)
        if 0x4E00 <= code <= 0x9FFF:
            cjk_count += 1
            if ch in _TRA_CHARS:
                tra_indicators += 1
            if ch in _SIM_CHARS:
                sim_indicators += 1
        # Hiragana
        elif 0x3040 <= code <= 0x309F:
            hiragana_count += 1
        # Katakana
        elif 0x30A0 <= code <= 0x30FF:
            katakana_count += 1
        # Hangul Syllables
        elif 0xAC00 <= code <= 0xD7AF:
            hangul_count += 1
        # Hangul Jamo
        elif 0x1100 <= code <= 0x11FF:
            hangul_count += 1
        # Hangul Compatibility Jamo
        elif 0x3130 <= code <= 0x318F:
            hangul_count += 1
        # Latin A-Z / a-z
        elif (0x41 <= code <= 0x5A) or (0x61 <= code <= 0x7A):
            eng_count += 1

    jpn_kana = hiragana_count + katakana_count
    total = cjk_count + eng_count + jpn_kana + hangul_count
    if total == 0:
        return ALL_OCR_LANGS

    # --- Japanese: presence of kana is a strong signal ---
    if jpn_kana > 0:
        # Kana exists -> definitely Japanese
        if eng_count > total * 0.2:
            return "jpn+eng"
        if tra_indicators > sim_indicators:
            return "jpn+chi_tra"  # Japanese + Trad. Chinese (common in some texts)
        return "jpn"

    # --- Korean: presence of Hangul is a strong signal ---
    if hangul_count > 0:
        if hangul_count / total > 0.3:
            if eng_count > total * 0.2:
                return "kor+eng"
            return "kor"
        # Mixed Korean + CJK
        if eng_count > total * 0.2:
            return "kor+eng"
        return "kor"

    # --- Chinese or English (no kana, no hangul) ---
    if cjk_count / total > 0.3:
        # Primarily CJK -- pick variant
        if tra_indicators > sim_indicators * 1.5:
            return "chi_tra"
        if sim_indicators > tra_indicators * 1.5:
            return "chi_sim"
        return "chi_tra+chi_sim"

    if cjk_count > 0:
        # Mixed CJK + English
        if tra_indicators > sim_indicators:
            return "eng+chi_tra"
        if sim_indicators > tra_indicators:
            return "eng+chi_sim"
        return "chi_tra+chi_sim+eng"

    return "eng"


def detect_language_from_pdf(pdf_path: Path) -> str:
    """Quick-scan PDF to detect language for OCR.

    Supports: English, Traditional Chinese, Simplified Chinese,
    Japanese (Hiragana, Katakana, Kanji), Korean (Hangul).
    """
    # Try existing extracted text first
    check = PDFCheckResult()
    _extract_text_pymupdf(pdf_path, check)
    if check.sample_text and not check.sample_text.startswith("("):
        detected = detect_language_from_text(check.sample_text)
        if detected != ALL_OCR_LANGS:
            return detected

    # Render first page and run quick Tesseract detection
    try:
        import fitz
        import tempfile
        doc = fitz.open(str(pdf_path))
        if len(doc) == 0:
            doc.close()
            return ALL_OCR_LANGS
        page = doc[0]
        mat = fitz.Matrix(150 / 72, 150 / 72)
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        doc.close()

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp.write(img_bytes)
            tmp_path = tmp.name
        try:
            # Use all installed languages for the detection pass
            detect_langs = _filter_ocr_languages(ALL_OCR_LANGS)
            r = run(["tesseract", tmp_path, "stdout",
                      "-l", detect_langs, "--psm", "3"], timeout=30)
            if r.returncode == 0 and r.stdout:
                return detect_language_from_text(r.stdout)
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
    except ImportError:
        pass
    except Exception:
        pass

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


def _filter_ocr_languages(requested: str) -> str:
    available = _check_tesseract_languages()
    if not available:
        return requested
    parts = requested.split("+")
    valid = [p for p in parts if p in available]
    if not valid:
        if "eng" in available:
            return "eng"
        return requested
    return "+".join(valid)


def _compress_page_ranges(pages: list[int]) -> str:
    """Convert [1,2,3,5,7,8,9] to '1-3,5,7-9' for ocrmypdf --pages."""
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


def run_ocr(input_pdf: Path, output_pdf: Path, language: str = "auto",
            dpi: int = 300, pages_to_ocr: list[int] | None = None) -> dict:
    """Run OCR on a PDF to add a searchable text layer.

    CRITICAL DESIGN: preserve original PDF structure, images, and links.

    Uses ocrmypdf with carefully chosen options:
    - output_type='pdf'   : avoids Ghostscript PDF/A conversion which strips
                            hyperlinks, bookmarks, TOC, and annotations
    - skip_text=True      : never re-OCR pages that already have text
    - pages=<list>        : only process image-only pages, leaving all other
                            pages completely untouched (bit-for-bit)
    - NO clean            : unpaper destroys real images/diagrams in ebooks
    - NO deskew           : ebook pages are digitally straight; deskewing
                            rotates pages and misaligns text overlay
    - optimize=0          : no re-encoding of images for maximum fidelity

    After OCR, restores any bookmarks/links/annotations that may have been
    lost from the processed pages.
    """
    try:
        import ocrmypdf
    except ImportError:
        raise RuntimeError(
            "ocrmypdf is not installed. Run: pip install ocrmypdf\n"
            "Also ensure tesseract is installed with language packs."
        )

    if language == "auto":
        print("  Auto-detecting language...")
        language = detect_language_from_pdf(input_pdf)

    language = _filter_ocr_languages(language)
    lang_label = LANG_LABELS.get(language, language)
    print(f"  OCR language: {lang_label} ({language})")

    ocr_kwargs = {
        "input_file": str(input_pdf),
        "output_file": str(output_pdf),
        "language": language,
        "output_type": "pdf",       # NOT pdfa — preserves links & annotations
        "skip_text": True,           # never re-OCR existing text pages
        "optimize": 0,               # no image re-encoding for max fidelity
        "image_dpi": dpi,            # DPI for images without resolution info
        "progress_bar": False,
        "jobs": min(os.cpu_count() or 1, 4),
        # Deliberately OMITTED:
        # - deskew: rotates pages, misaligns text overlay in digital ebooks
        # - clean:  runs unpaper which destroys real images/diagrams
    }

    # Target only image-only pages if we know which ones need OCR.
    # Pages NOT in this list are completely untouched — their links,
    # images, and structure are preserved bit-for-bit.
    # ocrmypdf accepts pages as a comma-separated string like "1,3,5-7"
    if pages_to_ocr:
        pages_str = _compress_page_ranges(pages_to_ocr)
        ocr_kwargs["pages"] = pages_str
        print(f"  Targeting {len(pages_to_ocr)} image-only page(s): {pages_str}"
              f"\n  All other pages left untouched")

    # Snapshot bookmarks and link annotations before OCR
    original_bookmarks = None
    original_annotations = {}
    try:
        original_bookmarks, original_annotations = _snapshot_pdf_metadata(input_pdf)
    except Exception as e:
        print(f"  (could not snapshot metadata: {e})")

    try:
        exit_code = ocrmypdf.ocr(**ocr_kwargs)
    except ocrmypdf.exceptions.PriorOcrFoundError:
        print("  All pages already have OCR text -- no processing needed")
        if input_pdf != output_pdf:
            shutil.copy2(input_pdf, output_pdf)
        return {
            "status": "already_has_text",
            "language": language,
            "lang_label": lang_label,
            "pages_ocrd": 0,
        }
    except ocrmypdf.exceptions.MissingDependencyError as e:
        raise RuntimeError(
            f"OCR dependency missing: {e}\n"
            "Ensure tesseract and language packs are installed:\n"
            "  apt-get install tesseract-ocr tesseract-ocr-eng "
            "tesseract-ocr-chi-tra tesseract-ocr-chi-sim "
            "tesseract-ocr-jpn tesseract-ocr-kor"
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

    # Restore any bookmarks/links/annotations lost during OCR
    try:
        _restore_pdf_metadata(output_pdf, original_bookmarks, original_annotations)
    except Exception as e:
        print(f"  (metadata restoration skipped: {e})")

    # Verify post-OCR
    post_check = PDFCheckResult()
    _extract_text_pymupdf(output_pdf, post_check)

    return {
        "status": "completed",
        "language": language,
        "lang_label": lang_label,
        "pages_total": post_check.total_pages,
        "pages_with_text": post_check.pages_with_text,
        "pages_still_image": len(post_check.pages_image_only),
    }


def _snapshot_pdf_metadata(pdf_path: Path):
    """Snapshot bookmarks (TOC) and per-page link annotations from a PDF.

    Returns (toc, annotations_dict) where:
    - toc is the list returned by doc.get_toc()
    - annotations_dict maps page_index -> list of link annotation dicts
    """
    try:
        import fitz
    except ImportError:
        return None, {}

    doc = fitz.open(str(pdf_path))
    toc = doc.get_toc(simple=False)  # [level, title, page, dest_dict]

    annotations = {}
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        page_links = []
        for link in page.get_links():
            # Preserve URI links (kind=2), internal goto (kind=1),
            # named destinations (kind=4), and goto-remote (kind=5)
            kind = link.get("kind", 0)
            if kind in (1, 2, 4, 5):
                page_links.append(link)
        if page_links:
            annotations[page_idx] = page_links

    doc.close()
    return toc, annotations


def _restore_pdf_metadata(pdf_path: Path, original_toc, original_annotations):
    """Restore bookmarks and link annotations to an OCR'd PDF.

    Compares the OCR'd output against the original snapshot and adds back
    any bookmarks or links that were lost during processing.
    """
    try:
        import fitz
    except ImportError:
        return

    if not original_toc and not original_annotations:
        return  # nothing to restore

    doc = fitz.open(str(pdf_path))
    modified = False

    # Restore TOC / bookmarks if they were lost
    if original_toc:
        current_toc = doc.get_toc()
        if len(current_toc) < len(original_toc):
            print(f"  Restoring {len(original_toc)} bookmarks/TOC entries...")
            try:
                doc.set_toc(original_toc)
                modified = True
            except Exception as e:
                print(f"  (TOC restore failed: {e})")

    # Restore link annotations per page
    if original_annotations:
        restored_count = 0
        for page_idx, orig_links in original_annotations.items():
            if page_idx >= len(doc):
                continue
            page = doc[page_idx]
            current_links = page.get_links()

            # If OCR'd page has fewer links, some were lost — clear and
            # re-insert all originals to avoid duplicates
            if len(current_links) < len(orig_links):
                # Remove existing links first
                for cl in current_links:
                    try:
                        page.delete_link(cl)
                    except Exception:
                        pass
                # Re-insert all original links
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
        try:
            doc.saveIncr()  # incremental save preserves existing structure
        except Exception:
            # Fallback: full save with minimal processing
            doc.save(str(pdf_path), garbage=0, deflate=False)
    doc.close()


# --- Link Verification (EPUB) --------------------------------------------

_LINK_ATTRS = {
    "a": ["href"], "area": ["href"], "link": ["href"],
    "script": ["src"], "img": ["src", "srcset"],
    "image": ["href", "{http://www.w3.org/1999/xlink}href"],
    "use": ["href", "{http://www.w3.org/1999/xlink}href"],
    "video": ["src", "poster"], "audio": ["src"],
    "source": ["src", "srcset"], "track": ["src"],
    "iframe": ["src"], "object": ["data"], "embed": ["src"],
    "blockquote": ["cite"], "q": ["cite"],
    "ins": ["cite"], "del": ["cite"],
}
_CSS_URL_RE = re.compile(r"""url\(\s*['"]?([^'"\)\s]+)['"]?\s*\)""", re.IGNORECASE)


def _resolve_epub_path(base_zip_path: str, href: str):
    parsed = urlparse(href)
    if parsed.scheme and parsed.scheme not in ("", "file"):
        return None
    if not parsed.path:
        return None
    raw_path = unquote(parsed.path)
    base_dir = str(PurePosixPath(base_zip_path).parent)
    resolved = raw_path if base_dir == "." else str(PurePosixPath(base_dir) / raw_path)
    parts = []
    for part in resolved.split("/"):
        if part == "..":
            if parts:
                parts.pop()
        elif part and part != ".":
            parts.append(part)
    return "/".join(parts)


def _collect_links_from_html(zip_path, text):
    links = []
    try:
        root = ET.fromstring(text.encode("utf-8", errors="replace"))
        for elem in root.iter():
            local_tag = elem.tag.split("}")[-1].lower() if "}" in elem.tag else elem.tag.lower()
            for attr in _LINK_ATTRS.get(local_tag, []):
                val = elem.get(attr, "").strip()
                if val:
                    if attr == "srcset":
                        for part in val.split(","):
                            candidate = part.strip().split()[0]
                            if candidate:
                                links.append(candidate)
                    else:
                        links.append(val)
    except ET.ParseError:
        for attr in ("href", "src", "data", "poster", "srcset", "cite"):
            for m in re.finditer(rf"""{attr}\s*=\s*['"]([^'"]+)['"]""", text, re.IGNORECASE):
                links.append(m.group(1).strip())
    for m in _CSS_URL_RE.finditer(text):
        links.append(m.group(1).strip())
    return links


def _collect_links_from_css(text):
    return [m.group(1).strip() for m in _CSS_URL_RE.finditer(text)]


def _collect_links_from_ncx(text):
    links = []
    try:
        root = ET.fromstring(text.encode("utf-8", errors="replace"))
        for elem in root.iter():
            local = elem.tag.split("}")[-1].lower() if "}" in elem.tag else elem.tag.lower()
            if local == "content":
                src = elem.get("src", "").strip()
                if src:
                    links.append(src)
    except ET.ParseError:
        for m in re.finditer(r"""src\s*=\s*['"]([^'"]+)['"]""", text, re.IGNORECASE):
            links.append(m.group(1).strip())
    return links


def _collect_links_from_nav(text):
    links = []
    try:
        root = ET.fromstring(text.encode("utf-8", errors="replace"))
        for elem in root.iter():
            local = elem.tag.split("}")[-1].lower() if "}" in elem.tag else elem.tag.lower()
            if local == "a":
                href = (elem.get("href") or "").strip()
                if href:
                    links.append(href)
    except ET.ParseError:
        for m in re.finditer(r"""href\s*=\s*['"]([^'"]+)['"]""", text, re.IGNORECASE):
            links.append(m.group(1).strip())
    return links


class LinkCheckResult:
    def __init__(self):
        self.total_links = 0
        self.external_links = 0
        self.fragment_links = 0
        self.internal_ok = 0
        self.broken = []
        self.encrypted_remaining = []
        self.warnings = []

    @property
    def has_errors(self):
        return bool(self.broken) or bool(self.encrypted_remaining)

    def summary(self):
        lines = [
            f"Links audited  : {self.total_links}",
            f"  External URLs : {self.external_links}",
            f"  Fragment-only : {self.fragment_links}",
            f"  Internal OK   : {self.internal_ok}",
            f"  Broken        : {len(self.broken)}",
        ]
        if self.encrypted_remaining:
            lines.append(f"  Still encrypted: {len(self.encrypted_remaining)} file(s)")
        if self.broken:
            lines.append("Broken links:")
            for src, href, resolved in self.broken[:20]:
                lines.append(f"  [{src}] -> {href!r}  (resolved: {resolved!r})")
            if len(self.broken) > 20:
                lines.append(f"  ... and {len(self.broken) - 20} more.")
        if self.warnings:
            lines.append("Warnings:")
            for w in self.warnings:
                lines.append(f"  {w}")
        return "\n".join(lines)


def verify_epub_links(epub_path):
    result = LinkCheckResult()
    if not epub_path.exists():
        result.warnings.append(f"EPUB file not found: {epub_path}")
        return result
    try:
        zf = zipfile.ZipFile(epub_path, "r")
    except zipfile.BadZipFile as e:
        result.warnings.append(f"Cannot open EPUB as zip: {e}")
        return result
    with zf:
        zip_names_lower = {n.lower(): n for n in zf.namelist()}
        zip_names_set = set(zf.namelist())
        def zip_has(path):
            return path in zip_names_set or path.lower() in zip_names_lower
        if "META-INF/encryption.xml" in zip_names_set:
            try:
                enc_xml = zf.read("META-INF/encryption.xml").decode("utf-8", errors="replace")
                enc_root = ET.fromstring(enc_xml)
                for elem in enc_root.iter():
                    local = elem.tag.split("}")[-1].lower() if "}" in elem.tag else elem.tag.lower()
                    if local == "cipherreference":
                        uri = elem.get("URI", "").strip()
                        if uri:
                            result.encrypted_remaining.append(uri)
            except Exception as e:
                result.warnings.append(f"Could not parse encryption.xml: {e}")
        opf_path = None
        if "META-INF/container.xml" in zip_names_set:
            try:
                container_xml = zf.read("META-INF/container.xml").decode("utf-8", errors="replace")
                c_root = ET.fromstring(container_xml)
                for elem in c_root.iter():
                    local = elem.tag.split("}")[-1].lower() if "}" in elem.tag else elem.tag.lower()
                    if local == "rootfile":
                        opf_path = elem.get("full-path", "").strip()
                        break
            except Exception:
                pass
        if not opf_path:
            opf_path = next((n for n in zf.namelist() if n.endswith(".opf")), None)
        manifest_items = {}
        spine_items = []
        nav_path = None
        ncx_path = None
        if opf_path:
            try:
                opf_xml = zf.read(opf_path).decode("utf-8", errors="replace")
                opf_root = ET.fromstring(opf_xml)
                for elem in opf_root.iter():
                    local = elem.tag.split("}")[-1].lower() if "}" in elem.tag else elem.tag.lower()
                    if local == "item":
                        item_id = elem.get("id", "")
                        href = elem.get("href", "").strip()
                        if href:
                            resolved = _resolve_epub_path(opf_path, href) or href
                            manifest_items[item_id] = resolved
                            props = elem.get("properties", "")
                            media_type = elem.get("media-type", "")
                            if "nav" in props:
                                nav_path = resolved
                            if media_type == "application/x-dtbncx+xml" or href.endswith(".ncx"):
                                ncx_path = resolved
                            result.total_links += 1
                            if not zip_has(resolved):
                                result.broken.append((opf_path, href, resolved))
                            else:
                                result.internal_ok += 1
                    elif local == "itemref":
                        idref = elem.get("idref", "")
                        if idref in manifest_items:
                            spine_items.append(manifest_items[idref])
            except Exception as e:
                result.warnings.append(f"Could not parse OPF: {e}")
        for zip_entry in zf.namelist():
            lower = zip_entry.lower()
            is_html = lower.endswith((".xhtml", ".html", ".htm", ".xml"))
            is_css = lower.endswith(".css")
            is_ncx = lower.endswith(".ncx") or zip_entry == ncx_path
            is_nav = zip_entry == nav_path
            if not (is_html or is_css or is_ncx or is_nav):
                continue
            try:
                text = zf.read(zip_entry).decode("utf-8", errors="replace")
            except Exception as e:
                result.warnings.append(f"Cannot read {zip_entry}: {e}")
                continue
            if is_css:
                raw_links = _collect_links_from_css(text)
            elif is_ncx:
                raw_links = _collect_links_from_ncx(text)
            elif is_nav:
                raw_links = _collect_links_from_nav(text) + _collect_links_from_html(zip_entry, text)
            else:
                raw_links = _collect_links_from_html(zip_entry, text)
            for href in raw_links:
                if not href:
                    continue
                result.total_links += 1
                parsed = urlparse(href)
                if parsed.scheme and parsed.scheme not in ("", "file"):
                    result.external_links += 1
                    continue
                if not parsed.path:
                    result.fragment_links += 1
                    continue
                resolved = _resolve_epub_path(zip_entry, href)
                if resolved is None:
                    result.external_links += 1
                    continue
                if zip_has(resolved):
                    result.internal_ok += 1
                else:
                    result.broken.append((zip_entry, href, resolved))
    return result


# --- Pipeline -------------------------------------------------------------


def convert_pipeline(acsm_path, output_dir):
    """Generator yielding (step, message) tuples.

    PDF path: steps 1-7 (step 7 = OCR if needed)
    EPUB path: steps 1-6 (no OCR)
    """
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
    yield (2, f"Detected format: {fmt.upper()}")

    # Step 3: Register device
    register_device()
    yield (3, "Device registered.")

    # Step 4: Download
    ext = ".pdf" if fmt == "pdf" else ".epub"
    drm_file = output_dir / f"{stem}_drm{ext}"
    fulfill_acsm(acsm_path, drm_file)
    yield (4, f"Downloaded: {drm_file.name}")

    # Step 5: Remove DRM
    output_file = output_dir / f"{stem}{ext}"
    remove_drm(drm_file, output_file)
    drm_file.unlink()
    yield (5, f"DRM removed: {output_file.name}")

    # Step 6: Verify readability
    if fmt == "pdf":
        print("Verifying PDF readability...")
        pdf_result = verify_pdf_readability(output_file)

        if pdf_result.encrypted:
            raise RuntimeError(
                "DRM removal incomplete: the PDF is still encrypted."
            )

        if pdf_result.probably_image_only:
            yield (6, (
                f"PDF scan: 0/{pdf_result.total_pages} pages have extractable text. "
                f"Image-only PDF detected -- will run OCR to add text layer."
            ))
        elif pdf_result.needs_ocr:
            img_count = len(pdf_result.pages_image_only)
            yield (6, (
                f"PDF scan: {pdf_result.pages_with_text}/{pdf_result.total_pages} pages "
                f"have text, {img_count} page(s) are image-only -- will OCR those pages."
            ))
        else:
            yield (6, (
                f"PDF verified: {pdf_result.pages_with_text}/{pdf_result.total_pages} pages "
                f"have readable, selectable text -- all OK. No OCR needed."
            ))

        # Step 7: OCR (if needed)
        if pdf_result.needs_ocr or pdf_result.probably_image_only:
            print("Running OCR to add text layer...")
            ocr_output = output_dir / f"{stem}_ocr.pdf"
            try:
                ocr_result = run_ocr(
                    input_pdf=output_file,
                    output_pdf=ocr_output,
                    language="auto",
                    dpi=300,
                    pages_to_ocr=pdf_result.pages_image_only or None,
                )
                if ocr_result["status"] == "already_has_text":
                    yield (7, (
                        f"OCR skipped -- all pages already have a text layer. "
                        f"Language: {ocr_result['lang_label']}"
                    ))
                else:
                    output_file.unlink()
                    ocr_output.rename(output_file)
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
            except RuntimeError as e:
                if ocr_output.exists():
                    ocr_output.unlink()
                yield (7, f"OCR failed: {e}. PDF available without text layer.")
        else:
            yield (7, "OCR not needed -- PDF is already fully readable.")

    else:
        # EPUB link verification
        print("Verifying link integrity...")
        link_result = verify_epub_links(output_file)
        if link_result.encrypted_remaining:
            files = ", ".join(link_result.encrypted_remaining[:5])
            raise RuntimeError(
                f"DRM removal incomplete: {len(link_result.encrypted_remaining)} file(s) "
                f"are still encrypted ({files})."
            )
        if link_result.broken:
            broken_count = len(link_result.broken)
            sample = link_result.broken[0]
            yield (6, (
                f"Link check: {link_result.internal_ok} OK, "
                f"{broken_count} broken (e.g. [{sample[0]}]->{sample[1]!r}). "
                f"EPUB usable but some links may not work."
            ))
        else:
            yield (6, (
                f"Links verified: {link_result.internal_ok} internal, "
                f"{link_result.external_links} external, "
                f"{link_result.fragment_links} anchors -- all OK."
            ))

    # Done
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
        description="Convert ACSM ebook tokens to DRM-free EPUB or PDF.",
        epilog="First run: python3 converter.py --setup",
    )
    parser.add_argument("acsm_file", nargs="?", help="Path to the .acsm file")
    parser.add_argument("--setup", action="store_true", help="Install dependencies and build tools")
    parser.add_argument("-o", "--output-dir", default="output", help="Output directory")
    parser.add_argument("--verify-only", metavar="FILE", help="Audit an existing EPUB or PDF")
    parser.add_argument("--ocr-only", metavar="FILE", help="Run OCR on an existing PDF")
    parser.add_argument("--ocr-lang", default="auto",
                        help="OCR language: auto, eng, chi_tra, chi_sim, jpn, kor (default: auto)")
    args = parser.parse_args()

    if args.verify_only:
        path = Path(args.verify_only)
        if path.suffix.lower() == ".pdf":
            result = verify_pdf_readability(path)
            print(result.summary())
            sys.exit(1 if result.has_errors else 0)
        else:
            result = verify_epub_links(path)
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

    if args.setup:
        do_setup()
        return
    if not args.acsm_file:
        parser.print_help()
        sys.exit(1)
    do_convert(args.acsm_file, args.output_dir)


if __name__ == "__main__":
    main()
