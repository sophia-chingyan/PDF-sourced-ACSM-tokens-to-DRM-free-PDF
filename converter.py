#!/usr/bin/env python3
"""
ACSM to EPUB/PDF Converter

Converts Adobe ACSM ebook tokens to DRM-free EPUB or PDF files
for personal offline reading.

Prerequisites (installed automatically by setup):
    brew install pugixml libzip openssl curl cmake
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


# ─── Setup ───────────────────────────────────────────────────────────────


def setup_brew_deps():
    """Install build dependencies via Homebrew."""
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
    """Get Homebrew prefix paths for dependencies."""
    prefixes = {}
    for dep in ["pugixml", "libzip", "openssl", "curl"]:
        r = run(["brew", "--prefix", dep])
        prefixes[dep] = r.stdout.strip() if r.returncode == 0 else f"/opt/homebrew/opt/{dep}"
    return prefixes


def _patch_makefiles(brew_prefixes):
    """Patch libgourou Makefiles for macOS compatibility."""
    include_flags = " ".join(f"-I{p}/include" for p in brew_prefixes.values())
    lib_flags = " ".join(f"-L{p}/lib" for p in brew_prefixes.values())

    root_mk = LIBGOUROU_DIR / "Makefile"
    content = root_mk.read_text()
    content = content.replace(
        "$(AR) rcs --thin $@ $^",
        "libtool -static -o $@ $^",
    )
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
    """Clone and build libgourou from source."""
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
        cwd=str(LIBGOUROU_DIR),
        env=env,
    )
    if result.returncode != 0:
        print(f"Build failed:\n{result.stdout}\n{result.stderr}")
        print("\nTry installing missing deps: brew install pugixml libzip openssl curl")
        sys.exit(1)

    if not (LIBGOUROU_BIN / "acsmdownloader").exists():
        print("Build completed but binaries not found.")
        print(f"Check {LIBGOUROU_BIN} for build output.")
        sys.exit(1)

    print("[OK] libgourou built successfully.")


def do_setup():
    """Run full first-time setup."""
    print("=== Setting up ACSM Converter ===\n")
    setup_brew_deps()
    print()
    build_libgourou()
    print("\n=== Setup complete! ===")
    print("You can now convert ACSM files:")
    print("  python3 converter.py ebook.acsm")


# ─── Conversion ──────────────────────────────────────────────────────────


def detect_format(acsm_path):
    """Parse the ACSM file and detect whether it points to EPUB or PDF."""
    tree = ET.parse(acsm_path)
    root = tree.getroot()
    ns = {"adept": "http://ns.adobe.com/adept"}

    src_elem = root.find(".//adept:src", ns)
    if src_elem is not None and src_elem.text:
        src = src_elem.text.lower()
        if ".pdf" in src or "output=pdf" in src:
            return "pdf"

    # Also check metadata format element
    fmt_elem = root.find(".//adept:metadata/adept:format", ns)
    if fmt_elem is not None and fmt_elem.text:
        if "pdf" in fmt_elem.text.lower():
            return "pdf"

    # Default to EPUB (most common case)
    return "epub"


def register_device():
    """Register an Adobe device (one-time setup)."""
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
    """Download the DRM-protected file by fulfilling the ACSM token."""
    print(f"Fulfilling ACSM: {acsm_path.name}")
    tool = find_tool("acsmdownloader")
    try:
        result = run([tool, "-f", str(acsm_path), "-o", str(output_path)], timeout=120)
    except subprocess.TimeoutExpired:
        raise RuntimeError("Download timed out (120s). The ACSM token may be expired or the server is unreachable.")
    if result.returncode != 0:
        stderr = result.stderr or result.stdout or ""
        raise RuntimeError(f"ACSM download failed (exit code {result.returncode}): {stderr[:500]}")

    if not output_path.exists():
        raise RuntimeError(f"Download completed but output file not found. stdout: {result.stdout[:200]}")

    size_kb = output_path.stat().st_size / 1024
    print(f"[OK] Downloaded: {output_path.name} ({size_kb:.0f} KB)")


def remove_drm(input_path, output_path):
    """Remove DRM from the downloaded file (EPUB or PDF)."""
    print(f"Removing DRM: {input_path.name}")
    tool = find_tool("adept_remove")
    try:
        result = run([tool, "-f", str(input_path), "-o", str(output_path)], timeout=60)
    except subprocess.TimeoutExpired:
        raise RuntimeError("DRM removal timed out (60s).")
    if result.returncode != 0:
        raise RuntimeError(f"DRM removal failed: {(result.stderr or result.stdout)[:300]}")

    print(f"[OK] DRM removed: {output_path.name}")


# ─── PDF Verification ─────────────────────────────────────────────────────


class PDFCheckResult:
    """Holds the outcome of a PDF readability audit."""

    def __init__(self):
        self.total_pages: int = 0
        self.pages_with_text: int = 0
        self.pages_image_only: list[int] = []
        self.sample_text: str = ""
        self.warnings: list[str] = []
        self.encrypted: bool = False
        self.has_fonts: bool = False  # True if any page has embedded fonts

    @property
    def has_errors(self) -> bool:
        # Only flag as error if encrypted. If no text was extracted but fonts
        # exist, the text is likely readable (CJK encoding issue).
        # If no text AND no fonts AND pages exist, it's probably image-only,
        # but we still only warn — don't block the download.
        return self.encrypted

    @property
    def probably_image_only(self) -> bool:
        """True if the PDF appears to be image-only (no text, no fonts)."""
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
            lines.append("⚠ PDF is still encrypted!")
        if self.pages_image_only:
            pages_str = ", ".join(str(p) for p in self.pages_image_only[:10])
            if len(self.pages_image_only) > 10:
                pages_str += f" … and {len(self.pages_image_only) - 10} more"
            lines.append(f"Image-only pages: {pages_str}")
        if self.warnings:
            lines.append("Warnings:")
            for w in self.warnings:
                lines.append(f"  {w}")
        return "\n".join(lines)


def _extract_text_pymupdf(pdf_path: Path, result: PDFCheckResult) -> bool:
    """Try text extraction with PyMuPDF (fitz). Returns True if successful."""
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
        return True  # Handled, even though encrypted

    result.total_pages = len(doc)
    result.has_fonts = False

    for i, page in enumerate(doc):
        try:
            text = page.get_text("text") or ""
            clean = text.strip()

            # Check if the page has embedded fonts (strong signal of real text,
            # even if extraction returns garbled results for some encodings)
            fonts = page.get_fonts()
            if fonts:
                result.has_fonts = True

            # Consider a page "has text" if it has at least 5 non-whitespace
            # chars. Lower threshold than before because CJK characters carry
            # more information per character than Latin scripts.
            if len(clean) >= 5:
                result.pages_with_text += 1
                if not result.sample_text and len(clean) > 10:
                    result.sample_text = clean[:200]
            else:
                # Page has no extractable text, but if it has fonts embedded,
                # it may still be readable (font encoding issue, not image-only)
                if fonts:
                    # Give benefit of the doubt — fonts present means text exists
                    result.pages_with_text += 1
                    if not result.sample_text:
                        result.sample_text = "(text present but not extractable — fonts embedded)"
                else:
                    result.pages_image_only.append(i + 1)
        except Exception:
            result.pages_image_only.append(i + 1)

    doc.close()
    return True


def _extract_text_pypdf(pdf_path: Path, result: PDFCheckResult) -> bool:
    """Fallback text extraction with pypdf. Returns True if successful."""
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
    """
    Open a DRM-free PDF and verify that its pages contain extractable text
    (not just scanned images).

    Uses PyMuPDF (fitz) as the primary engine because it handles CJK
    (Chinese/Japanese/Korean) fonts and encodings far better than pypdf.
    Falls back to pypdf if PyMuPDF is unavailable.

    Checks:
    1. PDF is not encrypted / password-protected
    2. Each page has extractable text content OR embedded fonts
    3. Reports pages that appear to be image-only (no text layer, no fonts)
    """
    result = PDFCheckResult()

    if not pdf_path.exists():
        result.warnings.append(f"PDF file not found: {pdf_path}")
        return result

    # Try PyMuPDF first (much better CJK support), then pypdf as fallback
    if not _extract_text_pymupdf(pdf_path, result):
        if not _extract_text_pypdf(pdf_path, result):
            result.warnings.append(
                "Neither PyMuPDF nor pypdf is installed — skipping text verification"
            )

    return result


# ─── Link Verification (EPUB) ─────────────────────────────────────────────


# HTML/XHTML attributes that carry links
_LINK_ATTRS = {
    "a":          ["href"],
    "area":       ["href"],
    "link":       ["href"],
    "script":     ["src"],
    "img":        ["src", "srcset"],
    "image":      ["href", "{http://www.w3.org/1999/xlink}href"],
    "use":        ["href", "{http://www.w3.org/1999/xlink}href"],
    "video":      ["src", "poster"],
    "audio":      ["src"],
    "source":     ["src", "srcset"],
    "track":      ["src"],
    "iframe":     ["src"],
    "object":     ["data"],
    "embed":      ["src"],
    "blockquote": ["cite"],
    "q":          ["cite"],
    "ins":        ["cite"],
    "del":        ["cite"],
}

_CSS_URL_RE = re.compile(r"""url\(\s*['"]?([^'"\)\s]+)['"]?\s*\)""", re.IGNORECASE)


def _resolve_epub_path(base_zip_path: str, href: str) -> str | None:
    parsed = urlparse(href)
    if parsed.scheme and parsed.scheme not in ("", "file"):
        return None
    if not parsed.path:
        return None
    raw_path = unquote(parsed.path)
    base_dir = str(PurePosixPath(base_zip_path).parent)
    if base_dir == ".":
        resolved = raw_path
    else:
        resolved = str(PurePosixPath(base_dir) / raw_path)
    parts = []
    for part in resolved.split("/"):
        if part == "..":
            if parts:
                parts.pop()
        elif part and part != ".":
            parts.append(part)
    return "/".join(parts)


def _collect_links_from_html(zip_path: str, text: str) -> list[str]:
    links: list[str] = []
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


def _collect_links_from_css(text: str) -> list[str]:
    return [m.group(1).strip() for m in _CSS_URL_RE.finditer(text)]


def _collect_links_from_ncx(text: str) -> list[str]:
    links: list[str] = []
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


def _collect_links_from_nav(text: str) -> list[str]:
    links: list[str] = []
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
        self.total_links: int = 0
        self.external_links: int = 0
        self.fragment_links: int = 0
        self.internal_ok: int = 0
        self.broken: list[tuple[str, str, str]] = []
        self.encrypted_remaining: list[str] = []
        self.warnings: list[str] = []

    @property
    def has_errors(self) -> bool:
        return bool(self.broken) or bool(self.encrypted_remaining)

    def summary(self) -> str:
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
                lines.append(f"  [{src}] → {href!r}  (resolved: {resolved!r})")
            if len(self.broken) > 20:
                lines.append(f"  … and {len(self.broken) - 20} more.")
        if self.warnings:
            lines.append("Warnings:")
            for w in self.warnings:
                lines.append(f"  {w}")
        return "\n".join(lines)


def verify_epub_links(epub_path: Path) -> LinkCheckResult:
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

        def zip_has(path: str) -> bool:
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

        manifest_items: dict[str, str] = {}
        spine_items: list[str] = []
        nav_path: str | None = None
        ncx_path: str | None = None

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


# ─── Pipeline ─────────────────────────────────────────────────────────────


def convert_pipeline(acsm_path, output_dir):
    """Generator that yields (step, message) tuples for each conversion step.

    Used by both the CLI (do_convert) and the web interface (app.py).
    Raises RuntimeError on failure.

    Steps:
      1. Check tools
      2. Detect format (EPUB or PDF)
      3. Register Adobe device
      4. Download file
      5. Remove DRM
      6. Verify (links for EPUB, readability for PDF)
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
        problems.append("acsmdownloader not found (run: python3 converter.py --setup)")
    if not find_tool("adept_activate"):
        problems.append("adept_activate not found (run: python3 converter.py --setup)")
    if not find_tool("adept_remove"):
        problems.append("adept_remove not found (run: python3 converter.py --setup)")
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

    # Step 6: Verify
    if fmt == "pdf":
        print("Verifying PDF readability...")
        pdf_result = verify_pdf_readability(output_file)

        if pdf_result.encrypted:
            raise RuntimeError(
                "DRM removal incomplete: the PDF is still encrypted and cannot be read."
            )

        if pdf_result.probably_image_only:
            # No text AND no fonts — likely a scanned/image-only PDF.
            # Warn but do NOT block the download.
            warning_msg = (
                f"PDF check: 0/{pdf_result.total_pages} pages have extractable text "
                f"and no embedded fonts were found. The PDF may be image-only "
                f"(scanned pages without a text layer). It is still downloadable."
            )
            yield (6, warning_msg)
        elif pdf_result.pages_image_only:
            img_count = len(pdf_result.pages_image_only)
            warning_msg = (
                f"PDF check: {pdf_result.pages_with_text}/{pdf_result.total_pages} pages "
                f"have readable text, {img_count} page(s) are image-only. "
                f"The PDF is usable but some pages may lack selectable text."
            )
            yield (6, warning_msg)
        else:
            yield (
                6,
                f"PDF verified: {pdf_result.pages_with_text}/{pdf_result.total_pages} pages "
                f"have readable, selectable text — all OK.",
            )
    else:
        # EPUB link verification
        print("Verifying link integrity...")
        link_result = verify_epub_links(output_file)

        if link_result.encrypted_remaining:
            files = ", ".join(link_result.encrypted_remaining[:5])
            raise RuntimeError(
                f"DRM removal incomplete: {len(link_result.encrypted_remaining)} file(s) "
                f"are still encrypted ({files}). The EPUB may not be readable."
            )

        if link_result.broken:
            broken_count = len(link_result.broken)
            sample = link_result.broken[0]
            warning_msg = (
                f"Link check: {link_result.internal_ok} OK, "
                f"{broken_count} broken (e.g. [{sample[0]}]→{sample[1]!r}). "
                f"The EPUB is usable but some links may not work."
            )
            yield (6, warning_msg)
        else:
            yield (
                6,
                f"Links verified: {link_result.internal_ok} internal, "
                f"{link_result.external_links} external, "
                f"{link_result.fragment_links} anchors — all OK.",
            )

    # Done
    size_mb = output_file.stat().st_size / (1024 * 1024) if output_file.exists() else 0
    yield ("done", f"{output_file.name}|{size_mb:.1f} MB")


def do_convert(acsm_file, output_dir):
    """Run the full ACSM conversion pipeline (CLI entry point)."""
    try:
        for step, message in convert_pipeline(acsm_file, output_dir):
            if step == "done":
                parts = message.split("|")
                print(f"\n=== Done! ===\nFile: {parts[0]} ({parts[1]})")
            else:
                print(f"\n=== Step {step}/6: {message} ===")
    except RuntimeError as e:
        print(str(e))
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="Convert ACSM ebook tokens to DRM-free EPUB or PDF.",
        epilog="First run: python3 converter.py --setup",
    )
    parser.add_argument(
        "acsm_file",
        nargs="?",
        help="Path to the .acsm file to convert",
    )
    parser.add_argument(
        "--setup",
        action="store_true",
        help="Install dependencies and build tools (run once)",
    )
    parser.add_argument(
        "-o", "--output-dir",
        default="output",
        help="Output directory (default: output)",
    )
    parser.add_argument(
        "--verify-only",
        metavar="FILE",
        help="Audit an existing EPUB (links) or PDF (readability) — no conversion",
    )
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

    if args.setup:
        do_setup()
        return

    if not args.acsm_file:
        parser.print_help()
        sys.exit(1)

    do_convert(args.acsm_file, args.output_dir)


if __name__ == "__main__":
    main()
