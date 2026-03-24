#!/usr/bin/env python3
"""Flask web interface for the ACSM to EPUB/PDF converter."""

import json
import os
import threading
import time
import zipfile
import xml.etree.ElementTree as ET
from collections import OrderedDict
from functools import wraps
from pathlib import Path

from flask import Flask, jsonify, make_response, render_template, request, send_from_directory, session, redirect, url_for

from converter import convert_pipeline, run_ocr_step

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24).hex())

APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

SCRIPT_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = SCRIPT_DIR / "uploads"
OUTPUT_DIR = SCRIPT_DIR / "output"
COVER_DIR = SCRIPT_DIR / "covers"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
COVER_DIR.mkdir(exist_ok=True)

TOTAL_STEPS = 7

STEP_LABELS = {
    1: "Checking tools...",
    2: "Detecting format...",
    3: "Registering Adobe device...",
    4: "Downloading ebook...",
    5: "Removing DRM...",
    6: "Verifying readability...",
    7: "Running OCR (if needed)...",
}

# Track active conversions: job_id -> {steps: [...], status, error}
active_jobs = {}


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if APP_PASSWORD and not session.get("authenticated"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        session["authenticated"] = True
        return redirect(url_for("index"))
    error = None
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["authenticated"] = True
            return redirect(url_for("index"))
        error = "Wrong password"
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


def extract_epub_cover(epub_path):
    cover_out = COVER_DIR / f"{epub_path.stem}.jpg"
    if cover_out.exists():
        return cover_out.name
    try:
        with zipfile.ZipFile(epub_path, "r") as zf:
            cover_name = _find_cover_in_opf(zf) or _find_cover_by_name(zf)
            if cover_name:
                data = zf.read(cover_name)
                ext = Path(cover_name).suffix or ".jpg"
                cover_out = COVER_DIR / f"{epub_path.stem}{ext}"
                cover_out.write_bytes(data)
                return cover_out.name
    except Exception:
        pass
    return None


def extract_pdf_cover(pdf_path):
    """Try to extract the first page of a PDF as a cover thumbnail."""
    cover_out = COVER_DIR / f"{pdf_path.stem}.jpg"
    if cover_out.exists():
        return cover_out.name
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(str(pdf_path))
        if len(doc) > 0:
            page = doc[0]
            # Render at a moderate resolution
            mat = fitz.Matrix(1.5, 1.5)
            pix = page.get_pixmap(matrix=mat)
            pix.save(str(cover_out))
            doc.close()
            return cover_out.name
        doc.close()
    except ImportError:
        pass
    except Exception:
        pass
    return None


def _find_cover_in_opf(zf):
    opf_path = next((n for n in zf.namelist() if n.endswith(".opf")), None)
    if not opf_path:
        return None
    opf_xml = zf.read(opf_path).decode("utf-8", errors="replace")
    root = ET.fromstring(opf_xml)
    cover_id = None
    for meta in root.iter():
        if meta.tag.endswith("}meta") or meta.tag == "meta":
            if meta.get("name") == "cover":
                cover_id = meta.get("content")
                break
    if not cover_id:
        for item in root.iter():
            if item.tag.endswith("}item") or item.tag == "item":
                if "cover-image" in (item.get("properties") or ""):
                    href = item.get("href")
                    if href:
                        opf_dir = str(Path(opf_path).parent)
                        return href if opf_dir == "." else f"{opf_dir}/{href}"
        return None
    for item in root.iter():
        if item.tag.endswith("}item") or item.tag == "item":
            if item.get("id") == cover_id:
                href = item.get("href")
                if href:
                    opf_dir = str(Path(opf_path).parent)
                    return href if opf_dir == "." else f"{opf_dir}/{href}"
    return None


def _find_cover_by_name(zf):
    for name in zf.namelist():
        lower = name.lower()
        if "cover" in lower and any(lower.endswith(ext) for ext in (".jpg", ".jpeg", ".png")):
            return name
    return None


def get_books():
    if not OUTPUT_DIR.exists():
        return []
    books = OrderedDict()
    for f in sorted(OUTPUT_DIR.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True):
        if f.suffix in (".epub", ".pdf"):
            stem = f.stem
            if stem not in books:
                books[stem] = {"stem": stem, "files": [], "cover": None}
            size_mb = f.stat().st_size / (1024 * 1024)
            ext_label = f.suffix[1:].upper()
            books[stem]["files"].append({
                "name": f.name,
                "size": f"{size_mb:.1f} MB",
                "ext": ext_label,
            })
            if not books[stem]["cover"]:
                if f.suffix == ".epub":
                    cover = extract_epub_cover(f)
                else:
                    cover = extract_pdf_cover(f)
                if cover:
                    books[stem]["cover"] = cover
    return list(books.values())


def run_conversion_job(job_id, acsm_path, output_dir):
    """Run conversion in a background thread, updating active_jobs."""
    import traceback
    job = active_jobs[job_id]
    print(f"[DEBUG] Job {job_id} started: acsm={acsm_path}, output={output_dir}", flush=True)
    try:
        job["current_step"] = 1
        job["current_label"] = STEP_LABELS[1]

        for step, message in convert_pipeline(str(acsm_path), str(output_dir)):
            print(f"[DEBUG] Job {job_id} step={step} message={message}", flush=True)
            if step == "done":
                job["steps"].append({"step": "done", "message": message})
                job["status"] = "done"
                job["done_message"] = message
            elif step == "ocr_prompt":
                # Pause — let user decide whether to run OCR
                parts = message.split("|")
                job["ocr_pdf_path"] = parts[0]
                job["ocr_pages"] = parts[1]
                job["status"] = "pending_ocr"
                job["current_step"] = 7
                job["current_label"] = "Waiting for your decision..."
                print(f"[DEBUG] Job {job_id} paused for OCR decision", flush=True)
                return  # stop the pipeline here
            else:
                step_num = int(step)
                is_warning = (
                    (step_num == 6 and ("broken" in message.lower() or "image-only" in message.lower()))
                    or (step_num == 7 and ("failed" in message.lower() or "could not" in message.lower()))
                )
                job["steps"].append({
                    "step": step_num,
                    "message": message,
                    "warning": is_warning,
                })
                next_step = step_num + 1
                if next_step <= TOTAL_STEPS:
                    job["current_step"] = next_step
                    job["current_label"] = STEP_LABELS.get(next_step, "")
    except RuntimeError as e:
        print(f"[DEBUG] Job {job_id} RuntimeError: {e}", flush=True)
        job["status"] = "error"
        job["error"] = str(e)
    except Exception as e:
        print(f"[DEBUG] Job {job_id} Exception: {e}\n{traceback.format_exc()}", flush=True)
        job["status"] = "error"
        job["error"] = f"Unexpected error: {e}"


def run_ocr_job(job_id):
    """Run OCR step in a background thread after user confirms."""
    import traceback
    job = active_jobs[job_id]
    pdf_path = job["ocr_pdf_path"]
    pages = [int(p) for p in job["ocr_pages"].split(",")]

    print(f"[DEBUG] Job {job_id} OCR started: pdf={pdf_path}, pages={len(pages)}", flush=True)
    try:
        job["status"] = "running"
        job["current_step"] = 7
        job["current_label"] = STEP_LABELS[7]

        for step, message in run_ocr_step(pdf_path, pages):
            print(f"[DEBUG] Job {job_id} OCR step={step} message={message}", flush=True)
            if step == "done":
                job["steps"].append({"step": "done", "message": message})
                job["status"] = "done"
                job["done_message"] = message
            else:
                step_num = int(step)
                is_warning = "failed" in message.lower() or "could not" in message.lower()
                job["steps"].append({
                    "step": step_num,
                    "message": message,
                    "warning": is_warning,
                })
    except Exception as e:
        print(f"[DEBUG] Job {job_id} OCR Exception: {e}\n{traceback.format_exc()}", flush=True)
        job["status"] = "error"
        job["error"] = f"OCR error: {e}"


@app.route("/")
@login_required
def index():
    books = get_books()
    resp = make_response(render_template("index.html", books=books))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp


@app.route("/upload", methods=["POST"])
@login_required
def upload():
    file = request.files.get("file")
    if not file or not file.filename:
        return {"error": "No file provided"}, 400
    if not file.filename.endswith(".acsm"):
        return {"error": "Only .acsm files are accepted"}, 400
    filename = Path(file.filename).name
    save_path = UPLOAD_DIR / filename
    file.save(save_path)
    return {"filename": filename}


@app.route("/start-convert/<filename>", methods=["POST"])
@login_required
def start_convert(filename):
    """Start conversion in background, return a job ID for polling."""
    filename = Path(filename).name
    acsm_path = UPLOAD_DIR / filename

    if not acsm_path.exists():
        return jsonify({"error": "File not found"}), 404

    job_id = f"{filename}_{int(time.time())}"
    active_jobs[job_id] = {
        "filename": filename,
        "status": "running",
        "steps": [],
        "current_step": 0,
        "current_label": "",
        "error": None,
        "done_message": None,
        "start_time": time.time(),
        "ocr_pdf_path": None,
        "ocr_pages": None,
    }

    t = threading.Thread(
        target=run_conversion_job,
        args=(job_id, acsm_path, OUTPUT_DIR),
        daemon=True,
    )
    t.start()

    return jsonify({"job_id": job_id})


@app.route("/job-status/<job_id>")
@login_required
def job_status(job_id):
    """Poll endpoint: returns current conversion progress."""
    if job_id not in active_jobs:
        return jsonify({"error": "Job not found"}), 404

    job = active_jobs[job_id]
    elapsed = round(time.time() - job["start_time"])

    resp_data = {
        "status": job["status"],
        "steps": job["steps"],
        "current_step": job["current_step"],
        "current_label": job["current_label"],
        "error": job["error"],
        "done_message": job["done_message"],
        "elapsed": elapsed,
    }

    # Include OCR info when waiting for user decision
    if job["status"] == "pending_ocr":
        pdf_path = job.get("ocr_pdf_path", "")
        pages_str = job.get("ocr_pages", "")
        page_count = len(pages_str.split(",")) if pages_str else 0
        resp_data["ocr_info"] = {
            "filename": Path(pdf_path).name if pdf_path else "",
            "page_count": page_count,
        }

    return jsonify(resp_data)


@app.route("/ocr-decision/<job_id>", methods=["POST"])
@login_required
def ocr_decision(job_id):
    """User decides whether to run OCR or skip."""
    if job_id not in active_jobs:
        return jsonify({"error": "Job not found"}), 404

    job = active_jobs[job_id]
    if job["status"] != "pending_ocr":
        return jsonify({"error": "Job is not waiting for OCR decision"}), 400

    data = request.get_json(silent=True) or {}
    choice = data.get("choice", "skip")

    if choice == "ocr":
        # Run OCR in background thread
        t = threading.Thread(
            target=run_ocr_job,
            args=(job_id,),
            daemon=True,
        )
        t.start()
        return jsonify({"status": "ocr_started"})

    else:
        # Skip OCR — mark as done with current file
        pdf_path = Path(job.get("ocr_pdf_path", ""))
        size_mb = pdf_path.stat().st_size / (1024 * 1024) if pdf_path.exists() else 0
        done_msg = f"{pdf_path.name}|{size_mb:.1f} MB"

        job["steps"].append({
            "step": 7,
            "message": "OCR skipped by user -- PDF downloaded as image-only.",
            "warning": False,
        })
        job["steps"].append({"step": "done", "message": done_msg})
        job["status"] = "done"
        job["done_message"] = done_msg
        return jsonify({"status": "skipped"})


@app.route("/download/<filename>")
@login_required
def download(filename):
    filename = Path(filename).name
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        return {"error": "File not found"}, 404
    resp = send_from_directory(OUTPUT_DIR, filename, as_attachment=True)

    @resp.call_on_close
    def cleanup():
        stem = Path(filename).stem
        try:
            file_path.unlink(missing_ok=True)
        except Exception:
            pass
        for d in (UPLOAD_DIR, COVER_DIR):
            for f in d.iterdir():
                if f.stem == stem or f.stem.startswith(stem):
                    try:
                        f.unlink(missing_ok=True)
                    except Exception:
                        pass

    return resp


@app.route("/cover/<filename>")
@login_required
def cover(filename):
    filename = Path(filename).name
    return send_from_directory(COVER_DIR, filename)


@app.route("/debug-status")
@login_required
def debug_status():
    """Debug endpoint to check server state."""
    import shutil
    jobs_summary = {}
    for jid, job in active_jobs.items():
        jobs_summary[jid] = {
            "status": job["status"],
            "steps_count": len(job["steps"]),
            "current_step": job["current_step"],
            "error": job["error"],
            "elapsed": round(time.time() - job["start_time"]),
        }
    upload_files = [f.name for f in UPLOAD_DIR.iterdir()] if UPLOAD_DIR.exists() else []
    output_files = [f.name for f in OUTPUT_DIR.iterdir()] if OUTPUT_DIR.exists() else []
    return jsonify({
        "active_jobs": jobs_summary,
        "upload_files": upload_files,
        "output_files": output_files,
        "acsmdownloader_found": shutil.which("acsmdownloader") or str(Path("libgourou/utils/acsmdownloader")),
        "libgourou_exists": (SCRIPT_DIR / "libgourou" / "utils" / "acsmdownloader").exists(),
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)
