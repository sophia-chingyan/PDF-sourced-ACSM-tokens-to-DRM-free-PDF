#!/usr/bin/env python3
"""Flask web interface for the ACSM-to-PDF converter."""

import os
import threading
import time
from collections import OrderedDict
from functools import wraps
from pathlib import Path

from flask import (
    Flask, jsonify, make_response, render_template, request,
    send_from_directory, session, redirect, url_for,
)
from authlib.integrations.flask_client import OAuth

from converter import convert_pipeline

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24).hex())

# ── Google OAuth config ────────────────────────────────────────────────────
#
# Set these three env vars in Zeabur:
#   GOOGLE_CLIENT_ID     — from Google Cloud Console → Credentials
#   GOOGLE_CLIENT_SECRET — from Google Cloud Console → Credentials
#   ALLOWED_EMAIL        — your Google account email (e.g. you@gmail.com)
#
# In Google Cloud Console → Credentials → OAuth 2.0 Client:
#   Authorised redirect URI: https://<your-zeabur-domain>/auth/google/callback
#
GOOGLE_CLIENT_ID     = os.environ.get("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET", "")
ALLOWED_EMAIL        = os.environ.get("ALLOWED_EMAIL", "")

oauth = OAuth(app)
oauth.register(
    name="google",
    client_id=GOOGLE_CLIENT_ID,
    client_secret=GOOGLE_CLIENT_SECRET,
    server_metadata_url="https://accounts.google.com/.well-known/openid-configuration",
    client_kwargs={"scope": "openid email profile"},
)

# ── Paths ──────────────────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = SCRIPT_DIR / "uploads"
OUTPUT_DIR = SCRIPT_DIR / "output"
COVER_DIR  = SCRIPT_DIR / "covers"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
COVER_DIR.mkdir(exist_ok=True)

TOTAL_STEPS = 6

STEP_LABELS = {
    1: "Checking tools...",
    2: "Detecting format...",
    3: "Registering Adobe device...",
    4: "Downloading ebook...",
    5: "Removing DRM...",
    6: "Verifying readability...",
}

active_jobs = {}
_active_jobs_lock = threading.Lock()


# ── Auth helpers ───────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("authenticated"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


@app.route("/login")
def login():
    # If OAuth is not configured, show a helpful error instead of a blank page.
    if not GOOGLE_CLIENT_ID or not GOOGLE_CLIENT_SECRET:
        error = (
            "Google OAuth is not configured. "
            "Set GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, and ALLOWED_EMAIL "
            "in your Zeabur environment variables."
        )
        return render_template("login.html", error=error)
    return render_template("login.html", error=None)


@app.route("/login/google")
def login_google():
    # Force the correct redirect URI explicitly — do not rely on url_for()
    # which can produce http:// behind a reverse proxy like Zeabur.
    base = os.environ.get("APP_BASE_URL", "").rstrip("/")
    if base:
        redirect_uri = f"{base}/auth/google/callback"
    else:
        redirect_uri = url_for("auth_callback", _external=True, _scheme="https")
    print(f"[DEBUG] OAuth redirect_uri = {redirect_uri}", flush=True)
    return oauth.google.authorize_redirect(redirect_uri)


@app.route("/auth/google/callback")
def auth_callback():
    try:
        token = oauth.google.authorize_access_token()
    except Exception as e:
        return render_template("login.html", error=f"OAuth error: {e}")

    user_info = token.get("userinfo")
    if not user_info:
        return render_template("login.html", error="Could not retrieve user info from Google.")

    email = user_info.get("email", "").lower().strip()
    allowed = ALLOWED_EMAIL.lower().strip()

    if not allowed:
        return render_template(
            "login.html",
            error="ALLOWED_EMAIL is not set. Add it to your Zeabur environment variables.",
        )

    if email != allowed:
        return render_template(
            "login.html",
            error=f"Access denied: {email} is not authorised to use this app.",
        )

    session["authenticated"] = True
    session["user_email"] = email
    session["user_name"] = user_info.get("name", email)
    return redirect(url_for("index"))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ── Library helpers ────────────────────────────────────────────────────────

def extract_pdf_cover(pdf_path):
    cover_out = COVER_DIR / f"{pdf_path.stem}.jpg"
    if cover_out.exists():
        return cover_out.name
    try:
        import fitz
        doc = fitz.open(str(pdf_path))
        if len(doc) > 0:
            page = doc[0]
            mat  = fitz.Matrix(1.5, 1.5)
            pix  = page.get_pixmap(matrix=mat)
            pix.save(str(cover_out))
            doc.close()
            return cover_out.name
        doc.close()
    except ImportError:
        pass
    except Exception:
        pass
    return None


def get_books():
    if not OUTPUT_DIR.exists():
        return [], 0
    books = OrderedDict()
    total_files = 0
    for f in sorted(OUTPUT_DIR.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True):
        if f.suffix != ".pdf":
            continue
        stem = f.stem
        if not stem:
            continue
        if stem not in books:
            books[stem] = {"stem": stem, "files": [], "cover": None}
        size_mb = f.stat().st_size / (1024 * 1024)
        books[stem]["files"].append({
            "name": f.name,
            "size": f"{size_mb:.1f} MB",
            "ext":  "PDF",
        })
        total_files += 1
        if not books[stem]["cover"]:
            cover = extract_pdf_cover(f)
            if cover:
                books[stem]["cover"] = cover
    return list(books.values()), total_files


def _prune_old_jobs():
    cutoff = time.time() - 7200
    with _active_jobs_lock:
        stale = [
            jid for jid, job in active_jobs.items()
            if job["status"] in ("done", "error") and job["start_time"] < cutoff
        ]
        for jid in stale:
            del active_jobs[jid]


# ── Conversion workers ─────────────────────────────────────────────────────

def run_conversion_job(job_id, acsm_path, output_dir):
    import traceback

    with _active_jobs_lock:
        job = active_jobs[job_id]

    print(f"[DEBUG] Job {job_id} started: acsm={acsm_path}, output={output_dir}", flush=True)
    try:
        job["current_step"]  = 1
        job["current_label"] = STEP_LABELS[1]

        for step, message in convert_pipeline(str(acsm_path), str(output_dir)):
            print(f"[DEBUG] Job {job_id} step={step} message={message}", flush=True)
            if step == "done":
                job["steps"].append({"step": "done", "message": message})
                job["status"]       = "done"
                job["done_message"] = message
            else:
                step_num   = int(step)
                is_warning = (
                    step_num == 6 and "broken" in message.lower()
                )
                job["steps"].append({"step": step_num, "message": message, "warning": is_warning})
                next_step = step_num + 1
                if next_step <= TOTAL_STEPS:
                    job["current_step"]  = next_step
                    job["current_label"] = STEP_LABELS.get(next_step, "")
    except RuntimeError as e:
        print(f"[DEBUG] Job {job_id} RuntimeError: {e}", flush=True)
        job["status"] = "error"
        job["error"]  = str(e)
    except Exception as e:
        print(f"[DEBUG] Job {job_id} Exception: {e}\n{traceback.format_exc()}", flush=True)
        job["status"] = "error"
        job["error"]  = f"Unexpected error: {e}"


# ── Routes ─────────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def index():
    books, total_files = get_books()
    resp = make_response(render_template("index.html", books=books))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp


@app.route("/library")
@login_required
def library():
    books, total_files = get_books()
    resp = make_response(render_template(
        "library.html",
        books=books,
        total_files=total_files,
    ))
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
    filename  = Path(file.filename).name
    save_path = UPLOAD_DIR / filename
    file.save(save_path)
    return {"filename": filename}


@app.route("/start-convert/<filename>", methods=["POST"])
@login_required
def start_convert(filename):
    _prune_old_jobs()
    filename  = Path(filename).name
    acsm_path = UPLOAD_DIR / filename
    if not acsm_path.exists():
        return jsonify({"error": "File not found"}), 404

    job_id = f"{filename}_{int(time.time())}"
    with _active_jobs_lock:
        active_jobs[job_id] = {
            "filename":     filename,
            "status":       "running",
            "steps":        [],
            "current_step": 0,
            "current_label": "",
            "error":        None,
            "done_message": None,
            "start_time":   time.time(),
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
    with _active_jobs_lock:
        if job_id not in active_jobs:
            return jsonify({"error": "Job not found"}), 404
        job = active_jobs[job_id]

    elapsed   = round(time.time() - job["start_time"])
    resp_data = {
        "status":        job["status"],
        "steps":         job["steps"],
        "current_step":  job["current_step"],
        "current_label": job["current_label"],
        "error":         job["error"],
        "done_message":  job["done_message"],
        "elapsed":       elapsed,
    }

    return jsonify(resp_data)


@app.route("/download/<filename>")
@login_required
def download(filename):
    filename  = Path(filename).name
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        return {"error": "File not found"}, 404
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


@app.route("/delete/<stem>", methods=["POST"])
@login_required
def delete_book(stem):
    stem = Path(stem).stem
    if not stem:
        return jsonify({"error": "Invalid stem"}), 400
    deleted = []
    for f in list(OUTPUT_DIR.iterdir()):
        if f.stem == stem and f.suffix == ".pdf":
            f.unlink(missing_ok=True)
            deleted.append(f.name)
    for d in (UPLOAD_DIR, COVER_DIR):
        for f in d.iterdir():
            if f.stem == stem:
                f.unlink(missing_ok=True)
    return jsonify({"deleted": deleted})


@app.route("/cover/<filename>")
@login_required
def cover(filename):
    filename = Path(filename).name
    return send_from_directory(COVER_DIR, filename)


@app.route("/debug-status")
@login_required
def debug_status():
    import shutil
    jobs_summary = {}
    with _active_jobs_lock:
        for jid, job in active_jobs.items():
            jobs_summary[jid] = {
                "status":       job["status"],
                "steps_count":  len(job["steps"]),
                "current_step": job["current_step"],
                "error":        job["error"],
                "elapsed":      round(time.time() - job["start_time"]),
            }
    upload_files = [f.name for f in UPLOAD_DIR.iterdir()] if UPLOAD_DIR.exists() else []
    output_files = [f.name for f in OUTPUT_DIR.iterdir()] if OUTPUT_DIR.exists() else []
    return jsonify({
        "active_jobs":          jobs_summary,
        "upload_files":         upload_files,
        "output_files":         output_files,
        "acsmdownloader_found": shutil.which("acsmdownloader") or str(Path("libgourou/utils/acsmdownloader")),
        "libgourou_exists":     (SCRIPT_DIR / "libgourou" / "utils" / "acsmdownloader").exists(),
        "logged_in_as":         session.get("user_email", "unknown"),
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)
