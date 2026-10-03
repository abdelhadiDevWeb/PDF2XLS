from __future__ import annotations

import hmac
import json
import logging
import os
import re
import shutil
import tempfile
import time
import uuid
from pathlib import Path

from flask import Flask, Response, abort, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from converter import SHEET_NAME, TESSERACT_AVAILABLE, convert_pdf, has_header

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("pdf2xls")

MAX_UPLOAD_MB = 50
JOB_TTL_SECONDS = 60 * 60
JOBS_DIR = Path(tempfile.gettempdir()) / "pdf2xls_jobs"
JOBS_DIR.mkdir(exist_ok=True)
JOB_ID = re.compile(r"^[0-9a-f]{32}$")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


@app.before_request
def require_password():
    """Optional protection for public deployments: set APP_PASSWORD to enable."""
    if not APP_PASSWORD or request.path == "/healthz":
        return None
    auth = request.authorization
    if auth and hmac.compare_digest((auth.password or "").encode(), APP_PASSWORD.encode()):
        return None
    return Response("Password required", 401, {"WWW-Authenticate": 'Basic realm="PDF to Excel"'})


def _cleanup_old_jobs() -> None:
    cutoff = time.time() - JOB_TTL_SECONDS
    for job_dir in JOBS_DIR.iterdir():
        try:
            expired = job_dir.is_dir() and job_dir.stat().st_mtime < cutoff
        except FileNotFoundError:  # removed concurrently by another worker
            continue
        if expired:
            shutil.rmtree(job_dir, ignore_errors=True)


def _preview(df, rows: int = 10) -> dict:
    data = json.loads(df.head(rows).to_json(orient="split", index=False))
    return {"columns": data["columns"], "rows": data["data"]}


@app.get("/healthz")
def healthz():
    return jsonify(status="ok", ocr=TESSERACT_AVAILABLE)


@app.get("/")
def index():
    return render_template("index.html", ocr_available=TESSERACT_AVAILABLE, max_mb=MAX_UPLOAD_MB)


@app.post("/api/convert")
def convert():
    _cleanup_old_jobs()

    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify(error="Please choose a PDF file."), 400
    if not upload.filename.lower().endswith(".pdf"):
        return jsonify(error="Only .pdf files are supported."), 400

    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir()
    pdf_path = job_dir / "input.pdf"
    upload.save(pdf_path)

    with pdf_path.open("rb") as fh:
        if fh.read(5) != b"%PDF-":
            shutil.rmtree(job_dir, ignore_errors=True)
            return jsonify(error="This file is not a valid PDF."), 400

    stem = Path(secure_filename(upload.filename)).stem or "converted"
    xlsx_path = job_dir / f"{stem}.xlsx"

    started = time.perf_counter()
    try:
        result = convert_pdf(pdf_path, xlsx_path, main_table_only=request.form.get("all_tables") != "1")
    except Exception as exc:
        log.exception("Conversion failed")
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify(error=f"Conversion failed: {exc}"), 500

    return jsonify(
        job_id=job_id,
        download_url=f"/api/download/{job_id}",
        filename=xlsx_path.name,
        seconds=round(time.perf_counter() - started, 1),
        sheet=SHEET_NAME,
        table_count=len(result.tables),
        pages=[vars(p) for p in result.pages],
        blocks=[
            {
                "pages": b.pages,
                "methods": sorted(set(b.methods)),
                "start_row": b.start_row,
                "rows": len(b.df),
                "cols": b.df.shape[1],
                "has_header": has_header(b.df),
                "preview": _preview(b.df),
            }
            for b in result.blocks
        ],
        warnings=result.warnings,
    )


@app.get("/api/download/<job_id>")
def download(job_id: str):
    if not JOB_ID.match(job_id):
        abort(404)
    files = list((JOBS_DIR / job_id).glob("*.xlsx"))
    if not files:
        abort(404)
    return send_file(files[0], as_attachment=True, download_name=files[0].name)


@app.errorhandler(413)
def too_large(_):
    return jsonify(error=f"File is too large (max {MAX_UPLOAD_MB} MB)."), 413


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
