from flask import (
    Flask, request, render_template_string, flash, redirect,
    send_file, url_for
)
from werkzeug.utils import secure_filename
import pytesseract
from PIL import Image
from textblob import TextBlob
from docx import Document
import uuid, io, os, tempfile, shutil

# ── Path to your local Tesseract executable ─────────────────────────────
pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)

# ── Flask setup ──────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = "change‑me‑in‑production"
ALLOWED_EXT = {"png", "jpg", "jpeg", "bmp", "tiff", "webp"}

# Temp storage for generated files
TMP_DIR = tempfile.mkdtemp(prefix="ocr_files_")

def allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXT

# ── HTML template (Bootstrap 5) ──────────────────────────────────────────
HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>OCR & Spell‑Correct</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css"
        rel="stylesheet">
</head>
<body class="bg-light py-5">
<div class="container">
  <h1 class="text-center mb-4">OCR → Spell‑Correct → Save&nbsp;as TXT/DOCX</h1>

  {% with m=get_flashed_messages() %}
    {% if m %}<div class="alert alert-warning">{{ m[0] }}</div>{% endif %}
  {% endwith %}

  <form method="POST" enctype="multipart/form-data" class="card p-4 shadow-sm">
    <div class="mb-3">
      <input class="form-control" type="file" name="image" required>
    </div>
    <div class="form-check mb-3">
      <input class="form-check-input" type="checkbox" name="spell" id="spell" checked>
      <label class="form-check-label" for="spell">Apply TextBlob spell‑correction</label>
    </div>
    <button class="btn btn-primary">Run OCR</button>
  </form>

  {% if text %}
  <div class="card mt-4 p-3 shadow-sm">
    <h5>Extracted&nbsp;Text{% if corrected %}&nbsp;(spell‑corrected){% endif %}:</h5>
    <pre style="white-space: pre-wrap;">{{ text }}</pre>
    <div class="mt-3">
      <a class="btn btn-outline-secondary me-2"
         href="{{ url_for('download', fid=fid, ext='txt') }}">Download TXT</a>
      <a class="btn btn-outline-secondary"
         href="{{ url_for('download', fid=fid, ext='docx') }}">Download DOCX</a>
    </div>
  </div>
  {% endif %}
</div>
</body>
</html>
"""

# ── Main page ────────────────────────────────────────────────────────────
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("image")
        if not file or file.filename == "":
            flash("No file selected.")
            return redirect(request.url)
        if not allowed(file.filename):
            flash("Unsupported file type.")
            return redirect(request.url)

        # ---------- OCR ----------
        img = Image.open(file.stream)
        raw_text = pytesseract.image_to_string(img)

        # ---------- Optional spell‑correction ----------
        apply_spell = request.form.get("spell") == "on"
        if apply_spell:
            corrected_text = str(TextBlob(raw_text).correct())
            text_out = corrected_text
        else:
            corrected_text = None
            text_out = raw_text

        # ---------- Persist result to temp files ----------
        fid = str(uuid.uuid4())
        txt_path = os.path.join(TMP_DIR, f"{fid}.txt")
        doc_path = os.path.join(TMP_DIR, f"{fid}.docx")

        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(text_out)

        doc = Document()
        doc.add_paragraph(text_out)
        doc.save(doc_path)

        return render_template_string(
            HTML,
            text=text_out,
            corrected=apply_spell,
            fid=fid
        )

    # GET
    return render_template_string(HTML, text=None)

# ── Download endpoints ───────────────────────────────────────────────────
@app.route("/download/<fid>.<ext>")
def download(fid, ext):
    if ext not in {"txt", "docx"}:
        flash("Invalid file type.")
        return redirect(url_for("index"))

    path = os.path.join(TMP_DIR, f"{fid}.{ext}")
    if not os.path.exists(path):
        flash("File expired or not found.")
        return redirect(url_for("index"))

    mime = "text/plain" if ext == "txt" else (
           "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    return send_file(path, as_attachment=True, download_name=f"ocr_result.{ext}", mimetype=mime)

# ── Clean up temp dir on exit (optional) ─────────────────────────────────
import atexit, shutil
@atexit.register
def _cleanup():
    shutil.rmtree(TMP_DIR, ignore_errors=True)

# ── Run ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)
