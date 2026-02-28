from flask import Flask, request, jsonify
import io

import fitz  # PyMuPDF
import pdfplumber

app = Flask(__name__)

@app.get("/")
def health():
    return "PDF Extractor is running"

def extract_with_pymupdf(pdf_bytes: bytes) -> str:
    text_all = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        t = page.get_text("text") or ""
        if t.strip():
            text_all.append(t)
    doc.close()
    return "\n".join(text_all).strip()

def extract_with_pdfplumber(pdf_bytes: bytes) -> str:
    text_all = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            if t.strip():
                text_all.append(t)
    return "\n".join(text_all).strip()

@app.post("/extract")
def extract_pdf():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file provided (form-data field name must be 'file')"}), 400

    f = request.files["file"]
    if not (f.filename or "").lower().endswith(".pdf"):
        return jsonify({"ok": False, "error": "File must be a PDF"}), 400

    pdf_bytes = f.read()
    if not pdf_bytes:
        return jsonify({"ok": False, "error": "Empty file bytes received"}), 400

    # 1) Try PyMuPDF
    try:
        text = extract_with_pymupdf(pdf_bytes)
        if text:
            return jsonify({"ok": True, "engine": "pymupdf", "text": text})
    except Exception as e:
        pymupdf_err = str(e)
    else:
        pymupdf_err = None

    # 2) Fallback pdfplumber
    try:
        text = extract_with_pdfplumber(pdf_bytes)
        if text:
            return jsonify({"ok": True, "engine": "pdfplumber", "text": text})
        return jsonify({
            "ok": True,
            "engine": "none",
            "text": "",
            "warning": "No extractable text found (could be image-based or weird encoding).",
            "pymupdf_error": pymupdf_err
        })
    except Exception as e:
        return jsonify({
            "ok": False,
            "error": str(e),
            "pymupdf_error": pymupdf_err
        }), 500
