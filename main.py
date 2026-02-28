from flask import Flask, request, jsonify
import os
import io

# 1) Primary extractor: PyMuPDF
import fitz  # PyMuPDF

# 2) Secondary extractor: pdfplumber (pdfminer)
import pdfplumber


app = Flask(__name__)


@app.get("/")
def health():
    return "OK"


def extract_with_pymupdf(pdf_bytes: bytes) -> str:
    text_parts = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page in doc:
            # "text" is usually best general-purpose
            t = page.get_text("text") or ""
            text_parts.append(t)
    return "\n".join(text_parts).strip()


def extract_with_pdfplumber(pdf_bytes: bytes) -> str:
    text_parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            text_parts.append(t)
    return "\n".join(text_parts).strip()


@app.post("/extract")
def extract():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file part 'file' in multipart/form-data"}), 400

    f = request.files["file"]
    filename = (f.filename or "").lower()

    if not filename.endswith(".pdf"):
        return jsonify({"ok": False, "error": "File must be a PDF (.pdf)"}), 400

    pdf_bytes = f.read()
    if not pdf_bytes or len(pdf_bytes) < 100:
        return jsonify({"ok": False, "error": "Empty/invalid PDF upload"}), 400

    warnings = []
    text = ""
    method_used = None

    # Try PyMuPDF first
    try:
        text = extract_with_pymupdf(pdf_bytes)
        method_used = "pymupdf"
    except Exception as e:
        warnings.append(f"pymupdf_failed: {str(e)}")

    # If empty, try pdfplumber
    if not text:
        try:
            text = extract_with_pdfplumber(pdf_bytes)
            method_used = "pdfplumber"
        except Exception as e:
            warnings.append(f"pdfplumber_failed: {str(e)}")

    # If still empty => we return a clear diagnostic (no silent empties)
    if not text:
        return jsonify({
            "ok": False,
            "error": "NO_TEXT_FOUND",
            "method_used": method_used,
            "warnings": warnings,
            "bytes": len(pdf_bytes),
        }), 200  # 200 on purpose, so Make can continue and decide fallback

    return jsonify({
        "ok": True,
        "text": text,
        "method_used": method_used,
        "warnings": warnings,
        "bytes": len(pdf_bytes),
        "chars": len(text),
    }), 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
