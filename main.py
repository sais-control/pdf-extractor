from flask import Flask, request, jsonify
import os
import io

import fitz  # PyMuPDF

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def root():
    # Healthcheck
    if request.method == "GET":
        return "PDF Extractor is running", 200

    # POST: Datei muss als multipart/form-data mit Feldname "file" kommen
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file part in request"}), 400

    f = request.files["file"]
    filename = (f.filename or "").lower()

    if not filename.endswith(".pdf"):
        return jsonify({"ok": False, "error": "File must be a PDF"}), 400

    try:
        pdf_bytes = f.read()
        text = []

        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page in doc:
            text.append(page.get_text("text"))

        extracted = "\n".join(text).strip()

        return jsonify({
            "ok": True,
            "text": extracted,
            "text_len": len(extracted)
        }), 200

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
