from flask import Flask, request, jsonify
import os
import fitz  # PyMuPDF

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def root():
    # Healthcheck
    if request.method == "GET":
        return "PDF Extractor is running", 200

    # POST = PDF extrahieren
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file provided"}), 400

    file = request.files["file"]
    filename = (file.filename or "").lower()

    if not filename.endswith(".pdf"):
        return jsonify({"ok": False, "error": "File must be a PDF"}), 400

    try:
        pdf_bytes = file.read()

        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text_parts = []
        for page in doc:
            t = page.get_text("text") or ""
            if t.strip():
                text_parts.append(t)
        text = "\n".join(text_parts)

        return jsonify({
            "ok": True,
            "text": text,
            "text_len": len(text)
        }), 200

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

