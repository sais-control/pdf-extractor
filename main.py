from flask import Flask, request, jsonify
import pdfplumber
import io
import os

app = Flask(__name__)

@app.route("/", methods=["GET"])
def health():
    return "PDF Extractor is running"

@app.route("/extract", methods=["POST"])
def extract_pdf():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]

    if not file.filename.endswith(".pdf"):
        return jsonify({"error": "File must be a PDF"}), 400

    try:
        pdf_bytes = file.read()
        text = ""

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""

        return jsonify({
            "success": True,
            "text": text
        })

    except Exception as e:
        return jsonify({
            "error": str(e)
        }), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
import functions_framework
import pdfplumber
from flask import Request, jsonify
import io


@functions_framework.http
def hello_http(request: Request):
    try:
        if request.method != "POST":
            return jsonify({
                "ok": False,
                "error": "Only POST requests allowed"
            }), 405

        if "file" not in request.files:
            return jsonify({
                "ok": False,
                "error": "No file part in request"
            }), 400

        file = request.files["file"]

        if file.filename == "":
            return jsonify({
                "ok": False,
                "error": "No selected file"
            }), 400

        if not file.filename.lower().endswith(".pdf"):
            return jsonify({
                "ok": False,
                "error": "File must be a PDF"
            }), 400

        pdf_bytes = file.read()

        extracted_text = ""

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    extracted_text += text + "\n"

        return jsonify({
            "ok": True,
            "text": extracted_text
        })

    except Exception as e:
        return jsonify({
            "ok": False,
            "error": str(e)
        }), 500
