from flask import Flask, request, jsonify
import fitz  # PyMuPDF
import pdfplumber
import io
import os
import re

app = Flask(__name__)


HEADER_MARKERS = [
    "rechnung", "rechnungsnummer", "rechn.-nr", "rechnung nr", "rechnung-nr",
    "beleg-nr", "belegnr", "beleg nr", "belegnummer",
    "datum", "rechnungsdatum", "belegdatum",
    "kunden-nr", "kundennr", "kunden nr", "kd-nr", "kd nr", "kundennummer",
    "blatt", "seite", "page"
]

POSITIONS_HEADER_MARKERS = [
    "artikel", "artikelnr", "artikel-nr", "art.-nr", "artnr",
    "menge", "einheit", "einh", "me",
    "preis", "wert", "betrag", "einzelpreis", "e-preis", "g-preis",
    "position", "pos", "pos.", "bezeichnung", "beschreibung",
    "preisdimension", "pos/wert",
    "artikel / bezeichnung", "menge / einzelpreis",
    "pos. artikel menge me pe preis wert sk",
    "position menge einh. bezeichnung e-preis g-preis"
]

ORDER_DELIVERY_MARKERS = [
    "lieferung", "lieferdatum", "lieferschein", "lieferschein-nr", "lieferscheinnr",
    "bestellnr", "bestell-nr", "bestellnummer", "bestellangaben", "bestellung",
    "auftrag", "auftragsnr", "auftrags-nr", "auftragsnummer",
    "auftr.text", "auftr text", "auftragstext",
    "auftr.nr", "auftr nr",
    "kommission", "projekt", "baustelle", "kostenstelle", "standort",
    "abholer", "ausweis-nr./abholer", "ausweis-nr", "selbstabholer", "abholung",
    "per lkw", "per paket", "versandart", "lieferanschrift", "lieferadresse",
    "lieferempfänger", "lieferempfaenger", "empfänger", "empfaenger",
    "ihr auftrag", "unser auftrag", "referenz", "referenznummer", "kundenzeichen"
]

TOTALS_MARKERS = [
    "nettowarenwert", "netto-summe", "nettobetrag", "netto gesamt", "netto",
    "warenwert", "gesamt", "gesamtbetrag", "gesamt-betrag",
    "endbetrag", "rechnungsbetrag", "bruttobetrag", "zwischensumme",
    "mwst", "mwst.", "ust", "ust.", "umsatzsteuer", "steuerbetrag",
    "steuerpflichtiger betrag", "skontofähiger betrag", "skontofaehiger betrag",
    "versandkosten", "fracht", "zuschlag", "gebühren", "gebuehren"
]

PAYMENT_MARKERS = [
    "zahlungskonditionen", "zahlungskondition",
    "zahlungsbedingung", "zahlungsbedingungen",
    "zahlungsart", "zahlweise",
    "zahlbar bis", "zahlbar sofort", "ohne abzug", "unter abzug",
    "skonto", "skontobetrag", "skontosatz", "skontodatum",
    "lastschrift", "abbuchung", "wird abgebucht", "bankeinzug",
    "zahlungsziel", "fälligkeit", "faelligkeit", "fälligkeitsdatum", "faelligkeitsdatum"
]

FOOTER_MARKERS = [
    "agb", "allgemeine geschäftsbedingungen", "allgemeine geschaeftsbedingungen",
    "iban", "bic", "swift", "bankverbindung", "bank",
    "ust-id", "ust-idnr", "ust id", "ustidnr",
    "registergericht", "amtsgericht", "handelsregister", "geschäftsführer", "geschaeftsfuehrer",
    "es gelten unsere", "verkauf, lieferung und versand erfolgen",
    "impressum", "www.", "http://", "https://"
]

PARTY_MARKERS = [
    "debitor", "ansprechpartner", "ihr ansprechpartner", "sachbearbeiter", "bearbeiter",
    "innendienst", "innend.", "außendienst", "aussendienst", "außend.", "aussend.",
    "telefon", "telefon-nr", "tel.", "fax", "e-mail", "email", "mail",
    "kunde", "rechnungsempfänger", "rechnungsempfaenger", "rechnung an",
    "rechnungsadresse", "lieferempfänger", "lieferempfaenger", "lieferanschrift",
    "lieferadresse", "anschrift"
]


def normalize_line(line: str) -> str:
    line = line.replace("\xa0", " ")
    line = re.sub(r"[ \t]+", " ", line)
    return line.strip()


def lower_clean(line: str) -> str:
    return normalize_line(line).lower()


def contains_any(line: str, markers: list[str]) -> bool:
    l = lower_clean(line)
    return any(marker in l for marker in markers)


def extract_text_pymupdf(pdf_bytes: bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []
    for i, page in enumerate(doc):
        text = page.get_text("text") or ""
        pages.append({
            "page": i + 1,
            "text": text
        })
    full_text = "\n".join(p["text"] for p in pages)
    return full_text, pages


def extract_text_pdfplumber(pdf_bytes: bytes):
    pages = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            pages.append({
                "page": i + 1,
                "text": text
            })
    full_text = "\n".join(p["text"] for p in pages)
    return full_text, pages


def clean_lines(text: str) -> list[str]:
    raw_lines = text.splitlines()
    lines = []
    for line in raw_lines:
        n = normalize_line(line)
        if n:
            lines.append(n)
    return lines


def find_positions_header_index(lines: list[str]) -> int | None:
    for i, line in enumerate(lines):
        l = lower_clean(line)

        score = 0
        if "artikel" in l:
            score += 1
        if "menge" in l:
            score += 1
        if "preis" in l or "e-preis" in l or "g-preis" in l:
            score += 1
        if "wert" in l or "betrag" in l or "pos/wert" in l:
            score += 1
        if "me" in l or "einh" in l or "einheit" in l:
            score += 1

        if score >= 3:
            return i

        if contains_any(line, POSITIONS_HEADER_MARKERS):
            # fallback: einzelne starke Kopfzeilen
            if any(x in l for x in ["artikel", "bezeichnung", "position", "pos."]):
                return i

    return None


def find_last_marker_index(lines: list[str], markers: list[str]) -> int | None:
    for i in range(len(lines) - 1, -1, -1):
        if contains_any(lines[i], markers):
            return i
    return None


def find_footer_start_index(lines: list[str]) -> int | None:
    for i in range(len(lines) - 1, -1, -1):
        if contains_any(lines[i], FOOTER_MARKERS):
            return i
    return None


def collect_order_delivery_indices(lines: list[str], positions_header_idx: int | None) -> set[int]:
    indices = set()
    upper_limit = positions_header_idx if positions_header_idx is not None else len(lines)

    i = 0
    while i < upper_limit:
        l = lower_clean(lines[i])

        if contains_any(lines[i], ORDER_DELIVERY_MARKERS):
            indices.add(i)

            # Hempelmann-/Auftragsblock-Regel:
            # nach Lieferung / AUFTR / Abholer / per LKW auch die Folgezeilen mitnehmen,
            # bis klar eine Position beginnt oder ein Summenmarker auftaucht.
            j = i + 1
            while j < upper_limit:
                next_l = lower_clean(lines[j])

                if contains_any(lines[j], TOTALS_MARKERS) or contains_any(lines[j], PAYMENT_MARKERS):
                    break

                # neue starke Tabellenkopfzeile -> stoppen
                if contains_any(lines[j], POSITIONS_HEADER_MARKERS):
                    break

                # grobe Positionszeile: kurze Artikelkennung am Anfang + Zahlen weiter rechts
                if re.match(r"^[A-Z0-9][A-Z0-9\-\/\.]{3,}", lines[j]) and re.search(r"\d", lines[j]):
                    # diese Zeile gehört meist schon zu Positionen, also NICHT mehr reinziehen
                    break

                # typische Folgezeilen für Lieferung/Adresse/AUFTR.TEXT/Abholer
                indices.add(j)
                j += 1

            i = j
            continue

        i += 1

    return indices


def build_blocks(lines: list[str]) -> dict:
    positions_header_idx = find_positions_header_index(lines)
    totals_idx = find_last_marker_index(lines, TOTALS_MARKERS)
    payment_idx = find_last_marker_index(lines, PAYMENT_MARKERS)
    footer_idx = find_footer_start_index(lines)

    # Footer bevorzugt abtrennen, aber nicht mitten in totals/payment schneiden
    if footer_idx is not None:
        if totals_idx is not None and footer_idx < totals_idx:
            footer_idx = None
        if payment_idx is not None and footer_idx < payment_idx:
            footer_idx = None

    # Positionsheader
    positions_header_block = ""
    if positions_header_idx is not None:
        positions_header_block = lines[positions_header_idx]

    # Order/Delivery
    order_delivery_indices = collect_order_delivery_indices(lines, positions_header_idx)

    # Header-/Party-Bereich: alles vor positions_header
    upper_end = positions_header_idx if positions_header_idx is not None else len(lines)
    upper_lines = lines[:upper_end]

    header_lines = []
    party_lines = []
    order_delivery_lines = []

    for i, line in enumerate(upper_lines):
        if i in order_delivery_indices:
            order_delivery_lines.append(line)
        elif contains_any(line, HEADER_MARKERS):
            header_lines.append(line)
        else:
            party_lines.append(line)

    # Positionsblock
    positions_start = positions_header_idx + 1 if positions_header_idx is not None else None

    # erstestes Ende bestimmen: totals oder payment oder footer
    possible_ends = [idx for idx in [totals_idx, payment_idx, footer_idx] if idx is not None]
    positions_end = min(possible_ends) if possible_ends else len(lines)

    positions_lines = []
    if positions_start is not None and positions_start < positions_end:
        for idx in range(positions_start, positions_end):
            # nichts aus order_delivery in positionen übernehmen
            if idx not in order_delivery_indices:
                positions_lines.append(lines[idx])

    # Totalsblock
    totals_lines = []
    if totals_idx is not None:
        totals_end = min([idx for idx in [payment_idx, footer_idx] if idx is not None and idx > totals_idx] or [len(lines)])
        totals_lines = lines[totals_idx:totals_end]

    # Paymentblock
    payment_lines = []
    if payment_idx is not None:
        payment_end = min([idx for idx in [footer_idx] if idx is not None and idx > payment_idx] or [len(lines)])
        payment_lines = lines[payment_idx:payment_end]

    # Footerblock
    footer_lines = []
    if footer_idx is not None:
        footer_lines = lines[footer_idx:]

    # Doppelte Zeilen aus party/header entfernen, wenn sie schon in order_delivery sind
    order_delivery_set = set(order_delivery_lines)
    header_lines = [x for x in header_lines if x not in order_delivery_set]
    party_lines = [x for x in party_lines if x not in order_delivery_set]

    # Party etwas säubern: keine Footer-/Totals-/Payment-Zeilen
    party_lines = [
        x for x in party_lines
        if not contains_any(x, TOTALS_MARKERS)
        and not contains_any(x, PAYMENT_MARKERS)
        and not contains_any(x, FOOTER_MARKERS)
    ]

    return {
        "header_block": "\n".join(header_lines).strip(),
        "party_block": "\n".join(party_lines).strip(),
        "order_delivery_block": "\n".join(order_delivery_lines).strip(),
        "positions_header_block": positions_header_block.strip(),
        "positions_block": "\n".join(positions_lines).strip(),
        "totals_block": "\n".join(totals_lines).strip(),
        "payment_block": "\n".join(payment_lines).strip(),
        "footer_block": "\n".join(footer_lines).strip(),
        "meta": {
            "positions_header_found": positions_header_idx is not None,
            "totals_found": totals_idx is not None,
            "payment_found": payment_idx is not None,
            "footer_found": footer_idx is not None
        }
    }


@app.route("/", methods=["GET"])
def health():
    return "PDF Extractor is running", 200


@app.route("/extract", methods=["POST"])
def extract_pdf():
    if "file" not in request.files:
        return jsonify({
            "ok": False,
            "error": "No file provided"
        }), 400

    file = request.files["file"]

    if not file.filename or not file.filename.lower().endswith(".pdf"):
        return jsonify({
            "ok": False,
            "error": "File must be a PDF"
        }), 400

    try:
        pdf_bytes = file.read()

        pymupdf_error = None
        pdfplumber_error = None

        text_full = ""
        pages = []
        engine = "none"

        # Hauptversuch: PyMuPDF
        try:
            text_full, pages = extract_text_pymupdf(pdf_bytes)
            engine = "pymupdf"
        except Exception as e:
            pymupdf_error = str(e)

        # Fallback: pdfplumber, falls kein Text
        if not text_full.strip():
            try:
                text_full, pages = extract_text_pdfplumber(pdf_bytes)
                engine = "pdfplumber"
            except Exception as e:
                pdfplumber_error = str(e)

        if not text_full.strip():
            return jsonify({
                "ok": True,
                "engine": engine,
                "text_full": "",
                "pages": pages,
                "blocks": {
                    "header_block": "",
                    "party_block": "",
                    "order_delivery_block": "",
                    "positions_header_block": "",
                    "positions_block": "",
                    "totals_block": "",
                    "payment_block": "",
                    "footer_block": ""
                },
                "meta": {
                    "pages_count": len(pages),
                    "chars": 0,
                    "positions_header_found": False,
                    "totals_found": False,
                    "payment_found": False,
                    "footer_found": False,
                    "warnings": ["No extractable text found in PDF"]
                },
                "pymupdf_error": pymupdf_error,
                "pdfplumber_error": pdfplumber_error
            }), 200

        lines = clean_lines(text_full)
        block_result = build_blocks(lines)

        return jsonify({
            "ok": True,
            "engine": engine,
            "text_full": text_full,
            "pages": pages,
            "blocks": {
                "header_block": block_result["header_block"],
                "party_block": block_result["party_block"],
                "order_delivery_block": block_result["order_delivery_block"],
                "positions_header_block": block_result["positions_header_block"],
                "positions_block": block_result["positions_block"],
                "totals_block": block_result["totals_block"],
                "payment_block": block_result["payment_block"],
                "footer_block": block_result["footer_block"]
            },
            "meta": {
                "pages_count": len(pages),
                "chars": len(text_full),
                "positions_header_found": block_result["meta"]["positions_header_found"],
                "totals_found": block_result["meta"]["totals_found"],
                "payment_found": block_result["meta"]["payment_found"],
                "footer_found": block_result["meta"]["footer_found"],
                "warnings": []
            }
        }), 200

    except Exception as e:
        return jsonify({
            "ok": False,
            "error": str(e)
        }), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
