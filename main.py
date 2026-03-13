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
    "blatt", "seite", "page",
    "bei schriftwechsel bitte angeben"
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
    "lieferschein nr", "lieferscheindatum",
    "bestellnr", "bestell-nr", "bestellnummer", "bestellangaben", "bestellung",
    "auftrag", "auftragsnr", "auftrags-nr", "auftragsnummer",
    "auftr.text", "auftr text", "auftragstext", "auftrags-text", "auftrags text",
    "auftr.nr", "auftr nr",
    "kommission", "projekt", "baustelle", "kostenstelle", "standort",
    "abholer", "ausweis-nr./abholer", "ausweis-nr", "selbstabholer", "abholung",
    "per lkw", "per paket", "versandart", "lieferanschrift", "lieferadresse",
    "lieferempfänger", "lieferempfaenger", "empfänger", "empfaenger",
    "ihr auftrag", "unser auftrag", "referenz", "referenznummer", "kundenzeichen",
    "vorgang entstand", "erfasst von", "online-warenkorbnr"
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

NOTICE_MARKERS = [
    "ab sofort möchten wir sie bitten",
    "ab sofort moechten wir sie bitten",
    "überweisungsankündigungen",
    "ueberweisungsankuendigungen",
    "zahlungsavis",
    "zentrale e-mail-adresse",
    "reibungslose zuordnung",
    "für ihre unterstützung bedanken wir uns",
    "fuer ihre unterstuetzung bedanken wir uns",
]

STAR_BLOCK_PATTERN = re.compile(r"^\*{5,}$")


def normalize_line(line: str) -> str:
    line = line.replace("\xa0", " ")
    line = re.sub(r"[ \t]+", " ", line)
    return line.strip()


def lower_clean(line: str) -> str:
    return normalize_line(line).lower()


def contains_any(line: str, markers: list[str]) -> bool:
    l = lower_clean(line)
    return any(marker in l for marker in markers)


def is_star_line(line: str) -> bool:
    return STAR_BLOCK_PATTERN.match(normalize_line(line)) is not None


def is_notice_line(line: str) -> bool:
    l = lower_clean(line)
    return contains_any(line, NOTICE_MARKERS) or "zahlungsavis" in l


def is_header_like(line: str) -> bool:
    l = lower_clean(line)
    if contains_any(line, HEADER_MARKERS):
        return True
    if "rechn.nr" in l or "rechn nr" in l:
        return True
    if "kd-nr" in l or "kunden-nr" in l:
        return True
    return False


def is_positions_header(line: str) -> bool:
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
        return True

    if contains_any(line, POSITIONS_HEADER_MARKERS):
        if any(x in l for x in ["artikel", "bezeichnung", "position", "pos."]):
            return True

    return False


def is_positions_like(line: str) -> bool:
    """
    Heuristik für echte Positionszeilen / Artikelzeilen.
    """
    raw = normalize_line(line)
    l = raw.lower()

    if not raw:
        return False

    # Sehr typische Artikelnummer-/Artikelkennungszeile
    if re.match(r"^[A-Z0-9][A-Z0-9\-/\.]{4,}$", raw):
        return True

    # Artikelcode am Anfang + irgendwo Zahl
    if re.match(r"^[A-Z0-9][A-Z0-9\-/\.]{3,}", raw) and re.search(r"\d", raw):
        return True

    # Mengen-/Preis-Zeilen
    if re.search(r"\d+,\d+\s*(m|stk|ein|pak|kg|l|qm|eim)\b", l):
        return True

    if re.search(r"\d+,\d+\s+\d+,\d+", raw):
        return True

    return False


def should_stop_order_block(line: str) -> bool:
    """
    Wann endet der order_delivery_block?
    """
    if is_positions_header(line):
        return True
    if contains_any(line, TOTALS_MARKERS):
        return True
    if contains_any(line, PAYMENT_MARKERS):
        return True
    if is_star_line(line):
        return True
    if is_notice_line(line):
        return True
    return False


def should_stop_positions_block(line: str) -> bool:
    """
    Wann endet der positions_block?
    """
    if contains_any(line, TOTALS_MARKERS):
        return True
    if contains_any(line, PAYMENT_MARKERS):
        return True
    if is_star_line(line):
        return True
    if is_notice_line(line):
        return True
    return False


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
        if is_positions_header(line):
            return i
    return None


def find_first_notice_or_star_index(lines: list[str], start_idx: int) -> int | None:
    for i in range(start_idx, len(lines)):
        if is_star_line(lines[i]) or is_notice_line(lines[i]):
            return i
    return None


def find_first_totals_or_payment_index(lines: list[str], start_idx: int) -> int | None:
    for i in range(start_idx, len(lines)):
        if contains_any(lines[i], TOTALS_MARKERS) or contains_any(lines[i], PAYMENT_MARKERS):
            return i
    return None


def find_footer_start_index(lines: list[str]) -> int | None:
    for i in range(len(lines) - 1, -1, -1):
        if contains_any(lines[i], FOOTER_MARKERS):
            return i
    return None


def collect_order_delivery_indices(lines: list[str], positions_header_idx: int | None) -> set[int]:
    """
    Sammelt explizit Liefer-/Auftrags-/Abholer-Block vor den Positionen.
    Wichtig für Hempelmann & ähnliche Rechnungen.
    """
    indices = set()
    if positions_header_idx is None:
        return indices

    upper_limit = len(lines)
    i = positions_header_idx + 1

    while i < upper_limit:
        line = lines[i]

        if contains_any(line, ORDER_DELIVERY_MARKERS):
            j = i
            while j < upper_limit:
                current = lines[j]

                if should_stop_order_block(current) and j != i:
                    break

                # stoppe, wenn echte Positionszeile beginnt und wir schon mind. 1 Folgezeile hatten
                if j != i and is_positions_like(current):
                    break

                indices.add(j)
                j += 1

            i = j
            continue

        # wenn nach Positionskopf direkt echte Positionszeile kommt, abbrechen
        if is_positions_like(line):
            break

        i += 1

    return indices


def split_blocks(lines: list[str]) -> dict:
    positions_header_idx = find_positions_header_index(lines)
    footer_idx = find_footer_start_index(lines)

    order_delivery_indices = collect_order_delivery_indices(lines, positions_header_idx)

    header_lines = []
    party_lines = []
    order_delivery_lines = []
    positions_header_block = ""
    positions_lines = []
    notice_lines = []
    totals_lines = []
    payment_lines = []
    footer_lines = []

    if positions_header_idx is None:
        # Fallback: kein Positionskopf gefunden -> fast alles oberer Bereich
        for line in lines:
            if contains_any(line, FOOTER_MARKERS):
                footer_lines.append(line)
            elif is_header_like(line):
                header_lines.append(line)
            else:
                party_lines.append(line)

        return {
            "header_block": "\n".join(header_lines).strip(),
            "party_block": "\n".join(party_lines).strip(),
            "order_delivery_block": "",
            "positions_header_block": "",
            "positions_block": "",
            "totals_block": "",
            "payment_block": "",
            "footer_block": "\n".join(footer_lines).strip(),
            "meta": {
                "positions_header_found": False,
                "totals_found": False,
                "payment_found": False,
                "footer_found": len(footer_lines) > 0
            }
        }

    positions_header_block = lines[positions_header_idx]

    # Oberer Bereich bis Positionskopf
    upper_lines = lines[:positions_header_idx]

    for idx, line in enumerate(upper_lines):
        if is_header_like(line):
            header_lines.append(line)
        elif contains_any(line, ORDER_DELIVERY_MARKERS):
            order_delivery_lines.append(line)
        else:
            party_lines.append(line)

    # Alles nach Positionskopf
    lower_lines = lines[positions_header_idx + 1:]

    # Step 1: notice/star block finden
    notice_start_rel = find_first_notice_or_star_index(lower_lines, 0)
    totals_start_rel = find_first_totals_or_payment_index(lower_lines, 0)

    positions_end_rel_candidates = [x for x in [notice_start_rel, totals_start_rel] if x is not None]
    positions_end_rel = min(positions_end_rel_candidates) if positions_end_rel_candidates else len(lower_lines)

    # Positionen bis vor notice/totals/payment
    for rel_idx in range(0, positions_end_rel):
        abs_idx = positions_header_idx + 1 + rel_idx
        line = lines[abs_idx]

        if abs_idx in order_delivery_indices:
            order_delivery_lines.append(line)
        else:
            positions_lines.append(line)

    # Notice Block
    notice_end_abs = None
    if notice_start_rel is not None:
        notice_abs = positions_header_idx + 1 + notice_start_rel
        k = notice_abs
        while k < len(lines):
            if contains_any(lines[k], TOTALS_MARKERS) or contains_any(lines[k], PAYMENT_MARKERS):
                break
            if contains_any(lines[k], FOOTER_MARKERS):
                break
            notice_lines.append(lines[k])
            k += 1
        notice_end_abs = k

    # Totals / Payment / Footer ab danach
    start_after_notice = notice_end_abs if notice_end_abs is not None else (positions_header_idx + 1 + positions_end_rel)

    tail_lines = lines[start_after_notice:]

    # Footer am Ende separat
    footer_start_in_tail = None
    for i, line in enumerate(tail_lines):
        if contains_any(line, FOOTER_MARKERS):
            footer_start_in_tail = i
            break

    tail_main = tail_lines[:footer_start_in_tail] if footer_start_in_tail is not None else tail_lines
    footer_lines = tail_lines[footer_start_in_tail:] if footer_start_in_tail is not None else []

    # Totals vs Payment trennen
    for line in tail_main:
        if contains_any(line, PAYMENT_MARKERS):
            payment_lines.append(line)
        elif contains_any(line, TOTALS_MARKERS):
            totals_lines.append(line)
        else:
            # unklare Restzeilen nach Positionen eher Payment zuordnen
            if line:
                payment_lines.append(line)

    # order_delivery zusätzlich aus expliziten Indizes nach Positionskopf ergänzen
    for idx in sorted(order_delivery_indices):
        if idx > positions_header_idx:
            line = lines[idx]
            if line not in order_delivery_lines:
                order_delivery_lines.append(line)

    # Aus positions_lines alles entfernen, was klar order_delivery / totals / payment / notice ist
    cleaned_positions = []
    for line in positions_lines:
        if contains_any(line, ORDER_DELIVERY_MARKERS):
            if line not in order_delivery_lines:
                order_delivery_lines.append(line)
            continue
        if should_stop_positions_block(line):
            if is_notice_line(line) or is_star_line(line):
                notice_lines.append(line)
            elif contains_any(line, TOTALS_MARKERS):
                totals_lines.append(line)
            elif contains_any(line, PAYMENT_MARKERS):
                payment_lines.append(line)
            continue
        cleaned_positions.append(line)
    positions_lines = cleaned_positions

    # notice_block an payment anhängen, damit GPT ihn separat von Positionen sieht
    if notice_lines:
        payment_lines = notice_lines + payment_lines

    # Dubletten bereinigen
    def dedupe_keep_order(items: list[str]) -> list[str]:
        seen = set()
        result = []
        for item in items:
            key = item.strip()
            if not key:
                continue
            if key not in seen:
                seen.add(key)
                result.append(item)
        return result

    header_lines = dedupe_keep_order(header_lines)
    party_lines = dedupe_keep_order(party_lines)
    order_delivery_lines = dedupe_keep_order(order_delivery_lines)
    positions_lines = dedupe_keep_order(positions_lines)
    totals_lines = dedupe_keep_order(totals_lines)
    payment_lines = dedupe_keep_order(payment_lines)
    footer_lines = dedupe_keep_order(footer_lines)

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
            "positions_header_found": True,
            "totals_found": len(totals_lines) > 0,
            "payment_found": len(payment_lines) > 0,
            "footer_found": len(footer_lines) > 0
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

        try:
            text_full, pages = extract_text_pymupdf(pdf_bytes)
            engine = "pymupdf"
        except Exception as e:
            pymupdf_error = str(e)

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
        block_result = split_blocks(lines)

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
