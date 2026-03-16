from flask import Flask, request, jsonify
import fitz  # PyMuPDF
import pdfplumber
import pytesseract
from PIL import Image
import io
import os
import re
from typing import Optional, List, Dict, Any

app = Flask(__name__)

# ============================================================
# MARKER / VARIANTEN
# ============================================================

DOCUMENT_TYPE_MARKERS = {
    "gutschrift": [
        "gutschrift",
        "leistungsgutschrift",
        "abrechnungsgutschrift",
        "rechnungskorrektur",
        "re-korrektur",
        "korrekturbeleg",
    ],
    "avis": [
        "lastschriftavis",
        "sepa-lastschriftavis",
        "zahlungsavis",
        "abbuchen",
        "abzugsbetrag",
        "zahlbetrag",
        "saldo",
    ],
    "lieferschein": [
        "lieferschein",
        "lieferscheinnr",
        "lieferschein-nr",
        "lieferscheinnummer",
        "entsorgungsnachweis",
    ],
    "rechnung": [
        "rechnung",
        "rechnungs-nr",
        "rechnungsnr",
        "rechnungsnummer",
        "invoice",
        "auftragsbestätigung/rechnung",
        "auftragsbestaetigung/rechnung",
    ],
}

HEADER_LABEL_VARIANTS = [
    "rechnung",
    "rechnungsnr",
    "rechnungs-nr",
    "rechnungsnummer",
    "rechn.nr",
    "rechnung nr",
    "beleg-nr",
    "belegnr",
    "beleg nr",
    "belegnummer",
    "re-korrekturnr",
    "korrekturnr",
    "gutschrift-nr",
    "gutschriftnr",
    "kundennr",
    "kunden-nr",
    "kunden nr",
    "kd-nr",
    "kd nr",
    "kundennummer",
    "datum",
    "beleg-datum",
    "belegdatum",
    "re-datum",
    "rechnungsdatum",
    "leistungsdatum",
    "lieferdatum",
    "seite",
    "blatt",
    "page",
    "seite 1 / 1",
    "bei schriftwechsel bitte angeben",
]

ORDER_LABEL_VARIANTS = [
    "auftrag",
    "auftragsnr",
    "auftrags-nr",
    "auftragsnummer",
    "auftrags-nr.",
    "unser auftrag",
    "ihr auftrag",
    "auftr.nr",
    "auftr nr",
    "auftr.text",
    "auftr text",
    "auftragstext",
    "auftragstext:",
    "bestell-nr",
    "bestellnr",
    "bestellnummer",
    "bestellangaben",
    "besteller",
    "kd-bestell-nr",
    "kd-besteller",
    "kd-bestell-datum",
    "lieferung",
    "lieferschein",
    "lieferschein-nr",
    "lieferscheinnr",
    "lieferscheinnummer",
    "liefersch.-nr",
    "lieferschein-nr.",
    "liefersch.-datum",
    "lieferdatum",
    "lieferanschrift",
    "lieferadresse",
    "lieferempfänger",
    "lieferempfaenger",
    "versandart",
    "versandbedingung",
    "versandanschrift",
    "lieferbedingung",
    "per lkw",
    "per paket",
    "abholung",
    "selbstabholer",
    "abholer",
    "ausweis-nr./abholer",
    "ausweis-nr",
    "kommission",
    "projekt",
    "baustelle",
    "kostenstelle",
    "standort",
    "fremdreferenz",
    "regulierer",
    "bearbeiter webshop",
    "vorgang entstand",
    "erfasst von",
]

TOTAL_LABEL_VARIANTS = [
    "summe positionen",
    "zwischensumme",
    "zwischensumme position",
    "zwischensumme vor steuer",
    "zwischensumme (netto)",
    "nettowarenwert",
    "netto-summe",
    "nettobetrag",
    "netto-betrag",
    "netto",
    "warenwert",
    "steuerpflichtiger betrag",
    "mwst",
    "mwst.",
    "mwst-betrag",
    "umsatzsteuer",
    "mehrwertsteuer",
    "ust",
    "ust.",
    "endbetrag",
    "gesamt",
    "gesamtbetrag",
    "gesamt-betrag",
    "brutto",
    "bruttobetrag",
    "summe",
]

PAYMENT_LABEL_VARIANTS = [
    "zahlungskonditionen",
    "zahlungskonditionen",
    "zahlungsbedingungen",
    "zahlungsbedingung",
    "zahlungsbedingung 14 tage",
    "zahlungsart",
    "zahlweise",
    "zahlbar bis",
    "zahlbar ohne abzug",
    "ohne abzug",
    "unter abzug",
    "skonto",
    "skontobetrag",
    "skontodatum",
    "skontosatz",
    "skontofähiger betrag",
    "skontofaehiger betrag",
    "fälligkeit",
    "faelligkeit",
    "fälligkeitsdatum",
    "faelligkeitsdatum",
    "abbuchen",
    "abgebucht",
    "lastschrift",
    "sepa-lastschrift",
    "sepa-lastschriftavis",
    "zahlungsavis",
    "netto / netto-tage",
    "skonto / skonto-tage",
    "valuta / valuta-tage",
]

FOOTER_LABEL_VARIANTS = [
    "iban",
    "bic",
    "swift",
    "bank",
    "bankverbindung",
    "ust-id",
    "ust-idnr",
    "ust-idnr.",
    "ust id",
    "ust.-ident-nr",
    "steuernummer",
    "registergericht",
    "amtsgericht",
    "geschäftsführung",
    "geschaeftsfuehrung",
    "geschäftsführer",
    "geschaeftsfuehrer",
    "www.",
    "agb",
    "allgemeine geschäftsbedingungen",
    "allgemeine geschaeftsbedingungen",
    "verkauf, lieferung und versand erfolgen",
]

PARTY_HINTS = [
    "firma",
    "kunde",
    "käufer",
    "kaeufer",
    "rechnungsempf",
    "rechnung an",
    "lieferadresse",
    "lieferanschrift",
    "lieferempfänger",
    "lieferempfaenger",
    "debitor",
    "ansprechpartner",
    "sachbearbeiter",
    "bearbeiter",
    "innend.",
    "innendienst",
    "außend.",
    "aussend.",
    "außendienst",
    "aussendienst",
    "telefon",
    "tel.",
    "fax",
    "e-mail",
    "email",
]

POSITIONS_HEADER_HINTS = [
    "pos",
    "position",
    "artikel",
    "artikelnr",
    "artikel-nr",
    "art.-nr",
    "bezeichnung",
    "menge",
    "einheit",
    "einh",
    "me",
    "ep",
    "pe",
    "preis",
    "einzelpreis",
    "e-preis",
    "betrag",
    "wert",
    "gesamt",
    "gp",
    "preisdimension",
]

NOTICE_MARKERS = [
    "zahlungsavis",
    "überweisungsankündigungen",
    "ueberweisungsankuendigungen",
    "zentrale e-mail-adresse",
    "für ihre unterstützung bedanken wir uns",
    "fuer ihre unterstuetzung bedanken wir uns",
]

STAR_LINE_RE = re.compile(r"^\*{5,}$")
AMOUNT_RE = re.compile(r"[-]?\d{1,3}(?:\.\d{3})*,\d{2}")
DATE_RE = re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b")
PERCENT_RE = re.compile(r"\d{1,3},\d{1,2}\s?%")
TOKEN_SPLIT_RE = re.compile(r"\s+")

# ============================================================
# HILFSFUNKTIONEN
# ============================================================

def norm(text: str) -> str:
    text = text.replace("\xa0", " ").replace("\ufeff", " ")
    text = text.replace("￾", "")  # kaputte OCR-Trennzeichen
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()

def lower(text: str) -> str:
    return norm(text).lower()

def contains_any(text: str, variants: List[str]) -> bool:
    lt = lower(text)
    return any(v in lt for v in variants)

def is_star_line(text: str) -> bool:
    return STAR_LINE_RE.match(norm(text)) is not None

def is_empty_or_decorative(text: str) -> bool:
    t = norm(text)
    if not t:
        return True
    if re.fullmatch(r"[_\-=\*]{4,}", t):
        return True
    return False

def split_lines(text: str) -> List[str]:
    return [norm(x) for x in text.splitlines() if not is_empty_or_decorative(x)]

def looks_like_positions_header(line: str) -> bool:
    lt = lower(line)
    score = 0
    for hint in POSITIONS_HEADER_HINTS:
        if hint in lt:
            score += 1
    return score >= 3

def looks_like_total_line(line: str) -> bool:
    lt = lower(line)
    return contains_any(lt, TOTAL_LABEL_VARIANTS) and bool(AMOUNT_RE.search(line) or "%" in line)

def looks_like_payment_line(line: str) -> bool:
    lt = lower(line)
    return contains_any(lt, PAYMENT_LABEL_VARIANTS)

def looks_like_footer_line(line: str) -> bool:
    lt = lower(line)
    return contains_any(lt, FOOTER_LABEL_VARIANTS)

def looks_like_order_line(line: str) -> bool:
    lt = lower(line)
    return contains_any(lt, ORDER_LABEL_VARIANTS)

def looks_like_header_line(line: str) -> bool:
    lt = lower(line)
    return contains_any(lt, HEADER_LABEL_VARIANTS)

def looks_like_party_line(line: str) -> bool:
    lt = lower(line)
    return contains_any(lt, PARTY_HINTS)

def is_amount_token(token: str) -> bool:
    return bool(AMOUNT_RE.fullmatch(token.strip()))

def count_amounts(text: str) -> int:
    return len(AMOUNT_RE.findall(text))

def tokenize_line(line: str) -> List[str]:
    parts = [p.strip() for p in TOKEN_SPLIT_RE.split(norm(line)) if p.strip()]
    return parts

def extract_text_pymupdf(pdf_bytes: bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []
    for i, page in enumerate(doc):
        text = page.get_text("text") or ""
        pages.append({"page": i + 1, "text": text})
    return "\n".join(p["text"] for p in pages), pages

def extract_text_pdfplumber(pdf_bytes: bytes):
    pages = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            pages.append({"page": i + 1, "text": text})
    return "\n".join(p["text"] for p in pages), pages

def extract_text_ocr(pdf_bytes: bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        text = pytesseract.image_to_string(img, lang="deu") or ""
        pages.append({"page": i + 1, "text": text})
    return "\n".join(p["text"] for p in pages), pages

# ============================================================
# DOKUMENTTYP
# ============================================================

def detect_document_type(text_full: str) -> str:
    lt = lower(text_full)
    if contains_any(lt, DOCUMENT_TYPE_MARKERS["gutschrift"]):
        return "gutschrift"
    if contains_any(lt, DOCUMENT_TYPE_MARKERS["avis"]):
        return "avis"
    if contains_any(lt, DOCUMENT_TYPE_MARKERS["lieferschein"]):
        return "lieferschein"
    if contains_any(lt, DOCUMENT_TYPE_MARKERS["rechnung"]):
        return "rechnung"
    return "sonstiges"

# ============================================================
# PAIR EXTRACTION
# ============================================================

def extract_inline_pairs(lines: List[str], variants: List[str]) -> List[Dict[str, str]]:
    pairs = []
    for line in lines:
        raw = norm(line)
        l = lower(raw)

        for v in variants:
            # label: value
            m = re.search(rf"(?i)\b({re.escape(v)})\b\s*[:\-]?\s*(.+)$", raw)
            if m:
                value = norm(m.group(2))
                if value and value.lower() != m.group(1).lower():
                    pairs.append({"label": norm(m.group(1)), "value": value, "source": "inline"})
                    break
    return dedupe_pairs(pairs)

def dedupe_pairs(pairs: List[Dict[str, str]]) -> List[Dict[str, str]]:
    seen = set()
    out = []
    for p in pairs:
        key = (lower(p["label"]), lower(p["value"]))
        if key not in seen:
            seen.add(key)
            out.append(p)
    return out

def split_header_label_row(line: str) -> List[str]:
    # versucht nebeneinanderstehende Kopf-Labels zu zerlegen
    cleaned = norm(line)
    cleaned = cleaned.replace("Rechn.Nr.", "Rechn.Nr")
    cleaned = cleaned.replace("KD-Nr.", "KD-Nr")
    labels = []

    known_labels = [
        "KD-Nr",
        "Rechn.Nr",
        "Rechnungsnummer",
        "Beleg-Nr",
        "Kundennr",
        "Kunden-Nr",
        "Datum",
        "Belegdatum",
        "Leistungsdatum",
        "Blatt",
        "Seite",
        "Rechnungsnr",
        "Rechnungs-Nr",
    ]

    temp = cleaned
    for kl in known_labels:
        temp = re.sub(rf"(?i)\b{re.escape(kl)}\b", f"|||{kl}|||", temp)

    chunks = [norm(x) for x in temp.split("|||") if norm(x)]
    for c in chunks:
        if any(lower(c) == lower(k) for k in known_labels):
            labels.append(c)

    return labels

def split_value_row_for_header(line: str) -> List[str]:
    tokens = tokenize_line(line)
    return tokens

def extract_vertical_header_pairs(lines: List[str]) -> List[Dict[str, str]]:
    pairs = []

    for i in range(len(lines) - 1):
        current = lines[i]
        nxt = lines[i + 1]

        labels = split_header_label_row(current)
        values = split_value_row_for_header(nxt)

        if not labels:
            continue
        if not values:
            continue

        # Heuristik für klassische Spaltenköpfe:
        # Wenn labels <= values, mappe der Reihe nach.
        # Wenn labels > values, nimm die letzten passenden Werte nicht blind,
        # sondern nur so viele wie vorhanden.
        # GPT bekommt später trotzdem blocks + full text.
        if len(labels) <= len(values):
            for idx, label in enumerate(labels):
                value = values[idx]
                pairs.append({"label": label, "value": value, "source": "vertical"})
        else:
            for idx, label in enumerate(labels[:len(values)]):
                value = values[idx]
                pairs.append({"label": label, "value": value, "source": "vertical_partial"})

    return dedupe_pairs(pairs)

def extract_same_line_table_pairs(lines: List[str], variants: List[str]) -> List[Dict[str, str]]:
    """
    Für Zeilen wie:
    Rechnungsnummer Datum
    7090918535 15.03.2026
    oder
    Unser Auftrag 105-VA... / Ihr Auftrag: ...
    """
    pairs = []

    for i in range(len(lines) - 1):
        current = norm(lines[i])
        nxt = norm(lines[i + 1])

        # Muster: Zeile mit mehreren Labels
        if sum(1 for v in variants if v in lower(current)) >= 2 and len(tokenize_line(nxt)) >= 2:
            labels = []
            for v in variants:
                if v in lower(current):
                    labels.append(v)

            values = tokenize_line(nxt)
            for idx, label in enumerate(labels[:len(values)]):
                pairs.append({"label": label, "value": values[idx], "source": "same_line_table"})

    return dedupe_pairs(pairs)

# ============================================================
# BLOCKS
# ============================================================

def find_positions_header_index(lines: List[str]) -> Optional[int]:
    for i, line in enumerate(lines):
        if looks_like_positions_header(line):
            return i
    return None

def build_blocks(lines: List[str]) -> Dict[str, str]:
    header_block = []
    party_block = []
    order_block = []
    positions_header_block = ""
    positions_block = []
    totals_block = []
    payment_block = []
    footer_block = []

    pos_idx = find_positions_header_index(lines)

    # Footer grob vom Ende aus
    footer_start = None
    for i in range(len(lines) - 1, -1, -1):
        if looks_like_footer_line(lines[i]):
            footer_start = i
            break

    effective_lines = lines[:footer_start] if footer_start is not None else lines
    footer_lines = lines[footer_start:] if footer_start is not None else []

    if pos_idx is None:
        # Fallback ohne Positionskopf
        for line in effective_lines:
            if looks_like_header_line(line):
                header_block.append(line)
            elif looks_like_order_line(line):
                order_block.append(line)
            elif looks_like_total_line(line):
                totals_block.append(line)
            elif looks_like_payment_line(line):
                payment_block.append(line)
            else:
                party_block.append(line)

        footer_block = footer_lines
        return {
            "header_block": "\n".join(header_block).strip(),
            "party_block": "\n".join(party_block).strip(),
            "order_delivery_block": "\n".join(order_block).strip(),
            "positions_header_block": "",
            "positions_block": "",
            "totals_block": "\n".join(totals_block).strip(),
            "payment_block": "\n".join(payment_block).strip(),
            "footer_block": "\n".join(footer_block).strip(),
        }

    positions_header_block = effective_lines[pos_idx]

    upper = effective_lines[:pos_idx]
    lower_part = effective_lines[pos_idx + 1:]

    for line in upper:
        if looks_like_header_line(line):
            header_block.append(line)
        elif looks_like_order_line(line):
            order_block.append(line)
        else:
            party_block.append(line)

    # order area / positions / totals / payment trennen
    mode = "positions"

    for line in lower_part:
        if is_star_line(line) or contains_any(line, NOTICE_MARKERS):
            mode = "payment"
            payment_block.append(line)
            continue

        if looks_like_payment_line(line) and mode != "positions":
            mode = "payment"
            payment_block.append(line)
            continue

        if looks_like_total_line(line):
            mode = "totals"
            totals_block.append(line)
            continue

        if looks_like_order_line(line) and mode == "positions" and len(positions_block) == 0:
            order_block.append(line)
            continue

        if mode == "positions":
            positions_block.append(line)
        elif mode == "totals":
            if looks_like_payment_line(line):
                payment_block.append(line)
            else:
                totals_block.append(line)
        else:
            payment_block.append(line)

    footer_block = footer_lines

    return {
        "header_block": "\n".join(header_block).strip(),
        "party_block": "\n".join(party_block).strip(),
        "order_delivery_block": "\n".join(order_block).strip(),
        "positions_header_block": positions_header_block.strip(),
        "positions_block": "\n".join(positions_block).strip(),
        "totals_block": "\n".join(totals_block).strip(),
        "payment_block": "\n".join(payment_block).strip(),
        "footer_block": "\n".join(footer_block).strip(),
    }

# ============================================================
# PARTY / ORDER DATA
# ============================================================

def extract_order_pairs(order_block_text: str) -> List[Dict[str, str]]:
    lines = split_lines(order_block_text)
    pairs = []

    # inline
    pairs.extend(extract_inline_pairs(lines, ORDER_LABEL_VARIANTS))

    # spezielle Logik für Zeilen wie:
    # "Lieferung 800 27639123-001 vom 28.01.2026 - per LKW -"
    for line in lines:
        if lower(line).startswith("lieferung"):
            value = norm(re.sub(r"(?i)^lieferung\s*", "", line))
            if value:
                pairs.append({"label": "Lieferung", "value": value, "source": "delivery_line"})

    return dedupe_pairs(pairs)

def extract_totals_pairs(totals_text: str) -> List[Dict[str, str]]:
    lines = split_lines(totals_text)
    pairs = []

    for line in lines:
        raw = norm(line)

        # Muster label : value
        m = re.match(r"^(.*?)(?:\s*[:]\s*|\s{2,})([-]?\d{1,3}(?:\.\d{3})*,\d{2}(?:\s*EUR)?)$", raw, re.IGNORECASE)
        if m:
            label = norm(m.group(1))
            value = norm(m.group(2))
            if label and value:
                pairs.append({"label": label, "value": value, "source": "totals"})
                continue

        # Muster "Mehrwertsteuer 19,00 % aus 613,92 116,64"
        amounts = AMOUNT_RE.findall(raw)
        if contains_any(raw, TOTAL_LABEL_VARIANTS) and amounts:
            pairs.append({"label": raw.replace(amounts[-1], "").strip(" :-"), "value": amounts[-1], "source": "totals_fallback"})

    return dedupe_pairs(pairs)

def extract_payment_pairs(payment_text: str) -> List[Dict[str, str]]:
    lines = split_lines(payment_text)
    pairs = []

    pairs.extend(extract_inline_pairs(lines, PAYMENT_LABEL_VARIANTS))

    for line in lines:
        raw = norm(line)
        if "zahlbar bis" in lower(raw):
            pairs.append({"label": "Zahlbar bis", "value": raw, "source": "payment_line"})
        elif "skonto" in lower(raw):
            pairs.append({"label": "Skonto", "value": raw, "source": "payment_line"})
        elif "lastschrift" in lower(raw):
            pairs.append({"label": "Lastschrift", "value": raw, "source": "payment_line"})

    return dedupe_pairs(pairs)

# ============================================================
# TABLE ROWS
# ============================================================

def is_probable_position_start(line: str) -> bool:
    t = tokenize_line(line)
    if not t:
        return False

    # typische Starts:
    # "1 790920001 ..."
    # "10 62 072 05 40 ..."
    # "Pos. 10 10 ST ..."
    # "YBT7630547 1000 1,000 ST ..."
    if lower(line).startswith("pos. "):
        return True

    if re.match(r"^\d+\s+\d", norm(line)):
        return True

    if re.match(r"^[A-Z0-9\-]{5,}\s+\d+", norm(line)):
        return True

    return False

def build_table_rows(positions_text: str) -> List[Dict[str, Any]]:
    lines = split_lines(positions_text)
    rows = []
    current = []

    for line in lines:
        l = lower(line)

        if "übertrag" in l or "uebertrag" in l:
            continue
        if looks_like_total_line(line) or looks_like_payment_line(line):
            continue
        if is_star_line(line):
            continue

        if is_probable_position_start(line):
            if current:
                rows.append(current)
            current = [line]
        else:
            if current:
                current.append(line)
            else:
                # falls Position ohne klaren Start kommt
                current = [line]

    if current:
        rows.append(current)

    out = []
    for idx, row_lines in enumerate(rows, start=1):
        flat = " ".join(row_lines)
        tokens = tokenize_line(flat)
        out.append({
            "row_index": idx,
            "raw_lines": row_lines,
            "raw_text": flat,
            "tokens": tokens,
        })

    return out

# ============================================================
# QUALITY
# ============================================================

def compute_quality(document_type: str,
                    blocks: Dict[str, str],
                    header_pairs: List[Dict[str, str]],
                    order_pairs: List[Dict[str, str]],
                    totals_pairs: List[Dict[str, str]],
                    payment_pairs: List[Dict[str, str]],
                    table_rows: List[Dict[str, Any]],
                    text_full: str,
                    ocr_used: bool) -> Dict[str, Any]:

    score = 0.0
    reasons = []

    if len(text_full.strip()) > 100:
        score += 0.10
    else:
        reasons.append("TEXT_SHORT")

    if blocks["header_block"]:
        score += 0.10
    else:
        reasons.append("HEADER_BLOCK_EMPTY")

    if header_pairs:
        score += 0.15
    else:
        reasons.append("HEADER_PAIRS_EMPTY")

    if document_type in ("rechnung", "gutschrift", "lieferschein"):
        if blocks["positions_header_block"]:
            score += 0.10
        else:
            reasons.append("POSITIONS_HEADER_EMPTY")

    if document_type in ("rechnung", "gutschrift", "lieferschein"):
        if table_rows:
            score += 0.20
        else:
            reasons.append("TABLE_ROWS_EMPTY")

    if document_type in ("rechnung", "gutschrift"):
        if totals_pairs:
            score += 0.15
        else:
            reasons.append("TOTALS_EMPTY")

    if blocks["order_delivery_block"] or order_pairs:
        score += 0.10

    if document_type in ("rechnung", "gutschrift", "avis"):
        if payment_pairs or blocks["payment_block"]:
            score += 0.10

    if ocr_used:
        score -= 0.05  # OCR ist okay, aber etwas unsicherer

    score = max(0.0, min(round(score, 2), 1.0))

    usable = score >= 0.50

    return {
        "score": score,
        "usable": usable,
        "reasons": reasons,
    }

def text_looks_bad(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return True
    if len(stripped) < 120:
        return True
    # zu wenige Zahlen + kaum Struktur
    if count_amounts(stripped) == 0 and len(split_lines(stripped)) < 8:
        return True
    return False

# ============================================================
# PIPELINE
# ============================================================

def build_output(text_full: str, pages: List[Dict[str, Any]], engine: str, ocr_used: bool,
                 pymupdf_error=None, pdfplumber_error=None, ocr_error=None):

    lines = split_lines(text_full)
    document_type = detect_document_type(text_full)
    blocks = build_blocks(lines)

    header_lines = split_lines(blocks["header_block"] + "\n" + blocks["party_block"])
    header_pairs = []
    header_pairs.extend(extract_inline_pairs(header_lines, HEADER_LABEL_VARIANTS))
    header_pairs.extend(extract_vertical_header_pairs(header_lines))
    header_pairs.extend(extract_same_line_table_pairs(header_lines, HEADER_LABEL_VARIANTS))
    header_pairs = dedupe_pairs(header_pairs)

    order_pairs = extract_order_pairs(blocks["order_delivery_block"])
    totals_pairs = extract_totals_pairs(blocks["totals_block"])
    payment_pairs = extract_payment_pairs(blocks["payment_block"])
    table_rows = build_table_rows(blocks["positions_block"])

    quality = compute_quality(
        document_type=document_type,
        blocks=blocks,
        header_pairs=header_pairs,
        order_pairs=order_pairs,
        totals_pairs=totals_pairs,
        payment_pairs=payment_pairs,
        table_rows=table_rows,
        text_full=text_full,
        ocr_used=ocr_used,
    )

    return {
        "ok": True,
        "meta": {
            "extractor": "cloudrun-v2",
            "text_engine": engine,
            "ocr_used": ocr_used,
            "page_count": len(pages),
            "chars": len(text_full),
        },
        "document": {
            "type": document_type,
            "language": "de",
        },
        "quality": quality,
        "header_pairs": header_pairs,
        "order_pairs": order_pairs,
        "totals_pairs": totals_pairs,
        "payment_pairs": payment_pairs,
        "party": {
            "party_block": blocks["party_block"],
        },
        "blocks": blocks,
        "table_rows": table_rows,
        "text_full": text_full,
        "pages": pages,
        "debug": {
            "pymupdf_error": pymupdf_error,
            "pdfplumber_error": pdfplumber_error,
            "ocr_error": ocr_error,
        }
    }

# ============================================================
# ROUTES
# ============================================================

@app.route("/", methods=["GET"])
def health():
    return "PDF Extractor v2 is running", 200

@app.route("/extract", methods=["POST"])
def extract_pdf():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file provided"}), 400

    file = request.files["file"]

    if not file.filename or not file.filename.lower().endswith(".pdf"):
        return jsonify({"ok": False, "error": "File must be a PDF"}), 400

    pdf_bytes = file.read()

    text_full = ""
    pages = []
    engine = "none"
    ocr_used = False

    pymupdf_error = None
    pdfplumber_error = None
    ocr_error = None

    # Stufe 1: PyMuPDF
    try:
        text_full, pages = extract_text_pymupdf(pdf_bytes)
        engine = "pymupdf"
    except Exception as e:
        pymupdf_error = str(e)

    # Stufe 2: pdfplumber
    if text_looks_bad(text_full):
        try:
            text_full_pp, pages_pp = extract_text_pdfplumber(pdf_bytes)
            if len(text_full_pp.strip()) > len(text_full.strip()):
                text_full, pages = text_full_pp, pages_pp
                engine = "pdfplumber"
        except Exception as e:
            pdfplumber_error = str(e)

    # Stufe 3: OCR
    if text_looks_bad(text_full):
        try:
            text_full_ocr, pages_ocr = extract_text_ocr(pdf_bytes)
            if len(text_full_ocr.strip()) > len(text_full.strip()):
                text_full, pages = text_full_ocr, pages_ocr
                engine = "ocr_tesseract"
                ocr_used = True
        except Exception as e:
            ocr_error = str(e)

    if not text_full.strip():
        return jsonify({
            "ok": False,
            "error": "No usable text extracted",
            "debug": {
                "pymupdf_error": pymupdf_error,
                "pdfplumber_error": pdfplumber_error,
                "ocr_error": ocr_error,
            }
        }), 200

    result = build_output(
        text_full=text_full,
        pages=pages,
        engine=engine,
        ocr_used=ocr_used,
        pymupdf_error=pymupdf_error,
        pdfplumber_error=pdfplumber_error,
        ocr_error=ocr_error,
    )

    return jsonify(result), 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
