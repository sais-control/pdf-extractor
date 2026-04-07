from flask import Flask, request, jsonify
import fitz # PyMuPDF
import pdfplumber
import pytesseract
from PIL import Image, ImageOps, ImageFilter
import io
import os
import re
from typing import List, Dict, Any, Optional, Tuple
from collections import Counter, defaultdict
from datetime import datetime, date, timezone, timedelta
from difflib import SequenceMatcher
import math
import unicodedata
app = Flask(__name__)

# ============================================================
# NORMALISIERUNG
# ============================================================

def norm(text: str) -> str:
    text = (text or "")
    text = text.replace("\xa0", " ").replace("\ufeff", " ").replace("￾", "")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()

def lower(text: str) -> str:
    return norm(text).lower()

def split_lines(text: str) -> List[str]:
    out = []
    for line in (text or "").splitlines():
        n = norm(line)
        if n:
            out.append(n)
    return out

def join_lines(lines: List[str]) -> str:
    return "\n".join([norm(x) for x in lines if norm(x)]).strip()

def tokenize(line: str) -> List[str]:
    return [x for x in re.split(r"\s+", norm(line)) if x]

# ============================================================
# MARKER
# ============================================================

DOC_TYPE_MARKERS = {
    "rechnung": [
        "rechnung",
        "invoice",
        "auftragsbestätigung/rechnung",
        "auftragsbestaetigung/rechnung",
    ],
    "gutschrift": [
        "gutschrift",
        "leistungsgutschrift",
        "abrechnungsgutschrift",
        "re-korrekturnr",
        "rechnungskorrektur",
    ],
    "lieferschein": [
        "lieferschein",
        "lieferscheinnr",
        "lieferscheinnummer",
        "entsorgungsnachweis",
    ],
    "avis": [
        "sepa-lastschriftavis",
        "sepa lastschriftavis",
        "lastschriftavis",
        "zahlungsavis",
        "saldo",
        "abzugsbetrag",
        "zahlbetrag",
        "nachstehende posten",
        "werden wir am",
    ],
}

HEADER_HINTS = [
    "rechnung",
    "rechnungsnr",
    "rechnungsnummer",
    "rechnungs-nr",
    "rechn.nr",
    "rechnung nr",
    "beleg-nr",
    "beleg nr",
    "belegnummer",
    "re-korrekturnr",
    "kundennr",
    "kunden-nr",
    "kunden nr",
    "kundennummer",
    "kd-nr",
    "kd nr",
    "datum",
    "belegdatum",
    "beleg-datum",
    "leistungsdatum",
    "blatt",
    "seite",
    "page",
    "bei schriftwechsel bitte angeben",
    "bitte bei zahlungen unbedingt angeben",
]

ADDRESS_HINTS = [
    "firma",
    "kunde",
    "käufer",
    "kaeufer",
    "rechnungsempf",
    "rechnung an",
    "lieferempfänger",
    "lieferempfaenger",
    "lieferanschrift",
    "lieferadresse",
    "ansprechpartner",
    "sachbearbeiter",
    "bearbeiter",
    "telefon",
    "tel.",
    "fax",
    "e-mail",
    "email",
    "debitor",
    "innend.",
    "innendienst",
    "außend.",
    "aussend.",
    "außendienst",
    "aussendienst",
]

ORDER_HINTS = [
    "auftrag",
    "auftragsnr",
    "auftragsnummer",
    "auftrags-nr",
    "unser auftrag",
    "ihr auftrag",
    "auftr.nr",
    "auftr nr",
    "auftr.text",
    "auftr text",
    "auftragstext",
    "bestell-nr",
    "bestellnr",
    "bestellnummer",
    "bestellangaben",
    "bestellung",
    "kd-bestell-nr",
    "kd-besteller",
    "lieferung",
    "lieferschein",
    "lieferschein-nr",
    "lieferscheinnr",
    "lieferscheinnummer",
    "liefersch.-nr",
    "liefersch.-datum",
    "lieferdatum",
    "lieferanschrift",
    "lieferadresse",
    "lieferempfänger",
    "lieferempfaenger",
    "versandart",
    "versandbedingung",
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
    "lieferadresse:",
    "adresse:",
    "für folgende adresse",
    "fuer folgende adresse",
    "ausgeführt bei",
    "ausgefuehrt bei",
    "betrifft baustelle",
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
    "pe",
    "ep",
    "preis",
    "einzelpreis",
    "e-preis",
    "betrag",
    "wert",
    "gesamt",
    "gp",
    "preisdimension",
    "nettopreis",
    "bruttopreis",
    "material",
    "art der leistung",
]

TOTAL_HINTS = [
    "summe positionen",
    "zwischensumme",
    "zwischensumme position",
    "zwischensumme (netto)",
    "zwischensumme vor steuer",
    "nettowarenwert",
    "nettobetrag",
    "netto-betrag",
    "netto",
    "warenwert",
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
    "brutto",
    "bruttobetrag",
    "zahlbetrag",
    "saldo",
    "gesamt ",
    "gesamtbetrag",
]

PAYMENT_HINTS = [
    "zahlungskonditionen",
    "zahlungsbedingungen",
    "zahlungsbedingung",
    "zahlungsart",
    "zahlweise",
    "zahlbar bis",
    "zahlbar ohne abzug",
    "ohne abzug",
    "unter abzug",
    "skonto",
    "skontobetrag",
    "skontofähiger betrag",
    "skontofaehiger betrag",
    "skontodatum",
    "skontosatz",
    "fälligkeit",
    "faelligkeit",
    "fälligkeitsdatum",
    "faelligkeitsdatum",
    "abbuchen",
    "abgebucht",
    "lastschrift",
    "sepa-lastschrift",
    "sepa lastschrift",
    "zahlung innerhalb",
    "wird zum",
    "wird von ihrem konto",
    "der betrag wird",
    "rechnungsbetrag bitte",
]

FOOTER_HINTS = [
    "iban",
    "bic",
    "swift",
    "bank",
    "bankverbindung",
    "ust-id",
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
    "eintragung in das handelsregister",
    "sitz:",
    "hrb",
    "hra",
]

NOTICE_HINTS = [
    "zahlungsavis",
    "überweisungsankündigungen",
    "ueberweisungsankuendigungen",
    "zentrale e-mail-adresse",
    "für ihre unterstützung bedanken wir uns",
    "fuer ihre unterstuetzung bedanken wir uns",
]

SUPPLIER_HINTS = {
    "hempelmann_gc": [
        "gc-gruppe.de",
        "bei schriftwechsel bitte angeben",
        "außend.",
        "innend.",
        "debitor :",
    ],
    "richter_frenzel": [
        "richter+frenzel",
        "richter+frenzel kassel",
        "r-f.de",
        "kommission:",
    ],
    "cl_bergmann": [
        "cl-bergmann.de",
        "beleg-nr.",
        "bestellangaben:",
        "zahlungskonditionen:",
    ],
    "weinmann_schanz": [
        "weinmann & schanz",
        "weinmann-schanz.de",
        "unser auftrag",
        "ihr auftrag",
    ],
    "vaillant": [
        "vaillant",
        "werkskundendienst",
        "leistungsgutschrift",
    ],
    "dittmar": [
        "bernhard dittmar",
        "dittmar-volkmarsen.de",
    ],
    "kowalski_service": [
        "kowalski-service",
        "garten und landschaftsbau",
        "hausmeister service",
        "trockenbau und innenausbau",
    ],
}

STAR_LINE_RE = re.compile(r"^\*{5,}$")
AMOUNT_RE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}")
DATE_RE = re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
IBAN_RE = re.compile(r"\b[A-Z]{2}\d{2}[A-Z0-9 ]{10,34}\b")

# ============================================================
# GENERISCHE TESTS
# ============================================================

def contains_any(text: str, variants: List[str]) -> bool:
    lt = lower(text)
    return any(v in lt for v in variants)

def is_star_line(line: str) -> bool:
    return STAR_LINE_RE.match(norm(line)) is not None

def has_amount(line: str) -> bool:
    return bool(AMOUNT_RE.search(line))

def has_date(line: str) -> bool:
    return bool(DATE_RE.search(line))

def looks_like_footer_line(line: str) -> bool:
    return contains_any(line, FOOTER_HINTS) or bool(IBAN_RE.search(line)) or bool(EMAIL_RE.search(line))

def looks_like_total_line(line: str) -> bool:
    return contains_any(line, TOTAL_HINTS) and (has_amount(line) or "%" in line)

def looks_like_payment_line(line: str) -> bool:
    return contains_any(line, PAYMENT_HINTS)

def looks_like_order_line(line: str) -> bool:
    return contains_any(line, ORDER_HINTS)

def looks_like_header_line(line: str) -> bool:
    return contains_any(line, HEADER_HINTS)

def looks_like_address_line(line: str) -> bool:
    l = lower(line)
    if re.search(r"\b\d{5}\b", line):
        return True
    if re.search(r"\bstr\.?\b|\bstraße\b|\bstrasse\b|\bweg\b|\ballee\b|\bplatz\b", l):
        return True
    return False

def looks_like_name_line(line: str) -> bool:
    toks = tokenize(line)
    if 1 <= len(toks) <= 5:
        alpha = [t for t in toks if re.search(r"[A-Za-zÄÖÜäöüß]", t)]
        if alpha and all(t[0].isupper() for t in alpha if t[0].isalpha()):
            if not looks_like_header_line(line) and not looks_like_order_line(line):
                return True
    return False

def is_single_letter_spaced_title(line: str) -> bool:
    toks = tokenize(line)
    return len(toks) >= 4 and all(len(t) == 1 and t.isalpha() for t in toks)

def looks_like_positions_header(line: str) -> bool:
    l = lower(line)
    score = sum(1 for h in POSITIONS_HEADER_HINTS if h in l)
    return score >= 2

def detect_supplier_hint(text_full: str) -> Optional[str]:
    l = lower(text_full)
    for key, variants in SUPPLIER_HINTS.items():
        if any(v in l for v in variants):
            return key
    return None

# ============================================================
# OCR / EXTRACTION
# ============================================================

def pil_to_bytes(img: Image.Image) -> bytes:
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    return bio.getvalue()

def score_ocr_text(text: str) -> int:
    t = norm(text)
    if not t:
        return 0

    score = 0
    score += min(len(t), 4000) // 10
    score += len(DATE_RE.findall(t)) * 25
    score += len(AMOUNT_RE.findall(t)) * 15
    score += t.lower().count("rechnung") * 40
    score += t.lower().count("gutschrift") * 40
    score += t.lower().count("betrag") * 10
    score += t.lower().count("gesamt") * 10
    score += t.lower().count("mwst") * 10
    score += t.lower().count("kommission") * 10

    lines = split_lines(t)
    score += min(len(lines), 120)

    return score

def preprocess_image_variants_for_ocr(img: Image.Image) -> List[Tuple[str, Image.Image]]:
    base = img.convert("L")
    base = ImageOps.exif_transpose(base)

    variants = []

    v1 = ImageOps.autocontrast(base)
    variants.append(("gray_autocontrast", v1))

    v2 = ImageOps.autocontrast(base).filter(ImageFilter.SHARPEN)
    variants.append(("gray_sharpen", v2))

    v3 = ImageOps.autocontrast(base)
    v3 = v3.point(lambda x: 255 if x > 180 else 0)
    variants.append(("threshold_180", v3))

    return variants

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

def extract_text_ocr_best(pdf_bytes: bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []

    max_ocr_pages = 3
    render_scale = 1.35
    tesseract_timeout_sec = 4

    for i, page in enumerate(doc):
        if i >= max_ocr_pages:
            break

        pix = page.get_pixmap(matrix=fitz.Matrix(render_scale, render_scale), alpha=False)
        original_img = Image.open(io.BytesIO(pix.tobytes("png")))

        best_text = ""
        best_score = -1
        best_variant_name = ""

        variants = preprocess_image_variants_for_ocr(original_img)

        for variant_name, img_variant in variants:
            try:
                text = pytesseract.image_to_string(
                    img_variant,
                    lang="deu",
                    timeout=tesseract_timeout_sec
                ) or ""
                score = score_ocr_text(text)

                if score > best_score:
                    best_score = score
                    best_text = text
                    best_variant_name = variant_name

            except RuntimeError:
                continue
            except Exception:
                continue

        pages.append({
            "page": i + 1,
            "text": best_text,
            "ocr_variant": best_variant_name,
            "ocr_score": best_score,
        })

    return "\n".join(p["text"] for p in pages), pages

def text_looks_bad(text: str) -> bool:
    t = norm(text)
    if not t:
        return True
    if len(t) < 120:
        return True
    return False

# ============================================================
# DOKUMENTTYP-HINWEIS
# ============================================================

def detect_document_type_hint(text_full: str) -> str:
    lines = split_lines(text_full)
    top = "\n".join(lines[:25])
    tl = lower(top)

    if contains_any(tl, DOC_TYPE_MARKERS["gutschrift"]):
        return "gutschrift"
    if contains_any(tl, DOC_TYPE_MARKERS["lieferschein"]):
        return "lieferschein"
    if contains_any(tl, DOC_TYPE_MARKERS["avis"]) and not contains_any(tl, DOC_TYPE_MARKERS["rechnung"]):
        return "avis"
    if contains_any(tl, DOC_TYPE_MARKERS["rechnung"]):
        return "rechnung"

    fl = lower(text_full)
    if contains_any(fl, DOC_TYPE_MARKERS["gutschrift"]):
        return "gutschrift"
    if contains_any(fl, DOC_TYPE_MARKERS["lieferschein"]):
        return "lieferschein"
    if contains_any(fl, DOC_TYPE_MARKERS["avis"]) and not contains_any(fl, DOC_TYPE_MARKERS["rechnung"]):
        return "avis"
    if contains_any(fl, DOC_TYPE_MARKERS["rechnung"]):
        return "rechnung"
    return "sonstiges"

# ============================================================
# HEADER-ROW-GROUPS
# ============================================================

def split_header_label_row(line: str) -> List[str]:
    known = [
        "KD-Nr.", "KD-Nr", "Kundennr.", "Kunden-Nr.", "Kunden-Nr", "Kundennr",
        "Rechn.Nr.", "Rechn.Nr", "Rechnungsnummer", "Rechnungsnr",
        "Beleg-Nr.", "Beleg-Nr", "Belegnummer",
        "Datum", "Belegdatum", "Leistungsdatum",
        "Blatt", "Seite"
    ]
    temp = norm(line)
    for k in known:
        temp = re.sub(rf"(?i)\b{re.escape(k)}\b", f"|||{k}|||", temp)
    parts = [norm(x) for x in temp.split("|||") if norm(x)]
    return [p for p in parts if any(lower(p) == lower(k) for k in known)]

def is_plausible_value_row(line: str) -> bool:
    if is_single_letter_spaced_title(line):
        return False
    if has_date(line) or has_amount(line):
        return True
    ts = tokenize(line)
    if len(ts) >= 2 and any(t.isdigit() for t in ts):
        return True
    if len(ts) >= 2 and any(re.fullmatch(r"[A-Z0-9\-\/\.]{4,}", t) for t in ts):
        return True
    return False

def build_header_row_groups(lines: List[str]) -> List[Dict[str, Any]]:
    groups = []
    for i in range(len(lines)):
        labels = split_header_label_row(lines[i])
        if not labels:
            continue

        value_lines = []
        if i + 1 < len(lines) and is_plausible_value_row(lines[i + 1]):
            value_lines.append(lines[i + 1])
        if i + 2 < len(lines) and is_plausible_value_row(lines[i + 2]):
            value_lines.append(lines[i + 2])

        if value_lines:
            groups.append({
                "label_line": lines[i],
                "labels": labels,
                "value_lines": value_lines,
            })
    return groups

# ============================================================
# POSITIONSGRUPPEN
# ============================================================

def line_has_qty_unit(line: str) -> bool:
    return bool(re.search(r"\d+,\d+\s*(stk|st|m|mtr|pa|dos|ein|kg|l|qm|pce)\b", lower(line)))

def line_has_price_value_pair(line: str) -> bool:
    return len(AMOUNT_RE.findall(line)) >= 2

def is_code_like(line: str) -> bool:
    return bool(re.fullmatch(r"[A-Z0-9][A-Z0-9\-\.\/]{3,}", norm(line)))

def has_trailing_amount(line: str) -> bool:
    return bool(re.search(r"-?\d{1,3}(?:\.\d{3})*,\d{2}\s*€?$", norm(line)))

def looks_like_service_description(line: str) -> bool:
    l = lower(line)

    if not norm(line):
        return False
    if looks_like_total_line(line) or looks_like_payment_line(line) or looks_like_footer_line(line):
        return False
    if is_star_line(line):
        return False
    if is_code_like(line):
        return False
    if len(tokenize(line)) < 2:
        return False
    if re.fullmatch(r"-?\d{1,3}(?:\.\d{3})*,\d{2}\s*€?", norm(line)):
        return False

    if any(x in l for x in [
        "fundament",
        "material",
        "entsorgung",
        "montage",
        "demontage",
        "arbeiten",
        "reparatur",
        "wartung",
        "service",
        "baustelle",
        "leistung",
        "bauen",
    ]):
        return True

    alpha_words = [t for t in tokenize(line) if re.search(r"[A-Za-zÄÖÜäöüß]", t)]
    return len(alpha_words) >= 2

def lookahead_has_position_evidence(lines: List[str], idx: int, window: int = 4) -> bool:
    for j in range(idx, min(len(lines), idx + window)):
        if line_has_qty_unit(lines[j]) or line_has_price_value_pair(lines[j]) or has_amount(lines[j]):
            return True
    return False

def is_probable_position_start(lines: List[str], idx: int) -> bool:
    line = norm(lines[idx])
    toks = tokenize(line)
    l = lower(line)

    if not line:
        return False
    if looks_like_order_line(line) or looks_like_payment_line(line) or looks_like_total_line(line):
        return False
    if looks_like_address_line(line) or looks_like_name_line(line):
        return False

    if re.match(r"^pos\.?\s*\d+", l):
        return True

    if re.match(r"^\d+\s+[A-Z0-9][A-Z0-9\-\/\.]{3,}", line):
        return True

    if is_code_like(line) and lookahead_has_position_evidence(lines, idx + 1, 4):
        return True

    if len(toks) >= 5 and toks[0].isdigit() and has_amount(line):
        return True

    # freie Leistungszeile mit späterem Betrag
    if looks_like_service_description(line) and lookahead_has_position_evidence(lines, idx + 1, 3):
        return True

    return False

def split_position_groups(position_lines: List[str]) -> List[List[str]]:
    cleaned = []

    for line in position_lines:
        l = lower(line)

        if not norm(line):
            continue
        if "übertrag" in l or "uebertrag" in l:
            continue
        if looks_like_total_line(line) or looks_like_payment_line(line):
            continue
        if looks_like_footer_line(line):
            continue
        if is_star_line(line):
            continue

        cleaned.append(line)

    if not cleaned:
        return []

    # --------------------------------------------------------
    # MODUS 1: klassische Artikel-/Materialrechnung
    # --------------------------------------------------------
    normal_starts = any(is_probable_position_start(cleaned, i) for i in range(len(cleaned)))

    if normal_starts:
        groups: List[List[str]] = []
        current: List[str] = []

        for i, line in enumerate(cleaned):
            if is_probable_position_start(cleaned, i):
                if current:
                    groups.append(current)
                current = [line]
            else:
                if current:
                    current.append(line)
                else:
                    current = [line]

        if current:
            groups.append(current)

        return groups

    # --------------------------------------------------------
    # MODUS 2: freie Leistungs-/Handwerkerrechnung / GP-only
    # --------------------------------------------------------
    groups: List[List[str]] = []
    current: List[str] = []

    for line in cleaned:
        current.append(line)

        if has_amount(line) or has_trailing_amount(line):
            groups.append(current)
            current = []

    if current:
        groups.append(current)

    groups = [g for g in groups if any(norm(x) for x in g)]
    return groups

# ============================================================
# BLOCK-STRUKTUR
# ============================================================

def find_positions_header_index(lines: List[str]) -> Optional[int]:
    for i, line in enumerate(lines):
        if looks_like_positions_header(line):
            return i
    return None

def find_footer_start_index(lines: List[str]) -> Optional[int]:
    count = 0
    first_idx = None
    for i in range(len(lines) - 1, -1, -1):
        if looks_like_footer_line(lines[i]):
            count += 1
            first_idx = i
        elif count >= 2:
            break
    return first_idx

def find_first_position_after_header(lines: List[str], start_idx: int) -> Optional[int]:
    for i in range(start_idx, len(lines)):
        if is_probable_position_start(lines, i):
            return i
        if looks_like_total_line(lines[i]) or looks_like_payment_line(lines[i]):
            return None
    return None

def build_structure(lines: List[str], supplier_hint: Optional[str]) -> Dict[str, Any]:
    footer_start = find_footer_start_index(lines)
    core_lines = lines[:footer_start] if footer_start is not None else lines
    footer_lines = lines[footer_start:] if footer_start is not None else []

    positions_header_idx = find_positions_header_index(core_lines)

    kopf_block_lines: List[str] = []
    adress_block_lines: List[str] = []
    auftrag_kommission_block_lines: List[str] = []
    positionskopf_block_lines: List[str] = []
    positions_lines: List[str] = []
    summenblock_lines: List[str] = []
    zahlungsblock_lines: List[str] = []

    fallback_before_positions: List[str] = []
    fallback_inside_positions: List[str] = []
    fallback_after_totals: List[str] = []
    fallback_global: List[str] = []

    if positions_header_idx is None:
        for line in core_lines:
            if looks_like_total_line(line):
                summenblock_lines.append(line)
            elif looks_like_payment_line(line) or contains_any(line, NOTICE_HINTS):
                zahlungsblock_lines.append(line)
            elif looks_like_order_line(line):
                auftrag_kommission_block_lines.append(line)
            elif looks_like_header_line(line):
                kopf_block_lines.append(line)
            else:
                adress_block_lines.append(line)
                fallback_global.append(line)
    else:
        upper = core_lines[:positions_header_idx]
        positionskopf_block_lines = [core_lines[positions_header_idx]]
        lower_part = core_lines[positions_header_idx + 1:]

        for line in upper:
            if looks_like_order_line(line):
                auftrag_kommission_block_lines.append(line)
            elif looks_like_header_line(line) or is_single_letter_spaced_title(line):
                kopf_block_lines.append(line)
            else:
                adress_block_lines.append(line)

        first_pos_rel = find_first_position_after_header(lower_part, 0)
        if first_pos_rel is None:
            first_pos_rel = len(lower_part)

        pre_position = lower_part[:first_pos_rel]
        rest = lower_part[first_pos_rel:]

        for line in pre_position:
            if looks_like_total_line(line):
                summenblock_lines.append(line)
            elif looks_like_payment_line(line) or contains_any(line, NOTICE_HINTS):
                zahlungsblock_lines.append(line)
            elif looks_like_order_line(line) or looks_like_address_line(line) or looks_like_name_line(line):
                auftrag_kommission_block_lines.append(line)
            else:
                auftrag_kommission_block_lines.append(line)
                fallback_before_positions.append(line)

        mode = "positions"
        for line in rest:
            if is_star_line(line) or contains_any(line, NOTICE_HINTS):
                mode = "payment"
                zahlungsblock_lines.append(line)
                continue

            if looks_like_total_line(line):
                mode = "totals"
                summenblock_lines.append(line)
                continue

            if looks_like_payment_line(line):
                mode = "payment"
                zahlungsblock_lines.append(line)
                continue

            if mode == "positions":
                positions_lines.append(line)
            elif mode == "totals":
                if looks_like_footer_line(line):
                    fallback_after_totals.append(line)
                else:
                    summenblock_lines.append(line)
            else:
                if looks_like_footer_line(line):
                    fallback_after_totals.append(line)
                else:
                    zahlungsblock_lines.append(line)

    if supplier_hint in ("hempelmann_gc", "richter_frenzel", "weinmann_schanz", "kowalski_service"):
        cleaned_positions = []
        for line in positions_lines:
            if looks_like_order_line(line) or looks_like_name_line(line) or looks_like_address_line(line):
                auftrag_kommission_block_lines.append(line)
                fallback_inside_positions.append(line)
            else:
                cleaned_positions.append(line)
        positions_lines = cleaned_positions

    position_groups_raw = split_position_groups(positions_lines)

    position_groups = []
    used_lines = set()

    for idx, group_lines in enumerate(position_groups_raw, start=1):
        for gl in group_lines:
            used_lines.add(gl)
        position_groups.append({
            "group_index": idx,
            "lines": group_lines,
            "text": join_lines(group_lines),
        })

    for line in positions_lines:
        if line not in used_lines:
            fallback_inside_positions.append(line)

    structure_warnings = []

    if not kopf_block_lines:
        structure_warnings.append("HEADER_WEAK")
    if not positionskopf_block_lines:
        structure_warnings.append("POSITIONS_HEADER_WEAK")
    if not position_groups:
        structure_warnings.append("POSITION_GROUPS_WEAK")
    if not summenblock_lines:
        structure_warnings.append("TOTALS_WEAK")
    if not auftrag_kommission_block_lines:
        structure_warnings.append("ORDER_BLOCK_WEAK")
    if fallback_before_positions or fallback_inside_positions or fallback_after_totals or fallback_global:
        structure_warnings.append("FALLBACK_LINES_PRESENT")

    return {
        "kopf_block_lines": kopf_block_lines,
        "adress_block_lines": adress_block_lines,
        "auftrag_kommission_block_lines": auftrag_kommission_block_lines,
        "positionskopf_block_lines": positionskopf_block_lines,
        "positionsgruppen": position_groups,
        "summenblock_lines": summenblock_lines,
        "zahlungsblock_lines": zahlungsblock_lines,
        "footer_block_lines": footer_lines,
        "fallback": {
            "fallback_before_positions": fallback_before_positions,
            "fallback_inside_positions": fallback_inside_positions,
            "fallback_after_totals": fallback_after_totals,
            "fallback_global": fallback_global,
        },
        "structure_warnings": structure_warnings,
    }

# ============================================================
# QUALITY
# ============================================================

def compute_structure_quality(structure: Dict[str, Any], text_full: str, ocr_used: bool) -> Dict[str, Any]:
    score = 0.0
    reasons = []

    if len(norm(text_full)) > 120:
        score += 0.10
    else:
        reasons.append("TEXT_SHORT")

    if structure["kopf_block_lines"]:
        score += 0.15
    else:
        reasons.append("KOPF_BLOCK_EMPTY")

    if structure["positionskopf_block_lines"]:
        score += 0.10
    else:
        reasons.append("POSITIONSKOPF_BLOCK_EMPTY")

    if structure["positionsgruppen"]:
        score += 0.25
    else:
        reasons.append("POSITIONSGRUPPEN_EMPTY")

    if structure["summenblock_lines"]:
        score += 0.15
    else:
        reasons.append("SUMMENBLOCK_EMPTY")

    if structure["zahlungsblock_lines"]:
        score += 0.10

    if structure["auftrag_kommission_block_lines"]:
        score += 0.10

    if structure["footer_block_lines"]:
        score += 0.05

    if ocr_used:
        score -= 0.05

    score = max(0.0, min(round(score, 2), 1.0))

    return {
        "score": score,
        "usable": score >= 0.60,
        "reasons": reasons
    }

# ============================================================
# OUTPUT
# ============================================================

def build_output(
    text_full: str,
    pages: List[Dict[str, Any]],
    text_engine: str,
    ocr_used: bool,
    known_betrieb_name: Optional[str],
    pymupdf_error=None,
    pdfplumber_error=None,
    ocr_error=None
) -> Dict[str, Any]:

    lines = split_lines(text_full)
    supplier_hint = detect_supplier_hint(text_full)
    document_type_hint = detect_document_type_hint(text_full)

    structure = build_structure(lines, supplier_hint=supplier_hint)
    quality = compute_structure_quality(structure, text_full, ocr_used)

    known_betrieb_match = False
    if known_betrieb_name:
        kb = lower(known_betrieb_name)
        if kb and kb in lower(text_full):
            known_betrieb_match = True

    header_row_groups = build_header_row_groups(
        structure["kopf_block_lines"] + structure["adress_block_lines"]
    )

    blocks = {
        "kopf_block": join_lines(structure["kopf_block_lines"]),
        "adress_block": join_lines(structure["adress_block_lines"]),
        "auftrag_kommission_block": join_lines(structure["auftrag_kommission_block_lines"]),
        "positionskopf_block": join_lines(structure["positionskopf_block_lines"]),
        "positionsblock": join_lines([line for g in structure["positionsgruppen"] for line in g["lines"]]),
        "summenblock": join_lines(structure["summenblock_lines"]),
        "zahlungsblock": join_lines(structure["zahlungsblock_lines"]),
        "footer_block": join_lines(structure["footer_block_lines"]),
    }

    return {
        "ok": True,
        "meta": {
            "extractor": "cloudrun-v4.3-structure",
            "text_engine": text_engine,
            "ocr_used": ocr_used,
            "page_count": len(pages),
            "chars": len(text_full),
        },
        "hints": {
            "document_type_hint": document_type_hint,
            "supplier_hint": supplier_hint,
            "known_betrieb_name": known_betrieb_name,
            "known_betrieb_match": known_betrieb_match,
        },
        "quality": quality,
        "structure": {
            **structure,
            "header_row_groups": header_row_groups,
        },
        "blocks": blocks,
        "pages": pages,
        "text_full": text_full,
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
    return "PDF Extractor v4.3 structure-first is running", 200

@app.route("/extract", methods=["POST"])
def extract_pdf():
    try:
        if "file" not in request.files:
            return jsonify({
                "ok": False,
                "error": "no_file_provided",
                "error_detail": "No file provided",
                "meta": {
                    "extractor": "cloudrun-v4.3-structure",
                    "text_engine": "none",
                    "ocr_used": False,
                    "page_count": 0,
                    "chars": 0,
                },
                "hints": {
                    "document_type_hint": "sonstiges",
                    "supplier_hint": None,
                    "known_betrieb_name": None,
                    "known_betrieb_match": False,
                },
                "quality": {
                    "score": 0.0,
                    "usable": False,
                    "reasons": ["NO_FILE_PROVIDED"]
                },
                "structure": {
                    "kopf_block_lines": [],
                    "adress_block_lines": [],
                    "auftrag_kommission_block_lines": [],
                    "positionskopf_block_lines": [],
                    "positionsgruppen": [],
                    "summenblock_lines": [],
                    "zahlungsblock_lines": [],
                    "footer_block_lines": [],
                    "fallback": {
                        "fallback_before_positions": [],
                        "fallback_inside_positions": [],
                        "fallback_after_totals": [],
                        "fallback_global": [],
                    },
                    "structure_warnings": ["EXTRACT_FAILED"],
                    "header_row_groups": [],
                },
                "blocks": {
                    "kopf_block": "",
                    "adress_block": "",
                    "auftrag_kommission_block": "",
                    "positionskopf_block": "",
                    "positionsblock": "",
                    "summenblock": "",
                    "zahlungsblock": "",
                    "footer_block": "",
                },
                "pages": [],
                "text_full": "",
                "debug": {
                    "pymupdf_error": None,
                    "pdfplumber_error": None,
                    "ocr_error": None,
                }
            }), 200

        file = request.files["file"]

        if not file.filename or not file.filename.lower().endswith(".pdf"):
            return jsonify({
                "ok": False,
                "error": "invalid_file_type",
                "error_detail": "File must be a PDF",
                "meta": {
                    "extractor": "cloudrun-v4.3-structure",
                    "text_engine": "none",
                    "ocr_used": False,
                    "page_count": 0,
                    "chars": 0,
                },
                "hints": {
                    "document_type_hint": "sonstiges",
                    "supplier_hint": None,
                    "known_betrieb_name": None,
                    "known_betrieb_match": False,
                },
                "quality": {
                    "score": 0.0,
                    "usable": False,
                    "reasons": ["INVALID_FILE_TYPE"]
                },
                "structure": {
                    "kopf_block_lines": [],
                    "adress_block_lines": [],
                    "auftrag_kommission_block_lines": [],
                    "positionskopf_block_lines": [],
                    "positionsgruppen": [],
                    "summenblock_lines": [],
                    "zahlungsblock_lines": [],
                    "footer_block_lines": [],
                    "fallback": {
                        "fallback_before_positions": [],
                        "fallback_inside_positions": [],
                        "fallback_after_totals": [],
                        "fallback_global": [],
                    },
                    "structure_warnings": ["EXTRACT_FAILED"],
                    "header_row_groups": [],
                },
                "blocks": {
                    "kopf_block": "",
                    "adress_block": "",
                    "auftrag_kommission_block": "",
                    "positionskopf_block": "",
                    "positionsblock": "",
                    "summenblock": "",
                    "zahlungsblock": "",
                    "footer_block": "",
                },
                "pages": [],
                "text_full": "",
                "debug": {
                    "pymupdf_error": None,
                    "pdfplumber_error": None,
                    "ocr_error": None,
                }
            }), 200

        known_betrieb_name = request.form.get("known_betrieb_name", "").strip() or None
        pdf_bytes = file.read()

        text_full = ""
        pages: List[Dict[str, Any]] = []
        text_engine = "none"
        ocr_used = False

        pymupdf_error = None
        pdfplumber_error = None
        ocr_error = None

        # 1) PyMuPDF
        try:
            text_full, pages = extract_text_pymupdf(pdf_bytes)
            text_engine = "pymupdf"
        except Exception as e:
            pymupdf_error = str(e)

        # 2) pdfplumber
        if text_looks_bad(text_full):
            try:
                text_pp, pages_pp = extract_text_pdfplumber(pdf_bytes)
                if len(norm(text_pp)) > len(norm(text_full)):
                    text_full, pages = text_pp, pages_pp
                    text_engine = "pdfplumber"
                if text_looks_bad(text_full):
                    raise ValueError("pdfplumber text still weak")
            except Exception as e:
                pdfplumber_error = str(e)

               # 3) OCR Best-Variant
        if text_looks_bad(text_full):
            try:
                text_ocr, pages_ocr = extract_text_ocr_best(pdf_bytes)
                if norm(text_ocr) and len(norm(text_ocr)) > len(norm(text_full)):
                    text_full, pages = text_ocr, pages_ocr
                    text_engine = "ocr_tesseract_best"
                    ocr_used = True
            except Exception as e:
                ocr_error = f"OCR_FAILED: {str(e)}"

        if not norm(text_full):
            return jsonify({
                "ok": False,
                "error": "no_usable_text_extracted",
                "error_detail": "No usable text could be extracted",
                "meta": {
                    "extractor": "cloudrun-v4.3-structure",
                    "text_engine": text_engine,
                    "ocr_used": ocr_used,
                    "page_count": len(pages),
                    "chars": len(text_full),
                },
                "hints": {
                    "document_type_hint": "sonstiges",
                    "supplier_hint": None,
                    "known_betrieb_name": known_betrieb_name,
                    "known_betrieb_match": False,
                },
                "quality": {
                    "score": 0.0,
                    "usable": False,
                    "reasons": ["NO_USABLE_TEXT_EXTRACTED"]
                },
                "structure": {
                    "kopf_block_lines": [],
                    "adress_block_lines": [],
                    "auftrag_kommission_block_lines": [],
                    "positionskopf_block_lines": [],
                    "positionsgruppen": [],
                    "summenblock_lines": [],
                    "zahlungsblock_lines": [],
                    "footer_block_lines": [],
                    "fallback": {
                        "fallback_before_positions": [],
                        "fallback_inside_positions": [],
                        "fallback_after_totals": [],
                        "fallback_global": [],
                    },
                    "structure_warnings": ["NO_USABLE_TEXT_EXTRACTED"],
                    "header_row_groups": [],
                },
                "blocks": {
                    "kopf_block": "",
                    "adress_block": "",
                    "auftrag_kommission_block": "",
                    "positionskopf_block": "",
                    "positionsblock": "",
                    "summenblock": "",
                    "zahlungsblock": "",
                    "footer_block": "",
                },
                "pages": pages,
                "text_full": text_full,
                "debug": {
                    "pymupdf_error": pymupdf_error,
                    "pdfplumber_error": pdfplumber_error,
                    "ocr_error": ocr_error,
                }
            }), 200

        try:
            result = build_output(
                text_full=text_full,
                pages=pages,
                text_engine=text_engine,
                ocr_used=ocr_used,
                known_betrieb_name=known_betrieb_name,
                pymupdf_error=pymupdf_error,
                pdfplumber_error=pdfplumber_error,
                ocr_error=ocr_error,
            )
            return jsonify(result), 200

        except Exception as e:
            return jsonify({
                "ok": False,
                "error": "parse_failed",
                "error_detail": str(e),
                "meta": {
                    "extractor": "cloudrun-v4.3-structure",
                    "text_engine": text_engine,
                    "ocr_used": ocr_used,
                    "page_count": len(pages),
                    "chars": len(text_full),
                },
                "hints": {
                    "document_type_hint": "sonstiges",
                    "supplier_hint": None,
                    "known_betrieb_name": known_betrieb_name,
                    "known_betrieb_match": False,
                },
                "quality": {
                    "score": 0.0,
                    "usable": False,
                    "reasons": ["PARSE_FAILED"]
                },
                "structure": {
                    "kopf_block_lines": [],
                    "adress_block_lines": [],
                    "auftrag_kommission_block_lines": [],
                    "positionskopf_block_lines": [],
                    "positionsgruppen": [],
                    "summenblock_lines": [],
                    "zahlungsblock_lines": [],
                    "footer_block_lines": [],
                    "fallback": {
                        "fallback_before_positions": [],
                        "fallback_inside_positions": [],
                        "fallback_after_totals": [],
                        "fallback_global": [],
                    },
                    "structure_warnings": ["PARSE_FAILED"],
                    "header_row_groups": [],
                },
                "blocks": {
                    "kopf_block": "",
                    "adress_block": "",
                    "auftrag_kommission_block": "",
                    "positionskopf_block": "",
                    "positionsblock": "",
                    "summenblock": "",
                    "zahlungsblock": "",
                    "footer_block": "",
                },
                "pages": pages,
                "text_full": text_full,
                "debug": {
                    "pymupdf_error": pymupdf_error,
                    "pdfplumber_error": pdfplumber_error,
                    "ocr_error": ocr_error,
                }
            }), 200

    except Exception as e:
        return jsonify({
            "ok": False,
            "error": "extract_failed",
            "error_detail": str(e),
            "meta": {
                "extractor": "cloudrun-v4.3-structure",
                "text_engine": "none",
                "ocr_used": False,
                "page_count": 0,
                "chars": 0,
            },
            "hints": {
                "document_type_hint": "sonstiges",
                "supplier_hint": None,
                "known_betrieb_name": None,
                "known_betrieb_match": False,
            },
            "quality": {
                "score": 0.0,
                "usable": False,
                "reasons": ["EXTRACT_FAILED"]
            },
            "structure": {
                "kopf_block_lines": [],
                "adress_block_lines": [],
                "auftrag_kommission_block_lines": [],
                "positionskopf_block_lines": [],
                "positionsgruppen": [],
                "summenblock_lines": [],
                "zahlungsblock_lines": [],
                "footer_block_lines": [],
                "fallback": {
                    "fallback_before_positions": [],
                    "fallback_inside_positions": [],
                    "fallback_after_totals": [],
                    "fallback_global": [],
                },
                "structure_warnings": ["EXTRACT_FAILED"],
                "header_row_groups": [],
            },
            "blocks": {
                "kopf_block": "",
                "adress_block": "",
                "auftrag_kommission_block": "",
                "positionskopf_block": "",
                "positionsblock": "",
                "summenblock": "",
                "zahlungsblock": "",
                "footer_block": "",
            },
            "pages": [],
            "text_full": "",
            "debug": {
                "pymupdf_error": None,
                "pdfplumber_error": None,
                "ocr_error": None,
            }
        }), 200
        
# ============================================================
# ANALYZE HELPERS / REPORT LOGIK V5
# ============================================================

KANON_HINWEIS_TYPEN = [
    "PREISABWEICHUNG",
    "MENGENABWEICHUNG",
    "PREISSPRUNG_AUFFAELLIG",
    "GEBUEHR_AUFFAELLIG",
    "GEBUEHR_POSITION",
    "GEBUEHR_ERKANNT",
    "GUTSCHRIFT_ERKANNT",
    "GUTSCHRIFT_POSITION",
    "SKONTO_ERKANNT",
    "SKONTO_ABWEICHUNG",
    "KOMMISSION_FEHLT",
    "KOMMISSION_UNKLAR",
    "DUPLIKAT_RECHNUNG",
    "DOPPELTE_POSITION",
    "MWST_UNPLAUSIBEL",
    "KONTO_ABWEICHUNG",
    "EXTRAKTION_FEHLER",
    "JSON_UNVOLLSTAENDIG",
    "RECHNUNGSNUMMER_FEHLT",
    "ARTIKELNUMMER_FEHLT",
    "ARTIKELNUMMER_UNGUELTIG",
    "LIEFERANT_NICHT_ERKANNT",
    "BETRIEB_NICHT_ERKANNT",
    "LIEFERANT_AUTO_ERSTELLT",
]

FACHLICHE_HINWEISE_BASIS = {
    "PREISABWEICHUNG",
    "MENGENABWEICHUNG",
    "PREISSPRUNG_AUFFAELLIG",
    "GEBUEHR_AUFFAELLIG",
    "GEBUEHR_POSITION",
    "SKONTO_ABWEICHUNG",
    "DUPLIKAT_RECHNUNG",
    "DOPPELTE_POSITION",
    "MWST_UNPLAUSIBEL",
    "KONTO_ABWEICHUNG",
    "KOMMISSION_FEHLT",
    "KOMMISSION_UNKLAR",
}

TECHNISCHE_HINWEISE_BASIS = {
    "EXTRAKTION_FEHLER",
    "JSON_UNVOLLSTAENDIG",
    "RECHNUNGSNUMMER_FEHLT",
    "ARTIKELNUMMER_FEHLT",
    "ARTIKELNUMMER_UNGUELTIG",
    "LIEFERANT_NICHT_ERKANNT",
    "BETRIEB_NICHT_ERKANNT",
    "LIEFERANT_AUTO_ERSTELLT",
    "SKONTO_ERKANNT",
    "GEBUEHR_ERKANNT",
    "GUTSCHRIFT_ERKANNT",
}

GENERISCHE_KOSTENSTELLEN = {
    "P1",
    "LAGER",
    "STANDARD",
    "INTERN",
    "SAMMEL",
}

PROJEKT_KATEGORIEN = {
    "GROSSHANDEL",
    "HERSTELLER",
    "SUBUNTERNEHMER",
}

NICHT_PROJEKT_KATEGORIEN = {
    "DIENSTLEISTER",
    "FIXKOSTEN",
    "WERKSTATT",
    "ARBEITSKLEIDUNG",
    "SONSTIGES",
}

LIEFERANTEN_TYP_MAPPING = {
    "GROSSHAENDLER": "GROSSHANDEL",
    "GROSSHÄNDLER": "GROSSHANDEL",
    "GROSSHANDEL": "GROSSHANDEL",
    "FACHGROSSHANDEL": "GROSSHANDEL",
    "HANDEL": "GROSSHANDEL",
    "HANDEL_ALLGEMEIN": "GROSSHANDEL",

    "HERSTELLER": "HERSTELLER",

    "SUBUNTERNEHMER": "SUBUNTERNEHMER",
    "SUB": "SUBUNTERNEHMER",
    "NACHUNTERNEHMER": "SUBUNTERNEHMER",

    "DIENSTLEISTER": "DIENSTLEISTER",
    "SERVICE": "DIENSTLEISTER",

    "FIXKOSTEN": "FIXKOSTEN",
    "WERKSTATT": "WERKSTATT",
    "ARBEITSKLEIDUNG": "ARBEITSKLEIDUNG",
    "SONSTIGES": "SONSTIGES",
    "FUHRPARK": "SONSTIGES",
    "MIETE": "SONSTIGES",
    "TRANSPORT": "SONSTIGES",
    "ENTSORGUNG": "SONSTIGES",
}

def normalize_text_basic(value):
    s = str(value or "").strip().lower()
    s = s.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    s = s.replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def normalize_code(value):
    s = normalize_text_basic(value)
    s = re.sub(r"[^a-z0-9]", "", s)
    return s

def normalize_name(value):
    s = normalize_text_basic(value)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_address(value):
    s = normalize_text_basic(value)
    if not s:
        return ""

    replacements = {
        "straße": "str",
        "strasse": "str",
        "str.": "str",
        "platz": "pl",
        "allee": "all",
        "deutschland": "",
        "de": "",
    }

    s = s.replace(",", " ")
    s = s.replace(".", " ")
    s = s.replace("-", " ")

    for k, v in replacements.items():
        s = s.replace(k, v)

    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_address_key(value):
    s = normalize_address(value)
    if not s:
        return ""

    s = f" {s} "
    noise_patterns = [
        r"\bdeutschland\b",
        r"\bde\b",
        r"\binnend\b",
        r"\baussend\b",
        r"\baußend\b",
        r"\bvom\b.*$",
        r"\babholung\b.*$",
        r"\blieferung\b.*$",
        r"\bkom\b.*$",
        r"\bupdate\b.*$",
        r"\bgebaeudetechnik\b",
        r"\bgmbh\b",
    ]
    for p in noise_patterns:
        s = re.sub(p, " ", s)

    s = re.sub(r"\s+", " ", s).strip()

    street_patterns = [
        r"([a-z0-9 ]+?\bstr\b\s+\d+[a-z]?)",
        r"([a-z0-9 ]+?\bweg\b\s+\d+[a-z]?)",
        r"([a-z0-9 ]+?\ballee\b\s+\d+[a-z]?)",
        r"([a-z0-9 ]+?\bplatz\b\s+\d+[a-z]?)",
        r"([a-z0-9 ]+?\bgasse\b\s+\d+[a-z]?)",
        r"([a-z0-9 ]+?\blehnhof\b\s+\d+[a-z]?)",
        r"([a-z0-9 ]+?\bweiden\b\s+\d+[a-z]?)",
        r"([a-z0-9 ]+?\brepsch\b\s+\d+[a-z]?)",
        r"([a-z0-9 ]+?\btrift\b\s+\d+[a-z]?)",
    ]

    candidate = ""
    for pattern in street_patterns:
        m = re.search(pattern, s)
        if m:
            candidate = m.group(1).strip()
            break

    if not candidate:
        tokens = s.split()
        haus_idx = None
        for i, t in enumerate(tokens):
            if re.fullmatch(r"\d+[a-z]?", t):
                haus_idx = i
                break

        if haus_idx is None:
            return ""

        start = max(0, haus_idx - 4)
        end = min(len(tokens), haus_idx + 1)
        candidate = " ".join(tokens[start:end]).strip()

    candidate = re.sub(r"\s+", " ", candidate).strip()

    plz_match = re.search(r"\b(\d{5})\b", s)
    plz = plz_match.group(1) if plz_match else ""

    return f"{candidate}|{plz}" if plz else candidate

def normalize_person_name_for_match(value):
    s = normalize_name(value)
    if not s:
        return ""

    s = s.replace(",", " ")
    s = re.sub(r"\s+", " ", s).strip()

    blacklist = {
        "lager", "webshop", "elements", "abholung", "abholer", "lieferung",
        "innend", "aussend", "außend", "kunde", "projekt", "baustelle",
        "van", "de", "meent"
    }

    nickname_map = {
        "alex": "alexander",
    }

    parts = []
    for p in s.split():
        if not p:
            continue
        if re.search(r"\d", p):
            continue
        if p in blacklist:
            continue
        p = nickname_map.get(p, p)
        parts.append(p)

    if not parts:
        return ""

    return " ".join(parts)

def is_full_person_name(value):
    s = normalize_person_name_for_match(value)
    if not s:
        return False

    parts = [p for p in s.split() if p]
    if len(parts) < 2:
        return False

    if any(re.search(r"\d", p) for p in parts):
        return False

    if any(len(p) < 2 for p in parts):
        return False

    blacklist = {
        "ks", "kb", "whg", "haus", "objekt", "projekt", "baustelle",
        "lager", "kunde", "webshop", "elements", "abholung"
    }

    if any(p in blacklist for p in parts):
        return False

    return True

def text_similarity(a, b):
    a = normalize_name(a)
    b = normalize_name(b)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()

def to_float_safe(value):
    if value is None:
        return 0.0

    if isinstance(value, (int, float)):
        try:
            return float(value)
        except Exception:
            return 0.0

    s = str(value).strip()
    if not s:
        return 0.0

    s = s.replace("€", "").replace("%", "").replace(" ", "").replace("\u00a0", "")

    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")

    s = re.sub(r"[^0-9\.\-]", "", s)

    if not s or s in ("-", ".", "-.", ".-"):
        return 0.0

    try:
        return float(s)
    except Exception:
        return 0.0

def parse_date_safe(value):
    if not value:
        return None

    if isinstance(value, date):
        return value

    if isinstance(value, datetime):
        return value.date()

    s = str(value).strip()
    if not s:
        return None

    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00")).date()
    except Exception:
        pass

    formats = [
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%Y/%m/%d",
        "%d-%m-%Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass

    return None

def canonical_hint_type(value):
    s = str(value or "").strip().upper()
    return s if s else "UNBEKANNT"

def canonical_hint_klasse(value):
    s = str(value or "").strip().upper()
    if s in ("FACHLICH", "TECHNISCH"):
        return s
    return ""

def canonical_pruefung_status(value):
    s = str(value or "").strip().upper()
    if s in ("OFFEN", "IN_PRUEFUNG", "ABGESCHLOSSEN"):
        return s
    return "OFFEN"

def canonical_gesamtbewertung(value):
    s = str(value or "").strip().upper()
    if s == "HINWEIS":
        return "HINWEIS"
    return "OK"

def canonical_dokumenttyp(value):
    s = str(value or "").strip().upper()
    if s == "GUTSCHRIFT":
        return "GUTSCHRIFT"
    if s == "RECHNUNG":
        return "RECHNUNG"
    return "SONSTIGES"

def canonical_lieferanten_kategorie(value):
    s = str(value or "").strip().upper()

    if not s:
        return "SONSTIGES"

    mapping = {
        "GROSSHAENDLER": "GROSSHANDEL",
        "GROSSHÄNDLER": "GROSSHANDEL",
        "GROSSHANDEL": "GROSSHANDEL",
        "FACHGROSSHANDEL": "GROSSHANDEL",
        "HANDEL": "GROSSHANDEL",
        "HANDEL_ALLGEMEIN": "GROSSHANDEL",

        "SUBUNTERNEHMER": "SUBUNTERNEHMER",
        "SUB": "SUBUNTERNEHMER",
        "NACHUNTERNEHMER": "SUBUNTERNEHMER",

        "DIENSTLEISTER": "DIENSTLEISTER",
        "SERVICE": "DIENSTLEISTER",

        "HERSTELLER": "HERSTELLER",
        "FIXKOSTEN": "FIXKOSTEN",
        "WERKSTATT": "WERKSTATT",
        "ARBEITSKLEIDUNG": "ARBEITSKLEIDUNG",
        "SONSTIGES": "SONSTIGES",
    }

    return mapping.get(s, s)

def map_lieferant_typ_to_kategorie(value):
    s = canonical_lieferanten_kategorie(value)
    return LIEFERANTEN_TYP_MAPPING.get(s, "SONSTIGES")

def is_projekt_kategorie(value):
    return map_lieferant_typ_to_kategorie(value) in PROJEKT_KATEGORIEN

def is_nicht_projekt_kategorie(value):
    return map_lieferant_typ_to_kategorie(value) in NICHT_PROJEKT_KATEGORIEN

def get_rechnung_id(r):
    return str(
        r.get("rechnung_id")
        or r.get("Rechnung_ID")
        or r.get("id")
        or r.get("ID")
        or ""
    ).strip()

def get_rechnungsnummer(r):
    return str(
        r.get("rechnungsnummer")
        or r.get("Rechnungsnummer")
        or ""
    ).strip()

def get_dokumenttyp(r):
    return canonical_dokumenttyp(
        r.get("dokumenttyp")
        or r.get("Dokumenttyp")
        or ""
    )

def get_lieferant_name(r):
    return str(
        r.get("lieferant_name")
        or r.get("Lieferant_Name")
        or r.get("lieferant")
        or r.get("Lieferant")
        or r.get("lieferant_id")
        or r.get("Lieferant_ID")
        or "UNBEKANNT"
    ).strip() or "UNBEKANNT"

def get_lieferant_id(r):
    return str(
        r.get("lieferant_id")
        or r.get("Lieferant_ID")
        or get_lieferant_name(r)
    ).strip()

def get_lieferant_typ_from_rechnung(r):
    return canonical_lieferanten_kategorie(
        r.get("lieferant_typ")
        or r.get("Lieferant_Typ")
        or r.get("lieferanten_typ")
        or r.get("Lieferanten_Typ")
        or r.get("lieferantenkategorie")
        or r.get("Lieferantenkategorie")
        or r.get("kategorie")
        or r.get("Kategorie")
        or ""
    )

def get_lieferant_typ_from_kontext_map(r, lieferanten_kontext_map):
    lid = get_lieferant_id(r)
    lname = get_lieferant_name(r)

    if lid and lid in lieferanten_kontext_map:
        return canonical_lieferanten_kategorie(lieferanten_kontext_map[lid].get("lieferant_typ"))

    lname_key = f"NAME::{normalize_name(lname)}"
    if lname_key in lieferanten_kontext_map:
        return canonical_lieferanten_kategorie(lieferanten_kontext_map[lname_key].get("lieferant_typ"))

    return ""

def get_lieferant_kategorie(r, lieferanten_kontext_map=None):
    lieferanten_kontext_map = lieferanten_kontext_map or {}

    typ = get_lieferant_typ_from_kontext_map(r, lieferanten_kontext_map)
    if not typ:
        typ = get_lieferant_typ_from_rechnung(r)

    return map_lieferant_typ_to_kategorie(typ)

def get_pruefung_status(r):
    return canonical_pruefung_status(
        r.get("pruefung_status")
        or r.get("Pruefung_Status")
        or ""
    )

def get_gesamtbewertung(r):
    return canonical_gesamtbewertung(
        r.get("gesamtbewertung")
        or r.get("Gesamtbewertung")
        or ""
    )

def get_brutto_summe(r):
    return to_float_safe(
        r.get("brutto_summe")
        or r.get("Brutto_Summe")
        or r.get("gesamt_brutto")
        or r.get("Gesamt_Brutto")
        or 0
    )

def get_netto_summe(r):
    return to_float_safe(
        r.get("netto_summe")
        or r.get("Netto_Summe")
        or r.get("gesamt_netto")
        or r.get("Gesamt_Netto")
        or 0
    )

def get_faelligkeitsdatum(r):
    return parse_date_safe(
        r.get("faelligkeitsdatum")
        or r.get("Faelligkeitsdatum")
        or r.get("zahlungsziel")
        or r.get("Zahlungsziel")
        or r.get("zahlungsziel_datum")
        or r.get("Zahlungsziel_Datum")
    )

def get_rechnungsdatum(r):
    return parse_date_safe(
        r.get("rechnungsdatum")
        or r.get("Rechnungsdatum")
    )

def get_eingangsdatum(r):
    return parse_date_safe(
        r.get("eingangsdatum")
        or r.get("Eingangsdatum")
    )

def get_skonto_datum(r):
    return parse_date_safe(
        r.get("skonto_datum")
        or r.get("Skonto_Datum")
        or r.get("skontodatum")
        or r.get("Skontodatum")
    )

def get_report_relevantes_datum(r):
    return get_eingangsdatum(r) or get_rechnungsdatum(r)

def is_im_zeitraum(d, zeitraum_start, zeitraum_ende):
    if not d:
        return False
    if zeitraum_start and d < zeitraum_start:
        return False
    if zeitraum_ende and d > zeitraum_ende:
        return False
    return True

def filter_rechnungen_fuer_report(rechnungen, zeitraum_start, zeitraum_ende):
    if not zeitraum_start and not zeitraum_ende:
        return list(rechnungen or [])

    out = []
    for r in (rechnungen or []):
        d = get_report_relevantes_datum(r)
        if is_im_zeitraum(d, zeitraum_start, zeitraum_ende):
            out.append(r)
    return out

def get_skonto_prozent(r):
    return to_float_safe(
        r.get("skonto_prozent")
        or r.get("Skonto_Prozent")
        or 0
    )

def get_skonto_betrag(r):
    return to_float_safe(
        r.get("skonto_betrag")
        or r.get("Skonto_Betrag")
        or 0
    )

def get_ablage_status(r):
    return str(
        r.get("ablage_status")
        or r.get("Ablage_Status")
        or ""
    ).strip().upper()

def get_referenznummer(r):
    return str(
        r.get("referenznummer")
        or r.get("Referenznummer")
        or r.get("rechnungsreferenznummer")
        or r.get("Rechnungsreferenznummer")
        or r.get("bezug_schluessel")
        or r.get("Bezug_Schluessel")
        or ""
    ).strip()

def is_abgelegt(r):
    v = get_ablage_status(r)
    if not v:
        return False
    return v not in ("OFFEN", "NICHT_ABGELEGT", "FEHLER")

def is_rechnung(r):
    return get_dokumenttyp(r) == "RECHNUNG"

def is_gutschrift(r):
    return get_dokumenttyp(r) == "GUTSCHRIFT"

def is_geprueft(r):
    return get_pruefung_status(r) == "ABGESCHLOSSEN"

def is_offen(r):
    return get_pruefung_status(r) != "ABGESCHLOSSEN"

def is_rechnung_auffaellig(r):
    return get_gesamtbewertung(r) == "HINWEIS"

def is_rechnung_unauffaellig(r):
    return get_gesamtbewertung(r) == "OK"

def is_generic_kostenstelle(value):
    raw = str(value or "").strip().upper()
    if not raw:
        return False
    normed = normalize_code(raw)
    canon_generisch = {normalize_code(x) for x in GENERISCHE_KOSTENSTELLEN}
    if raw in GENERISCHE_KOSTENSTELLEN:
        return True
    if normed in canon_generisch:
        return True
    return False

def choose_best_value(values):
    values = [str(v).strip() for v in values if str(v or "").strip()]
    if not values:
        return ""
    c = Counter(values)
    return c.most_common(1)[0][0]

def unique_nonempty(values):
    out = []
    seen = set()
    for v in values:
        s = str(v or "").strip()
        if not s:
            continue
        if s in seen:
            continue
        seen.add(s)
        out.append(s)
    return out

def determine_hinweis_klasse(h, rechnung_map, betriebskontext=None):
    betriebskontext = betriebskontext or {}
    rid = str(
        h.get("rechnung_id")
        or h.get("Rechnung_ID")
        or ""
    ).strip()

    rechnung = rechnung_map.get(rid, {})
    dokumenttyp = get_dokumenttyp(rechnung)
    hinweis_typ = canonical_hint_type(
        h.get("hinweis_typ")
        or h.get("Hinweis_Typ")
    )

    vorhandene_klasse = canonical_hint_klasse(
        h.get("hinweis_klasse")
        or h.get("Hinweis_Klasse")
    )
    if vorhandene_klasse in ("FACHLICH", "TECHNISCH"):
        return vorhandene_klasse

    if hinweis_typ in FACHLICHE_HINWEISE_BASIS:
        return "FACHLICH"

    if hinweis_typ in TECHNISCHE_HINWEISE_BASIS:
        return "TECHNISCH"

    if hinweis_typ in ("KOMMISSION_FEHLT", "KOMMISSION_UNKLAR"):
        if bool(betriebskontext.get("kommission_pruefen", True)):
            return "FACHLICH"
        return "TECHNISCH"

    if hinweis_typ in ("GUTSCHRIFT_ERKANNT", "GUTSCHRIFT_POSITION"):
        if dokumenttyp == "GUTSCHRIFT":
            return "TECHNISCH"
        return "FACHLICH"

    return "TECHNISCH"

def build_lieferanten_kontext_map(lieferanten_kontext):
    result = {}

    for row in (lieferanten_kontext or []):
        lieferant_id = str(
            row.get("lieferant_id")
            or row.get("Lieferant_ID")
            or row.get("id")
            or row.get("ID")
            or ""
        ).strip()

        lieferant_name = str(
            row.get("lieferant_name")
            or row.get("Lieferant_Name")
            or row.get("name")
            or row.get("Name")
            or ""
        ).strip()

        raw_typ = str(
            row.get("lieferant_typ")
            or row.get("Lieferant_Typ")
            or row.get("lieferantenkategorie")
            or row.get("Lieferantenkategorie")
            or row.get("kategorie")
            or row.get("Kategorie")
            or row.get("typ")
            or row.get("Typ")
            or ""
        ).strip()

        canonical_typ = canonical_lieferanten_kategorie(raw_typ)

        entry = {
            "lieferant_id": lieferant_id,
            "lieferant_name": lieferant_name,
            "lieferant_typ": canonical_typ,
            "lieferantenkategorie": canonical_typ,
            "projekt_kategorie": map_lieferant_typ_to_kategorie(canonical_typ),
            "ist_projektlevant": map_lieferant_typ_to_kategorie(canonical_typ) in PROJEKT_KATEGORIEN,
        }

        if lieferant_id:
            result[lieferant_id] = entry

        if lieferant_name:
            result[f"NAME::{normalize_name(lieferant_name)}"] = entry

    return result

def has_plausible_street_number(address_key):
    s = str(address_key or "").strip().lower()
    if not s:
        return False

    left = s.split("|")[0].strip()

    if not re.search(r"\b\d+[a-z]?\b", left):
        return False

    street_markers = [
        "str", "weg", "allee", "platz", "gasse", "ring", "stieg",
        "ufer", "pfad", "trift", "weiden", "lehnhof"
    ]
    if not any(m in left for m in street_markers):
        return False

    return True

def addresses_refer_to_same_place(a, b):
    a = str(a or "").strip()
    b = str(b or "").strip()

    if not a or not b:
        return False

    a_left, _, a_plz = a.partition("|")
    b_left, _, b_plz = b.partition("|")

    a_left = a_left.strip()
    b_left = b_left.strip()
    a_plz = a_plz.strip()
    b_plz = b_plz.strip()

    if a_plz and b_plz and a_plz != b_plz:
        return False

    a_num = re.search(r"\b(\d+[a-z]?)\b", a_left)
    b_num = re.search(r"\b(\d+[a-z]?)\b", b_left)

    if a_num and b_num and a_num.group(1) != b_num.group(1):
        return False

    sim = text_similarity(a_left, b_left)
    return sim >= 0.88
    
def is_same_betriebsadresse(a, b):
    a = str(a or "").strip()
    b = str(b or "").strip()

    if not a or not b:
        return False

    a_left, _, a_plz = a.partition("|")
    b_left, _, b_plz = b.partition("|")

    a_left = a_left.strip()
    b_left = b_left.strip()
    a_plz = a_plz.strip()
    b_plz = b_plz.strip()

    if a_plz and b_plz and a_plz != b_plz:
        return False

    a_num = re.search(r"\b(\d+[a-z]?)\b", a_left)
    b_num = re.search(r"\b(\d+[a-z]?)\b", b_left)

    if a_num and b_num and a_num.group(1) != b_num.group(1):
        return False

    sim = text_similarity(a_left, b_left)
    return sim >= 0.96

def is_plausible_project_kostenstelle(value):
    raw = str(value or "").strip()
    if not raw:
        return False

    if is_generic_kostenstelle(raw):
        return False

    cleaned = raw.upper()
    cleaned = cleaned.replace(" ", "").replace("-", "").replace("_", "").replace("/", "").replace("\\", "")

    letters = re.findall(r"[A-ZÄÖÜ]", cleaned)
    digits = re.sub(r"\D", "", cleaned)

    if letters:
        return True

    if 1 <= len(digits) <= 8:
        return True

    return False


def is_weak_project_kostenstelle(value):
    raw = str(value or "").strip()
    if not raw:
        return False

    return not is_plausible_project_kostenstelle(raw)


def extract_clean_baustelle_text(value):
    raw = str(value or "").strip()
    if not raw:
        return ""

    s = re.sub(r"\s+", " ", raw).strip()

    noise_patterns = [
        r"\bvom\b.*$",
        r"\babholung\b.*$",
        r"\bselbstabholer\b.*$",
        r"\babholer\b.*$",
        r"\blieferung\b.*$",
        r"\bupdate\b.*$",
        r"\bkommission\b.*$",
        r"\bkostenstelle\b.*$",
    ]
    for p in noise_patterns:
        s = re.sub(p, "", s, flags=re.IGNORECASE).strip()

    address_patterns = [
        r"([A-Za-zÄÖÜäöüß0-9\-\./ ]+?\b(?:str\.?|straße|strasse|weg|allee|platz|gasse|ring|ufer|pfad|stieg|trift|lehnhof|weiden)\s+\d+[A-Za-z]?(?:,\s*|\s+)\d{5}\s+[A-Za-zÄÖÜäöüß\-]+(?:\s+[A-Za-zÄÖÜäöüß\-]+)*)",
    ]

    for pattern in address_patterns:
        m = re.search(pattern, s, flags=re.IGNORECASE)
        if m:
            candidate = re.sub(r"\s+", " ", m.group(1)).strip(" ,.-")
            key = normalize_address_key(candidate)
            if has_plausible_street_number(key):
                return candidate

    key = normalize_address_key(s)
    if has_plausible_street_number(key):
        return re.sub(r"\s+", " ", s).strip(" ,.-")

    return ""

def normalize_kostenstelle_match(value):
    raw = str(value or "").strip().upper()
    if not raw:
        return ""

    raw = raw.replace(" ", "").replace("_", "").replace("/", "").replace("\\", "").replace("-", "")
    raw = raw.replace("Ä", "AE").replace("Ö", "OE").replace("Ü", "UE").replace("ß", "SS")

    if raw in {"P1", "LAGER", "STANDARD", "INTERN", "SAMMEL"}:
        return raw

    m = re.search(r"([A-Z]*)(\d{4,})", raw)
    if m:
        prefix = m.group(1)
        digits = m.group(2)

        if prefix in {"P", "PA", "S"}:
            return f"{prefix}{digits}"

        return digits

    return raw

def kostenstelle_match(a, b):
    a_norm = normalize_kostenstelle_match(a)
    b_norm = normalize_kostenstelle_match(b)

    if not a_norm or not b_norm:
        return False

    if a_norm == b_norm:
        return True

    a_digits = re.sub(r"^[A-Z]+", "", a_norm)
    b_digits = re.sub(r"^[A-Z]+", "", b_norm)

    if a_digits and b_digits and a_digits == b_digits and len(a_digits) >= 4:
        return True

    return False

def extract_project_features(r, betriebsadresse_key=""):
    kostenstelle = r.get("kostenstelle") or r.get("Kostenstelle") or ""
    kommission = r.get("kommission") or r.get("Kommission") or ""
    baustelle = r.get("baustelle") or r.get("Baustelle") or ""
    projekt_text = (
        r.get("projekt_hinweis_text")
        or r.get("Projekt_Hinweis_Text")
        or r.get("kommission_hinweis")
        or r.get("Kommission_Hinweis")
        or ""
    )

    kostenstelle_raw = str(kostenstelle or "").strip()
    kommission_raw = str(kommission or "").strip()
    baustelle_raw_original = str(baustelle or "").strip()
    baustelle_raw = extract_clean_baustelle_text(baustelle_raw_original)
    projekt_text_raw = str(projekt_text or "").strip()

    current_address_key = normalize_address_key(baustelle_raw)

    is_betriebsadresse = bool(
        betriebsadresse_key
        and current_address_key
        and is_same_betriebsadresse(current_address_key, betriebsadresse_key)
    )

    return {
        "kostenstelle_raw": kostenstelle_raw,
        "kommission_raw": kommission_raw,
        "baustelle_raw": baustelle_raw,
        "projekt_text_raw": projekt_text_raw,
        "kostenstelle_norm": normalize_code(kostenstelle_raw),
        "kostenstelle_match_key": normalize_kostenstelle_match(kostenstelle_raw),
        "kommission_norm": normalize_person_name_for_match(kommission_raw),
        "baustelle_norm": normalize_address(baustelle_raw),
        "address_key": current_address_key,
        "is_betriebsadresse": is_betriebsadresse,
        "effective_address_key": "" if is_betriebsadresse else current_address_key,
        "betriebsadresse_key": current_address_key if is_betriebsadresse else "",
        "projekt_text_norm": normalize_name(projekt_text_raw),
        "kostenstelle_generisch": is_generic_kostenstelle(kostenstelle_raw),
    }

def is_noise_kommission(value):
    s = normalize_person_name_for_match(value)
    if not s:
        return True

    bad = {
        "webshop", "abholung", "abholer", "elements", "lager",
        "innend", "aussend", "außend", "kunde", "projekt", "baustelle"
    }
    parts = [p for p in s.split() if p]
    if not parts:
        return True

    if all(p in bad for p in parts):
        return True

    if len(parts) == 1 and parts[0] in bad:
        return True

    return False

def strict_person_match(a, b):
    a = normalize_person_name_for_match(a)
    b = normalize_person_name_for_match(b)

    if not a or not b:
        return False

    a_parts = [x for x in a.split() if x]
    b_parts = [x for x in b.split() if x]

    if len(a_parts) < 2 or len(b_parts) < 2:
        return False

    if a == b:
        return True

    a_first = a_parts[0]
    b_first = b_parts[0]
    a_last = a_parts[-1]
    b_last = b_parts[-1]

    if a_last == b_last:
        if a_first == b_first:
            return True
        if a_first.startswith(b_first) or b_first.startswith(a_first):
            return True
        if SequenceMatcher(None, a_first, b_first).ratio() >= 0.80:
            return True

    sim = SequenceMatcher(None, a, b).ratio()
    return sim >= 0.90

def is_project_relevant_rechnung(r, lieferanten_kontext_map=None):
    return get_lieferant_kategorie(r, lieferanten_kontext_map) in PROJEKT_KATEGORIEN

def is_non_project_rechnung(r, lieferanten_kontext_map=None):
    return get_lieferant_kategorie(r, lieferanten_kontext_map) not in PROJEKT_KATEGORIEN

def build_project_cluster_supplier_stats(cluster_rechnungen, lieferanten_kontext_map=None):
    lieferanten_kontext_map = lieferanten_kontext_map or {}

    stats = {}

    for r in (cluster_rechnungen or []):
        lieferant_id = get_lieferant_id(r)
        lieferant_name = get_lieferant_name(r)
        dokumenttyp = get_dokumenttyp(r)
        brutto = round(get_brutto_summe(r), 2)
        netto = round(get_netto_summe(r), 2)
        kategorie = get_lieferant_kategorie(r, lieferanten_kontext_map)

        key = lieferant_id or lieferant_name or "UNBEKANNT"

        if key not in stats:
            stats[key] = {
                "lieferant_id": lieferant_id,
                "lieferant_name": lieferant_name,
                "projekt_kategorie": kategorie,
                "anzahl_dokumente": 0,
                "anzahl_rechnungen": 0,
                "anzahl_gutschriften": 0,
                "summe_brutto": 0.0,
                "summe_netto": 0.0,
                "rechnung_summe_brutto": 0.0,
                "gutschrift_summe_brutto": 0.0,
                "rechnungsnummern": [],
            }

        item = stats[key]
        item["anzahl_dokumente"] += 1
        item["summe_brutto"] = round(item["summe_brutto"] + brutto, 2)
        item["summe_netto"] = round(item["summe_netto"] + netto, 2)

        if is_rechnung(r):
            item["anzahl_rechnungen"] += 1
            item["rechnung_summe_brutto"] = round(item["rechnung_summe_brutto"] + brutto, 2)

        if is_gutschrift(r):
            item["anzahl_gutschriften"] += 1
            item["gutschrift_summe_brutto"] = round(item["gutschrift_summe_brutto"] + brutto, 2)

        rnr = get_rechnungsnummer(r)
        if rnr:
            item["rechnungsnummern"].append(rnr)

    result = list(stats.values())
    for item in result:
        item["rechnungsnummern"] = unique_nonempty(item["rechnungsnummern"])

    result.sort(key=lambda x: x.get("summe_brutto", 0), reverse=True)
    return result

def build_project_clusters(rechnungen, lieferanten_kontext_map=None, betriebsadresse_key=""):
    lieferanten_kontext_map = lieferanten_kontext_map or {}

    all_docs = list(rechnungen)
    prepared = []

    def get_cluster_pf(adjusted_kategorie):
        kat = str(adjusted_kategorie or "").strip().upper()

        if kat == "GROSSHANDEL":
            return "GROSSHANDEL"

        if kat in {"HERSTELLER", "SUBUNTERNEHMER"}:
            return "BAUSTELLE_ONLY"

        return "BETRIEBSKOSTEN"

    def is_offene_projektkosten_kategorie(kategorie):
        kat = str(kategorie or "").strip().upper()
        return kat in {"GROSSHANDEL", "HERSTELLER", "SUBUNTERNEHMER"}

    def get_effective_kostenstelle_for_cluster(feat, rechnung):
        kategorie = get_lieferant_kategorie(rechnung, lieferanten_kontext_map)
        pfad = get_cluster_pf(kategorie)

        if pfad != "GROSSHANDEL":
            return ""

        raw = str(feat.get("kostenstelle_raw") or "").strip()
        if not raw:
            return ""
        if feat.get("kostenstelle_generisch"):
            return ""
        if not is_plausible_project_kostenstelle(raw):
            return ""

        return raw

    def is_betriebskosten_kategorie(rechnung):
        kategorie = get_lieferant_kategorie(rechnung, lieferanten_kontext_map)
        return get_cluster_pf(kategorie) == "BETRIEBSKOSTEN"

    def is_lager_item(feat):
        values = [
            str(feat.get("kostenstelle_raw") or "").strip().lower(),
            str(feat.get("kommission_raw") or "").strip().lower(),
            str(feat.get("baustelle_raw") or "").strip().lower(),
            str(feat.get("projekt_text_raw") or "").strip().lower(),
        ]
        joined = " ".join(values)

        lager_marker = [
            "p1", "lager", "lagerbestand", "lagerware", "lagerartikel"
        ]
        return any(x in joined for x in lager_marker)

    def get_effective_address_key(feat):
        key = str(feat.get("effective_address_key") or "").strip()
        if not key:
            return ""
        if not has_plausible_street_number(key):
            return ""
        return key

    def get_betriebsadresse_cluster_key(feat):
        key = str(feat.get("betriebsadresse_key") or "").strip()
        if not key:
            return ""
        if not has_plausible_street_number(key):
            return ""
        return key

    def make_cluster_from_item(item, reason):
        feat = item["feat"]
        rechnung = item["rechnung"]
        brutto = get_brutto_summe(rechnung)
        netto = get_netto_summe(rechnung)

        kategorie = get_lieferant_kategorie(rechnung, lieferanten_kontext_map)
        effective_address_key = get_effective_address_key(feat)
        betriebsadresse_cluster_key = get_betriebsadresse_cluster_key(feat)
        effective_kostenstelle_raw = get_effective_kostenstelle_for_cluster(feat, rechnung)
        kostenstelle_match_key = normalize_kostenstelle_match(effective_kostenstelle_raw)


        material = brutto if (is_rechnung(rechnung) and kategorie in {"GROSSHANDEL", "HERSTELLER"}) else 0.0
        subunternehmer = brutto if (is_rechnung(rechnung) and kategorie == "SUBUNTERNEHMER") else 0.0
        sonstiges = brutto if (is_rechnung(rechnung) and kategorie not in {"GROSSHANDEL", "HERSTELLER", "SUBUNTERNEHMER"}) else 0.0

        return {
            "rechnungen": [rechnung],
            "rechnung_ids": [get_rechnung_id(rechnung)] if get_rechnung_id(rechnung) else [],
            "rechnungsnummern": [get_rechnungsnummer(rechnung)] if get_rechnungsnummer(rechnung) else [],
            "kostenstelle_values": [effective_kostenstelle_raw] if effective_kostenstelle_raw else [],
            "kommission_values": [feat["kommission_raw"]] if feat["kommission_raw"] and not is_noise_kommission(feat["kommission_raw"]) else [],
            "baustelle_values": [feat["baustelle_raw"]] if feat["baustelle_raw"] else [],
            "projekt_text_values": [feat["projekt_text_raw"]] if feat["projekt_text_raw"] else [],
            "address_keys": {effective_address_key} if effective_address_key else set(),
            "betriebsadresse_keys": {betriebsadresse_cluster_key} if betriebsadresse_cluster_key else set(),
            "kostenstelle_norms": {normalize_code(effective_kostenstelle_raw)} if effective_kostenstelle_raw else set(),
            "kostenstelle_match_keys": {kostenstelle_match_key} if kostenstelle_match_key else set(),
            "kommission_norms": {feat["kommission_norm"]} if feat["kommission_norm"] and not is_noise_kommission(feat["kommission_raw"]) else set(),
            "match_hinweise": [reason],
            "is_lager": is_lager_item(feat),
            "is_betriebsadresse_cluster": bool(betriebsadresse_cluster_key and not effective_address_key),
            "rechnung_summe_brutto": round(brutto if is_rechnung(rechnung) else 0.0, 2),
            "rechnung_summe_netto": round(netto if is_rechnung(rechnung) else 0.0, 2),
            "gutschrift_summe_brutto": round(brutto if is_gutschrift(rechnung) else 0.0, 2),
            "gutschrift_summe_netto": round(netto if is_gutschrift(rechnung) else 0.0, 2),
            "nettoeffekt_brutto": round(brutto, 2),
            "nettoeffekt_netto": round(netto, 2),
            "kostenstruktur_material_brutto": round(material, 2),
            "kostenstruktur_subunternehmer_brutto": round(subunternehmer, 2),
            "kostenstruktur_sonstiges_brutto": round(sonstiges, 2),
            "anzahl_rechnungen": 1 if is_rechnung(rechnung) else 0,
            "anzahl_gutschriften": 1 if is_gutschrift(rechnung) else 0,
            "anzahl_dokumente": 1,
        }

    def add_item_to_cluster(cluster, item, reason):
        feat = item["feat"]
        rechnung = item["rechnung"]
        brutto = get_brutto_summe(rechnung)
        netto = get_netto_summe(rechnung)

        kategorie = get_lieferant_kategorie(rechnung, lieferanten_kontext_map)
        effective_address_key = get_effective_address_key(feat)
        betriebsadresse_cluster_key = get_betriebsadresse_cluster_key(feat)
        effective_kostenstelle_raw = get_effective_kostenstelle_for_cluster(feat, rechnung)
        kostenstelle_match_key = normalize_kostenstelle_match(effective_kostenstelle_raw)

        rid = get_rechnung_id(rechnung)
        rnr = get_rechnungsnummer(rechnung)

        if rid:
            cluster["rechnung_ids"].append(rid)
        if rnr:
            cluster["rechnungsnummern"].append(rnr)

        cluster["rechnungen"].append(rechnung)

        if effective_kostenstelle_raw:
            cluster["kostenstelle_values"].append(effective_kostenstelle_raw)

        if feat["kommission_raw"] and not is_noise_kommission(feat["kommission_raw"]):
            cluster["kommission_values"].append(feat["kommission_raw"])

        if feat["baustelle_raw"]:
            cluster["baustelle_values"].append(feat["baustelle_raw"])

        if feat["projekt_text_raw"]:
            cluster["projekt_text_values"].append(feat["projekt_text_raw"])

        if effective_address_key and not feat.get("is_betriebsadresse"):
            cluster["address_keys"].add(effective_address_key)

        if betriebsadresse_cluster_key:
            cluster["betriebsadresse_keys"].add(betriebsadresse_cluster_key)

        if effective_kostenstelle_raw:
            cluster["kostenstelle_norms"].add(normalize_code(effective_kostenstelle_raw))

        if kostenstelle_match_key:
            cluster["kostenstelle_match_keys"].add(kostenstelle_match_key)

        if feat["kommission_norm"] and not is_noise_kommission(feat["kommission_raw"]):
            cluster["kommission_norms"].add(feat["kommission_norm"])

        cluster["match_hinweise"].append(reason)

        if is_rechnung(rechnung):
            cluster["rechnung_summe_brutto"] = round(cluster["rechnung_summe_brutto"] + brutto, 2)
            cluster["rechnung_summe_netto"] = round(cluster["rechnung_summe_netto"] + netto, 2)
            cluster["anzahl_rechnungen"] += 1

            if kategorie in {"GROSSHANDEL", "HERSTELLER"}:
                cluster["kostenstruktur_material_brutto"] = round(cluster["kostenstruktur_material_brutto"] + brutto, 2)
            elif kategorie == "SUBUNTERNEHMER":
                cluster["kostenstruktur_subunternehmer_brutto"] = round(cluster["kostenstruktur_subunternehmer_brutto"] + brutto, 2)
            else:
                cluster["kostenstruktur_sonstiges_brutto"] = round(cluster["kostenstruktur_sonstiges_brutto"] + brutto, 2)

        if is_gutschrift(rechnung):
            cluster["gutschrift_summe_brutto"] = round(cluster["gutschrift_summe_brutto"] + brutto, 2)
            cluster["gutschrift_summe_netto"] = round(cluster["gutschrift_summe_netto"] + netto, 2)
            cluster["anzahl_gutschriften"] += 1

        cluster["nettoeffekt_brutto"] = round(cluster["nettoeffekt_brutto"] + brutto, 2)
        cluster["nettoeffekt_netto"] = round(cluster["nettoeffekt_netto"] + netto, 2)
        cluster["anzahl_dokumente"] += 1

        if cluster.get("betriebsadresse_keys"):
            cluster["is_betriebsadresse_cluster"] = True
        if cluster.get("address_keys"):
            cluster["is_betriebsadresse_cluster"] = False

    def strong_address_match(feat, cluster):
        feat_key = get_effective_address_key(feat)

        if not feat_key or not cluster["address_keys"]:
            return False

        for existing_key in cluster["address_keys"]:
            if addresses_refer_to_same_place(feat_key, existing_key):
                return True

        return False

    def can_attach_orphan_to_cluster_by_kostenstelle(item, cluster):
        feat = item["feat"]

        if not feat["kostenstelle_raw"]:
            return ""

        if feat["kostenstelle_generisch"]:
            return ""

        for existing_value in cluster.get("kostenstelle_values", []):
            if not existing_value:
                continue
            if is_generic_kostenstelle(existing_value):
                continue
            if kostenstelle_match(feat["kostenstelle_raw"], existing_value):
                return "KOSTENSTELLE_NACHGEZOGEN"

        for existing_key in cluster.get("kostenstelle_match_keys", set()):
            if not existing_key:
                continue
            if kostenstelle_match(feat["kostenstelle_raw"], existing_key):
                return "KOSTENSTELLE_NACHGEZOGEN"

        return ""

    def can_attach_address_only_to_cluster(item, cluster):
        feat = item["feat"]

        if feat["kostenstelle_raw"]:
            return ""

        if not get_effective_address_key(feat):
            return ""

        if not cluster.get("kostenstelle_values"):
            return ""

        if strong_address_match(feat, cluster):
            return "ADRESSE_AN_BESTEHENDE_KOSTENSTELLE"

        return ""

    def can_attach_orphan_to_cluster_by_name(item, cluster):
        feat = item["feat"]

        if feat["kostenstelle_raw"] or get_effective_address_key(feat):
            return ""

        if not feat["kommission_norm"]:
            return ""

        if is_noise_kommission(feat["kommission_raw"]):
            return ""

        if not is_full_person_name(feat["kommission_raw"]):
            return ""

        for existing_value in cluster.get("kommission_values", []):
            if not existing_value:
                continue
            if strict_person_match(feat["kommission_raw"], existing_value):
                return "KOMMISSION_NACHGEZOGEN"

        for existing_norm in cluster.get("kommission_norms", set()):
            if not existing_norm:
                continue
            if strict_person_match(feat["kommission_norm"], existing_norm):
                return "KOMMISSION_NACHGEZOGEN"

        return ""

    def can_merge_orphan_clusters_by_kostenstelle(a, b):
        a_keys = {
            normalize_kostenstelle_match(x)
            for x in a.get("kostenstelle_values", [])
            if x and not is_generic_kostenstelle(x)
        }

        b_keys = {
            normalize_kostenstelle_match(x)
            for x in b.get("kostenstelle_values", [])
            if x and not is_generic_kostenstelle(x)
        }

        a_keys.discard("")
        b_keys.discard("")

        if not a_keys or not b_keys:
            return False

        return not a_keys.isdisjoint(b_keys)

    def can_merge_orphan_clusters_by_name(a, b):
        a_names = [
            x for x in a.get("kommission_values", [])
            if x and is_full_person_name(x) and not is_noise_kommission(x)
        ]
        b_names = [
            x for x in b.get("kommission_values", [])
            if x and is_full_person_name(x) and not is_noise_kommission(x)
        ]

        for av in a_names:
            for bv in b_names:
                if strict_person_match(av, bv):
                    return True

        for an in a.get("kommission_norms", set()):
            if not an:
                continue
            for bn in b.get("kommission_norms", set()):
                if not bn:
                    continue
                if strict_person_match(an, bn):
                    return True

        return False

    def get_attachable_clusters_for_betriebsadresse():
        clusters = []

        for cluster in address_clusters:
            if not cluster.get("is_betriebsadresse_cluster", False):
                clusters.append(cluster)

        if lager_cluster is not None and not lager_cluster.get("is_betriebsadresse_cluster", False):
            clusters.append(lager_cluster)

        for cluster in standalone_clusters:
            if not cluster.get("is_betriebsadresse_cluster", False):
                clusters.append(cluster)

        return clusters

    def build_betriebsadresse_rest_group_key(item):
        feat = item["feat"]

        ks_key = normalize_kostenstelle_match(feat.get("kostenstelle_raw"))
        if ks_key and not feat.get("kostenstelle_generisch", False):
            return f"KS::{ks_key}"

        person_raw = feat.get("kommission_raw") or ""
        person_norm = normalize_person_name_for_match(person_raw)
        if person_norm and is_full_person_name(person_raw) and not is_noise_kommission(person_raw):
            return f"NAME::{person_norm}"

        return "REST"

    def merge_two_clusters(a, b):
        merged = {
            "rechnungen": list(a["rechnungen"]) + list(b["rechnungen"]),
            "rechnung_ids": list(a["rechnung_ids"]) + list(b["rechnung_ids"]),
            "rechnungsnummern": list(a["rechnungsnummern"]) + list(b["rechnungsnummern"]),
            "kostenstelle_values": list(a["kostenstelle_values"]) + list(b["kostenstelle_values"]),
            "kommission_values": list(a["kommission_values"]) + list(b["kommission_values"]),
            "baustelle_values": list(a["baustelle_values"]) + list(b["baustelle_values"]),
            "projekt_text_values": list(a["projekt_text_values"]) + list(b["projekt_text_values"]),
            "address_keys": set(a["address_keys"]) | set(b["address_keys"]),
            "betriebsadresse_keys": set(a.get("betriebsadresse_keys", set())) | set(b.get("betriebsadresse_keys", set())),
            "kostenstelle_norms": set(a["kostenstelle_norms"]) | set(b["kostenstelle_norms"]),
            "kostenstelle_match_keys": set(a.get("kostenstelle_match_keys", set())) | set(b.get("kostenstelle_match_keys", set())),
            "kommission_norms": set(a["kommission_norms"]) | set(b["kommission_norms"]),
            "match_hinweise": list(a["match_hinweise"]) + list(b["match_hinweise"]) + ["CLUSTER_ZUSAMMENGEFUEHRT"],
            "is_lager": a["is_lager"] or b["is_lager"],
            "is_betriebsadresse_cluster": a.get("is_betriebsadresse_cluster", False) or b.get("is_betriebsadresse_cluster", False),
            "rechnung_summe_brutto": round(a["rechnung_summe_brutto"] + b["rechnung_summe_brutto"], 2),
            "rechnung_summe_netto": round(a["rechnung_summe_netto"] + b["rechnung_summe_netto"], 2),
            "gutschrift_summe_brutto": round(a["gutschrift_summe_brutto"] + b["gutschrift_summe_brutto"], 2),
            "gutschrift_summe_netto": round(a["gutschrift_summe_netto"] + b["gutschrift_summe_netto"], 2),
            "nettoeffekt_brutto": round(a["nettoeffekt_brutto"] + b["nettoeffekt_brutto"], 2),
            "nettoeffekt_netto": round(a["nettoeffekt_netto"] + b["nettoeffekt_netto"], 2),
            "kostenstruktur_material_brutto": round(a["kostenstruktur_material_brutto"] + b["kostenstruktur_material_brutto"], 2),
            "kostenstruktur_subunternehmer_brutto": round(a["kostenstruktur_subunternehmer_brutto"] + b["kostenstruktur_subunternehmer_brutto"], 2),
            "kostenstruktur_sonstiges_brutto": round(a["kostenstruktur_sonstiges_brutto"] + b["kostenstruktur_sonstiges_brutto"], 2),
            "anzahl_rechnungen": a["anzahl_rechnungen"] + b["anzahl_rechnungen"],
            "anzahl_gutschriften": a["anzahl_gutschriften"] + b["anzahl_gutschriften"],
            "anzahl_dokumente": a["anzahl_dokumente"] + b["anzahl_dokumente"],
        }

        if merged["address_keys"]:
            merged["is_betriebsadresse_cluster"] = False

        return merged

    def should_merge_clusters(a, b):
        if a["is_lager"] != b["is_lager"]:
            return False

        if a["is_lager"] and b["is_lager"]:
            return True

        if a.get("is_betriebsadresse_cluster") or b.get("is_betriebsadresse_cluster"):
            return False

        address_match = False
        if a["address_keys"] and b["address_keys"]:
            for ka in a["address_keys"]:
                for kb in b["address_keys"]:
                    if addresses_refer_to_same_place(ka, kb):
                        address_match = True
                        break
                if address_match:
                    break

        if not address_match:
            return False

        a_kostenstelle = choose_best_value(a["kostenstelle_values"])
        b_kostenstelle = choose_best_value(b["kostenstelle_values"])

        strong_kostenstelle_match = False
        if a_kostenstelle and b_kostenstelle:
            if is_plausible_project_kostenstelle(a_kostenstelle) and is_plausible_project_kostenstelle(b_kostenstelle):
                if kostenstelle_match(a_kostenstelle, b_kostenstelle):
                    strong_kostenstelle_match = True

        a_kommission = normalize_person_name_for_match(choose_best_value(a["kommission_values"]))
        b_kommission = normalize_person_name_for_match(choose_best_value(b["kommission_values"]))

        person_match = False
        if a_kommission and b_kommission:
            if is_full_person_name(a_kommission) and is_full_person_name(b_kommission):
                if strict_person_match(a_kommission, b_kommission):
                    person_match = True

        if strong_kostenstelle_match:
            return True

        if person_match:
            return True

        if a["address_keys"] and b["address_keys"]:
            return True

        if not a_kostenstelle and not b_kostenstelle and not a_kommission and not b_kommission:
            return True

        return False


    def build_cluster_result(cluster, idx):
        kostenstelle = choose_best_value(cluster["kostenstelle_values"])
        kommission = choose_best_value(cluster["kommission_values"])
        baustelle = choose_best_value(cluster["baustelle_values"])

        if cluster.get("is_betriebsadresse_cluster", False):
            kostenstellen_unique = []
            kostenstellen_seen = set()
            for v in cluster.get("kostenstelle_values", []):
                raw = str(v or "").strip()
                if not raw:
                    continue
                if is_generic_kostenstelle(raw):
                    continue
                k = normalize_kostenstelle_match(raw)
                if not k:
                    continue
                if k in kostenstellen_seen:
                    continue
                kostenstellen_seen.add(k)
                kostenstellen_unique.append(raw)

            namen_unique = []
            namen_seen = set()
            for v in cluster.get("kommission_values", []):
                raw = str(v or "").strip()
                if not raw:
                    continue
                if is_noise_kommission(raw):
                    continue
                if not is_full_person_name(raw):
                    continue
                n = normalize_person_name_for_match(raw)
                if not n:
                    continue
                if n in namen_seen:
                    continue
                namen_seen.add(n)
                namen_unique.append(raw)

            if len(kostenstellen_unique) > 1:
                kostenstelle = ""
            elif len(kostenstellen_unique) == 1:
                kostenstelle = kostenstellen_unique[0]

            if len(namen_unique) > 1:
                kommission = ""
            elif len(namen_unique) == 1:
                kommission = namen_unique[0]

        confidence = 0.45

        if cluster["is_lager"]:
            confidence = 0.78
        else:
            if cluster["address_keys"]:
                confidence += 0.28
            elif cluster.get("is_betriebsadresse_cluster") and baustelle:
                confidence += 0.10
            if is_plausible_project_kostenstelle(kostenstelle):
                confidence += 0.18
            if kommission and is_full_person_name(kommission):
                confidence += 0.05
            if cluster["anzahl_dokumente"] >= 2:
                confidence += 0.08
            if "CLUSTER_ZUSAMMENGEFUEHRT" in cluster["match_hinweise"]:
                confidence += 0.02
            if "ADRESSE_AN_BESTEHENDE_KOSTENSTELLE" in cluster["match_hinweise"]:
                confidence += 0.07
            if "KOSTENSTELLE_NACHGEZOGEN" in cluster["match_hinweise"]:
                confidence += 0.07
            if "KOMMISSION_NACHGEZOGEN" in cluster["match_hinweise"]:
                confidence += 0.03

        confidence = min(round(confidence, 2), 0.98)

        if confidence >= 0.82:
            status = "SICHER"
        elif confidence >= 0.65:
            status = "MITTEL"
        else:
            status = "UNSICHER"

        projekt_name_report = ""
        if "BETRIEBSKOSTEN" in cluster.get("match_hinweise", []):
            projekt_name_report = "Betriebskosten"
        elif "OFFENE_PROJEKTKOSTEN" in cluster.get("match_hinweise", []):
            projekt_name_report = "Offene Projektkosten"
        elif cluster["is_lager"]:
            projekt_name_report = "Lager"
        elif baustelle and is_plausible_project_kostenstelle(kostenstelle):
            projekt_name_report = f"{baustelle} / {kostenstelle}"
        elif baustelle:
            projekt_name_report = baustelle
        elif is_plausible_project_kostenstelle(kostenstelle):
            projekt_name_report = kostenstelle
        elif kommission and is_full_person_name(kommission):
            projekt_name_report = kommission
        else:
            projekt_name_report = f"Projektcluster {idx}"

        lieferanten_stats = build_project_cluster_supplier_stats(
            cluster_rechnungen=cluster.get("rechnungen", []),
            lieferanten_kontext_map=lieferanten_kontext_map,
        )

        return {
            "projekt_cluster_id": f"PC_{idx:04d}",
            "projekt_name_report": projekt_name_report,
            "erkannte_baustelle": baustelle,
            "erkannte_kostenstelle": kostenstelle,
            "erkannte_kostenstellen_alle": unique_nonempty(cluster.get("kostenstelle_values", [])),
            "erkannte_kostenstellen_plausibel": [
                x for x in unique_nonempty(cluster.get("kostenstelle_values", []))
                if is_plausible_project_kostenstelle(x)
            ],
            "erkannte_kommission": kommission,
            "projekt_summe_brutto": round(cluster["nettoeffekt_brutto"], 2),
            "projekt_summe_netto": round(cluster["nettoeffekt_netto"], 2),
            "rechnung_summe_brutto": round(cluster["rechnung_summe_brutto"], 2),
            "rechnung_summe_netto": round(cluster["rechnung_summe_netto"], 2),
            "gutschrift_summe_brutto": round(cluster["gutschrift_summe_brutto"], 2),
            "gutschrift_summe_netto": round(cluster["gutschrift_summe_netto"], 2),
            "nettoeffekt_brutto": round(cluster["nettoeffekt_brutto"], 2),
            "nettoeffekt_netto": round(cluster["nettoeffekt_netto"], 2),
            "anzahl_rechnungen": cluster["anzahl_rechnungen"],
            "anzahl_gutschriften": cluster["anzahl_gutschriften"],
            "anzahl_dokumente": cluster["anzahl_dokumente"],
            "confidence": confidence,
            "status": status,
            "is_lager": bool(cluster.get("is_lager", False)),
            "is_betriebskosten": "BETRIEBSKOSTEN" in cluster.get("match_hinweise", []),
            "is_offene_projektkosten": "OFFENE_PROJEKTKOSTEN" in cluster.get("match_hinweise", []),
            "zugeordnete_rechnung_ids": unique_nonempty(cluster["rechnung_ids"]),
            "zugeordnete_rechnungsnummern": unique_nonempty(cluster["rechnungsnummern"]),
            "match_hinweise": sorted(list(set(cluster["match_hinweise"]))),
            "offen_unterbestimmt": False,
            "kostenstruktur_material_brutto": round(cluster["kostenstruktur_material_brutto"], 2),
            "kostenstruktur_subunternehmer_brutto": round(cluster["kostenstruktur_subunternehmer_brutto"], 2),
            "kostenstruktur_sonstiges_brutto": round(cluster["kostenstruktur_sonstiges_brutto"], 2),
            "lieferanten_breakdown": lieferanten_stats,
        }

    for r in all_docs:
        if not is_project_relevant_rechnung(r, lieferanten_kontext_map):
            continue

        feat = extract_project_features(r, betriebsadresse_key=betriebsadresse_key)
        prepared.append({
            "rechnung": r,
            "feat": feat,
        })

    address_clusters = []
    orphan_items = []
    betriebsadresse_items = []
    offene_projektkosten_items = []
    lager_cluster = None


    for item in prepared:
        feat = item["feat"]
        rechnung = item["rechnung"]
        kategorie = get_lieferant_kategorie(rechnung, lieferanten_kontext_map)
        pfad = get_cluster_pf(kategorie)
        effective_kostenstelle_raw = get_effective_kostenstelle_for_cluster(feat, rechnung)
        effective_address_key = get_effective_address_key(feat)

        if is_lager_item(feat):
            if lager_cluster is None:
                lager_cluster = make_cluster_from_item(item, "LAGER_DIREKT")
            else:
                add_item_to_cluster(lager_cluster, item, "LAGER_DIREKT")
            continue

        if pfad == "GROSSHANDEL":
            if effective_kostenstelle_raw:
                matched = False

                for cluster in address_clusters:
                    reason = can_attach_orphan_to_cluster_by_kostenstelle(item, cluster)
                    if reason:
                        add_item_to_cluster(cluster, item, reason)
                        matched = True
                        break

                if not matched:
                    address_clusters.append(make_cluster_from_item(item, "KOSTENSTELLE_DIREKT"))
                continue

            if effective_address_key:
                matched = False

                for cluster in address_clusters:
                    reason = can_attach_address_only_to_cluster(item, cluster)
                    if reason:
                        add_item_to_cluster(cluster, item, reason)
                        matched = True
                        break

                if not matched:
                    for cluster in address_clusters:
                        if strong_address_match(feat, cluster):
                            add_item_to_cluster(cluster, item, "ADRESSE_GLEICH")
                            matched = True
                            break

                if not matched:
                    address_clusters.append(make_cluster_from_item(item, "ADRESSE_NEU"))
                continue

            offene_projektkosten_items.append(item)
            continue

        if pfad == "BAUSTELLE_ONLY":
            if effective_address_key:
                matched = False

                for cluster in address_clusters:
                    if strong_address_match(feat, cluster):
                        add_item_to_cluster(cluster, item, "ADRESSE_GLEICH")
                        matched = True
                        break

                if not matched:
                    address_clusters.append(make_cluster_from_item(item, "ADRESSE_NEU"))
                continue

            offene_projektkosten_items.append(item)
            continue

        orphan_items.append(item)


    pruef_clusters_basis = list(address_clusters)
    if lager_cluster is not None:
        pruef_clusters_basis.append(lager_cluster)
        
    offene_projektkosten_clusters = []

    remaining_after_kostenstelle = []

    for item in orphan_items:
        candidates = []

        for cluster in pruef_clusters_basis:
            reason = can_attach_orphan_to_cluster_by_kostenstelle(item, cluster)
            if reason:
                candidates.append((cluster, reason))

        if len(candidates) == 1:
            cluster, reason = candidates[0]
            add_item_to_cluster(cluster, item, reason)
        else:
            remaining_after_kostenstelle.append(item)

    remaining_after_name = []

    for item in remaining_after_kostenstelle:
        candidates = []

        for cluster in pruef_clusters_basis:
            reason = can_attach_orphan_to_cluster_by_name(item, cluster)
            if reason:
                candidates.append((cluster, reason))

        if len(candidates) == 1:
            cluster, reason = candidates[0]
            add_item_to_cluster(cluster, item, reason)
        else:
            remaining_after_name.append(item)

    for item in offene_projektkosten_items:
        offene_projektkosten_clusters.append(make_cluster_from_item(item, "OFFENE_PROJEKTKOSTEN"))

    standalone_clusters = []
    for item in remaining_after_name:
        standalone_clusters.append(make_cluster_from_item(item, "EINZELFALL"))

    changed_orphan = True
    while changed_orphan:
        changed_orphan = False
        merged = []
        used = set()

        for i in range(len(standalone_clusters)):
            if i in used:
                continue

            current = standalone_clusters[i]

            for j in range(i + 1, len(standalone_clusters)):
                if j in used:
                    continue

                other = standalone_clusters[j]

                if can_merge_orphan_clusters_by_kostenstelle(current, other):
                    current = merge_two_clusters(current, other)
                    current["match_hinweise"].append("ORPHAN_KOSTENSTELLE_MERGE")
                    used.add(j)
                    changed_orphan = True

            merged.append(current)

        standalone_clusters = merged

    changed_orphan = True
    while changed_orphan:
        changed_orphan = False
        merged = []
        used = set()

        for i in range(len(standalone_clusters)):
            if i in used:
                continue

            current = standalone_clusters[i]

            for j in range(i + 1, len(standalone_clusters)):
                if j in used:
                    continue

                other = standalone_clusters[j]

                if can_merge_orphan_clusters_by_name(current, other):
                    current = merge_two_clusters(current, other)
                    current["match_hinweise"].append("ORPHAN_NAME_MERGE")
                    used.add(j)
                    changed_orphan = True

            merged.append(current)

        standalone_clusters = merged

    remaining_betriebsadresse_after_kostenstelle = []

    for item in betriebsadresse_items:
        candidates = []

        for cluster in get_attachable_clusters_for_betriebsadresse():
            reason = can_attach_orphan_to_cluster_by_kostenstelle(item, cluster)
            if reason:
                candidates.append((cluster, "BETRIEBSADRESSE_" + reason))

        if len(candidates) == 1:
            cluster, reason = candidates[0]
            add_item_to_cluster(cluster, item, reason)
        else:
            remaining_betriebsadresse_after_kostenstelle.append(item)

    remaining_betriebsadresse_after_name = []

    for item in remaining_betriebsadresse_after_kostenstelle:
        candidates = []

        for cluster in get_attachable_clusters_for_betriebsadresse():
            reason = can_attach_orphan_to_cluster_by_name(item, cluster)
            if reason:
                candidates.append((cluster, "BETRIEBSADRESSE_" + reason))

        if len(candidates) == 1:
            cluster, reason = candidates[0]
            add_item_to_cluster(cluster, item, reason)
        else:
            remaining_betriebsadresse_after_name.append(item)

    betriebsadresse_rest_clusters = []

    for item in remaining_betriebsadresse_after_name:
        feat = item["feat"]
        bkey = get_betriebsadresse_cluster_key(feat)
        rest_group_key = build_betriebsadresse_rest_group_key(item)

        matched = False
        if bkey:
            for cluster in betriebsadresse_rest_clusters:
                existing_bkeys = cluster.get("betriebsadresse_keys", set())
                existing_group_key = cluster.get("betriebsadresse_rest_group_key", "")

                if bkey in existing_bkeys and rest_group_key == existing_group_key:
                    add_item_to_cluster(cluster, item, "BETRIEBSADRESSE_REST")
                    matched = True
                    break

        if not matched:
            new_cluster = make_cluster_from_item(item, "BETRIEBSADRESSE_REST")
            new_cluster["betriebsadresse_rest_group_key"] = rest_group_key
            betriebsadresse_rest_clusters.append(new_cluster)

    all_clusters = list(address_clusters)
    if lager_cluster is not None:
        all_clusters.append(lager_cluster)
    all_clusters.extend(offene_projektkosten_clusters)
    all_clusters.extend(standalone_clusters)
    all_clusters.extend(betriebsadresse_rest_clusters)

    changed = True
    while changed:
        changed = False
        merged_clusters = []
        used = set()

        for i in range(len(all_clusters)):
            if i in used:
                continue

            current = all_clusters[i]

            for j in range(i + 1, len(all_clusters)):
                if j in used:
                    continue

                other = all_clusters[j]
                if should_merge_clusters(current, other):
                    current = merge_two_clusters(current, other)
                    used.add(j)
                    changed = True

            merged_clusters.append(current)

        all_clusters = merged_clusters

    result = []
    for idx, cluster in enumerate(all_clusters, start=1):
        result.append(build_cluster_result(cluster, idx))

    result.sort(key=lambda x: x.get("nettoeffekt_brutto", 0), reverse=True)
    return result

def build_project_report(projekt_cluster, rechnung_lookup=None, lieferanten_kontext_map=None):
    rechnung_lookup = rechnung_lookup or []
    lieferanten_kontext_map = lieferanten_kontext_map or {}

    out = []

    for cluster in projekt_cluster:
        material = round(to_float_safe(cluster.get("kostenstruktur_material_brutto")), 2)
        subunternehmer = round(to_float_safe(cluster.get("kostenstruktur_subunternehmer_brutto")), 2)
        sonstiges = round(to_float_safe(cluster.get("kostenstruktur_sonstiges_brutto")), 2)

        projekt_name = str(cluster.get("projekt_name_report") or "").strip()
        if not projekt_name:
            projekt_name = str(cluster.get("erkannte_baustelle") or "").strip()
        if not projekt_name:
            projekt_name = str(cluster.get("erkannte_kostenstelle") or "").strip()
        if not projekt_name:
            projekt_name = str(cluster.get("erkannte_kommission") or "").strip()

        lieferanten_breakdown = list(cluster.get("lieferanten_breakdown") or [])
        lieferanten_anzahl = len(lieferanten_breakdown)

        groesster_lieferant = {}
        if lieferanten_breakdown:
            top = lieferanten_breakdown[0]
            groesster_lieferant = {
                "lieferant_id": str(top.get("lieferant_id") or "").strip(),
                "lieferant_name": str(top.get("lieferant_name") or "").strip(),
                "projekt_kategorie": str(top.get("projekt_kategorie") or "").strip(),
                "summe_brutto": round(to_float_safe(top.get("summe_brutto")), 2),
                "summe_netto": round(to_float_safe(top.get("summe_netto")), 2),
                "anzahl_dokumente": int(top.get("anzahl_dokumente") or 0),
            }

        top_3_lieferanten = []
        for item in lieferanten_breakdown[:3]:
            top_3_lieferanten.append({
                "lieferant_id": str(item.get("lieferant_id") or "").strip(),
                "lieferant_name": str(item.get("lieferant_name") or "").strip(),
                "projekt_kategorie": str(item.get("projekt_kategorie") or "").strip(),
                "summe_brutto": round(to_float_safe(item.get("summe_brutto")), 2),
                "summe_netto": round(to_float_safe(item.get("summe_netto")), 2),
                "anzahl_dokumente": int(item.get("anzahl_dokumente") or 0),
            })

        out.append({
            "projekt_name": projekt_name,
            "erkannte_baustelle": str(cluster.get("erkannte_baustelle") or "").strip(),
            "erkannte_kostenstelle": str(cluster.get("erkannte_kostenstelle") or "").strip(),
            "erkannte_kostenstellen_alle": unique_nonempty(cluster.get("erkannte_kostenstellen_alle") or []),
            "erkannte_kostenstellen_plausibel": unique_nonempty(cluster.get("erkannte_kostenstellen_plausibel") or []),
            "erkannte_kommission": str(cluster.get("erkannte_kommission") or "").strip(),
            "is_betriebskosten": bool(cluster.get("is_betriebskosten", False)),
            "is_offene_projektkosten": bool(cluster.get("is_offene_projektkosten", False)),

            "status": str(cluster.get("status") or "").strip(),
            "confidence": to_float_safe(cluster.get("confidence")),
            "anzahl_rechnungen": int(cluster.get("anzahl_rechnungen") or 0),
            "anzahl_gutschriften": int(cluster.get("anzahl_gutschriften") or 0),
            "anzahl_dokumente": int(cluster.get("anzahl_dokumente") or 0),
            "rechnung_summe_brutto": round(to_float_safe(cluster.get("rechnung_summe_brutto")), 2),
            "rechnung_summe_netto": round(to_float_safe(cluster.get("rechnung_summe_netto")), 2),
            "gutschrift_summe_brutto": round(to_float_safe(cluster.get("gutschrift_summe_brutto")), 2),
            "gutschrift_summe_netto": round(to_float_safe(cluster.get("gutschrift_summe_netto")), 2),
            "nettoeffekt_brutto": round(to_float_safe(cluster.get("nettoeffekt_brutto")), 2),
            "nettoeffekt_netto": round(to_float_safe(cluster.get("nettoeffekt_netto")), 2),
            "kostenstruktur": {
                "material_brutto": material,
                "subunternehmer_brutto": subunternehmer,
                "sonstiges_brutto": sonstiges,
            },
            "kostenstruktur_anteile": {
                "material_anteil_prozent": round((material / to_float_safe(cluster.get("nettoeffekt_brutto")) * 100), 2) if to_float_safe(cluster.get("nettoeffekt_brutto")) else 0.0,
                "subunternehmer_anteil_prozent": round((subunternehmer / to_float_safe(cluster.get("nettoeffekt_brutto")) * 100), 2) if to_float_safe(cluster.get("nettoeffekt_brutto")) else 0.0,
                "sonstiges_anteil_prozent": round((sonstiges / to_float_safe(cluster.get("nettoeffekt_brutto")) * 100), 2) if to_float_safe(cluster.get("nettoeffekt_brutto")) else 0.0,
            },
            "zugeordnete_rechnungsnummern": unique_nonempty(cluster.get("zugeordnete_rechnungsnummern") or []),
            "match_hinweise": unique_nonempty(cluster.get("match_hinweise") or []),
            "lieferanten_anzahl": lieferanten_anzahl,
            "groesster_lieferant": groesster_lieferant,
            "top_3_lieferanten": top_3_lieferanten,
            "lieferanten_breakdown": lieferanten_breakdown,
        })

    out.sort(key=lambda x: x.get("nettoeffekt_brutto", 0), reverse=True)
    return out

def build_hinweis_breakdown(hinweise):
    counter = Counter()
    for h in hinweise:
        typ = canonical_hint_type(
            h.get("hinweis_typ") or h.get("Hinweis_Typ")
        )
        counter[typ] += 1

    breakdown = {}
    for typ in KANON_HINWEIS_TYPEN:
        breakdown[typ] = counter.get(typ, 0)

    extra = {k: v for k, v in counter.items() if k not in breakdown}
    breakdown.update(extra)

    return breakdown

def build_top_lieferanten(rechnungen):
    supplier_map = defaultdict(lambda: {
        "lieferant_name": "",
        "anzahl_rechnungen": 0,
        "summe_brutto": 0.0,
        "auffaellige_rechnungen": 0,
        "gepruefte_rechnungen": 0,
    })

    for r in rechnungen:
        supplier = get_lieferant_name(r)
        brutto = get_brutto_summe(r)

        item = supplier_map[supplier]
        item["lieferant_name"] = supplier
        item["anzahl_rechnungen"] += 1
        item["summe_brutto"] += brutto

        if is_rechnung_auffaellig(r):
            item["auffaellige_rechnungen"] += 1

        if is_geprueft(r):
            item["gepruefte_rechnungen"] += 1

    result = []
    for _, v in supplier_map.items():
        v["summe_brutto"] = round(v["summe_brutto"], 2)
        result.append(v)

    result.sort(key=lambda x: x["summe_brutto"], reverse=True)
    return result[:10]

def build_payment_section(rechnungen, payment_start, payment_end):
    faellige = []
    summe_faellig = 0.0
    skonto_chancen = []

    for r in rechnungen:
        fad = get_faelligkeitsdatum(r)
        brutto = get_brutto_summe(r)
        skonto_prozent = get_skonto_prozent(r)
        skonto_betrag_raw = get_skonto_betrag(r)

        hat_echtes_skonto = False
        if skonto_prozent > 0:
            hat_echtes_skonto = True
        elif skonto_betrag_raw > 0 and brutto > 0 and skonto_betrag_raw < brutto:
            hat_echtes_skonto = True

        effektiver_skonto_betrag = round(skonto_betrag_raw, 2) if hat_echtes_skonto else 0.0

        if fad and payment_start and payment_end and payment_start <= fad <= payment_end:
            entry = {
                "rechnung_id": get_rechnung_id(r),
                "rechnungsnummer": get_rechnungsnummer(r),
                "lieferant_name": get_lieferant_name(r),
                "faelligkeitsdatum": str(fad or ""),
                "brutto_summe": round(brutto, 2),
                "skonto_betrag": effektiver_skonto_betrag,
                "skonto_prozent": round(skonto_prozent, 2) if hat_echtes_skonto else 0.0,
                "hat_skonto": hat_echtes_skonto,
            }

            faellige.append(entry)
            summe_faellig += brutto

            if hat_echtes_skonto and is_rechnung(r):
                skonto_chancen.append({
                    "rechnung_id": get_rechnung_id(r),
                    "rechnungsnummer": get_rechnungsnummer(r),
                    "lieferant_name": get_lieferant_name(r),
                    "skonto_prozent": round(skonto_prozent, 2),
                    "skonto_betrag": effektiver_skonto_betrag,
                    "brutto_summe": round(brutto, 2),
                    "faelligkeitsdatum": str(fad or ""),
                })

    faellige.sort(key=lambda x: x["brutto_summe"], reverse=True)
    skonto_chancen.sort(key=lambda x: x["skonto_betrag"], reverse=True)

    return {
        "basis_start": str(payment_start) if payment_start else "",
        "basis_ende": str(payment_end) if payment_end else "",
        "faellige_rechnungen_anzahl": len(faellige),
        "summe_faellig": round(summe_faellig, 2),
        "faellige_rechnungen": faellige[:20],
        "skonto_chancen": skonto_chancen[:20],
    }


def build_fachliche_hinweis_details(fachliche_hinweise, rechnung_map):
    details = []

    for h in fachliche_hinweise:
        rid = str(
            h.get("rechnung_id")
            or h.get("Rechnung_ID")
            or ""
        ).strip()

        r = rechnung_map.get(rid, {})

        details.append({
            "rechnung_id": rid,
            "rechnungsnummer": get_rechnungsnummer(r),
            "lieferant_name": get_lieferant_name(r),
            "dokumenttyp": get_dokumenttyp(r),
            "brutto_summe": round(get_brutto_summe(r), 2),
            "hinweis_typ": canonical_hint_type(h.get("hinweis_typ") or h.get("Hinweis_Typ")),
            "schweregrad": str(h.get("schweregrad") or h.get("Schweregrad") or "").strip().upper(),
            "kurzbeschreibung": str(h.get("kurzbeschreibung") or h.get("Kurzbeschreibung") or "").strip(),
            "aktion_empfohlen": str(h.get("aktion_empfohlen") or h.get("Aktion_Empfohlen") or "").strip().upper(),
            "bezug_schluessel": str(h.get("bezug_schluessel") or h.get("Bezug_Schluessel") or "").strip(),
            "artikelnummer": str(h.get("artikelnummer") or h.get("Artikelnummer") or "").strip(),
            "positionsbezug": str(h.get("positionsbezug") or h.get("Positionsbezug") or "").strip(),
        })

    details.sort(key=lambda x: (x["brutto_summe"], x["hinweis_typ"]), reverse=True)
    return details

def build_gutschrift_details(gutschriften, lookup_rechnungen):
    by_rechnungsnummer = {}
    by_rechnung_id = {}

    for r in lookup_rechnungen:
        nr = get_rechnungsnummer(r)
        rid = get_rechnung_id(r)

        if nr:
            by_rechnungsnummer[nr] = r
        if rid:
            by_rechnung_id[rid] = r

    details = []

    for r in gutschriften:
        referenznummer = get_referenznummer(r)
        ursprung = None

        if referenznummer:
            ursprung = by_rechnungsnummer.get(referenznummer) or by_rechnung_id.get(referenznummer)

        status = "KEINE_REFERENZ"
        if referenznummer:
            status = "REFERENZ_NICHT_GEFUNDEN"
        if ursprung:
            status = "ZUGEORDNET"

        details.append({
            "rechnung_id": get_rechnung_id(r),
            "rechnungsnummer": get_rechnungsnummer(r),
            "referenznummer": referenznummer,
            "lieferant_name": get_lieferant_name(r),
            "brutto_summe": round(get_brutto_summe(r), 2),
            "netto_summe": round(get_netto_summe(r), 2),
            "rechnungsdatum": str(get_rechnungsdatum(r) or ""),
            "faelligkeitsdatum": str(get_faelligkeitsdatum(r) or ""),
            "status": status,
            "ursprungsrechnung_rechnung_id": get_rechnung_id(ursprung) if ursprung else "",
            "ursprungsrechnung_rechnungsnummer": get_rechnungsnummer(ursprung) if ursprung else "",
            "ursprungsrechnung_rechnungsdatum": str(get_rechnungsdatum(ursprung) or "") if ursprung else "",
            "ursprungsrechnung_brutto_summe": round(get_brutto_summe(ursprung), 2) if ursprung else 0.0,
            "ursprungsrechnung_netto_summe": round(get_netto_summe(ursprung), 2) if ursprung else 0.0,
            "ursprungsrechnung_lieferant_name": get_lieferant_name(ursprung) if ursprung else "",
            "ursprungsrechnung_dokumenttyp": get_dokumenttyp(ursprung) if ursprung else "",
        })

    details.sort(key=lambda x: abs(x["brutto_summe"]), reverse=True)
    return details

def build_wichtige_faelle(fachliche_hinweis_details, gutschrift_details):
    typ_prio = {
        "PREISABWEICHUNG": 100,
        "MENGENABWEICHUNG": 95,
        "PREISSPRUNG_AUFFAELLIG": 90,
        "MWST_UNPLAUSIBEL": 88,
        "DUPLIKAT_RECHNUNG": 85,
        "GEBUEHR_AUFFAELLIG": 80,
        "GEBUEHR_POSITION": 78,
        "SKONTO_ABWEICHUNG": 75,
        "KOMMISSION_FEHLT": 70,
        "KOMMISSION_UNKLAR": 68,
        "DOPPELTE_POSITION": 65,
        "KONTO_ABWEICHUNG": 60,
        "GUTSCHRIFT_POSITION": 55,
        "GUTSCHRIFT_ERKANNT": 50,
    }

    faelle = []

    for d in fachliche_hinweis_details:
        prio = typ_prio.get(d["hinweis_typ"], 40)
        faelle.append({
            "typ": d["hinweis_typ"],
            "prioritaet": prio,
            "rechnung_id": d["rechnung_id"],
            "rechnungsnummer": d["rechnungsnummer"],
            "lieferant_name": d["lieferant_name"],
            "betrag": d["brutto_summe"],
            "kurztext": d["kurzbeschreibung"] or d["hinweis_typ"],
            "artikelnummer": d["artikelnummer"],
            "positionsbezug": d["positionsbezug"],
        })

    for g in gutschrift_details:
        if g["status"] == "ZUGEORDNET":
            faelle.append({
                "typ": "GUTSCHRIFT_ZUGEORDNET",
                "prioritaet": 45,
                "rechnung_id": g["rechnung_id"],
                "rechnungsnummer": g["rechnungsnummer"],
                "lieferant_name": g["lieferant_name"],
                "betrag": g["brutto_summe"],
                "kurztext": f"Gutschrift referenziert Ursprung {g['ursprungsrechnung_rechnungsnummer']}",
                "artikelnummer": "",
                "positionsbezug": "",
            })

    faelle.sort(key=lambda x: (x["prioritaet"], abs(x["betrag"])), reverse=True)
    return faelle[:10]

def build_project_report_meta(projekt_cluster_report):
    all_items = list(projekt_cluster_report or [])
    countable_items = [x for x in all_items if is_countable_project_cluster(x)]

    sorted_items = sorted(
        countable_items,
        key=lambda x: x.get("nettoeffekt_brutto", 0),
        reverse=True
    )

    anzahl_sicher = sum(1 for x in sorted_items if str(x.get("status") or "").upper() == "SICHER")
    anzahl_mittel = sum(1 for x in sorted_items if str(x.get("status") or "").upper() == "MITTEL")
    anzahl_unsicher = sum(1 for x in sorted_items if str(x.get("status") or "").upper() == "UNSICHER")

    unsicherste = sorted(
        list(sorted_items),
        key=lambda x: (x.get("confidence", 1), -x.get("nettoeffekt_brutto", 0))
    )[:5]

    ausgefilterte = [x for x in all_items if not is_countable_project_cluster(x)]

    return {
        "anzahl_projekte": len(sorted_items),
        "anzahl_sicher": anzahl_sicher,
        "anzahl_mittel": anzahl_mittel,
        "anzahl_unsicher": anzahl_unsicher,
        "top_3_nach_nettoeffekt_brutto": sorted_items[:3],
        "top_5_unsicherste_projekte": unsicherste,
        "ausgefilterte_cluster_anzahl": len(ausgefilterte),
        "ausgefilterte_cluster_top10": sorted(
            ausgefilterte,
            key=lambda x: x.get("nettoeffekt_brutto", 0),
            reverse=True
        )[:10],
    }


def is_countable_project_cluster(cluster):
    if not cluster:
        return False

    if bool(cluster.get("is_lager")):
        return True
    if bool(cluster.get("is_betriebskosten")):
        return False

    if bool(cluster.get("is_offene_projektkosten")):
        return False

    anzahl_dokumente = int(cluster.get("anzahl_dokumente") or 0)
    confidence = float(cluster.get("confidence") or 0)

    kostenstelle = str(cluster.get("erkannte_kostenstelle") or "").strip()
    baustelle = str(cluster.get("erkannte_baustelle") or "").strip()
    kommission = str(cluster.get("erkannte_kommission") or "").strip()

    hat_starke_kostenstelle = is_plausible_project_kostenstelle(kostenstelle)

    hat_baustelle = bool(baustelle)
    hat_vollname = bool(kommission and is_full_person_name(kommission))

    if hat_starke_kostenstelle:
        return True

    if hat_baustelle and confidence >= 0.70:
        return True

    if anzahl_dokumente >= 2 and hat_vollname and confidence >= 0.75:
        return True

    if anzahl_dokumente >= 2 and confidence >= 0.78:
        return True

    return False


def build_project_cluster_diagnostics(projekt_cluster):
    stats = {
        "cluster_gesamt": 0,
        "cluster_sicher": 0,
        "cluster_mittel": 0,
        "cluster_unsicher": 0,
        "cluster_lager": 0,
        "cluster_mit_kostenstelle": 0,
        "cluster_mit_baustelle": 0,
        "cluster_mit_kommission": 0,
        "cluster_mit_lieferanten_breakdown": 0,
        "cluster_per_kostenstelle_direkt": 0,
        "cluster_per_adresse_neu": 0,
        "cluster_per_adresse_gleich": 0,
        "cluster_mit_adresse_an_bestehende_kostenstelle": 0,
        "cluster_mit_kostenstelle_nachgezogen": 0,
        "cluster_mit_kommission_nachgezogen": 0,
        "cluster_mit_betriebsadresse_rest": 0,
        "cluster_zusammengefuehrt": 0,
        "zugeordnete_dokumente_gesamt": 0,
        "zugeordnete_rechnungen_gesamt": 0,
        "zugeordnete_gutschriften_gesamt": 0,
    }

    for cluster in (projekt_cluster or []):
        stats["cluster_gesamt"] += 1
        stats["zugeordnete_dokumente_gesamt"] += int(cluster.get("anzahl_dokumente") or 0)
        stats["zugeordnete_rechnungen_gesamt"] += int(cluster.get("anzahl_rechnungen") or 0)
        stats["zugeordnete_gutschriften_gesamt"] += int(cluster.get("anzahl_gutschriften") or 0)

        status = str(cluster.get("status") or "").upper()
        if status == "SICHER":
            stats["cluster_sicher"] += 1
        elif status == "MITTEL":
            stats["cluster_mittel"] += 1
        else:
            stats["cluster_unsicher"] += 1

        if cluster.get("is_lager"):
            stats["cluster_lager"] += 1

        if str(cluster.get("erkannte_kostenstelle") or "").strip():
            stats["cluster_mit_kostenstelle"] += 1

        if str(cluster.get("erkannte_baustelle") or "").strip():
            stats["cluster_mit_baustelle"] += 1

        if str(cluster.get("erkannte_kommission") or "").strip():
            stats["cluster_mit_kommission"] += 1

        if cluster.get("lieferanten_breakdown"):
            stats["cluster_mit_lieferanten_breakdown"] += 1

        hints = set(cluster.get("match_hinweise") or [])

        if "KOSTENSTELLE_DIREKT" in hints:
            stats["cluster_per_kostenstelle_direkt"] += 1
        if "ADRESSE_NEU" in hints:
            stats["cluster_per_adresse_neu"] += 1
        if "ADRESSE_GLEICH" in hints:
            stats["cluster_per_adresse_gleich"] += 1
        if "ADRESSE_AN_BESTEHENDE_KOSTENSTELLE" in hints:
            stats["cluster_mit_adresse_an_bestehende_kostenstelle"] += 1
        if "KOSTENSTELLE_NACHGEZOGEN" in hints:
            stats["cluster_mit_kostenstelle_nachgezogen"] += 1
        if "KOMMISSION_NACHGEZOGEN" in hints:
            stats["cluster_mit_kommission_nachgezogen"] += 1
        if "BETRIEBSADRESSE_REST" in hints:
            stats["cluster_mit_betriebsadresse_rest"] += 1
        if "CLUSTER_ZUSAMMENGEFUEHRT" in hints:
            stats["cluster_zusammengefuehrt"] += 1

    return stats

def build_non_project_supplier_summary(rechnungen_report, lieferanten_kontext_map):
    supplier_map = defaultdict(lambda: {
        "lieferant_id": "",
        "lieferant_name": "",
        "lieferanten_typ": "",
        "kosten_kategorie": "",
        "anzahl_dokumente": 0,
        "anzahl_rechnungen": 0,
        "anzahl_gutschriften": 0,
        "summe_brutto": 0.0,
        "summe_netto": 0.0,
    })

    for r in rechnungen_report:
        if is_project_relevant_rechnung(r, lieferanten_kontext_map):
            continue

        lieferant_id = get_lieferant_id(r)
        lieferant_name = get_lieferant_name(r)
        key = lieferant_id or lieferant_name or "UNBEKANNT"

        item = supplier_map[key]
        item["lieferant_id"] = lieferant_id
        item["lieferant_name"] = lieferant_name
        item["lieferanten_typ"] = get_lieferant_typ_from_kontext_map(r, lieferanten_kontext_map) or get_lieferant_typ_from_rechnung(r)
        item["kosten_kategorie"] = get_lieferant_kategorie(r, lieferanten_kontext_map)
        item["anzahl_dokumente"] += 1
        item["summe_brutto"] += get_brutto_summe(r)
        item["summe_netto"] += get_netto_summe(r)

        if is_rechnung(r):
            item["anzahl_rechnungen"] += 1
        if is_gutschrift(r):
            item["anzahl_gutschriften"] += 1

    result = []
    for _, v in supplier_map.items():
        v["summe_brutto"] = round(v["summe_brutto"], 2)
        v["summe_netto"] = round(v["summe_netto"], 2)
        result.append(v)

    result.sort(key=lambda x: abs(x["summe_brutto"]), reverse=True)
    return result

def build_betriebskosten_report(rechnungen_report, lieferanten_kontext_map, betriebsadresse_raw=""):
    kategorien_map = defaultdict(lambda: {
        "kosten_kategorie": "",
        "zugeordnete_baustelle": betriebsadresse_raw or "",
        "anzahl_dokumente": 0,
        "anzahl_rechnungen": 0,
        "anzahl_gutschriften": 0,
        "summe_brutto": 0.0,
        "summe_netto": 0.0,
        "lieferanten": set(),
    })

    lieferanten_map = defaultdict(lambda: {
        "lieferant_name": "",
        "kosten_kategorie": "",
        "zugeordnete_baustelle": betriebsadresse_raw or "",
        "anzahl_dokumente": 0,
        "anzahl_rechnungen": 0,
        "anzahl_gutschriften": 0,
        "summe_brutto": 0.0,
        "summe_netto": 0.0,
    })

    rechnungen_details = []

    for r in rechnungen_report:
        kategorie = get_lieferant_kategorie(r, lieferanten_kontext_map)

        if kategorie in {"GROSSHANDEL", "HERSTELLER", "SUBUNTERNEHMER"}:
            continue

        lieferant_name = get_lieferant_name(r)
        lieferant_id = get_lieferant_id(r)
        lieferant_key = lieferant_id or lieferant_name or "UNBEKANNT"

        brutto = get_brutto_summe(r)
        netto = get_netto_summe(r)

        kat_item = kategorien_map[kategorie]
        kat_item["kosten_kategorie"] = kategorie
        kat_item["anzahl_dokumente"] += 1
        kat_item["summe_brutto"] += brutto
        kat_item["summe_netto"] += netto
        kat_item["lieferanten"].add(lieferant_name)

        if is_rechnung(r):
            kat_item["anzahl_rechnungen"] += 1
        if is_gutschrift(r):
            kat_item["anzahl_gutschriften"] += 1

        lf_item = lieferanten_map[lieferant_key]
        lf_item["lieferant_name"] = lieferant_name
        lf_item["kosten_kategorie"] = kategorie
        lf_item["anzahl_dokumente"] += 1
        lf_item["summe_brutto"] += brutto
        lf_item["summe_netto"] += netto

        if is_rechnung(r):
            lf_item["anzahl_rechnungen"] += 1
        if is_gutschrift(r):
            lf_item["anzahl_gutschriften"] += 1

        rechnungen_details.append({
            "rechnung_id": get_rechnung_id(r),
            "rechnungsnummer": get_rechnungsnummer(r),
            "dokumenttyp": get_dokumenttyp(r),
            "lieferant_id": lieferant_id,
            "lieferant_name": lieferant_name,
            "kosten_kategorie": kategorie,
            "zugeordnete_baustelle": betriebsadresse_raw or "",
            "rechnungsdatum": str(get_rechnungsdatum(r) or ""),
            "eingangsdatum": str(get_eingangsdatum(r) or ""),
            "faelligkeitsdatum": str(get_faelligkeitsdatum(r) or ""),
            "brutto_summe": round(brutto, 2),
            "netto_summe": round(netto, 2),
            "pruefung_status": get_pruefung_status(r),
            "gesamtbewertung": get_gesamtbewertung(r),
            "ablage_status": get_ablage_status(r),
        })

    kategorien = []
    for _, v in kategorien_map.items():
        kategorien.append({
            "kosten_kategorie": v["kosten_kategorie"],
            "zugeordnete_baustelle": v["zugeordnete_baustelle"],
            "anzahl_lieferanten": len(v["lieferanten"]),
            "anzahl_dokumente": v["anzahl_dokumente"],
            "anzahl_rechnungen": v["anzahl_rechnungen"],
            "anzahl_gutschriften": v["anzahl_gutschriften"],
            "summe_brutto": round(v["summe_brutto"], 2),
            "summe_netto": round(v["summe_netto"], 2),
        })

    kategorien.sort(key=lambda x: abs(x["summe_brutto"]), reverse=True)

    lieferanten = []
    for _, v in lieferanten_map.items():
        lieferanten.append({
            "lieferant_name": v["lieferant_name"],
            "kosten_kategorie": v["kosten_kategorie"],
            "zugeordnete_baustelle": v["zugeordnete_baustelle"],
            "anzahl_dokumente": v["anzahl_dokumente"],
            "anzahl_rechnungen": v["anzahl_rechnungen"],
            "anzahl_gutschriften": v["anzahl_gutschriften"],
            "summe_brutto": round(v["summe_brutto"], 2),
            "summe_netto": round(v["summe_netto"], 2),
        })

    lieferanten.sort(key=lambda x: abs(x["summe_brutto"]), reverse=True)
    rechnungen_details.sort(key=lambda x: abs(x["brutto_summe"]), reverse=True)

    return {
        "betriebsadresse": betriebsadresse_raw or "",
        "kategorien": kategorien,
        "lieferanten_top20": lieferanten[:20],
        "rechnungen_top50": rechnungen_details[:50],
        "summe_brutto_gesamt": round(sum(x["summe_brutto"] for x in kategorien), 2),
        "summe_netto_gesamt": round(sum(x["summe_netto"] for x in kategorien), 2),
        "anzahl_kategorien": len(kategorien),
        "anzahl_lieferanten": len(lieferanten),
        "anzahl_rechnungen_gesamt": sum(x["anzahl_rechnungen"] for x in kategorien),
        "anzahl_gutschriften_gesamt": sum(x["anzahl_gutschriften"] for x in kategorien),
        "anzahl_dokumente_gesamt": sum(x["anzahl_dokumente"] for x in kategorien),
    }

    
def build_email_summary(summary, fachlicher_breakdown, payment, top_lieferanten, meta_report_type):
    parts = []

    parts.append(
        f"Im Zeitraum wurden {summary['rechnungen_gesamt']} Rechnungen mit einem Gesamtvolumen von {summary['summe_brutto_rechnungen']:.2f} verarbeitet."
    )

    parts.append(
        f"Davon waren {summary['unauffaellig']} unauffällig und {summary['auffaellig']} auffällig."
    )

    if summary["gutschriften_gesamt"] > 0:
        parts.append(f"Zusätzlich wurden {summary['gutschriften_gesamt']} Gutschriften erkannt.")

    dominante = sorted(
        [(k, v) for k, v in fachlicher_breakdown.items() if v > 0],
        key=lambda x: x[1],
        reverse=True
    )[:3]

    if dominante:
        txt = ", ".join([f"{k}: {v}" for k, v in dominante])
        parts.append(f"Die wichtigsten Hinweisarten waren {txt}.")

    if payment["faellige_rechnungen_anzahl"] > 0:
        parts.append(
            f"Im Zahlungsfenster liegen {payment['faellige_rechnungen_anzahl']} Rechnungen mit insgesamt {payment['summe_faellig']:.2f}."
        )

    if top_lieferanten:
        top = top_lieferanten[0]
        parts.append(
            f"Größter Lieferant im {meta_report_type} war {top['lieferant_name']} mit {top['summe_brutto']:.2f} Volumen."
        )

    top3_skonto = round(sum(x["skonto_betrag"] for x in payment["skonto_chancen"][:3]), 2)
    if top3_skonto > 0:
        parts.append(
            f"Zusätzlich bestehen relevante Skonto-Chancen, allein die Top-3 liegen bei {top3_skonto:.2f}."
        )

    return " ".join(parts).strip()
# ============================================================
# ANALYSE / REPORT-HELPER
# ============================================================

def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        if value in [None, "", False]:
            return default
        return float(value)
    except Exception:
        return default


def safe_int(value: Any, default: int = 0) -> int:
    try:
        if value in [None, "", False]:
            return default
        return int(float(value))
    except Exception:
        return default


def round2(value: Any) -> float:
    return round(safe_float(value), 2)


def compact_nonzero_dict(d: Dict[str, Any]) -> Dict[str, Any]:
    out = {}
    for k, v in (d or {}).items():
        if isinstance(v, (int, float)):
            if v != 0:
                out[k] = v
        elif v not in [None, "", [], {}]:
            out[k] = v
    return out


def take_top(items: List[Dict[str, Any]], limit: int = 10, sort_key: Optional[str] = None, reverse: bool = True) -> List[Dict[str, Any]]:
    data = list(items or [])
    if sort_key:
        data.sort(key=lambda x: safe_float(x.get(sort_key, 0)), reverse=reverse)
    return data[:limit]

PROJECT_RELEVANT_KATEGORIEN = {
    "GROSSHANDEL",
    "HERSTELLER",
    "SUBUNTERNEHMER",
}

NON_PROJECT_KATEGORIEN = {
    "DIENSTLEISTER",
    "WERKSTATT",
    "FIXKOSTEN",
    "ARBEITSKLEIDUNG",
    "SONSTIGES",
}


def strip_lieferant_item(x: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "lieferant_name": x.get("lieferant_name", ""),
        "lieferanten_typ": x.get("lieferanten_typ") or x.get("kosten_kategorie") or x.get("lieferantenkategorie") or "",
        "anzahl_rechnungen": safe_int(x.get("anzahl_rechnungen", 0)),
        "anzahl_gutschriften": safe_int(x.get("anzahl_gutschriften", 0)),
        "summe_brutto": round2(x.get("summe_brutto", 0)),
        "summe_netto": round2(x.get("summe_netto", 0)),
        "auffaellige_rechnungen": safe_int(x.get("auffaellige_rechnungen", 0)),
        "gepruefte_rechnungen": safe_int(x.get("gepruefte_rechnungen", 0)),
    }


def strip_payment_item(x: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "lieferant_name": x.get("lieferant_name", ""),
        "rechnungsnummer": x.get("rechnungsnummer", ""),
        "faelligkeitsdatum": x.get("faelligkeitsdatum", ""),
        "brutto_summe": round2(x.get("brutto_summe", 0)),
        "skonto_prozent": round2(x.get("skonto_prozent", 0)),
        "skonto_betrag": round2(x.get("skonto_betrag", 0)),
    }


def strip_gutschrift_item(x: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "lieferant_name": x.get("lieferant_name", ""),
        "rechnungsnummer": x.get("rechnungsnummer", ""),
        "rechnungsdatum": x.get("rechnungsdatum", ""),
        "referenznummer": x.get("referenznummer", ""),
        "status": x.get("status", ""),
        "brutto_summe": round2(x.get("brutto_summe", 0)),
        "netto_summe": round2(x.get("netto_summe", 0)),
    }


def strip_project_for_report(p: Dict[str, Any]) -> Dict[str, Any]:
    kostenstruktur = p.get("kostenstruktur") or {}

    material_brutto = kostenstruktur.get("material_brutto", p.get("kostenstruktur_material_brutto", 0))
    subunternehmer_brutto = kostenstruktur.get("subunternehmer_brutto", p.get("kostenstruktur_subunternehmer_brutto", 0))
    sonstiges_brutto = kostenstruktur.get("sonstiges_brutto", p.get("kostenstruktur_sonstiges_brutto", 0))

    top_lieferanten = []
    for l in (p.get("top_3_lieferanten") or [])[:3]:
        top_lieferanten.append({
            "lieferant_name": l.get("lieferant_name", ""),
            "projekt_kategorie": l.get("projekt_kategorie", ""),
            "anzahl_dokumente": safe_int(l.get("anzahl_dokumente", 0)),
            "summe_brutto": round2(l.get("summe_brutto", 0)),
            "summe_netto": round2(l.get("summe_netto", 0)),
        })

    return {
        "projekt_name": p.get("projekt_name") or p.get("projekt_name_report") or "",
        "status": p.get("status", ""),
        "confidence": round2(p.get("confidence", 0)),
        "erkannte_baustelle": p.get("erkannte_baustelle", ""),
        "erkannte_kostenstelle": p.get("erkannte_kostenstelle", ""),
        "erkannte_kostenstellen_alle": unique_nonempty(p.get("erkannte_kostenstellen_alle") or []),
        "erkannte_kostenstellen_plausibel": unique_nonempty(p.get("erkannte_kostenstellen_plausibel") or []),
        "erkannte_kommission": p.get("erkannte_kommission", ""),
        "is_betriebskosten": bool(p.get("is_betriebskosten", False)),
        "is_offene_projektkosten": bool(p.get("is_offene_projektkosten", False)),
        "anzahl_dokumente": safe_int(p.get("anzahl_dokumente", 0)),
        "anzahl_rechnungen": safe_int(p.get("anzahl_rechnungen", 0)),
        "anzahl_gutschriften": safe_int(p.get("anzahl_gutschriften", 0)),
        "rechnung_summe_brutto": round2(p.get("rechnung_summe_brutto", 0)),
        "rechnung_summe_netto": round2(p.get("rechnung_summe_netto", 0)),
        "gutschrift_summe_brutto": round2(p.get("gutschrift_summe_brutto", 0)),
        "gutschrift_summe_netto": round2(p.get("gutschrift_summe_netto", 0)),
        "nettoeffekt_brutto": round2(p.get("nettoeffekt_brutto", 0)),
        "nettoeffekt_netto": round2(p.get("nettoeffekt_netto", 0)),
        "kostenstruktur": {
            "material_brutto": round2(material_brutto),
            "subunternehmer_brutto": round2(subunternehmer_brutto),
            "sonstiges_brutto": round2(sonstiges_brutto),
        },
        "top_lieferanten": top_lieferanten,
        "match_hinweise": (p.get("match_hinweise") or [])[:5],
    }

def filter_countable_projects_for_report(projects: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [p for p in (projects or []) if is_countable_project_cluster(p)]

def aggregate_payment_by_lieferant(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    bucket = {}
    for x in items or []:
        name = x.get("lieferant_name", "") or "UNBEKANNT"
        if name not in bucket:
            bucket[name] = {
                "lieferant_name": name,
                "anzahl_rechnungen": 0,
                "summe_brutto": 0.0,
                "summe_skonto": 0.0,
            }
        bucket[name]["anzahl_rechnungen"] += 1
        bucket[name]["summe_brutto"] += safe_float(x.get("brutto_summe", 0))
        bucket[name]["summe_skonto"] += safe_float(x.get("skonto_betrag", 0))

    out = list(bucket.values())
    out.sort(key=lambda z: z["summe_brutto"], reverse=True)

    return [
        {
            "lieferant_name": z["lieferant_name"],
            "anzahl_rechnungen": z["anzahl_rechnungen"],
            "summe_brutto": round(z["summe_brutto"], 2),
            "summe_skonto": round(z["summe_skonto"], 2),
        }
        for z in out[:10]
    ]


def aggregate_skonto_by_lieferant(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    bucket = {}
    for x in items or []:
        name = x.get("lieferant_name", "") or "UNBEKANNT"
        if name not in bucket:
            bucket[name] = {
                "lieferant_name": name,
                "anzahl_rechnungen": 0,
                "summe_skonto": 0.0,
                "summe_brutto": 0.0,
            }
        bucket[name]["anzahl_rechnungen"] += 1
        bucket[name]["summe_skonto"] += safe_float(x.get("skonto_betrag", 0))
        bucket[name]["summe_brutto"] += safe_float(x.get("brutto_summe", 0))

    out = list(bucket.values())
    out.sort(key=lambda z: z["summe_skonto"], reverse=True)

    return [
        {
            "lieferant_name": z["lieferant_name"],
            "anzahl_rechnungen": z["anzahl_rechnungen"],
            "summe_skonto": round(z["summe_skonto"], 2),
            "summe_brutto": round(z["summe_brutto"], 2),
        }
        for z in out[:10]
    ]


def build_compact_analysis_response(
    *,
    mode: str,
    betrieb_id: str,
    zeitraum_start: Any,
    zeitraum_ende: Any,
    zeitraum: Dict[str, Any],
    payment_start: Any,
    payment_end: Any,

    summary: Dict[str, Any],

    fachliche_hinweise: List[Dict[str, Any]],
    technische_hinweise: List[Dict[str, Any]],
    fachlicher_breakdown: Dict[str, Any],
    technischer_breakdown: Dict[str, Any],
    fachliche_hinweise_by_rechnung: Dict[str, Any],
    technische_hinweise_by_rechnung: Dict[str, Any],
    fachliche_hinweis_details: List[Dict[str, Any]],

    top_lieferanten: List[Dict[str, Any]],
    non_project_lieferanten: List[Dict[str, Any]],
    lieferanten_kontext: List[Dict[str, Any]],

    gutschriften: List[Dict[str, Any]],
    gutschrift_details: List[Dict[str, Any]],

    projekt_cluster: List[Dict[str, Any]],
    projekt_cluster_report: List[Dict[str, Any]],
    projekt_report_meta: Dict[str, Any],
    unklare_projekte: List[Dict[str, Any]],

    payment: Dict[str, Any],
    wichtige_faelle: List[Dict[str, Any]],
    email_summary: str,

    project_relevante_docs: List[Dict[str, Any]],
    non_project_docs: List[Dict[str, Any]],
    projekt_cluster_diagnostics: Dict[str, Any],
    betriebskosten_report: Dict[str, Any],
) -> Dict[str, Any]:


    fachlicher_breakdown_compact = compact_nonzero_dict(fachlicher_breakdown)
    technischer_breakdown_compact = compact_nonzero_dict(technischer_breakdown)

    countable_project_clusters = filter_countable_projects_for_report(projekt_cluster_report or [])
    non_countable_project_clusters = [p for p in (projekt_cluster_report or []) if not is_countable_project_cluster(p)]

    final_report_projects = [strip_project_for_report(p) for p in countable_project_clusters]
    unklare_report_projects = [strip_project_for_report(p) for p in non_countable_project_clusters]

    offene_projektkosten_report = [
        x for x in unklare_report_projects
        if x.get("is_offene_projektkosten")
    ]

    betriebskosten_cluster_report = [
        x for x in unklare_report_projects
        if x.get("is_betriebskosten")
    ]

    top_projekte_report = take_top(
        final_report_projects,
        limit=10,
        sort_key="nettoeffekt_brutto",
        reverse=True
    )

    kritische_projekte_report = take_top(
        [p for p in final_report_projects if p.get("status") in ["MITTEL", "UNSICHER"]],
        limit=10,
        sort_key="nettoeffekt_brutto",
        reverse=True
    )

    top_lieferanten_compact = [strip_lieferant_item(x) for x in (top_lieferanten or [])[:10]]
    non_project_lieferanten_compact = [strip_lieferant_item(x) for x in (non_project_lieferanten or [])[:10]]

    gutschriften_details_compact = [strip_gutschrift_item(x) for x in (gutschrift_details or [])[:10]]

    faellige_rechnungen_compact = [strip_payment_item(x) for x in (payment.get("faellige_rechnungen") or [])[:10]]
    skonto_chancen_compact = [strip_payment_item(x) for x in (payment.get("skonto_chancen") or [])[:10]]

    payment_by_lieferant = aggregate_payment_by_lieferant(payment.get("faellige_rechnungen") or [])
    skonto_by_lieferant = aggregate_skonto_by_lieferant(payment.get("skonto_chancen") or [])

    report_summary = {
        "rechnungen_gesamt": safe_int(summary.get("rechnungen_gesamt", 0)),
        "gutschriften_gesamt": safe_int(summary.get("gutschriften_gesamt", 0)),
        "unauffaellig": safe_int(summary.get("unauffaellig", 0)),
        "auffaellig": safe_int(summary.get("auffaellig", 0)),
        "summe_brutto_rechnungen": round2(summary.get("summe_brutto_rechnungen", 0)),
        "summe_brutto_gutschriften": round2(summary.get("summe_brutto_gutschriften", 0)),
        "summe_brutto_nettoeffekt": round2(summary.get("summe_brutto_nettoeffekt", 0)),
        "summe_netto_rechnungen": round2(summary.get("summe_netto_rechnungen", 0)),
        "summe_netto_gutschriften": round2(summary.get("summe_netto_gutschriften", 0)),
        "summe_netto_nettoeffekt": round2(summary.get("summe_netto_nettoeffekt", 0)),
        "projekt_relevante_dokumente": safe_int(summary.get("projekt_relevante_dokumente", 0)),
        "nicht_projekt_relevante_dokumente": safe_int(summary.get("nicht_projekt_relevante_dokumente", 0)),
        "geprueft": safe_int(summary.get("geprueft", 0)),
        "offen": safe_int(summary.get("offen", 0)),
        "abgelegt": safe_int(summary.get("abgelegt", 0)),
        "nicht_abgelegt": safe_int(summary.get("nicht_abgelegt", 0)),
    }

    projekte_summary = {
        "anzahl_projekte": safe_int(projekt_report_meta.get("anzahl_projekte", 0)),
        "anzahl_sicher": safe_int(projekt_report_meta.get("anzahl_sicher", 0)),
        "anzahl_mittel": safe_int(projekt_report_meta.get("anzahl_mittel", 0)),
        "anzahl_unsicher": safe_int(projekt_report_meta.get("anzahl_unsicher", 0)),
        "anzahl_top_projekte_im_output": len(top_projekte_report),
        "anzahl_kritische_projekte_im_output": len(kritische_projekte_report),
        "anzahl_unklare_cluster": len(unklare_report_projects),
            "anzahl_offene_projektkosten_cluster": len(offene_projektkosten_report),
            "anzahl_betriebskosten_cluster": len(betriebskosten_cluster_report),
        "ausgefilterte_cluster_anzahl": safe_int(projekt_report_meta.get("ausgefilterte_cluster_anzahl", 0)),
    }

    hinweise_summary = {
        "fachlich_anzahl": len(fachliche_hinweise or []),
        "technisch_anzahl": len(technische_hinweise or []),
        "fachlich_breakdown": fachlicher_breakdown_compact,
        "technisch_breakdown": technischer_breakdown_compact,
    }

    gutschriften_summary = {
        "anzahl": len(gutschriften or []),
        "summe_brutto": round(sum(get_brutto_summe(r) for r in (gutschriften or [])), 2),
        "summe_netto": round(sum(get_netto_summe(r) for r in (gutschriften or [])), 2),
        "details_top10": gutschriften_details_compact,
    }

    zahlungen_summary = {
        "basis_start": str(payment_start) if payment_start else "",
        "basis_ende": str(payment_end) if payment_end else "",
        "faellige_rechnungen_anzahl": safe_int(payment.get("faellige_rechnungen_anzahl", 0)),
        "summe_faellig": round2(payment.get("summe_faellig", 0)),
        "faellige_rechnungen_top10": faellige_rechnungen_compact,
        "faellige_rechnungen_nach_lieferant_top10": payment_by_lieferant,
        "skonto_chancen_top10": skonto_chancen_compact,
        "skonto_nach_lieferant_top10": skonto_by_lieferant,
    }

    lieferanten_report = {
        "top_lieferanten": top_lieferanten_compact,
        "nicht_projekt_relevant_top10": non_project_lieferanten_compact,
    }

    projekte_report = {
        "top_projekte": top_projekte_report,
        "kritische_projekte": kritische_projekte_report,
        "unklare_projektzuordnungen": take_top(
            [
                x for x in unklare_report_projects
                if not x.get("is_betriebskosten") and not x.get("is_offene_projektkosten")
            ],
            limit=10,
            sort_key="nettoeffekt_brutto",
            reverse=True
        ),
        "offene_projektkosten": take_top(
            offene_projektkosten_report,
            limit=10,
            sort_key="nettoeffekt_brutto",
            reverse=True
        ),
        "betriebskostenblock": {
            "betriebsadresse": betriebskosten_report.get("betriebsadresse", ""),
            "summe_brutto_gesamt": round2(betriebskosten_report.get("summe_brutto_gesamt", 0)),
            "summe_netto_gesamt": round2(betriebskosten_report.get("summe_netto_gesamt", 0)),
            "anzahl_kategorien": safe_int(betriebskosten_report.get("anzahl_kategorien", 0)),
            "anzahl_lieferanten": safe_int(betriebskosten_report.get("anzahl_lieferanten", 0)),
            "anzahl_rechnungen_gesamt": safe_int(betriebskosten_report.get("anzahl_rechnungen_gesamt", 0)),
            "anzahl_gutschriften_gesamt": safe_int(betriebskosten_report.get("anzahl_gutschriften_gesamt", 0)),
            "anzahl_dokumente_gesamt": safe_int(betriebskosten_report.get("anzahl_dokumente_gesamt", 0)),
            "kategorien": betriebskosten_report.get("kategorien", [])[:10],
            "lieferanten_top20": betriebskosten_report.get("lieferanten_top20", [])[:20],
            "rechnungen_top50": betriebskosten_report.get("rechnungen_top50", [])[:50],
        },

    }

    debug_output = {
        "internal_diagnostics": {
            "technische_hinweise_anzahl": len(technische_hinweise or []),
            "rechnungen_mit_fachlichen_hinweisen": len(fachliche_hinweise_by_rechnung or {}),
            "rechnungen_mit_technischen_hinweisen": len(technische_hinweise_by_rechnung or {}),
            "project_relevante_dokumente": len(project_relevante_docs or []),
            "non_project_relevante_dokumente": len(non_project_docs or []),
            "projekt_cluster_diagnostics": projekt_cluster_diagnostics or {},
        },
        "fachliche_hinweise_details_top25": (fachliche_hinweis_details or [])[:25],
        "unklare_projektzuordnungen_top10": take_top(
            unklare_report_projects,
            limit=10,
            sort_key="nettoeffekt_brutto",
            reverse=True
        ),

        "wichtige_faelle_top10": (wichtige_faelle or [])[:10],
        "lieferanten_kontext_top50": (lieferanten_kontext or [])[:50],
        "email_summary": email_summary or "",
    }

    return {
        "ok": True,
        "meta": {
            "report_type": mode,
            "betrieb_id": betrieb_id,
            "zeitraum_start": str(zeitraum_start) if zeitraum_start else str(zeitraum.get("start") or ""),
            "zeitraum_ende": str(zeitraum_ende) if zeitraum_ende else str(zeitraum.get("ende") or ""),
            "payment_basis_start": str(payment_start) if payment_start else "",
            "payment_basis_ende": str(payment_end) if payment_end else "",
            "generated_at": datetime.utcnow().isoformat() + "Z",
            "analyze_version": "v6-compact-report-output",
        },
        "report_summary": report_summary,
        "projekte_summary": projekte_summary,
        "hinweise_summary": hinweise_summary,
        "gutschriften_summary": gutschriften_summary,
        "zahlungen_summary": zahlungen_summary,
        "lieferanten_report": lieferanten_report,
        "projekte_report": projekte_report,
        "report_highlights": {
            "email_summary": email_summary or "",
            "wichtige_faelle_top10": (wichtige_faelle or [])[:10],
        },
        "debug": debug_output,
    }


@app.route("/analyze", methods=["POST"])
def analyze():
    try:
        data = request.get_json() or {}

        meta = data.get("meta", {}) or {}

        mode = (
            meta.get("report_type")
            or data.get("mode")
            or "wochen_report"
        )

        betrieb_id = (
            meta.get("betrieb_id")
            or data.get("betrieb_id")
            or ""
        )

        zeitraum = {
            "start": meta.get("zeitraum_start") or "",
            "ende": meta.get("zeitraum_ende") or "",
        }

        rechnungen = data.get("rechnungen", []) or []
        hinweise = data.get("hinweise", []) or []

        betriebskontext = data.get("betriebskontext", {}) or {}
        betriebsadresse_raw = (
            betriebskontext.get("betriebsadresse")
            or betriebskontext.get("Betriebsadresse")
            or ""
        ).strip()

        betriebsadresse_key = normalize_address_key(betriebsadresse_raw)

        lieferanten_kontext = data.get("lieferanten_kontext", []) or []

        if "kommission_pruefen" not in betriebskontext:
            betriebskontext["kommission_pruefen"] = True

        zeitraum_start = parse_date_safe(zeitraum.get("start"))
        zeitraum_ende = parse_date_safe(zeitraum.get("ende"))

        rechnungen_report = filter_rechnungen_fuer_report(
            rechnungen=rechnungen,
            zeitraum_start=zeitraum_start,
            zeitraum_ende=zeitraum_ende,
        )

        report_rechnung_ids = {
            get_rechnung_id(r)
            for r in rechnungen_report
            if get_rechnung_id(r)
        }

        hinweise_report = []
        for h in hinweise:
            rid = str(h.get("rechnung_id") or h.get("Rechnung_ID") or "").strip()
            if rid and rid in report_rechnung_ids:
                hinweise_report.append(h)

        rechnung_map = {}
        for r in rechnungen_report:
            rid = get_rechnung_id(r)
            if rid:
                rechnung_map[rid] = r

        lookup_rechnungen = list(rechnungen)
        lieferanten_kontext_map = build_lieferanten_kontext_map(lieferanten_kontext)

        fachliche_hinweise = []
        technische_hinweise = []

        for h in hinweise_report:
            klasse = determine_hinweis_klasse(h, rechnung_map, betriebskontext=betriebskontext)
            h_copy = dict(h)
            h_copy["hinweis_klasse_effektiv"] = klasse

            if klasse == "FACHLICH":
                fachliche_hinweise.append(h_copy)
            else:
                technische_hinweise.append(h_copy)

        fachliche_hinweise_by_rechnung = defaultdict(list)
        technische_hinweise_by_rechnung = defaultdict(list)

        for h in fachliche_hinweise:
            rid = str(h.get("rechnung_id") or h.get("Rechnung_ID") or "").strip()
            if rid:
                fachliche_hinweise_by_rechnung[rid].append(h)

        for h in technische_hinweise:
            rid = str(h.get("rechnung_id") or h.get("Rechnung_ID") or "").strip()
            if rid:
                technische_hinweise_by_rechnung[rid].append(h)

        rechnungen_nur = [r for r in rechnungen_report if is_rechnung(r)]
        gutschriften = [r for r in rechnungen_report if is_gutschrift(r)]

        project_relevante_docs = [
            r for r in rechnungen_report
            if is_project_relevant_rechnung(r, lieferanten_kontext_map)
        ]

        non_project_docs = [
            r for r in rechnungen_report
            if not is_project_relevant_rechnung(r, lieferanten_kontext_map)
        ]

        today_base = datetime.utcnow().date()
        payment_start = today_base
        payment_end = today_base
        if zeitraum_ende:
            try:
                delta_days = max((zeitraum_ende - zeitraum_start).days, 0) if zeitraum_start else 7
                delta_days = min(max(delta_days, 0), 31)
                payment_end = today_base + timedelta(days=delta_days)
            except Exception:
                payment_end = today_base + timedelta(days=7)
        else:
            payment_end = today_base + timedelta(days=7)

        summary = {
            "rechnungen_gesamt": len(rechnungen_nur),
            "gutschriften_gesamt": len(gutschriften),
            "geprueft": sum(1 for r in rechnungen_report if is_geprueft(r)),
            "offen": sum(1 for r in rechnungen_report if is_offen(r)),
            "auffaellig": sum(1 for r in rechnungen_nur if is_rechnung_auffaellig(r)),
            "unauffaellig": sum(1 for r in rechnungen_nur if is_rechnung_unauffaellig(r)),
            "abgelegt": sum(1 for r in rechnungen_report if is_abgelegt(r)),
            "nicht_abgelegt": sum(1 for r in rechnungen_report if not is_abgelegt(r)),
            "summe_brutto_rechnungen": round(sum(get_brutto_summe(r) for r in rechnungen_nur), 2),
            "summe_netto_rechnungen": round(sum(get_netto_summe(r) for r in rechnungen_nur), 2),
            "summe_brutto_gutschriften": round(sum(get_brutto_summe(r) for r in gutschriften), 2),
            "summe_netto_gutschriften": round(sum(get_netto_summe(r) for r in gutschriften), 2),
            "summe_brutto_nettoeffekt": round(sum(get_brutto_summe(r) for r in rechnungen_report), 2),
            "summe_netto_nettoeffekt": round(sum(get_netto_summe(r) for r in rechnungen_report), 2),
            "hinweise_fachlich_gesamt": len(fachliche_hinweise),
            "hinweise_technisch_gesamt": len(technische_hinweise),
            "projekt_relevante_dokumente": len(project_relevante_docs),
            "nicht_projekt_relevante_dokumente": len(non_project_docs),
        }

        fachlicher_breakdown = build_hinweis_breakdown(fachliche_hinweise)
        technischer_breakdown = build_hinweis_breakdown(technische_hinweise)

        payment = build_payment_section(rechnungen_nur, payment_start, payment_end)
        top_lieferanten = build_top_lieferanten(rechnungen_report)

        projekt_cluster = build_project_clusters(
            rechnungen=rechnungen_report,
            lieferanten_kontext_map=lieferanten_kontext_map,
            betriebsadresse_key=betriebsadresse_key,
        )

        projekt_cluster_report = build_project_report(
            projekt_cluster=projekt_cluster,
            rechnung_lookup=lookup_rechnungen,
            lieferanten_kontext_map=lieferanten_kontext_map,
        )

        projekt_report_meta = build_project_report_meta(projekt_cluster_report)
        projekt_cluster_diagnostics = build_project_cluster_diagnostics(projekt_cluster)

        unklare_projekte = [
            x for x in projekt_cluster
            if x.get("confidence", 0) < 0.75
        ][:20]

        fachliche_hinweis_details = build_fachliche_hinweis_details(fachliche_hinweise, rechnung_map)
        gutschrift_details = build_gutschrift_details(gutschriften, lookup_rechnungen)
        wichtige_faelle = build_wichtige_faelle(fachliche_hinweis_details, gutschrift_details)

        email_summary = build_email_summary(
            summary=summary,
            fachlicher_breakdown=fachlicher_breakdown,
            payment=payment,
            top_lieferanten=top_lieferanten,
            meta_report_type=mode
        )

        non_project_lieferanten = build_non_project_supplier_summary(
            rechnungen_report=rechnungen_report,
            lieferanten_kontext_map=lieferanten_kontext_map,
        )

        betriebskosten_report = build_betriebskosten_report(
            rechnungen_report=rechnungen_report,
            lieferanten_kontext_map=lieferanten_kontext_map,
            betriebsadresse_raw=betriebsadresse_raw,
        )

        response_payload = build_compact_analysis_response(
            mode=mode,
            betrieb_id=betrieb_id,
            zeitraum_start=zeitraum_start,
            zeitraum_ende=zeitraum_ende,
            zeitraum=zeitraum,

            payment_start=payment_start,
            payment_end=payment_end,

            summary=summary,

            fachliche_hinweise=fachliche_hinweise,
            technische_hinweise=technische_hinweise,
            fachlicher_breakdown=fachlicher_breakdown,
            technischer_breakdown=technischer_breakdown,
            fachliche_hinweise_by_rechnung=fachliche_hinweise_by_rechnung,
            technische_hinweise_by_rechnung=technische_hinweise_by_rechnung,
            fachliche_hinweis_details=fachliche_hinweis_details,

            top_lieferanten=top_lieferanten,
            non_project_lieferanten=non_project_lieferanten,
            betriebskosten_report=betriebskosten_report,
            lieferanten_kontext=lieferanten_kontext,

            gutschriften=gutschriften,
            gutschrift_details=gutschrift_details,

            projekt_cluster=projekt_cluster,
            projekt_cluster_report=projekt_cluster_report,
            projekt_report_meta=projekt_report_meta,
            unklare_projekte=unklare_projekte,

            payment=payment,
            wichtige_faelle=wichtige_faelle,
            email_summary=email_summary,

            project_relevante_docs=project_relevante_docs,
            non_project_docs=non_project_docs,
            projekt_cluster_diagnostics=projekt_cluster_diagnostics,
        )

        return jsonify(response_payload), 200

    except Exception as e:
        return jsonify({
            "ok": False,
            "error": "analyze_failed",
            "error_detail": str(e)
        }), 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
