from flask import Flask, request, jsonify
import fitz  # PyMuPDF
import pdfplumber
import pytesseract
from PIL import Image, ImageOps, ImageFilter
import io
import os
import re
from typing import List, Dict, Any, Optional, Tuple
from collections import Counter, defaultdict
from datetime import datetime, date, timezone
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

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(1.8, 1.8), alpha=False)
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
                    timeout=8
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
            except BaseException as e:
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
        
REPORT_HINT_TYPES = [
    "ARTIKELNUMMER_FEHLT",
    "ARTIKELNUMMER_UNGUELTIG",
    "BETRIEB_NICHT_ERKANNT",
    "DOPPELTE_POSITION",
    "DUPLIKAT_RECHNUNG",
    "EINHEIT_ABWEICHUNG",
    "EXTRAKTION_FEHLER",
    "GEBUEHR_ERKANNT",
    "GEBUEHR_POSITION",
    "GUTSCHRIFT_ERKANNT",
    "GUTSCHRIFT_POSITION",
    "JSON_UNVOLLSTAENDIG",
    "KOMMISSION_FEHLT",
    "KOMMISSION_UNKLAR",
    "KONTO_ABWEICHUNG",
    "LIEFERANT_AUTO_ERSTELLT",
    "LIEFERANT_NICHT_ERKANNT",
    "MENGENABWEICHUNG",
    "MWST_UNPLAUSIBEL",
    "PREISABWEICHUNG",
    "PREISSPRUNG_AUFFAELLIG",
    "RECHNUNGSNUMMER_FEHLT",
    "SKONTO_ABWEICHUNG",
    "SKONTO_ERKANNT",
]

HINT_PRIORITY = {
    "PREISABWEICHUNG": 100,
    "MENGENABWEICHUNG": 95,
    "PREISSPRUNG_AUFFAELLIG": 90,
    "GEBUEHR_ERKANNT": 85,
    "GEBUEHR_POSITION": 80,
    "GUTSCHRIFT_ERKANNT": 78,
    "GUTSCHRIFT_POSITION": 78,
    "SKONTO_ABWEICHUNG": 75,
    "KOMMISSION_UNKLAR": 72,
    "KOMMISSION_FEHLT": 70,
    "DUPLIKAT_RECHNUNG": 68,
    "DOPPELTE_POSITION": 66,
    "MWST_UNPLAUSIBEL": 64,
    "LIEFERANT_NICHT_ERKANNT": 62,
    "BETRIEB_NICHT_ERKANNT": 60,
    "JSON_UNVOLLSTAENDIG": 55,
    "EXTRAKTION_FEHLER": 55,
    "ARTIKELNUMMER_UNGUELTIG": 40,
    "ARTIKELNUMMER_FEHLT": 35,
    "SKONTO_ERKANNT": 20,
    "LIEFERANT_AUTO_ERSTELLT": 15,
    "KONTO_ABWEICHUNG": 10,
    "EINHEIT_ABWEICHUNG": 10,
}

GENERIC_COSTCENTER_VALUES = {
    "", "0", "1", "p1", "p-1", "projekt", "lager", "test", "kaju"
}


def strip_accents(text):
    text = str(text or "")
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(ch)
    )


def normalize_spaces(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def normalize_basic(text):
    text = strip_accents(text).lower().strip()
    text = text.replace("–", "-").replace("—", "-")
    text = text.replace("_", " ")
    text = re.sub(r"[;,|]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_name_report(text):
    text = normalize_basic(text)
    if not text:
        return ""
    text = re.sub(r"\bauftr\.?text\b[: ]*", " ", text)
    text = re.sub(r"\bauftr\.?nr\b[: ]*", " ", text)
    text = re.sub(r"\binnend\.?\b", " ", text)
    text = re.sub(r"\baußend\.?\b", " ", text)
    text = re.sub(r"\baussend\.?\b", " ", text)
    text = re.sub(r"\bwebshop\b", "webshop", text)
    text = re.sub(r"\bprojektleiter\b[: ]*", " ", text)
    text = re.sub(r"\bbearbeiter\b[: ]*", " ", text)
    text = re.sub(r"\bmonteur\b[: ]*", " ", text)
    text = re.sub(r"\bserviceauftrag\b[: ]*\d+\b", " ", text)
    text = re.sub(r"\bauftragsnummer\b[: ]*\d+\b", " ", text)
    text = re.sub(r"\btechnische einheit\b[: ]*[\wäöüß\- ]+", " ", text)
    text = re.sub(r"\bdurchgefuhrte arbeiten\b[: ]*.*", " ", text)
    text = re.sub(r"\bausfuhrdatum\b[: ]*[\d\-.]+", " ", text)
    text = re.sub(r"[^a-z0-9äöüß\- ]", " ", text)
    text = normalize_spaces(text)
    return text


def normalize_address_report(text):
    text = normalize_basic(text)
    if not text:
        return ""

    text = text.replace("strasse", "str")
    text = text.replace("straße", "str")
    text = text.replace("str.", "str")
    text = re.sub(r"\bdeutschland\b", " ", text)
    text = re.sub(r"\bde\b", " ", text)
    text = re.sub(r"\bjuni gebaudetechnik gmbh\b", " ", text)
    text = re.sub(r"\bcornelius gellert str 104\b", " ", text)
    text = re.sub(r"\bniestetal\b", " niestetal ", text)
    text = re.sub(r"\binnend\.?\b", " ", text)
    text = re.sub(r"\baußend\.?\b", " ", text)
    text = re.sub(r"\baussend\.?\b", " ", text)
    text = re.sub(r"\bobjekt\b[: ]*", " ", text)
    text = re.sub(r"\bfur das objekt\b[: ]*", " ", text)
    text = re.sub(r"\bfuer das objekt\b[: ]*", " ", text)
    text = re.sub(r"\bin\b", " ", text)
    text = re.sub(r"[^a-z0-9äöüß\- ]", " ", text)
    text = re.sub(r"(\d+)\s*([a-z])\b", r"\1\2", text)
    text = normalize_spaces(text)
    return text


def normalize_costcenter_report(text):
    text = normalize_basic(text)
    if not text:
        return ""
    text = re.sub(r"[^a-z0-9\-_/]", "", text)
    text = text.upper()
    return text


def is_generic_costcenter(text):
    value = normalize_costcenter_report(text).lower()
    if value in GENERIC_COSTCENTER_VALUES:
        return True
    if re.fullmatch(r"p\d{1,2}", value or ""):
        return True
    if re.fullmatch(r"s\d{1,3}", value or ""):
        return True
    return False


def parse_float_de(value):
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return 0.0
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    text = text.replace("€", "").replace("%", "").replace(" ", "")
    text = text.replace("\u00A0", "")

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        pass

    text = re.sub(r"[^0-9\.\-]", "", text)
    if text in {"", "-", ".", "-."}:
        return 0.0

    try:
        return float(text)
    except Exception:
        return 0.0


def parse_int_safe(value):
    try:
        return int(float(str(value).replace(",", ".").strip()))
    except Exception:
        return 0


def parse_iso_date_flexible(value):
    if not value:
        return None

    value = str(value).strip()
    if not value:
        return None

    if len(value) == 10 and value.count("-") == 2:
        try:
            return datetime.strptime(value, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        except Exception:
            return None

    if value.endswith("Z"):
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        except Exception:
            pass

    try:
        dt = datetime.fromisoformat(value)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return None


def is_in_period(value, start_dt, end_dt):
    dt = parse_iso_date_flexible(value)
    if not dt:
        return False
    return start_dt <= dt <= end_dt


def detect_document_type(invoice):
    raw_type = normalize_basic(invoice.get("dokumenttyp"))
    brutto = parse_float_de(invoice.get("brutto_summe"))

    if "gutschrift" in raw_type:
        return "GUTSCHRIFT"
    if "lastschrift" in raw_type:
        return "LASTSCHRIFTAVIS"
    if brutto < 0:
        return "GUTSCHRIFT"
    return "RECHNUNG"


def normalize_supplier_name(text):
    text = normalize_basic(text)
    text = re.sub(r"\blieferant\b[_ ]*", "", text)
    text = normalize_spaces(text).upper()
    return text or "UNBEKANNT"


def normalize_project_text(text):
    text = normalize_basic(text)
    if not text:
        return ""
    text = text.replace("auftr.text", "auftrtext")
    text = text.replace("auftr.nr", "auftrnr")
    text = re.sub(r"[^a-z0-9äöüß\- ]", " ", text)
    text = normalize_spaces(text)
    return text


def extract_name_from_project_text(text):
    text = normalize_project_text(text)
    if not text:
        return ""

    patterns = [
        r"auftrtext[: ]+([a-zäöüß\- ]{3,80})",
        r"auftrnr[: ]+([a-z0-9\-_/]{2,40})",
    ]
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            value = normalize_name_report(m.group(1))
            if value:
                return value

    return ""


def build_invoice_record(raw):
    invoice = dict(raw or {})

    invoice["rechnung_id"] = str(invoice.get("rechnung_id") or "").strip()
    invoice["rechnungsnummer"] = str(invoice.get("rechnungsnummer") or "").strip()
    invoice["rechnungsdatum"] = str(invoice.get("rechnungsdatum") or "").strip()
    invoice["eingangsdatum"] = str(invoice.get("eingangsdatum") or "").strip()
    invoice["lieferant_id"] = str(invoice.get("lieferant_id") or "").strip()
    invoice["lieferant_name"] = normalize_supplier_name(invoice.get("lieferant_name"))
    invoice["brutto_summe"] = parse_float_de(invoice.get("brutto_summe"))
    invoice["netto_summe"] = parse_float_de(invoice.get("netto_summe"))
    invoice["dokumenttyp"] = detect_document_type(invoice)
    invoice["gesamtbewertung"] = normalize_basic(invoice.get("gesamtbewertung")).upper() or "OK"
    invoice["hinweise_anzahl"] = parse_int_safe(invoice.get("hinweise_anzahl"))
    invoice["faelligkeitsdatum"] = str(invoice.get("faelligkeitsdatum") or "").strip()
    invoice["skonto_prozent"] = parse_float_de(invoice.get("skonto_prozent"))
    invoice["skonto_betrag"] = parse_float_de(invoice.get("skonto_betrag"))
    invoice["kommission"] = normalize_spaces(invoice.get("kommission"))
    invoice["kostenstelle"] = normalize_spaces(invoice.get("kostenstelle"))
    invoice["baustelle"] = normalize_spaces(invoice.get("baustelle"))
    invoice["projekt_hinweis_text"] = normalize_spaces(invoice.get("projekt_hinweis_text"))

    invoice["_kommission_norm"] = normalize_name_report(invoice["kommission"]) or extract_name_from_project_text(invoice["projekt_hinweis_text"])
    invoice["_kostenstelle_norm"] = normalize_costcenter_report(invoice["kostenstelle"])
    invoice["_baustelle_norm"] = normalize_address_report(invoice["baustelle"] or invoice["projekt_hinweis_text"])
    invoice["_projekt_text_norm"] = normalize_project_text(invoice["projekt_hinweis_text"])

    invoice["_severity_seed"] = 0
    return invoice


def build_hint_record(raw):
    hint = dict(raw or {})
    hint["betrieb_id"] = str(hint.get("betrieb_id") or "").strip()
    hint["rechnung_id"] = str(hint.get("rechnung_id") or "").strip()
    hint["hinweis_typ"] = str(hint.get("hinweis_typ") or "").strip().upper()
    hint["schweregrad"] = normalize_basic(hint.get("schweregrad"))
    hint["status"] = normalize_basic(hint.get("status"))
    hint["erkannt_am"] = str(hint.get("erkannt_am") or "").strip()
    hint["bezug_typ"] = str(hint.get("bezug_typ") or "").strip()
    hint["bezug_schluessel"] = str(hint.get("bezug_schluessel") or "").strip()
    hint["kurzbeschreibung"] = normalize_spaces(hint.get("kurzbeschreibung"))
    hint["aktion_empfohlen"] = normalize_spaces(hint.get("aktion_empfohlen"))
    return hint


def dedupe_invoices(invoices):
    grouped = defaultdict(list)

    for inv in invoices:
        key = (
            inv.get("lieferant_name", ""),
            inv.get("rechnungsnummer", ""),
            round(abs(inv.get("brutto_summe", 0.0)), 2),
            inv.get("dokumenttyp", ""),
            inv.get("rechnungsdatum", "") or inv.get("eingangsdatum", ""),
        )
        grouped[key].append(inv)

    deduped = []
    for _, items in grouped.items():
        if len(items) == 1:
            deduped.append(items[0])
            continue

        def score_invoice(x):
            completeness = sum(
                1 for field in [
                    x.get("rechnung_id"),
                    x.get("rechnungsnummer"),
                    x.get("rechnungsdatum"),
                    x.get("eingangsdatum"),
                    x.get("lieferant_name"),
                    x.get("faelligkeitsdatum"),
                    x.get("kommission"),
                    x.get("kostenstelle"),
                    x.get("baustelle"),
                    x.get("projekt_hinweis_text"),
                ] if str(field or "").strip()
            )
            return (
                completeness,
                len(str(x.get("projekt_hinweis_text") or "")),
                len(str(x.get("baustelle") or "")),
                len(str(x.get("kommission") or "")),
                len(str(x.get("rechnung_id") or "")),
            )

        best = sorted(items, key=score_invoice, reverse=True)[0]
        deduped.append(best)

    return deduped


def merge_duplicate_project_clusters(clusters):
    merged = []
    used = set()

    for i, base in enumerate(clusters):
        if i in used:
            continue

        current = dict(base)
        current["zugeordnete_rechnung_ids"] = list(base["zugeordnete_rechnung_ids"])
        current["match_hinweise"] = list(base["match_hinweise"])

        for j in range(i + 1, len(clusters)):
            if j in used:
                continue
            other = clusters[j]

            same_strong_costcenter = (
                current.get("_kostenstelle_norm")
                and other.get("_kostenstelle_norm")
                and current["_kostenstelle_norm"] == other["_kostenstelle_norm"]
                and not is_generic_costcenter(current["_kostenstelle_norm"])
            )
            same_address = (
                current.get("_baustelle_norm")
                and other.get("_baustelle_norm")
                and current["_baustelle_norm"] == other["_baustelle_norm"]
            )
            same_name = (
                current.get("_kommission_norm")
                and other.get("_kommission_norm")
                and (
                    current["_kommission_norm"] == other["_kommission_norm"]
                    or current["_kommission_norm"] in other["_kommission_norm"]
                    or other["_kommission_norm"] in current["_kommission_norm"]
                )
            )

            should_merge = False
            if same_strong_costcenter and (same_address or same_name):
                should_merge = True
            elif same_address and same_name:
                should_merge = True
            elif same_address and current.get("_kostenstelle_norm") == other.get("_kostenstelle_norm") and current.get("_kostenstelle_norm"):
                should_merge = True

            if should_merge:
                used.add(j)
                current["zugeordnete_rechnung_ids"] = list(
                    dict.fromkeys(current["zugeordnete_rechnung_ids"] + other["zugeordnete_rechnung_ids"])
                )
                current["projekt_summe_brutto"] += other["projekt_summe_brutto"]
                current["anzahl_rechnungen"] = len(current["zugeordnete_rechnung_ids"])
                current["match_hinweise"] = list(
                    dict.fromkeys(current["match_hinweise"] + other["match_hinweise"] + ["NACHMERGE"])
                )
                current["confidence"] = max(current["confidence"], other["confidence"])

                if not current.get("erkannte_kostenstelle") and other.get("erkannte_kostenstelle"):
                    current["erkannte_kostenstelle"] = other["erkannte_kostenstelle"]
                    current["_kostenstelle_norm"] = other.get("_kostenstelle_norm", "")
                if not current.get("erkannte_baustelle") and other.get("erkannte_baustelle"):
                    current["erkannte_baustelle"] = other["erkannte_baustelle"]
                    current["_baustelle_norm"] = other.get("_baustelle_norm", "")
                if not current.get("erkannte_kommission") and other.get("erkannte_kommission"):
                    current["erkannte_kommission"] = other["erkannte_kommission"]
                    current["_kommission_norm"] = other.get("_kommission_norm", "")

        merged.append(current)

    return merged


def build_project_clusters_report(invoices):
    clusters = []
    cluster_index = 1

    for inv in invoices:
        if inv.get("dokumenttyp") not in {"RECHNUNG", "GUTSCHRIFT"}:
            continue

        name_norm = inv.get("_kommission_norm", "")
        cost_norm = inv.get("_kostenstelle_norm", "")
        addr_norm = inv.get("_baustelle_norm", "")
        brutto = inv.get("brutto_summe", 0.0)

        best_cluster_idx = None
        best_score = -1
        best_reasons = []

        for idx, cl in enumerate(clusters):
            score = 0
            reasons = []

            cl_cost = cl.get("_kostenstelle_norm", "")
            cl_addr = cl.get("_baustelle_norm", "")
            cl_name = cl.get("_kommission_norm", "")

            if cost_norm and cl_cost and cost_norm == cl_cost:
                if is_generic_costcenter(cost_norm):
                    score += 20
                    reasons.append("KOSTENSTELLE_GENERISCH_GLEICH")
                else:
                    score += 65
                    reasons.append("KOSTENSTELLE_GLEICH")

            if addr_norm and cl_addr and addr_norm == cl_addr:
                score += 45
                reasons.append("BAUSTELLE_GLEICH")

            if name_norm and cl_name:
                if name_norm == cl_name:
                    score += 35
                    reasons.append("KOMMISSION_GLEICH")
                elif name_norm in cl_name or cl_name in name_norm:
                    score += 22
                    reasons.append("KOMMISSION_AEHNLICH")

            if cost_norm and cl_cost and cost_norm == cl_cost and addr_norm and cl_addr and addr_norm != cl_addr:
                reasons.append("BAUSTELLE_ABWEICHEND_ABER_KOSTENSTELLE_STARK")

            if score > best_score:
                best_score = score
                best_cluster_idx = idx
                best_reasons = reasons

        should_attach = best_score >= 60 or (
            best_score >= 45 and addr_norm and name_norm
        )

        if should_attach and best_cluster_idx is not None:
            cl = clusters[best_cluster_idx]
            cl["zugeordnete_rechnung_ids"].append(inv["rechnung_id"])
            cl["projekt_summe_brutto"] += brutto
            cl["anzahl_rechnungen"] = len(cl["zugeordnete_rechnung_ids"])
            cl["match_hinweise"] = list(dict.fromkeys(cl["match_hinweise"] + best_reasons))

            if not cl.get("erkannte_kostenstelle") and inv.get("kostenstelle"):
                cl["erkannte_kostenstelle"] = inv["kostenstelle"]
                cl["_kostenstelle_norm"] = cost_norm
            if not cl.get("erkannte_baustelle") and inv.get("baustelle"):
                cl["erkannte_baustelle"] = inv["baustelle"]
                cl["_baustelle_norm"] = addr_norm
            if not cl.get("erkannte_kommission") and inv.get("kommission"):
                cl["erkannte_kommission"] = inv["kommission"]
                cl["_kommission_norm"] = name_norm

            cl["confidence"] = min(0.98, max(cl["confidence"], 0.5 + (best_score / 120.0)))
        else:
            confidence = 0.5
            if cost_norm and not is_generic_costcenter(cost_norm):
                confidence += 0.3
            if addr_norm:
                confidence += 0.05
            if name_norm:
                confidence += 0.1
            confidence = min(confidence, 0.98)

            clusters.append({
                "projekt_cluster_id": f"PC_{cluster_index:04d}",
                "erkannte_kostenstelle": inv.get("kostenstelle") or "",
                "erkannte_kommission": inv.get("kommission") or "",
                "erkannte_baustelle": inv.get("baustelle") or "",
                "projekt_summe_brutto": brutto,
                "anzahl_rechnungen": 1,
                "confidence": confidence,
                "match_hinweise": ["NEUES_CLUSTER"],
                "status": "sicher" if confidence >= 0.8 else "mittel" if confidence >= 0.55 else "unsicher",
                "offen_unterbestimmt": not bool(cost_norm or addr_norm or name_norm),
                "zugeordnete_rechnung_ids": [inv["rechnung_id"]],
                "_kostenstelle_norm": cost_norm,
                "_kommission_norm": name_norm,
                "_baustelle_norm": addr_norm,
            })
            cluster_index += 1

    clusters = merge_duplicate_project_clusters(clusters)

    for cl in clusters:
        cl["anzahl_rechnungen"] = len(cl["zugeordnete_rechnung_ids"])
        cl["match_hinweise"] = list(dict.fromkeys(cl["match_hinweise"]))
        cl["status"] = "sicher" if cl["confidence"] >= 0.8 else "mittel" if cl["confidence"] >= 0.55 else "unsicher"

    def cluster_sort_key(c):
        return (c["projekt_summe_brutto"], c["anzahl_rechnungen"], c["confidence"])

    clusters = sorted(clusters, key=cluster_sort_key, reverse=True)

    for cl in clusters:
        cl.pop("_kostenstelle_norm", None)
        cl.pop("_kommission_norm", None)
        cl.pop("_baustelle_norm", None)

    return clusters


def aggregate_hints(hints):
    counts = {hint_type: 0 for hint_type in REPORT_HINT_TYPES}
    extra_counts = Counter()

    for h in hints:
        hint_type = h.get("hinweis_typ") or ""
        if hint_type in counts:
            counts[hint_type] += 1
        else:
            extra_counts[hint_type or "UNBEKANNT"] += 1

    for key, value in extra_counts.items():
        counts[key] = value

    return counts


def build_hint_index(hints):
    by_invoice = defaultdict(list)
    for h in hints:
        by_invoice[h.get("rechnung_id")].append(h)
    return by_invoice


def build_supplier_summary(invoices, hint_index):
    grouped = defaultdict(lambda: {
        "anzahl_rechnungen": 0,
        "auffaellige_rechnungen": 0,
        "hinweise_gesamt": 0,
        "lieferant_name": "UNBEKANNT",
        "summe_brutto": 0.0,
    })

    for inv in invoices:
        supplier = inv.get("lieferant_name") or "UNBEKANNT"
        rec = grouped[supplier]
        rec["lieferant_name"] = supplier
        rec["anzahl_rechnungen"] += 1
        rec["summe_brutto"] += inv.get("brutto_summe", 0.0)
        inv_hints = hint_index.get(inv.get("rechnung_id"), [])
        rec["hinweise_gesamt"] += len(inv_hints)
        if inv.get("gesamtbewertung") != "OK" or inv_hints:
            rec["auffaellige_rechnungen"] += 1

    rows = sorted(grouped.values(), key=lambda x: x["summe_brutto"], reverse=True)
    for row in rows:
        row["summe_brutto"] = round(row["summe_brutto"], 2)
    return rows


def build_payment_summary(invoices, end_dt):
    faellige = []
    skonto = []

    for inv in invoices:
        brutto = inv.get("brutto_summe", 0.0)
        dokumenttyp = inv.get("dokumenttyp")

        if dokumenttyp == "RECHNUNG" and brutto > 0:
            due = parse_iso_date_flexible(inv.get("faelligkeitsdatum"))
            if due and end_dt and due <= end_dt:
                faellige.append({
                    "rechnung_id": inv.get("rechnung_id"),
                    "rechnungsnummer": inv.get("rechnungsnummer"),
                    "lieferant_name": inv.get("lieferant_name"),
                    "faelligkeitsdatum": inv.get("faelligkeitsdatum"),
                    "brutto_summe": round(brutto, 2),
                })

        skonto_prozent = inv.get("skonto_prozent", 0.0)
        skonto_betrag = inv.get("skonto_betrag", 0.0)
        if dokumenttyp == "RECHNUNG" and brutto > 0 and skonto_prozent > 0 and skonto_betrag > 0:
            skonto.append({
                "rechnung_id": inv.get("rechnung_id"),
                "rechnungsnummer": inv.get("rechnungsnummer"),
                "lieferant_name": inv.get("lieferant_name"),
                "skonto_prozent": round(skonto_prozent, 2),
                "skonto_betrag": round(skonto_betrag, 2),
            })

    faellige = sorted(
        faellige,
        key=lambda x: parse_float_de(x.get("brutto_summe")),
        reverse=True
    )
    skonto = sorted(
        skonto,
        key=lambda x: parse_float_de(x.get("skonto_betrag")),
        reverse=True
    )

    return {
        "faellige_rechnungen": faellige,
        "faellige_rechnungen_anzahl": len(faellige),
        "summe_faellig": round(sum(x["brutto_summe"] for x in faellige), 2),
        "skonto_chancen": skonto[:20],
    }


def invoice_priority_score(inv, invoice_hints):
    hint_types = [h.get("hinweis_typ") for h in invoice_hints]
    max_hint_priority = max([HINT_PRIORITY.get(ht, 5) for ht in hint_types] + [0])
    amount_score = min(abs(inv.get("brutto_summe", 0.0)) / 100.0, 50.0)

    doc_bonus = 8 if inv.get("dokumenttyp") == "GUTSCHRIFT" else 0
    status_bonus = 8 if inv.get("gesamtbewertung") != "OK" else 0
    hint_count_bonus = min(len(invoice_hints) * 4, 20)

    return max_hint_priority + amount_score + doc_bonus + status_bonus + hint_count_bonus


def build_important_invoices(invoices, hint_index, limit=10):
    rows = []

    for inv in invoices:
        inv_hints = hint_index.get(inv.get("rechnung_id"), [])
        if not inv_hints and inv.get("gesamtbewertung") == "OK":
            continue

        hint_types = sorted(set(h.get("hinweis_typ") for h in inv_hints if h.get("hinweis_typ")))
        rows.append({
            "rechnung_id": inv.get("rechnung_id"),
            "rechnungsnummer": inv.get("rechnungsnummer"),
            "lieferant_name": inv.get("lieferant_name"),
            "brutto_summe": round(inv.get("brutto_summe", 0.0), 2),
            "hinweise_anzahl": len(inv_hints),
            "hinweis_typen": hint_types,
            "_score": invoice_priority_score(inv, inv_hints),
        })

    rows = sorted(rows, key=lambda x: (x["_score"], abs(x["brutto_summe"])), reverse=True)
    for row in rows:
        row.pop("_score", None)
    return rows[:limit]


def build_email_summary(summary, hint_breakdown, supplier_rows, payment, mode):
    rechnungen_gesamt = summary["rechnungen_gesamt"]
    summe_brutto = summary["summe_brutto"]
    rechnungen_ok = summary["rechnungen_ok"]
    rechnungen_auffaellig = summary["rechnungen_auffaellig"]
    gutschriften_anzahl = summary["gutschriften_anzahl"]

    sorted_hints = sorted(
        [(k, v) for k, v in hint_breakdown.items() if v > 0],
        key=lambda x: (HINT_PRIORITY.get(x[0], 0), x[1]),
        reverse=True
    )
    top_hints = sorted_hints[:3]
    hint_text = ", ".join([f"{k}: {v}" for k, v in top_hints]) if top_hints else "keine"

    supplier_text = "kein Lieferant"
    if supplier_rows:
        supplier_text = f"{supplier_rows[0]['lieferant_name']} mit {supplier_rows[0]['summe_brutto']:.2f} Volumen"

    faellig_text = ""
    if payment["faellige_rechnungen_anzahl"] > 0:
        faellig_text = (
            f" Aktuell sind {payment['faellige_rechnungen_anzahl']} Rechnungen "
            f"mit insgesamt {payment['summe_faellig']:.2f} fällig."
        )

    skonto_text = ""
    if payment["skonto_chancen"]:
        top_skonto = payment["skonto_chancen"][:3]
        top_skonto_sum = round(sum(x["skonto_betrag"] for x in top_skonto), 2)
        skonto_text = f" Zusätzlich bestehen relevante Skonto-Chancen, allein die Top-3 liegen bei {top_skonto_sum:.2f}."

    report_label = mode or "Report"

    return (
        f"Im Zeitraum wurden {rechnungen_gesamt} Rechnungen mit einem Gesamtvolumen von {summe_brutto:.2f} verarbeitet. "
        f"Davon waren {rechnungen_ok} unauffällig und {rechnungen_auffaellig} auffällig. "
        f"Zusätzlich wurden {gutschriften_anzahl} Gutschriften erkannt. "
        f"Die wichtigsten Hinweisarten waren {hint_text}. "
        f"Größter Lieferant im {report_label} war {supplier_text}."
        f"{faellig_text}{skonto_text}"
    )


@app.route("/analyze", methods=["POST"])
def analyze():
    try:
        data = request.get_json() or {}

        mode = str(data.get("mode") or "wochen_report").strip()
        betrieb_id = str(data.get("betrieb_id") or "").strip()
        zeitraum = data.get("zeitraum", {}) or {}

        start_raw = zeitraum.get("start")
        end_raw = zeitraum.get("ende")

        rechnungen_raw = data.get("rechnungen", []) or []
        hinweise_raw = data.get("hinweise", []) or []
        historische_raw = data.get("historische_rechnungen", []) or []

        start_dt = parse_iso_date_flexible(start_raw)
        end_dt = parse_iso_date_flexible(end_raw)

        rechnungen = [build_invoice_record(x) for x in rechnungen_raw]
        historische_rechnungen = [build_invoice_record(x) for x in historische_raw]
        hinweise = [build_hint_record(x) for x in hinweise_raw]

        if start_dt and end_dt:
            rechnungen = [
                r for r in rechnungen
                if is_in_period(r.get("rechnungsdatum") or r.get("eingangsdatum"), start_dt, end_dt)
            ]
            hinweise = [
                h for h in hinweise
                if is_in_period(h.get("erkannt_am"), start_dt, end_dt)
            ]

        if betrieb_id:
            hinweise = [
                h for h in hinweise
                if not h.get("betrieb_id") or h.get("betrieb_id") == betrieb_id
            ]

        rechnungen = dedupe_invoices(rechnungen)

        hint_index = build_hint_index(hinweise)

        for inv in rechnungen:
            inv_hints = hint_index.get(inv["rechnung_id"], [])
            inv["hinweise_anzahl"] = max(inv.get("hinweise_anzahl", 0), len(inv_hints))
            inv["_severity_seed"] = invoice_priority_score(inv, inv_hints)

            if inv_hints and inv.get("gesamtbewertung") == "OK":
                inv["gesamtbewertung"] = "HINWEIS"

        clustering_pool = rechnungen + historische_rechnungen
        clustering_pool = dedupe_invoices(clustering_pool)
        projekt_cluster = build_project_clusters_report(clustering_pool)

        current_ids = {r["rechnung_id"] for r in rechnungen}
        report_clusters = []
        for cl in projekt_cluster:
            current_cluster_ids = [rid for rid in cl["zugeordnete_rechnung_ids"] if rid in current_ids]
            if not current_cluster_ids:
                continue
            copied = dict(cl)
            copied["zugeordnete_rechnung_ids"] = current_cluster_ids
            copied["anzahl_rechnungen"] = len(current_cluster_ids)
            report_clusters.append(copied)

        hinweis_breakdown = aggregate_hints(hinweise)
        supplier_rows = build_supplier_summary(rechnungen, hint_index)
        payment = build_payment_summary(rechnungen, end_dt)
        wichtige_rechnungen = build_important_invoices(rechnungen, hint_index, limit=10)

        rechnungen_gesamt = len(rechnungen)
        gutschriften_anzahl = sum(1 for r in rechnungen if r.get("dokumenttyp") == "GUTSCHRIFT")
        rechnungen_auffaellig = sum(
            1 for r in rechnungen
            if r.get("gesamtbewertung") != "OK" or len(hint_index.get(r.get("rechnung_id"), [])) > 0
        )
        rechnungen_ok = max(rechnungen_gesamt - rechnungen_auffaellig, 0)
        summe_brutto = round(sum(r.get("brutto_summe", 0.0) for r in rechnungen), 2)
        hinweise_gesamt = len(hinweise)

        summary = {
            "rechnungen_gesamt": rechnungen_gesamt,
            "summe_brutto": summe_brutto,
            "rechnungen_ok": rechnungen_ok,
            "rechnungen_auffaellig": rechnungen_auffaellig,
            "gutschriften_anzahl": gutschriften_anzahl,
            "faellige_rechnungen": payment["faellige_rechnungen_anzahl"],
            "summe_faellig": payment["summe_faellig"],
            "hinweise_gesamt": hinweise_gesamt,
        }

        unklare_cluster = [
            cl for cl in report_clusters
            if cl.get("status") in {"mittel", "unsicher"} or cl.get("offen_unterbestimmt")
        ]

        email_summary = build_email_summary(
            summary=summary,
            hint_breakdown=hinweis_breakdown,
            supplier_rows=supplier_rows,
            payment=payment,
            mode=mode
        )

        result = {
            "ok": True,
            "meta": {
                "report_type": mode,
                "betrieb_id": betrieb_id,
                "zeitraum_start": start_raw,
                "zeitraum_ende": end_raw,
                "generated_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
            },
            "summary": summary,
            "hinweis_breakdown": hinweis_breakdown,
            "lieferanten": {
                "top_lieferanten": supplier_rows[:10]
            },
            "projekt_cluster": report_clusters,
            "payment": payment,
            "email_summary": email_summary,
            "optional_details": {
                "wichtige_rechnungen": wichtige_rechnungen,
                "unklare_projektzuordnungen": unklare_cluster[:20],
            }
        }

        return jsonify(result), 200

    except Exception as e:
        return jsonify({
            "ok": False,
            "error": str(e),
        }), 200
        
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
