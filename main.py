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
from datetime import datetime, date
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
        
# ============================================================
# ANALYZE HELPERS
# ============================================================

KANON_HINWEIS_TYPEN = [
    "PREISABWEICHUNG",
    "MENGENABWEICHUNG",
    "PREISSPRUNG_AUFFAELLIG",
    "GEBUEHR_ERKANNT",
    "GEBUEHR_POSITION",
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

def normalize_text_basic(value):
    s = str(value or "").strip().lower()
    s = s.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
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
    repl = {
        "straße": "str",
        "strasse": "str",
        "str.": "str",
        "straße.": "str",
        "platz": "pl",
        "allee": "all",
    }
    for k, v in repl.items():
        s = s.replace(k, v)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

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
        return float(value)

    s = str(value).strip()
    if not s:
        return 0.0

    s = s.replace("€", "").replace("%", "").replace(" ", "")
    s = s.replace("\u00a0", "")
    s = s.replace(".", "").replace(",", ".")

    s = re.sub(r"[^0-9\.\-]", "", s)

    if not s or s in ("-", ".", "-.", ".-"):
        return 0.0

    try:
        return float(s)
    except:
        return 0.0

def parse_date_safe(value):
    if not value:
        return None

    if isinstance(value, date):
        return value

    s = str(value).strip()
    if not s:
        return None

    formats = [
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%Y/%m/%d",
        "%d-%m-%Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except:
            pass

    return None

def canonical_hint_type(value):
    s = str(value or "").strip().upper()
    if s in KANON_HINWEIS_TYPEN:
        return s
    return s if s else "UNBEKANNT"

def get_rechnung_id(r):
    return (
        r.get("rechnung_id")
        or r.get("Rechnung_ID")
        or r.get("id")
        or r.get("ID")
        or ""
    )

def get_lieferant_name(r):
    return (
        r.get("lieferant_name")
        or r.get("Lieferant_Name")
        or r.get("lieferant")
        or r.get("Lieferant")
        or r.get("lieferant_id")
        or r.get("Lieferant_ID")
        or "UNBEKANNT"
    )

def get_dokumenttyp(r):
    return str(
        r.get("dokumenttyp")
        or r.get("Dokumenttyp")
        or ""
    ).strip().upper()

def is_gutschrift(r):
    dt = get_dokumenttyp(r)
    if "GUTSCHRIFT" in dt:
        return True

    brutto = (
        r.get("brutto_summe")
        or r.get("Brutto_Summe")
        or r.get("gesamt_brutto")
        or r.get("Gesamt_Brutto")
    )
    return to_float_safe(brutto) < 0

def is_rechnung_auffaellig(r, hinweise_by_rechnung):
    rid = get_rechnung_id(r)
    gesamtbewertung = str(
        r.get("gesamtbewertung")
        or r.get("Gesamtbewertung")
        or r.get("pruefung_status")
        or r.get("Pruefung_Status")
        or ""
    ).strip().upper()

    if rid and len(hinweise_by_rechnung.get(rid, [])) > 0:
        return True

    if "AUFFAELLIG" in gesamtbewertung:
        return True
    if "WARN" in gesamtbewertung:
        return True
    if "FEHLER" in gesamtbewertung:
        return True
    if "PRUEFEN" in gesamtbewertung:
        return True

    hinweise_anzahl = (
        r.get("hinweise_anzahl")
        or r.get("Hinweise_Anzahl")
        or 0
    )
    try:
        if int(hinweise_anzahl) > 0:
            return True
    except:
        pass

    return False

def get_brutto_summe(r):
    return to_float_safe(
        r.get("brutto_summe")
        or r.get("Brutto_Summe")
        or r.get("gesamt_brutto")
        or r.get("Gesamt_Brutto")
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

def is_faellig(r, zeitraum_ende):
    fad = get_faelligkeitsdatum(r)
    if not fad or not zeitraum_ende:
        return False
    return fad <= zeitraum_ende

def extract_project_features(r):
    kostenstelle = (
        r.get("kostenstelle")
        or r.get("Kostenstelle")
        or ""
    )
    kommission = (
        r.get("kommission")
        or r.get("Kommission")
        or ""
    )
    baustelle = (
        r.get("baustelle")
        or r.get("Baustelle")
        or ""
    )
    projekt_text = (
        r.get("projekt_hinweis_text")
        or r.get("Projekt_Hinweis_Text")
        or r.get("kommission_hinweis")
        or r.get("Kommission_Hinweis")
        or ""
    )

    return {
        "kostenstelle_raw": str(kostenstelle or "").strip(),
        "kommission_raw": str(kommission or "").strip(),
        "baustelle_raw": str(baustelle or "").strip(),
        "projekt_text_raw": str(projekt_text or "").strip(),

        "kostenstelle_norm": normalize_code(kostenstelle),
        "kommission_norm": normalize_name(kommission),
        "baustelle_norm": normalize_address(baustelle),
        "projekt_text_norm": normalize_name(projekt_text),
    }

def choose_best_value(values):
    values = [str(v).strip() for v in values if str(v or "").strip()]
    if not values:
        return ""
    c = Counter(values)
    return c.most_common(1)[0][0]

def cluster_match_score(cluster, feat):
    score = 0.0
    reasons = []

    if feat["kostenstelle_norm"] and cluster["kostenstelle_norm"]:
        if feat["kostenstelle_norm"] == cluster["kostenstelle_norm"]:
            score += 1.0
            reasons.append("KOSTENSTELLE_GLEICH")
        else:
            return 0.0, ["KOSTENSTELLE_KONFLIKT"]

    if feat["baustelle_norm"] and cluster["baustelle_norm"]:
        if feat["baustelle_norm"] == cluster["baustelle_norm"]:
            score += 0.9
            reasons.append("BAUSTELLE_GLEICH")
        else:
            sim = text_similarity(feat["baustelle_norm"], cluster["baustelle_norm"])
            if sim >= 0.88:
                score += 0.65
                reasons.append("BAUSTELLE_AEHNLICH")
            elif cluster["kostenstelle_norm"] and feat["kostenstelle_norm"] and cluster["kostenstelle_norm"] == feat["kostenstelle_norm"]:
                reasons.append("BAUSTELLE_ABWEICHEND_ABER_KOSTENSTELLE_STARK")
            elif sim < 0.55 and not (
                cluster["kostenstelle_norm"] and feat["kostenstelle_norm"] and cluster["kostenstelle_norm"] == feat["kostenstelle_norm"]
            ):
                return 0.0, ["BAUSTELLE_KONFLIKT"]

    if feat["kommission_norm"] and cluster["kommission_norm"]:
        if feat["kommission_norm"] == cluster["kommission_norm"]:
            score += 0.7
            reasons.append("KOMMISSION_GLEICH")
        else:
            sim = text_similarity(feat["kommission_norm"], cluster["kommission_norm"])
            if sim >= 0.86:
                score += 0.45
                reasons.append("KOMMISSION_AEHNLICH")

    if feat["projekt_text_norm"]:
        if cluster["kostenstelle_norm"] and cluster["kostenstelle_norm"] in feat["projekt_text_norm"]:
            score += 0.35
            reasons.append("TEXT_ENTHAELT_KOSTENSTELLE")
        if cluster["kommission_norm"] and cluster["kommission_norm"] in feat["projekt_text_norm"]:
            score += 0.25
            reasons.append("TEXT_ENTHAELT_KOMMISSION")
        if cluster["baustelle_norm"] and cluster["baustelle_norm"] in feat["projekt_text_norm"]:
            score += 0.25
            reasons.append("TEXT_ENTHAELT_BAUSTELLE")

    return score, reasons

def build_project_clusters(rechnungen):
    clusters = []
    cluster_seq = 1

    for r in rechnungen:
        feat = extract_project_features(r)
        rid = get_rechnung_id(r)
        brutto = get_brutto_summe(r)

        has_any_project_data = any([
            feat["kostenstelle_norm"],
            feat["kommission_norm"],
            feat["baustelle_norm"],
            feat["projekt_text_norm"],
        ])

        if not has_any_project_data:
            continue

        best_idx = None
        best_score = 0.0
        best_reasons = []

        for idx, cluster in enumerate(clusters):
            score, reasons = cluster_match_score(cluster, feat)
            if score > best_score:
                best_score = score
                best_idx = idx
                best_reasons = reasons

        match_threshold = 1.0

        if best_idx is not None and best_score >= match_threshold:
            c = clusters[best_idx]
            c["rechnung_ids"].append(rid)
            c["summe_brutto"] += brutto
            c["rechnungen"].append(r)

            if feat["kostenstelle_raw"]:
                c["kostenstelle_values"].append(feat["kostenstelle_raw"])
            if feat["kommission_raw"]:
                c["kommission_values"].append(feat["kommission_raw"])
            if feat["baustelle_raw"]:
                c["baustelle_values"].append(feat["baustelle_raw"])
            if feat["projekt_text_raw"]:
                c["projekt_text_values"].append(feat["projekt_text_raw"])

            c["match_reasons"].extend(best_reasons)

            if feat["kostenstelle_norm"]:
                c["kostenstelle_norm"] = feat["kostenstelle_norm"]
            if feat["kommission_norm"] and not c["kommission_norm"]:
                c["kommission_norm"] = feat["kommission_norm"]
            if feat["baustelle_norm"] and not c["baustelle_norm"]:
                c["baustelle_norm"] = feat["baustelle_norm"]

        else:
            clusters.append({
                "projekt_cluster_id": f"PC_{cluster_seq:04d}",
                "rechnung_ids": [rid],
                "summe_brutto": brutto,
                "rechnungen": [r],

                "kostenstelle_norm": feat["kostenstelle_norm"],
                "kommission_norm": feat["kommission_norm"],
                "baustelle_norm": feat["baustelle_norm"],

                "kostenstelle_values": [feat["kostenstelle_raw"]] if feat["kostenstelle_raw"] else [],
                "kommission_values": [feat["kommission_raw"]] if feat["kommission_raw"] else [],
                "baustelle_values": [feat["baustelle_raw"]] if feat["baustelle_raw"] else [],
                "projekt_text_values": [feat["projekt_text_raw"]] if feat["projekt_text_raw"] else [],
                "match_reasons": ["NEUES_CLUSTER"],
            })
            cluster_seq += 1

    result = []

    for c in clusters:
        kostenstelle = choose_best_value(c["kostenstelle_values"])
        kommission = choose_best_value(c["kommission_values"])
        baustelle = choose_best_value(c["baustelle_values"])

        confidence = 0.35
        if kostenstelle:
            confidence += 0.30
        if baustelle:
            confidence += 0.20
        if kommission:
            confidence += 0.15
        if len(c["rechnung_ids"]) >= 2:
            confidence += 0.10

        confidence = min(round(confidence, 2), 0.98)

        status = "sicher"
        if confidence < 0.55:
            status = "unsicher"
        elif confidence < 0.75:
            status = "mittel"

        result.append({
            "projekt_cluster_id": c["projekt_cluster_id"],
            "erkannte_kostenstelle": kostenstelle,
            "erkannte_kommission": kommission,
            "erkannte_baustelle": baustelle,
            "projekt_summe_brutto": round(c["summe_brutto"], 2),
            "anzahl_rechnungen": len(c["rechnung_ids"]),
            "confidence": confidence,
            "status": status,
            "zugeordnete_rechnung_ids": c["rechnung_ids"],
            "match_hinweise": sorted(list(set(c["match_reasons"]))),
            "offen_unterbestimmt": True if not (kostenstelle or kommission or baustelle) else False,
        })

    result.sort(key=lambda x: x["projekt_summe_brutto"], reverse=True)
    return result

def build_hinweis_breakdown(hinweise):
    counter = Counter()
    for h in hinweise:
        typ = canonical_hint_type(
            h.get("hinweis_typ")
            or h.get("Hinweis_Typ")
        )
        counter[typ] += 1

    breakdown = {}
    for typ in KANON_HINWEIS_TYPEN:
        breakdown[typ] = counter.get(typ, 0)

    # unbekannte mitnehmen
    extra = {k: v for k, v in counter.items() if k not in breakdown}
    breakdown.update(extra)

    return breakdown

def build_top_lieferanten(rechnungen, hinweise_by_rechnung):
    supplier_map = defaultdict(lambda: {
        "lieferant_name": "",
        "anzahl_rechnungen": 0,
        "summe_brutto": 0.0,
        "auffaellige_rechnungen": 0,
        "hinweise_gesamt": 0,
    })

    for r in rechnungen:
        supplier = str(get_lieferant_name(r))
        rid = get_rechnung_id(r)
        brutto = get_brutto_summe(r)

        item = supplier_map[supplier]
        item["lieferant_name"] = supplier
        item["anzahl_rechnungen"] += 1
        item["summe_brutto"] += brutto
        item["hinweise_gesamt"] += len(hinweise_by_rechnung.get(rid, []))

        if is_rechnung_auffaellig(r, hinweise_by_rechnung):
            item["auffaellige_rechnungen"] += 1

    result = []
    for _, v in supplier_map.items():
        v["summe_brutto"] = round(v["summe_brutto"], 2)
        result.append(v)

    result.sort(key=lambda x: x["summe_brutto"], reverse=True)
    return result[:10]

def build_wichtige_rechnungen(rechnungen, hinweise_by_rechnung):
    items = []

    for r in rechnungen:
        rid = get_rechnung_id(r)
        hinweise = hinweise_by_rechnung.get(rid, [])
        if not hinweise:
            continue

        items.append({
            "rechnung_id": rid,
            "rechnungsnummer": (
                r.get("rechnungsnummer")
                or r.get("Rechnungsnummer")
                or ""
            ),
            "lieferant_name": get_lieferant_name(r),
            "brutto_summe": round(get_brutto_summe(r), 2),
            "hinweise_anzahl": len(hinweise),
            "hinweis_typen": sorted(list(set(
                canonical_hint_type(h.get("hinweis_typ") or h.get("Hinweis_Typ"))
                for h in hinweise
            ))),
        })

    items.sort(key=lambda x: (x["hinweise_anzahl"], x["brutto_summe"]), reverse=True)
    return items[:10]

def build_payment_section(rechnungen, zeitraum_ende):
    faellige = []
    summe_faellig = 0.0
    skonto_chancen = []

    for r in rechnungen:
        if is_faellig(r, zeitraum_ende):
            brutto = get_brutto_summe(r)
            summe_faellig += brutto
            faellige.append({
                "rechnung_id": get_rechnung_id(r),
                "rechnungsnummer": (
                    r.get("rechnungsnummer")
                    or r.get("Rechnungsnummer")
                    or ""
                ),
                "lieferant_name": get_lieferant_name(r),
                "faelligkeitsdatum": str(get_faelligkeitsdatum(r) or ""),
                "brutto_summe": round(brutto, 2),
            })

        skonto_prozent = get_skonto_prozent(r)
        skonto_betrag = get_skonto_betrag(r)
        if skonto_prozent > 0 or skonto_betrag > 0:
            skonto_chancen.append({
                "rechnung_id": get_rechnung_id(r),
                "rechnungsnummer": (
                    r.get("rechnungsnummer")
                    or r.get("Rechnungsnummer")
                    or ""
                ),
                "lieferant_name": get_lieferant_name(r),
                "skonto_prozent": skonto_prozent,
                "skonto_betrag": round(skonto_betrag, 2),
            })

    faellige.sort(key=lambda x: x["brutto_summe"], reverse=True)
    skonto_chancen.sort(key=lambda x: x["skonto_betrag"], reverse=True)

    return {
        "faellige_rechnungen_anzahl": len(faellige),
        "summe_faellig": round(summe_faellig, 2),
        "faellige_rechnungen": faellige[:20],
        "skonto_chancen": skonto_chancen[:20],
    }

def build_email_summary(summary, breakdown, payment, top_lieferanten):
    parts = []

    parts.append(
        f"Im Zeitraum wurden {summary['rechnungen_gesamt']} Rechnungen mit einem Gesamtvolumen von {summary['summe_brutto']:.2f} verarbeitet."
    )

    parts.append(
        f"Davon waren {summary['rechnungen_ok']} unauffällig und {summary['rechnungen_auffaellig']} auffällig."
    )

    if summary["gutschriften_anzahl"] > 0:
        parts.append(f"Zusätzlich wurden {summary['gutschriften_anzahl']} Gutschriften erkannt.")

    dominante = sorted(
        [(k, v) for k, v in breakdown.items() if v > 0],
        key=lambda x: x[1],
        reverse=True
    )[:3]

    if dominante:
        txt = ", ".join([f"{k}: {v}" for k, v in dominante])
        parts.append(f"Die häufigsten Hinweisarten waren {txt}.")

    if payment["faellige_rechnungen_anzahl"] > 0:
        parts.append(
            f"Aktuell sind {payment['faellige_rechnungen_anzahl']} Rechnungen mit insgesamt {payment['summe_faellig']:.2f} fällig."
        )

    if top_lieferanten:
        top = top_lieferanten[0]
        parts.append(
            f"Größter Lieferant im Zeitraum war {top['lieferant_name']} mit {top['summe_brutto']:.2f} Volumen."
        )

    return " ".join(parts).strip()
        
@app.route("/analyze", methods=["POST"])
def analyze():
    try:
        data = request.get_json() or {}

        mode = data.get("mode", "weekly_report")
        betrieb_id = data.get("betrieb_id")
        zeitraum = data.get("zeitraum", {}) or {}
        rechnungen = data.get("rechnungen", []) or []
        hinweise = data.get("hinweise", []) or []
        historische_rechnungen = data.get("historische_rechnungen", []) or []

        zeitraum_start = parse_date_safe(zeitraum.get("start"))
        zeitraum_ende = parse_date_safe(zeitraum.get("ende"))

        hinweise_by_rechnung = defaultdict(list)
        for h in hinweise:
            rid = (
                h.get("rechnung_id")
                or h.get("Rechnung_ID")
                or h.get("bezug_schluessel")
                or h.get("Bezug_Schluessel")
                or ""
            )
            if rid:
                hinweise_by_rechnung[str(rid)].append(h)

        rechnungen_gesamt = len(rechnungen)
        summe_brutto = round(sum(get_brutto_summe(r) for r in rechnungen), 2)

        rechnungen_auffaellig = sum(
            1 for r in rechnungen if is_rechnung_auffaellig(r, hinweise_by_rechnung)
        )
        rechnungen_ok = max(rechnungen_gesamt - rechnungen_auffaellig, 0)

        gutschriften_anzahl = sum(1 for r in rechnungen if is_gutschrift(r))
        hinweise_gesamt = len(hinweise)

        hinweis_breakdown = build_hinweis_breakdown(hinweise)
        payment = build_payment_section(rechnungen, zeitraum_ende)
        top_lieferanten = build_top_lieferanten(rechnungen, hinweise_by_rechnung)

        cluster_input = list(rechnungen)
        if historische_rechnungen:
            cluster_input.extend(historische_rechnungen)

        projekt_cluster = build_project_clusters(cluster_input)

        wichtige_rechnungen = build_wichtige_rechnungen(rechnungen, hinweise_by_rechnung)
        unklare_projekte = [p for p in projekt_cluster if p["status"] == "unsicher"][:20]

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

        email_summary = build_email_summary(
            summary=summary,
            breakdown=hinweis_breakdown,
            payment=payment,
            top_lieferanten=top_lieferanten
        )

        return jsonify({
            "ok": True,
            "meta": {
                "report_type": mode,
                "betrieb_id": betrieb_id,
                "zeitraum_start": str(zeitraum_start) if zeitraum_start else zeitraum.get("start"),
                "zeitraum_ende": str(zeitraum_ende) if zeitraum_ende else zeitraum.get("ende"),
                "generated_at": datetime.utcnow().isoformat() + "Z"
            },
            "summary": summary,
            "hinweis_breakdown": hinweis_breakdown,
            "lieferanten": {
                "top_lieferanten": top_lieferanten
            },
            "projekt_cluster": projekt_cluster,
            "payment": payment,
            "email_summary": email_summary,
            "optional_details": {
                "wichtige_rechnungen": wichtige_rechnungen,
                "unklare_projektzuordnungen": unklare_projekte
            }
        }), 200

    except Exception as e:
        return jsonify({
            "ok": False,
            "error": "analyze_failed",
            "error_detail": str(e)
        }), 200

        
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
