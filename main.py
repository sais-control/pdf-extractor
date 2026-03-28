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
        
# ============================================================
# ANALYZE HELPERS / REPORT LOGIK V3
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

LIEFERANTEN_TYP_MAPPING = {
    "GROSSHAENDLER": "material",
    "GROSSHÄNDLER": "material",
    "GROSSHANDEL": "material",
    "FACHGROSSHANDEL": "material",
    "HERSTELLER": "material",
    "LIEFERANT": "material",
    "MATERIAL": "material",
    "SUBUNTERNEHMER": "subunternehmer",
    "NACHUNTERNEHMER": "subunternehmer",
    "DIENSTLEISTER": "subunternehmer",
    "HANDWERKER": "subunternehmer",
    "ELEKTRIKER": "subunternehmer",
    "MONTEUR": "subunternehmer",
    "SONSTIGES": "sonstiges",
    "FUHRPARK": "sonstiges",
    "MIETE": "sonstiges",
    "TRANSPORT": "sonstiges",
    "ENTSORGUNG": "sonstiges",
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
        "cornelius gellert": "gellert",
        "cornelius-gellert": "gellert",
        "corn gellert": "gellert",
        "corn gellertstr": "gellertstr",
        "corn gellert str": "gellertstr",
        "cornelius gellert str": "gellertstr",
        "cornelius gellertstr": "gellertstr",
        "corn gellertstraße": "gellertstr",
        "cornelius gellertstraße": "gellertstr",
        "corn gellert strasse": "gellertstr",
        "cornelius gellert strasse": "gellertstr",
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

    blacklist = {
        "lager", "webshop", "elements", "abholung", "lieferung",
        "innend", "aussend", "außend", "kunde", "projekt", "baustelle"
    }

    parts = [p for p in s.split() if p and p not in blacklist and not re.search(r"\d", p)]
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

    return True

def looks_like_company_address(value):
    s = normalize_text_basic(value)
    if not s:
        return False

    company_markers = [
        "gmbh", "mbh", "ug", "gbr", "kg", "ag", "ohg",
        "gebaeudetechnik", "haustechnik", "elektro", "sanitaer",
        "heizung", "klima", "technik"
    ]
    return any(m in s for m in company_markers)

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
        except:
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
    except:
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
    except:
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
        except:
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
    return "RECHNUNG"

def canonical_lieferant_typ(value):
    s = str(value or "").strip().upper()
    if not s:
        return ""
    return s

def map_lieferant_typ_to_kategorie(value):
    s = canonical_lieferant_typ(value)
    return LIEFERANTEN_TYP_MAPPING.get(s, "sonstiges")

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
    return canonical_lieferant_typ(
        r.get("lieferant_typ")
        or r.get("Lieferant_Typ")
        or r.get("lieferanten_typ")
        or r.get("Lieferanten_Typ")
        or r.get("kategorie")
        or r.get("Kategorie")
        or ""
    )

def get_lieferant_typ_from_kontext_map(r, lieferanten_kontext_map):
    lid = get_lieferant_id(r)
    lname = get_lieferant_name(r)

    if lid and lid in lieferanten_kontext_map:
        return canonical_lieferant_typ(lieferanten_kontext_map[lid].get("lieferant_typ"))

    lname_norm = normalize_name(lname)
    for _, entry in lieferanten_kontext_map.items():
        if normalize_name(entry.get("lieferant_name")) == lname_norm:
            return canonical_lieferant_typ(entry.get("lieferant_typ"))

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

def get_hinweise_anzahl_rechnung(r):
    v = (
        r.get("hinweise_anzahl")
        or r.get("Hinweise_Anzahl")
        or 0
    )
    try:
        return int(v)
    except:
        return 0

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

def is_faellig(r, zeitraum_ende):
    fad = get_faelligkeitsdatum(r)
    if not fad or not zeitraum_ende:
        return False
    return fad <= zeitraum_ende

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

def extract_project_features(r):
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
    baustelle_raw = str(baustelle or "").strip()
    projekt_text_raw = str(projekt_text or "").strip()

    return {
        "kostenstelle_raw": kostenstelle_raw,
        "kommission_raw": kommission_raw,
        "baustelle_raw": baustelle_raw,
        "projekt_text_raw": projekt_text_raw,
        "kostenstelle_norm": normalize_code(kostenstelle_raw),
        "kommission_norm": normalize_person_name_for_match(kommission_raw),
        "baustelle_norm": normalize_address(baustelle_raw),
        "address_key": normalize_address_key(baustelle_raw),
        "projekt_text_norm": normalize_name(projekt_text_raw),
        "kostenstelle_generisch": is_generic_kostenstelle(kostenstelle_raw),
    }
    
def canonical_lieferanten_kategorie(value):
    s = str(value or "").strip().upper()

    if not s:
        return "SONSTIGES"

    mapping = {
        "GROSSHANDEL": "GROSSHANDEL",
        "GROSSHANDEL": "GROSSHANDEL",
        "HANDEL": "GROSSHANDEL",
        "HANDEL_ALLGEMEIN": "GROSSHANDEL",

        "SUBUNTERNEHMER": "SUBUNTERNEHMER",
        "SUB": "SUBUNTERNEHMER",

        "DIENSTLEISTER": "DIENSTLEISTER",
        "SERVICE": "DIENSTLEISTER",

        "HERSTELLER": "HERSTELLER",

        "FIXKOSTEN": "FIXKOSTEN",

        "WERKSTATT": "WERKSTATT",

        "ARBEITSKLEIDUNG": "ARBEITSKLEIDUNG",

        "SONSTIGES": "SONSTIGES",
    }

    return mapping.get(s, s)

def build_lieferanten_kontext_map(lieferanten_kontext):
    result = {}

    for item in (lieferanten_kontext or []):
        lieferant_id = str(
            item.get("lieferant_id")
            or item.get("Lieferant_ID")
            or ""
        ).strip()

        lieferant_name = str(
            item.get("lieferant_name")
            or item.get("Lieferant_Name")
            or ""
        ).strip()

        lieferant_typ = str(
            item.get("lieferant_typ")
            or item.get("Lieferant_Typ")
            or item.get("kategorie")
            or item.get("Kategorie")
            or ""
        ).strip()

        entry = {
            "lieferant_id": lieferant_id,
            "lieferant_name": lieferant_name,
            "lieferant_typ": lieferant_typ,
        }

        if lieferant_id:
            result[lieferant_id] = entry
        elif lieferant_name:
            result[f"NAME::{normalize_name(lieferant_name)}"] = entry

    return result

def build_lieferanten_kontext_output(lieferanten_kontext):
    out = []

    for item in (lieferanten_kontext or []):
        lieferant_id = str(
            item.get("lieferant_id")
            or item.get("Lieferant_ID")
            or ""
        ).strip()

        lieferant_name = str(
            item.get("lieferant_name")
            or item.get("Lieferant_Name")
            or ""
        ).strip()

        lieferant_typ = str(
            item.get("lieferant_typ")
            or item.get("Lieferant_Typ")
            or item.get("kategorie")
            or item.get("Kategorie")
            or ""
        ).strip()

        out.append({
            "lieferant_id": lieferant_id,
            "lieferant_name": lieferant_name,
            "lieferant_typ": lieferant_typ,
            "kosten_kategorie": map_lieferant_typ_to_kategorie(lieferant_typ),
        })

    out.sort(key=lambda x: (x["lieferant_name"], x["lieferant_id"]))
    return out

def build_project_clusters(rechnungen, lieferanten_kontext_map=None):
    lieferanten_kontext_map = lieferanten_kontext_map or {}

    prepared = []

    def is_lager_item(feat):
        values = [
            str(feat.get("kostenstelle_raw") or "").strip().lower(),
            str(feat.get("kommission_raw") or "").strip().lower(),
            str(feat.get("baustelle_raw") or "").strip().lower(),
            str(feat.get("projekt_text_raw") or "").strip().lower(),
        ]
        joined = " ".join(values)
        return ("p1" in joined) or ("lager" in joined) or ("lagerbestand" in joined)

    def make_cluster_from_item(item, reason):
        feat = item["feat"]
        rechnung = item["rechnung"]
        brutto = get_brutto_summe(rechnung)

        kategorie = get_lieferant_kategorie(rechnung, lieferanten_kontext_map)

        material = brutto if (is_rechnung(rechnung) and kategorie == "material") else 0.0
        subunternehmer = brutto if (is_rechnung(rechnung) and kategorie == "subunternehmer") else 0.0
        sonstiges = brutto if (is_rechnung(rechnung) and kategorie == "sonstiges") else 0.0

        return {
            "rechnungen": [rechnung],
            "rechnung_ids": [get_rechnung_id(rechnung)],
            "rechnungsnummern": [get_rechnungsnummer(rechnung)] if get_rechnungsnummer(rechnung) else [],
            "kostenstelle_values": [feat["kostenstelle_raw"]] if feat["kostenstelle_raw"] else [],
            "kommission_values": [feat["kommission_raw"]] if feat["kommission_raw"] else [],
            "baustelle_values": [feat["baustelle_raw"]] if feat["baustelle_raw"] else [],
            "projekt_text_values": [feat["projekt_text_raw"]] if feat["projekt_text_raw"] else [],
            "address_keys": {feat["address_key"]} if feat["address_key"] else set(),
            "kostenstelle_norms": {feat["kostenstelle_norm"]} if feat["kostenstelle_norm"] else set(),
            "kommission_norms": {feat["kommission_norm"]} if feat["kommission_norm"] else set(),
            "match_hinweise": [reason],
            "is_lager": is_lager_item(feat),
            "rechnung_summe_brutto": round(brutto if is_rechnung(rechnung) else 0.0, 2),
            "gutschrift_summe_brutto": round(brutto if is_gutschrift(rechnung) else 0.0, 2),
            "nettoeffekt_brutto": round(brutto, 2),
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

        kategorie = get_lieferant_kategorie(rechnung, lieferanten_kontext_map)

        if get_rechnung_id(rechnung):
            cluster["rechnung_ids"].append(get_rechnung_id(rechnung))
        if get_rechnungsnummer(rechnung):
            cluster["rechnungsnummern"].append(get_rechnungsnummer(rechnung))

        cluster["rechnungen"].append(rechnung)

        if feat["kostenstelle_raw"]:
            cluster["kostenstelle_values"].append(feat["kostenstelle_raw"])
        if feat["kommission_raw"]:
            cluster["kommission_values"].append(feat["kommission_raw"])
        if feat["baustelle_raw"]:
            cluster["baustelle_values"].append(feat["baustelle_raw"])
        if feat["projekt_text_raw"]:
            cluster["projekt_text_values"].append(feat["projekt_text_raw"])

        if feat["address_key"]:
            cluster["address_keys"].add(feat["address_key"])
        if feat["kostenstelle_norm"]:
            cluster["kostenstelle_norms"].add(feat["kostenstelle_norm"])
        if feat["kommission_norm"]:
            cluster["kommission_norms"].add(feat["kommission_norm"])

        cluster["match_hinweise"].append(reason)

        if is_rechnung(rechnung):
            cluster["rechnung_summe_brutto"] = round(cluster["rechnung_summe_brutto"] + brutto, 2)
            cluster["anzahl_rechnungen"] += 1

            if kategorie == "material":
                cluster["kostenstruktur_material_brutto"] = round(cluster["kostenstruktur_material_brutto"] + brutto, 2)
            elif kategorie == "subunternehmer":
                cluster["kostenstruktur_subunternehmer_brutto"] = round(cluster["kostenstruktur_subunternehmer_brutto"] + brutto, 2)
            else:
                cluster["kostenstruktur_sonstiges_brutto"] = round(cluster["kostenstruktur_sonstiges_brutto"] + brutto, 2)

        if is_gutschrift(rechnung):
            cluster["gutschrift_summe_brutto"] = round(cluster["gutschrift_summe_brutto"] + brutto, 2)
            cluster["anzahl_gutschriften"] += 1

        cluster["nettoeffekt_brutto"] = round(cluster["nettoeffekt_brutto"] + brutto, 2)
        cluster["anzahl_dokumente"] += 1

    def strong_address_match(feat, cluster):
        if not feat["address_key"] or not cluster["address_keys"]:
            return False
        if feat["address_key"] in cluster["address_keys"]:
            return True

        for existing_key in cluster["address_keys"]:
            sim = text_similarity(feat["address_key"], existing_key)
            if sim >= 0.93:
                return True
        return False

    def strict_person_match(a, b):
        if not a or not b:
            return False
        if a == b:
            return True
        sim = text_similarity(a, b)
        return sim >= 0.93

    def can_attach_orphan_to_cluster(item, cluster):
        feat = item["feat"]

        # 1) harte Kostenstelle
        if (
            feat["kostenstelle_norm"]
            and not feat["kostenstelle_generisch"]
            and feat["kostenstelle_norm"] in cluster["kostenstelle_norms"]
        ):
            return "KOSTENSTELLE_NACHGEZOGEN"

        # 2) eindeutige Personen-Kommission
        if feat["kommission_norm"] and is_full_person_name(feat["kommission_raw"]):
            for existing_norm in cluster["kommission_norms"]:
                if strict_person_match(feat["kommission_norm"], existing_norm):
                    return "KOMMISSION_NACHGEZOGEN"

        return ""

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
            "kostenstelle_norms": set(a["kostenstelle_norms"]) | set(b["kostenstelle_norms"]),
            "kommission_norms": set(a["kommission_norms"]) | set(b["kommission_norms"]),
            "match_hinweise": list(a["match_hinweise"]) + list(b["match_hinweise"]) + ["CLUSTER_ZUSAMMENGEFUEHRT"],
            "is_lager": a["is_lager"] or b["is_lager"],
            "rechnung_summe_brutto": round(a["rechnung_summe_brutto"] + b["rechnung_summe_brutto"], 2),
            "gutschrift_summe_brutto": round(a["gutschrift_summe_brutto"] + b["gutschrift_summe_brutto"], 2),
            "nettoeffekt_brutto": round(a["nettoeffekt_brutto"] + b["nettoeffekt_brutto"], 2),
            "kostenstruktur_material_brutto": round(a["kostenstruktur_material_brutto"] + b["kostenstruktur_material_brutto"], 2),
            "kostenstruktur_subunternehmer_brutto": round(a["kostenstruktur_subunternehmer_brutto"] + b["kostenstruktur_subunternehmer_brutto"], 2),
            "kostenstruktur_sonstiges_brutto": round(a["kostenstruktur_sonstiges_brutto"] + b["kostenstruktur_sonstiges_brutto"], 2),
            "anzahl_rechnungen": a["anzahl_rechnungen"] + b["anzahl_rechnungen"],
            "anzahl_gutschriften": a["anzahl_gutschriften"] + b["anzahl_gutschriften"],
            "anzahl_dokumente": a["anzahl_dokumente"] + b["anzahl_dokumente"],
        }
        return merged

    def should_merge_clusters(a, b):
        if a["is_lager"] != b["is_lager"]:
            return False

        if a["is_lager"] and b["is_lager"]:
            return True

        a_address = choose_best_value(a["baustelle_values"])
        b_address = choose_best_value(b["baustelle_values"])
        a_kommission = normalize_person_name_for_match(choose_best_value(a["kommission_values"]))
        b_kommission = normalize_person_name_for_match(choose_best_value(b["kommission_values"]))
        a_kostenstelle = choose_best_value(a["kostenstelle_values"])
        b_kostenstelle = choose_best_value(b["kostenstelle_values"])

        # gleiche/nahe Adresse
        address_match = False
        if a["address_keys"] and b["address_keys"]:
            for ka in a["address_keys"]:
                for kb in b["address_keys"]:
                    if ka == kb or text_similarity(ka, kb) >= 0.93:
                        address_match = True
                        break
                if address_match:
                    break

        # gleiche starke Kostenstelle
        strong_kostenstelle_match = False
        if a_kostenstelle and b_kostenstelle:
            if (not is_generic_kostenstelle(a_kostenstelle)) and (not is_generic_kostenstelle(b_kostenstelle)):
                if normalize_code(a_kostenstelle) == normalize_code(b_kostenstelle):
                    strong_kostenstelle_match = True

        # gleiche Personen-Kommission
        person_match = False
        if a_kommission and b_kommission:
            if is_full_person_name(a_kommission) and is_full_person_name(b_kommission):
                if strict_person_match(a_kommission, b_kommission):
                    person_match = True

        if address_match and (strong_kostenstelle_match or person_match):
            return True

        # Spezialfall: gleiche echte Adresse, eine Seite mit Kostenstelle, andere ohne
        if address_match and (a_address or b_address):
            return True

        return False

    def build_cluster_result(cluster, idx):
        kostenstelle = choose_best_value(cluster["kostenstelle_values"])
        kommission = choose_best_value(cluster["kommission_values"])
        baustelle = choose_best_value(cluster["baustelle_values"])

        confidence = 0.55

        if cluster["is_lager"]:
            confidence = 0.80
        else:
            if cluster["address_keys"]:
                confidence += 0.20
            if kostenstelle and not is_generic_kostenstelle(kostenstelle):
                confidence += 0.10
            if kommission and is_full_person_name(kommission):
                confidence += 0.08
            if cluster["anzahl_dokumente"] >= 2:
                confidence += 0.07
            if "CLUSTER_ZUSAMMENGEFUEHRT" in cluster["match_hinweise"]:
                confidence += 0.03

        confidence = min(round(confidence, 2), 0.98)

        if confidence >= 0.82:
            status = "sicher"
        elif confidence >= 0.65:
            status = "mittel"
        else:
            status = "unsicher"

        projekt_name_report = ""
        if cluster["is_lager"]:
            projekt_name_report = "Lager / P1"
        elif baustelle and kostenstelle and not is_generic_kostenstelle(kostenstelle):
            projekt_name_report = f"{baustelle} / {kostenstelle}"
        elif baustelle:
            projekt_name_report = baustelle
        elif kostenstelle and not is_generic_kostenstelle(kostenstelle):
            projekt_name_report = kostenstelle
        elif kommission:
            projekt_name_report = kommission
        else:
            projekt_name_report = ""

        return {
            "projekt_cluster_id": f"PC_{idx:04d}",
            "projekt_name_report": projekt_name_report,
            "erkannte_baustelle": baustelle,
            "erkannte_kostenstelle": kostenstelle,
            "erkannte_kommission": kommission,
            "projekt_summe_brutto": round(cluster["nettoeffekt_brutto"], 2),
            "rechnung_summe_brutto": round(cluster["rechnung_summe_brutto"], 2),
            "gutschrift_summe_brutto": round(cluster["gutschrift_summe_brutto"], 2),
            "nettoeffekt_brutto": round(cluster["nettoeffekt_brutto"], 2),
            "anzahl_rechnungen": cluster["anzahl_rechnungen"],
            "anzahl_gutschriften": cluster["anzahl_gutschriften"],
            "anzahl_dokumente": cluster["anzahl_dokumente"],
            "confidence": confidence,
            "status": status,
            "zugeordnete_rechnung_ids": unique_nonempty(cluster["rechnung_ids"]),
            "zugeordnete_rechnungsnummern": unique_nonempty(cluster["rechnungsnummern"]),
            "match_hinweise": sorted(list(set(cluster["match_hinweise"]))),
            "offen_unterbestimmt": False,
        }

    # vorbereiten
    for r in rechnungen:
        feat = extract_project_features(r)
        prepared.append({
            "rechnung": r,
            "feat": feat,
        })

    address_clusters = []
    orphan_items = []
    lager_cluster = None

    # Phase 1: adressbasierte Cluster
    for item in prepared:
        feat = item["feat"]

        if is_lager_item(feat):
            if lager_cluster is None:
                lager_cluster = make_cluster_from_item(item, "LAGER_DIREKT")
            else:
                add_item_to_cluster(lager_cluster, item, "LAGER_DIREKT")
            continue

        if feat["address_key"]:
            matched = False
            for cluster in address_clusters:
                if strong_address_match(feat, cluster):
                    add_item_to_cluster(cluster, item, "ADRESSE_GLEICH")
                    matched = True
                    break

            if not matched:
                address_clusters.append(make_cluster_from_item(item, "ADRESSE_GLEICH"))
        else:
            orphan_items.append(item)

    # Phase 2: adresslose Fälle nur vorsichtig nachziehen
    standalone_clusters = []

    for item in orphan_items:
        candidates = []

        for cluster in address_clusters:
            reason = can_attach_orphan_to_cluster(item, cluster)
            if reason:
                candidates.append((cluster, reason))

        if len(candidates) == 1:
            cluster, reason = candidates[0]
            add_item_to_cluster(cluster, item, reason)
        else:
            standalone_clusters.append(make_cluster_from_item(item, "EINZELFALL"))

    all_clusters = list(address_clusters)
    if lager_cluster is not None:
        all_clusters.append(lager_cluster)
    all_clusters.extend(standalone_clusters)

    # Phase 3: Cluster zusammenführen
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

def build_project_report_list(projekt_cluster, lieferanten_kontext_map=None):
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

        out.append({
            "projekt_name": projekt_name,
            "erkannte_baustelle": str(cluster.get("erkannte_baustelle") or "").strip(),
            "erkannte_kostenstelle": str(cluster.get("erkannte_kostenstelle") or "").strip(),
            "erkannte_kommission": str(cluster.get("erkannte_kommission") or "").strip(),
            "status": str(cluster.get("status") or "").strip(),
            "confidence": to_float_safe(cluster.get("confidence")),
            "anzahl_rechnungen": int(cluster.get("anzahl_rechnungen") or 0),
            "anzahl_gutschriften": int(cluster.get("anzahl_gutschriften") or 0),
            "anzahl_dokumente": int(cluster.get("anzahl_dokumente") or 0),
            "rechnung_summe_brutto": round(to_float_safe(cluster.get("rechnung_summe_brutto")), 2),
            "gutschrift_summe_brutto": round(to_float_safe(cluster.get("gutschrift_summe_brutto")), 2),
            "nettoeffekt_brutto": round(to_float_safe(cluster.get("nettoeffekt_brutto")), 2),
            "kostenstruktur": {
                "material_brutto": material,
                "subunternehmer_brutto": subunternehmer,
                "sonstiges_brutto": sonstiges,
            },
            "zugeordnete_rechnungsnummern": unique_nonempty(cluster.get("zugeordnete_rechnungsnummern") or []),
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
                "rechnungsnummer": get_rechnungsnummer(r),
                "lieferant_name": get_lieferant_name(r),
                "faelligkeitsdatum": str(get_faelligkeitsdatum(r) or ""),
                "brutto_summe": round(brutto, 2),
            })

        skonto_prozent = get_skonto_prozent(r)
        skonto_betrag = get_skonto_betrag(r)

        if skonto_prozent > 0 and skonto_betrag > 0 and get_brutto_summe(r) > 0 and is_rechnung(r):
            skonto_chancen.append({
                "rechnung_id": get_rechnung_id(r),
                "rechnungsnummer": get_rechnungsnummer(r),
                "lieferant_name": get_lieferant_name(r),
                "skonto_prozent": skonto_prozent,
                "skonto_betrag": round(skonto_betrag, 2),
                "brutto_summe": round(get_brutto_summe(r), 2),
                "faelligkeitsdatum": str(get_faelligkeitsdatum(r) or ""),
            })

    faellige.sort(key=lambda x: x["brutto_summe"], reverse=True)
    skonto_chancen.sort(key=lambda x: x["skonto_betrag"], reverse=True)

    return {
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

        kategorie = str(
            row.get("lieferantenkategorie")
            or row.get("Lieferantenkategorie")
            or row.get("kategorie")
            or row.get("Kategorie")
            or row.get("typ")
            or row.get("Typ")
            or ""
        ).strip().lower()

        canonical = canonical_lieferanten_kategorie(kategorie)

        entry = {
            "lieferant_id": lieferant_id,
            "lieferant_name": lieferant_name,
            "lieferantenkategorie": canonical,
        }

        if lieferant_id:
            result[lieferant_id] = entry
        if lieferant_name:
            result[lieferant_name] = entry

    return result


def build_gutschrift_details(gutschriften, lookup_rechnungen):
    by_rechnungsnummer = {}
    for r in lookup_rechnungen:
        nr = get_rechnungsnummer(r)
        if nr:
            by_rechnungsnummer[nr] = r

    details = []

    for r in gutschriften:
        referenznummer = get_referenznummer(r)
        ursprung = by_rechnungsnummer.get(referenznummer)

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
            "rechnungsdatum": str(get_rechnungsdatum(r) or ""),
            "faelligkeitsdatum": str(get_faelligkeitsdatum(r) or ""),
            "status": status,
            "ursprungsrechnung_rechnung_id": get_rechnung_id(ursprung) if ursprung else "",
            "ursprungsrechnung_rechnungsnummer": get_rechnungsnummer(ursprung) if ursprung else "",
            "ursprungsrechnung_rechnungsdatum": str(get_rechnungsdatum(ursprung) or "") if ursprung else "",
            "ursprungsrechnung_brutto_summe": round(get_brutto_summe(ursprung), 2) if ursprung else 0.0,
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
    sorted_items = sorted(
        list(projekt_cluster_report or []),
        key=lambda x: x.get("nettoeffekt_brutto", 0),
        reverse=True
    )

    return {
        "anzahl_projekte": len(sorted_items),
        "top_3_nach_nettoeffekt_brutto": sorted_items[:3],
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
            f"Aktuell sind {payment['faellige_rechnungen_anzahl']} Rechnungen mit insgesamt {payment['summe_faellig']:.2f} fällig."
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

@app.route("/analyze", methods=["POST"])
def analyze():
    try:
        data = request.get_json() or {}

        # === META / INPUT MAPPING (angepasst an dein Make JSON) ===
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

        # === HAUPTDATEN ===
        rechnungen = data.get("rechnungen", []) or []
        hinweise = data.get("hinweise", []) or []
        historische_rechnungen = data.get("historische_rechnungen", []) or []

        # === KONTEXT ===
        betriebskontext = data.get("betriebskontext", {}) or {}
        lieferanten_kontext = data.get("lieferanten_kontext", []) or []

        if "kommission_pruefen" not in betriebskontext:
            betriebskontext["kommission_pruefen"] = True

        # === DATUM PARSING ===
        zeitraum_start = parse_date_safe(zeitraum.get("start"))
        zeitraum_ende = parse_date_safe(zeitraum.get("ende"))

        rechnung_map = {}
        for r in rechnungen:
            rid = get_rechnung_id(r)
            if rid:
                rechnung_map[rid] = r

        lookup_rechnungen = list(rechnungen) + list(historische_rechnungen)
        lieferanten_kontext_map = build_lieferanten_kontext_map(lieferanten_kontext)

        fachliche_hinweise = []
        technische_hinweise = []

        for h in hinweise:
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

        rechnungen_nur = [r for r in rechnungen if is_rechnung(r)]
        gutschriften = [r for r in rechnungen if is_gutschrift(r)]

        summary = {
            "rechnungen_gesamt": len(rechnungen_nur),
            "gutschriften_gesamt": len(gutschriften),
            "geprueft": sum(1 for r in rechnungen if is_geprueft(r)),
            "offen": sum(1 for r in rechnungen if is_offen(r)),
            "auffaellig": sum(1 for r in rechnungen_nur if is_rechnung_auffaellig(r)),
            "unauffaellig": sum(1 for r in rechnungen_nur if is_rechnung_unauffaellig(r)),
            "abgelegt": sum(1 for r in rechnungen if is_abgelegt(r)),
            "nicht_abgelegt": sum(1 for r in rechnungen if not is_abgelegt(r)),
            "summe_brutto_rechnungen": round(sum(get_brutto_summe(r) for r in rechnungen_nur), 2),
            "summe_brutto_gutschriften": round(sum(get_brutto_summe(r) for r in gutschriften), 2),
            "summe_brutto_nettoeffekt": round(sum(get_brutto_summe(r) for r in rechnungen), 2),
            "hinweise_fachlich_gesamt": len(fachliche_hinweise),
            "hinweise_technisch_gesamt": len(technische_hinweise),
        }

        fachlicher_breakdown = build_hinweis_breakdown(fachliche_hinweise)
        technischer_breakdown = build_hinweis_breakdown(technische_hinweise)

        payment = build_payment_section(rechnungen_nur, zeitraum_ende)
        top_lieferanten = build_top_lieferanten(rechnungen)

        projekt_cluster = build_project_clusters(
            rechnungen=rechnungen,
            historische_rechnungen=historische_rechnungen,
            lieferanten_kontext_map=lieferanten_kontext_map,
        )

        projekt_cluster_report = build_project_report(
            projekt_cluster=projekt_cluster,
            rechnung_lookup=lookup_rechnungen,
            lieferanten_kontext_map=lieferanten_kontext_map,
        )

        projekt_cluster_report_top10 = sorted(
            projekt_cluster_report,
            key=lambda x: x.get("nettoeffekt_brutto", 0),
            reverse=True
        )[:10]

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

        return jsonify({
            "ok": True,
            "meta": {
                "report_type": mode,
                "betrieb_id": betrieb_id,
                "zeitraum_start": str(zeitraum_start) if zeitraum_start else str(zeitraum.get("start") or ""),
                "zeitraum_ende": str(zeitraum_ende) if zeitraum_ende else str(zeitraum.get("ende") or ""),
                "generated_at": datetime.utcnow().isoformat() + "Z"
            },
            "summary": summary,
            "verarbeitung": {
                "geprueft": summary["geprueft"],
                "offen": summary["offen"],
                "abgelegt": summary["abgelegt"],
                "nicht_abgelegt": summary["nicht_abgelegt"],
            },
            "hinweise": {
                "fachlich_anzahl": len(fachliche_hinweise),
                "technisch_anzahl": len(technische_hinweise),
                "fachlich_breakdown": fachlicher_breakdown,
                "technisch_breakdown": technischer_breakdown,
            },
            "lieferanten": {
                "top_lieferanten": top_lieferanten
            },
            "lieferanten_kontext": {
                "anzahl": len(lieferanten_kontext),
                "eintraege": lieferanten_kontext[:50],
            },
            "gutschriften": {
                "anzahl": len(gutschriften),
                "summe_brutto": round(sum(get_brutto_summe(r) for r in gutschriften), 2),
                "details": gutschrift_details,
            },
            "projekt_cluster": projekt_cluster,
            "projekt_cluster_report": projekt_cluster_report,
            "projekt_report_meta": build_project_report_meta(projekt_cluster_report),
            "payment": payment,
            "report_highlights": {
                "wichtigste_faelle": wichtige_faelle
            },
            "email_summary": email_summary,
            "optional_details": {
                "fachliche_hinweise_details": fachliche_hinweis_details[:25],
                "unklare_projektzuordnungen": unklare_projekte,
                "projekt_cluster_report_top10": projekt_cluster_report_top10,
            },
            "internal_diagnostics": {
                "technische_hinweise_anzahl": len(technische_hinweise),
                "technische_breakdown": technischer_breakdown,
                "rechnungen_mit_fachlichen_hinweisen": len(fachliche_hinweise_by_rechnung),
                "rechnungen_mit_technischen_hinweisen": len(technische_hinweise_by_rechnung),
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
