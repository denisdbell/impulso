#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DNI verification by OCR — compares the PHOTO of the document against the
``Clientes`` sheet and flags anomalies (DNI mismatch, different name, DNI not in
Clientes, unreadable, front/back disagreement, etc.).

It is the Python port of the proven logic in ``VerificacionDNI.gs`` (same Drive
folder, same fields and same anomaly catalog), so results stay consistent with
the «Verificación de DNI» sheet. OCR is **indicative**: it may contain errors and
does NOT replace manual verification of the document.

Flow:
    python3 bcra/fetch_dni.py      # download the photos to out/dni/<person>/…
    python3 bcra/verify_dni.py     # OCR + comparison → Markdown report

The OCR engine is chosen by the OCR_ENGINE env var (``auto`` default):
**Apple Vision** (ocr_vision.py — macOS, much better on phone photos, reads
HEIC, applies EXIF rotation) when available, falling back to **local
Tesseract**. Both are fully local — the images never leave the machine.

    pip install pyobjc-framework-Vision pyobjc-framework-Quartz   # Vision (macOS)
    brew install tesseract tesseract-lang      # Tesseract + Spanish data
    pip install pytesseract pillow

Usage:
    python3 bcra/verify_dni.py [--fetch] [--person <CUIL|DNI|name>] [--export out/_dni.md]

Note: the literals that match the physical Argentine DNI (APELLIDO, NOMBRE,
VENCIMIENTO, the month abbreviations, the `spa` OCR model) and the Clientes
column names are intentionally kept in Spanish — they describe external data.
"""

import argparse
import datetime
import os
import re
import sys
import unicodedata
from pathlib import Path

import analyze_impulso  # norm_dni / norm_cuil / load_tables / build_clients

HERE = Path(__file__).resolve().parent
DEFAULT_PHOTO_DIR = HERE / "out" / "dni"
DEFAULT_XLSX = HERE / "out" / "_impulso.xlsx"
DEFAULT_EXPORT = HERE / "out" / "_dni.md"
OCR_LANG = os.environ.get("OCR_LANG", "spa")

# Status / match-route / anomaly strings (shown to the user).
OK, REVIEW, NO_MATCH = "OK", "Review", "No match"

# Spanish months on the DNI front, for loose expiry parsing (external data).
_MONTHS = {m: i + 1 for i, m in enumerate(
    ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"])}


# --------------------------------------------------------------------------- #
# Normalization helpers (ports of normName_ / normDni_)
# --------------------------------------------------------------------------- #
def norm_name(v):
    """Lowercase, strip accents, collapse whitespace (port of normName_)."""
    s = unicodedata.normalize("NFD", str(v or ""))
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def contains_all_tokens(haystack, needle):
    """True when every token of `needle` (≥2 tokens) appears in `haystack`."""
    toks = [t for t in str(needle or "").split(" ") if t]
    if len(toks) < 2:
        return False
    return all(t in haystack for t in toks)


# --------------------------------------------------------------------------- #
# OCR seam (monkeypatched in tests — the _ocrStub/TEST_MODE equivalent)
# --------------------------------------------------------------------------- #
def _load_ocr():
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        raise RuntimeError(
            "OCR not available. Install the local engine:\n"
            "    brew install tesseract tesseract-lang\n"
            "    pip install pytesseract pillow")
    try:  # iPhone photos arrive as .heic — opt-in decoder (pip install pillow-heif)
        from pillow_heif import register_heif_opener
        register_heif_opener()
    except ImportError:
        pass
    return Image, pytesseract


OCR_ENGINE = os.environ.get("OCR_ENGINE", "auto").lower()  # auto | vision | tesseract


def ocr_image_text(path):
    """OCR an image file to text. Engine per OCR_ENGINE: Apple Vision when
    available (auto/vision), Tesseract otherwise — with Tesseract as the auto
    fallback when Vision reads nothing. An undecodable image yields '' (the
    record gets flagged, the batch keeps going)."""
    if OCR_ENGINE in ("auto", "vision"):
        try:
            import ocr_vision
            if ocr_vision.available():
                text = ocr_vision.ocr_image_text(path)
                if text.strip() or OCR_ENGINE == "vision":
                    return text
        except Exception as exc:  # noqa: BLE001 — Vision hiccup → Tesseract still runs
            print("Vision OCR failed on %s: %s" % (path, exc), file=sys.stderr)
            if OCR_ENGINE == "vision":
                return ""
    Image, pytesseract = _load_ocr()
    try:
        with Image.open(path) as im:
            if (im.format or "").upper() not in ("JPEG", "PNG", "TIFF", "BMP", "GIF", "WEBP"):
                im = im.convert("RGB")  # e.g. HEIF — pytesseract only feeds tesseract its known formats
            try:
                return pytesseract.image_to_string(im, lang=OCR_LANG)
            except pytesseract.TesseractError:
                return pytesseract.image_to_string(im)  # no/invalid lang data → default
    except Exception as exc:  # noqa: BLE001 — corrupt/unsupported file ≠ dead batch
        print("OCR skipped %s: %s" % (path, exc), file=sys.stderr)
        return ""


# --------------------------------------------------------------------------- #
# Field parsers (ports of the .gs parsers — the DNI labels stay in Spanish)
# --------------------------------------------------------------------------- #
def clean_field(s):
    s = re.sub(r"[:|]+", " ", str(s or ""))
    s = re.sub(r"\s+", " ", s).strip()
    return re.sub(r"\d.*$", "", s).strip()


def extract_dni(s, loose=False):
    """A 7-8 digit DNI from text (port of extractDni_)."""
    if not s:
        return ""
    text = str(s)
    best = ""
    for m in re.finditer(r"\b\d{1,2}\.\d{3}\.\d{3}\b", text):      # 32.456.789
        d = re.sub(r"\D", "", m.group(0))
        if 7 <= len(d) <= 8 and len(d) > len(best):
            best = d
    for m in re.finditer(r"\b\d{7,8}\b", text):                    # 32456789
        if len(m.group(0)) > len(best):
            best = m.group(0)
    if loose and not best:
        for m in re.finditer(r"\b\d{1,2} \d{3} \d{3}\b", text):    # 32 456 789
            d = re.sub(r"\D", "", m.group(0))
            if 7 <= len(d) <= 8 and len(d) > len(best):
                best = d
    return best


def _empty_fields():
    return {"apellido": "", "nombre": "", "dni": "", "sexo": "", "nacimiento": "",
            "nacionalidad": "", "cuil": "", "vencimiento": "", "domicilio": ""}


_CUIL_PREFIXES = ("20", "23", "24", "25", "26", "27", "30", "33", "34")


def _valid_cuil(cuil, dni=""):
    """Plausibility: known prefix + (when the DNI is known) it must embed it —
    kills false positives like the 11-digit Trámite N° on the card front."""
    c = re.sub(r"\D", "", str(cuil or ""))
    if len(c) != 11 or c[:2] not in _CUIL_PREFIXES:
        return ""
    d = re.sub(r"\D", "", str(dni or ""))
    if d and c[2:10].lstrip("0") != d.lstrip("0"):
        return ""
    return c


_EN_LABEL_RE = re.compile(
    r"^[\s/|:._-]*(SURNAME|NAME|SEX|NATIONALITY|DATE\s+OF\s+(BIRTH|ISSUE|EXPIRY)|"
    r"EXPIRY(\s+DATE)?|DOCUMENT|IDENTIFICATION)?[\s/|:._-]*$")


def parse_constancia(text):
    """ANSES «Constancia de CUIL» — some clients photograph this instead of the
    card. It still yields the full name (Titular, surname first), DNI (DU) and
    CUIL."""
    out = _empty_fields()
    U = str(text).replace("\r", "").upper()
    lines = [s.strip() for s in U.split("\n") if s.strip()]
    for i, line in enumerate(lines):
        if re.fullmatch(r"TITULAR[:.]?", line) and i + 1 < len(lines):
            toks = [t for t in re.sub(r"[^A-ZÁÉÍÓÚÑÜ ]", " ", lines[i + 1]).split() if t]
            if toks:
                out["apellido"], out["nombre"] = toks[0], " ".join(toks[1:])
            break
    m = re.search(r"\bDU\s*[:.]?\s*(\d{7,8})\b", U)
    if m:
        out["dni"] = m.group(1)
    m = re.search(r"\b(\d{2})-?(\d{8})-?(\d)\b", U)
    if m:
        out["cuil"] = _valid_cuil("".join(m.groups()), out["dni"])
    if not out["dni"] and out["cuil"]:
        out["dni"] = out["cuil"][2:10].lstrip("0")
    return out


def parse_front(text):
    """Parse the front by the DNI's Spanish labels (port of parseDniFront_)."""
    out = _empty_fields()
    if not text:
        return out
    U = str(text).replace("\r", "").upper()
    if "CONSTANCIA DE CUIL" in U:
        return parse_constancia(text)
    lines = [s.strip() for s in U.split("\n") if s.strip()]

    def after_label(re_label):
        for i, line in enumerate(lines):
            m = re.search(re_label, line)
            if m:
                rest = line[m.end():].lstrip(": ").strip()
                # New-format DNI labels are bilingual («APELLIDO / Surname») with the
                # value on the NEXT line — don't mistake the English half for the value.
                if rest and not _EN_LABEL_RE.match(rest):
                    return rest
                if i + 1 < len(lines):
                    return lines[i + 1]
        return ""

    out["apellido"] = clean_field(after_label(r"APELLIDO"))
    out["nombre"] = clean_field(after_label(r"NOMBRES?"))
    sx = after_label(r"SEXO")
    out["sexo"] = "F" if re.search(r"\bF\b|FEMENINO", sx) else ("M" if re.search(r"\bM\b|MASCULINO", sx) else "")
    out["nacionalidad"] = clean_field(after_label(r"NACIONALIDAD"))
    out["nacimiento"] = re.sub(r"\s+", " ", after_label(r"FECHA DE NAC|NACIMIENTO").replace("|", " ")).strip()
    out["vencimiento"] = re.sub(r"\s+", " ", after_label(r"VENCIMIENTO|VENC\.|EXPIRY").replace("|", " ")).strip()
    out["dni"] = extract_dni(after_label(r"DOCUMENTO|D\.?N\.?I\.?|N[º°O]\.?\s*DOC"), loose=True) or extract_dni(U)
    m_cuil = re.search(r"\b(\d{2}-?\d{8}-?\d)\b", U)
    if m_cuil:
        out["cuil"] = _valid_cuil(m_cuil.group(1), out["dni"])
    return out


# Any printed label on either card side — a VALUE never matches this.
_ANY_LABEL_RE = re.compile(
    r"APELLIDO|SURNAME|NOMBRE|NAME|SEXO|\bSEX\b|NACIONALIDAD|NATIONALITY|FECHA|DATE|"
    r"EMISI|ISSUE|VENCIMIENTO|EXPIR|TR[AÁ]MITE|DOCUMENTO|DOCUMENT|EJEMPLAR|"
    r"REP[UÚ]BLICA|MERCOSUR|REGISTRO|MINISTERIO|INTERIOR|FIRMA|SIGNATURE|IDENTIFICADO|"
    r"OF\.?\s*IDENT")


def _value_below(obs, label_re):
    """The text observation sitting right UNDER the matching label, in the same
    column — ID cards are laid out in columns, so plain line order interleaves
    unrelated fields (Vision geometry fixes what next-line parsing can't)."""
    best = None
    for L in obs:
        if not re.search(label_re, L["text"].upper()):
            continue
        for v in obs:
            t = v["text"].strip()
            if not t or v is L or _ANY_LABEL_RE.search(t.upper()):
                continue
            gap = v["y0"] - L["y1"]
            if gap < -0.005 or gap > 0.06:                      # just below the label
                continue
            overlap = min(v["x1"], L["x1"]) - max(v["x0"], L["x0"])
            if overlap <= 0:                                    # same column
                continue
            score = (gap, -overlap)
            if best is None or score < best[0]:
                best = (score, t)
        if best:
            break
    return best[1] if best else ""


def _rot_cw(o):
    """Remap a box as if the card were rotated 90° (card-down == image -x)."""
    return {"text": o["text"], "x0": o["y0"], "x1": o["y1"],
            "y0": 1 - o["x1"], "y1": 1 - o["x0"]}


def _rot_ccw(o):
    return {"text": o["text"], "x0": 1 - o["y1"], "x1": 1 - o["y0"],
            "y0": o["x0"], "y1": o["x1"]}


def _key_label_score(obs):
    return sum(1 for pat in (r"APELLIDO", r"NOMBRE", r"DOCUMENTO", r"NACIMIENTO")
               if _value_below(obs, pat))


def _best_orientation(obs):
    """Photos of the card are often taken sideways; Vision still reads the text
    but the boxes stay in image space. Keep whichever remapping resolves the
    most label→value pairs (ties prefer the untransformed boxes)."""
    variants = [obs, [_rot_cw(o) for o in obs], [_rot_ccw(o) for o in obs]]
    return max(variants, key=_key_label_score)


def parse_front_obs(obs):
    """Geometry-aware front parse (Vision observations): each value is matched
    to the label box it sits under. Regex-based fields (DNI/CUIL) and anything
    geometry can't resolve keep the plain-text parse as the baseline."""
    joined = "\n".join(o["text"] for o in obs)
    if "CONSTANCIA DE CUIL" in joined.upper():   # ANSES certificate, not a card
        return parse_front(joined)
    out = parse_front(joined)
    obs = _best_orientation(obs)
    v = _value_below(obs, r"APELLIDO")
    if v:
        out["apellido"] = clean_field(v) or out["apellido"]
    v = _value_below(obs, r"NOMBRE")
    if v:
        out["nombre"] = clean_field(v) or out["nombre"]
    v = _value_below(obs, r"SEXO")
    if v:
        out["sexo"] = ("F" if re.search(r"\bF\b|FEMENINO", v.upper())
                       else ("M" if re.search(r"\bM\b|MASCULINO", v.upper()) else out["sexo"]))
    v = _value_below(obs, r"NACIONALIDAD")
    if v:
        out["nacionalidad"] = clean_field(v) or out["nacionalidad"]
    v = _value_below(obs, r"NACIMIENTO")
    if v:
        out["nacimiento"] = re.sub(r"\s+", " ", v).strip()
    v = _value_below(obs, r"VENCIMIENTO|EXPIR")
    if v:
        out["vencimiento"] = re.sub(r"\s+", " ", v).strip()
    d = extract_dni(_value_below(obs, r"DOCUMENTO"), loose=True)
    if d:
        out["dni"] = d
    return out


def parse_front_image(path):
    """Engine-aware front parse: geometric with Vision, label text otherwise."""
    if OCR_ENGINE in ("auto", "vision"):
        try:
            import ocr_vision
            if ocr_vision.available():
                obs = ocr_vision.ocr_observations(path)
                if obs:
                    return parse_front_obs(obs)
        except Exception as exc:  # noqa: BLE001 — fall through to the text engines
            print("Vision front parse failed on %s: %s" % (path, exc), file=sys.stderr)
            if OCR_ENGINE == "vision":
                return _empty_fields()
    return parse_front(ocr_image_text(path))


# Lines that END an address value on the card back (next printed label/zone;
# «•» joins locality•country in the birth-place zone — addresses use dashes).
_ADDR_STOP_RE = re.compile(r"LUGAR|NACIM|CUIL|PULGAR|EJEMPLAR|MOTIVO|MINIST|REGISTRO|IDARG|<|•")


def parse_domicilio(text):
    """«DOMICILIO: …» from the back — label-based, keeps digits (addresses have
    numbers, so clean_field does NOT apply), follows up to 2 wrapped lines."""
    if not text:
        return ""
    lines = [s.strip() for s in str(text).replace("\r", "").upper().split("\n") if s.strip()]
    for i, line in enumerate(lines):
        m = re.search(r"DOMICILIO", line)
        if not m:
            continue
        parts = [line[m.end():].lstrip(":. ").strip()]
        for j in range(i + 1, min(i + 3, len(lines))):   # wrapped continuation lines
            if _ADDR_STOP_RE.search(lines[j]):
                break
            parts.append(lines[j])
        val = re.sub(r"\s+", " ", " ".join(p for p in parts if p)).strip(" -·")
        return val
    return ""


def extract_mrz_lines(text):
    """Candidate MRZ lines — dominated by [A-Z0-9<], containing '<' (port)."""
    if not text:
        return []
    res = []
    for s in str(text).upper().split("\n"):
        s = re.sub(r"\s+", "", s)
        if len(s) < 20 or len(s) > 40 or "<" not in s:
            continue
        good = len(re.findall(r"[A-Z0-9<]", s))
        if good >= len(s) * 0.8:
            res.append(s)
    return res


def mrz_date(yymmdd):
    """YYMMDD → raw dd/mm/yy (century not inferred — port of mrzDate_)."""
    if not re.match(r"^\d{6}$", str(yymmdd or "")):
        return ""
    return yymmdd[4:6] + "/" + yymmdd[2:4] + "/" + yymmdd[0:2]


def parse_back(text):
    """Parse the back: TD1 MRZ (3 lines); fall back to front labels (port).
    Domicilio and the labeled CUIL are read from the printed zone either way."""
    out = _empty_fields()
    if not text:
        return out
    domicilio = parse_domicilio(text)
    m_cuil = re.search(r"CUIL[:.\s]*(\d{2})[-\s.]?(\d{7,8})[-\s.]?(\d)\b", str(text).upper())
    cuil_back = (m_cuil.group(1) + m_cuil.group(2).zfill(8) + m_cuil.group(3)) if m_cuil else ""
    mrz = extract_mrz_lines(text)
    if len(mrz) >= 3:
        # OCR may return the three TD1 lines in ANY order (rotated photos) —
        # pick them by content: doc line, dates/sex line, name line.
        l1 = next((s for s in mrz if "IDARG" in s or re.match(r"^ID[A-Z]{3}", s)), mrz[-3])
        l2 = next((s for s in mrz if re.search(r"\d{6}\d?[MFX<]\d{6}", s)), mrz[-2])
        l3 = next((s for s in mrz if s not in (l1, l2) and "<<" in s
                   and not re.search(r"\d{4}", s)), mrz[-1])
        L = [l1, l2, l3]
        parts = l3.split("<<")
        out["apellido"] = re.sub(r"\s+", " ", parts[0].replace("<", " ")).strip() if parts else ""
        out["nombre"] = re.sub(r"\s+", " ", " ".join(parts[1:]).replace("<", " ")).strip()
        m2 = re.search(r"(\d{6})\d?([MFX<])(\d{6})", l2)
        if m2:
            out["nacimiento"] = mrz_date(m2.group(1))
            out["sexo"] = m2.group(2) if m2.group(2) in ("M", "F") else ""
            out["vencimiento"] = mrz_date(m2.group(3))
        m_nac = re.search(r"[MFX<]\d{6}\d?([A-Z]{3})", l2)
        if m_nac:
            out["nacionalidad"] = m_nac.group(1)
        # The doc line glues letters to digits («IDARG31678708<4») — \b-based
        # extract_dni can't split that, so take the first 7-8 digit run.
        m_doc = re.search(r"(\d{7,8})", l1)
        out["dni"] = (m_doc.group(1) if m_doc else "") or extract_dni(" ".join(L).replace("<", " "))
        out["domicilio"] = domicilio
        out["cuil"] = _valid_cuil(cuil_back, out["dni"])
        return out
    out = parse_front(text)
    out["domicilio"] = domicilio
    out["cuil"] = _valid_cuil(cuil_back, out["dni"]) or out["cuil"]
    return out


def merge_fields(front, back):
    """Combine front + back, preferring non-empty; set front_back_mismatch when
    the two DNIs differ (port of mergeDniFields_)."""
    front = front or {}
    back = back or {}

    def pick(a, b):
        a = str(a or "").strip()
        return a or str(b or "").strip()

    merged = {
        "apellido": pick(front.get("apellido"), back.get("apellido")),
        "nombre": pick(front.get("nombre"), back.get("nombre")),
        "dni": "",
        "sexo": pick(front.get("sexo"), back.get("sexo")),
        "nacimiento": pick(front.get("nacimiento"), back.get("nacimiento")),
        "nacionalidad": pick(front.get("nacionalidad"), back.get("nacionalidad")),
        "cuil": pick(back.get("cuil"), front.get("cuil")),      # back has it labeled
        "vencimiento": pick(front.get("vencimiento"), back.get("vencimiento")),
        "domicilio": pick(back.get("domicilio"), front.get("domicilio")),  # lives on the back
        "front_back_mismatch": False,
    }
    df, db = analyze_impulso.norm_dni(front.get("dni")), analyze_impulso.norm_dni(back.get("dni"))
    if df and db:
        merged["dni"] = df
        merged["front_back_mismatch"] = (df != db)
    else:
        merged["dni"] = df or db
    return merged


def expiry_date(s):
    """Best-effort expiry as a date when a 4-digit year is present (MRZ 2-digit
    years are left ambiguous, matching the .gs 'sin inferir el siglo')."""
    if not s:
        return None
    U = str(s).upper()
    m = re.search(r"\b(\d{1,2})[/\-. ](\d{1,2})[/\-. ](\d{4})\b", U)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.search(r"\b(\d{1,2})[/\-. ]([A-Z]{3})[/\-. ](\d{4})\b", U)
        if not m or m.group(2) not in _MONTHS:
            return None
        d, mo, y = int(m.group(1)), _MONTHS[m.group(2)], int(m.group(3))
    try:
        return datetime.date(y, mo, d)
    except ValueError:
        return None


def birth_date(s):
    """V-50 — best-effort birth date from the DNI's «FECHA DE NACIMIENTO» field
    (same DD/MM/YYYY or DD-MMM-YYYY formats as the other dates). None if absent."""
    if not s:
        return None
    U = str(s).upper()
    m = re.search(r"\b(\d{1,2})[/\-. ](\d{1,2})[/\-. ](\d{4})\b", U)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.search(r"\b(\d{1,2})[/\-. ]([A-Z]{3})[/\-. ](\d{4})\b", U)
        if not m or m.group(2) not in _MONTHS:
            return None
        d, mo, y = int(m.group(1)), _MONTHS[m.group(2)], int(m.group(3))
    try:
        return datetime.date(y, mo, d)
    except ValueError:
        return None


def _dhash(path, size=8):
    """V-48 — 64-bit difference hash of an image, or None if it can't be read.
    Robust to resizing/recompression; near-identical images share a low Hamming distance."""
    try:
        from PIL import Image
        try:
            from pillow_heif import register_heif_opener
            register_heif_opener()
        except Exception:
            pass
    except ImportError:
        return None
    try:
        with Image.open(path) as im:
            im = im.convert("L").resize((size + 1, size), Image.BILINEAR)
            px = list(im.getdata())
    except Exception:
        return None
    bits = 0
    for row in range(size):
        base = row * (size + 1)
        for col in range(size):
            bits = (bits << 1) | (1 if px[base + col] > px[base + col + 1] else 0)
    return bits


def _hamming(a, b):
    return bin(a ^ b).count("1")


# --------------------------------------------------------------------------- #
# File grouping (one subfolder = one person)
# --------------------------------------------------------------------------- #
_IMG_EXT = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".heic"}
_SIDE_RE = re.compile(r"\b(dorso|reverso|back|trasero|posterior)\b", re.I)  # Spanish/English side tokens
_STAMP_RE = re.compile(r"(\d{8}[-_ ]\d{6})")


def _side_of(name):
    return "dorso" if _SIDE_RE.search(name) else "frente"


def _stamp_of(name):
    m = _STAMP_RE.search(name)
    return re.sub(r"\D", "", m.group(1)) if m else ""


def _persona_and_dni(subfolder_name):
    """Split 'Facundo Acosta 42886617' → ('Facundo Acosta', '42886617')."""
    m = re.search(r"(\d{7,8})\s*$", subfolder_name.strip())
    dni = m.group(1) if m else ""
    persona = subfolder_name[:m.start()].strip() if m else subfolder_name.strip()
    persona = re.sub(r"\b(dni|cuil)\b", " ", persona, flags=re.I)
    return re.sub(r"\s+", " ", persona).strip(), dni


def group_photos(photo_dir):
    """Walk photo_dir/<subfolder>/* → one group per subfolder, keeping the most
    recent front and back (filename timestamp, then mtime)."""
    photo_dir = Path(photo_dir)
    groups = []
    if not photo_dir.exists():
        return groups
    for sub in sorted([p for p in photo_dir.iterdir() if p.is_dir()], key=lambda p: p.name.lower()):
        persona, dni_folder = _persona_and_dni(sub.name)
        g = {"subfolder": sub.name, "persona": persona or sub.name, "dni_folder": dni_folder,
             "frente": None, "dorso": None}
        best = {"frente": (-1, ""), "dorso": (-1, "")}  # (mtime, stamp)
        for f in sub.iterdir():
            if not f.is_file() or f.suffix.lower() not in _IMG_EXT:
                continue
            if not any(k in f.name.lower() for k in ("dni", "cuil")):
                continue
            side = _side_of(f.name)
            try:
                mtime = f.stat().st_mtime
            except OSError:
                mtime = 0
            stamp = _stamp_of(f.name)
            cur = best[side]
            if (mtime, stamp) > cur:
                best[side] = (mtime, stamp)
                g[side] = f
        if g["frente"] or g["dorso"]:
            groups.append(g)
    return groups


# --------------------------------------------------------------------------- #
# Matching against Clientes + anomaly catalog
# --------------------------------------------------------------------------- #
def match_cliente(fields, persona_from_file, dni_folder, clients):
    """Match the scanned document to a client (port of matchClienteForDni_).
    Primary by DNI (OCR, else the folder DNI); fallback by name."""
    by_dni = {}
    for c in clients.values():
        if c.get("dni") and c["dni"] not in by_dni:
            by_dni[c["dni"]] = c
    dn = analyze_impulso.norm_dni(fields.get("dni")) or analyze_impulso.norm_dni(dni_folder)
    if dn and dn in by_dni:
        via = "By DNI" if analyze_impulso.norm_dni(fields.get("dni")) else "By folder"
        return {"cliente": by_dni[dn], "via": via, "ambiguo": False}

    by_file = norm_name(persona_from_file)
    by_ocr = norm_name(" ".join([fields.get("apellido", ""), fields.get("nombre", "")]).strip())

    def match_name(target):
        if not target:
            return []
        hits = []
        for o in clients.values():
            n = norm_name(o.get("name"))
            if n and (n == target or contains_all_tokens(n, target) or contains_all_tokens(target, n)):
                hits.append(o)
        return hits

    lst = match_name(by_file) or match_name(by_ocr)
    uniq = {}
    for o in lst:
        uniq[o["client_id"]] = o
    dedup = list(uniq.values())
    if len(dedup) == 1:
        return {"cliente": dedup[0], "via": "By name (file)", "ambiguo": False}
    if len(dedup) > 1:
        return {"cliente": None, "via": NO_MATCH, "ambiguo": True}
    return {"cliente": None, "via": NO_MATCH, "ambiguo": False}


def anomalies(fields, match, group, dni_persona_counts, today, photo_dup=False):
    """Anomaly catalog → {estado, labels} (port of dniAnomalies_ + extras)."""
    labels = []
    dn_ocr = analyze_impulso.norm_dni(fields.get("dni"))
    eff = dn_ocr or analyze_impulso.norm_dni(group.get("dni_folder"))
    cli = match.get("cliente")

    if not eff:
        labels.append("Unreadable DNI")
    if match.get("ambiguo"):
        labels.append("Ambiguous name match")

    if not cli:
        if eff and not match.get("ambiguo"):
            labels.append("DNI not found in Clientes")
    else:
        if eff and cli.get("dni") and eff != cli["dni"]:
            labels.append("DNI mismatch")
        ocr_name = norm_name(" ".join([fields.get("apellido", ""), fields.get("nombre", "")]).strip())
        cli_name = norm_name(cli.get("name"))
        if (ocr_name and cli_name and ocr_name != cli_name
                and not contains_all_tokens(cli_name, ocr_name)
                and not contains_all_tokens(ocr_name, cli_name)):
            labels.append("Name mismatch")

    if fields.get("front_back_mismatch"):
        labels.append("Front and back mismatch")
    if not group.get("frente"):
        labels.append("Missing front")
    if not group.get("dorso"):
        labels.append("Missing back")
    if not str(fields.get("apellido", "")).strip() or not str(fields.get("nombre", "")).strip():
        labels.append("Incomplete OCR")
    if eff and dni_persona_counts.get(eff, 0) > 1:
        labels.append("DNI shared across people")
    if expiry_date(fields.get("vencimiento")) and expiry_date(fields.get("vencimiento")) < today:
        labels.append("Document expired")
    # V-48 — the same DNI photo appears under a different person
    if photo_dup:
        labels.append("Duplicate DNI photo (same image as another person)")
    # V-50 — age from the birth date: under 18 is ineligible; an implausible age hints at a bad read/forgery
    bd = birth_date(fields.get("nacimiento"))
    if bd:
        age = (today - bd).days // 365
        if age < 18:
            labels.append("Underage (under 18)")
        elif age > 100:
            labels.append("Implausible age (%d)" % age)

    return {"estado": REVIEW if labels else OK, "labels": labels}


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def load_clients(xlsx_path, today=None):
    today = today or datetime.date.today()
    tables = analyze_impulso.load_tables(str(xlsx_path))
    return analyze_impulso.build_clients(tables, today)


def group_signature(g):
    """Fingerprint of a group's chosen photos (name + mtime): when unchanged, a
    cached OCR result can be reused instead of re-scanning."""
    parts = []
    for side in ("frente", "dorso"):
        f = g.get(side)
        if f:
            try:
                parts.append("%s:%d" % (f.name, int(f.stat().st_mtime)))
            except OSError:
                parts.append(f.name)
    return "|".join(parts)


def build_verifications(photo_dir, clients, today=None, progress=lambda m: None,
                        fields_cache=None):
    """OCR every scanned subfolder and compare to Clientes → list of records.
    `fields_cache` (optional, MUTATED): {subfolder: {"sig", "fields"}} — groups
    whose photo signature is unchanged reuse the cached OCR fields (matching and
    anomalies are always recomputed against the current Clientes)."""
    today = today or datetime.date.today()
    groups = group_photos(photo_dir)
    progress("Reading %d DNI folder(s)…" % len(groups))

    # Pass 1 — OCR + parse all groups (so DNI-shared-across-people can be computed).
    for i, g in enumerate(groups, 1):
        sig = group_signature(g)
        cached = (fields_cache or {}).get(g["subfolder"])
        if cached and cached.get("sig") == sig and isinstance(cached.get("fields"), dict):
            g["_fields"] = dict(_empty_fields(), **cached["fields"])
            g["_cached"] = True
            continue
        progress("(%d/%d) %s" % (i, len(groups), g["persona"]))
        front = parse_front_image(g["frente"]) if g["frente"] else _empty_fields()
        back = parse_back(ocr_image_text(g["dorso"])) if g["dorso"] else _empty_fields()
        g["_fields"] = merge_fields(front, back)
        if fields_cache is not None:
            fields_cache[g["subfolder"]] = {"sig": sig, "fields": g["_fields"]}

    # how many distinct people each effective DNI appears under
    dni_persona = {}
    for g in groups:
        eff = analyze_impulso.norm_dni(g["_fields"].get("dni")) or analyze_impulso.norm_dni(g["dni_folder"])
        if eff:
            dni_persona.setdefault(eff, set()).add(norm_name(g["persona"]))
    dni_persona_counts = {k: len(v) for k, v in dni_persona.items()}

    # V-48 — perceptual-hash the photos; the SAME image reused by a different person is a strong lead.
    # Cheap next to OCR; O(G²) Hamming per side over ≤ a few hundred groups.
    for g in groups:
        g["_ph_frente"] = _dhash(g["frente"]) if g["frente"] else None
        g["_ph_dorso"] = _dhash(g["dorso"]) if g["dorso"] else None
    dup_photo = set()   # subfolders whose front/back matches a DIFFERENT person's
    PH_MAX = 6          # Hamming distance ≤ this on a 64-bit dHash = effectively the same image
    for side in ("_ph_frente", "_ph_dorso"):
        hashed = [g for g in groups if g.get(side) is not None]
        for i in range(len(hashed)):
            for j in range(i + 1, len(hashed)):
                a, b = hashed[i], hashed[j]
                if norm_name(a["persona"]) == norm_name(b["persona"]):
                    continue    # same person across sessions → not a duplicate
                if _hamming(a[side], b[side]) <= PH_MAX:
                    dup_photo.add(a["subfolder"])
                    dup_photo.add(b["subfolder"])

    # Pass 2 — match + anomalies → records
    records = []
    for g in groups:
        fields = g["_fields"]
        match = match_cliente(fields, g["persona"], g["dni_folder"], clients)
        an = anomalies(fields, match, g, dni_persona_counts, today, photo_dup=g["subfolder"] in dup_photo)
        cli = match["cliente"]
        estado = an["estado"] if cli else NO_MATCH
        records.append({
            "persona": g["persona"], "subfolder": g["subfolder"],
            "dni_ocr": fields.get("dni", ""), "dni_folder": g["dni_folder"],
            "apellido": fields.get("apellido", ""), "nombre": fields.get("nombre", ""),
            "sexo": fields.get("sexo", ""), "nacimiento": fields.get("nacimiento", ""),
            "nacionalidad": fields.get("nacionalidad", ""), "cuil": fields.get("cuil", ""),
            "vencimiento": fields.get("vencimiento", ""), "domicilio": fields.get("domicilio", ""),
            "front_back_mismatch": fields.get("front_back_mismatch", False),
            "cliente_id": cli["client_id"] if cli else "",
            "cliente_name": cli["name"] if cli else "",
            "cliente_dni": cli["dni"] if cli else "",
            "cliente_cuil": cli["cuil"] if cli else "",
            "via": match["via"], "estado": estado, "anomalias": an["labels"],
            "frente": str(g["frente"]) if g["frente"] else "",
            "dorso": str(g["dorso"]) if g["dorso"] else "",
            "ocr_cached": bool(g.get("_cached")),
        })
    records.sort(key=lambda r: ({REVIEW: 0, NO_MATCH: 1, OK: 2}.get(r["estado"], 3), r["persona"].lower()))
    return records


# --------------------------------------------------------------------------- #
# Rendering
# --------------------------------------------------------------------------- #
_DISCLAIMER = ("_OCR reading is **indicative**: it may contain errors and does not "
               "replace manual verification of the document._")


def render_md(records):
    gen = datetime.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M")
    n_rev = sum(1 for r in records if r["estado"] != OK)
    L = ["# DNI Verification (OCR)", "",
         "_Generated %s · %d document(s), %d to review._" % (gen, len(records), n_rev), "",
         _DISCLAIMER, "",
         "| Person | DNI (doc) | DNI (client) | Name (doc) | Name (client) | Match | Status | Anomalies |",
         "|---|---|---|---|---|---|---|---|"]
    for r in records:
        doc_name = (" ".join([r["apellido"], r["nombre"]]).strip()) or "—"
        L.append("| %s | %s | %s | %s | %s | %s | %s | %s |" % (
            r["persona"] or "—", r["dni_ocr"] or r["dni_folder"] or "—", r["cliente_dni"] or "—",
            doc_name, r["cliente_name"] or "—", r["via"], r["estado"],
            ", ".join(r["anomalias"]) or "—"))
    return "\n".join(L)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def parse_args(argv):
    p = argparse.ArgumentParser(
        description="OCR clients' DNI photos and compare them to the Clientes sheet.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("--fetch", action="store_true", help="download the photos first (fetch_dni.py)")
    p.add_argument("--person", help="only report one person (matches person name or DNI)")
    p.add_argument("--photos", default=str(DEFAULT_PHOTO_DIR), help="folder with the DNI photos")
    p.add_argument("--xlsx", default=str(DEFAULT_XLSX), help="Impulso workbook (.xlsx export)")
    p.add_argument("--export", nargs="?", const=str(DEFAULT_EXPORT), help="also write the report to this .md")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    if args.fetch:
        import fetch_dni
        cred = fetch_dni.fetch_impulso.resolve_credentials(None, HERE)
        folder = fetch_dni.resolve_folder_id(None, HERE)
        fetch_dni.download_all(args.photos, cred, folder, progress=print)

    clients = load_clients(args.xlsx)
    records = build_verifications(args.photos, clients, progress=lambda m: print(m, file=sys.stderr))
    if args.person:
        key = args.person.strip().lower()
        digits = re.sub(r"\D", "", args.person)
        records = [r for r in records
                   if key in r["persona"].lower()
                   or (digits and digits in (r["dni_ocr"], r["dni_folder"], r["cliente_dni"]))]
    text = render_md(records)
    print(text)
    if args.export:
        Path(args.export).parent.mkdir(parents=True, exist_ok=True)
        Path(args.export).write_text(text + "\n", encoding="utf-8")
        print("\nSaved %s" % args.export, file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
