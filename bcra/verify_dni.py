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

The OCR engine is **local Tesseract** (the images never leave the machine):

    brew install tesseract tesseract-lang      # binary + Spanish data
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
    return Image, pytesseract


def ocr_image_text(path):
    """OCR an image file to text with Tesseract. Tries the Spanish model, falling
    back to the default language if `spa` data isn't installed."""
    Image, pytesseract = _load_ocr()
    with Image.open(path) as im:
        try:
            return pytesseract.image_to_string(im, lang=OCR_LANG)
        except pytesseract.TesseractError:
            return pytesseract.image_to_string(im)  # no/invalid lang data → default


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
            "nacionalidad": "", "cuil": "", "vencimiento": ""}


def parse_front(text):
    """Parse the front by the DNI's Spanish labels (port of parseDniFront_)."""
    out = _empty_fields()
    if not text:
        return out
    U = str(text).replace("\r", "").upper()
    lines = [s.strip() for s in U.split("\n") if s.strip()]

    def after_label(re_label):
        for i, line in enumerate(lines):
            m = re.search(re_label, line)
            if m:
                rest = line[m.end():].lstrip(": ").strip()
                if rest:
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
        out["cuil"] = re.sub(r"\D", "", m_cuil.group(1))
    return out


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
    """Parse the back: TD1 MRZ (3 lines); fall back to front labels (port)."""
    out = _empty_fields()
    if not text:
        return out
    mrz = extract_mrz_lines(text)
    if len(mrz) >= 3:
        L = mrz[-3:]
        l1, l2, l3 = L[0], L[1], L[2]
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
        out["dni"] = extract_dni(l1.replace("<", " ")) or extract_dni(" ".join(L).replace("<", " "))
        return out
    return parse_front(text)


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
        "cuil": pick(front.get("cuil"), back.get("cuil")),
        "vencimiento": pick(front.get("vencimiento"), back.get("vencimiento")),
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


def anomalies(fields, match, group, dni_persona_counts, today):
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

    return {"estado": REVIEW if labels else OK, "labels": labels}


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def load_clients(xlsx_path, today=None):
    today = today or datetime.date.today()
    tables = analyze_impulso.load_tables(str(xlsx_path))
    return analyze_impulso.build_clients(tables, today)


def build_verifications(photo_dir, clients, today=None, progress=lambda m: None):
    """OCR every scanned subfolder and compare to Clientes → list of records."""
    today = today or datetime.date.today()
    groups = group_photos(photo_dir)
    progress("Reading %d DNI folder(s)…" % len(groups))

    # Pass 1 — OCR + parse all groups (so DNI-shared-across-people can be computed).
    for i, g in enumerate(groups, 1):
        progress("(%d/%d) %s" % (i, len(groups), g["persona"]))
        front = parse_front(ocr_image_text(g["frente"])) if g["frente"] else _empty_fields()
        back = parse_back(ocr_image_text(g["dorso"])) if g["dorso"] else _empty_fields()
        g["_fields"] = merge_fields(front, back)

    # how many distinct people each effective DNI appears under
    dni_persona = {}
    for g in groups:
        eff = analyze_impulso.norm_dni(g["_fields"].get("dni")) or analyze_impulso.norm_dni(g["dni_folder"])
        if eff:
            dni_persona.setdefault(eff, set()).add(norm_name(g["persona"]))
    dni_persona_counts = {k: len(v) for k, v in dni_persona.items()}

    # Pass 2 — match + anomalies → records
    records = []
    for g in groups:
        fields = g["_fields"]
        match = match_cliente(fields, g["persona"], g["dni_folder"], clients)
        an = anomalies(fields, match, g, dni_persona_counts, today)
        cli = match["cliente"]
        estado = an["estado"] if cli else NO_MATCH
        records.append({
            "persona": g["persona"], "subfolder": g["subfolder"],
            "dni_ocr": fields.get("dni", ""), "dni_folder": g["dni_folder"],
            "apellido": fields.get("apellido", ""), "nombre": fields.get("nombre", ""),
            "sexo": fields.get("sexo", ""), "nacimiento": fields.get("nacimiento", ""),
            "nacionalidad": fields.get("nacionalidad", ""), "cuil": fields.get("cuil", ""),
            "vencimiento": fields.get("vencimiento", ""),
            "front_back_mismatch": fields.get("front_back_mismatch", False),
            "cliente_id": cli["client_id"] if cli else "",
            "cliente_name": cli["name"] if cli else "",
            "cliente_dni": cli["dni"] if cli else "",
            "cliente_cuil": cli["cuil"] if cli else "",
            "via": match["via"], "estado": estado, "anomalias": an["labels"],
            "frente": str(g["frente"]) if g["frente"] else "",
            "dorso": str(g["dorso"]) if g["dorso"] else "",
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
