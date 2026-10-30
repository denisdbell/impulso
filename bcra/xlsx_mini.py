#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Minimal, dependency-free .xlsx reader (Python std-lib only: zipfile + xml).

Just enough of the OOXML SpreadsheetML format to pull tabular data out of a
Google-Sheets "Download ▸ Microsoft Excel (.xlsx)" export: sheet names in
workbook order, the shared-string table, cell values (number / shared-string /
inline-string / boolean / formula result), and date detection (serial →
datetime) via each cell's number format.

This is NOT a general xlsx library — no styles, merges, or formula evaluation.
It exists so the rest of `bcra/` stays std-lib only (like fetch_bcra.py /
analyze_bcra.py) instead of pulling in openpyxl. Correctness is validated against
openpyxl on the real workbook (see README ▸ "Historial Impulso").

read_xlsx(path) -> dict: {sheet_name: [row, ...]} where each row is a list of
cell values (str | float | int | datetime.datetime | bool | None), padded to the
sheet's widest row so column indices are stable.
"""

import datetime
import re
import zipfile
import xml.etree.ElementTree as ET

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKGREL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

# Built-in numFmtId values that denote dates/times (ECMA-376 §18.8.30 + the
# East-Asian date formats). Everything else is a date only if its custom format
# code contains date tokens (see _code_is_date).
_BUILTIN_DATE_FMTS = {14, 15, 16, 17, 18, 19, 20, 21, 22,
                      27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
                      45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58}

# Excel's day 0 is 1899-12-30 (the 1900 date system, including its Feb-29-1900
# leap-year bug which only matters for serials ≤ 60 — never real dates here).
_EPOCH = datetime.datetime(1899, 12, 30)


def _code_is_date(code):
    """True if a number-format code renders a date/time (heuristic, openpyxl-style)."""
    if not code:
        return False
    s = code.split(";")[0]                 # first (positive) section
    s = re.sub(r'\[[^\]]*\]', '', s)       # [Red], [$-409], [h] color/locale/elapsed
    s = re.sub(r'"[^"]*"', '', s)          # "literal text"
    s = re.sub(r'\\.', '', s)              # \escaped char
    s = s.lower()
    return any(ch in s for ch in "ymdh")


def _col_to_idx(ref):
    """'B7' / 'AB' -> zero-based column index."""
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _serial_to_dt(serial):
    """Excel serial -> datetime; returns the raw float if the serial is out of range
    (some Google exports carry corrupt serials in unused study tabs)."""
    try:
        dt = _EPOCH + datetime.timedelta(days=float(serial))
    except (OverflowError, ValueError):
        return float(serial)
    if dt.year < 1900 or dt.year > 9999:
        return float(serial)
    # Whole-day serial -> keep a midnight datetime (callers use isinstance + .date()).
    return dt


def _text_of(node):
    """Concatenate every <t> under a <si>/<is> node (handles rich-text runs)."""
    parts = []
    for t in node.iter(_MAIN + "t"):
        parts.append(t.text or "")
    return "".join(parts)


def read_xlsx(path):
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())

        # --- shared strings ---
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(_MAIN + "si"):
                shared.append(_text_of(si))

        # --- styles: which cellXfs indices are dates ---
        date_style = []
        if "xl/styles.xml" in names:
            st = ET.fromstring(z.read("xl/styles.xml"))
            custom = {}
            numfmts = st.find(_MAIN + "numFmts")
            if numfmts is not None:
                for nf in numfmts.findall(_MAIN + "numFmt"):
                    custom[int(nf.get("numFmtId"))] = nf.get("formatCode") or ""
            cellxfs = st.find(_MAIN + "cellXfs")
            if cellxfs is not None:
                for xf in cellxfs.findall(_MAIN + "xf"):
                    fid = int(xf.get("numFmtId") or 0)
                    is_date = fid in _BUILTIN_DATE_FMTS or _code_is_date(custom.get(fid))
                    date_style.append(is_date)

        # --- workbook: sheet name -> r:id, in document (tab) order ---
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheet_rids = []
        for sh in wb.find(_MAIN + "sheets").findall(_MAIN + "sheet"):
            sheet_rids.append((sh.get("name"), sh.get(_R + "id")))

        # --- rels: r:id -> worksheet part path ---
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_target = {}
        for rel in rels.findall(_PKGREL + "Relationship"):
            tgt = rel.get("Target")
            if tgt.startswith("/"):
                part = tgt.lstrip("/")
            else:
                part = "xl/" + tgt if not tgt.startswith("xl/") else tgt
            rid_target[rel.get("Id")] = part

        out = {}
        for name, rid in sheet_rids:
            part = rid_target.get(rid)
            if not part or part not in names:
                out[name] = []
                continue
            out[name] = _read_sheet(ET.fromstring(z.read(part)), shared, date_style)
        return out


def _read_sheet(ws, shared, date_style):
    data = ws.find(_MAIN + "sheetData")
    if data is None:
        return []
    rows = []
    maxw = 0
    for row in data.findall(_MAIN + "row"):
        cells = {}
        for c in row.findall(_MAIN + "c"):
            ref = c.get("r") or ""
            ci = _col_to_idx(ref) if ref else len(cells)
            t = c.get("t")
            s = c.get("s")
            val = None
            if t == "inlineStr":
                isn = c.find(_MAIN + "is")
                val = _text_of(isn) if isn is not None else None
            else:
                v = c.find(_MAIN + "v")
                raw = v.text if v is not None else None
                if raw is None:
                    val = None
                elif t == "s":
                    val = shared[int(raw)] if raw.isdigit() and int(raw) < len(shared) else ""
                elif t == "b":
                    val = (raw == "1")
                elif t in ("str", "e"):
                    val = raw                      # formula string / error text
                else:                              # number (t None or "n")
                    try:
                        num = float(raw)
                    except ValueError:
                        val = raw
                    else:
                        si = int(s) if (s is not None and s.isdigit()) else None
                        if si is not None and si < len(date_style) and date_style[si]:
                            val = _serial_to_dt(num)
                        elif num == int(num):
                            val = int(num)
                        else:
                            val = num
            cells[ci] = val
        if cells:
            maxw = max(maxw, max(cells) + 1)
        rows.append(cells)
    # Pad every row to the widest observed column.
    return [[r.get(i) for i in range(maxw)] for r in rows]
