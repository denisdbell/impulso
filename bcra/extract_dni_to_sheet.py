#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extract each client's DNI data from the photos in the shared Drive folder and
write it to a NEW TAB of the live Impulso spreadsheet, one row per document set,
linked to the client:

  «Datos DNI (OCR)»  —  ID Cliente (clickable: jumps to the client's row in
  «Clientes»), the OCR-extracted fields (DNI, apellido, nombres, sexo,
  nacimiento, nacionalidad, CUIL, vencimiento), the verification status and
  anomalies, plus links to the person's Drive folder and both photos.

Reuses the proven pipeline: fetch_dni.py (download the photos from the folder
that has «DNI» in the file names), verify_dni.py (local Tesseract OCR + match
against Clientes), fetch_impulso.py (service-account auth + spreadsheet id).
Sheet headers are Spanish to match the workbook (locale es-AR).

Requirements (same as the other fetch tools — use a python that has them):

    pip install google-auth requests pytesseract pillow
    brew install tesseract tesseract-lang

One-time setup (both steps, otherwise the write fails with a clear 403):
  1. Enable the **Google Sheets API** in the service account's Cloud project
     (console.developers.google.com → APIs → Google Sheets API → Enable).
  2. From the sheet OWNER's account, share the spreadsheet with the
     service-account address ("client_email" in service_account.json) as
     **Editor** — Viewer is enough to read, but this tool WRITES the new tab.

Usage:
    python3 bcra/extract_dni_to_sheet.py                # OCR cache → write sheet
    python3 bcra/extract_dni_to_sheet.py --fetch        # download new photos first
    python3 bcra/extract_dni_to_sheet.py --incremental  # only OCR new/changed folders
    python3 bcra/extract_dni_to_sheet.py --dry-run      # print table, write nothing
    python3 bcra/extract_dni_to_sheet.py --csv          # also save out/_dni_datos.csv
    python3 bcra/extract_dni_to_sheet.py --no-links     # offline: skip Drive links

Also callable in-process via run(...) — used by the dashboard menu («Tools ▸
Extract DNI data») and the hourly cron wrapper (cron_extract_dni.sh).

OCR is indicative — the sheet includes the status/anomaly columns so every row
can be reviewed by eye against the linked photos.
"""

import argparse
import csv
import datetime
import json
import logging
import re
import sys
from pathlib import Path

import fetch_dni      # Drive listing/download + folder id resolution
import fetch_impulso  # service-account credentials + spreadsheet id resolution
import verify_dni     # OCR, field parsing, client matching

HERE = Path(__file__).resolve().parent
DEFAULT_XLSX = HERE / "out" / "_impulso.xlsx"
DEFAULT_PHOTOS = HERE / "out" / "dni"
DEFAULT_SHEET = "Datos DNI (OCR)"
CLIENTES_SHEET = "Clientes"
# OCR-result cache for --incremental: {subfolder: {"sig", "fields"}} — folders
# whose photos are unchanged are NOT re-scanned (matching is always recomputed).
CACHE_FILE = HERE / "out" / "_dni_ocr_cache.json"

# Write scope for the new tab + read-only Drive for the photo links.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive.readonly"]
SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets/%s"

# Spanish headers — the sheet lives inside the es-AR production workbook.
HEADERS = ["ID Cliente", "Cliente", "Estado", "Coincidencia", "Anomalías",
           "DNI (OCR)", "Apellido (OCR)", "Nombres (OCR)", "Sexo", "Nacimiento",
           "Nacionalidad", "CUIL (OCR)", "Vencimiento",
           "Domicilio (OCR)", "Dirección declarada",
           "Carpeta en Drive", "Frente", "Dorso", "Actualizado"]


# --------------------------------------------------------------------------- #
# auth + Drive photo links
# --------------------------------------------------------------------------- #
def _session(cred_path):
    try:
        from google.oauth2 import service_account
        from google.auth.transport.requests import AuthorizedSession
    except ImportError:
        raise SystemExit("Missing dependencies. Install with:\n    pip install google-auth requests")
    creds = service_account.Credentials.from_service_account_file(str(cred_path), scopes=SCOPES)
    return AuthorizedSession(creds)


def drive_photo_index(session, folder_id, timeout=60, progress=lambda m: None):
    """Metadata-only listing: subfolder name → {'_folder': id, filename: fileId}.
    Lets every sheet row link straight to the person's folder and photos."""
    idx = {}
    subs = fetch_dni._list_children(session, folder_id, timeout, only_folders=True)
    for i, sub in enumerate(subs, 1):
        name = sub.get("name") or sub["id"]
        progress("(%d/%d) links %s" % (i, len(subs), name))
        entry = {"_folder": sub["id"]}
        try:
            for f in fetch_dni._list_children(session, sub["id"], timeout):
                entry[f["name"]] = f["id"]
        except SystemExit as exc:         # one unreadable subfolder must not kill the run
            logging.warning("link listing failed for %s: %s", name, exc)
        idx[name] = entry
    return idx


# --------------------------------------------------------------------------- #
# Sheets API helpers (REST via AuthorizedSession — no extra client library)
# --------------------------------------------------------------------------- #
def _check(resp, what, cred_path):
    if resp.status_code == 403:
        if "SERVICE_DISABLED" in resp.text:
            raise SystemExit(
                "HTTP 403 on %s — the Google SHEETS API is disabled in the service account's "
                "Cloud project.\nEnable it (one click, logged in as the account that created "
                "service_account.json), wait ~2 min, retry:\n"
                "    https://console.developers.google.com/apis/api/sheets.googleapis.com/overview"
                % what)
        raise SystemExit(
            "HTTP 403 on %s — the service account can't write the spreadsheet.\n"
            "From the sheet OWNER's account, share it as **Editor** with:\n"
            "    %s" % (what, fetch_impulso._client_email(cred_path)))
    if resp.status_code == 404:
        raise SystemExit("HTTP 404 on %s — the service account can't see the spreadsheet. "
                         "Share it with the client_email in service_account.json." % what)
    if resp.status_code != 200:
        raise SystemExit("%s failed (HTTP %d): %s" % (what, resp.status_code, resp.text[:300]))
    return resp.json()


def sheet_gids(session, ssid, cred_path, timeout=60):
    data = _check(session.get(SHEETS_API % ssid,
                              params={"fields": "sheets(properties(sheetId,title))"},
                              timeout=timeout), "spreadsheet metadata", cred_path)
    return {s["properties"]["title"]: s["properties"]["sheetId"] for s in data.get("sheets", [])}


def get_values(session, ssid, rng, cred_path, timeout=60):
    data = _check(session.get((SHEETS_API % ssid) + "/values/" + rng, timeout=timeout),
                  "read %s" % rng, cred_path)
    return data.get("values", [])


def _col_letter(i):
    """0-based column index → A1 letter(s)."""
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def clientes_live_rows(session, ssid, cred_path, timeout=60):
    """Live «Clientes»: ID Cliente → row number (for the per-client hyperlink).
    Read live (not from the xlsx) so links survive row moves in the sheet."""
    hdr = get_values(session, ssid, "'%s'!1:1" % CLIENTES_SHEET, cred_path, timeout)
    row0 = hdr[0] if hdr else []
    try:
        ci = row0.index("ID Cliente")
    except ValueError:
        return {}
    col = _col_letter(ci)
    vals = get_values(session, ssid, "'%s'!%s2:%s" % (CLIENTES_SHEET, col, col), cred_path, timeout)
    return {str(v[0]).strip(): i + 2 for i, v in enumerate(vals) if v and str(v[0]).strip()}


# --------------------------------------------------------------------------- #
# row building
# --------------------------------------------------------------------------- #
def _hyperlink(url, label):
    """es-AR locale: formula arguments separate with ';' (',' is the decimal mark)."""
    label = str(label if label not in (None, "") else url).replace('"', "'")
    return '=HYPERLINK("%s";"%s")' % (url, label)


_HL_RE = re.compile(r'^=HYPERLINK\("([^"]+)";"([^"]*)"\)$')


def _plain(cell):
    """CSV version of a cell: 'label (url)'; in-sheet #gid links → just the label."""
    m = _HL_RE.match(str(cell))
    if not m:
        return cell
    url, label = m.group(1), m.group(2)
    return label if url.startswith("#") else "%s (%s)" % (label, url)


def _sort_key(r):
    m = re.search(r"(\d+)", r["cliente_id"] or "")
    return (0, int(m.group(1))) if m else (1, 10 ** 9), (r["persona"] or "").lower()


def declared_addresses(xlsx_path):
    """«Clientes» ID Cliente → Dirección declarada (for eyeball comparison with
    the OCR'd domicilio). Last occurrence of a duplicated header wins."""
    import xlsx_mini
    rows = xlsx_mini.read_xlsx(str(xlsx_path)).get(CLIENTES_SHEET, [])
    if not rows:
        return {}
    idx = {}
    for i, h in enumerate(rows[0]):
        h = str(h or "").strip()
        if h:
            idx[h] = i
    ci, di = idx.get("ID Cliente"), idx.get("Dirección")
    if ci is None or di is None:
        return {}
    out = {}
    for r in rows[1:]:
        cid = str(r[ci] if ci < len(r) else "").strip()
        if cid:
            out[cid] = str(r[di] if di < len(r) and r[di] is not None else "").strip()
    return out


def build_rows(records, photo_idx, client_rows, clientes_gid, stamp, decl_addr=None):
    """records (verify_dni) → one sheet row per scanned document set, sorted by
    client id (unmatched folders last)."""
    rows = []
    for r in sorted(records, key=_sort_key):
        sub = photo_idx.get(r["subfolder"], {})
        cid = r["cliente_id"]
        if cid and cid in client_rows and clientes_gid is not None:
            id_cell = _hyperlink("#gid=%d&range=A%d" % (clientes_gid, client_rows[cid]), cid)
        else:
            id_cell = cid or "—"
        folder_cell = (_hyperlink("https://drive.google.com/drive/folders/" + sub["_folder"],
                                  r["subfolder"]) if sub.get("_folder") else r["subfolder"])

        def photo_cell(path, label):
            if not path:
                return "—"
            fid = sub.get(Path(path).name, "")
            return _hyperlink("https://drive.google.com/file/d/%s/view" % fid, label) if fid else label

        rows.append([
            id_cell, r["cliente_name"] or "—", r["estado"], r["via"],
            ", ".join(r["anomalias"]) or "—",
            r["dni_ocr"] or r["dni_folder"] or "—", r["apellido"] or "—", r["nombre"] or "—",
            r["sexo"] or "—", r["nacimiento"] or "—", r["nacionalidad"] or "—",
            r["cuil"] or "—", r["vencimiento"] or "—",
            r.get("domicilio") or "—", (decl_addr or {}).get(cid, "") or "—",
            folder_cell, photo_cell(r["frente"], "Ver frente"), photo_cell(r["dorso"], "Ver dorso"),
            stamp,
        ])
    return rows


# --------------------------------------------------------------------------- #
# sheet writing (idempotent: create tab if missing, then clear + rewrite)
# --------------------------------------------------------------------------- #
def write_sheet(session, ssid, sheet_name, rows, cred_path, timeout=60):
    gids = sheet_gids(session, ssid, cred_path, timeout)
    if sheet_name not in gids:
        _check(session.post((SHEETS_API % ssid) + ":batchUpdate",
                            json={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]},
                            timeout=timeout), "create sheet", cred_path)
        gids = sheet_gids(session, ssid, cred_path, timeout)
    gid = gids[sheet_name]

    _check(session.post((SHEETS_API % ssid) + "/values/" + ("'%s'" % sheet_name) + ":clear",
                        json={}, timeout=timeout), "clear sheet", cred_path)
    _check(session.put((SHEETS_API % ssid) + "/values/" + ("'%s'!A1" % sheet_name),
                       params={"valueInputOption": "USER_ENTERED"},
                       json={"values": [HEADERS] + rows}, timeout=timeout),
           "write values", cred_path)

    # Light formatting: bold blue header + frozen first row + sane widths.
    fmt = [
        {"repeatCell": {
            "range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {
                "backgroundColor": {"red": 0.11, "green": 0.27, "blue": 0.53},
                "textFormat": {"bold": True,
                               "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "wrapStrategy": "WRAP"}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,wrapStrategy)"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": gid, "gridProperties": {"frozenRowCount": 1}},
            "fields": "gridProperties.frozenRowCount"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": gid, "dimension": "COLUMNS",
                      "startIndex": 0, "endIndex": len(HEADERS)},
            "properties": {"pixelSize": 130}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {   # Cliente + Anomalías wider
            "range": {"sheetId": gid, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2},
            "properties": {"pixelSize": 210}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": gid, "dimension": "COLUMNS", "startIndex": 4, "endIndex": 5},
            "properties": {"pixelSize": 260}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {   # Domicilio (OCR) + Dirección declarada + Carpeta
            "range": {"sheetId": gid, "dimension": "COLUMNS", "startIndex": 13, "endIndex": 16},
            "properties": {"pixelSize": 240}, "fields": "pixelSize"}},
    ]
    _check(session.post((SHEETS_API % ssid) + ":batchUpdate", json={"requests": fmt},
                        timeout=timeout), "format sheet", cred_path)
    return gid


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def parse_args(argv):
    p = argparse.ArgumentParser(
        description="OCR the clients' DNI photos and write them to a new tab of the "
                    "live spreadsheet, linked to each client.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("--fetch", action="store_true", help="download new photos from Drive first")
    p.add_argument("--incremental", action="store_true",
                   help="only OCR folders with new/changed photos (cache in out/_dni_ocr_cache.json)")
    p.add_argument("--dry-run", action="store_true", help="print the rows; write nothing")
    p.add_argument("--no-links", action="store_true",
                   help="skip the Drive link lookup (offline photo cache only)")
    p.add_argument("--csv", nargs="?", const=str(HERE / "out" / "_dni_datos.csv"),
                   help="also save a plain-text CSV copy (links become 'label (url)')")
    p.add_argument("--sheet-name", default=DEFAULT_SHEET)
    p.add_argument("--photos", default=str(DEFAULT_PHOTOS))
    p.add_argument("--xlsx", default=str(DEFAULT_XLSX), help="Impulso workbook (client matching)")
    p.add_argument("--spreadsheet-id", help="else IMPULSO_SHEET_ID env / .impulso_file_id")
    p.add_argument("--folder-id", help="DNI photos folder (else DNI_FOLDER_ID / default)")
    p.add_argument("--credentials", help="service-account JSON key (else service_account.json)")
    p.add_argument("--timeout", type=int, default=60)
    return p.parse_args(argv)


def run(fetch=False, incremental=False, no_links=False, dry_run=False, csv_path=None,
        sheet_name=DEFAULT_SHEET, photos=str(DEFAULT_PHOTOS), xlsx=str(DEFAULT_XLSX),
        spreadsheet_id=None, folder_id=None, credentials=None, timeout=60,
        progress=lambda m: None):
    """Full pipeline (fetch → OCR → links → sheet) → summary dict. Reusable entry
    point for the CLI, the dashboard menu and the hourly cron. With
    incremental=True only folders with NEW or changed photos are OCR'd (cache in
    out/_dni_ocr_cache.json); the sheet is always rewritten in full."""
    cred = fetch_impulso.resolve_credentials(credentials, HERE)
    folder_id = fetch_dni.resolve_folder_id(folder_id, HERE)

    if fetch:
        fetch_dni.download_all(photos, cred, folder_id, timeout, progress=progress)

    if not Path(xlsx).exists():
        raise SystemExit("Workbook not found: %s — run fetch_impulso.py first." % xlsx)
    clients = verify_dni.load_clients(xlsx)

    cache = None
    if incremental:
        cache = {}
        if CACHE_FILE.exists():
            try:
                cache = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
            except ValueError:
                cache = {}
    records = verify_dni.build_verifications(photos, clients, progress=progress,
                                             fields_cache=cache)
    if not records:
        raise SystemExit("No DNI photos under %s — run with --fetch first." % photos)
    if incremental:
        CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")

    session = photo_idx = client_rows = None
    clientes_gid = None
    ssid = ""
    if not (dry_run and no_links):
        ssid = fetch_impulso.resolve_file_id(spreadsheet_id, HERE)
        session = _session(cred)
    if not no_links and session is not None:
        progress("Resolving Drive photo links…")
        photo_idx = drive_photo_index(session, folder_id, timeout, progress=progress)
    if session is not None:
        try:
            gids = sheet_gids(session, ssid, cred, timeout)
            clientes_gid = gids.get(CLIENTES_SHEET)
            client_rows = clientes_live_rows(session, ssid, cred, timeout)
        except SystemExit as exc:   # reads are best-effort; only the write hard-fails
            progress("Client-row links unavailable — %s" % exc)

    stamp = datetime.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M")
    rows = build_rows(records, photo_idx or {}, client_rows or {}, clientes_gid, stamp,
                      decl_addr=declared_addresses(xlsx))

    if csv_path:
        path = Path(csv_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", newline="", encoding="utf-8-sig") as fh:  # BOM: Excel-safe accents
            w = csv.writer(fh)
            w.writerow(HEADERS)
            for row in rows:
                w.writerow([_plain(c) for c in row])
        progress("CSV saved: %s" % path)

    summary = {
        "rows": len(rows),
        "matched": sum(1 for r in records if r["cliente_id"]),
        "cached": sum(1 for r in records if r.get("ocr_cached")),
        "new": [{"persona": r["persona"], "cliente_id": r["cliente_id"],
                 "cliente_name": r["cliente_name"], "estado": r["estado"],
                 "dni": r["dni_ocr"] or r["dni_folder"],
                 "anomalias": r["anomalias"]}
                for r in records if not r.get("ocr_cached")],
        "estados": {k: sum(1 for r in records if r["estado"] == k)
                    for k in (verify_dni.OK, verify_dni.REVIEW, verify_dni.NO_MATCH)},
        "sheet": sheet_name, "wrote": False,
        "url": ("https://docs.google.com/spreadsheets/d/" + ssid) if ssid else "",
        "preview": rows[:8],
    }
    if not dry_run:
        write_sheet(session, ssid, sheet_name, rows, cred, timeout)
        summary["wrote"] = True
        progress("Wrote %d row(s) (%d matched to a client) to «%s» — %s"
                 % (len(rows), summary["matched"], sheet_name, summary["url"]))
    return summary


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    summary = run(fetch=args.fetch, incremental=args.incremental, no_links=args.no_links,
                  dry_run=args.dry_run, csv_path=args.csv, sheet_name=args.sheet_name,
                  photos=args.photos, xlsx=args.xlsx, spreadsheet_id=args.spreadsheet_id,
                  folder_id=args.folder_id, credentials=args.credentials,
                  timeout=args.timeout, progress=lambda m: print(m, file=sys.stderr))
    if args.incremental:
        logging.info("OCR: %d newly scanned, %d reused from cache.",
                     len(summary["new"]), summary["cached"])
    if args.dry_run:
        n_rev = summary["rows"] - summary["estados"].get(verify_dni.OK, 0)
        print("\nDRY RUN — %d row(s) (%d to review), nothing written. Sheet: «%s»"
              % (summary["rows"], n_rev, args.sheet_name))
        for row in summary["preview"]:
            print(" | ".join(str(c) for c in row[:8]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
