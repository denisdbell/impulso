#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Download the clients' DNI photos from the shared Google Drive folder using the
same Google **service account** as fetch_impulso.py — no browser, fully headless.

The Drive folder holds one **subfolder per person**, whose name embeds the DNI,
e.g.  ``Facundo Acosta 42886617/``  containing the two photos:

    Facundo Acosta DNI 20261029-214950.jpeg        ← front  (anverso)
    Facundo Acosta DNI dorso 20261029-214950.jpeg  ← back   (dorso, has the MRZ)

Files are cached under ``out/dni/<subfolder>/<filename>`` (``out/`` is git-ignored,
so the PII never gets committed) and the download is resumable — already-present
files are skipped. ``verify_dni.py`` then OCRs them and compares to the Clientes
sheet.

This file, like fetch_impulso.py, needs the Google client libraries:

    pip install google-auth requests

One-time setup (same service account as fetch_impulso.py):
  1. Open bcra/service_account.json and copy the  "client_email".
  2. In Drive, **share the DNI folder (and its subfolders) with that address as
     Viewer** — otherwise the listing returns 404 (the SA can't see it).

The folder id is read from (in order): --folder-id, the DNI_FOLDER_ID env var, a
git-ignored bcra/.dni_folder_id file, or the default below (the same id already
present in VerificacionDNI.gs — it is not a secret).

Usage:
    python3 bcra/fetch_dni.py                      # → out/dni/<persona>/…
    python3 bcra/fetch_dni.py --folder-id <ID> --credentials path/to/key.json
"""

import argparse
import logging
import os
import sys
from pathlib import Path

import fetch_impulso  # reuse resolve_credentials / _client_email and the SA auth pattern

# Same folder id as the Apps Script (VerificacionDNI.gs) — not secret.
DEFAULT_FOLDER_ID = "10lrgWdEkQdykjocMpRR-DlPeRQjg_a43"
FOLDER_ID_FILE = ".dni_folder_id"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
FOLDER_MIME = "application/vnd.google-apps.folder"
# Only download files whose name mentions one of these (skip contracts, receipts…).
KEYWORDS = ("dni", "cuil")
DEFAULT_OUT = "out/dni"


def resolve_folder_id(arg, here):
    if arg:
        return arg.strip()
    env = os.environ.get("DNI_FOLDER_ID")
    if env:
        return env.strip()
    f = here / FOLDER_ID_FILE
    if f.exists():
        val = f.read_text(encoding="utf-8").strip()
        if val:
            return val
    return DEFAULT_FOLDER_ID


def _session(cred_path):
    try:
        from google.oauth2 import service_account
        from google.auth.transport.requests import AuthorizedSession
    except ImportError:
        raise SystemExit("Missing dependencies. Install with:\n    pip install google-auth requests")
    creds = service_account.Credentials.from_service_account_file(str(cred_path), scopes=SCOPES)
    return AuthorizedSession(creds)


def _list_children(session, parent_id, timeout, only_folders=False):
    """Yield {id,name,mimeType} for the direct children of a folder (paged)."""
    url = "https://www.googleapis.com/drive/v3/files"
    q = "'%s' in parents and trashed=false" % parent_id
    if only_folders:
        q += " and mimeType='%s'" % FOLDER_MIME
    page = None
    out = []
    while True:
        params = {
            "q": q,
            "fields": "nextPageToken, files(id,name,mimeType)",
            "pageSize": 1000,
            "supportsAllDrives": "true",
            "includeItemsFromAllDrives": "true",
        }
        if page:
            params["pageToken"] = page
        resp = session.get(url, params=params, timeout=timeout)
        if resp.status_code == 404:
            raise SystemExit(
                "HTTP 404 — the service account can't see folder %s.\n"
                "Share the DNI folder (Viewer) with the service-account address "
                "(the \"client_email\" in service_account.json)." % parent_id)
        if resp.status_code != 200:
            raise SystemExit("Drive list failed (HTTP %d): %s" % (resp.status_code, resp.text[:300]))
        data = resp.json()
        out.extend(data.get("files", []))
        page = data.get("nextPageToken")
        if not page:
            return out


def _wanted(name):
    low = (name or "").lower()
    return any(k in low for k in KEYWORDS)


def _download_file(session, file_id, dest, timeout):
    url = "https://www.googleapis.com/drive/v3/files/%s" % file_id
    resp = session.get(url, params={"alt": "media", "supportsAllDrives": "true"},
                       timeout=timeout, stream=True)
    if resp.status_code != 200:
        raise RuntimeError("download HTTP %d: %s" % (resp.status_code, resp.text[:200]))
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + ".part")
    with open(tmp, "wb") as fh:
        for chunk in resp.iter_content(chunk_size=65536):
            if chunk:
                fh.write(chunk)
    tmp.replace(dest)
    return dest.stat().st_size


def download_all(out_dir, cred_path, folder_id, timeout=60, progress=lambda m: None):
    """Download every DNI photo into out_dir/<subfolder>/<filename>. Resumable
    (skips files already on disk). Returns a stats dict."""
    out_dir = Path(out_dir)
    session = _session(cred_path)
    progress("Listing people folders…")
    subfolders = _list_children(session, folder_id, timeout, only_folders=True)
    stats = {"folders": len(subfolders), "downloaded": 0, "skipped": 0, "errors": 0}
    for i, sub in enumerate(sorted(subfolders, key=lambda f: (f.get("name") or "").lower()), 1):
        name = sub.get("name") or sub["id"]
        progress("(%d/%d) %s" % (i, len(subfolders), name))
        try:
            files = _list_children(session, sub["id"], timeout)
        except SystemExit as exc:
            logging.warning("skip %s: %s", name, exc)
            stats["errors"] += 1
            continue
        for f in files:
            if str(f.get("mimeType", "")).startswith("application/vnd.google-apps"):
                continue  # native Google Docs/Sheets, not a photo
            if not _wanted(f.get("name")):
                continue
            dest = out_dir / name / f["name"]
            if dest.exists() and dest.stat().st_size > 0:
                stats["skipped"] += 1
                continue
            try:
                _download_file(session, f["id"], dest, timeout)
                stats["downloaded"] += 1
            except Exception as exc:  # noqa: BLE001 — keep going, log per file
                logging.warning("download failed for %s/%s: %s", name, f.get("name"), exc)
                stats["errors"] += 1
    progress("DNI photos: %d downloaded, %d already on disk, %d errors (%d folders)."
             % (stats["downloaded"], stats["skipped"], stats["errors"], stats["folders"]))
    return stats


def parse_args(argv):
    here = Path(__file__).resolve().parent
    p = argparse.ArgumentParser(
        description="Download clients' DNI photos from Drive via a Google service account.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("--folder-id", help="Drive folder id (else DNI_FOLDER_ID env / .dni_folder_id / default)")
    p.add_argument("--credentials", help="service-account JSON key (else env / service_account.json)")
    p.add_argument("--out", default=str(here / DEFAULT_OUT), help="output folder for the photos")
    p.add_argument("--timeout", type=int, default=60)
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    here = Path(__file__).resolve().parent
    folder_id = resolve_folder_id(args.folder_id, here)
    cred = fetch_impulso.resolve_credentials(args.credentials, here)
    logging.info("Downloading DNI photos from folder %s…", folder_id)
    stats = download_all(args.out, cred, folder_id, args.timeout, progress=logging.info)
    logging.info("Now run:  python3 %s", (here / "verify_dni.py"))
    return 0 if stats["errors"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
