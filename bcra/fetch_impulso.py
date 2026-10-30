#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Download the private Impulso Crédito spreadsheet as .xlsx using a Google
**service account** — no Claude connector, no browser, fully headless.

It exports the bound Sheet via the Drive API (files.export → xlsx) and writes
`out/_impulso.xlsx`, which analyze_impulso.py / analyze_combined.py then read.
This is the ONLY file in bcra/ that needs third-party libraries; the analysis
tools stay std-lib. Install once:

    pip install google-auth requests

One-time setup (≈5 min, done once):
  1. In Google Cloud Console (any project): enable the **Google Drive API**.
  2. Create a **service account**; under it, add a **JSON key** and download it.
     Save it as  bcra/service_account.json  (git-ignored).
  3. Open the JSON, copy the  "client_email"  (…@….iam.gserviceaccount.com) and
     **share the spreadsheet with that address as Viewer** (like sharing with a
     person). Without this the export returns 404 — the service account has no
     access until the sheet is shared with it.

The spreadsheet's file id is read from (in order): --file-id, the
IMPULSO_SHEET_ID env var, or a git-ignored `bcra/.impulso_file_id` file — so the
real id never lives in committed code.

Usage:
    python3 bcra/fetch_impulso.py                 # → out/_impulso.xlsx
    python3 bcra/fetch_impulso.py --file-id <ID> --credentials path/to/key.json
    python3 bcra/analyze_impulso.py --all         # then analyze as usual
"""

import argparse
import json
import logging
import os
import sys
from pathlib import Path

EXPORT_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
DEFAULT_OUT = "out/_impulso.xlsx"
DEFAULT_CRED = "service_account.json"
FILE_ID_FILE = ".impulso_file_id"


def resolve_file_id(arg, here):
    if arg:
        return arg.strip()
    env = os.environ.get("IMPULSO_SHEET_ID")
    if env:
        return env.strip()
    f = here / FILE_ID_FILE
    if f.exists():
        val = f.read_text(encoding="utf-8").strip()
        if val:
            return val
    raise SystemExit(
        "No spreadsheet id. Pass --file-id, set IMPULSO_SHEET_ID, or put it in %s" % f)


def resolve_credentials(arg, here):
    for cand in (arg, os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"), str(here / DEFAULT_CRED)):
        if cand and Path(cand).exists():
            return Path(cand)
    raise SystemExit(
        "Service-account key not found. Download the JSON key and save it as %s,\n"
        "or pass --credentials / set GOOGLE_APPLICATION_CREDENTIALS. See this file's header." %
        (here / DEFAULT_CRED))


def _client_email(cred_path):
    try:
        return json.loads(Path(cred_path).read_text(encoding="utf-8")).get("client_email", "")
    except (json.JSONDecodeError, OSError):
        return ""


def fetch(file_id, cred_path, out_path, timeout):
    try:
        from google.oauth2 import service_account
        from google.auth.transport.requests import AuthorizedSession
    except ImportError:
        raise SystemExit("Missing dependencies. Install with:\n    pip install google-auth requests")

    creds = service_account.Credentials.from_service_account_file(str(cred_path), scopes=SCOPES)
    session = AuthorizedSession(creds)
    url = "https://www.googleapis.com/drive/v3/files/%s/export" % file_id
    resp = session.get(url, params={"mimeType": EXPORT_MIME}, timeout=timeout)

    if resp.status_code == 404:
        email = _client_email(cred_path)
        raise SystemExit(
            "HTTP 404 — the service account can't see this sheet.\n"
            "Share the spreadsheet (Viewer) with the service-account address:\n    %s" %
            (email or "(see client_email in your JSON key)"))
    if resp.status_code == 403 and "exportSizeLimit" in resp.text:
        raise SystemExit("HTTP 403 — the export exceeds Drive's 10 MB limit for files.export.")
    if resp.status_code != 200:
        raise SystemExit("Drive export failed (HTTP %d): %s" % (resp.status_code, resp.text[:400]))

    data = resp.content
    if data[:2] != b"PK":  # .xlsx is a zip; anything else is an error page/JSON
        raise SystemExit("Unexpected response (not an .xlsx). First bytes: %r" % data[:80])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(data)
    return len(data)


def parse_args(argv):
    here = Path(__file__).resolve().parent
    p = argparse.ArgumentParser(
        description="Download the Impulso spreadsheet as .xlsx via a Google service account.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("--file-id", help="spreadsheet id (else IMPULSO_SHEET_ID env or .impulso_file_id)")
    p.add_argument("--credentials", help="service-account JSON key (else env / service_account.json)")
    p.add_argument("--out", default=str(here / DEFAULT_OUT), help="output .xlsx path")
    p.add_argument("--timeout", type=int, default=60)
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    here = Path(__file__).resolve().parent
    file_id = resolve_file_id(args.file_id, here)
    cred = resolve_credentials(args.credentials, here)
    out = Path(args.out)

    logging.info("Exporting sheet %s…", file_id)
    n = fetch(file_id, cred, out, args.timeout)
    logging.info("Saved %s (%.0f KB)", out, n / 1024)
    logging.info("Now run:  python3 %s --all", (here / "analyze_impulso.py"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
