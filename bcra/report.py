#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
One entry point for the whole credit picture: BCRA (external bureau), Impulso
Crédito (first-party repayment), and the combined view — for a single person or
for everyone. It just orchestrates the three analyzers (analyze_bcra.py,
analyze_impulso.py, analyze_combined.py); all scoring logic lives there.

Usage:
    # Everyone: prints the BCRA, Impulso and combined rankings; saves the 3 .md
    python3 bcra/report.py --all

    # One person (CUIL, DNI, or ID Cliente): prints all three reports + saves
    # a single out/<id>.full.report.md dossier
    python3 bcra/report.py 20123456789
    python3 bcra/report.py 12345678            # by DNI
    python3 bcra/report.py C-0007              # by client id

Missing BCRA files are **downloaded automatically** (the polite, rate-limited
fetch_bcra path) — one CUIL for a person, every borrower's CUIL for --all. Use
--no-fetch to stay offline, --insecure for BCRA's TLS cert-chain quirk.

Options: --xlsx, --out, --impulso-weight, --no-fetch, --insecure, --no-save. The
Impulso workbook defaults to out/_impulso.xlsx (see fetch_impulso.py to download it
with a service account, or export it manually).
"""

import argparse
import datetime
import json
import logging
import re
import sys
from pathlib import Path

import analyze_bcra
import analyze_impulso
import analyze_combined
import fetch_bcra  # std-lib based; the polite downloader reused for auto-fetch

_RULE = "═" * 78

# Auto-fetch pacing (mirrors fetch_bcra's defaults: sequential + polite delay).
FETCH_DELAY, FETCH_JITTER, FETCH_RETRIES, FETCH_TIMEOUT = 2.0, 1.0, 4, 30


def _banner(title):
    return "\n".join(["", _RULE, "  " + title, _RULE])


def _ensure_bcra(cuils, out_dir, insecure):
    """Download a BCRA file for any 11-digit CUIL that isn't already on disk.
    Returns the number fetched. Missing files are the norm on a fresh checkout;
    fetch_one is resumable (skips existing) and writes a file even for a 404
    ('no record'), so this never re-hits the API for the same person twice."""
    out = Path(out_dir)
    missing, seen = [], set()
    for c in cuils:
        if c and len(c) == 11 and c not in seen:
            seen.add(c)
            if not (out / ("%s.json" % c)).exists():
                missing.append(c)
    if not missing:
        return 0
    logging.info("Downloading %d missing BCRA file(s) from the BCRA (polite, ~%.0fs each)…",
                 len(missing), FETCH_DELAY + FETCH_JITTER)
    opener = fetch_bcra.make_opener(insecure)
    pace = fetch_bcra.make_pace(FETCH_DELAY, FETCH_JITTER)
    want = list(fetch_bcra.ENDPOINTS.items())
    done = 0
    for i, c in enumerate(missing, 1):
        try:
            fetch_bcra.fetch_one(c, out, opener, want, FETCH_TIMEOUT, FETCH_RETRIES, pace)
            done += 1
            logging.info("  [%d/%d] %s ✓", i, len(missing), c)
        except Exception as e:  # one bad CUIL must not abort the whole run
            logging.warning("  [%d/%d] %s — fetch failed: %s", i, len(missing), c, e)
    return done


# --------------------------------------------------------------------------- #
# Everyone
# --------------------------------------------------------------------------- #
def report_all(out_dir, xlsx, today, iw, bw, no_save, do_fetch, insecure):
    out = Path(out_dir)
    icohort = analyze_impulso.build_cohort(xlsx, today)
    if do_fetch:  # pull BCRA data for every borrower's CUIL that's missing
        _ensure_bcra([c["cuil"] for c in icohort if c.get("cuil")], out_dir, insecure)
    bcohort = analyze_bcra.build_cohort(out_dir) if out.exists() else []
    combined = analyze_combined.build_combined(out_dir, xlsx, today, iw, bw)

    print(_banner("BCRA — EXTERNAL BUREAU RANKING"))
    print(analyze_bcra.render_leaderboard_cli(bcohort) if bcohort
          else "  (no BCRA files in %s — run fetch_bcra.py first)" % out_dir)
    print(_banner("IMPULSO CRÉDITO — FIRST-PARTY REPAYMENT RANKING"))
    print(analyze_impulso.render_leaderboard_cli(icohort) if icohort
          else "  (no borrowers found in the workbook)")
    print(_banner("COMBINED — BCRA + IMPULSO"))
    print(analyze_combined.render_cli(combined) if combined else "  (nothing to rank)")

    if not no_save:
        out.mkdir(parents=True, exist_ok=True)
        if bcohort:
            (out / "_ranking.md").write_text(analyze_bcra.render_leaderboard_md(bcohort) + "\n", encoding="utf-8")
        if icohort:
            (out / "_impulso_ranking.md").write_text(analyze_impulso.render_leaderboard_md(icohort) + "\n", encoding="utf-8")
        if combined:
            (out / "_combined_ranking.md").write_text(analyze_combined.render_md(combined) + "\n", encoding="utf-8")
        logging.info("\nSaved _ranking.md, _impulso_ranking.md, _combined_ranking.md in %s", out_dir)
    return 0


# --------------------------------------------------------------------------- #
# One person
# --------------------------------------------------------------------------- #
def _bcra_for(cuil, out_dir):
    """Return analyze_bcra's analysis dict for a CUIL, or None if no file on disk.
    (Downloading is handled up front by _ensure_bcra.)"""
    dest = Path(out_dir) / ("%s.json" % cuil)
    if dest.exists():
        try:
            return analyze_bcra.analyze_cuil(json.loads(dest.read_text(encoding="utf-8")))
        except (json.JSONDecodeError, OSError):
            return None
    return None


def report_person(token, out_dir, xlsx, today, iw, bw, do_fetch, insecure, no_save):
    tables = analyze_impulso.load_tables(xlsx)
    clients = analyze_impulso.build_clients(tables, today)
    cl = analyze_impulso.find_client(clients, token)
    digits = re.sub(r"\D", "", token)
    cuil = (cl["cuil"] if cl and cl["cuil"] else "") or (digits if len(digits) == 11 else "")

    if cuil and do_fetch:  # auto-download the BCRA file if it isn't on disk yet
        _ensure_bcra([cuil], out_dir, insecure)

    bcohort = analyze_bcra.build_cohort(out_dir) if Path(out_dir).exists() else []
    icohort = analyze_impulso.build_cohort(xlsx, today)

    # ---- BCRA ----
    a_bcra = _bcra_for(cuil, out_dir) if cuil else None
    if a_bcra:
        rank = analyze_bcra.rank_of(bcohort, a_bcra["cuil"] or cuil, a_bcra["score"], a_bcra["category"])
        bcra_block = analyze_bcra.render_cli(a_bcra, rank)
    elif cuil:
        bcra_block = "  No BCRA data for %s (download failed or --no-fetch). Retry: python3 bcra/analyze_bcra.py %s" % (cuil, cuil)
    else:
        bcra_block = "  No CUIL on file for this person — BCRA lookup needs an 11-digit CUIL."

    # ---- Impulso ----
    a_imp = analyze_impulso.analyze_client(cl, today) if cl else None
    if a_imp:
        key = a_imp["cuil"] or ("DNI:" + a_imp["dni"] if a_imp["dni"] else "CID:" + a_imp["client_id"])
        irank = analyze_impulso.rank_of(icohort, key, a_imp["score"], a_imp["category"]) if a_imp["has_record"] else None
        imp_block = analyze_impulso.render_client_cli(a_imp, irank)
    else:
        imp_block = "  No Impulso loan record matched '%s'." % token

    # ---- Combined (light records mirror what the --all ranking uses) ----
    bcra_light = None
    if a_bcra:
        bcra_light = {"cuil": a_bcra["cuil"] or cuil, "name": a_bcra["name"], "score": a_bcra["score"],
                      "rating": a_bcra["rating"], "category": a_bcra["category"], "decision": a_bcra["decision"]}
    imp_light = None
    if a_imp and a_imp["has_record"]:
        imp_light = {"cuil": a_imp["cuil"], "name": a_imp["name"], "score": a_imp["score"],
                     "rating": a_imp["rating"], "category": a_imp["category"], "decision": a_imp["decision"],
                     "n_loans": a_imp["n_loans"]}
    rec = analyze_combined.combine(bcra_light, imp_light, iw, bw) if (bcra_light or imp_light) else None
    combined_block = analyze_combined.render_one(rec) if rec else "  No data in either source for '%s'." % token

    name = (a_imp["name"] if a_imp and a_imp.get("name") else "") or (a_bcra["name"] if a_bcra else "") or token
    print(_banner("COMBINED VERDICT — %s" % name))
    print(combined_block)
    print(_banner("BCRA — EXTERNAL BUREAU REPORT"))
    print(bcra_block)
    print(_banner("IMPULSO CRÉDITO — FIRST-PARTY REPAYMENT REPORT"))
    print(imp_block)

    if not no_save and (a_bcra or a_imp):
        out = Path(out_dir)
        out.mkdir(parents=True, exist_ok=True)
        stem = cuil or (cl["dni"] if cl and cl["dni"] else (cl["client_id"] if cl else digits or "person"))
        parts = ["# Full Credit Dossier — %s" % name, "",
                 "_Generated %s._" % datetime.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M"), ""]
        if rec:
            parts += ["## Combined verdict", "", "```", analyze_combined.render_one(rec), "```"]
        if a_bcra:
            r = analyze_bcra.rank_of(bcohort, a_bcra["cuil"] or cuil, a_bcra["score"], a_bcra["category"])
            parts += ["", "---", "", analyze_bcra.render_md(a_bcra, r)]
        if a_imp:
            key = a_imp["cuil"] or ("DNI:" + a_imp["dni"] if a_imp["dni"] else "CID:" + a_imp["client_id"])
            r = analyze_impulso.rank_of(icohort, key, a_imp["score"], a_imp["category"]) if a_imp["has_record"] else None
            parts += ["", "---", "", analyze_impulso.render_client_md(a_imp, r)]
        path = out / ("%s.full.report.md" % stem)
        path.write_text("\n".join(parts) + "\n", encoding="utf-8")
        logging.info("\nSaved %s", path)
    return 0


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def parse_args(argv):
    here = Path(__file__).resolve().parent
    p = argparse.ArgumentParser(
        description="BCRA + Impulso + combined credit reports, for one person or everyone.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("who", nargs="?", help="CUIL, DNI, or ID Cliente (omit with --all)")
    p.add_argument("--all", action="store_true", help="all three rankings for the whole portfolio")
    p.add_argument("--xlsx", default=str(here / analyze_impulso.DEFAULT_XLSX), help="Impulso workbook (.xlsx)")
    p.add_argument("--out", default=str(here / "out"), help="BCRA json + report folder")
    p.add_argument("--impulso-weight", type=float, default=analyze_combined.IMPULSO_WEIGHT,
                   help="weight on the Impulso score in the combined blend (0-1)")
    p.add_argument("--no-fetch", action="store_true",
                   help="don't auto-download missing BCRA files (offline; use what's on disk)")
    p.add_argument("--insecure", action="store_true",
                   help="disable TLS verification when fetching (BCRA cert-chain fallback)")
    p.add_argument("--no-save", action="store_true", help="print only; don't write .md files")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    xlsx = Path(args.xlsx)
    today = datetime.date.today()
    iw = max(0.0, min(1.0, args.impulso_weight))
    bw = 1.0 - iw

    if not xlsx.exists():
        logging.error("Impulso workbook not found: %s\n"
                      "Download it (python3 bcra/fetch_impulso.py) or export: File ▸ Download ▸ .xlsx.", xlsx)
        return 2

    do_fetch = not args.no_fetch
    if args.all:
        return report_all(args.out, str(xlsx), today, iw, bw, args.no_save, do_fetch, args.insecure)
    if not args.who:
        logging.error("Pass a CUIL/DNI/ID Cliente, or --all. See --help.")
        return 2
    return report_person(args.who, args.out, str(xlsx), today, iw, bw, do_fetch, args.insecure, args.no_save)


if __name__ == "__main__":
    sys.exit(main())
