#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Combined credit ranking: BCRA (external bureau) + Impulso Crédito (first-party
repayment) joined per person by CUIL.

- BCRA side  : analyze_bcra.build_cohort(out/) — every out/<cuil>.json already fetched.
- Impulso side: analyze_impulso.build_cohort(out/_impulso.xlsx) — every borrower.
- Join key   : the 11-digit CUIL (Clientes!CUIL ↔ out/<cuil>.json filename).

A blended 0-100 score weights the first-party Impulso signal above the bureau
signal (direct repayment behaviour with us is the stronger predictor). The combined
decision is driven by that blended score, with two safety rules: a REJECT on either
source escalates a clean OK to REVIEW, and a unanimous REJECT stays REJECT — so a
good first-party payer is never auto-rejected by the bureau alone. People present in
only one source are carried through on that source alone and flagged "BCRA only" /
"Impulso only".

Usage:
    python3 bcra/analyze_combined.py                 # ranking → out/_combined_ranking.md
    python3 bcra/analyze_combined.py <cuil>          # one person's combined verdict
    python3 bcra/analyze_combined.py --help
"""

import argparse
import datetime
import logging
import re
import sys
from pathlib import Path

import analyze_bcra
import analyze_impulso

# First-party repayment behaviour is the stronger predictor of repayment to us,
# so it outweighs the external bureau score. Tunable via --impulso-weight.
IMPULSO_WEIGHT = 0.6
BCRA_WEIGHT = 0.4

# A decision counts as a real signal only if it isn't "no record".
_NEUTRAL_DECISIONS = {"NO RECORD", "No history", "—", None, ""}


def grade_letter(rating):
    """'B - Medium-Low Risk' → 'B'; '—'/empty → '—'."""
    return rating[0] if rating and rating[0] in "ABCDE" else "—"


def combined_decision(combined, source_decisions):
    """Score-driven (so the Impulso/BCRA weighting governs), with two safety rules:
    a REJECT on *either* source escalates a clean OK to REVIEW, and a *unanimous*
    REJECT stays REJECT. A good first-party payer is therefore never auto-rejected
    by the bureau alone, but neither side's hard reject is ever silently cleared."""
    present = [d for d in source_decisions if d not in _NEUTRAL_DECISIONS]
    if not present:
        return "NO RECORD"
    base = "OK" if combined >= 60 else "REVIEW" if combined >= 35 else "REJECT"
    if all(d == "REJECT" for d in present):
        return "REJECT"                                  # unanimous reject
    if base == "OK" and any(d == "REJECT" for d in present):
        return "REVIEW"                                  # one side rejects → at least review
    return base


def combine(bcra, imp, iw=IMPULSO_WEIGHT, bw=BCRA_WEIGHT):
    """Merge a BCRA cohort record and/or an Impulso cohort record (either may be None)."""
    bs = bcra["score"] if bcra else None
    is_ = imp["score"] if imp else None
    if bs is not None and is_ is not None:
        combined = round(bw * bs + iw * is_)
        source = "both"
    elif is_ is not None:
        combined, source = is_, "Impulso only"
    else:
        combined, source = bs, "BCRA only"

    decision = combined_decision(combined, [bcra["decision"] if bcra else None,
                                            imp["decision"] if imp else None])
    name = (imp["name"] if imp and imp.get("name") else "") or (bcra["name"] if bcra else "") or ""
    cuil = (imp["cuil"] if imp and imp.get("cuil") else "") or (bcra["cuil"] if bcra else "")
    return {
        "cuil": cuil, "name": name, "source": source,
        "combined": combined, "rating": analyze_impulso.rate_of(combined),
        "decision": decision,
        "bcra_score": bs, "bcra_rating": bcra["rating"] if bcra else None,
        "bcra_category": bcra["category"] if bcra else None,
        "bcra_decision": bcra["decision"] if bcra else None,
        "imp_score": is_, "imp_rating": imp["rating"] if imp else None,
        "imp_category": imp["category"] if imp else None,
        "imp_decision": imp["decision"] if imp else None,
        "imp_loans": imp["n_loans"] if imp else None,
    }


def build_combined(out_dir, xlsx, today=None, iw=IMPULSO_WEIGHT, bw=BCRA_WEIGHT):
    bcohort = analyze_bcra.build_cohort(out_dir)
    icohort = analyze_impulso.build_cohort(xlsx, today)
    bmap = {c["cuil"]: c for c in bcohort if c.get("cuil")}

    rows, used = [], set()
    for imp in icohort:                       # every Impulso borrower
        bcra = bmap.get(imp["cuil"]) if imp.get("cuil") else None
        if bcra:
            used.add(imp["cuil"])
        rows.append(combine(bcra, imp, iw, bw))
    for cuil, bcra in bmap.items():           # BCRA-only (prospects, never borrowed)
        if cuil not in used:
            rows.append(combine(bcra, None, iw, bw))
    return rows


# --------------------------------------------------------------------------- #
# Renderers
# --------------------------------------------------------------------------- #
def _src_tag(s):
    return {"both": "", "BCRA only": " (BCRA only)", "Impulso only": " (Impulso only)"}.get(s, "")


def render_cli(rows):
    ranked = sorted(rows, key=lambda r: (-(r["combined"] if r["combined"] is not None else -1), r["name"]))
    L = ["COMBINED CREDIT RANKING — BCRA + Impulso  (%d people, best → worst)" % len(ranked),
         "  weights: Impulso %.0f%% · BCRA %.0f%%   ·   decision from combined score (reject on either → ≥review)" % (IMPULSO_WEIGHT * 100, BCRA_WEIGHT * 100),
         "",
         "  %-4s %-13s %6s %4s %7s %7s %-9s %s" % ("#", "CUIL", "Comb.", "Rt", "BCRA", "Impulso", "Decision", "Name")]
    for i, r in enumerate(ranked, 1):
        bs = "%d" % r["bcra_score"] if r["bcra_score"] is not None else "—"
        is_ = "%d" % r["imp_score"] if r["imp_score"] is not None else "—"
        L.append("  %-4d %-13s %6s %4s %7s %7s %-9s %s%s" % (
            i, r["cuil"] or "—",
            ("%d" % r["combined"]) if r["combined"] is not None else "—",
            grade_letter(r["rating"]), bs, is_, r["decision"],
            (r["name"] or "")[:28], _src_tag(r["source"])))
    if ranked:
        vals = [r["combined"] for r in ranked if r["combined"] is not None]
        L += ["", "  Average combined score: %.0f   ·   joined both sources: %d   ·   single-source: %d"
              % (sum(vals) / len(vals), sum(1 for r in ranked if r["source"] == "both"),
                 sum(1 for r in ranked if r["source"] != "both"))]
    return "\n".join(L)


def render_md(rows):
    ranked = sorted(rows, key=lambda r: (-(r["combined"] if r["combined"] is not None else -1), r["name"]))
    gen = datetime.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M")
    both = sum(1 for r in ranked if r["source"] == "both")
    L = ["# Combined Credit Ranking — BCRA + Impulso Crédito", "",
         "_Generated %s · %d people, best → worst._" % (gen, len(ranked)), "",
         "Blended score weights **first-party Impulso repayment %.0f%%** and **BCRA bureau %.0f%%**. "
         "The **decision follows the combined score** (≥60 OK · ≥35 REVIEW · else REJECT), escalating a "
         "clean OK to REVIEW when either source rejects, and holding REJECT when both do — so a good "
         "first-party payer is never auto-rejected by the bureau alone. "
         "People in only one source are carried on that source alone and tagged accordingly "
         "(%d joined on both; %d single-source)."
         % (IMPULSO_WEIGHT * 100, BCRA_WEIGHT * 100, both, len(ranked) - both), "",
         "| # | CUIL | Combined | Rating | BCRA | Impulso | Decision | Source | Name |",
         "|---|---|---:|---|---:|---:|---|---|---|"]
    for i, r in enumerate(ranked, 1):
        bs = ("%d (%s)" % (r["bcra_score"], grade_letter(r["bcra_rating"]))) if r["bcra_score"] is not None else "—"
        is_ = ("%d (%s)" % (r["imp_score"], grade_letter(r["imp_rating"]))) if r["imp_score"] is not None else "—"
        L.append("| %d | %s | %s | %s | %s | %s | %s | %s | %s |" % (
            i, r["cuil"] or "—",
            ("**%d**" % r["combined"]) if r["combined"] is not None else "—",
            grade_letter(r["rating"]), bs, is_, r["decision"],
            {"both": "both", "BCRA only": "BCRA", "Impulso only": "Impulso"}.get(r["source"], r["source"]),
            r["name"] or ""))
    if ranked:
        vals = [r["combined"] for r in ranked if r["combined"] is not None]
        L += ["", "**Average combined score:** %.0f" % (sum(vals) / len(vals)),
              "", "> **Notes.** Combined = %.0f%%·Impulso + %.0f%%·BCRA (first-party repayment weighs "
              "more as the stronger predictor). The decision follows the combined score, escalated to "
              "REVIEW if either source rejects and held at REJECT only when both do. The BCRA score is "
              "history-weighted bureau standing; the Impulso score is first-party repayment behaviour. "
              "See `<cuil>.report.md` (BCRA) and `<cuil>.impulso.report.md` (Impulso) for the detail "
              "behind each person." % (IMPULSO_WEIGHT * 100, BCRA_WEIGHT * 100)]
    return "\n".join(L)


def render_one(r):
    L = ["=" * 66,
         " COMBINED CREDIT VERDICT — %s" % (r["name"] or "(no name)"),
         " CUIL %s   ·   source: %s" % (r["cuil"] or "—", r["source"]),
         "=" * 66, "",
         "  Combined score   : %s / 100   (%s)" % (r["combined"], r["rating"]),
         "  Decision         : %s   (from combined score; reject on either → ≥review)" % r["decision"],
         "  Weights          : Impulso %.0f%% · BCRA %.0f%%" % (IMPULSO_WEIGHT * 100, BCRA_WEIGHT * 100),
         "",
         "  BCRA (bureau)    : %s   ·   %s   ·   %s" % (
             "%d" % r["bcra_score"] if r["bcra_score"] is not None else "no record",
             r["bcra_rating"] or "—", r["bcra_category"] or "—"),
         "  Impulso (ours)   : %s   ·   %s   ·   %s%s" % (
             "%d" % r["imp_score"] if r["imp_score"] is not None else "no loans",
             r["imp_rating"] or "—", r["imp_category"] or "—",
             "  [%d loan(s)]" % r["imp_loans"] if r["imp_loans"] else ""),
         "",
         "  Detail: out/%s.report.md (BCRA) · out/%s.impulso.report.md (Impulso)" % (
             r["cuil"] or "<cuil>", r["cuil"] or "<cuil>"),
         "=" * 66]
    return "\n".join(L)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def parse_args(argv):
    here = Path(__file__).resolve().parent
    p = argparse.ArgumentParser(
        description="Combined BCRA + Impulso credit ranking, joined by CUIL.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("cuil", nargs="?", help="one CUIL for a combined verdict (default: full ranking)")
    p.add_argument("--xlsx", default=str(here / analyze_impulso.DEFAULT_XLSX), help="Impulso workbook (.xlsx)")
    p.add_argument("--out", default=str(here / "out"), help="BCRA json + report folder")
    p.add_argument("--impulso-weight", type=float, default=IMPULSO_WEIGHT,
                   help="weight on the Impulso score (0-1); BCRA gets the remainder")
    p.add_argument("--no-save", action="store_true", help="print only; don't write the .md")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    out_dir, xlsx = Path(args.out), Path(args.xlsx)
    iw = max(0.0, min(1.0, args.impulso_weight))
    bw = 1.0 - iw
    today = datetime.date.today()

    if not xlsx.exists():
        logging.error("Impulso workbook not found: %s (export: File ▸ Download ▸ .xlsx).", xlsx)
        return 2
    if not out_dir.exists():
        logging.error("BCRA folder not found: %s (run fetch_bcra.py first).", out_dir)
        return 2

    rows = build_combined(str(out_dir), str(xlsx), today, iw, bw)
    if not rows:
        logging.error("Nothing to rank — no BCRA json and no Impulso borrowers.")
        return 2

    if args.cuil:
        key = re.sub(r"\D", "", args.cuil)
        match = next((r for r in rows if r["cuil"] == key), None)
        if not match:
            logging.error("CUIL %s not found in either source.", key)
            return 2
        print(render_one(match))
        return 0

    print(render_cli(rows))
    if not args.no_save:
        out_dir.mkdir(parents=True, exist_ok=True)
        path = out_dir / "_combined_ranking.md"
        path.write_text(render_md(rows) + "\n", encoding="utf-8")
        logging.info("\nSaved %s", path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
