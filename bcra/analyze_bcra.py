#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Analyze BCRA Central de Deudores data for a CUIL and produce an English credit report.

Reads the raw JSON produced by fetch_bcra.py (bcra/out/<cuil>.json). For one CUIL it:
  - reconstructs the payment history from the 24-month series,
  - classifies the payment behavior into a named category
    (Consistent / On and off / Deteriorating / Recovering / Chronically delinquent / …),
  - computes industry-standard debt metrics (DPD buckets, % performing, volatility,
    trend, recency, exposure) and a 0-100 composite risk score + A-E rating,
  - ranks the person against the other debtors already downloaded in out/.

The report is printed to the command line AND saved to out/<cuil>.report.md.

If the person's JSON is not on disk yet, it is fetched first by reusing the polite,
rate-limited logic in fetch_bcra.py (fetch_one) — there is only one HTTP path.

The metrics mirror the Apps Script logic in LoanManagerV2.gs so scores/ratings line up
with the spreadsheet: bcraAggEntidades_ (per-period aggregation, montos in thousands
×1000), bcraScore_ (0.6·situación-base + 0.4·%performing ± trend, A-E with situación
floor), the consultarBcra_ decision rule, bcraSituLabel_, and analyzeBcraSeries_
(monthly totals, month-over-month net = payment/drawdown, per-creditor trend).

Usage:
    python3 bcra/analyze_bcra.py <cuil>     # one person → CLI + out/<cuil>.report.md
    python3 bcra/analyze_bcra.py --all      # leaderboard of all out/*.json → CLI + out/_ranking.md
    python3 bcra/analyze_bcra.py --help
"""

import argparse
import json
import logging
import re
import sys
from datetime import datetime
from pathlib import Path

import fetch_bcra  # same folder: reuse make_opener, make_pace, ENDPOINTS, fetch_one

# --------------------------------------------------------------------------- #
# Reference tables (English). Situación 1-6 ↔ standard delinquency / DPD buckets.
# --------------------------------------------------------------------------- #
STATUS_EN = {
    0: ("No classification", "—"),
    1: ("Current", "0-31 days past due"),
    2: ("Special follow-up (low risk)", "31-90 days past due"),
    3: ("Substandard / troubled (medium risk)", "90-180 days past due"),
    4: ("Doubtful (high insolvency risk)", "180-365 days past due"),
    5: ("Loss / uncollectible", "over 365 days past due"),
    6: ("Loss (technical ruling)", "uncollectible by technical ruling"),
}
BARS = {0: " ", 1: "▁", 2: "▃", 3: "▅", 4: "▆", 5: "▇", 6: "█"}
# Compact status names for narrow table columns (full names live in STATUS_EN).
STATUS_SHORT = {0: "—", 1: "Current", 2: "Special follow-up", 3: "Substandard",
                4: "Doubtful", 5: "Loss", 6: "Loss (technical)"}

# Behavior-categorization thresholds (tunable).
PERFORMING_MAX_SIT = 2     # situación ≤ 2 (and 0 = no debt) counts as "performing / on"
DELINQUENT_MIN_SIT = 3     # situación ≥ 3 counts as "delinquent / off"
CHRONIC_RATIO = 0.5        # ≥ this share of months delinquent → chronic

# History-aware scoring. LEVEL_SCORE maps a month's worst situación to a 0-100 standing
# (0 = no debt that month = clean). The score is a recency-weighted average of these,
# so recent delinquency weighs most and keeps hurting even after the debt is paid off.
LEVEL_SCORE = {0: 100, 1: 100, 2: 75, 3: 50, 4: 25, 5: 5, 6: 0}
RECENCY_DECAY = 0.88       # per-month weight decay (newest = 1); ~6-month half-life
RECENT_WINDOW = 6          # months treated as "recent" for the rating floor / review trigger


# --------------------------------------------------------------------------- #
# JSON helpers
# --------------------------------------------------------------------------- #
def _results(section):
    """results{} of a saved endpoint section (handles null / missing)."""
    if isinstance(section, dict):
        r = section.get("results")
        return r if isinstance(r, dict) else {}
    return {}


def denom_of(combined):
    for k in ("deudas", "historicas", "chequesRechazados"):
        r = _results(combined.get(k))
        if r.get("denominacion"):
            return str(r["denominacion"]).strip()
    return ""


def agg_entities(entidades, already_pesos):
    """Mirror of bcraAggEntidades_: {total, perf (sit≤2), worst, deudas[]}. Montos ARS."""
    mult = 1 if already_pesos else 1000
    total = perf = 0.0
    worst = 0
    deudas = []
    for e in entidades or []:
        m = float(e.get("monto") or 0) * mult
        s = int(e.get("situacion") or 0)
        total += m
        if s <= 2:
            perf += m
        if s > worst:
            worst = s
        deudas.append({"entidad": str(e.get("entidad") or "").strip(), "situacion": s, "monto": m})
    deudas.sort(key=lambda d: -d["situacion"])
    return {"total": total, "perf": perf, "worst": worst, "deudas": deudas}


def latest_snapshot(combined):
    """Current debt: /Deudas latest period, else last historical period. (entidades, periodo)."""
    d = _results(combined.get("deudas"))
    per = d.get("periodos") or []
    if per:
        return per[0].get("entidades") or [], str(per[0].get("periodo") or "")
    h = _results(combined.get("historicas"))
    hper = sorted(h.get("periodos") or [], key=lambda p: str(p.get("periodo")))
    if hper:
        return hper[-1].get("entidades") or [], str(hper[-1].get("periodo") or "")
    return [], ""


def history_series(combined):
    """Date-ordered monthly series from /Deudas/Historicas. Each item: periodo, worst,
    total (ARS), perf (ARS), entidades (raw, thousands)."""
    h = _results(combined.get("historicas"))
    per = sorted(h.get("periodos") or [], key=lambda p: str(p.get("periodo")))
    out = []
    for p in per:
        agg = agg_entities(p.get("entidades") or [], already_pesos=False)
        out.append({
            "periodo": str(p.get("periodo") or ""),
            "worst": agg["worst"], "total": agg["total"], "perf": agg["perf"],
            "entidades": p.get("entidades") or [],
        })
    return out


def cheques_count(combined):
    r = _results(combined.get("chequesRechazados"))
    n = 0
    for c in (r.get("causales") or []):
        for e in (c.get("entidades") or []):
            n += len(e.get("detalle") or [])
    return n


# --------------------------------------------------------------------------- #
# Scoring (reimplementation of bcraScore_) and decision (consultarBcra_)
# --------------------------------------------------------------------------- #
def _rate(score, floor_worst):
    """Map a 0-100 score to an A-E rating, floored by a worst-situación level so the
    grade never contradicts a recent BCRA classification."""
    letters = ["A", "B", "C", "D", "E"]
    if score >= 80:
        rating = "A - Low Risk"
    elif score >= 60:
        rating = "B - Medium-Low Risk"
    elif score >= 40:
        rating = "C - Medium Risk"
    elif score >= 20:
        rating = "D - High Risk"
    else:
        rating = "E - Very High Risk"
    if floor_worst >= 5 and letters.index(rating[0]) < 4:
        rating = "E - Very High Risk"
    elif floor_worst == 4 and letters.index(rating[0]) < 3:
        rating = "D - High Risk"
    return rating


def risk_score(worst_series, current_worst, recent_worst, pct_months, pct_performante, has_record):
    """History-aware 0-100 risk score (higher = lower risk) + A-E rating.

    Extends the spreadsheet's bcraScore_ (current standing only) by weighting the
    24-month history with recency: the score is 70% a recency-weighted average of
    each month's standing (LEVEL_SCORE) + 30% the share of months performing. Recent
    delinquency keeps the score down even after payoff. The A-E floor uses the worst
    situación of the last RECENT_WINDOW months. Scores here intentionally differ from
    the sheet's current-only quick check.
    """
    if not has_record:
        return 85, "No history"
    n = len(worst_series)
    if n == 0:
        # No monthly history: fall back to current standing (sheet-style).
        base = LEVEL_SCORE.get(current_worst, 0 if current_worst > 6 else 100)
        score = round(0.6 * base + 0.4 * (pct_performante if pct_performante is not None else 100))
        floor = current_worst
    else:
        num = den = 0.0
        for i, w in enumerate(worst_series):
            g = RECENCY_DECAY ** (n - 1 - i)        # newest month weight 1, older decays
            num += g * LEVEL_SCORE.get(w, 0)
            den += g
        history_comp = (num / den) if den else 100  # recency-weighted monthly standing
        perf = pct_months if pct_months is not None else 100
        score = round(0.7 * history_comp + 0.3 * perf)
        floor = recent_worst
    score = max(0, min(100, score))
    return score, _rate(score, floor)


def decision_of(current_worst, recent_worst, share, delta6m, has_record):
    if not has_record:
        return "NO RECORD"
    if current_worst >= 4 and share < 0.5:
        return "REJECT"
    # Recent serious delinquency warrants review even if the debt is currently paid off.
    if delta6m >= 2 or current_worst >= 4 or recent_worst >= 4:
        return "REVIEW"
    return "OK"


def trend_label(delta6m):
    if delta6m >= 2:
        return "Deteriorating (worsened ≥2 levels in 6 mo)"
    if delta6m >= 1:
        return "Slightly deteriorating"
    if delta6m <= -2:
        return "Improving (recovered ≥2 levels in 6 mo)"
    if delta6m <= -1:
        return "Slightly improving"
    return "Stable"


# --------------------------------------------------------------------------- #
# Payment-behavior categorization (date-based)
# --------------------------------------------------------------------------- #
def categorize(worst_series, latest_worst, delta6m):
    """Classify payment behavior from the date-ordered worst-situación series.
    Returns (category, explanation)."""
    n = len(worst_series)
    if n == 0:
        return "No record", "No debt history on file in the BCRA."

    delinq = [w >= DELINQUENT_MIN_SIT for w in worst_series]
    nonperf_ratio = sum(delinq) / n
    # Delinquency episodes = maximal runs of delinquent months.
    episodes, prev = 0, False
    for d in delinq:
        if d and not prev:
            episodes += 1
        prev = d
    latest_perf = latest_worst <= PERFORMING_MAX_SIT  # 0 (no debt) counts as performing

    if all(w == 1 for w in worst_series):
        return "Consistent — always current", "Situación 1 in all %d reported months." % n
    if episodes == 0:
        return ("Consistent — performing",
                "Always performing (≤ situación 2) across %d months; never delinquent." % n)
    if episodes >= 2:
        return ("On and off (volatile)",
                "%d separate delinquency episodes (crossed in and out of default)." % episodes)
    # episodes == 1 (a single delinquency spell)
    if latest_perf:
        return ("Recovering",
                "One past delinquency spell; currently performing again (6-mo trend %+d)." % (-delta6m))
    if nonperf_ratio >= CHRONIC_RATIO:
        return ("Chronically delinquent",
                "Delinquent in %d%% of the last %d months, still unresolved." % (round(100 * nonperf_ratio), n))
    return ("Deteriorating",
            "Recently entered delinquency (6-mo trend %+d), not yet recovered." % delta6m)


# --------------------------------------------------------------------------- #
# Per-creditor trend (mirror of analyzeBcraSeries_ porEntidad)
# --------------------------------------------------------------------------- #
def per_creditor(series, snapshot_deudas):
    n = len(series)
    idx6 = n - 7 if n >= 7 else 0
    idx12 = n - 13 if n >= 13 else 0

    def monto_of(ents, name):
        for e in ents:
            if str(e.get("entidad") or "").strip() == name:
                return float(e.get("monto") or 0) * 1000
        return 0.0

    def sit_of(ents, name):
        for e in ents:
            if str(e.get("entidad") or "").strip() == name:
                return int(e.get("situacion") or 0)
        return 0

    snap_by = {d["entidad"]: d for d in snapshot_deudas}
    names, seen = [], set()
    for d in snapshot_deudas:
        if d["entidad"] not in seen:
            seen.add(d["entidad"]); names.append(d["entidad"])
    for s in series:
        for e in s["entidades"]:
            nm = str(e.get("entidad") or "").strip()
            if nm and nm not in seen:
                seen.add(nm); names.append(nm)

    rows = []
    for name in names:
        snap = snap_by.get(name)
        actual = snap["monto"] if snap else 0.0
        situ = snap["situacion"] if snap else 0
        if not snap:
            for s in reversed(series):
                v = sit_of(s["entidades"], name)
                if v:
                    situ = v; break
        primero = None
        for s in series:
            mm = monto_of(s["entidades"], name)
            if mm > 0:
                primero = mm; break
        if primero is None:
            primero = actual
        rows.append({
            "entidad": name, "situacion": situ, "actual": actual,
            "hace6m": monto_of(series[idx6]["entidades"], name) if series else 0.0,
            "hace12m": monto_of(series[idx12]["entidades"], name) if series else 0.0,
            "neto": actual - primero,
        })
    rows.sort(key=lambda r: (-r["situacion"], -r["actual"]))
    return rows


# --------------------------------------------------------------------------- #
# Full analysis of one CUIL
# --------------------------------------------------------------------------- #
def analyze_cuil(combined):
    cuil = str(combined.get("cuil") or "")
    name = denom_of(combined)
    series = history_series(combined)
    snap_ents, snap_periodo = latest_snapshot(combined)
    snap = agg_entities(snap_ents, already_pesos=False)

    worst_series = [s["worst"] for s in series]
    n = len(worst_series)
    current_worst = snap["worst"]  # = peorSituación (bcraScore_/decision use current snapshot)
    worst_24m = max(worst_series) if worst_series else current_worst
    latest_hist = worst_series[-1] if n else current_worst
    six_ago = worst_series[-7] if n >= 7 else latest_hist
    delta6m = latest_hist - six_ago

    share = (snap["perf"] / snap["total"]) if snap["total"] else 1.0
    pct_performante = round(share * 1000) / 10  # one decimal, like consultarBcra_
    months_perf = sum(1 for w in worst_series if w <= PERFORMING_MAX_SIT)  # 0 = no debt = clean
    pct_months = round(100 * months_perf / n) if n else None

    # Status volatility: boundary crossings in the performing/non-performing series.
    perf_flags = [w <= PERFORMING_MAX_SIT for w in worst_series]
    transitions = sum(1 for i in range(1, n) if perf_flags[i] != perf_flags[i - 1])

    # Recency: months since the most recent delinquent (≥3) month.
    last_delinq = None
    for i, w in enumerate(worst_series):
        if w >= DELINQUENT_MIN_SIT:
            last_delinq = i
    recency = (n - 1 - last_delinq) if last_delinq is not None else None
    recent_worst = max(worst_series[-RECENT_WINDOW:]) if worst_series else current_worst

    has_record = bool(series) or bool(snap["deudas"])
    score, rating = risk_score(worst_series, current_worst, recent_worst,
                               pct_months, pct_performante, has_record)
    decision = decision_of(current_worst, recent_worst, share, delta6m, has_record)
    category, category_why = categorize(worst_series, latest_hist, delta6m)

    # Payment history (month-over-month net; delta<0 = payment/amortization).
    neto = []
    for i, s in enumerate(series):
        delta = 0.0 if i == 0 else s["total"] - series[i - 1]["total"]
        neto.append({"periodo": s["periodo"], "worst": s["worst"], "total": s["total"],
                     "delta": delta, "has_prev": i > 0})
    payments = [x for x in neto if x["has_prev"] and x["delta"] < 0]
    total_repaid = -sum(x["delta"] for x in payments)
    largest_payment = -min((x["delta"] for x in payments), default=0.0)

    return {
        "cuil": cuil, "name": name,
        "generated": datetime.now().astimezone().strftime("%Y-%m-%d %H:%M"),
        "has_record": has_record,
        "snap_periodo": snap_periodo,
        "current_worst": current_worst, "worst_24m": worst_24m,
        "pct_performante": pct_performante, "pct_months": pct_months,
        "months": n, "months_perf": months_perf,
        "transitions": transitions, "delta6m": delta6m, "recency": recency,
        "entidades": len(snap["deudas"]), "total_debt": snap["total"],
        "cheques": cheques_count(combined),
        "score": score, "rating": rating, "decision": decision,
        "trend_label": trend_label(delta6m),
        "category": category, "category_why": category_why,
        "worst_series": worst_series,
        "neto": neto, "payments_count": len(payments),
        "total_repaid": total_repaid, "largest_payment": largest_payment,
        "creditors": per_creditor(series, snap["deudas"]),
    }


# --------------------------------------------------------------------------- #
# Formatting helpers
# --------------------------------------------------------------------------- #
def money(x):
    return "$ " + format(int(round(x or 0)), ",d").replace(",", ".")


def signed_money(x):
    x = x or 0
    return ("-" if x < 0 else "") + "$ " + format(int(round(abs(x))), ",d").replace(",", ".")


def fmt_period(p):
    p = str(p)
    return p[:4] + "-" + p[4:6] if len(p) >= 6 else p


def pct_or_na(v):
    return "%d%%" % v if v is not None else "n/a"


def recency_txt(r):
    if r is None:
        return "no delinquency on record"
    return "current month" if r == 0 else "%d month(s) ago" % r


# --------------------------------------------------------------------------- #
# Cohort ranking
# --------------------------------------------------------------------------- #
def build_cohort(out_dir):
    """Score every out/<cuil>.json already on disk. Returns list of light records."""
    cohort = []
    for f in sorted(Path(out_dir).glob("*.json")):
        if f.name.startswith("_"):
            continue
        try:
            c = json.loads(f.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            continue
        a = analyze_cuil(c)
        cohort.append({"cuil": a["cuil"] or f.stem, "name": a["name"],
                       "score": a["score"], "rating": a["rating"],
                       "category": a["category"], "decision": a["decision"],
                       "has_record": a["has_record"]})
    return cohort


def rank_of(cohort, target_cuil, target_score, target_category):
    """Rank the target among the cohort by risk score (higher score = lower risk =
    rank 1). Tie-aware: `better`/`worse`/`equal` count peers by score."""
    m = len(cohort)
    if m == 0:
        return None
    peers = [c for c in cohort if c["cuil"] != target_cuil]
    npeers = len(peers)
    better = sum(1 for c in peers if c["score"] > target_score)   # lower risk than target
    worse = sum(1 for c in peers if c["score"] < target_score)    # higher risk than target
    equal = npeers - better - worse
    rank = better + 1                                             # competition rank; ties share it
    avg = sum(c["score"] for c in cohort) / m
    same_cat = sum(1 for c in peers if c["category"] == target_category)
    return {"rank": rank, "of": m, "npeers": npeers, "better": better, "worse": worse,
            "equal": equal, "avg": avg, "same_cat": same_cat}


# --------------------------------------------------------------------------- #
# Renderers
# --------------------------------------------------------------------------- #
def render_cli(a, rank):
    W = 66
    L = ["=" * W,
         " BCRA CREDIT REPORT — %s" % (a["name"] or "(name not reported)"),
         " CUIL %s   ·   generated %s" % (a["cuil"], a["generated"]),
         "=" * W, "",
         "VERDICT",
         "  Risk score       : %d / 100   (%s)  [history-weighted]" % (a["score"], a["rating"]),
         "  Decision          : %s" % a["decision"],
         "  Payment profile   : %s" % a["category"],
         "  Basis             : %s" % a["category_why"],
         "  6-month trend     : %s" % a["trend_label"]]

    if not a["has_record"]:
        L += ["", "  No debt found in the BCRA Central de Deudores (no record)."]
        L += _render_rank_lines(a, rank)
        L += ["", _footer_note(), "=" * W]
        return "\n".join(L)

    cw = STATUS_EN.get(a["current_worst"], ("?", "?"))
    w24 = STATUS_EN.get(a["worst_24m"], ("?", "?"))
    L += ["", "INDUSTRY-STANDARD METRICS",
          "  Current status    : sit %d — %s (%s)" % (a["current_worst"], cw[0], cw[1]),
          "  Worst (24 mo)     : sit %d — %s (%s)" % (a["worst_24m"], w24[0], w24[1]),
          "  Balance performing: %.1f%% of current debt (amount-weighted)" % a["pct_performante"],
          "  Months performing : %d / %d (%s)" % (a["months_perf"], a["months"], pct_or_na(a["pct_months"])),
          "  Status volatility : %d boundary transition(s)" % a["transitions"],
          "  Last delinquency  : %s" % recency_txt(a["recency"]),
          "  Active creditors  : %d" % a["entidades"],
          "  Total reported debt: %s" % money(a["total_debt"]),
          "  Rejected checks   : %d%s" % (a["cheques"], "  ⚠" if a["cheques"] else "")]

    # Payment history
    L += ["", "PAYMENT HISTORY  (inferred from month-over-month balance changes)",
          "  %-9s %-4s %-18s %15s   %s" % ("Period", "Sit", "Status", "Total debt", "Δ vs prior")]
    for x in a["neto"]:
        if x["has_prev"] and x["delta"] != 0:
            d = "%s (%s)" % (signed_money(x["delta"]), "payment" if x["delta"] < 0 else "increase")
        else:
            d = "—"
        L.append("  %-9s %-4d %-18s %15s   %s" % (
            fmt_period(x["periodo"]), x["worst"], STATUS_SHORT.get(x["worst"], "?"),
            money(x["total"]), d))
    L += ["",
          "  Inferred payments : %d month(s) with reductions · total ≈ %s · largest ≈ %s"
          % (a["payments_count"], money(a["total_repaid"]), money(a["largest_payment"]))]
    if a["worst_series"]:
        L += ["  Status timeline   : %s" % "".join(BARS.get(w, " ") for w in a["worst_series"]),
              "  (old→new, levels) : %s" % " ".join(str(w) for w in a["worst_series"])]

    # Creditors
    if a["creditors"]:
        L += ["", "CREDITORS",
              "  %-26s %-5s %14s %14s %14s %14s" % ("Entity", "Sit", "Now", "6 mo ago", "12 mo ago", "Net")]
        for c in a["creditors"]:
            L.append("  %-26s %-4d %14s %14s %14s %14s" % (
                (c["entidad"] or "(unnamed)")[:26], c["situacion"],
                money(c["actual"]), money(c["hace6m"]), money(c["hace12m"]),
                signed_money(c["neto"])))

    L += _render_rank_lines(a, rank)
    L += ["", _footer_note(), "=" * W]
    return "\n".join(L)


def _render_rank_lines(a, rank):
    if not rank or rank["of"] <= 1:
        return ["", "COHORT RANKING", "  Only this debtor downloaded — nothing to compare yet",
                "  (run fetch_bcra.py or analyze --all to populate the cohort)."]
    pos = "above average" if a["score"] >= rank["avg"] else "below average"
    tie = "   (%d tied at this score)" % rank["equal"] if rank["equal"] else ""
    return ["", "COHORT RANKING  (%d debtors downloaded; rank 1 = lowest risk)" % rank["of"],
            "  Risk rank         : %d of %d" % (rank["rank"], rank["of"]),
            "  Score             : %d   ·   cohort average %.0f (%s)" % (a["score"], rank["avg"], pos),
            "  Safer than        : %d of %d peers%s" % (rank["worse"], rank["npeers"], tie),
            "  Same profile      : %d of %d peers share '%s'" % (rank["same_cat"], rank["npeers"], a["category"])]


def _footer_note():
    return ("NOTES\n"
            "  • BCRA situación is a monthly classification snapshot, not a transaction\n"
            "    ledger; \"payments\" are inferred from reported-balance reductions.\n"
            "  • Amounts in ARS (BCRA reports in thousands; ×1000 applied).\n"
            "  • Score is history-weighted (recency-decayed 24-month standing), so it\n"
            "    intentionally differs from the spreadsheet's current-only quick check.")


def render_md(a, rank):
    L = ["# BCRA Credit Report — %s" % (a["name"] or "(name not reported)"),
         "",
         "**CUIL:** %s  ·  **Generated:** %s" % (a["cuil"], a["generated"]),
         "",
         "## Verdict",
         "",
         "| | |",
         "|---|---|",
         "| Risk score (history-weighted) | **%d / 100** (%s) |" % (a["score"], a["rating"]),
         "| Decision | **%s** |" % a["decision"],
         "| Payment profile | **%s** |" % a["category"],
         "| Basis | %s |" % a["category_why"],
         "| 6-month trend | %s |" % a["trend_label"]]

    if not a["has_record"]:
        L += ["", "> No debt found in the BCRA Central de Deudores (no record)."]
        L += _md_rank(a, rank)
        L += ["", _md_footer()]
        return "\n".join(L)

    cw = STATUS_EN.get(a["current_worst"], ("?", "?"))
    w24 = STATUS_EN.get(a["worst_24m"], ("?", "?"))
    L += ["", "## Industry-standard metrics", "",
          "| Metric | Value |", "|---|---|",
          "| Current status | sit %d — %s (%s) |" % (a["current_worst"], cw[0], cw[1]),
          "| Worst status (24 mo) | sit %d — %s (%s) |" % (a["worst_24m"], w24[0], w24[1]),
          "| Balance performing (amount-weighted) | %.1f%% |" % a["pct_performante"],
          "| Months performing | %d / %d (%s) |" % (a["months_perf"], a["months"], pct_or_na(a["pct_months"])),
          "| Status volatility | %d transition(s) |" % a["transitions"],
          "| Last delinquency | %s |" % recency_txt(a["recency"]),
          "| Active creditors | %d |" % a["entidades"],
          "| Total reported debt | %s |" % money(a["total_debt"]),
          "| Rejected checks | %d%s |" % (a["cheques"], " ⚠" if a["cheques"] else "")]

    L += ["", "## Payment history",
          "", "_Inferred from month-over-month balance changes (Δ < 0 = payment)._", "",
          "| Period | Sit | Status | Total debt | Δ vs prior |",
          "|---|---|---|---:|---:|"]
    for x in a["neto"]:
        if x["has_prev"] and x["delta"] != 0:
            d = "%s (%s)" % (signed_money(x["delta"]), "payment" if x["delta"] < 0 else "increase")
        else:
            d = "—"
        L.append("| %s | %d | %s | %s | %s |" % (
            fmt_period(x["periodo"]), x["worst"], STATUS_SHORT.get(x["worst"], "?"), money(x["total"]), d))
    L += ["",
          "**Inferred payments:** %d month(s) with reductions · total ≈ %s · largest ≈ %s"
          % (a["payments_count"], money(a["total_repaid"]), money(a["largest_payment"]))]
    if a["worst_series"]:
        L += ["", "**Status timeline (old→new):** `%s`" % " ".join(str(w) for w in a["worst_series"])]

    if a["creditors"]:
        L += ["", "## Creditors", "",
              "| Entity | Sit | Now | 6 mo ago | 12 mo ago | Net |",
              "|---|---|---:|---:|---:|---:|"]
        for c in a["creditors"]:
            L.append("| %s | %d | %s | %s | %s | %s |" % (
                c["entidad"] or "(unnamed)", c["situacion"],
                money(c["actual"]), money(c["hace6m"]), money(c["hace12m"]), signed_money(c["neto"])))

    L += _md_rank(a, rank)
    L += ["", _md_footer()]
    return "\n".join(L)


def _md_rank(a, rank):
    if not rank or rank["of"] <= 1:
        return ["", "## Cohort ranking", "", "_Only this debtor downloaded — nothing to compare yet._"]
    pos = "above average" if a["score"] >= rank["avg"] else "below average"
    tie = "  (%d tied at this score)" % rank["equal"] if rank["equal"] else ""
    return ["", "## Cohort ranking", "",
            "Compared against **%d** debtors downloaded (rank 1 = lowest risk):" % rank["of"], "",
            "- **Risk rank %d of %d**" % (rank["rank"], rank["of"]),
            "- Score %d vs cohort average %.0f (%s)" % (a["score"], rank["avg"], pos),
            "- Safer than **%d of %d** peers%s" % (rank["worse"], rank["npeers"], tie),
            "- Same profile ('%s'): %d of %d peers" % (a["category"], rank["same_cat"], rank["npeers"])]


def _md_footer():
    return ("---\n\n"
            "> **Notes.** BCRA *situación* is a monthly classification snapshot, not a transaction "
            "ledger; \"payments\" are inferred from reported-balance reductions. Amounts in ARS "
            "(BCRA reports in thousands; ×1000 applied). The score is history-weighted "
            "(recency-decayed 24-month standing) and intentionally differs from the "
            "spreadsheet's current-only quick check in `LoanManagerV2.gs`.")


# --------------------------------------------------------------------------- #
# Leaderboard (--all)
# --------------------------------------------------------------------------- #
def render_leaderboard_cli(cohort):
    ranked = sorted(cohort, key=lambda c: (-c["score"], c["cuil"]))
    L = ["BCRA DEBTOR LEADERBOARD  (%d debtor%s, best → worst)" % (len(ranked), "" if len(ranked) == 1 else "s"),
         "",
         "  %-4s %-13s %6s %-22s %-26s %s" % ("#", "CUIL", "Score", "Rating", "Payment profile", "Name")]
    for i, c in enumerate(ranked, 1):
        L.append("  %-4d %-13s %6d %-22s %-26s %s" % (
            i, c["cuil"], c["score"], c["rating"], c["category"][:26], (c["name"] or "")[:30]))
    if ranked:
        avg = sum(c["score"] for c in ranked) / len(ranked)
        L += ["", "  Cohort average score: %.0f" % avg]
    return "\n".join(L)


def render_leaderboard_md(cohort):
    ranked = sorted(cohort, key=lambda c: (-c["score"], c["cuil"]))
    gen = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M")
    L = ["# BCRA Debtor Ranking", "", "_Generated %s · %d debtor(s), best → worst._" % (gen, len(ranked)),
         "", "| # | CUIL | Score | Rating | Payment profile | Decision | Name |",
         "|---|---|---:|---|---|---|---|"]
    for i, c in enumerate(ranked, 1):
        L.append("| %d | %s | %d | %s | %s | %s | %s |" % (
            i, c["cuil"], c["score"], c["rating"], c["category"], c["decision"], c["name"] or ""))
    if ranked:
        avg = sum(c["score"] for c in ranked) / len(ranked)
        L += ["", "**Cohort average score:** %.0f" % avg]
    return "\n".join(L)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def parse_args(argv):
    here = Path(__file__).resolve().parent
    p = argparse.ArgumentParser(
        description="Analyze BCRA data for a CUIL and produce an English credit report.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("cuil", nargs="?", help="CUIL to analyze (with or without dashes)")
    p.add_argument("--all", action="store_true", help="rank all downloaded debtors (leaderboard)")
    p.add_argument("--out", default=str(here / "out"), help="data/report folder")
    p.add_argument("--no-save", action="store_true", help="print only; don't write the .md report")
    p.add_argument("--no-fetch", action="store_true", help="don't auto-fetch a missing target (error instead)")
    # Pass-through fetch options (used only when auto-fetching a missing CUIL).
    p.add_argument("--delay", type=float, default=2.0)
    p.add_argument("--jitter", type=float, default=1.0)
    p.add_argument("--retries", type=int, default=4)
    p.add_argument("--timeout", type=int, default=30)
    p.add_argument("--insecure", action="store_true", help="disable TLS verification (BCRA cert-chain fallback)")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    out_dir = Path(args.out)

    if args.all:
        if not out_dir.exists():
            logging.error("Output folder not found: %s", out_dir)
            return 2
        cohort = build_cohort(out_dir)
        if not cohort:
            logging.error("No <cuil>.json files in %s — run fetch_bcra.py first.", out_dir)
            return 2
        print(render_leaderboard_cli(cohort))
        if not args.no_save:
            path = out_dir / "_ranking.md"
            path.write_text(render_leaderboard_md(cohort) + "\n", encoding="utf-8")
            logging.info("\nSaved %s", path)
        return 0

    if not args.cuil:
        logging.error("Pass a CUIL to analyze, or --all for the leaderboard. See --help.")
        return 2

    cuil = re.sub(r"\D", "", args.cuil)
    if len(cuil) != 11:
        logging.error("'%s' is not an 11-digit CUIL.", args.cuil)
        return 2

    dest = out_dir / ("%s.json" % cuil)
    if not dest.exists():
        if args.no_fetch:
            logging.error("No data for %s at %s (and --no-fetch is set).", cuil, dest)
            return 2
        logging.info("No local data for %s — fetching from BCRA…", cuil)
        if args.insecure:
            logging.warning("TLS verification disabled (--insecure).")
        opener = fetch_bcra.make_opener(args.insecure)
        pace = fetch_bcra.make_pace(args.delay, args.jitter)
        want = list(fetch_bcra.ENDPOINTS.items())
        fetch_bcra.fetch_one(cuil, out_dir, opener, want, args.timeout, args.retries, pace)

    try:
        combined = json.loads(dest.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        logging.error("Could not read %s: %s", dest, e)
        return 2

    a = analyze_cuil(combined)
    cohort = build_cohort(out_dir)  # includes the target; ranking excludes self for "peers"
    rank = rank_of(cohort, a["cuil"] or cuil, a["score"], a["category"])

    print(render_cli(a, rank))
    if not args.no_save:
        path = out_dir / ("%s.report.md" % cuil)
        path.write_text(render_md(a, rank) + "\n", encoding="utf-8")
        logging.info("\nSaved %s", path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
