#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Analyze each person's **Impulso Crédito** loan history and produce a repayment
report + a standalone Impulso ranking — the first-party counterpart to the BCRA
report (analyze_bcra.py). Together they feed analyze_combined.py.

Input is the bound Impulso spreadsheet exported to .xlsx (File ▸ Download ▸
Microsoft Excel). The private sheet holds real borrower PII, so the download and
all reports live in the git-ignored `out/` folder; only this code is committed.
Download it with:  File ▸ Download ▸ Microsoft Excel (.xlsx) → save as
`bcra/out/_impulso.xlsx`  (the default path).

For one person it reconstructs every loan (on-time / late / overdue / in a plan /
frozen), scores 0-100 with the same A-E rating scale as the BCRA report, assigns
a payment profile, and ranks them against the rest of the portfolio. The
behavioural facts come from the normalized model:

  - Prestatarios : one row per loan (Capital, Vencimiento, Total a Pagar/Pagado,
                   Saldo, Estado, Plan desde, Congelado el, ID Cliente).
  - Clientes     : identity + the CUIL join key (used by analyze_combined.py).
  - Pagos        : dated payments → on-time vs late for closed loans.
  - Cuotas       : per-installment due dates + state → finer overdue detection.

The spreadsheet's `Estado` is a live formula whose cached export value can be
stale (e.g. a loan flagged VENCIDO whose due date is still in the future). So the
effective status is **derived from dates + balances** and reconciled with Estado;
a divergence is flagged in the per-loan table rather than trusted blindly.

Usage:
    python3 bcra/analyze_impulso.py --all                 # ranking → out/_impulso_ranking.md
    python3 bcra/analyze_impulso.py <cuil|dni|ID Cliente> # one person → out/<cuil>.impulso.report.md
    python3 bcra/analyze_impulso.py --help
"""

import argparse
import datetime
import logging
import re
import sys
from pathlib import Path

import xlsx_mini  # same folder: dependency-free .xlsx reader

DEFAULT_XLSX = "out/_impulso.xlsx"

# --------------------------------------------------------------------------- #
# Scoring model (parallels analyze_bcra.py so the two scores blend coherently)
# --------------------------------------------------------------------------- #
# Each loan gets a 0-100 "standing" from its outcome; the client score is a
# recency-weighted average of those + a performing-share term, with a rating
# floor driven by the current worst standing (overdue > frozen/plan).
RECENCY_DECAY = 0.85        # per-loan weight decay (newest loan = 1)
RECENT_LOANS = 3            # most recent loans that drive the REVIEW trigger
CHRONIC_RATIO = 0.5         # ≥ this share of loans late/overdue → chronic
PAID_EPS = 0.5             # ARS tolerance for "fully paid"

# Grace buckets (days) → standing, for a loan that WAS repaid but repaid late.
def _repaid_late_level(dl):
    if dl <= 3:  return 85
    if dl <= 7:  return 70
    if dl <= 14: return 55
    if dl <= 30: return 40
    return 25

# Standing for a loan still open and currently overdue, by days past due.
def _overdue_level(dl):
    if dl <= 7:  return 35
    if dl <= 14: return 25
    if dl <= 30: return 15
    if dl <= 60: return 8
    return 2

_LEVEL = {
    "repaid_ontime": 100,   # closed, last payment on/before due date
    "repaid_untimed": 92,   # closed, no dated payment to prove timing (benefit of the doubt)
    "current": 90,          # open, not yet due, in good standing
    "frozen": 45,           # CONGELADO — mora paused (hardship/negotiation)
    "in_plan": 40,          # EN PLAN — restructured after falling behind
}
# "performing" = clean standings that count toward the performing share.
_PERFORMING = {"repaid_ontime", "repaid_untimed", "current"}
# human labels for the per-loan table
_STATUS_LABEL = {
    "repaid_ontime": "Repaid on time", "repaid_untimed": "Repaid (timing n/a)",
    "repaid_late": "Repaid late", "current": "Current (not due)",
    "overdue": "Overdue", "in_plan": "In a plan", "frozen": "Frozen",
}


# --------------------------------------------------------------------------- #
# Value helpers
# --------------------------------------------------------------------------- #
def norm_cuil(v):
    """11-digit CUIL from a cell that may be a number (e.g. 27205346606.0) or text."""
    if v in (None, ""):
        return ""
    if isinstance(v, bool):
        return ""
    if isinstance(v, (int, float)):
        s = str(int(round(v)))
    else:
        s = re.sub(r"\D", "", str(v))
    return s if len(s) == 11 else ""


def norm_dni(v):
    if v in (None, ""):
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return str(int(round(v)))
    return re.sub(r"\D", "", str(v))


def num(v):
    if isinstance(v, bool) or v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s not in ("", "-", ".", "-.") else 0.0
    except ValueError:
        return 0.0


def as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return None


def effective_due(fecha, term, venc_stored):
    """Due date = Fecha Préstamo + Plazo (días), the contractual business rule.
    The stored `Vencimiento` is a live formula whose exported value is corrupt for
    a few trouble loans (pushed months into the future, inconsistent with the term);
    the derived date is consistent for the whole normal portfolio, so prefer it and
    fall back to the stored column only when Fecha/Plazo are missing."""
    if fecha and term and term > 0:
        return fecha + datetime.timedelta(days=int(round(term)))
    return venc_stored


def money(x):
    return "$ " + format(int(round(x or 0)), ",d").replace(",", ".")


# --------------------------------------------------------------------------- #
# Workbook → tables
# --------------------------------------------------------------------------- #
class Table:
    """A sheet as header + data rows, with first-wins header→index lookup
    (the workbook has a few duplicate header labels in columns we don't use)."""

    def __init__(self, rows):
        self.header = [str(x).strip() if x is not None else "" for x in (rows[0] if rows else [])]
        self.rows = rows[1:] if len(rows) > 1 else []
        self.idx = {}
        for i, h in enumerate(self.header):
            if h and h not in self.idx:
                self.idx[h] = i

    def get(self, row, name, default=None):
        i = self.idx.get(name)
        if i is None or i >= len(row):
            return default
        return row[i]

    def records(self, key):
        """Rows whose `key` column is non-empty."""
        ki = self.idx.get(key)
        if ki is None:
            return []
        out = []
        for r in self.rows:
            if ki < len(r) and r[ki] not in (None, ""):
                out.append(r)
        return out


def load_tables(path):
    sheets = xlsx_mini.read_xlsx(path)
    need = ["Prestatarios", "Clientes", "Pagos", "Cuotas"]
    missing = [s for s in need if s not in sheets]
    if missing:
        raise SystemExit("Workbook %s is missing sheet(s): %s" % (path, ", ".join(missing)))
    return {s: Table(sheets[s]) for s in need}


# --------------------------------------------------------------------------- #
# Loan classification
# --------------------------------------------------------------------------- #
def _loan_days_late(due, saldo, cuotas, today):
    """Max days past due across the loan and any unpaid installments (0 if current)."""
    dl = 0
    if saldo > PAID_EPS and due and today > due:
        dl = (today - due).days
    for cu in cuotas:
        cstate = str(cu.get("estado") or "").strip().upper()
        csaldo = cu.get("saldo")
        cdue = cu.get("due")
        unpaid = cstate != "PAGADA" and (csaldo is None or csaldo > PAID_EPS)
        if unpaid and cdue and today > cdue:
            dl = max(dl, (today - cdue).days)
    return dl


def classify_loan(loan, today):
    """Return (status, days_late) from dates/balances, reconciled with Estado."""
    estado = str(loan["estado"] or "").strip().upper()
    total, paid, saldo = loan["total"], loan["paid"], loan["saldo"]
    due = loan["venc"]
    is_paid = saldo <= PAID_EPS or (total > 0 and paid >= total - PAID_EPS)

    # Current relationship states take priority over a one-off due date.
    if estado == "CONGELADO":
        return "frozen", _loan_days_late(due, saldo, loan["cuotas"], today)
    if estado == "EN PLAN":
        return "in_plan", _loan_days_late(due, saldo, loan["cuotas"], today)

    if is_paid or estado in ("PAGADO", "SALDADO"):
        last_pay = max((p for p in loan["pay_dates"]), default=None)
        if last_pay and due:
            dl = (last_pay - due).days
            return ("repaid_late", dl) if dl > 0 else ("repaid_ontime", 0)
        return "repaid_untimed", 0

    dl = _loan_days_late(due, saldo, loan["cuotas"], today)
    return ("overdue", dl) if dl > 0 else ("current", 0)


def loan_level(status, dl):
    if status == "repaid_late":
        return _repaid_late_level(dl)
    if status == "overdue":
        return _overdue_level(dl)
    return _LEVEL.get(status, 50)


# --------------------------------------------------------------------------- #
# Build the per-client model
# --------------------------------------------------------------------------- #
def build_clients(tables, today):
    P, C, Pay, Cu = tables["Prestatarios"], tables["Clientes"], tables["Pagos"], tables["Cuotas"]

    # payments by loan id
    pays = {}
    for r in Pay.records("ID Préstamo"):
        lid = str(Pay.get(r, "ID Préstamo")).strip()
        d = as_date(Pay.get(r, "Fecha de Pago"))
        pays.setdefault(lid, []).append(d)
    for lid in pays:
        pays[lid] = [d for d in pays[lid] if d]

    # cuotas by loan id
    cuotas = {}
    for r in Cu.records("ID Préstamo"):
        lid = str(Cu.get(r, "ID Préstamo")).strip()
        cuotas.setdefault(lid, []).append({
            "due": as_date(Cu.get(r, "Fecha de Vencimiento")),
            "estado": Cu.get(r, "Estado"),
            "saldo": num(Cu.get(r, "Saldo Cuota")),
        })

    # client identity
    clients = {}
    for r in C.records("ID Cliente"):
        cid = str(C.get(r, "ID Cliente")).strip()
        clients[cid] = {
            "client_id": cid,
            "name": str(C.get(r, "Nombre") or "").strip(),
            "cuil": norm_cuil(C.get(r, "CUIL")),
            "dni": norm_dni(C.get(r, "DNI")),
            "loans": [],
        }

    # loans → attach to client
    for r in P.records("ID Préstamo"):
        lid = str(P.get(r, "ID Préstamo")).strip()
        cid = str(P.get(r, "ID Cliente") or "").strip()
        fecha = as_date(P.get(r, "Fecha Préstamo"))
        term = num(P.get(r, "Plazo (días)"))
        venc_stored = as_date(P.get(r, "Vencimiento"))
        loan = {
            "loan_id": lid, "client_id": cid,
            "capital": num(P.get(r, "Capital")),
            "term": term,
            "total": num(P.get(r, "Total a Pagar")),
            "paid": num(P.get(r, "Total Pagado")),
            "saldo": num(P.get(r, "Saldo Pendiente")),
            "estado": P.get(r, "Estado"),
            "fecha": fecha,
            "venc": effective_due(fecha, term, venc_stored),
            "venc_stored": venc_stored,
            "pay_dates": pays.get(lid, []),
            "cuotas": cuotas.get(lid, []),
        }
        loan["status"], loan["days_late"] = classify_loan(loan, today)
        loan["level"] = loan_level(loan["status"], loan["days_late"])
        if cid not in clients:          # loan referencing a client not in Clientes
            name = str(P.get(r, "Nombre") or "").strip()
            clients[cid] = {"client_id": cid, "name": name, "cuil": "",
                            "dni": norm_dni(P.get(r, "DNI")), "loans": []}
        clients[cid]["loans"].append(loan)

    return clients


# --------------------------------------------------------------------------- #
# Rating / decision / category
# --------------------------------------------------------------------------- #
def rate_of(score):
    if score >= 80:
        return "A - Low Risk"
    if score >= 60:
        return "B - Medium-Low Risk"
    if score >= 40:
        return "C - Medium Risk"
    if score >= 20:
        return "D - High Risk"
    return "E - Very High Risk"


_FLOOR_LABEL = {"C": "C - Medium Risk", "D": "D - High Risk", "E": "E - Very High Risk"}


def apply_floor(rating, floor_letter):
    """Never let the grade read better than the current worst standing allows."""
    order = "ABCDE"
    if floor_letter and order.index(rating[0]) < order.index(floor_letter):
        return _FLOOR_LABEL[floor_letter]
    return rating


def analyze_client(cl, today):
    loans = sorted(cl["loans"], key=lambda l: (l["fecha"] or datetime.date(1900, 1, 1)))
    n = len(loans)

    a = {
        "client_id": cl["client_id"], "name": cl["name"],
        "cuil": cl["cuil"], "dni": cl["dni"],
        "generated": datetime.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M"),
        "n_loans": n, "has_record": n > 0, "loans": loans,
    }

    if n == 0:
        a.update(score=0, rating="—", decision="NO RECORD",
                 category="No record", category_why="No loans on file at Impulso Crédito.",
                 total_borrowed=0, total_repaid=0, outstanding=0, counts={}, on_time_rate=None,
                 worst_now="—", avg_days_late=None)
        return a

    # recency-weighted standing + performing share
    num_ = den = 0.0
    for i, l in enumerate(loans):               # oldest→newest; newest weight 1
        g = RECENCY_DECAY ** (n - 1 - i)
        num_ += g * l["level"]
        den += g
    history_comp = num_ / den if den else 100
    performing = sum(1 for l in loans if l["status"] in _PERFORMING)
    perf_share = performing / n
    score = max(0, min(100, round(0.7 * history_comp + 0.3 * 100 * perf_share)))

    # current worst standing → rating floor + category priority
    overdue_now = [l for l in loans if l["status"] == "overdue"]
    cur_overdue_days = max((l["days_late"] for l in overdue_now), default=0)
    in_plan_now = any(l["status"] == "in_plan" for l in loans)
    frozen_now = any(l["status"] == "frozen" for l in loans)

    floor = None
    if overdue_now:
        floor = "E" if cur_overdue_days > 30 else "D"
    elif in_plan_now or frozen_now:
        floor = "C"
    rating = apply_floor(rate_of(score), floor)

    late_cnt = sum(1 for l in loans if l["status"] in ("repaid_late", "overdue"))
    late_ratio = late_cnt / n
    recent = loans[-RECENT_LOANS:]
    recent_trouble = any(l["status"] in ("repaid_late", "overdue") for l in recent)

    # decision
    if overdue_now and (cur_overdue_days > 30 or (late_ratio >= CHRONIC_RATIO and n >= 2)):
        decision = "REJECT"
    elif overdue_now or in_plan_now or frozen_now or recent_trouble or late_ratio >= CHRONIC_RATIO:
        decision = "REVIEW"
    else:
        decision = "OK"

    # category (priority: current trouble → history pattern → clean)
    closed = [l for l in loans if l["status"] in ("repaid_ontime", "repaid_untimed", "repaid_late")]
    if overdue_now:
        category = "Currently overdue"
        why = "%d loan(s) past due now (worst %d days)." % (len(overdue_now), cur_overdue_days)
    elif in_plan_now:
        category = "In a payment plan"
        why = "Currently repaying under a restructured plan."
    elif frozen_now:
        category = "Frozen (hardship)"
        why = "A loan is frozen (mora paused) pending resolution."
    elif late_ratio >= CHRONIC_RATIO and n >= 2:
        category = "Chronically late"
        why = "%d of %d loans were late or went overdue." % (late_cnt, n)
    elif late_cnt > 0:
        category = "Occasional late payer"
        why = "%d late repayment(s) across %d loans; otherwise current." % (late_cnt, n)
    elif not closed:
        category = "New — no closed loans yet"
        why = "%d active loan(s), none completed yet; paying as agreed." % n
    elif n >= 2:
        category = "Excellent — always on time"
        why = "All %d loans repaid on time, no delinquency ever." % n
    else:
        category = "Reliable — pays on time"
        why = "Single loan handled on time."

    # metrics
    counts = {}
    for l in loans:
        counts[l["status"]] = counts.get(l["status"], 0) + 1
    late_days = [l["days_late"] for l in loans if l["status"] in ("repaid_late", "overdue") and l["days_late"] > 0]
    worst_now = ("Overdue %d days" % cur_overdue_days if overdue_now else
                 "In a payment plan" if in_plan_now else
                 "Frozen" if frozen_now else "In good standing")

    a.update(
        score=score, rating=rating, decision=decision,
        category=category, category_why=why,
        total_borrowed=sum(l["capital"] for l in loans),
        total_repaid=sum(l["paid"] for l in loans),
        outstanding=sum(l["saldo"] for l in loans if l["status"] in ("current", "overdue", "in_plan", "frozen")),
        counts=counts, on_time_rate=round(100 * perf_share),
        avg_days_late=(round(sum(late_days) / len(late_days)) if late_days else None),
        worst_now=worst_now, history_comp=round(history_comp),
    )
    return a


# --------------------------------------------------------------------------- #
# Cohort + ranking
# --------------------------------------------------------------------------- #
def build_cohort(path, today=None):
    """Analyze every client holding ≥1 loan. Returns light records (for the
    Impulso ranking and for analyze_combined.py). Keyed by CUIL when known,
    else by 'DNI:<dni>' / client id so impulso-only people stay distinct."""
    today = today or datetime.date.today()
    tables = load_tables(path)
    clients = build_clients(tables, today)
    cohort = []
    for cl in clients.values():
        if not cl["loans"]:
            continue
        a = analyze_client(cl, today)
        key = a["cuil"] or ("DNI:" + a["dni"] if a["dni"] else "CID:" + a["client_id"])
        cohort.append({
            "key": key, "cuil": a["cuil"], "dni": a["dni"], "name": a["name"],
            "score": a["score"], "rating": a["rating"], "category": a["category"],
            "decision": a["decision"], "n_loans": a["n_loans"], "has_record": True,
        })
    return cohort


def rank_of(cohort, target_key, target_score, target_category):
    m = len(cohort)
    if m == 0:
        return None
    peers = [c for c in cohort if c["key"] != target_key]
    npeers = len(peers)
    better = sum(1 for c in peers if c["score"] > target_score)
    worse = sum(1 for c in peers if c["score"] < target_score)
    equal = npeers - better - worse
    avg = sum(c["score"] for c in cohort) / m
    same_cat = sum(1 for c in peers if c["category"] == target_category)
    return {"rank": better + 1, "of": m, "npeers": npeers, "better": better,
            "worse": worse, "equal": equal, "avg": avg, "same_cat": same_cat}


# --------------------------------------------------------------------------- #
# Renderers
# --------------------------------------------------------------------------- #
def render_client_cli(a, rank):
    W = 66
    L = ["=" * W,
         " IMPULSO CRÉDITO — REPAYMENT REPORT",
         " %s   ·   generated %s" % (a["name"] or "(no name)", a["generated"]),
         " CUIL %s   ·   DNI %s   ·   client %s" % (a["cuil"] or "—", a["dni"] or "—", a["client_id"]),
         "=" * W, "",
         "VERDICT",
         "  Impulso score     : %s" % ("%d / 100   (%s)" % (a["score"], a["rating"]) if a["has_record"] else "no record"),
         "  Decision          : %s" % a["decision"],
         "  Payment profile   : %s" % a["category"],
         "  Basis             : %s" % a["category_why"]]
    if not a["has_record"]:
        L += ["", "  No loans on file at Impulso Crédito.", "", "=" * W]
        return "\n".join(L)

    c = a["counts"]
    L += ["", "PORTFOLIO WITH IMPULSO",
          "  Loans             : %d" % a["n_loans"],
          "  On-time rate      : %d%% of loans in good standing" % a["on_time_rate"],
          "  Current standing  : %s" % a["worst_now"],
          "  Total borrowed    : %s" % money(a["total_borrowed"]),
          "  Total repaid      : %s" % money(a["total_repaid"]),
          "  Outstanding       : %s" % money(a["outstanding"]),
          "  Breakdown         : " + ", ".join("%s×%d" % (_STATUS_LABEL.get(k, k), v) for k, v in sorted(c.items())),
          "  Avg days late     : %s" % ("%d" % a["avg_days_late"] if a["avg_days_late"] is not None else "—")]

    L += ["", "LOANS  (oldest → newest)",
          "  %-10s %-11s %12s %5s %-20s %s" % ("Loan", "Date", "Capital", "Term", "Status", "Late")]
    for l in a["loans"]:
        flag = _reconcile_flag(l)
        L.append("  %-10s %-11s %12s %4dd %-20s %s%s" % (
            l["loan_id"], (l["fecha"].isoformat() if l["fecha"] else "—"), money(l["capital"]),
            int(l["term"]), _STATUS_LABEL.get(l["status"], l["status"]),
            ("%dd" % l["days_late"] if l["days_late"] else "—"), flag))

    L += _rank_lines_cli(a, rank)
    L += ["", _footer_note(), "=" * W]
    return "\n".join(L)


def _reconcile_flag(l):
    """Note when the stored Vencimiento is inconsistent with Fecha+Plazo (corrupt
    export value): we used the derived due date instead."""
    stored = l.get("venc_stored")
    due = l.get("venc")
    if stored and due and abs((stored - due).days) > 3:
        return "  ⚠ stored Vencimiento %s inconsistent with term; used %s" % (
            stored.isoformat(), due.isoformat())
    return ""


def _rank_lines_cli(a, rank):
    if not rank or rank["of"] <= 1:
        return ["", "PORTFOLIO RANKING", "  Only this borrower in the portfolio — nothing to compare."]
    pos = "above average" if a["score"] >= rank["avg"] else "below average"
    tie = "   (%d tied)" % rank["equal"] if rank["equal"] else ""
    return ["", "PORTFOLIO RANKING  (%d borrowers; rank 1 = best payer)" % rank["of"],
            "  Rank              : %d of %d" % (rank["rank"], rank["of"]),
            "  Score             : %d   ·   portfolio average %.0f (%s)" % (a["score"], rank["avg"], pos),
            "  Better than       : %d of %d peers%s" % (rank["worse"], rank["npeers"], tie),
            "  Same profile      : %d of %d peers share '%s'" % (rank["same_cat"], rank["npeers"], a["category"])]


def _footer_note():
    return ("NOTES\n"
            "  • Impulso score is first-party repayment behaviour with Impulso Crédito\n"
            "    (on-time / late / overdue / plan / frozen), recency-weighted. It is\n"
            "    independent of, and complementary to, the BCRA score.\n"
            "  • Effective status is derived from due dates + balances and reconciled\n"
            "    with the sheet's Estado; a ⚠ marks a stale cached label.\n"
            "  • Amounts in ARS.")


def render_client_md(a, rank):
    L = ["# Impulso Crédito — Repayment Report — %s" % (a["name"] or "(no name)"),
         "",
         "**CUIL:** %s  ·  **DNI:** %s  ·  **Client:** %s  ·  **Generated:** %s"
         % (a["cuil"] or "—", a["dni"] or "—", a["client_id"], a["generated"]),
         "",
         "## Verdict", "", "| | |", "|---|---|",
         "| Impulso score | **%s** |" % ("%d / 100 (%s)" % (a["score"], a["rating"]) if a["has_record"] else "no record"),
         "| Decision | **%s** |" % a["decision"],
         "| Payment profile | **%s** |" % a["category"],
         "| Basis | %s |" % a["category_why"]]
    if not a["has_record"]:
        L += ["", "> No loans on file at Impulso Crédito.", "", _md_footer()]
        return "\n".join(L)

    c = a["counts"]
    L += ["", "## Portfolio with Impulso", "", "| Metric | Value |", "|---|---|",
          "| Loans | %d |" % a["n_loans"],
          "| On-time rate | %d%% of loans in good standing |" % a["on_time_rate"],
          "| Current standing | %s |" % a["worst_now"],
          "| Total borrowed | %s |" % money(a["total_borrowed"]),
          "| Total repaid | %s |" % money(a["total_repaid"]),
          "| Outstanding | %s |" % money(a["outstanding"]),
          "| Breakdown | %s |" % ", ".join("%s×%d" % (_STATUS_LABEL.get(k, k), v) for k, v in sorted(c.items())),
          "| Avg days late | %s |" % ("%d" % a["avg_days_late"] if a["avg_days_late"] is not None else "—")]

    L += ["", "## Loans", "", "_Oldest → newest._", "",
          "| Loan | Date | Capital | Term | Status | Days late | Note |",
          "|---|---|---:|---:|---|---:|---|"]
    for l in a["loans"]:
        note = _reconcile_flag(l).strip().lstrip("⚠ ").strip()
        L.append("| %s | %s | %s | %dd | %s | %s | %s |" % (
            l["loan_id"], (l["fecha"].isoformat() if l["fecha"] else "—"), money(l["capital"]),
            int(l["term"]), _STATUS_LABEL.get(l["status"], l["status"]),
            (l["days_late"] if l["days_late"] else "—"), ("⚠ " + note) if note else ""))

    L += _rank_lines_md(a, rank)
    L += ["", _md_footer()]
    return "\n".join(L)


def _rank_lines_md(a, rank):
    if not rank or rank["of"] <= 1:
        return ["", "## Portfolio ranking", "", "_Only this borrower in the portfolio._"]
    pos = "above average" if a["score"] >= rank["avg"] else "below average"
    tie = "  (%d tied)" % rank["equal"] if rank["equal"] else ""
    return ["", "## Portfolio ranking", "",
            "Compared against **%d** borrowers (rank 1 = best payer):" % rank["of"], "",
            "- **Rank %d of %d**" % (rank["rank"], rank["of"]),
            "- Score %d vs portfolio average %.0f (%s)" % (a["score"], rank["avg"], pos),
            "- Better than **%d of %d** peers%s" % (rank["worse"], rank["npeers"], tie),
            "- Same profile ('%s'): %d of %d peers" % (a["category"], rank["same_cat"], rank["npeers"])]


def _md_footer():
    return ("---\n\n"
            "> **Notes.** The Impulso score reflects first-party repayment behaviour with "
            "Impulso Crédito (on-time / late / overdue / plan / frozen), recency-weighted and "
            "independent of the BCRA score. Effective status is derived from due dates + balances "
            "and reconciled with the sheet's `Estado` (⚠ = stale cached label). Amounts in ARS.")


def render_leaderboard_cli(cohort):
    ranked = sorted(cohort, key=lambda c: (-c["score"], c["name"]))
    L = ["IMPULSO CRÉDITO — BORROWER RANKING  (%d borrower%s, best → worst)"
         % (len(ranked), "" if len(ranked) == 1 else "s"), "",
         "  %-4s %6s %-22s %-26s %-9s %-4s %s" % ("#", "Score", "Rating", "Payment profile", "Decision", "Lns", "Name")]
    for i, c in enumerate(ranked, 1):
        L.append("  %-4d %6d %-22s %-26s %-9s %-4d %s" % (
            i, c["score"], c["rating"], c["category"][:26], c["decision"], c["n_loans"], (c["name"] or "")[:30]))
    if ranked:
        L += ["", "  Portfolio average score: %.0f" % (sum(c["score"] for c in ranked) / len(ranked))]
    return "\n".join(L)


def render_leaderboard_md(cohort):
    ranked = sorted(cohort, key=lambda c: (-c["score"], c["name"]))
    gen = datetime.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M")
    L = ["# Impulso Crédito — Borrower Ranking", "",
         "_Generated %s · %d borrower(s), best → worst. First-party repayment behaviour only._"
         % (gen, len(ranked)), "",
         "| # | CUIL | Score | Rating | Payment profile | Decision | Loans | Name |",
         "|---|---|---:|---|---|---|---:|---|"]
    for i, c in enumerate(ranked, 1):
        L.append("| %d | %s | %d | %s | %s | %s | %d | %s |" % (
            i, c["cuil"] or "—", c["score"], c["rating"], c["category"], c["decision"], c["n_loans"], c["name"] or ""))
    if ranked:
        L += ["", "**Portfolio average score:** %.0f" % (sum(c["score"] for c in ranked) / len(ranked))]
    return "\n".join(L)


# --------------------------------------------------------------------------- #
# Lookup + main
# --------------------------------------------------------------------------- #
def find_client(clients, token):
    """Match a client by CUIL, DNI, or ID Cliente (case-insensitive)."""
    digits = re.sub(r"\D", "", token)
    tok = token.strip().upper()
    for cl in clients.values():
        if cl["cuil"] and cl["cuil"] == digits:
            return cl
    for cl in clients.values():
        if cl["dni"] and cl["dni"] == digits:
            return cl
    for cl in clients.values():
        if cl["client_id"].upper() == tok:
            return cl
    return None


def parse_args(argv):
    here = Path(__file__).resolve().parent
    p = argparse.ArgumentParser(
        description="Analyze Impulso Crédito loan history and rank borrowers.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("who", nargs="?", help="CUIL, DNI, or ID Cliente to report on")
    p.add_argument("--all", action="store_true", help="rank all borrowers → out/_impulso_ranking.md")
    p.add_argument("--xlsx", default=str(here / DEFAULT_XLSX), help="Impulso workbook (.xlsx export)")
    p.add_argument("--out", default=str(here / "out"), help="report output folder")
    p.add_argument("--no-save", action="store_true", help="print only; don't write the .md")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    xlsx = Path(args.xlsx)
    out_dir = Path(args.out)
    today = datetime.date.today()

    if not xlsx.exists():
        logging.error("Workbook not found: %s\n"
                      "Export the sheet: File ▸ Download ▸ Microsoft Excel (.xlsx) → save there.", xlsx)
        return 2

    if args.all:
        cohort = build_cohort(str(xlsx), today)
        if not cohort:
            logging.error("No borrowers with loans found in %s.", xlsx)
            return 2
        print(render_leaderboard_cli(cohort))
        if not args.no_save:
            out_dir.mkdir(parents=True, exist_ok=True)
            path = out_dir / "_impulso_ranking.md"
            path.write_text(render_leaderboard_md(cohort) + "\n", encoding="utf-8")
            logging.info("\nSaved %s", path)
        return 0

    if not args.who:
        logging.error("Pass a CUIL/DNI/ID Cliente, or --all for the ranking. See --help.")
        return 2

    tables = load_tables(str(xlsx))
    clients = build_clients(tables, today)
    cl = find_client(clients, args.who)
    if not cl:
        logging.error("No client matched '%s' (by CUIL, DNI, or ID Cliente).", args.who)
        return 2

    a = analyze_client(cl, today)
    cohort = build_cohort(str(xlsx), today)
    key = a["cuil"] or ("DNI:" + a["dni"] if a["dni"] else "CID:" + a["client_id"])
    rank = rank_of(cohort, key, a["score"], a["category"]) if a["has_record"] else None

    print(render_client_cli(a, rank))
    if not args.no_save:
        out_dir.mkdir(parents=True, exist_ok=True)
        stem = a["cuil"] or a["dni"] or a["client_id"]
        path = out_dir / ("%s.impulso.report.md" % stem)
        path.write_text(render_client_md(a, rank) + "\n", encoding="utf-8")
        logging.info("\nSaved %s", path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
