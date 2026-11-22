#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Repayment likelihood — a **predictive, data-calibrated** profile of who repays
on time, learned from the actual Impulso book + BCRA files in ``out/``.

Unlike the descriptive scores (analyze_bcra / analyze_impulso / analyze_combined,
which summarize past behaviour with hand-set weights), this module:

  1. builds a **leakage-free training set** — one row per *resolved* loan, with
     features computed strictly as of that loan's origination date;
  2. fits an **L2-regularized logistic regression** (pure stdlib IRLS/Newton —
     no sklearn/numpy) for P(loan repaid with ≤ 3 days of delay);
  3. validates it honestly with **leave-one-client-out cross-validation** plus
     baselines, and only "ships" probabilities when a quality gate passes;
  4. scores **existing borrowers** (next-loan probability), **BCRA-only
     prospects**, and **pending applicants** from «Nuevos Prestatarios»;
  5. renders the **repayment profile** (which factors matter in THIS book) to
     ``out/_likelihood.md``.

Statistical honesty: at ~75 labeled loans / ~19 bad outcomes this is a
calibrated, validated summary of a term/amount/history table — NOT a learned
risk engine. Coefficients are directional (EPV ≈ 3); the dominant driver is the
loan term; BCRA situación shows no measurable signal in the *approved* book
(range restriction — approval already filters the worst files), so this score
complements, and does not replace, the upstream BCRA reject gate.

Usage:
    python3 bcra/analyze_likelihood.py --all          # profile + ranking → out/_likelihood.md
    python3 bcra/analyze_likelihood.py <cuil|dni|C-x> # one person, with per-feature drivers
    python3 bcra/analyze_likelihood.py --selftest     # deterministic numeric self-checks
"""

import argparse
import datetime
import json
import logging
import math
import random
import re
import sys
from pathlib import Path

import xlsx_mini
import analyze_impulso  # Table, load_tables, build_clients, norm_dni/cuil, rate_of, money
import analyze_bcra     # history_series, latest_snapshot, denom_of

HERE = Path(__file__).resolve().parent
DEFAULT_XLSX = HERE / "out" / "_impulso.xlsx"
DEFAULT_OUT = HERE / "out"
DEFAULT_EXPORT = DEFAULT_OUT / "_likelihood.md"

# ---------------------------------------------------------------------------- #
# Model constants (pre-registered — no in-loop feature/λ selection)
# ---------------------------------------------------------------------------- #
GRACE_DAYS = 3          # "on time" = repaid within this many days past due
L2_LAMBDA = 2.0         # ridge strength, fixed a priori (EPV≈3 regime)
L2_SENSITIVITY = (0.5, 2.0, 8.0)
BOOT_REPS = 200         # client-level bootstrap replicates
BOOT_SEED = 42
BCRA_BAD_SIT = 3        # situación ≥ 3 counts as a bad BCRA mark
BCRA_WINDOW = 12        # months before origination scanned for bad marks
PLAZO_LONG_DAYS = 30    # term ≥ this → "long" (the dominant risk factor)
CAP_SPLIT = 100_000     # profile bucket split for Capital (report only)
N_PRIOR_CAP = 3         # n_prior_good is capped here
AUC_GATE = 0.65         # shipping gate: LOOCV AUC must reach this
P_CLAMP = (0.02, 0.98)  # never display 0%/100%
STALE_DAYS = 30         # warn when the export looks older than this

FEATURES = ["plazo_long", "logcap", "n_prior_good", "prior_any_bad",
            "bcra_missing", "bcra_bad12"]

FEATURE_LABEL = {
    "plazo_long": "term ≥ %d days" % PLAZO_LONG_DAYS,
    "logcap": "loan amount (log)",
    "n_prior_good": "prior on-time loans (capped at %d)" % N_PRIOR_CAP,
    "prior_any_bad": "a prior late/overdue loan",
    "bcra_missing": "no usable BCRA file",
    "bcra_bad12": "BCRA situación ≥ %d in prior %d months" % (BCRA_BAD_SIT, BCRA_WINDOW),
}

# Display decision reuses the combined-score cutoffs so ranking row colors work.
def decision_of_p(pct):
    return "OK" if pct >= 60 else ("REVIEW" if pct >= 35 else "REJECT")


def tier_of_p(pct):
    """A–E letter on the same scale as the other scores (A ≥ 80 … E < 20)."""
    return analyze_impulso.rate_of(pct)[0]


# ---------------------------------------------------------------------------- #
# Labels (per resolved loan)
# ---------------------------------------------------------------------------- #
def label_loan(loan, today):
    """1 = repaid ≤ GRACE_DAYS late · 0 = definitively not · None = indeterminate.

    overdue/in_plan/frozen past the grace window are resolved-BAD (they already
    failed "repaid on time") even if they eventually repay; inside the window
    they are indeterminate, never prematurely bad."""
    st, dl = loan["status"], loan["days_late"]
    if st == "repaid_ontime":
        return 1
    if st == "repaid_late":
        return 1 if dl <= GRACE_DAYS else 0
    if st == "repaid_untimed":
        return None          # closed but timing unverifiable
    if st == "current":
        return None          # censored
    if st in ("overdue", "in_plan", "frozen"):
        if dl > GRACE_DAYS:
            return 0
        venc = loan.get("venc")
        if venc and (today - venc).days > GRACE_DAYS:
            return 0
        return None
    return None


# ---------------------------------------------------------------------------- #
# BCRA as-of features
# ---------------------------------------------------------------------------- #
def load_bcra_map(out_dir):
    """cuil -> combined json for every real <cuil>.json in out_dir (placeholders
    parse fine but yield an empty series → "missing" downstream)."""
    out = {}
    for p in sorted(Path(out_dir).glob("*.json")):
        if not re.fullmatch(r"\d{11}", p.stem):
            continue
        try:
            out[p.stem] = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
    return out


def _full_series(bjson):
    """historicas series + the current /Deudas period appended when newer.
    Only `periodo` and `worst` are used here, so the montos scale is irrelevant."""
    series = analyze_bcra.history_series(bjson)
    ents, per = analyze_bcra.latest_snapshot(bjson)
    if per and (not series or per > series[-1]["periodo"]):
        worst = max((int(e.get("situacion") or 0) for e in ents), default=0)
        series = series + [{"periodo": per, "worst": worst}]
    return series


def bcra_asof(bjson, asof):
    """BCRA state strictly BEFORE `asof`'s month: periods with periodo < asof's
    YYYYMM. The asof-month período reflects end-of-month state published weeks
    later, so including it would leak post-origination information into training
    features (train and live scoring use the same strict rule for consistency).
    None = missing/unusable (no json, placeholder, or the fetched series starts
    at/after asof's month)."""
    if not bjson:
        return None
    series = _full_series(bjson)
    if not series:
        return None
    ym = "%04d%02d" % (asof.year, asof.month)
    kept = [s for s in series if s["periodo"] and s["periodo"] < ym]
    if not kept:
        return None          # series exists but starts at/after origination month
    window = kept[-BCRA_WINDOW:]
    return {"worst12": max(s["worst"] for s in window), "nperiods": len(kept)}


# ---------------------------------------------------------------------------- #
# Prior internal history, strictly as of a date
# ---------------------------------------------------------------------------- #
def _prior_state(prior, d):
    """'good' | 'bad' | None for a prior loan, judged with information available
    at date d only (no peeking at how it eventually ended)."""
    st, dl = prior["status"], prior["days_late"]
    if st == "repaid_untimed":
        return None                          # timing unknowable, before or after d
    last_pay = max(prior["pay_dates"], default=None)
    if st in ("repaid_ontime", "repaid_late") and last_pay and last_pay < d:
        return "good" if dl <= GRACE_DAYS else "bad"
    venc = prior.get("venc")
    if venc and (d - venc).days > GRACE_DAYS:
        return "bad"                          # already past grace and unresolved at d
    return None                               # open and current at d


def prior_history(loans_sorted, d):
    """(n_prior_good capped, prior_any_bad) over same-client loans originated before d."""
    good = bad = 0
    for p in loans_sorted:
        if not p["fecha"] or p["fecha"] >= d:
            continue
        st = _prior_state(p, d)
        if st == "good":
            good += 1
        elif st == "bad":
            bad += 1
    return min(good, N_PRIOR_CAP), (1 if bad else 0)


# ---------------------------------------------------------------------------- #
# Features (single path for training AND scoring)
# ---------------------------------------------------------------------------- #
def featurize(capital, plazo, n_prior_good, prior_any_bad, bcra):
    """bcra = result of bcra_asof() or None (missing)."""
    return {
        "plazo_long": 1.0 if (plazo or 0) >= PLAZO_LONG_DAYS else 0.0,
        "logcap": math.log(max(float(capital or 0), 1.0)),
        "n_prior_good": float(min(int(n_prior_good or 0), N_PRIOR_CAP)),
        "prior_any_bad": 1.0 if prior_any_bad else 0.0,
        "bcra_missing": 0.0 if bcra else 1.0,
        "bcra_bad12": 1.0 if (bcra and bcra["worst12"] >= BCRA_BAD_SIT) else 0.0,
    }


def build_dataset(clients, bcra_map, today):
    """One row per labeled loan: {loan_id, client_id, cuil, name, fecha, y, feats,
    status, days_late}. Also returns exclusion counts for the report."""
    rows, excluded = [], {"current": 0, "untimed": 0, "in_grace": 0, "no_fecha": 0}
    for cl in clients.values():
        loans = sorted([l for l in cl["loans"]],
                       key=lambda l: (l["fecha"] or datetime.date(1900, 1, 1)))
        for l in loans:
            if not l["fecha"]:
                excluded["no_fecha"] += 1
                continue
            y = label_loan(l, today)
            if y is None:
                st = l["status"]
                excluded["current" if st == "current" else
                         "untimed" if st == "repaid_untimed" else "in_grace"] += 1
                continue
            npg, pab = prior_history(loans, l["fecha"])
            bc = bcra_asof(bcra_map.get(cl["cuil"]), l["fecha"]) if cl["cuil"] else None
            rows.append({
                "loan_id": l["loan_id"], "client_id": cl["client_id"],
                "cuil": cl["cuil"], "dni": cl["dni"], "name": cl["name"],
                "fecha": l["fecha"], "y": y, "status": l["status"],
                "days_late": l["days_late"],
                "feats": featurize(l["capital"], l["term"], npg, pab, bc),
            })
    rows.sort(key=lambda r: (r["client_id"], r["fecha"], r["loan_id"]))
    return rows, excluded


# ---------------------------------------------------------------------------- #
# Pure-python ridge logistic regression (IRLS / Newton)
# ---------------------------------------------------------------------------- #
def _sigmoid(z):
    if z >= 0:
        ez = math.exp(-z) if z < 700 else 0.0
        return 1.0 / (1.0 + ez)
    ez = math.exp(z) if z > -700 else 0.0
    return ez / (1.0 + ez)


def _gauss_solve(A, b):
    """Solve A·x = b by Gaussian elimination with partial pivoting (A: n×n lists)."""
    n = len(A)
    M = [A[i][:] + [b[i]] for i in range(n)]
    for col in range(n):
        piv = max(range(col, n), key=lambda r: abs(M[r][col]))
        if abs(M[piv][col]) < 1e-12:
            raise ArithmeticError("singular system in IRLS")
        M[col], M[piv] = M[piv], M[col]
        pv = M[col][col]
        for r in range(col + 1, n):
            f = M[r][col] / pv
            if f:
                for c in range(col, n + 1):
                    M[r][c] -= f * M[col][c]
    x = [0.0] * n
    for r in range(n - 1, -1, -1):
        s = M[r][n] - sum(M[r][c] * x[c] for c in range(r + 1, n))
        x[r] = s / M[r][r]
    return x


def _xrow(feats, feat_names, mean_logcap, std_logcap):
    """Design row [1, x1..xk]; only logcap is standardized."""
    out = [1.0]
    for f in feat_names:
        v = feats.get(f, 0.0)
        if f == "logcap":
            v = (v - mean_logcap) / std_logcap
        out.append(v)
    return out


def fit(rows, l2=L2_LAMBDA, feats=None):
    """Ridge logistic via IRLS. Deterministic: zero init, fixed order, no RNG.
    Intercept unpenalized. Returns the model dict used by predict()."""
    feat_names = list(feats) if feats else list(FEATURES)
    logs = [r["feats"]["logcap"] for r in rows]
    mean_lc = sum(logs) / len(logs) if logs else 0.0
    var_lc = sum((v - mean_lc) ** 2 for v in logs) / len(logs) if logs else 0.0
    std_lc = max(math.sqrt(var_lc), 1e-9)

    X = [_xrow(r["feats"], feat_names, mean_lc, std_lc) for r in rows]
    y = [float(r["y"]) for r in rows]
    k = len(feat_names) + 1
    w = [0.0] * k
    iters, converged = 0, False
    for iters in range(1, 51):
        p = [_sigmoid(sum(wj * xj for wj, xj in zip(w, x))) for x in X]
        # gradient of penalized NLL (we solve for the Newton step directly)
        g = [0.0] * k
        for i, x in enumerate(X):
            resid = y[i] - p[i]
            for j in range(k):
                g[j] += x[j] * resid
        for j in range(1, k):                      # intercept unpenalized
            g[j] -= l2 * w[j]
        A = [[0.0] * k for _ in range(k)]
        for i, x in enumerate(X):
            wi = max(p[i] * (1.0 - p[i]), 1e-9)
            for a in range(k):
                xa = x[a] * wi
                for b in range(a, k):
                    A[a][b] += xa * x[b]
        for a in range(k):
            for b in range(a):
                A[a][b] = A[b][a]
        for j in range(1, k):
            A[j][j] += l2
        delta = _gauss_solve(A, g)
        w = [wj + dj for wj, dj in zip(w, delta)]
        if max(abs(d) for d in delta) < 1e-10:
            converged = True
            break
    n_bad = sum(1 for v in y if v == 0.0)
    return {
        "w": w, "features": feat_names,
        "mean_logcap": mean_lc, "std_logcap": std_lc,
        "n": len(rows), "n_bad": n_bad, "n_good": len(rows) - n_bad,
        "base_rate": (len(rows) - n_bad) / len(rows) if rows else 0.0,
        "l2": l2, "iters": iters, "converged": converged,
    }


def predict(model, feats):
    x = _xrow(feats, model["features"], model["mean_logcap"], model["std_logcap"])
    return _sigmoid(sum(wj * xj for wj, xj in zip(model["w"], x)))


def _penalized_grad(model, rows):
    """max|gradient| of the penalized NLL at the fitted weights (selftest)."""
    k = len(model["features"]) + 1
    g = [0.0] * k
    for r in rows:
        x = _xrow(r["feats"], model["features"], model["mean_logcap"], model["std_logcap"])
        resid = float(r["y"]) - _sigmoid(sum(wj * xj for wj, xj in zip(model["w"], x)))
        for j in range(k):
            g[j] += x[j] * resid
    for j in range(1, k):
        g[j] -= model["l2"] * model["w"][j]
    return max(abs(v) for v in g)


# ---------------------------------------------------------------------------- #
# Validation: leave-one-CLIENT-out CV + metrics + baselines + gate
# ---------------------------------------------------------------------------- #
def _auc(pairs):
    """Mann-Whitney AUC over (p, y) pairs; ties count 0.5. None if degenerate."""
    pos = [p for p, y in pairs if y == 1]
    neg = [p for p, y in pairs if y == 0]
    if not pos or not neg:
        return None
    wins = ties = 0
    for pp in pos:
        for pn in neg:
            if pp > pn:
                wins += 1
            elif pp == pn:
                ties += 1
    return (wins + 0.5 * ties) / (len(pos) * len(neg))


def _metrics(pairs):
    n = len(pairs)
    brier = sum((p - y) ** 2 for p, y in pairs) / n
    eps = 1e-6
    logloss = -sum(y * math.log(max(p, eps)) + (1 - y) * math.log(max(1 - p, eps))
                   for p, y in pairs) / n
    acc = sum(1 for p, y in pairs if (p >= 0.5) == (y == 1)) / n
    return {"auc": _auc(pairs), "brier": brier, "logloss": logloss, "acc": acc, "n": n}


def _calibration(pairs, bins=3):
    """Tercile calibration: [{lo, hi, n, mean_p, obs}] over held-out predictions.
    Ties sort by p only (label-neutral) so bin boundaries never depend on y."""
    srt = sorted(pairs, key=lambda t: t[0])
    n = len(srt)
    if n < bins:
        return []
    out, start = [], 0
    for b in range(bins):
        size = n // bins + (1 if b < n % bins else 0)
        chunk = srt[start:start + size]
        start += size
        ps = [p for p, _ in chunk]
        ys = [y for _, y in chunk]
        out.append({"lo": min(ps), "hi": max(ps), "n": len(chunk),
                    "mean_p": sum(ps) / len(chunk), "obs": sum(ys) / len(chunk)})
    return out


def loocv_by_client(rows, l2=L2_LAMBDA, feats=None):
    """Leave-one-CLIENT-out: refit (incl. restandardization) without each client,
    predict their loans. Returns pooled metrics + calibration + raw preds."""
    by_client = {}
    for r in rows:
        by_client.setdefault(r["client_id"], []).append(r)
    preds = []
    for cid in sorted(by_client):
        train = [r for r in rows if r["client_id"] != cid]
        if not train:
            continue
        m = fit(train, l2, feats)
        for r in by_client[cid]:
            preds.append((r["loan_id"], predict(m, r["feats"]), r["y"]))
    pairs = [(p, y) for _, p, y in preds]
    out = _metrics(pairs)
    out["calibration"] = _calibration(pairs)
    out["preds"] = preds
    out["folds"] = len(by_client)
    return out


def loocv_constant(rows):
    """Baseline: every held-out loan gets the train-fold base rate."""
    by_client = {}
    for r in rows:
        by_client.setdefault(r["client_id"], []).append(r)
    pairs = []
    for cid in sorted(by_client):
        train = [r for r in rows if r["client_id"] != cid]
        if not train:
            continue
        base = sum(r["y"] for r in train) / len(train)
        pairs += [(base, r["y"]) for r in by_client[cid]]
    return _metrics(pairs)


def gate(validation, constant_brier):
    """shipped iff AUC ≥ AUC_GATE AND Brier beats the constant baseline AND the
    tercile observed rates are monotone nondecreasing."""
    auc = validation.get("auc")
    if auc is None or auc < AUC_GATE:
        return False, "LOOCV AUC %s < %.2f" % ("—" if auc is None else "%.2f" % auc, AUC_GATE)
    if validation["brier"] >= constant_brier:
        return False, "Brier %.3f does not beat the base-rate baseline (%.3f)" % (
            validation["brier"], constant_brier)
    cal = validation.get("calibration") or []
    obs = [b["obs"] for b in cal]
    if any(obs[i] > obs[i + 1] + 1e-9 for i in range(len(obs) - 1)):
        return False, "calibration terciles are not monotone"
    return True, "AUC %.2f · Brier %.3f < baseline %.3f · terciles monotone" % (
        auc, validation["brier"], constant_brier)


def bootstrap_intervals(rows, l2=L2_LAMBDA, reps=BOOT_REPS, seed=BOOT_SEED):
    """Client-level bootstrap → 95% percentile intervals on each odds ratio."""
    by_client = {}
    for r in rows:
        by_client.setdefault(r["client_id"], []).append(r)
    cids = sorted(by_client)
    rng = random.Random(seed)
    samples = {f: [] for f in FEATURES}
    done = 0
    for _ in range(reps):
        sample = []
        for _ in cids:
            sample += by_client[rng.choice(cids)]
        if not sample or all(r["y"] == 1 for r in sample) or all(r["y"] == 0 for r in sample):
            continue
        try:
            m = fit(sample, l2)
        except ArithmeticError:
            continue
        for j, f in enumerate(FEATURES, start=1):
            samples[f].append(math.exp(m["w"][j]))
        done += 1
    out = {}
    for f in FEATURES:
        s = sorted(samples[f])
        if len(s) >= 20:
            lo = s[max(0, int(round(0.025 * (len(s) - 1))))]
            hi = s[min(len(s) - 1, int(round(0.975 * (len(s) - 1))))]
            out[f] = (lo, hi)
        else:
            out[f] = None
    return {"or_ci": out, "reps_done": done}


# ---------------------------------------------------------------------------- #
# Scoring (people + applicants)
# ---------------------------------------------------------------------------- #
def _drivers(model, feats):
    """Top ± contributions as percentage points: Δp of dropping each feature
    from the logit (honest leave-one-feature-out attribution)."""
    x = _xrow(feats, model["features"], model["mean_logcap"], model["std_logcap"])
    z = sum(wj * xj for wj, xj in zip(model["w"], x))
    p = _sigmoid(z)
    items = []
    for j, f in enumerate(model["features"], start=1):
        contrib = model["w"][j] * x[j]
        if abs(contrib) < 1e-9:
            continue
        dpp = (p - _sigmoid(z - contrib)) * 100.0
        items.append((dpp, f))
    items.sort(key=lambda t: -abs(t[0]))
    or_ci = model.get("or_ci") or {}
    parts = []
    for dpp, f in items[:3]:
        label = FEATURE_LABEL[f]
        if f == "n_prior_good":
            label = "%d prior on-time loan(s)" % int(feats.get("n_prior_good", 0))
        if f == "logcap":
            label = "loan amount"
        ci = or_ci.get(f)
        weak = " (uncertain)" if ci and ci[0] <= 1.0 <= ci[1] else ""
        parts.append("%+d pp %s%s" % (int(round(dpp)), label, weak))
    return " · ".join(parts) if parts else "near the portfolio base rate"


def _n_similar(dataset_rows, feats):
    """Training loans in the same term-bucket × history-class cell + their rate."""
    has_hist = feats["n_prior_good"] > 0 or feats["prior_any_bad"] > 0
    cell = [r for r in dataset_rows
            if r["feats"]["plazo_long"] == feats["plazo_long"]
            and ((r["feats"]["n_prior_good"] > 0 or r["feats"]["prior_any_bad"] > 0) == has_hist)]
    if not cell:
        return 0, None
    return len(cell), sum(r["y"] for r in cell) / len(cell)


def _pct(p):
    pct = int(round(100 * min(max(p, P_CLAMP[0]), P_CLAMP[1])))
    return max(2, min(98, pct))


def _score_common(model, dataset_rows, feats, basis, shipped):
    p = predict(model, feats)
    pct = _pct(p)
    n_sim, sim_rate = _n_similar(dataset_rows, feats)
    return {
        "probability": pct, "tier": tier_of_p(pct),
        "decision": decision_of_p(pct), "basis": basis,
        "drivers": _drivers(model, feats),
        "n_similar": n_sim,
        "similar_rate": (None if sim_rate is None else int(round(100 * sim_rate))),
        "shipped": shipped, "feats": feats,
    }


def score_person(model, dataset_rows, client, bcra_map, today, shipped=True,
                 median_cap=100_000.0):
    """P(next loan on time) for an existing client. Reference loan = last Capital
    (fallbacks: median of own loans, then the PORTFOLIO median — never a $1
    placeholder, which would extrapolate far outside training support) at 30
    days — the assumption is surfaced in the UI/report."""
    loans = sorted([l for l in client["loans"] if l["fecha"]],
                   key=lambda l: l["fecha"])
    npg, pab = prior_history(loans, today + datetime.timedelta(days=1))
    cap = loans[-1]["capital"] if loans else 0.0
    if not cap and loans:
        caps = sorted(l["capital"] for l in loans if l["capital"])
        cap = caps[len(caps) // 2] if caps else 0.0
    assumed = not cap
    if assumed:
        cap = median_cap
    bc = bcra_asof(bcra_map.get(client["cuil"]), today) if client["cuil"] else None
    feats = featurize(cap, PLAZO_LONG_DAYS, npg, pab, bc)
    basis = "repeat" if (npg > 0 or pab) else ("new" if bc else "no-BCRA")
    out = _score_common(model, dataset_rows, feats, basis, shipped)
    if assumed:
        out["drivers"] = ("amount unknown → assumed portfolio median %s · "
                          % analyze_impulso.money(cap)) + out["drivers"]
    out.update(token=(client["cuil"] or ("DNI:" + client["dni"] if client["dni"]
                                         else "CID:" + client["client_id"])),
               cuil=client["cuil"], dni=client["dni"], name=client["name"],
               client_id=client["client_id"],
               ref_capital=cap, ref_term=PLAZO_LONG_DAYS)
    return out


def score_bcra_only(model, dataset_rows, cuil, bjson, median_cap, today, shipped=True):
    """A BCRA-only prospect (no internal record): portfolio-median reference loan."""
    bc = bcra_asof(bjson, today)
    feats = featurize(median_cap, PLAZO_LONG_DAYS, 0, 0, bc)
    out = _score_common(model, dataset_rows, feats, "BCRA only", shipped)
    out.update(token=cuil, cuil=cuil, dni="", client_id="",
               name=analyze_bcra.denom_of(bjson) or "(BCRA record)",
               ref_capital=median_cap, ref_term=PLAZO_LONG_DAYS)
    return out


_PLAZO_HEADERS = ["Plazo (días)", "Plazo (dias)", "Plazo (meses)", "Plazo"]


def pending_applicants(xlsx_path):
    """Pending rows of «Nuevos Prestatarios»: Resultado empty AND Rechazar? falsy.
    Values of the Plazo column are DAYS whatever the header says."""
    wb = xlsx_mini.read_xlsx(str(xlsx_path))
    rows = wb.get("Nuevos Prestatarios")
    if not rows:
        return []
    T = analyze_impulso.Table(rows)

    def _s(v):
        return "" if v is None else str(v).strip()

    def _truthy(v):
        return _s(v).upper() in ("SÍ", "SI", "TRUE", "1", "YES", "VERDADERO", "X")

    out = []
    for r in T.records("Nombre completo"):
        if _s(T.get(r, "Resultado")):
            continue
        if _truthy(T.get(r, "Rechazar?")):
            continue
        plazo = 0.0
        for h in _PLAZO_HEADERS:
            if h in T.idx:
                plazo = analyze_impulso.num(T.get(r, h))
                if plazo:
                    break
        out.append({
            "submitted": analyze_impulso.as_date(T.get(r, "Fecha de Envío")),
            "name": _s(T.get(r, "Nombre completo")),
            "dni": analyze_impulso.norm_dni(T.get(r, "DNI")),
            "cuil": analyze_impulso.norm_cuil(T.get(r, "CUIL")),
            "amount": analyze_impulso.num(T.get(r, "Monto Solicitado")),
            "term_days": plazo,
        })
    return out


def pending_cuils(xlsx_path):
    """CUILs of pending applicants (lets the GUI pre-fetch their BCRA files)."""
    try:
        return [a["cuil"] for a in pending_applicants(xlsx_path) if a["cuil"]]
    except Exception:
        return []


def score_applicant(model, dataset_rows, app, clients, bcra_map, today, shipped=True,
                    median_cap=100_000.0):
    """Pending applicant scored with their ACTUAL requested amount/term; internal
    history reused when DNI/CUIL matches an existing client. A blank/unparsable
    Monto falls back to the portfolio median (flagged) — never to a $1 loan."""
    match = None
    for cl in clients.values():
        if (app["cuil"] and cl["cuil"] == app["cuil"]) or \
           (app["dni"] and cl["dni"] == app["dni"]):
            match = cl
            break
    npg = pab = 0
    if match:
        loans = sorted([l for l in match["loans"] if l["fecha"]], key=lambda l: l["fecha"])
        npg, pab = prior_history(loans, today + datetime.timedelta(days=1))
    cuil = app["cuil"] or (match["cuil"] if match else "")
    bc = bcra_asof(bcra_map.get(cuil), today) if cuil else None
    amount = app["amount"] or median_cap
    feats = featurize(amount, app["term_days"] or PLAZO_LONG_DAYS, npg, pab, bc)
    basis = "repeat" if match and (npg > 0 or pab) else "new"
    out = _score_common(model, dataset_rows, feats, basis, shipped)
    if not app["amount"]:
        out["drivers"] = ("amount missing → assumed portfolio median %s · "
                          % analyze_impulso.money(amount)) + out["drivers"]
    out.update(token=(cuil or ("DNI:" + app["dni"] if app["dni"] else app["name"])),
               cuil=cuil, dni=app["dni"], name=app["name"], client_id=(match["client_id"] if match else ""),
               submitted=app["submitted"], amount=app["amount"], term_days=app["term_days"])
    return out


# ---------------------------------------------------------------------------- #
# Univariate repayment profile (report section 3)
# ---------------------------------------------------------------------------- #
def _bucket_rows(dataset, picker, buckets):
    out = []
    for label, pred in buckets:
        sel = [r for r in dataset if pred(picker(r))]
        if not sel:
            continue
        good = sum(r["y"] for r in sel)
        out.append({"bucket": label, "n": len(sel), "good": good,
                    "rate": good / len(sel)})
    return out


def build_profile(dataset, bcra_map, clients):
    base = sum(r["y"] for r in dataset) / len(dataset) if dataset else 0.0
    prof = []

    def add(factor, rows, note=""):
        for b in rows:
            b.update(factor=factor, lift=(b["rate"] - base), note=note)
            prof.append(b)

    add("Term (Plazo)", _bucket_rows(
        dataset, lambda r: r["feats"]["plazo_long"],
        [("15 days", lambda v: v == 0.0), ("≥ 30 days", lambda v: v == 1.0)]),
        "dominant factor; drives the model")
    add("Capital", _bucket_rows(
        dataset, lambda r: math.exp(r["feats"]["logcap"]),
        [("≤ $100k", lambda v: v <= CAP_SPLIT), ("> $100k", lambda v: v > CAP_SPLIT)]))
    add("Prior internal history", _bucket_rows(
        dataset, lambda r: r["feats"],
        [("none (first loan)", lambda f: f["n_prior_good"] == 0 and f["prior_any_bad"] == 0),
         ("≥1 prior, all on time", lambda f: f["n_prior_good"] > 0 and f["prior_any_bad"] == 0),
         ("any prior late/overdue", lambda f: f["prior_any_bad"] == 1.0)]))
    add("BCRA at origination", _bucket_rows(
        dataset, lambda r: r["feats"],
        [("worst sit ≤ 2 (12m)", lambda f: f["bcra_missing"] == 0 and f["bcra_bad12"] == 0),
         ("worst sit ≥ 3 (12m)", lambda f: f["bcra_bad12"] == 1.0),
         ("no usable BCRA file", lambda f: f["bcra_missing"] == 1.0)]),
        "flat by design of the approval filter (range restriction); "
        "missing-file gap is a correlate, not necessarily a cause")
    return prof


# ---------------------------------------------------------------------------- #
# Orchestrator (GUI entry point)
# ---------------------------------------------------------------------------- #
def build_likelihood(xlsx, out_dir, today=None, clients=None, l2=L2_LAMBDA):
    """Everything the GUI needs: model + validation + gate + profile + scores +
    applicants. Deterministic for identical inputs and `today`."""
    today = today or datetime.date.today()
    if clients is None:
        tables = analyze_impulso.load_tables(str(xlsx))
        clients = analyze_impulso.build_clients(tables, today)
    bcra_map = load_bcra_map(out_dir)

    dataset, excluded = build_dataset(clients, bcra_map, today)
    if len(dataset) < 20 or sum(1 for r in dataset if r["y"] == 0) < 5:
        return {"error": "not enough resolved loans to fit (%d labeled)" % len(dataset),
                "rows": [], "applicants": [], "profile": [], "shipped": False}

    model = fit(dataset, l2)
    validation = loocv_by_client(dataset, l2)
    base = loocv_constant(dataset)
    term_only = loocv_by_client(dataset, l2, feats=["plazo_long"])
    shipped, gate_reason = gate(validation, base["brier"])
    boot = bootstrap_intervals(dataset, l2)
    profile = build_profile(dataset, bcra_map, clients)

    # λ sensitivity (reported, never used for selection)
    sens = []
    for lam in L2_SENSITIVITY:
        v = loocv_by_client(dataset, l2=lam)
        sens.append({"l2": lam, "auc": v["auc"], "brier": v["brier"]})

    all_caps = sorted(l["capital"] for cl in clients.values() for l in cl["loans"]
                      if l["capital"])
    median_cap = all_caps[len(all_caps) // 2] if all_caps else 100_000.0

    model["or_ci"] = boot["or_ci"]      # lets _drivers flag unsupported factors

    rows, seen_cuils = [], set()
    for cl in clients.values():
        if not cl["loans"]:
            continue
        # workbook note/template rows create identity-less stub clients — skip them
        if not (cl["name"] or cl["cuil"] or cl["dni"]):
            continue
        rec = score_person(model, dataset, cl, bcra_map, today, shipped, median_cap)
        rows.append(rec)
        if cl["cuil"]:
            seen_cuils.add(cl["cuil"])
    for cuil, bjson in sorted(bcra_map.items()):
        if cuil in seen_cuils:
            continue
        if not _full_series(bjson):
            continue                      # placeholder/404 file
        rows.append(score_bcra_only(model, dataset, cuil, bjson, median_cap, today, shipped))
    rows.sort(key=lambda r: (-r["probability"], r["name"]))

    applicants = [score_applicant(model, dataset, a, clients, bcra_map, today, shipped, median_cap)
                  for a in pending_applicants(xlsx)]
    applicants.sort(key=lambda r: (-r["probability"], r["name"]))

    max_book = max((l["fecha"] for cl in clients.values() for l in cl["loans"]
                    if l["fecha"]), default=None)
    stale = bool(max_book and (today - max_book).days > STALE_DAYS)

    return {
        "model": model, "validation": validation,
        "baselines": {"constant": base, "term_only": term_only},
        "sensitivity": sens, "bootstrap": boot,
        "shipped": shipped, "gate_reason": gate_reason,
        "profile": profile, "rows": rows, "applicants": applicants,
        "dataset_n": len(dataset), "dataset_bad": model["n_bad"],
        "excluded": excluded, "today": today.isoformat(),
        "max_book_date": (max_book.isoformat() if max_book else ""),
        "stale": stale,
        "label_def": "repaid with ≤ %d days of delay" % GRACE_DAYS,
    }


# ---------------------------------------------------------------------------- #
# Renderers
# ---------------------------------------------------------------------------- #
def _fmt_auc(v):
    return "—" if v is None else "%.2f" % v


def quality_line(lk):
    if not lk or lk.get("error"):
        return "Model unavailable%s." % (" — " + lk["error"] if lk and lk.get("error") else "")
    v, b = lk["validation"], lk["baselines"]["constant"]
    line = ("Trained on %d resolved loans (%d on time · %d not) · "
            "leave-one-client-out AUC %s · Brier %.3f (base %.3f) · %s — indicative, not a verdict."
            % (lk["dataset_n"], lk["dataset_n"] - lk["dataset_bad"], lk["dataset_bad"],
               _fmt_auc(v["auc"]), v["brier"], b["brier"],
               "probabilities shown" if lk["shipped"] else
               "insufficient signal — factor profile only"))
    if lk.get("stale"):
        line += "  ⚠ workbook export looks stale (last loan %s)." % lk["max_book_date"]
    return line


def render_md(lk):
    if not lk or lk.get("error"):
        return "# Repayment likelihood\n\n_%s_\n" % quality_line(lk)
    gen = datetime.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M")
    m, v = lk["model"], lk["validation"]
    L = ["# Repayment likelihood — profile & model", "",
         "_Generated %s · data through %s · as-of date %s._" % (gen, lk["max_book_date"], lk["today"]),
         "", "**%s**" % quality_line(lk), ""]

    L += ["## 1 · What counts as \"repaid on time\"", "",
          "A loan is **good** when it was %s. Loans repaid later, or already more than "
          "%d days past due (even if they eventually pay), are **bad**." % (lk["label_def"], GRACE_DAYS), "",
          "| Outcome | Loans |", "|---|---:|",
          "| Labeled good | %d |" % (lk["dataset_n"] - lk["dataset_bad"]),
          "| Labeled bad | %d |" % lk["dataset_bad"],
          "| Excluded — still active (censored) | %d |" % lk["excluded"]["current"],
          "| Excluded — inside the grace window | %d |" % lk["excluded"]["in_grace"],
          "| Excluded — repaid, timing unverifiable | %d |" % lk["excluded"]["untimed"],
          "| Excluded — no origination date | %d |" % lk["excluded"]["no_fecha"], ""]

    base = (lk["dataset_n"] - lk["dataset_bad"]) / lk["dataset_n"]
    L += ["## 2 · Repayment profile (what actually predicts on-time payment here)", "",
          "_Base on-time rate: **%d%%**. Lift = bucket rate − base._" % round(100 * base), "",
          "| Factor | Bucket | Loans | On time | Rate | Lift |", "|---|---|---:|---:|---:|---:|"]
    last_factor = None
    for b in lk["profile"]:
        fac = b["factor"] if b["factor"] != last_factor else ""
        last_factor = b["factor"]
        L.append("| %s | %s | %d | %d | %d%% | %+d pp |" % (
            fac, b["bucket"], b["n"], b["good"], round(100 * b["rate"]), round(100 * b["lift"])))
    notes = sorted({b["note"] for b in lk["profile"] if b.get("note")})
    for nt in notes:
        L.append("")
        L.append("> %s." % nt)
    L.append("")

    L += ["## 3 · Model (ridge logistic, λ=%.1f, fit on all %d labeled loans)" % (m["l2"], m["n"]), "",
          "| Feature | Coef | Odds ratio | 95% CI (client bootstrap) | Reads as |",
          "|---|---:|---:|---|---|"]
    ci = lk["bootstrap"]["or_ci"]
    for j, f in enumerate(m["features"], start=1):
        orr = math.exp(m["w"][j])
        c = ci.get(f)
        ci_s = ("%.2f – %.2f" % c) if c else "—"
        if abs(m["w"][j]) < 1e-9:
            direction = "no measurable effect on"
        else:
            direction = "raises" if m["w"][j] > 0 else "lowers"
        if c and c[0] <= 1.0 <= c[1] and abs(m["w"][j]) >= 1e-9:
            direction += " (CI includes 1 — uncertain)"
        L.append("| %s | %+.2f | %.2f | %s | %s P(on time) |" % (
            FEATURE_LABEL[f], m["w"][j], orr, ci_s, direction))
    L += ["", "_Intercept %.2f · converged in %d IRLS iteration(s) · bootstrap %d/%d reps._"
          % (m["w"][0], m["iters"], lk["bootstrap"]["reps_done"], BOOT_REPS),
          "", "_λ sensitivity (reported, not tuned):_ " +
          " · ".join("λ=%.1f → AUC %s, Brier %.3f" % (s["l2"], _fmt_auc(s["auc"]), s["brier"])
                     for s in lk["sensitivity"]), ""]

    bl = lk["baselines"]
    L += ["## 4 · Honest validation (leave-one-client-out, %d folds)" % v["folds"], "",
          "| Model | AUC | Brier | Log-loss | Acc@0.5 |", "|---|---:|---:|---:|---:|",
          "| **This model** | %s | %.3f | %.3f | %d%% |" % (
              _fmt_auc(v["auc"]), v["brier"], v["logloss"], round(100 * v["acc"])),
          "| Term-only | %s | %.3f | %.3f | %d%% |" % (
              _fmt_auc(bl["term_only"]["auc"]), bl["term_only"]["brier"],
              bl["term_only"]["logloss"], round(100 * bl["term_only"]["acc"])),
          "| Base rate (constant) | — | %.3f | %.3f | %d%% |" % (
              bl["constant"]["brier"], bl["constant"]["logloss"], round(100 * bl["constant"]["acc"])),
          "", "Calibration (terciles of held-out predictions):", "",
          "| Predicted | n | Mean p | Observed |", "|---|---:|---:|---:|"]
    for c in v["calibration"]:
        L.append("| %d–%d%% | %d | %d%% | %d%% |" % (
            round(100 * c["lo"]), round(100 * c["hi"]), c["n"],
            round(100 * c["mean_p"]), round(100 * c["obs"])))
    L += ["", "**Shipping gate:** %s → **%s**." % (
        lk["gate_reason"], "probabilities shown" if lk["shipped"] else "factor profile only"),
        "", "> The existing Impulso/combined scores are **not** used as benchmarks: they are "
        "computed from the same outcomes this model predicts (label leakage).", ""]

    L += ["## 5 · People (P of repaying the next loan on time)", "",
          "_Existing clients: reference loan = last capital at %d days. "
          "BCRA-only prospects: portfolio-median capital._" % PLAZO_LONG_DAYS, "",
          "| # | Name | CUIL | P(repay) | Tier | Basis | Key drivers |",
          "|---|---|---|---:|---|---|---|"]
    for i, r in enumerate(lk["rows"], 1):
        L.append("| %d | %s | %s | %s | %s | %s | %s |" % (
            i, r["name"] or "—", r["cuil"] or "—",
            ("%d%%" % r["probability"]) if lk["shipped"] else "—",
            r["tier"] if lk["shipped"] else "—", r["basis"], r["drivers"]))
    L.append("")

    if lk["applicants"]:
        L += ["## 6 · Pending applicants", "",
              "| Submitted | Name | DNI | Amount | Term | P(repay) | Tier | Basis |",
              "|---|---|---|---:|---:|---:|---|---|"]
        for a in lk["applicants"]:
            L.append("| %s | %s | %s | %s | %dd | %s | %s | %s |" % (
                a["submitted"].isoformat() if a["submitted"] else "—",
                a["name"], a["dni"] or "—", analyze_impulso.money(a["amount"]),
                int(a["term_days"] or 0),
                ("%d%%" % a["probability"]) if lk["shipped"] else "—",
                a["tier"] if lk["shipped"] else "—", a["basis"]))
        L.append("")
    else:
        L += ["## 6 · Pending applicants", "", "_None pending in this export._", ""]

    L += ["## 7 · How to read a score", "",
          "- The %% is a calibrated probability that the **next/requested loan** is %s." % lk["label_def"],
          "- Tier letters follow the house scale: A ≥ 80% · B ≥ 60% · C ≥ 40% · D ≥ 20% · E < 20%.",
          "- Example (fictional): a new borrower asking $80.000 at 15 days with a clean BCRA file "
          "scores near the 15-day bucket's observed rate; the same person asking $300.000 at 90 days "
          "drops toward the long-term bucket's rate.",
          "- Scores are clamped to 2–98%: this book is far too small to justify certainty.", "",
          "## 8 · Limits (read before acting on a number)", "",
          "- **N = %d labeled loans, %d bad events** → EPV ≈ %.1f: coefficients are directional, "
          "not precise; intervals are wide." % (lk["dataset_n"], lk["dataset_bad"],
                                                lk["dataset_bad"] / len(FEATURES)),
          "- The book spans ~3.5 months — no seasonality, no macro cycle, heavy recency.",
          "- **Range restriction:** the model describes the *approved* book. BCRA situación looks "
          "flat here because approval already filters bad files — keep the upstream BCRA gate.",
          "- The missing-BCRA-file gap is a correlate of trouble, not necessarily a cause.",
          "- Term and amount are collinear by the tier rules; ridge shares the credit between them.",
          "- The model refits on every load; the shipping gate re-decides each time whether "
          "probabilities are displayed at all.", ""]
    return "\n".join(L)


def render_cli(lk):
    if not lk or lk.get("error"):
        return quality_line(lk)
    L = ["REPAYMENT LIKELIHOOD — %s" % quality_line(lk), "",
         "  %-4s %-28s %-13s %9s %5s %-10s %s" %
         ("#", "Name", "CUIL", "P(repay)", "Tier", "Basis", "Key drivers")]
    for i, r in enumerate(lk["rows"], 1):
        L.append("  %-4d %-28s %-13s %9s %5s %-10s %s" % (
            i, (r["name"] or "—")[:28], r["cuil"] or "—",
            ("%d%%" % r["probability"]) if lk["shipped"] else "—",
            r["tier"] if lk["shipped"] else "—", r["basis"], r["drivers"][:60]))
    if lk["applicants"]:
        L += ["", "PENDING APPLICANTS"]
        for a in lk["applicants"]:
            L.append("  %-28s %10s %4dd  %4s %s" % (
                a["name"][:28], analyze_impulso.money(a["amount"]), int(a["term_days"] or 0),
                ("%d%%" % a["probability"]) if lk["shipped"] else "—", a["basis"]))
    return "\n".join(L)


def render_person_md(res):
    shipped = res.get("shipped", True)
    headline = ("**P(next loan on time): %d%% · Tier %s · basis %s**"
                % (res["probability"], res["tier"], res["basis"])) if shipped else \
               ("**P(next loan on time): — (shipping gate failed; factor readout only) "
                "· basis %s**" % res["basis"])
    L = ["## 📈 Repayment likelihood — %s" % (res.get("name") or "—"), "", headline, ""]
    if res.get("ref_capital") is not None:
        L.append("_Reference loan: %s at %d days (assumption — scores change with the "
                 "actual request)._" % (analyze_impulso.money(res["ref_capital"]),
                                        res.get("ref_term", PLAZO_LONG_DAYS)))
    if res.get("amount") is not None:
        L.append("_Requested: %s at %d days._" % (analyze_impulso.money(res["amount"]),
                                                  int(res.get("term_days") or 0)))
    L += ["", "Key drivers: %s" % res["drivers"]]
    if res.get("n_similar"):
        L.append("Similar resolved loans (same term/history cell): %d · observed on-time %s%%"
                 % (res["n_similar"], res["similar_rate"]))
    if not res.get("shipped", True):
        L.append("")
        L.append("> Shipping gate failed — treat as a factor readout, not a probability.")
    return "\n".join(L)


def render_person_cli(res):
    return re.sub(r"[#*_>`]", "", render_person_md(res))


# ---------------------------------------------------------------------------- #
# Self-test (deterministic numeric checks; no workbook needed)
# ---------------------------------------------------------------------------- #
def _selftest():
    errs = []
    n_checks = [0]

    def ck(ok, msg):
        n_checks[0] += 1
        if not ok:
            errs.append(msg)

    # 1) golden sigmoid values through predict()
    m = {"w": [0.0, math.log(3.0)], "features": ["plazo_long"],
         "mean_logcap": 0.0, "std_logcap": 1.0}
    ck(abs(predict(m, {"plazo_long": 0.0, "logcap": 0.0}) - 0.5) <= 1e-12,
       "sigmoid(0) != 0.5")
    ck(abs(predict(m, {"plazo_long": 1.0, "logcap": 0.0}) - 0.75) <= 1e-12,
       "sigmoid(ln 3) != 0.75")

    # 2) tiny fixed dataset: fit converges, gradient ~ 0, refit bit-identical,
    #    monotone in a positive feature
    rows = []
    spec = [  # (plazo_long, capital, n_prior_good, y) — synthetic, fictional
        (0, 50_000, 0, 1), (0, 60_000, 0, 1), (0, 80_000, 1, 1), (0, 90_000, 2, 1),
        (1, 200_000, 0, 0), (1, 250_000, 0, 0), (1, 300_000, 1, 1), (1, 150_000, 2, 1),
        (0, 70_000, 0, 1), (1, 220_000, 0, 0), (1, 180_000, 3, 1), (0, 55_000, 1, 1),
    ]
    for i, (pl, cap, npg, y) in enumerate(spec):
        rows.append({"loan_id": "T-%d" % i, "client_id": "T-%d" % i, "y": y,
                     "fecha": datetime.date(2026, 1, 1),
                     "feats": featurize(cap, 30 if pl else 15, npg, 0, None)})
    m1 = fit(rows)
    m2 = fit(rows)
    ck(m1["w"] == m2["w"], "refit not bit-identical")
    ck(m1["converged"], "IRLS did not converge on the tiny dataset")
    ck(_penalized_grad(m1, rows) <= 1e-6,
       "gradient not ~0 at optimum (%.2e)" % _penalized_grad(m1, rows))
    f0 = featurize(100_000, 30, 0, 0, None)
    f3 = featurize(100_000, 30, 3, 0, None)
    ck(0.0 < predict(m1, f0) < 1.0 and 0.0 < predict(m1, f3) < 1.0,
       "probabilities out of (0,1)")
    ck(predict(m1, f3) > predict(m1, f0),
       "p not monotone in n_prior_good on the synthetic set")

    # 3) AUC sanity
    ck(_auc([(0.9, 1), (0.8, 1), (0.2, 0), (0.1, 0)]) == 1.0,
       "AUC of a perfect separator != 1")
    ck(abs(_auc([(0.5, 1), (0.5, 0)]) - 0.5) <= 1e-12, "tied AUC != 0.5")

    # 4) label logic
    today = datetime.date(2026, 6, 1)
    mk = lambda st, dl, venc=None: {"status": st, "days_late": dl, "venc": venc}
    ck(label_loan(mk("repaid_ontime", 0), today) == 1, "label repaid_ontime")
    ck(label_loan(mk("repaid_late", 3), today) == 1, "label repaid_late<=3")
    ck(label_loan(mk("repaid_late", 4), today) == 0, "label repaid_late>3")
    ck(label_loan(mk("current", 0), today) is None, "label current")
    ck(label_loan(mk("overdue", 10), today) == 0, "label overdue>grace")
    ck(label_loan(mk("overdue", 0, venc=today - datetime.timedelta(days=2)), today) is None,
       "label overdue inside grace should be None")

    if errs:
        for e in errs:
            logging.error("SELFTEST FAIL: %s", e)
        return 1
    logging.info("selftest OK (%d checks)", n_checks[0])
    return 0


# ---------------------------------------------------------------------------- #
# CLI
# ---------------------------------------------------------------------------- #
def parse_args(argv):
    p = argparse.ArgumentParser(
        description="Data-calibrated repayment-likelihood profile + scores.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("who", nargs="?", help="CUIL, DNI, or ID Cliente to report on")
    p.add_argument("--all", action="store_true", help="profile + full ranking → out/_likelihood.md")
    p.add_argument("--xlsx", default=str(DEFAULT_XLSX), help="Impulso workbook (.xlsx export)")
    p.add_argument("--out", default=str(DEFAULT_OUT), help="BCRA json folder / report output folder")
    p.add_argument("--no-save", action="store_true", help="print only; don't write the .md")
    p.add_argument("--l2", type=float, default=L2_LAMBDA, help="ridge strength (reporting runs only)")
    p.add_argument("--selftest", action="store_true", help="run deterministic numeric self-checks")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    if args.selftest:
        return _selftest()

    xlsx, out_dir = Path(args.xlsx), Path(args.out)
    if not xlsx.exists():
        logging.error("Workbook not found: %s", xlsx)
        return 2

    lk = build_likelihood(str(xlsx), str(out_dir), l2=args.l2)
    if lk.get("error"):
        logging.error("%s", lk["error"])
        return 2

    if args.all or not args.who:
        print(render_cli(lk))
        if not args.no_save:
            out_dir.mkdir(parents=True, exist_ok=True)
            path = out_dir / "_likelihood.md"
            path.write_text(render_md(lk) + "\n", encoding="utf-8")
            logging.info("\nSaved %s", path)
        return 0

    digits = re.sub(r"\D", "", args.who)
    tok = args.who.strip().upper()
    hit = None
    for r in lk["rows"] + lk["applicants"]:
        if (digits and (r.get("cuil") == digits or r.get("dni") == digits)) or \
           (r.get("client_id", "").upper() == tok) or (r.get("token", "").upper() == tok):
            hit = r
            break
    if not hit:
        logging.error("No person matched '%s' (by CUIL, DNI, or ID Cliente).", args.who)
        return 2
    print(render_person_cli(hit))
    print()
    print(quality_line(lk))
    if not args.no_save:
        out_dir.mkdir(parents=True, exist_ok=True)
        stem = hit.get("cuil") or hit.get("dni") or hit.get("client_id") or "person"
        path = out_dir / ("%s.likelihood.report.md" % stem)
        path.write_text(render_person_md(hit) + "\n\n" + quality_line(lk) + "\n", encoding="utf-8")
        logging.info("Saved %s", path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
