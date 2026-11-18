#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Device & fraud insights — mines the «Dispositivos» sheet of the Impulso workbook
for device/network leads, WITHOUT OCR and with no new dependencies (reads the .xlsx
with xlsx_mini, like the rest of bcra).

«Dispositivos» logs one row per public submission (loan application, prospect,
contract signing, freeze/plan request) with the device/anti-fraud trail the forms
capture: a stable fingerprint, IP + IP-geolocation, GPS, timezone, OS/browser,
dwell time, automation (webdriver), etc. This module turns that raw log into:
  • clusters: the same device fingerprint / IP / GPS / persistent token used by
    several *distinct people* (the strongest ring signal — it spans applicants,
    prospects AND loans, not just existing clients).
  • flags: per-submission leads (automation, IP-vs-address mismatch, bot-fast dwell,
    VPN/datacenter IP, locale/clock inconsistency, pasted identity, EXIF photo
    mismatch) — each only when its backing column exists.
  • log: a normalized, searchable device log (newest first).

Every column is resolved BY HEADER NAME and is optional, so the tool works on the
current 31-column schema AND on an enriched one (new columns light up automatically).

Results are **indicative** — leads for manual review, not verdicts (a shared IP can
be a household/NAT; low-precision GPS repeats; one person submits several times).

Usage:
    python3 bcra/analyze_devices.py [--xlsx out/_impulso.xlsx] [--export out/_dispositivos.md]

Note: the «Dispositivos» sheet and column names are intentionally kept in Spanish —
they are the actual labels in the workbook.
"""

import argparse
import datetime
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

import xlsx_mini
import analyze_impulso  # Table, norm_dni, norm_cuil

HERE = Path(__file__).resolve().parent
DEFAULT_XLSX = HERE / "out" / "_impulso.xlsx"
DEFAULT_EXPORT = HERE / "out" / "_dispositivos.md"

SHEET = "Dispositivos"

# Cluster thresholds (how many DISTINCT people sharing a key becomes a lead).
DEVICE_MIN = 2   # same fingerprint across ≥2 people  → high
TOKEN_MIN = 2    # same persistent token across ≥2 people → high
IP_MIN = 2       # same IP across ≥2 people (≥3 is stronger)
GPS_MIN = 2      # identical GPS fix across ≥2 people

FAST_DWELL_S = 3         # a form filled in ≤3 s looks automated/scripted
RISK_HIGH, RISK_REVIEW, RISK_OK = "High", "Review", "OK"
_RISK_ORDER = {RISK_HIGH: 0, RISK_REVIEW: 1, RISK_OK: 2}

# ISP substrings that hint at a hosting/VPN/datacenter connection (indicative only).
_HOSTING_RE = re.compile(
    r"\b(vpn|proxy|hosting|datacenter|data center|cloud|ovh|digitalocean|linode|"
    r"amazon|aws|microsoft azure|google llc|google cloud|hetzner|vultr|m247|"
    r"leaseweb|choopa|server|colo)\b", re.I)
_GPS_RE = re.compile(r"^-?\d{1,3}\.\d+,\s*-?\d{1,3}\.\d+$")


# --------------------------------------------------------------------------- #
# small helpers
# --------------------------------------------------------------------------- #
def _s(v):
    return "" if v is None else str(v).strip()


def _truthy(v):
    return _s(v).upper() in ("SÍ", "SI", "TRUE", "1", "YES", "VERDADERO", "X")


def _num(v):
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return None


def _norm_name(v):
    s = unicodedata.normalize("NFD", _s(v))
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def _table(wb, name):
    rows = wb.get(name)
    return analyze_impulso.Table(rows) if rows else None


def _fmt_dt(v):
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, datetime.date):
        return v.isoformat()
    return _s(v)


def _short(s, n=10):
    s = _s(s)
    return (s[:n] + "…") if len(s) > n else s


def _is_real_gps(s):
    return bool(_GPS_RE.match(_s(s)))


def _who(entry):
    """'Name (DNI)' for a person entry, best-effort."""
    nm = entry.get("name") or "—"
    dn = entry.get("dni")
    return "%s (%s)" % (nm, dn) if dn else nm


def _name_tok(name_norm):
    """Order-insensitive name key: 'romina vanesa belizan' == 'belizan romina vanesa'."""
    return " ".join(sorted(name_norm.split()))


def _distinct_people(entries):
    """Collapse a bucket of Dispositivos rows into DISTINCT people.

    A person is keyed by DNI when present. A name-only row (e.g. a prospect without a
    DNI) folds into a DNI person when one name's tokens are a subset of the other's —
    so 'Campos Catalina' (prospect) and 'Campos Bianca Catalina' (DNI applicant) count
    as one person, not a false shared device. Otherwise it counts on its own.
    Returns {person_key: representative_entry} (newest representative per person)."""
    people, dni_tok = {}, {}
    for e in entries:                                   # DNI people first
        if not e["dni"]:
            continue
        key = "D:" + e["dni"]
        dni_tok[key] = dni_tok.get(key, set()) | set(e["name_norm"].split())
        prev = people.get(key)
        if prev is None or _ts_key(e["ts"]) >= _ts_key(prev["ts"]):
            people[key] = e
    for e in entries:                                   # then name-only rows
        if e["dni"] or not e["name_norm"]:
            continue
        toks = set(e["name_norm"].split())
        # `dt and …` skips DNI people with no name (empty token set) — otherwise the empty
        # set is a subset of everything and would absorb every unrelated name-only prospect.
        if len(toks) >= 2 and any(dt and (toks <= dt or dt <= toks) for dt in dni_tok.values()):
            continue                                    # same person as a DNI row here
        key = "N:" + _name_tok(e["name_norm"])
        prev = people.get(key)
        if prev is None or _ts_key(e["ts"]) >= _ts_key(prev["ts"]):
            people[key] = e
    return people


def _ts_key(ts):
    if isinstance(ts, datetime.datetime):
        return ts
    if isinstance(ts, datetime.date):
        return datetime.datetime(ts.year, ts.month, ts.day)
    return datetime.datetime.min


# --------------------------------------------------------------------------- #
# core
# --------------------------------------------------------------------------- #
def build_devices(xlsx_path, clients=None, today=None):
    """Return {summary, clusters, flags, log, available} from the «Dispositivos» sheet."""
    today = today or datetime.date.today()
    wb = xlsx_mini.read_xlsx(str(xlsx_path))
    D = _table(wb, SHEET)
    if not D:
        return {"summary": {}, "clusters": [], "flags": [], "log": [], "available": False}

    # identity resolution from Clientes: dni -> cuil / name (for open-person + nicer labels)
    dni_cuil, dni_name = {}, {}
    C = _table(wb, "Clientes")
    if C:
        for r in C.records("ID Cliente"):
            dn = analyze_impulso.norm_dni(C.get(r, "DNI (norm)") or C.get(r, "DNI"))
            if not dn:
                continue
            dni_cuil.setdefault(dn, analyze_impulso.norm_cuil(C.get(r, "CUIL")))
            dni_name.setdefault(dn, _s(C.get(r, "Nombre")))

    # optional / future columns, matched loosely by header keywords (present-or-absent)
    def find_col(*kw):
        for h in D.header:
            hl = h.lower()
            if all(k in hl for k in kw):
                return h
        return None
    col_vpn = find_col("vpn") or find_col("datacenter") or find_col("proxy")
    col_token = find_col("token")
    col_skew = find_col("reloj") or find_col("skew")
    col_locale = find_col("consistencia") or find_col("locale")
    col_paste = find_col("pegado") or find_col("paste")
    col_exifgps = find_col("exif", "gps") or find_col("foto", "gps")
    col_exifdev = (find_col("exif", "equipo") or find_col("exif", "modelo")
                   or find_col("foto", "modelo") or find_col("foto", "dispositivo"))

    recs = []
    fp_idx, ip_idx, gps_idx, tok_idx = (defaultdict(list), defaultdict(list),
                                        defaultdict(list), defaultdict(list))
    for r in D.rows:
        if not any(x not in (None, "") for x in r):
            continue
        dni = analyze_impulso.norm_dni(D.get(r, "DNI"))
        name = _s(D.get(r, "Nombre")) or (dni_name.get(dni, "") if dni else "")
        cuil = dni_cuil.get(dni, "") if dni else ""
        rec = {
            "ts": D.get(r, "Fecha y hora"),
            "dni": dni, "name": name, "name_norm": _norm_name(name),
            "token": cuil or dni,                       # for open-person in the GUI
            "loan": _s(D.get(r, "ID Préstamo")),
            "fp": _s(D.get(r, "Huella de dispositivo (SHA-256)")),
            "ip": _s(D.get(r, "IP")),
            "iploc": _s(D.get(r, "Ubicación (IP)")),
            "gps": _s(D.get(r, "GPS (lat,long)")),
            "os": _s(D.get(r, "Equipo / SO")),
            "browser": _s(D.get(r, "Navegador")),
            "screen": _s(D.get(r, "Pantalla")),
            "tz": _s(D.get(r, "Zona horaria")),
            "net": _s(D.get(r, "Red")),
            "origin": _s(D.get(r, "Origen de ubicación")),
            "dwell": _num(D.get(r, "Tiempo de lectura (s)")),
            "webdriver": _truthy(D.get(r, "Automatización (webdriver)")),
            "token_val": _s(D.get(r, col_token)) if col_token else "",
            "vpn": _truthy(D.get(r, col_vpn)) if col_vpn else False,
            "skew": _num(D.get(r, col_skew)) if col_skew else None,
            "locale_bad": (bool(_s(D.get(r, col_locale))) and not _truthy(D.get(r, col_locale))) if col_locale else False,
            "paste_id": _truthy(D.get(r, col_paste)) if col_paste else False,
            "exif_gps": _s(D.get(r, col_exifgps)) if col_exifgps else "",
            "exif_dev": _s(D.get(r, col_exifdev)) if col_exifdev else "",
            "riesgo": RISK_OK, "flags": [],
        }
        recs.append(rec)
        if rec["fp"] and rec["fp"] not in ("N/D",):
            fp_idx[rec["fp"]].append(rec)
        if rec["ip"] and rec["ip"] not in ("N/D", "", "0.0.0.0"):
            ip_idx[rec["ip"]].append(rec)
        if _is_real_gps(rec["gps"]):
            gps_idx[rec["gps"]].append(rec)
        if rec["token_val"]:
            tok_idx[rec["token_val"]].append(rec)

    # ---- clusters: a key shared by ≥N distinct people ----
    def _cluster(kind, idx, minimum, high):
        out = []
        for key, entries in idx.items():
            people = _distinct_people(entries)
            n = len(people)
            if n < minimum:
                continue
            reps = sorted(people.values(), key=lambda e: _ts_key(e["ts"]), reverse=True)
            riesgo = RISK_HIGH if (high or n >= 3) else RISK_REVIEW
            last = max((_ts_key(e["ts"]) for e in entries), default=datetime.datetime.min)
            first = min((_ts_key(e["ts"]) for e in entries), default=datetime.datetime.min)
            out.append({
                "kind": kind, "key": key, "key_short": _short(key, 12),
                "n": n, "people": [{"name": e["name"], "dni": e["dni"], "token": e["token"]} for e in reps],
                "who": " · ".join(_who(e) for e in reps),
                "first": _fmt_dt(first) if first != datetime.datetime.min else "",
                "last": _fmt_dt(last) if last != datetime.datetime.min else "",
                "riesgo": riesgo,
                # tag the member rows so the log/flags inherit the cluster risk
                "_entries": entries,
            })
            for e in entries:
                e.setdefault("_clusters", []).append(kind)
                if riesgo == RISK_HIGH:
                    e["riesgo"] = RISK_HIGH
                elif e["riesgo"] != RISK_HIGH:
                    e["riesgo"] = RISK_REVIEW
        return out

    clusters = []
    clusters += _cluster("Device", fp_idx, DEVICE_MIN, high=True)
    clusters += _cluster("Token", tok_idx, TOKEN_MIN, high=True)
    clusters += _cluster("IP", ip_idx, IP_MIN, high=False)
    clusters += _cluster("GPS", gps_idx, GPS_MIN, high=False)
    clusters.sort(key=lambda c: (_RISK_ORDER[c["riesgo"]], -c["n"], c["kind"]))

    # ---- per-record flags ----
    for e in recs:
        fl = e["flags"]
        if e["webdriver"]:
            fl.append("Automation (webdriver)")
            e["riesgo"] = RISK_HIGH
        if e["vpn"] or (e["iploc"] and _HOSTING_RE.search(e["iploc"])):
            fl.append("VPN / hosting IP" if e["vpn"] else "Possible VPN/hosting ISP")
            if e["riesgo"] != RISK_HIGH:
                e["riesgo"] = RISK_REVIEW
        if "revisar" in e["iploc"].lower():
            loc = re.sub(r"^.*?revisar vs domicilio\)\s*", "", e["iploc"], flags=re.I)
            fl.append("IP location vs address mismatch — " + (loc or e["iploc"]))
            if e["riesgo"] != RISK_HIGH:
                e["riesgo"] = RISK_REVIEW
        if e["dwell"] is not None and 0 < e["dwell"] <= FAST_DWELL_S:
            fl.append("Bot-fast dwell (%ds)" % int(e["dwell"]))
            if e["riesgo"] != RISK_HIGH:
                e["riesgo"] = RISK_REVIEW
        if e["locale_bad"]:
            fl.append("Locale/timezone inconsistency")
            if e["riesgo"] != RISK_HIGH:
                e["riesgo"] = RISK_REVIEW
        if e["skew"] is not None and abs(e["skew"]) >= 120000:  # ≥2 min client/server clock gap
            fl.append("Clock skew %+ds" % int(e["skew"] / 1000))
            if e["riesgo"] != RISK_HIGH:
                e["riesgo"] = RISK_REVIEW
        if e["paste_id"]:
            fl.append("Identity fields pasted")
            if e["riesgo"] != RISK_HIGH:
                e["riesgo"] = RISK_REVIEW
        if e["exif_gps"] and "revisar" in e["exif_gps"].lower():
            fl.append("Photo GPS vs address mismatch")
            if e["riesgo"] != RISK_HIGH:
                e["riesgo"] = RISK_REVIEW
        if e["exif_dev"] and "revisar" in e["exif_dev"].lower():
            fl.append("Photo device ≠ submitting device")
            if e["riesgo"] != RISK_HIGH:
                e["riesgo"] = RISK_REVIEW
        for kind in e.get("_clusters", []):
            fl.append("Shared %s" % kind.lower())

    flags = [{
        "name": e["name"] or "—", "dni": e["dni"] or "—", "token": e["token"],
        "when": _fmt_dt(e["ts"]), "riesgo": e["riesgo"], "loan": e["loan"],
        "flags": e["flags"], "flagtext": "; ".join(e["flags"]),
    } for e in recs if e["flags"]]
    flags.sort(key=lambda f: (_RISK_ORDER.get(f["riesgo"], 3), f["name"].lower()))

    # ---- normalized log (newest first) ----
    for e in recs:
        e["when"] = _fmt_dt(e["ts"])
    log = sorted(recs, key=lambda e: _ts_key(e["ts"]), reverse=True)

    # ---- summary ----
    fps = {e["fp"] for e in recs if e["fp"] and e["fp"] != "N/D"}
    summary = {
        "records": len(recs),
        "unique_fp": len(fps),
        "shared_fp": sum(1 for c in clusters if c["kind"] == "Device"),
        "shared_ip": sum(1 for c in clusters if c["kind"] == "IP"),
        "shared_gps": sum(1 for c in clusters if c["kind"] == "GPS"),
        "shared_token": sum(1 for c in clusters if c["kind"] == "Token"),
        "webdriver": sum(1 for e in recs if e["webdriver"]),
        "vpn": sum(1 for e in recs if e["vpn"] or (e["iploc"] and _HOSTING_RE.search(e["iploc"]))),
        "geo_mismatch": sum(1 for e in recs if "revisar" in e["iploc"].lower()),
        "fast_dwell": sum(1 for e in recs if e["dwell"] is not None and 0 < e["dwell"] <= FAST_DWELL_S),
        "flagged": len(flags),
    }
    return {"summary": summary, "clusters": clusters, "flags": flags, "log": log, "available": True}


# --------------------------------------------------------------------------- #
# rendering (CLI / export)
# --------------------------------------------------------------------------- #
_DISCLAIMER = ("_**Indicative** device leads for manual review, not verdicts. A shared IP "
               "can be a household/NAT; GPS repeats at low precision; one person may submit "
               "several times._")


def render_md(dev):
    gen = datetime.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M")
    if not dev or not dev.get("available"):
        return "# Impulso Crédito — Device insights\n\n_No «Dispositivos» sheet in the workbook._"
    s = dev["summary"]
    L = ["# Impulso Crédito — Device & Fraud Insights (Dispositivos)", "",
         "_Generated %s · %d submission(s) · %d unique device(s)._" % (gen, s["records"], s["unique_fp"]),
         "", _DISCLAIMER, "",
         "## Summary", "",
         "| Records | Unique devices | Shared devices | Shared IPs | Shared tokens | Webdriver | VPN/hosting | Geo-mismatch | Bot-fast |",
         "|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
         "| %d | %d | %d | %d | %d | %d | %d | %d | %d |" % (
             s["records"], s["unique_fp"], s["shared_fp"], s["shared_ip"], s["shared_token"],
             s["webdriver"], s["vpn"], s["geo_mismatch"], s["fast_dwell"]),
         "", "## Clusters (same key, several people)", "",
         "| Risk | Type | Key | People | Who | Last seen |",
         "|---|---|---|---:|---|---|"]
    for c in dev["clusters"]:
        L.append("| %s | %s | %s | %d | %s | %s |" % (
            c["riesgo"], c["kind"], c["key_short"], c["n"], c["who"], c["last"]))
    if not dev["clusters"]:
        L.append("| — | — | _none_ | 0 | | |")
    L += ["", "## Flagged submissions", "",
          "| Risk | Client | DNI | When | Flags |", "|---|---|---|---|---|"]
    for f in dev["flags"]:
        L.append("| %s | %s | %s | %s | %s |" % (f["riesgo"], f["name"], f["dni"], f["when"], f["flagtext"]))
    if not dev["flags"]:
        L.append("| — | _none_ | | | |")
    return "\n".join(L)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def parse_args(argv):
    p = argparse.ArgumentParser(
        description="Mine the «Dispositivos» sheet for device/fraud insights.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("--xlsx", default=str(DEFAULT_XLSX), help="Impulso workbook (.xlsx export)")
    p.add_argument("--export", nargs="?", const=str(DEFAULT_EXPORT), help="also write the report to this .md")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    dev = build_devices(args.xlsx)
    text = render_md(dev)
    print(text)
    if args.export:
        Path(args.export).parent.mkdir(parents=True, exist_ok=True)
        Path(args.export).write_text(text + "\n", encoding="utf-8")
        print("\nSaved %s" % args.export, file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
