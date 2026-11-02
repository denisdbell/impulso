#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fraud signals — mines the Impulso workbook (``out/_impulso.xlsx``) for indications
of suspicious identity/behavior, WITHOUT OCR and with no new dependencies (it only
reads the .xlsx with xlsx_mini, like the rest of bcra).

It cross-references records across sheets (Clientes, Dispositivos, Firmas,
Prestatarios, Pagos, Nuevos Prestatarios) and flags, per client, signals such as:
  • Duplicate identity: shared DNI/CUIL/phone/email/address, inconsistent CUIL↔DNI,
    invalid email, flagged as duplicate, blocked with an active loan.
  • Device/network: same device fingerprint, same IP (or signing IP) across several
    clients, shared GPS, automation (webdriver).
  • Reference ring: a reference phone that equals a client phone, or is reused.
  • Third-party account / mule: disbursement to a third party; one account holder
    receiving for several borrowers.
  • Consistency: a loan's or payment's Name/DNI differing from the linked client.

It can optionally fold in the DNI-OCR anomalies (verify_dni) as one more category.
Results are **indicative**: leads to review manually, not verdicts.

Usage:
    python3 bcra/analyze_fraud.py [--xlsx out/_impulso.xlsx] [--export out/_fraude.md]

Note: the Clientes/Dispositivos/… sheet and column names, and the stored status
values (PAGADO, SÍ/NO, …), are intentionally kept in Spanish — they are the actual
labels in the workbook.
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
DEFAULT_EXPORT = HERE / "out" / "_fraude.md"

# Signal weights and the risk thresholds built from them.
HIGH, MED, LOW = 40, 20, 10
IP_SHARE_MIN = 3   # an IP shared by this many distinct clients is a signal
GPS_SHARE_MIN = 3  # same for an identical GPS fix

# Risk levels (shown to the user).
RISK_HIGH, RISK_REVIEW, RISK_OK = "High", "Review", "OK"


# --------------------------------------------------------------------------- #
# small helpers
# --------------------------------------------------------------------------- #
def _s(v):
    return "" if v is None else str(v).strip()


def _truthy(v):
    return _s(v).upper() in ("SÍ", "SI", "TRUE", "1", "YES", "VERDADERO", "X")


def _phone(v):
    d = re.sub(r"\D", "", _s(v))
    return d[-10:] if len(d) >= 8 else ""


def _email(v):
    return _s(v).lower()


def _addr(v):
    a = re.sub(r"\s+", " ", _s(v)).lower()
    return a if len(a) >= 8 else ""


def _norm_name(v):
    s = unicodedata.normalize("NFD", _s(v))
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def _cuil_check_ok(cuil):
    """AFIP mod-11 check digit (returns True when we can't judge)."""
    if len(cuil) != 11:
        return True
    w = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    dv = 11 - (sum(int(cuil[i]) * w[i] for i in range(10)) % 11)
    if dv == 11:
        dv = 0
    elif dv == 10:
        dv = 9
    return dv == int(cuil[10])


def _table(wb, name):
    rows = wb.get(name)
    return analyze_impulso.Table(rows) if rows else None


def _own_account_yes(v):
    """¿Cuenta propia? — blank defaults to 'yes' (own account); only an explicit
    NO/FALSE means a third-party account."""
    s = _s(v).upper()
    if not s:
        return True
    return s not in ("NO", "FALSE", "0", "N")


# --------------------------------------------------------------------------- #
# core
# --------------------------------------------------------------------------- #
def build_fraud(xlsx_path, dni_records=None, today=None):
    """Return a list of per-client fraud records, worst first."""
    today = today or datetime.date.today()
    wb = xlsx_mini.read_xlsx(str(xlsx_path))
    C = _table(wb, "Clientes")
    if not C:
        raise SystemExit("Workbook %s has no 'Clientes' sheet." % xlsx_path)

    # ---- 1) read Clientes into per-client dicts ----
    clients = {}        # cid -> dict
    dni_to_cid = {}
    for r in C.records("ID Cliente"):
        cid = _s(C.get(r, "ID Cliente"))
        dni = analyze_impulso.norm_dni(C.get(r, "DNI (norm)") or C.get(r, "DNI"))
        c = {
            "client_id": cid,
            "name": _s(C.get(r, "Nombre")),
            "dni": dni,
            "cuil": analyze_impulso.norm_cuil(C.get(r, "CUIL")),
            "email": _email(C.get(r, "Correo (norm)") or C.get(r, "Correo")),
            "phone": _phone(C.get(r, "Teléfono (norm)") or C.get(r, "Teléfono")),
            "addr": _addr(C.get(r, "Dirección")),
            "correo_valido_raw": C.get(r, "Correo válido"),
            "duplicado": _truthy(C.get(r, "Duplicado")),
            "bloqueado": _truthy(C.get(r, "Bloqueado")),
            "ref_phones": [p for p in (_phone(C.get(r, "Ref 1 Teléfono")), _phone(C.get(r, "Ref 2 Teléfono"))) if p],
            "signals": [],
        }
        clients[cid] = c
        if dni and dni not in dni_to_cid:
            dni_to_cid[dni] = cid

    # ---- 2) cross-client indexes (within Clientes) ----
    def index(attr):
        idx = defaultdict(set)
        for cid, c in clients.items():
            v = c[attr]
            if v:
                idx[v].add(cid)
        return idx

    by_dni = index("dni")
    by_cuil = index("cuil")
    by_email = index("email")
    by_phone = index("phone")
    by_addr = index("addr")
    own_phones = {c["phone"] for c in clients.values() if c["phone"]}
    ref_owner = defaultdict(set)
    for cid, c in clients.items():
        for rp in c["ref_phones"]:
            ref_owner[rp].add(cid)

    # ---- 3) loans: active flag, loan→client, denormalized name/DNI ----
    P = _table(wb, "Prestatarios")
    active_cid, loan_cid = set(), {}
    titular_owner = defaultdict(set)
    inconsistent_loan = set()
    thirdparty_cid = set()
    if P:
        for r in P.records("ID Préstamo"):
            lid = _s(P.get(r, "ID Préstamo"))
            cid = _s(P.get(r, "ID Cliente"))
            loan_cid[lid] = cid
            estado = _s(P.get(r, "Estado")).upper()
            if estado and estado not in ("PAGADO", "SALDADO", "RECHAZADO", "ANULADO"):
                active_cid.add(cid)
            if cid in clients:
                ld = analyze_impulso.norm_dni(P.get(r, "DNI"))
                ln = _norm_name(P.get(r, "Nombre"))
                cc = clients[cid]
                if (ld and cc["dni"] and ld != cc["dni"]) or (ln and _norm_name(cc["name"]) and ln != _norm_name(cc["name"])):
                    inconsistent_loan.add(cid)
            if not _own_account_yes(P.get(r, "¿Cuenta propia?")):
                if cid in clients:
                    thirdparty_cid.add(cid)
            tdni = analyze_impulso.norm_dni(P.get(r, "DNI del titular"))
            if tdni and cid:
                titular_owner[tdni].add(cid)

    # Nuevos Prestatarios: third-party + account holder (prospects keyed by DNI)
    NP = _table(wb, "Nuevos Prestatarios")
    if NP:
        for r in NP.records("DNI"):
            dni = analyze_impulso.norm_dni(NP.get(r, "DNI"))
            cid = dni_to_cid.get(dni)
            if cid and not _own_account_yes(NP.get(r, "¿Cuenta propia?")):
                thirdparty_cid.add(cid)
            tdni = analyze_impulso.norm_dni(NP.get(r, "DNI del titular"))
            if tdni and cid:
                titular_owner[tdni].add(cid)

    # Pagos: denormalized name/DNI vs the loan's client
    Pay = _table(wb, "Pagos")
    inconsistent_pay = set()
    if Pay:
        for r in Pay.records("ID Préstamo"):
            cid = loan_cid.get(_s(Pay.get(r, "ID Préstamo")))
            if cid in clients:
                pd = analyze_impulso.norm_dni(Pay.get(r, "DNI"))
                if pd and clients[cid]["dni"] and pd != clients[cid]["dni"]:
                    inconsistent_pay.add(cid)

    # ---- 4) devices / network ----
    def _dev_key(get, r):
        cid = _s(get(r, "ID Cliente"))
        if cid in clients:
            return cid
        return dni_to_cid.get(analyze_impulso.norm_dni(get(r, "DNI")), "")

    fp_idx, ip_idx, gps_idx = defaultdict(set), defaultdict(set), defaultdict(set)
    webdriver_cid = set()
    D = _table(wb, "Dispositivos")
    if D:
        for r in D.rows:
            cid = _dev_key(D.get, r)
            if not cid:
                continue
            fp = _s(D.get(r, "Huella de dispositivo (SHA-256)"))
            ip = _s(D.get(r, "IP"))
            gps = _s(D.get(r, "GPS (lat,long)"))
            if fp:
                fp_idx[fp].add(cid)
            if ip:
                ip_idx[ip].add(cid)
            if gps:
                gps_idx[gps].add(cid)
            if _truthy(D.get(r, "Automatización (webdriver)")):
                webdriver_cid.add(cid)

    # Firmas: signing IP shared across distinct DNIs
    F = _table(wb, "Firmas")
    sign_ip = defaultdict(set)
    if F:
        for r in F.records("ID Préstamo"):
            ip = _s(F.get(r, "IP"))
            dni = analyze_impulso.norm_dni(F.get(r, "DNI"))
            if ip and dni:
                sign_ip[ip].add(dni)
    shared_sign_dnis = {dni for ip, dnis in sign_ip.items() if len(dnis) >= 2 for dni in dnis}

    def _dev_shares(idx, cid, minimum=2):
        """Max distinct-client count among this client's buckets in idx."""
        best = 0
        for bucket in idx.values():
            if cid in bucket and len(bucket) >= minimum:
                best = max(best, len(bucket))
        return best

    # ---- 5) per-client signal assembly ----
    def add(c, code, label, weight, detail=""):
        c["signals"].append({"code": code, "label": label, "weight": weight, "detail": detail})

    for cid, c in clients.items():
        # duplicate identity
        if c["dni"] and len(by_dni[c["dni"]]) > 1:
            add(c, "dni_dup", "DNI repeated across clients", HIGH, "%d clients" % len(by_dni[c["dni"]]))
        if c["cuil"] and len(by_cuil[c["cuil"]]) > 1:
            add(c, "cuil_dup", "CUIL repeated across clients", HIGH, "%d clients" % len(by_cuil[c["cuil"]]))
        if c["cuil"]:
            embedded = c["cuil"][2:10].lstrip("0")
            if (c["dni"] and embedded and embedded != c["dni"].lstrip("0")) or not _cuil_check_ok(c["cuil"]):
                add(c, "cuil_dni", "CUIL↔DNI inconsistent", MED)
        if c["phone"] and len(by_phone[c["phone"]]) > 1:
            add(c, "phone_dup", "Shared phone", MED, "%d clients" % len(by_phone[c["phone"]]))
        if c["email"] and len(by_email[c["email"]]) > 1:
            add(c, "email_dup", "Shared email", MED, "%d clients" % len(by_email[c["email"]]))
        if c["correo_valido_raw"] is not None and _s(c["correo_valido_raw"]) and not _truthy(c["correo_valido_raw"]):
            add(c, "email_bad", "Invalid email", LOW)
        if c["addr"] and len(by_addr[c["addr"]]) > 1:
            add(c, "addr_dup", "Shared address", LOW, "%d clients" % len(by_addr[c["addr"]]))
        if c["duplicado"]:
            add(c, "flag_dup", "Flagged as duplicate", MED)
        if c["bloqueado"] and cid in active_cid:
            add(c, "blocked", "Blocked with active loan", HIGH)

        # device / network
        fp_n = _dev_shares(fp_idx, cid)
        if fp_n:
            add(c, "device", "Shared device fingerprint", HIGH, "%d clients" % fp_n)
        ip_n = _dev_shares(ip_idx, cid, IP_SHARE_MIN)
        if ip_n:
            add(c, "ip", "Shared IP", MED, "%d clients" % ip_n)
        gps_n = _dev_shares(gps_idx, cid, GPS_SHARE_MIN)
        if gps_n:
            add(c, "gps", "Shared GPS", LOW, "%d clients" % gps_n)
        if cid in webdriver_cid:
            add(c, "webdriver", "Automation (webdriver)", HIGH)
        if c["dni"] and c["dni"] in shared_sign_dnis:
            add(c, "sign_ip", "Signing IP shared across DNIs", MED)

        # reference ring
        if any(rp in own_phones for rp in c["ref_phones"]):
            add(c, "ref_self", "Reference = a client phone", MED)
        if any(len(ref_owner[rp]) > 1 for rp in c["ref_phones"]):
            add(c, "ref_dup", "Shared reference phone", LOW)

        # third-party account / mule
        if cid in thirdparty_cid:
            add(c, "thirdparty", "Disbursement to a third-party account", LOW)

        # consistency
        if cid in inconsistent_loan:
            add(c, "loan_mismatch", "Loan Name/DNI differs from client", MED)
        if cid in inconsistent_pay:
            add(c, "pay_mismatch", "Payment Name/DNI differs from client", LOW)

    # account holder shared across >1 borrower → flag those borrowers
    for tdni, owners in titular_owner.items():
        if len(owners) > 1:
            for cid in owners:
                if cid in clients:
                    add(clients[cid], "titular_shared", "Account holder receives for multiple borrowers",
                        HIGH, "%d borrowers" % len(owners))

    # ---- 6) fold in DNI-OCR anomalies (one more category) ----
    if dni_records:
        amap = {
            "DNI mismatch": ("dni_photo", "DNI photo mismatch", HIGH),
            "Name mismatch": ("name_photo", "Name photo mismatch", MED),
            "Front and back mismatch": ("front_back", "DNI front/back mismatch", MED),
            "DNI shared across people": ("doc_shared", "Same DNI on multiple people (photo)", HIGH),
            "Document expired": ("doc_expired", "Document expired", LOW),
        }
        for rec in dni_records:
            cid = rec.get("cliente_id")
            if cid not in clients:
                continue
            for lbl in rec.get("anomalias", []):
                if lbl in amap:
                    code, text, w = amap[lbl]
                    add(clients[cid], code, text, w, "photo")

    # ---- 7) score + risk ----
    records = []
    for c in clients.values():
        score = sum(s["weight"] for s in c["signals"])
        has_high = any(s["weight"] >= HIGH for s in c["signals"])
        riesgo = RISK_HIGH if (has_high or score >= HIGH) else (RISK_REVIEW if score > 0 else RISK_OK)
        records.append({
            "client_id": c["client_id"], "name": c["name"], "dni": c["dni"], "cuil": c["cuil"],
            "riesgo": riesgo, "score": score, "n_signals": len(c["signals"]), "signals": c["signals"],
        })
    order = {RISK_HIGH: 0, RISK_REVIEW: 1, RISK_OK: 2}
    records.sort(key=lambda r: (order.get(r["riesgo"], 3), -r["score"], r["name"].lower()))
    return records


# --------------------------------------------------------------------------- #
# rendering
# --------------------------------------------------------------------------- #
_DISCLAIMER = ("_**Indicative** signals for manual review, not verdicts. Several have "
               "legitimate explanations (cohabitants, same neighborhood, low GPS precision)._")


def render_md(records):
    gen = datetime.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M")
    flagged = [r for r in records if r["riesgo"] != RISK_OK]
    high = sum(1 for r in records if r["riesgo"] == RISK_HIGH)
    L = ["# Impulso Crédito — Fraud Signals", "",
         "_Generated %s · %d client(s), %d with signals (%d high risk)._"
         % (gen, len(records), len(flagged), high), "",
         _DISCLAIMER, "",
         "| Risk | Client | DNI | CUIL | # | Signals |",
         "|---|---|---|---|---:|---|"]
    for r in flagged:
        sigs = ", ".join("%s%s" % (s["label"], (" (%s)" % s["detail"]) if s["detail"] else "")
                         for s in r["signals"])
        L.append("| %s | %s | %s | %s | %d | %s |" % (
            r["riesgo"], r["name"] or "—", r["dni"] or "—", r["cuil"] or "—", r["n_signals"], sigs))
    if not flagged:
        L.append("| — | _no signals_ | | | 0 | |")
    return "\n".join(L)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def parse_args(argv):
    p = argparse.ArgumentParser(
        description="Mine the Impulso workbook for cross-record fraud signals.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("--xlsx", default=str(DEFAULT_XLSX), help="Impulso workbook (.xlsx export)")
    p.add_argument("--export", nargs="?", const=str(DEFAULT_EXPORT), help="also write the report to this .md")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    records = build_fraud(args.xlsx)
    text = render_md(records)
    print(text)
    if args.export:
        Path(args.export).parent.mkdir(parents=True, exist_ok=True)
        Path(args.export).write_text(text + "\n", encoding="utf-8")
        print("\nSaved %s" % args.export, file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
