#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Anti-fraud RULES mirror — evaluates the new Apps Script validations against the
workbook, so the dashboard can show, per person, exactly what production would say:

  V-38  Max ONE of the three references may be an existing client (phone match).
  V-39  Configurable max of distinct persons sharing IP / device fingerprint /
        persistent token / GPS / declared address — each connected person shown
        WITH the status of their loans.
  V-40  Phone-connection map (informative): references who are clients, clients
        sharing the person's phone or listing them as a reference, references
        shared across borrowers, and matches against the «Prospectos» sheet.

Pure stdlib (xlsx via xlsx_mini), read-only, faithful port of Validaciones.gs /
Dispositivos.gs (same normalization, same insertion-order person lists, same
loan-status summaries, config read from the workbook's «Configuración» sheet).
All OUTPUT text is English (production messages are Spanish): sheet statuses are
translated for display only; matching always runs on the raw Spanish values.
V-39 here answers: “if this person applied again TODAY, what would happen?” —
the same question approveApplicantV2_ asks (last Dispositivos row + address).

Usage:
    python3 bcra/analyze_rules.py --all            # report → out/_rules.md
    python3 bcra/analyze_rules.py <cuil|dni|C-x>   # one person
Indicative mirror for review — the blocking itself happens in Apps Script.
"""

import argparse
import datetime
import logging
import re
import sys
import unicodedata
from collections import defaultdict
from math import asin, cos, radians, sin, sqrt
from pathlib import Path

import xlsx_mini
import analyze_impulso  # norm_dni/norm_cuil/money/as_date patterns (same folder)

HERE = Path(__file__).resolve().parent
DEFAULT_XLSX = HERE / "out" / "_impulso.xlsx"
DEFAULT_EXPORT = HERE / "out" / "_rules.md"

VIGENTE_EXCL = ("PAGADO", "SALDADO")      # blacklist, like the .gs (ST.PAID/ST.CLEARED)
GRACE_DEFAULTS = {"on": True, "max": 2, "radioM": 50.0, "ipDias": 30}

# All generated text is ENGLISH (the production messages are Spanish): the
# dashboard is English, so sheet statuses are translated for display only —
# matching/logic always runs on the raw Spanish values.
ETIQ = {"ip": "the same IP", "fp": "the same device (fingerprint)",
        "token": "the same device (token)", "gps": "the same GPS location",
        "addr": "the same declared address"}

LOAN_STATE_EN = {"ACTIVO": "ACTIVE", "VENCIDO": "OVERDUE", "CONGELADO": "FROZEN",
                 "PAGADO": "PAID", "SALDADO": "SETTLED"}
PROSPECT_STATE_EN = {"NUEVO": "NEW", "CONTACTADO": "CONTACTED",
                     "APROBADO": "APPROVED", "RECHAZADO": "REJECTED"}


# --------------------------------------------------------------------------- #
# value helpers (mirror Validaciones.gs semantics)
# --------------------------------------------------------------------------- #
def _cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def norm_dni(v):
    d = re.sub(r"\D", "", _cell(v))
    return d if 7 <= len(d) <= 8 else ""


def norm_phone(v):
    d = re.sub(r"\D", "", _cell(v))
    if not d:
        return ""
    if len(d) > 10 and d.startswith("54"):
        d = d[2:]
    if len(d) == 11 and d[0] == "9":
        d = d[1:]
    if len(d) != 10 or d[0] not in "123" or re.fullmatch(r"(\d)\1{9}", d):
        return ""
    return "+54" + d


def norm_address(v):
    s = unicodedata.normalize("NFD", _cell(v))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^a-z0-9ñ]+", " ", s.lower())
    s = re.sub(r"\s+", " ", s).strip()
    return s if len(s) >= 8 else ""


def parse_gps(s):
    m = re.fullmatch(r"\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*", _cell(s))
    return (float(m.group(1)), float(m.group(2))) if m else None


def haversine_m(a, b):
    R = 6371000.0
    dlat, dlon = radians(b[0] - a[0]), radians(b[1] - a[1])
    s = sin(dlat / 2) ** 2 + cos(radians(a[0])) * cos(radians(b[0])) * sin(dlon / 2) ** 2
    return 2 * R * asin(sqrt(s))


def _lastwins_idx(rows):
    """Header → index, LAST occurrence wins (like clientesHeaderMap_ — the live
    Clientes sheet has a duplicate 'Ref 2 Teléfono' whose data lives in the last)."""
    idx = {}
    for i, h in enumerate(rows[0] if rows else []):
        h = _cell(h)
        if h:
            idx[h] = i
    return idx


# --------------------------------------------------------------------------- #
# workbook model
# --------------------------------------------------------------------------- #
class RulesData:
    """One read of the workbook → everything the three rules need."""

    def __init__(self, xlsx_path, today=None):
        self.today = today or datetime.date.today()
        self.now = datetime.datetime.now()
        wb = xlsx_mini.read_xlsx(str(xlsx_path))

        # ---- Configuración (V-39 knobs, same keys production reads) ----
        self.cfg = dict(GRACE_DEFAULTS)
        for r in wb.get("Configuración", [])[1:]:
            k = _cell(r[0] if len(r) > 0 else "")
            v = _cell(r[1] if len(r) > 1 else "")
            if k == "Límite por dispositivo/ubicación (SÍ/NO)":
                self.cfg["on"] = v.upper() in ("SÍ", "SI")
            elif k == "Máximo de personas por dispositivo/ubicación":
                try:
                    self.cfg["max"] = max(1, int(round(float(re.sub(r"[^0-9.]", "", v) or 0))) or 2)
                except ValueError:
                    pass
            elif k == "Radio GPS de coincidencia (metros)":
                try:
                    self.cfg["radioM"] = float(re.sub(r"[^0-9.]", "", v) or 50)
                except ValueError:
                    pass
            elif k == "Ventana de coincidencia de IP (días)":
                try:
                    self.cfg["ipDias"] = int(float(re.sub(r"[^0-9.]", "", v) or 30))
                except ValueError:
                    pass

        # ---- Clientes (last-wins headers; re-normalized phones like the fixed .gs) ----
        crows = wb.get("Clientes", [])
        ci = _lastwins_idx(crows)
        self.clientes = []
        for r in crows[1:]:
            g = lambda k: _cell(r[ci[k]]) if k in ci and ci[k] < len(r) else ""
            cid, dni = g("ID Cliente"), norm_dni(g("DNI (norm)") or g("DNI"))
            if not cid and not dni:
                continue
            self.clientes.append({
                "id": cid, "name": g("Nombre"), "dniNorm": dni,
                "phoneNorm": norm_phone(g("Teléfono (norm)")) or norm_phone(g("Teléfono")),
                "address": g("Dirección"),
                "refNames": [g("Ref 1 Nombre"), g("Ref 2 Nombre"), g("Ref 3 Nombre")],
                "refPhones": [norm_phone(g("Ref 1 Teléfono")), norm_phone(g("Ref 2 Teléfono")),
                              norm_phone(g("Ref 3 Teléfono"))],
            })
        for c in self.clientes:
            c["refsN"] = [p for p in c["refPhones"] if p]
        self.by_dni = {c["dniNorm"]: c for c in self.clientes if c["dniNorm"]}

        # ---- Prestatarios → loans per client (status + saldo) ----
        prows = wb.get("Prestatarios", [])
        pidx = _lastwins_idx(prows)
        self.loans = defaultdict(list)
        for r in prows[1:]:
            g = lambda k: r[pidx[k]] if k in pidx and pidx[k] < len(r) else None
            lid, cid = _cell(g("ID Préstamo")), _cell(g("ID Cliente"))
            if not lid or not cid or len(lid) > 20:      # skip the workbook's note row
                continue
            self.loans[cid].append({
                "loanId": lid, "estado": _cell(g("Estado")).upper(),
                "saldo": float(g("Saldo Pendiente") or 0) if isinstance(g("Saldo Pendiente"), (int, float)) else 0.0,
            })

        # ---- Prospectos ----
        rrows = wb.get("Prospectos", [])
        ridx = _lastwins_idx(rrows)
        self.prospectos = []
        for rix, r in enumerate(rrows[1:], start=2):
            g = lambda k: _cell(r[ridx[k]]) if k in ridx and ridx[k] < len(r) else ""
            name, cuil = g("Nombre"), re.sub(r"\D", "", g("CUIL"))
            if not name and not cuil:
                continue
            fv = r[ridx["Fecha de Captura"]] if "Fecha de Captura" in ridx and ridx["Fecha de Captura"] < len(r) else None
            self.prospectos.append({
                "name": name, "cuil": cuil, "row": rix,
                "dni": norm_dni(cuil[2:10]) if len(cuil) == 11 else "",
                "phoneN": norm_phone(g("Teléfono")), "estado": g("Estado").upper() or "NUEVO",
                "fecha": fv.strftime("%Y-%m-%d") if isinstance(fv, datetime.datetime) else "",
            })

        # ---- Dispositivos (positional base cols + Token by header) ----
        drows = wb.get("Dispositivos", [])
        didx = _lastwins_idx(drows)
        tok_i = didx.get("Token persistente")
        self.disp = []
        for r in drows[1:]:
            at = lambda i: r[i] if i is not None and i < len(r) else None
            t = at(0)
            ip = _cell(at(21))
            row = {"t": t if isinstance(t, datetime.datetime) else None,
                   "dniNorm": norm_dni(at(3)), "clientId": _cell(at(2)),
                   "ip": ip if ip and ip != "N/D" else "", "fp": _cell(at(6)),
                   "token": _cell(at(tok_i)) if tok_i is not None else "",
                   "gps": parse_gps(at(23)),
                   "gpsAcc": float(at(24)) if isinstance(at(24), (int, float)) else 0.0}
            if row["dniNorm"] or row["clientId"]:
                self.disp.append(row)
        self.id_to_dni = {c["id"]: c["dniNorm"] for c in self.clientes if c["id"] and c["dniNorm"]}

    # ---------------- shared formatting ---------------- #
    def resumen_prestamos(self, cid):
        loans = self.loans.get(cid, [])
        if not loans:
            return "no loans"
        vig = [l for l in loans if l["estado"] not in VIGENTE_EXCL]
        if not vig:
            return "no open loans (%d settled)" % len(loans)
        return " · ".join("%s %s %s" % (l["loanId"], LOAN_STATE_EN.get(l["estado"], l["estado"]),
                                        analyze_impulso.money(l["saldo"])) for l in vig)

    def quien_detalle(self, k):
        c = self.by_dni.get(k) if re.fullmatch(r"\d{7,8}", k or "") else None
        if c and c["id"]:
            return "DNI %s (%s — %s)" % (k, c["id"], self.resumen_prestamos(c["id"]))
        if re.fullmatch(r"\d{7,8}", k or ""):
            return "DNI " + k
        return k

    # ---------------- V-39 ---------------- #
    def _canon(self, r):
        return r["dniNorm"] or self.id_to_dni.get(r["clientId"], "") or r["clientId"]

    def _contar(self, sig):
        out = {"ip": {}, "fp": {}, "token": {}, "gps": {}}
        ip_lim = datetime.timedelta(days=max(1, self.cfg["ipDias"]))
        for r in self.disp:
            own = (sig["dniNorm"] and r["dniNorm"] == sig["dniNorm"]) or \
                  (sig["clientId"] and r["clientId"] and r["clientId"] == sig["clientId"])
            if own:
                continue
            k = self._canon(r)
            if not k:
                continue
            if sig["ip"] and r["ip"] == sig["ip"] and r["t"] and (self.now - r["t"]) <= ip_lim:
                out["ip"][k] = 1
            if sig["deviceFp"] and r["fp"] == sig["deviceFp"]:
                out["fp"][k] = 1
            if sig["token"] and r["token"] == sig["token"]:
                out["token"][k] = 1
            if sig["gps"] and r["gps"] and haversine_m(sig["gps"], r["gps"]) <= max(self.cfg["radioM"], r["gpsAcc"]):
                out["gps"][k] = 1
        return out

    def check_v39(self, sig):
        if not self.cfg["on"]:
            return {"ok": True, "exceeded": None, "alertas": []}
        m = self._contar({"dniNorm": sig.get("dniNorm", ""), "clientId": sig.get("clientId", ""),
                          "ip": sig.get("ip", ""), "deviceFp": sig.get("deviceFp", ""),
                          "token": sig.get("token", ""), "gps": parse_gps(sig.get("gps", ""))})
        senales = [("ip", sig.get("ip", ""), m["ip"]), ("fp", sig.get("deviceFp", ""), m["fp"]),
                   ("token", sig.get("token", ""), m["token"]), ("gps", sig.get("gps", ""), m["gps"])]
        addr_n = norm_address(sig.get("address", ""))
        if addr_n:
            otros = {}
            for o in self.clientes:
                if o["dniNorm"] and o["dniNorm"] != sig.get("dniNorm") and norm_address(o["address"]) == addr_n:
                    otros[o["dniNorm"]] = 1
            senales.append(("addr", sig.get("address", ""), otros))
        exceeded, alertas = [], []
        for key, valor, otros in senales:
            if not valor or not otros:
                continue
            arr = list(otros)
            if 1 + len(arr) > self.cfg["max"]:
                exceeded.append({"senal": key, "etiqueta": ETIQ[key], "valor": valor,
                                 "otros": arr, "max": self.cfg["max"]})
            else:
                alertas.append("Shares %s with %d other person(s): %s"
                               % (ETIQ[key], len(arr), ", ".join(self.quien_detalle(k) for k in arr)))
        return {"ok": not exceeded, "exceeded": exceeded or None, "alertas": alertas}

    def exceeded_items(self, exceeded):
        """One sentence per exceeded V-39 signal — rendered as bullets in the GUI."""
        return ["Shares %s «%s» with %d other person(s): %s — max %d"
                % (e["etiqueta"], e["valor"], len(e["otros"]),
                   ", ".join(self.quien_detalle(k) for k in e["otros"]), e["max"])
                for e in (exceeded or [])]

    def msg_prestamista(self, exceeded, sin_cola=False):
        return "V-39 — " + " · ".join(self.exceeded_items(exceeded)) + \
               ("." if sin_cola else ". In the sheet, «Anular límites» overrides this at approval.")

    def ultima_fila(self, dni, cid=""):
        for r in reversed(self.disp):
            if (dni and r["dniNorm"] == dni) or (cid and r["clientId"] == cid):
                return r
        return None

    def v39_for_person(self, dni, cid, address):
        """Same question as approveApplicantV2_: last device row + declared address."""
        own = self.ultima_fila(dni, cid)
        return self.check_v39({
            "dniNorm": dni, "clientId": cid,
            "ip": own["ip"] if own else "", "deviceFp": own["fp"] if own else "",
            "token": own["token"] if own else "",
            "gps": ("%.5f,%.5f" % own["gps"]) if own and own["gps"] else "",
            "address": address}), own

    # ---------------- V-38 / V-40 ---------------- #
    def ref_cliente_match(self, phone_n, self_dni):
        if not phone_n:
            return None
        return next((c for c in self.clientes if c["phoneNorm"] == phone_n and c["dniNorm"] != self_dni), None)

    def v38_for_client(self, c):
        ms = [self.ref_cliente_match(p, c["dniNorm"]) for p in c["refPhones"]]
        det = ["Ref %d (%s) → client %s (%s) — %s"
               % (i + 1, c["refNames"][i] or "—", m["id"], m["name"], self.resumen_prestamos(m["id"]))
               for i, m in enumerate(ms) if m]
        return sum(1 for m in ms if m), det

    def conexiones(self, app_phone, refs, self_dni, skip_prospect_row=None):
        """V-40 connection items. `skip_prospect_row` excludes the person's own
        «Prospectos» row when scoring a prospect (row number, 1-based)."""
        items, seen = [], set()

        def add(k, t):
            if k not in seen:
                seen.add(k)
                items.append(t)

        for c in self.clientes:
            if app_phone and c["phoneNorm"] == app_phone and c["dniNorm"] and c["dniNorm"] != self_dni:
                add("tel-cli:" + c["id"], "Same phone as client %s (%s) — %s"
                    % (c["id"], c["name"], self.resumen_prestamos(c["id"])))
            if app_phone and c["dniNorm"] != self_dni and app_phone in c["refsN"]:
                add("es-ref-de:" + c["id"], "Listed as a reference by %s (%s) — %s"
                    % (c["id"], c["name"], self.resumen_prestamos(c["id"])))
        for slot, rname, ph in refs:
            if not ph:
                continue
            cli = self.ref_cliente_match(ph, self_dni)
            if cli:
                add("ref-cli:%s:%s" % (slot, cli["id"]), "%s (%s) is client %s (%s) — %s"
                    % (slot, rname, cli["id"], cli["name"], self.resumen_prestamos(cli["id"])))
            for c in self.clientes:
                if c["dniNorm"] == self_dni:
                    continue
                if ph in c["refsN"]:
                    add("ref-comp:%s:%s" % (slot, c["id"]),
                        "%s shares a phone with a reference of %s (%s) — %s"
                        % (slot, c["id"], c["name"], self.resumen_prestamos(c["id"])))
        for p in self.prospectos:
            if skip_prospect_row is not None and p["row"] == skip_prospect_row:
                continue
            est = PROSPECT_STATE_EN.get(p["estado"], p["estado"]) + \
                  ((", captured " + p["fecha"]) if p["fecha"] else "")
            key = str(p["row"])
            if self_dni and p["dni"] and p["dni"] == self_dni:
                add("pro-self:" + key, "Also registered as prospect «%s» (%s)" % (p["name"], est))
                continue
            if app_phone and p["phoneN"] == app_phone:
                add("pro-tel:" + key, "Same phone as prospect «%s» (%s)" % (p["name"], est))
            for slot, _rn, ph in refs:
                if ph and p["phoneN"] == ph:
                    add("pro-ref:%s:%s" % (slot, key), "%s is prospect «%s» (%s)" % (slot, p["name"], est))
        return items


# --------------------------------------------------------------------------- #
# build (GUI entry point)
# --------------------------------------------------------------------------- #
def build_rules(xlsx_path, today=None):
    """{'cfg', 'summary', 'v38', 'v39', 'v40', 'prospects'} — row dicts ready for
    the dashboard tables; never raises for a missing optional sheet."""
    d = RulesData(xlsx_path, today)
    v38, v39, v40, pros = [], [], [], []

    for c in d.clientes:
        if not (c["name"] or c["dniNorm"]):
            continue
        token = c["dniNorm"] or c["id"]

        n_ref, det38 = d.v38_for_client(c)
        if any(c["refPhones"]):
            v38.append({
                "name": c["name"] or "—", "client_id": c["id"], "dni": c["dniNorm"] or "—",
                "cuil": "", "n": n_ref,
                "verdict": "BLOCKS" if n_ref > 1 else ("flagged" if n_ref == 1 else "OK"),
                "items": det38, "detail": "  |  ".join(det38) or "—",
                "decision": "REJECT" if n_ref > 1 else ("REVIEW" if n_ref == 1 else "OK"),
                "token": token,
            })

        chk, own = d.v39_for_person(c["dniNorm"], c["id"], c["address"])
        if not chk["ok"]:
            items39, status, dec = d.exceeded_items(chk["exceeded"]), "BLOCKS", "REJECT"
        elif chk["alertas"]:
            items39, status, dec = chk["alertas"], "alert", "REVIEW"
        else:
            items39, status, dec = [], "OK", "OK"
        v39.append({
            "name": c["name"] or "—", "client_id": c["id"], "dni": c["dniNorm"] or "—", "cuil": "",
            "status": status, "nsig": len(chk["exceeded"] or []),
            "signals": ("last device row: " + ((own["ip"] or "—") + " · " + ((own["fp"] or "—")[:12]))
                        if own else "no device data — address only"),
            "items": items39, "detail": "  |  ".join(items39) or "—",
            "decision": dec, "token": token,
        })

        refs = [("Ref %d" % (i + 1), c["refNames"][i] or "—", c["refPhones"][i]) for i in range(3)]
        items = d.conexiones(c["phoneNorm"], refs, c["dniNorm"])
        if items:
            v40.append({
                "name": c["name"] or "—", "client_id": c["id"], "dni": c["dniNorm"] or "—", "cuil": "",
                "n": len(items), "items": items, "detail": "  |  ".join(items),
                "decision": "REVIEW", "token": token,
            })

    for p in d.prospectos:
        chk, own = d.v39_for_person(p["dni"], "", "") if p["dni"] else ({"ok": True, "exceeded": None, "alertas": []}, None)
        items = d.conexiones(p["phoneN"], [], p["dni"], skip_prospect_row=p["row"])
        warn = []
        if not chk["ok"]:
            warn += ["⚠ V-39 blocks — " + w for w in d.exceeded_items(chk["exceeded"])]
        else:
            warn += ["⚠ V-39 — " + a for a in chk["alertas"]]
        warn += ["🔗 V-40 — " + i for i in items]
        if not warn:
            continue
        pros.append({
            "name": p["name"] or "—", "cuil": p["cuil"] or "—",
            "estado": PROSPECT_STATE_EN.get(p["estado"], p["estado"]),
            "fecha": p["fecha"] or "—",
            "items": warn, "detail": "  |  ".join(warn),
            "decision": "REJECT" if not chk["ok"] else "REVIEW",
            "token": p["cuil"] or p["dni"] or p["name"],
        })

    v38.sort(key=lambda r: (-r["n"], r["name"]))
    v39.sort(key=lambda r: ({"BLOCKS": 0, "alert": 1, "OK": 2}[r["status"]], r["name"]))
    v40.sort(key=lambda r: (-r["n"], r["name"]))

    summary = {
        "cfg": d.cfg,
        "v38_blocks": sum(1 for r in v38 if r["n"] > 1),
        "v38_flagged": sum(1 for r in v38 if r["n"] == 1),
        "v39_blocks": sum(1 for r in v39 if r["status"] == "BLOCKS"),
        "v39_alerts": sum(1 for r in v39 if r["status"] == "alert"),
        "v40_people": len(v40),
        "prospects": len(pros),
        "clients": len(d.clientes),
    }
    return {"cfg": d.cfg, "summary": summary, "v38": v38, "v39": v39, "v40": v40,
            "prospects": pros, "today": d.today.isoformat()}


def quality_line(rules):
    if not rules:
        return "Rules mirror unavailable."
    s, cfg = rules["summary"], rules["cfg"]
    return ("Mirror of the Apps Script rules, as of %s · V-39 config: %s, max %d persons, GPS %dm, IP window %dd · "
            "V-38: %d would block / %d flagged · V-39: %d would block / %d alerts · V-40: %d people connected · "
            "%d prospects flagged — indicative, blocking happens in the sheet."
            % (rules["today"], "ON" if cfg["on"] else "OFF", cfg["max"], int(cfg["radioM"]), cfg["ipDias"],
               s["v38_blocks"], s["v38_flagged"], s["v39_blocks"], s["v39_alerts"], s["v40_people"], s["prospects"]))


def find_person(rules, cuil="", dni="", client_id=""):
    """Per-person rule rows for the Person view card. A CUIL embeds the DNI in
    digits 3-10 (the last digit is the check digit)."""
    if not rules:
        return None
    dnis = {d for d in (dni, norm_dni(cuil[2:10]) if len(cuil or "") == 11 else "") if d}
    out = {"v38": None, "v39": None, "v40": None}
    for key in ("v38", "v39", "v40"):
        for r in rules.get(key, []):
            if (client_id and r.get("client_id") == client_id) or (r.get("dni") in dnis):
                out[key] = r
                break
    return out if any(out.values()) else None


# --------------------------------------------------------------------------- #
# renderers
# --------------------------------------------------------------------------- #
def render_md(rules):
    if not rules:
        return "# Anti-fraud rules\n\n_Unavailable._\n"
    gen = datetime.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M")
    L = ["# Anti-fraud rules — V-38 · V-39 · V-40", "",
         "_Generated %s._  %s" % (gen, quality_line(rules)), ""]
    cell = lambda r: "<br>".join(r.get("items") or []) or "—"   # one finding per line
    L += ["## V-38 — references who are clients (max 1)", "",
          "| Person | DNI | Client refs | Verdict | Detail |", "|---|---|---:|---|---|"]
    for r in rules["v38"]:
        if r["n"] == 0:
            continue
        L.append("| %s (%s) | %s | %d | %s | %s |" % (r["name"], r["client_id"], r["dni"], r["n"], r["verdict"], cell(r)))
    L += ["", "## V-39 — device/location limit (if they applied again today)", "",
          "| Person | DNI | Status | Detail |", "|---|---|---|---|"]
    for r in rules["v39"]:
        if r["status"] == "OK":
            continue
        L.append("| %s (%s) | %s | %s | %s |" % (r["name"], r["client_id"], r["dni"], r["status"], cell(r)))
    L += ["", "## V-40 — phone connections (informative)", "",
          "| Person | DNI | # | Connections |", "|---|---|---:|---|"]
    for r in rules["v40"]:
        L.append("| %s (%s) | %s | %d | %s |" % (r["name"], r["client_id"], r["dni"], r["n"], cell(r)))
    L += ["", "## Prospects — warnings at registration", "",
          "| Prospect | CUIL | Status | Captured | Warnings |", "|---|---|---|---|---|"]
    for r in rules["prospects"]:
        L.append("| %s | %s | %s | %s | %s |" % (r["name"], r["cuil"], r["estado"], r["fecha"], cell(r)))
    L += ["", "> Loan-status notation: `L-0034 OVERDUE $ 12.000 · …` = that person's open loans; "
          "`no open loans (N settled)` = everything repaid. Sheet statuses are translated for "
          "display (ACTIVO→ACTIVE, VENCIDO→OVERDUE, PAGADO→PAID, SALDADO→SETTLED). Counts for the "
          "IP signal depend on the run date (rolling %d-day window, same as production)." % rules["cfg"]["ipDias"], ""]
    return "\n".join(L)


def render_person_md(pr):
    L = ["## 🚩 Fraud rule checks — V-38 · V-39 · V-40"]
    if pr.get("v38"):
        r = pr["v38"]
        L.append("- **V-38 · references** — %d client reference(s), verdict %s" % (r["n"], r["verdict"]))
        L += ["  - " + i for i in r.get("items") or []]
    if pr.get("v39"):
        r = pr["v39"]
        L.append("- **V-39 · device/location** — %s (%s)" % (r["status"], r["signals"]))
        L += ["  - " + i for i in r.get("items") or []]
    if pr.get("v40"):
        r = pr["v40"]
        L.append("- **V-40 · phone connections** — %d connection(s)" % r["n"])
        L += ["  - " + i for i in r.get("items") or []]
    return "\n".join(L)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main(argv=None):
    p = argparse.ArgumentParser(description="Mirror of the Apps Script anti-fraud rules (V-38/39/40).",
                                formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("who", nargs="?", help="CUIL, DNI, or ID Cliente")
    p.add_argument("--all", action="store_true", help="full report → out/_rules.md")
    p.add_argument("--xlsx", default=str(DEFAULT_XLSX))
    p.add_argument("--out", default=str(HERE / "out"))
    p.add_argument("--no-save", action="store_true")
    args = p.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    if not Path(args.xlsx).exists():
        logging.error("Workbook not found: %s", args.xlsx)
        return 2
    rules = build_rules(args.xlsx)
    if args.all or not args.who:
        print(quality_line(rules))
        print()
        print(render_md(rules))
        if not args.no_save:
            out = Path(args.out)
            out.mkdir(parents=True, exist_ok=True)
            (out / "_rules.md").write_text(render_md(rules) + "\n", encoding="utf-8")
            logging.info("Saved %s", out / "_rules.md")
        return 0
    digits = re.sub(r"\D", "", args.who)
    pr = find_person(rules, cuil=digits if len(digits) == 11 else "",
                     dni=digits if 7 <= len(digits) <= 8 else "",
                     client_id=args.who.strip().upper())
    if not pr:
        logging.error("No rule rows for '%s'.", args.who)
        return 2
    print(render_person_md(pr))
    return 0


if __name__ == "__main__":
    sys.exit(main())
