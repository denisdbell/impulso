#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Descarga por lotes de la Central de Deudores del BCRA, un archivo por persona.

Lee una lista de CUILs (uno por línea) y, para cada uno, consulta la API pública
del BCRA y guarda un JSON por persona en una carpeta de salida.

Diseñado para NO ser marcado como DDoS ni ser limitado (throttling):
  - Pide de a uno por vez (secuencial, sin concurrencia).
  - Espacia cada pedido con una pausa + jitter aleatorio.
  - Reintenta con backoff exponencial ante 429 / 5xx (respeta Retry-After).
  - Envía un User-Agent honesto y Accept: application/json.
  - Reanuda sin volver a pegarle a la API (saltea los que ya tienen archivo).

Solo usa la biblioteca estándar de Python 3 (no requiere pip install).

Uso típico (tras pegar los CUILs en bcra/cuils.txt):
    python3 bcra/fetch_bcra.py

Ver todas las opciones:
    python3 bcra/fetch_bcra.py --help
"""

import argparse
import csv
import json
import logging
import random
import re
import socket
import ssl
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

# Mismo endpoint base que usa el Apps Script (LoanManagerV2.gs: BCRA_BASE).
BCRA_BASE = "https://api.bcra.gob.ar/CentralDeDeudores/v1.0"

# Endpoints de la Central de Deudores. Clave de salida -> ruta de la API.
ENDPOINTS = {
    "deudas": "Deudas",                        # deuda actual (último período)
    "historicas": "Deudas/Historicas",         # 24 meses de historia (la URL pedida)
    "chequesRechazados": "Deudas/ChequesRechazados",  # cheques rechazados
}

USER_AGENT = "impulso-bcra-batch/1.0 (contacto: denisdbell@gmail.com)"


# --------------------------------------------------------------------------- #
# Entrada: lectura y validación de CUILs
# --------------------------------------------------------------------------- #
def read_cuils(path):
    """Lee el archivo de entrada: un CUIL por línea.

    Ignora líneas vacías y comentarios (#). De cada línea toma solo los dígitos,
    así tolera formatos como 20-12345678-9. De-duplica preservando el orden.
    Devuelve (cuils_validos, lineas_invalidas).
    """
    cuils, invalidos, vistos = [], [], set()
    for n, raw in enumerate(Path(path).read_text(encoding="utf-8").splitlines(), 1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        digits = re.sub(r"\D", "", line)
        if len(digits) != 11:
            invalidos.append((n, raw, "no tiene 11 dígitos"))
            continue
        if digits in vistos:
            continue
        vistos.add(digits)
        cuils.append(digits)
    return cuils, invalidos


def cuil_check_digit_ok(cuil):
    """Valida el dígito verificador (mod-11 AFIP). Solo para advertir, no bloquea.

    Replica la lógica de cuilCheckDigit_ en LoanManagerV2.gs (reimplementada aquí
    porque este script es independiente del proyecto Apps Script).
    """
    weights = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    s = sum(int(cuil[i]) * weights[i] for i in range(10))
    rem = s % 11
    dv = 11 - rem
    if dv == 11:
        dv = 0
    elif dv == 10:
        # Caso especial de AFIP (suele implicar prefijo 23); lo aceptamos sin falla.
        return True
    return dv == int(cuil[10])


# --------------------------------------------------------------------------- #
# HTTP: pedido con reintentos, backoff y pausa cortés
# --------------------------------------------------------------------------- #
def make_opener(insecure):
    ctx = ssl.create_default_context()
    if insecure:
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    return urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))


def _retry_after_seconds(err):
    """Lee el header Retry-After (segundos) si viene; None si no se puede."""
    try:
        val = err.headers.get("Retry-After")
        if val is None:
            return None
        return int(float(val))
    except (TypeError, ValueError):
        return None


def http_get(opener, url, timeout, retries, pace):
    """Un GET con reintentos. Devuelve {status, data, error}.

    - 200: {status:200, data:<json>, error:None}
    - 404: {status:404, data:None, error:'sin registro (404)'}  (sin reintento)
    - 429/5xx: reintenta con backoff; si se agota, status con error.
    - errores de red: reintenta con backoff; si se agota, status None con error.
    `pace` es un callable que duerme (pausa + jitter) antes del primer intento.
    """
    attempt = 0
    while True:
        if attempt == 0:
            pace()  # espaciado cortés antes de cada llamada nueva
        try:
            req = urllib.request.Request(
                url,
                headers={"User-Agent": USER_AGENT, "Accept": "application/json"},
                method="GET",
            )
            with opener.open(req, timeout=timeout) as resp:
                code = resp.getcode()
                raw = resp.read().decode("utf-8", "replace")
                try:
                    data = json.loads(raw)
                except json.JSONDecodeError:
                    data = None
                return {"status": code, "data": data, "error": None if data is not None else "JSON inválido"}
        except urllib.error.HTTPError as e:
            code = e.code
            if code == 404:
                return {"status": 404, "data": None, "error": "sin registro (404)"}
            if code == 429 or code >= 500:
                if attempt >= retries:
                    return {"status": code, "data": None, "error": "agotados reintentos (HTTP %d)" % code}
                wait = _retry_after_seconds(e) or (5 * (3 ** attempt))
                logging.warning("    HTTP %d — reintento %d/%d en %ds", code, attempt + 1, retries, wait)
                time.sleep(wait)
                attempt += 1
                continue
            return {"status": code, "data": None, "error": "HTTP %d" % code}
        except (urllib.error.URLError, ssl.SSLError, socket.timeout, TimeoutError) as e:
            if attempt >= retries:
                return {"status": None, "data": None, "error": "red: %s" % e}
            wait = 5 * (3 ** attempt)
            logging.warning("    error de red (%s) — reintento %d/%d en %ds", e, attempt + 1, retries, wait)
            time.sleep(wait)
            attempt += 1
            continue


def make_pace(delay, jitter):
    """Devuelve un callable que duerme (pausa base + jitter) antes de cada pedido."""
    def pace():
        time.sleep(delay + random.uniform(0, jitter))
    return pace


def fetch_one(cuil, out_dir, opener, want, timeout, retries, pace, force=False):
    """Descarga (o lee de disco) los datos de la Central de Deudores de UN CUIL.

    Punto único de descarga por persona, reutilizado por el lote (main) y por el
    analizador (analyze_bcra.py), para que exista una sola lógica cortés de red.

    Devuelve dict: {combined, skipped, errors, statuses, dest}.
      - combined : datos combinados ({cuil, fetched_at, deudas, historicas,
                   chequesRechazados, _meta}); desde disco si ya existía.
      - skipped  : True si se devolvió un archivo ya existente (sin pegarle a la API).
      - errors   : lista de (cuil, seccion, motivo) para fallos no-404/no-200.
      - statuses : {seccion: codigo_http}.
      - dest     : Path del archivo JSON.
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = out_dir / ("%s.json" % cuil)

    if dest.exists() and not force:
        try:
            combined = json.loads(dest.read_text(encoding="utf-8"))
            statuses = (combined.get("_meta") or {}).get("http_status", {}) or {}
            return {"combined": combined, "skipped": True, "errors": [], "statuses": statuses, "dest": dest}
        except (json.JSONDecodeError, OSError):
            pass  # archivo ilegible → se vuelve a descargar

    combined = {"cuil": cuil, "fetched_at": datetime.now().astimezone().isoformat(timespec="seconds")}
    meta_status, meta_urls, errors = {}, {}, []
    for key, pathseg in want:
        url = "%s/%s/%s" % (BCRA_BASE, pathseg, cuil)
        res = http_get(opener, url, timeout, retries, pace)
        combined[key] = res["data"]
        meta_status[key] = res["status"]
        meta_urls[key] = url
        if res["status"] not in (200, 404):
            errors.append((cuil, key, res["error"] or "error"))
    combined["_meta"] = {"http_status": meta_status, "source_urls": meta_urls}
    dest.write_text(json.dumps(combined, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"combined": combined, "skipped": False, "errors": errors, "statuses": meta_status, "dest": dest}


# --------------------------------------------------------------------------- #
# Resumen para el índice CSV (lee directo del results de la API, sin puntajes)
# --------------------------------------------------------------------------- #
def summarize(combined):
    """Extrae campos livianos para el índice. Tolera secciones ausentes."""
    denom, peor, ent_count, chq_count = "", "", "", ""

    def results_of(key):
        sec = combined.get(key)
        return (sec.get("results") or {}) if isinstance(sec, dict) else {}

    # Denominación: de la primera sección que la tenga.
    for key in ("deudas", "historicas", "chequesRechazados"):
        r = results_of(key)
        if r.get("denominacion"):
            denom = r["denominacion"]
            break

    # Peor situación y cantidad de entidades: del período más reciente de deudas,
    # o del primero de historicas si no hay deuda actual.
    for key in ("deudas", "historicas"):
        r = results_of(key)
        periodos = r.get("periodos") or []
        if periodos:
            ents = periodos[0].get("entidades") or []
            ent_count = len(ents)
            sits = [e.get("situacion") for e in ents if isinstance(e.get("situacion"), int)]
            if sits:
                peor = max(sits)
            break

    # Cheques rechazados: total de detalles bajo causales -> entidades -> detalle.
    r = results_of("chequesRechazados")
    try:
        n = 0
        for c in (r.get("causales") or []):
            for e in (c.get("entidades") or []):
                n += len(e.get("detalle") or [])
        chq_count = n
    except (AttributeError, TypeError):
        chq_count = ""

    return denom, peor, ent_count, chq_count


# --------------------------------------------------------------------------- #
# Principal
# --------------------------------------------------------------------------- #
def parse_args(argv):
    here = Path(__file__).resolve().parent
    p = argparse.ArgumentParser(
        description="Descarga por lotes de la Central de Deudores del BCRA (un archivo por CUIL).",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--input", default=str(here / "cuils.txt"), help="archivo con un CUIL por línea")
    p.add_argument("--out", default=str(here / "out"), help="carpeta de salida")
    p.add_argument("--endpoints", choices=["all", "historicas"], default="all",
                   help="'all' = Deudas + Historicas + ChequesRechazados; 'historicas' = solo la histórica")
    p.add_argument("--delay", type=float, default=2.0, help="pausa base (seg) entre pedidos")
    p.add_argument("--jitter", type=float, default=1.0, help="aleatorio extra (seg) sumado a la pausa")
    p.add_argument("--retries", type=int, default=4, help="reintentos ante 429/5xx/red")
    p.add_argument("--timeout", type=int, default=30, help="timeout por pedido (seg)")
    p.add_argument("--force", action="store_true", help="re-descargar aunque el archivo ya exista")
    p.add_argument("--insecure", action="store_true",
                   help="desactivar verificación TLS (recurso ante el certificado incompleto del BCRA)")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    in_path = Path(args.input)
    if not in_path.exists():
        logging.error("No existe el archivo de entrada: %s", in_path)
        logging.error("Pegá los CUILs (uno por línea) en ese archivo y volvé a correr.")
        return 2

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    cuils, invalidos = read_cuils(in_path)
    for n, raw, why in invalidos:
        logging.warning("línea %d inválida (%s): %r", n, why, raw)
    if not cuils:
        logging.error("No hay CUILs válidos en %s", in_path)
        return 2

    if args.insecure:
        logging.warning("TLS sin verificar (--insecure): úsalo solo si el certificado del BCRA falla.")
    opener = make_opener(args.insecure)

    want = list(ENDPOINTS.items()) if args.endpoints == "all" else [("historicas", ENDPOINTS["historicas"])]
    pace = make_pace(args.delay, args.jitter)

    errors = [("línea %d" % n, why, raw) for n, raw, why in invalidos]
    index_rows = []
    counts = {"ok": 0, "sin_registro": 0, "fallo": 0, "salteado": 0}
    start = time.time()
    total = len(cuils)

    try:
        for i, cuil in enumerate(cuils, 1):
            if not cuil_check_digit_ok(cuil):
                logging.warning("[%d/%d] %s — dígito verificador no coincide (se consulta igual)", i, total, cuil)

            r = fetch_one(cuil, out_dir, opener, want, args.timeout, args.retries, pace, force=args.force)
            combined, statuses = r["combined"], r["statuses"]
            errors.extend(r["errors"])
            denom, peor, ent, chq = summarize(combined)
            index_rows.append([cuil, denom, peor, ent, chq, json.dumps(statuses, ensure_ascii=False), r["dest"].name])

            if r["skipped"]:
                counts["salteado"] += 1
                logging.info("[%d/%d] %s — salteado (ya existe)", i, total, cuil)
                continue

            # Clasificación del resultado para el log/resumen.
            any_data = any(combined.get(k) is not None for k, _ in want)
            status_vals = list(statuses.values())
            if any_data:
                counts["ok"] += 1
                estado = "ok"
            elif status_vals and all(s == 404 for s in status_vals):
                counts["sin_registro"] += 1
                estado = "sin registro (404)"
            elif 200 in status_vals:
                counts["ok"] += 1
                estado = "ok (sin datos)"
            else:
                counts["fallo"] += 1
                estado = "FALLO"

            logging.info("[%d/%d] %s → %s | deudas:%s cheques:%s | %s",
                         i, total, cuil, denom or "—",
                         ent if ent != "" else "?", chq if chq != "" else "?", estado)

    except KeyboardInterrupt:
        logging.warning("\nInterrumpido. Se guarda el progreso parcial (índice y errores).")

    finally:
        # Índice CSV (utf-8-sig para que Excel lea bien los acentos).
        index_path = out_dir / "_index.csv"
        with index_path.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["cuil", "denominacion", "peor_situacion", "entidades",
                        "cheques_rechazados", "http_status", "archivo"])
            w.writerows(index_rows)

        # Log de errores.
        err_path = out_dir / "_errors.log"
        if errors:
            err_path.write_text(
                "\n".join("%s\t%s\t%s" % (a, b, c) for a, b, c in errors) + "\n",
                encoding="utf-8",
            )
        elif err_path.exists():
            err_path.unlink()

    elapsed = time.time() - start
    logging.info(
        "\nListo en %.0fs — ok:%d  sin registro:%d  fallo:%d  salteado:%d  (índice: %s)",
        elapsed, counts["ok"], counts["sin_registro"], counts["fallo"], counts["salteado"],
        out_dir / "_index.csv",
    )
    if errors:
        logging.info("Errores registrados en %s", out_dir / "_errors.log")
    return 1 if counts["fallo"] else 0


if __name__ == "__main__":
    sys.exit(main())
