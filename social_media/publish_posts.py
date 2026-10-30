#!/usr/bin/env python3
"""
Impulso Crédito — publish the feature posts to Instagram (feed, 1:1).

Reads feat_XX.png + feat_XX.txt, hosts each image on ImgBB, and publishes it as
a feed post with its caption via the Instagram API (Instagram Login).

SAFEGUARDS (publishing is irreversible + rate-limited):
  • Dry-run by default — prints what it WOULD post. Use --publish to post.
  • --limit N   max posts this run (default 5). Instagram caps ~50 posts/24h.
  • --delay S   seconds between posts (default 45; use 300 for 5 min apart).
  • Skips any feature image already published (matched by perceptual hash),
    so it's safe to re-run and resume where you left off.

Every step is timestamped and written to BOTH the console and publish.log.

Env (from instagram/.env): IG_ACCESS_TOKEN, IG_USER_ID, IMGBB_API_KEY.

Examples:
  python3 publish_posts.py                              # dry run — show the plan
  python3 publish_posts.py --publish --limit 30 --delay 300    # all 30, 5 min apart
"""

import base64
import io
import logging
import os
import sys
import time
import warnings
from datetime import datetime, timedelta
from pathlib import Path

warnings.filterwarnings("ignore")        # silence the LibreSSL/urllib3 notice

import numpy as np
import requests
from dotenv import load_dotenv
from PIL import Image

HERE = Path(__file__).parent
load_dotenv(HERE.parent / "instagram" / ".env")
load_dotenv(HERE / ".env")
IG_API = "https://graph.instagram.com/v21.0"
TOKEN = os.environ.get("IG_ACCESS_TOKEN", "")
UID = os.environ.get("IG_USER_ID", "")
IMGBB = os.environ.get("IMGBB_API_KEY", "")

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s  %(message)s", datefmt="%H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout),
              logging.FileHandler(HERE / "publish.log", encoding="utf-8")])
log = logging.getLogger("publish")


def ahash(img):
    g = np.asarray(img.convert("L").resize((16, 16), Image.LANCZOS), dtype=np.float32)
    return (g > g.mean()).flatten()

def hamming(a, b):
    return int(np.count_nonzero(a != b))


def published_hashes():
    out, url, params = [], f"{IG_API}/me/media", {
        "fields": "id,media_type,media_url", "limit": 50, "access_token": TOKEN}
    while url:
        r = requests.get(url, params=params, timeout=30)
        if not r.ok:
            log.error("ERROR reading media: %s %s", r.status_code, r.text[:300]); sys.exit(1)
        j = r.json()
        for m in j.get("data", []):
            if m.get("media_type") == "IMAGE" and m.get("media_url"):
                try:
                    im = Image.open(io.BytesIO(requests.get(m["media_url"], timeout=30).content))
                    out.append(ahash(im))
                except Exception:
                    pass
        url = j.get("paging", {}).get("next"); params = None
    return out


def upload_imgbb(path):
    b64 = base64.b64encode(path.read_bytes()).decode()
    r = requests.post("https://api.imgbb.com/1/upload",
                      data={"key": IMGBB, "image": b64}, timeout=60)
    r.raise_for_status()
    return r.json()["data"]["url"]


def publish_feed(image_url, caption):
    r = requests.post(f"{IG_API}/{UID}/media",
                      data={"image_url": image_url, "caption": caption, "access_token": TOKEN}, timeout=30)
    if not r.ok:
        log.error("   ✗ container error: %s %s", r.status_code, r.text[:300]); return None
    cid = r.json()["id"]
    log.info("   contenedor creado: %s (procesando…)", cid)
    for _ in range(20):
        s = requests.get(f"{IG_API}/{cid}", params={"fields": "status_code", "access_token": TOKEN}, timeout=30)
        st = s.json().get("status_code") if s.ok else None
        if st == "FINISHED":
            break
        if st == "ERROR":
            log.error("   ✗ procesamiento falló: %s", s.text[:300]); return None
        time.sleep(2)
    r2 = requests.post(f"{IG_API}/{UID}/media_publish",
                       data={"creation_id": cid, "access_token": TOKEN}, timeout=30)
    if not r2.ok:
        log.error("   ✗ publish error: %s %s", r2.status_code, r2.text[:300]); return None
    return r2.json()["id"]


def main(do_publish, limit, delay):
    if not (TOKEN and UID):
        log.error("Falta IG_ACCESS_TOKEN / IG_USER_ID en instagram/.env"); return
    posts = sorted(HERE.glob("feat_*.png"))
    if not posts:
        log.error("No hay feat_*.png — corré generate_feature_posts.py primero."); return

    log.info("Leyendo posts ya publicados (para no duplicar)…")
    pub = published_hashes()
    todo = []
    for p in posts:
        h = ahash(Image.open(p))
        if any(hamming(h, ph) <= 8 for ph in pub):
            continue
        cap = p.with_suffix(".txt")
        if not cap.exists():
            log.warning("  (sin caption para %s — salteado)", p.name); continue
        todo.append((p, cap.read_text(encoding="utf-8")))

    log.info("%d posts · %d ya publicados · %d pendientes.",
             len(posts), len(posts) - len(todo), len(todo))
    batch = todo[:limit]

    if not do_publish:
        log.info("DRY RUN — publicaría %d ahora (1.ª línea de cada caption):", len(batch))
        for p, cap in batch:
            log.info("  • %s: %s", p.name, cap.splitlines()[0])
        log.info("Agregá --publish para publicar. Volumen con --limit / --delay.")
        return

    if not IMGBB:
        log.error("Falta IMGBB_API_KEY en instagram/.env"); return
    total_min = round((len(batch) - 1) * delay / 60) if len(batch) > 1 else 0
    log.info("── Publicando %d post(s), 1 cada %ds (~%d min en total) ──",
             len(batch), delay, total_min)
    done = 0
    for idx, (p, cap) in enumerate(batch):
        log.info("[%d/%d] ENVIANDO %s …", idx + 1, len(batch), p.name)
        try:
            url = upload_imgbb(p)
            log.info("   imagen en ImgBB: %s", url)
        except Exception as e:
            log.error("   ✗ falló la subida a ImgBB: %s", e); continue
        mid = publish_feed(url, cap)
        if mid:
            log.info("   ✓ PUBLICADO %s  (media id %s)", p.name, mid); done += 1
        else:
            log.error("   ✗ no se pudo publicar %s", p.name)
        if idx < len(batch) - 1:
            nxt = (datetime.now() + timedelta(seconds=delay)).strftime("%H:%M:%S")
            log.info("   … próximo post a las %s (espera %ds)\n", nxt, delay)
            time.sleep(delay)
    log.info("── Listo: %d publicado(s), %d pendiente(s). Volvé a correr para continuar. ──",
             done, len(todo) - done)


if __name__ == "__main__":
    args = sys.argv[1:]
    def opt(name, default):
        return int(args[args.index(name) + 1]) if name in args and args.index(name) + 1 < len(args) else default
    main(do_publish="--publish" in args, limit=opt("--limit", 5), delay=opt("--delay", 45))
