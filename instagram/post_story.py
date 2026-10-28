#!/usr/bin/env python3
"""
Impulso Crédito — Daily Instagram Story Poster

Generates a branded 1080×1920 story image with Claude-written content
and publishes it via the Instagram Graph API.

Usage:
    python3 post_story.py

Setup: copy .env.example → .env and fill in the four keys.
"""

import base64
import logging
import os
import sys
import textwrap
import time
from datetime import date
from pathlib import Path

import anthropic
import requests
from dotenv import load_dotenv
from PIL import Image, ImageDraw, ImageFont

load_dotenv(Path(__file__).parent / ".env")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ── Brand colours ──────────────────────────────────────────────────────────
NAVY   = (47,  71,  136)   # #2f4788
TEAL   = (15,  118, 110)   # #0f766e
TEAL_D = (10,  80,  75)    # darker teal for CTA box fill
WHITE  = (255, 255, 255)
SLATE  = (148, 163, 184)   # #94a3b8
TEAL_L = (165, 243, 252)   # #a5f3fc

# ── Canvas ─────────────────────────────────────────────────────────────────
W, H     = 1080, 1920
TMP_FILE     = Path(__file__).parent / "story_tmp.png"
PREVIEW_FILE = Path(__file__).parent / "story_preview.png"  # kept on --dry-run
IG_API   = "https://graph.instagram.com/v21.0"   # Instagram API with Instagram Login (no FB Page)

FONT_BOLD   = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
FONT_NORMAL = "/System/Library/Fonts/Supplemental/Arial.ttf"


def _font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    primary = FONT_BOLD if bold else FONT_NORMAL
    fallbacks = [
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
            else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold
            else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    for path in [primary] + fallbacks:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def _centered_x(draw: ImageDraw.ImageDraw, text: str,
                 font: ImageFont.FreeTypeFont) -> int:
    bb = draw.textbbox((0, 0), text, font=font)
    return (W - (bb[2] - bb[0])) // 2


# ── Step 1: Claude generates the daily message ─────────────────────────────

def generate_message() -> str:
    today = date.today().strftime("%A %d de %B de %Y")
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    resp = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=120,
        messages=[{"role": "user", "content": (
            f"Escribí un mensaje corto para una historia de Instagram de Impulso Crédito, "
            f"un servicio de préstamos personales en zona norte y AMBA, Buenos Aires, Argentina. "
            f"Hoy es {today}. Usá español argentino informal (\"vos\"). "
            f"Máximo 100 caracteres. Cálido, urgente, práctico. "
            f"Variá el tema según el día: tasas, velocidad, simplicidad, WhatsApp, "
            f"confianza, o un tip financiero. "
            f"Devolvé SOLO el texto del mensaje, sin comillas ni explicaciones."
        )}],
    )
    return resp.content[0].text.strip()


# ── Step 2: Render branded 1080×1920 PNG ───────────────────────────────────

def render_story(message: str) -> None:
    today_str = date.today().strftime("%A %d de %B")

    img  = Image.new("RGB", (W, H), NAVY)
    draw = ImageDraw.Draw(img)

    # Teal accent bars
    draw.rectangle([0, 0, W, 8], fill=TEAL)
    draw.rectangle([0, H - 8, W, H], fill=TEAL)

    # Brand name
    f_brand = _font(88, bold=True)
    brand = "Impulso Crédito"
    draw.text((_centered_x(draw, brand, f_brand), 110), brand, font=f_brand, fill=WHITE)

    # Tagline
    f_tag = _font(38)
    tag = "Préstamos personales · Zona Norte + AMBA"
    draw.text((_centered_x(draw, tag, f_tag), 230), tag, font=f_tag, fill=TEAL_L)

    # Divider below tagline
    draw.line([120, 310, W - 120, 310], fill=TEAL, width=2)

    # Main message — wrapped and vertically centred in the middle band
    f_msg    = _font(66, bold=True)
    lines    = textwrap.wrap(message, width=20)
    line_h   = 86
    total_h  = len(lines) * line_h
    y_msg    = (H // 2) - (total_h // 2) - 40
    for i, line in enumerate(lines):
        draw.text((_centered_x(draw, line, f_msg), y_msg + i * line_h),
                  line, font=f_msg, fill=WHITE)

    # Horizontal separator beneath message
    sep_y = y_msg + total_h + 50
    draw.line([120, sep_y, W - 120, sep_y], fill=TEAL, width=3)

    # Date
    f_date = _font(42)
    draw.text((_centered_x(draw, today_str, f_date), sep_y + 24),
              today_str, font=f_date, fill=SLATE)

    # CTA box
    cta_y = H - 300
    draw.rounded_rectangle([60, cta_y, W - 60, cta_y + 200], radius=28, fill=TEAL_D)

    f_num = _font(52, bold=True)
    num   = "+54 9 11 3925-5583"
    draw.text((_centered_x(draw, num, f_num), cta_y + 28), num, font=f_num, fill=TEAL_L)

    f_cta = _font(36)
    cta   = "Pedí tu préstamo hoy · WhatsApp"
    draw.text((_centered_x(draw, cta, f_cta), cta_y + 110), cta, font=f_cta, fill=WHITE)

    img.save(TMP_FILE, "PNG")
    log.info("Image rendered → %s", TMP_FILE)


# ── Step 3: Upload to ImgBB (public URL for Instagram to fetch) ─────────────

def upload_to_imgbb() -> str:
    with open(TMP_FILE, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    r = requests.post(
        "https://api.imgbb.com/1/upload",
        data={"key": os.environ["IMGBB_API_KEY"], "image": b64},
        timeout=60,
    )
    r.raise_for_status()
    url = r.json()["data"]["url"]          # direct i.ibb.co image link
    log.info("Hosted at %s", url)
    return url


# ── Step 4: Post to Instagram ──────────────────────────────────────────────

def post_to_instagram(image_url: str) -> str:
    token = os.environ["IG_ACCESS_TOKEN"]
    uid   = os.environ["IG_USER_ID"]

    # Publish as a STORY: the image is 1080×1920 (9:16), which the feed rejects
    # (feed photos must be 4:5 … 1.91:1). media_type=STORIES matches the render.
    r = requests.post(
        f"{IG_API}/{uid}/media",
        data={"image_url": image_url, "media_type": "STORIES", "access_token": token},
        timeout=30,
    )
    if not r.ok:                       # surface Facebook's actual reason, not just the status
        log.error("IG container error %s: %s", r.status_code, r.text[:800])
    r.raise_for_status()
    creation_id = r.json()["id"]
    log.info("Container: %s", creation_id)

    # The container isn't publishable until Instagram finishes processing it; publishing
    # too early returns "Media Not Found". Poll status_code until FINISHED.
    status = None
    for _ in range(20):
        s = requests.get(f"{IG_API}/{creation_id}",
                         params={"fields": "status_code", "access_token": token}, timeout=30)
        status = s.json().get("status_code") if s.ok else None
        if status == "FINISHED":
            break
        if status == "ERROR":
            log.error("Container processing failed: %s", s.text[:400])
            raise RuntimeError("Instagram could not process the image container.")
        time.sleep(2)
    log.info("Container status: %s", status)

    r2 = requests.post(
        f"{IG_API}/{uid}/media_publish",
        data={"creation_id": creation_id, "access_token": token},
        timeout=30,
    )
    if not r2.ok:
        log.error("IG publish error %s: %s", r2.status_code, r2.text[:800])
    r2.raise_for_status()
    media_id = r2.json()["id"]
    log.info("Published: %s", media_id)
    return media_id


# ── Main ───────────────────────────────────────────────────────────────────

def main(dry_run: bool = False, message=None) -> None:
    log.info("── Impulso Crédito story · %s%s ──",
             date.today(), "  (DRY RUN)" if dry_run else "")
    try:
        message = message or generate_message()
        log.info("Message: %s", message)
        render_story(message)
        if dry_run:
            TMP_FILE.replace(PREVIEW_FILE)
            log.info("DRY RUN — nothing uploaded or posted. Preview → %s", PREVIEW_FILE)
            return
        image_url = upload_to_imgbb()
        post_to_instagram(image_url)
        log.info("Done.")
    except Exception as e:
        log.error("Failed: %s", e)
        raise
    finally:
        if not dry_run and TMP_FILE.exists():
            TMP_FILE.unlink()


if __name__ == "__main__":
    argv = sys.argv[1:]
    dry = "--dry-run" in argv
    custom = None
    if "--message" in argv:          # optional: render a fixed text (offline, no Claude call)
        i = argv.index("--message")
        custom = argv[i + 1] if i + 1 < len(argv) else None
    main(dry_run=dry, message=custom)
