#!/usr/bin/env python3
"""
Impulso Crédito — Instagram Story video (1080×1920, ~12.5s).

Three phases:
  A) "Arranque seguro" splash — IC logo scales/fades in, wordmark, loading bar.
  B) App demo — a phone mockup scrolls through the real app screenshots.
  C) CTA outro.

Encoded with imageio's bundled ffmpeg (libx264 / yuv420p). Run: python3 make_video.py
"""

import math
from pathlib import Path

import imageio
import numpy as np
from PIL import Image, ImageDraw, ImageFilter

from generate_assets import (font, vgradient, ic_badge, BG_TOP, BG_BOT,
                             BLUE, GOLD, WHITE, SLATE)

HERE = Path(__file__).parent
SHOTS_DIR = Path.home() / ".claude/image-cache/7c33ccb2-f5ce-4af3-9302-4a8054b036b0"
W, H, FPS = 1080, 1920, 30

# ── Easing ───────────────────────────────────────────────────────────────────
def clamp(x, a=0.0, b=1.0): return max(a, min(b, x))
def ease_out_cubic(x): x = clamp(x); return 1 - (1 - x) ** 3
def ease_in_out(x): x = clamp(x); return x * x * (3 - 2 * x)
def ease_out_back(x):
    x = clamp(x); c1 = 1.70158; c3 = c1 + 1
    return 1 + c3 * (x - 1) ** 3 + c1 * (x - 1) ** 2
def lerp(a, b, t): return a + (b - a) * t


# ── Static assets ────────────────────────────────────────────────────────────
BG = vgradient(W, H, BG_TOP, BG_BOT)                     # RGBA navy gradient
_d = ImageDraw.Draw(BG)
_d.rectangle([0, 0, W, 8], fill=GOLD)
_d.rectangle([0, H - 8, W, H], fill=GOLD)
BADGE_MASTER = ic_badge(560)                             # resized per frame

# phone geometry
PHONE_W, PHONE_H = 664, 1300
PX0 = (W - PHONE_W) // 2; PY0 = 360; PX1 = PX0 + PHONE_W; PY1 = PY0 + PHONE_H
BEZEL = 24
SX, SY = PX0 + BEZEL, PY0 + BEZEL
SW, SH = PHONE_W - 2 * BEZEL, PHONE_H - 2 * BEZEL

# phone body (shadow + shell), static RGBA
def _build_phone():
    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    sh = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle([PX0, PY0 + 22, PX1, PY1 + 22], radius=66, fill=(0, 0, 0, 160))
    layer.alpha_composite(sh.filter(ImageFilter.GaussianBlur(42)))
    body = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    bd = ImageDraw.Draw(body)
    bd.rounded_rectangle([PX0, PY0, PX1, PY1], radius=60, fill=(9, 13, 24, 255))
    bd.rounded_rectangle([PX0, PY0, PX1, PY1], radius=60, outline=(64, 86, 128, 255), width=3)
    layer.alpha_composite(body)
    return layer

PHONE_BASE = _build_phone()
SCREEN_MASK = Image.new("L", (SW, SH), 0)
ImageDraw.Draw(SCREEN_MASK).rounded_rectangle([0, 0, SW - 1, SH - 1], radius=40, fill=255)

# app content strip: the 3 real screenshots, scaled to screen width, stacked
def _build_strip():
    strip_parts, tops, y = [], [], 0
    for n in (30, 31, 32):
        im = Image.open(SHOTS_DIR / f"{n}.png").convert("RGB")
        h = round(SW * im.height / im.width)
        strip_parts.append(im.resize((SW, h), Image.LANCZOS)); tops.append(y); y += h
    strip = Image.new("RGB", (SW, y), (18, 27, 45))
    yy = 0
    for p in strip_parts:
        strip.paste(p, (0, yy)); yy += p.height
    return strip, tops

STRIP, SECTION_TOPS = _build_strip()
SCROLL_MAX = max(0, STRIP.height - SH)


# ── Text helpers (alpha-aware) ───────────────────────────────────────────────
def ctext(draw, cx, y, text, f, rgb, a=1.0, tracking=0):
    col = (*rgb, int(255 * clamp(a)))
    if tracking:
        total = sum(draw.textlength(ch, font=f) for ch in text) + tracking * (len(text) - 1)
        x = cx - total / 2
        for ch in text:
            draw.text((x, y), ch, font=f, fill=col); x += draw.textlength(ch, font=f) + tracking
    else:
        draw.text((cx - draw.textlength(text, font=f) / 2, y), text, font=f, fill=col)

def wordmark(draw, cx, y, size, a=1.0):
    f = font(size, bold=True)
    w1 = draw.textlength("Impulso ", font=f); w2 = draw.textlength("Crédito", font=f)
    x = cx - (w1 + w2) / 2
    draw.text((x, y), "Impulso ", font=f, fill=(*BLUE, int(255 * clamp(a))))
    draw.text((x + w1, y), "Crédito", font=f, fill=(*GOLD, int(255 * clamp(a))))

def with_alpha(img, a):
    if a >= 1: return img
    img = img.copy(); img.putalpha(img.split()[3].point(lambda p: int(p * clamp(a)))); return img

def paste_badge(frame, cx, cy, disp, a=1.0):
    scale = disp / 560.0
    b = BADGE_MASTER.resize((int(BADGE_MASTER.width * scale), int(BADGE_MASTER.height * scale)), Image.LANCZOS)
    frame.alpha_composite(with_alpha(b, a), (int(cx - b.width / 2), int(cy - b.height / 2)))


# ── Phase A: splash ──────────────────────────────────────────────────────────
def render_splash(p):
    frame = BG.copy(); d = ImageDraw.Draw(frame)
    bp = ease_out_back((p - 0.05) / 0.30)                 # badge pop
    ba = ease_out_cubic((p - 0.05) / 0.22)
    paste_badge(frame, W / 2, 740, lerp(300, 500, bp), ba)
    wordmark(d, W / 2, 1010, 96, clamp((p - 0.34) / 0.16))
    ctext(d, W / 2, 1130, "Préstamos rápidos y confiables", font(40), SLATE,
          clamp((p - 0.42) / 0.15), tracking=1)
    # loading bar + label
    la = clamp((p - 0.48) / 0.12)
    ctext(d, W / 2, 1300, "A R R A N Q U E   S E G U R O", font(32, bold=True), GOLD, la, tracking=2)
    bw, bx, by = 440, (W - 440) // 2, 1380
    d.rounded_rectangle([bx, by, bx + bw, by + 16], radius=8, fill=(255, 255, 255, int(40 * la)))
    fillw = int(bw * ease_in_out((p - 0.52) / 0.36))
    if fillw > 10:
        d.rounded_rectangle([bx, by, bx + fillw, by + 16], radius=8, fill=(*GOLD, int(255 * la)))
    # fade out at the very end
    fo = clamp((p - 0.92) / 0.08)
    if fo > 0:
        ov = Image.new("RGBA", (W, H), (BG_BOT[0], BG_BOT[1], BG_BOT[2], int(255 * fo)))
        frame.alpha_composite(ov)
    return frame


# ── Phase B: app demo ────────────────────────────────────────────────────────
def render_app(p):
    frame = BG.copy(); d = ImageDraw.Draw(frame)
    # header
    ha = ease_out_cubic(p / 0.12)
    paste_badge(frame, W / 2 - 170, 150, 86, ha)
    wordmark(d, W / 2 + 48, 120, 48, ha)
    ctext(d, W / 2, 232, "Pedí tu préstamo en minutos", font(52, bold=True), WHITE, ha)

    # phone entrance (fade + slight slide up)
    ent = ease_out_cubic(p / 0.14); dy = int((1 - ent) * 70); pa = ent
    layer = PHONE_BASE.copy()

    # scrolling screen content
    scroll = ease_in_out((p - 0.14) / 0.80) * SCROLL_MAX
    view = Image.new("RGB", (SW, SH), (18, 27, 45))
    view.paste(STRIP, (0, -int(scroll)))
    layer.paste(view, (SX, SY), SCREEN_MASK)
    # dynamic island + screen rim
    ld = ImageDraw.Draw(layer)
    ld.rounded_rectangle([W // 2 - 78, PY0 + 16, W // 2 + 78, PY0 + 46], radius=15, fill=(0, 0, 0, 255))
    ld.rounded_rectangle([SX, SY, SX + SW - 1, SY + SH - 1], radius=40, outline=(70, 92, 134, 120), width=2)
    # tap/scroll indicator
    pulse = 0.5 + 0.5 * math.sin(p * 22)
    tx, ty = SX + int(SW * 0.60), SY + int(SH * 0.52) - int(scroll * 0.03 % 40)
    rr = 26 + int(12 * pulse)
    tap = Image.new("RGBA", (W, H), (0, 0, 0, 0)); td = ImageDraw.Draw(tap)
    td.ellipse([tx - rr, ty - rr, tx + rr, ty + rr], outline=(255, 255, 255, int(150 * (1 - pulse))), width=4)
    td.ellipse([tx - 16, ty - 16, tx + 16, ty + 16], fill=(255, 255, 255, 180))
    layer.alpha_composite(tap)

    frame.alpha_composite(with_alpha(layer, pa), (0, dy))

    # step caption (based on scroll position)
    center = scroll + SH / 2
    step = 1 + sum(1 for t in SECTION_TOPS[1:] if center >= t)
    labels = {1: "1 · Identificación", 2: "2 · Tus datos", 3: "3 · El préstamo"}
    ctext(d, W / 2, 1700, labels.get(step, ""), font(44, bold=True), WHITE, pa)
    ctext(d, W / 2, 1760, "100% online · con tu correo y DNI", font(32), SLATE, pa)
    return frame


# ── Phase C: outro ───────────────────────────────────────────────────────────
def render_outro(p):
    frame = BG.copy(); d = ImageDraw.Draw(frame)
    a = ease_out_cubic(p / 0.22)
    paste_badge(frame, W / 2, 640, 300, a)
    wordmark(d, W / 2, 870, 92, a)
    yb = 1040
    d.rounded_rectangle([W / 2 - 40, yb, W / 2 + 40, yb + 9], radius=4, fill=(*GOLD, int(255 * a)))
    ctext(d, W / 2, yb + 50, "Pedí el tuyo 100% online", font(58, bold=True), WHITE, a)
    ctext(d, W / 2, yb + 140, "@impulso.credito.ars", font(44, bold=True), GOLD, a)
    ctext(d, W / 2, yb + 210, "Zona Norte + AMBA", font(34), SLATE, a)
    fi = clamp((0.14 - p) / 0.14)                          # fade in from navy
    if fi > 0:
        frame.alpha_composite(Image.new("RGBA", (W, H), (*BG_BOT, int(255 * fi))))
    return frame


# ── Assemble ─────────────────────────────────────────────────────────────────
def build():
    A, B, C = int(3.3 * FPS), int(6.6 * FPS), int(2.7 * FPS)
    out = HERE / "story_video.mp4"
    previews = {int(A * 0.55): "preview_splash.png",
                A + int(B * 0.55): "preview_app.png",
                A + B + int(C * 0.5): "preview_outro.png"}
    writer = imageio.get_writer(out, fps=FPS, codec="libx264", macro_block_size=8,
                                pixelformat="yuv420p", ffmpeg_params=["-crf", "20", "-preset", "medium"])
    total = A + B + C
    for i in range(total):
        if i < A:
            fr = render_splash(i / (A - 1))
        elif i < A + B:
            fr = render_app((i - A) / (B - 1))
        else:
            fr = render_outro((i - A - B) / (C - 1))
        rgb = fr.convert("RGB")
        if i in previews:
            rgb.save(HERE / previews[i], "PNG")
        writer.append_data(np.asarray(rgb))
    writer.close()
    print(f"Wrote {out}  ({total} frames, {total / FPS:.1f}s)")


if __name__ == "__main__":
    build()
