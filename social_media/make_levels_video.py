#!/usr/bin/env python3
"""
Impulso Crédito — "Escalera de niveles" (graduación) Instagram Story video.

1080×1920 (9:16), 30 fps, 15 s — tamaño ideal para una historia de Instagram.
Explica, con animación de motion-graphics de marca, cómo funciona la escalera de
graduación (V-32): se arranca en $150.000, se sube +$50.000 por nivel pagando a
tiempo, hasta un tope de $500.000; atrasarse hace bajar de nivel. Cierra con el
logo IC (mismo estilo que impulso_add.mp4 / make_video.py).

Fases:
  A) Intro  — etiqueta + título "Subí de nivel".
  B) Escalada — la escalera se enciende peldaño a peldaño mientras la moneda sube.
  C) Premio — "Hasta $500.000 · Accedé a montos más altos" + estrellas.
  D) Aviso  — "¿Te atrasás? Bajás de nivel" (la moneda resbala un peldaño).
  E) Cierre — badge IC + marca denominativa + CTA.

Codificado con el ffmpeg incluido en imageio (libx264 / yuv420p).
Correr:  python3 make_levels_video.py
"""

import math
from pathlib import Path

import imageio
import numpy as np
from PIL import Image, ImageDraw, ImageFilter

from generate_assets import (font, vgradient, ic_badge, BG_TOP, BG_BOT,
                             BLUE, GOLD, WHITE, SLATE)

HERE = Path(__file__).parent
W, H, FPS = 1080, 1920, 30
GREEN = (61, 190, 107)
RED = (226, 96, 96)

# ── Easing ─────────────────────────────────────────────────────────────────────
def clamp(x, a=0.0, b=1.0): return max(a, min(b, x))
def ease_out_cubic(x): x = clamp(x); return 1 - (1 - x) ** 3
def ease_in_out(x): x = clamp(x); return x * x * (3 - 2 * x)
def ease_out_back(x):
    x = clamp(x); c1 = 1.70158; c3 = c1 + 1
    return 1 + c3 * (x - 1) ** 3 + c1 * (x - 1) ** 2
def lerp(a, b, t): return a + (b - a) * t
def money(v): return "$" + f"{v:,.0f}".replace(",", ".")


def ctext(d, cx, y, text, f, rgb, a=1.0, tracking=0):
    col = (*rgb, int(255 * clamp(a)))
    if tracking:
        total = sum(d.textlength(ch, font=f) for ch in text) + tracking * (len(text) - 1)
        x = cx - total / 2
        for ch in text:
            d.text((x, y), ch, font=f, fill=col); x += d.textlength(ch, font=f) + tracking
    else:
        d.text((cx - d.textlength(text, font=f) / 2, y), text, font=f, fill=col)


# ── Static background (navy gradient + gold bars + faint dots) ──────────────────
def _bg():
    img = vgradient(W, H, BG_TOP, BG_BOT)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, 8], fill=GOLD)
    d.rectangle([0, H - 8, W, H], fill=GOLD)
    for (x, y, r, a) in [(180, 520, 7, 40), (920, 700, 9, 36), (240, 1200, 6, 34),
                         (860, 1320, 8, 40), (540, 300, 5, 30), (700, 1500, 6, 30)]:
        d.ellipse([x - r, y - r, x + r, y + r], fill=(*BLUE, a))
    return img

BG = _bg()
NAVY = (*BG_BOT,)

# ── Geometry: eight stacked level-bars (+$50.000 each), coin climbs the left rail
NLV = 8
TOP = NLV - 1
BAR_W, BAR_H = 724, 94               # bar size
BXC = 606                           # bar centre x  (spans 244..968)
BY = [1498 - k * 150 for k in range(NLV)]            # bar centres, bottom→top
AMOUNTS = [150000 + k * 50000 for k in range(NLV)]   # 150k → 500k, +$50k/nivel
COIN_R = 44
COIN_X = 128                        # left rail x for the climbing coin


def bar_y(pos):
    pos = clamp(pos, 0, TOP); i = min(int(pos), TOP - 1); frac = pos - i
    return lerp(BY[i], BY[i + 1], ease_in_out(frac))


def _platform(lit):
    pad = 30
    im = Image.new("RGBA", (BAR_W + 2 * pad, BAR_H + 2 * pad), (0, 0, 0, 0))
    sh = Image.new("RGBA", im.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle([pad, pad + 10, pad + BAR_W, pad + BAR_H + 10],
                                         radius=36, fill=(0, 0, 0, 150))
    im.alpha_composite(sh.filter(ImageFilter.GaussianBlur(16)))
    d = ImageDraw.Draw(im)
    box = [pad, pad, pad + BAR_W, pad + BAR_H]
    if lit:
        d.rounded_rectangle(box, radius=36, fill=(*GOLD, 255))
        gl = Image.new("RGBA", im.size, (0, 0, 0, 0))                          # gloss (composited)
        ImageDraw.Draw(gl).rounded_rectangle([box[0] + 8, box[1] + 6, box[2] - 8, box[1] + BAR_H // 2],
                                             radius=28, fill=(255, 255, 255, 55))
        im.alpha_composite(gl)
    else:
        d.rounded_rectangle(box, radius=36, fill=(40, 58, 96, 255))
        ol = Image.new("RGBA", im.size, (0, 0, 0, 0))
        ImageDraw.Draw(ol).rounded_rectangle(box, radius=36, outline=(*GOLD, 110), width=2)
        im.alpha_composite(ol)
    return im, pad

PLAT_LIT, PLAT_PAD = _platform(True)
PLAT_UNLIT, _ = _platform(False)


def _coin():
    pad = 22
    S = COIN_R * 2
    im = Image.new("RGBA", (S + 2 * pad, S + 2 * pad), (0, 0, 0, 0))
    sh = Image.new("RGBA", im.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).ellipse([pad, pad + 6, pad + S, pad + S + 6], fill=(0, 0, 0, 150))
    im.alpha_composite(sh.filter(ImageFilter.GaussianBlur(10)))
    d = ImageDraw.Draw(im)
    d.ellipse([pad, pad, pad + S, pad + S], fill=(*GOLD, 255))
    ov = Image.new("RGBA", im.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)      # ring + gloss
    od.ellipse([pad + 7, pad + 7, pad + S - 7, pad + S - 7], outline=(255, 255, 255, 150), width=3)
    od.ellipse([pad + 5, pad + 5, pad + S * 0.7, pad + S * 0.5], fill=(255, 255, 255, 46))
    im.alpha_composite(ov)
    f = font(int(COIN_R * 1.15), bold=True)
    bb = d.textbbox((0, 0), "$", font=f)
    d.text((pad + (S - (bb[2] - bb[0])) / 2 - bb[0], pad + (S - (bb[3] - bb[1])) / 2 - bb[1]),
           "$", font=f, fill=(44, 34, 10, 255))
    return im

COIN = _coin()


def _star(size, col):
    im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    cx = cy = size / 2
    R, r = size * 0.48, size * 0.20
    pts = []
    for i in range(10):
        ang = -math.pi / 2 + i * math.pi / 5
        rad = R if i % 2 == 0 else r
        pts.append((cx + rad * math.cos(ang), cy + rad * math.sin(ang)))
    d.polygon(pts, fill=(*col, 255))
    return im

STAR_GOLD = _star(46, GOLD)
STAR_DARK = _star(46, (54, 42, 12))
STAR_BURST = _star(84, GOLD)


def paste(frame, sprite, cx, cy, a=1.0):
    if a <= 0: return
    img = sprite
    if a < 1:
        img = sprite.copy(); img.putalpha(sprite.split()[3].point(lambda p: int(p * clamp(a))))
    frame.alpha_composite(img, (int(cx - sprite.width / 2), int(cy - sprite.height / 2)))


def draw_check(layer, cx, cy, col, a=1.0):
    d = ImageDraw.Draw(layer)
    c = (*col, int(255 * clamp(a)))
    d.line([(cx - 15, cy), (cx - 4, cy + 12)], fill=c, width=8)
    d.line([(cx - 4, cy + 12), (cx + 17, cy - 14)], fill=c, width=8)


def draw_clock(layer, cx, cy, a=1.0):
    d = ImageDraw.Draw(layer)
    r = 32
    col = (*RED, int(255 * clamp(a)))
    d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=col, width=6)
    d.line([(cx, cy), (cx, cy - 18)], fill=col, width=6)
    d.line([(cx, cy), (cx + 13, cy + 7)], fill=col, width=6)


# ── The ladder, drawn onto its own RGBA layer ───────────────────────────────────
def ladder_layer(pos, coin_xy=None, reveal=1.0, burst_a=0.0):
    """pos: continuous climb index 0..4 → which bars are lit (reached). coin_xy
    overrides the coin position; otherwise it sits by rung round(pos). reveal fades
    the whole structure in; burst_a fires celebratory stars at the top bar."""
    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    lit = [k == 0 or pos >= k - 0.12 for k in range(NLV)]
    af = font(50, bold=True)
    tf = font(26, bold=True)
    bx0, bx1 = BXC - BAR_W / 2, BXC + BAR_W / 2
    for k in range(NLV):
        a = reveal * (1.0 if lit[k] else 0.30)
        paste(layer, PLAT_LIT if lit[k] else PLAT_UNLIT, BXC, BY[k], a)
        ink = (34, 28, 10) if lit[k] else SLATE
        # left marker: dark ✓ when reached, faint gold ★ when still locked
        if lit[k]:
            draw_check(layer, bx0 + 56, BY[k], (38, 94, 52), reveal)
        else:
            paste(layer, STAR_GOLD, bx0 + 56, BY[k], reveal * 0.5)
        d.text((bx0 + 116, BY[k]), money(AMOUNTS[k]), font=af,
               fill=(*ink, int(255 * a)), anchor="lm")
        tag = "inicial" if k == 0 else ("tope" if k == TOP else "")
        if tag:
            d.text((bx1 - 52, BY[k]), tag, font=tf, fill=(*ink, int(230 * a)), anchor="rm")
        elif lit[k]:
            paste(layer, STAR_DARK, bx1 - 54, BY[k], reveal * 0.85)
    # coin on the left rail
    if coin_xy is None:
        coin_xy = (COIN_X, BY[int(round(clamp(pos, 0, TOP)))])
    paste(layer, COIN, coin_xy[0], coin_xy[1], reveal)
    # celebratory burst at the top bar
    if burst_a > 0:
        for (dx, dy) in [(-150, -16), (160, -10), (-80, -84), (100, -92),
                         (0, -120), (225, -48), (-225, -44)]:
            sc = 0.5 + 0.9 * burst_a
            s = STAR_BURST.resize((max(1, int(STAR_BURST.width * sc)),) * 2, Image.LANCZOS)
            paste(layer, s, BXC + dx * burst_a, BY[TOP] + dy - 10 * burst_a,
                  clamp(1.4 - burst_a) * 0.85)
    return layer


def climb_coin_xy(pos):
    i = min(int(pos), TOP - 1); frac = pos - i
    x = COIN_X - 40 * math.sin(math.pi * frac)      # arc out to the left on each hop
    return (x, bar_y(pos))


def footer(frame, a, climb=True):
    d = ImageDraw.Draw(frame)
    if climb:
        pill = "Pagá a tiempo"
        pf = font(40, bold=True)
        pw = d.textlength(pill, font=pf) + 64
        x0 = W / 2 - pw / 2
        d.rounded_rectangle([x0, 1636, x0 + pw, 1636 + 70], radius=35, fill=(*GOLD, int(235 * a)))
        d.text((W / 2 - d.textlength(pill, font=pf) / 2, 1650), pill, font=pf, fill=(20, 18, 10, int(255 * a)))
    ctext(d, W / 2, 1740, "+$50.000 por nivel  ·  tope $500.000", font(32, bold=True), SLATE, a)


# ── Phases ──────────────────────────────────────────────────────────────────────
def render_intro(p):
    frame = BG.copy(); d = ImageDraw.Draw(frame)
    la = ease_out_cubic(p / 0.28)
    ctext(d, W / 2, 360, "E S C A L E R A   D E   N I V E L E S", font(36, bold=True), GOLD, la, tracking=2)
    pop = ease_out_back((p - 0.22) / 0.45)
    ta = ease_out_cubic((p - 0.22) / 0.30)
    big = font(int(lerp(120, 150, pop)), bold=True)
    ctext(d, W / 2, lerp(560, 520, pop), "Subí de nivel", big, WHITE, ta)
    ctext(d, W / 2, 760, "Pagá a tiempo y accedé a más", font(46), SLATE, clamp((p - 0.45) / 0.3))
    # faint ladder teasing in at the bottom
    lad = ladder_layer(0, reveal=clamp((p - 0.5) / 0.4) * 0.5)
    frame.alpha_composite(lad)
    fo = clamp((p - 0.9) / 0.1)
    if fo > 0:
        frame.alpha_composite(Image.new("RGBA", (W, H), (*NAVY, int(120 * fo))))
    return frame


def render_climb(p):
    frame = BG.copy(); d = ImageDraw.Draw(frame)
    cp = clamp((p - 0.05) / 0.9)
    pos = cp * 4
    ctext(d, W / 2, 180, "Pagá a tiempo y subí", font(78, bold=True), WHITE)
    ctext(d, W / 2, 288, "cada préstamo devuelto en fecha te sube de nivel", font(34), SLATE)
    lad = ladder_layer(pos, coin_xy=climb_coin_xy(pos), reveal=1.0)
    frame.alpha_composite(lad)
    footer(frame, clamp(p / 0.12))
    return frame


def render_reward(p):
    frame = BG.copy(); d = ImageDraw.Draw(frame)
    lad = ladder_layer(4, reveal=1.0, burst_a=ease_out_cubic(p / 0.6))
    frame.alpha_composite(lad)
    ta = ease_out_cubic(p / 0.3)
    ctext(d, W / 2, 170, "Hasta " + money(500000), font(96, bold=True), GOLD, ta)
    ctext(d, W / 2, 300, "Accedé a montos más altos", font(48, bold=True), WHITE, clamp((p - 0.2) / 0.3))
    footer(frame, 1.0, climb=False)
    return frame


def render_warn(p):
    frame = BG.copy(); d = ImageDraw.Draw(frame)
    pos = lerp(4, 3, ease_in_out(clamp((p - 0.15) / 0.7)))
    cy = bar_y(pos)
    lad = ladder_layer(pos, coin_xy=(COIN_X, cy), reveal=1.0)
    frame.alpha_composite(lad)
    # red clock above the slipping coin + caution headline
    draw_clock(frame, COIN_X, cy - 104, ease_out_cubic(p / 0.3))
    ta = ease_out_cubic(p / 0.25)
    ctext(d, W / 2, 180, "¿Te atrasás?", font(82, bold=True), WHITE, ta)
    ctext(d, W / 2, 300, "Bajás de nivel", font(66, bold=True), RED, clamp((p - 0.15) / 0.25))
    ctext(d, W / 2, 1740, "Mantené tu nivel pagando en fecha", font(34, bold=True), SLATE, ta)
    return frame


BADGE_MASTER = ic_badge(560)
def paste_badge(frame, cx, cy, disp, a=1.0):
    scale = disp / 560.0
    b = BADGE_MASTER.resize((int(BADGE_MASTER.width * scale), int(BADGE_MASTER.height * scale)), Image.LANCZOS)
    paste(frame, b, cx, cy, a)

def wordmark(d, cx, y, size, a=1.0):
    f = font(size, bold=True)
    w1 = d.textlength("Impulso ", font=f); w2 = d.textlength("Crédito", font=f)
    x = cx - (w1 + w2) / 2
    d.text((x, y), "Impulso ", font=f, fill=(*BLUE, int(255 * clamp(a))))
    d.text((x + w1, y), "Crédito", font=f, fill=(*GOLD, int(255 * clamp(a))))


def render_outro(p):
    frame = BG.copy(); d = ImageDraw.Draw(frame)
    # ladder dissolves, navy wipes it clean
    lad_a = clamp(1 - p / 0.28)
    if lad_a > 0:
        frame.alpha_composite(ladder_layer(4, reveal=lad_a))
    ov = clamp((p - 0.10) / 0.18)
    if ov > 0:
        frame.alpha_composite(Image.new("RGBA", (W, H), (*NAVY, int(255 * ov))))
    bp = ease_out_back((p - 0.24) / 0.34)
    ba = ease_out_cubic((p - 0.24) / 0.26)
    paste_badge(frame, W / 2, 700, lerp(320, 500, bp), ba)
    wordmark(d, W / 2, 980, 96, clamp((p - 0.42) / 0.16))
    ctext(d, W / 2, 1110, "Pagá a tiempo y crecé con nosotros", font(44), SLATE,
          clamp((p - 0.5) / 0.16), tracking=1)
    yb = 1260
    d.rounded_rectangle([W / 2 - 40, yb, W / 2 + 40, yb + 9], radius=4,
                        fill=(*GOLD, int(255 * clamp((p - 0.55) / 0.15))))
    ctext(d, W / 2, yb + 48, "Consultá por WhatsApp", font(52, bold=True), WHITE, clamp((p - 0.58) / 0.16))
    ctext(d, W / 2, yb + 132, "@impulso.credito.ars", font(44, bold=True), GOLD, clamp((p - 0.64) / 0.16))
    ctext(d, W / 2, yb + 206, "Zona Norte + AMBA", font(34), SLATE, clamp((p - 0.7) / 0.16))
    return frame


# ── Assemble ────────────────────────────────────────────────────────────────────
SECS = [("intro", 2.2, render_intro), ("climb", 7.0, render_climb),
        ("reward", 1.8, render_reward), ("warn", 1.5, render_warn),
        ("outro", 2.5, render_outro)]


def build():
    out = HERE / "story_niveles.mp4"
    spans = [(name, int(round(dur * FPS)), fn) for name, dur, fn in SECS]
    total = sum(n for _, n, _ in spans)
    previews = {"climb": "nivpreview_climb.png", "reward": "nivpreview_reward.png",
                "outro": "nivpreview_outro.png"}
    writer = imageio.get_writer(out, fps=FPS, codec="libx264", macro_block_size=8,
                                pixelformat="yuv420p", ffmpeg_params=["-crf", "20", "-preset", "medium"])
    i = 0
    for name, n, fn in spans:
        for k in range(n):
            fr = fn(k / max(1, n - 1)).convert("RGB")
            if name in previews and k == n // 2:
                fr.save(HERE / previews[name], "PNG")
            writer.append_data(np.asarray(fr))
            i += 1
    writer.close()
    print(f"Wrote {out}  ({total} frames, {total / FPS:.1f}s)")


if __name__ == "__main__":
    build()
