#!/usr/bin/env python3
"""
Impulso Crédito — 30 feature posts (1080×1080) with custom icons + graphics.

Same navy/blue/gold brand as the base posts, but built for ads: a bold focal
icon medallion, accent shapes/dots, strong hierarchy, gold CTA. Writes
feat_01.png … feat_30.png + matching feat_XX.txt captions and feature_captions.md.

Run:  python3 generate_feature_posts.py
"""

import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

from generate_assets import (font, vgradient, ic_badge, BG_TOP, BG_BOT,
                             BLUE, GOLD, WHITE, SLATE)
from generate_captions import CONTACT, HASHTAGS

HERE = Path(__file__).parent
S = 1080
NAVY_FILL = (18, 27, 45)


# ── low-level stroke helper (rounded caps + smooth joints) ───────────────────
def stroke(d, pts, col, w):
    if len(pts) >= 2:
        d.line(pts, fill=col, width=w, joint="curve")
    r = max(1, w // 2)
    for x, y in (pts[0], pts[-1]):
        d.ellipse([x - r, y - r, x + r, y + r], fill=col)


# ── icon library (drawn on an N×N canvas, center = N/2) ──────────────────────
def _g(N):                       # common geometry: center, half-size, stroke width
    return N / 2, N / 2, N * 0.34, max(6, int(N * 0.055))

def i_clock(d, N, c):
    cx, cy, s, w = _g(N)
    d.ellipse([cx - s, cy - s, cx + s, cy + s], outline=c, width=w)
    stroke(d, [(cx, cy), (cx, cy - s * 0.55)], c, w)
    stroke(d, [(cx, cy), (cx + s * 0.4, cy + s * 0.15)], c, w)

def i_bolt(d, N, c):
    cx, cy, s, w = _g(N)
    d.polygon([(cx + 0.15 * s, cy - s), (cx - 0.55 * s, cy + 0.12 * s),
               (cx - 0.05 * s, cy + 0.12 * s), (cx - 0.2 * s, cy + s),
               (cx + 0.6 * s, cy - 0.25 * s), (cx + 0.08 * s, cy - 0.25 * s)], fill=c)

def i_phone(d, N, c):
    cx, cy, s, w = _g(N)
    d.rounded_rectangle([cx - 0.6 * s, cy - s, cx + 0.6 * s, cy + s], radius=int(s * 0.22), outline=c, width=w)
    stroke(d, [(cx - 0.15 * s, cy - 0.78 * s), (cx + 0.15 * s, cy - 0.78 * s)], c, w)
    d.ellipse([cx - 0.08 * s, cy + 0.68 * s, cx + 0.08 * s, cy + 0.84 * s], fill=c)

def i_chat(d, N, c):
    cx, cy, s, w = _g(N)
    d.rounded_rectangle([cx - s, cy - 0.8 * s, cx + s, cy + 0.45 * s], radius=int(s * 0.32), outline=c, width=w)
    d.polygon([(cx - 0.35 * s, cy + 0.4 * s), (cx - 0.05 * s, cy + 0.4 * s), (cx - 0.35 * s, cy + 0.9 * s)], fill=c)
    for dx in (-0.4, 0, 0.4):
        d.ellipse([cx + dx * s - 0.08 * s, cy - 0.28 * s, cx + dx * s + 0.08 * s, cy - 0.12 * s], fill=c)

def i_doc(d, N, c):
    cx, cy, s, w = _g(N)
    d.rounded_rectangle([cx - 0.6 * s, cy - s, cx + 0.6 * s, cy + s], radius=int(s * 0.14), outline=c, width=w)
    for i, fy in enumerate((-0.45, -0.1, 0.25, 0.6)):
        xr = 0.35 if i == 3 else 0.4
        stroke(d, [(cx - 0.35 * s, cy + fy * s), (cx + xr * s, cy + fy * s)], c, int(w * 0.8))

def i_pen(d, N, c):
    cx, cy, s, w = _g(N)
    d.rounded_rectangle([cx - 0.75 * s, cy - s, cx + 0.35 * s, cy + s], radius=int(s * 0.14), outline=c, width=w)
    for fy in (-0.5, -0.18):
        stroke(d, [(cx - 0.5 * s, cy + fy * s), (cx + 0.1 * s, cy + fy * s)], c, int(w * 0.8))
    stroke(d, [(cx - 0.1 * s, cy + 0.9 * s), (cx + 0.9 * s, cy - 0.1 * s)], c, w)   # pen body
    d.polygon([(cx + 0.9 * s, cy - 0.1 * s), (cx + 1.0 * s, cy + 0.1 * s), (cx + 0.75 * s, cy + 0.15 * s)], fill=c)

def i_id(d, N, c):
    cx, cy, s, w = _g(N)
    d.rounded_rectangle([cx - s, cy - 0.65 * s, cx + s, cy + 0.65 * s], radius=int(s * 0.18), outline=c, width=w)
    d.ellipse([cx - 0.7 * s, cy - 0.3 * s, cx - 0.15 * s, cy + 0.25 * s], outline=c, width=int(w * 0.8))
    for fy in (-0.18, 0.1, 0.35):
        stroke(d, [(cx + 0.05 * s, cy + fy * s), (cx + 0.7 * s, cy + fy * s)], c, int(w * 0.7))

def i_check(d, N, c):
    cx, cy, s, w = _g(N)
    d.ellipse([cx - s, cy - s, cx + s, cy + s], outline=c, width=w)
    stroke(d, [(cx - 0.42 * s, cy + 0.02 * s), (cx - 0.1 * s, cy + 0.38 * s), (cx + 0.48 * s, cy - 0.38 * s)], c, w)

def i_refresh(d, N, c):
    cx, cy, s, w = _g(N)
    d.arc([cx - s, cy - s, cx + s, cy + s], 40, 300, fill=c, width=w)
    ax, ay = cx + s * math.cos(math.radians(40)), cy + s * math.sin(math.radians(40))
    d.polygon([(ax, ay), (ax - 0.3 * s, ay - 0.05 * s), (ax + 0.02 * s, ay + 0.32 * s)], fill=c)

def i_calendar(d, N, c):
    cx, cy, s, w = _g(N)
    d.rounded_rectangle([cx - s, cy - 0.8 * s, cx + s, cy + s], radius=int(s * 0.14), outline=c, width=w)
    stroke(d, [(cx - s + w, cy - 0.35 * s), (cx + s - w, cy - 0.35 * s)], c, int(w * 0.8))
    stroke(d, [(cx - 0.55 * s, cy - s), (cx - 0.55 * s, cy - 0.7 * s)], c, w)
    stroke(d, [(cx + 0.55 * s, cy - s), (cx + 0.55 * s, cy - 0.7 * s)], c, w)
    for ry in (0.1, 0.5):
        for rx in (-0.5, 0, 0.5):
            d.ellipse([cx + rx * s - 0.07 * s, cy + ry * s - 0.07 * s, cx + rx * s + 0.07 * s, cy + ry * s + 0.07 * s], fill=c)

def i_coins(d, N, c):
    cx, cy, s, w = _g(N)
    for k, oy in enumerate((0.55, 0.15, -0.25)):
        top = cy + oy * s
        d.ellipse([cx - 0.8 * s, top - 0.18 * s, cx + 0.8 * s, top + 0.18 * s], outline=c, width=int(w * 0.8))

def i_chart(d, N, c):
    cx, cy, s, w = _g(N)
    base = cy + 0.8 * s
    for i, h in enumerate((0.5, 0.9, 1.3)):
        x = cx - 0.6 * s + i * 0.6 * s
        d.rounded_rectangle([x - 0.16 * s, base - h * s, x + 0.16 * s, base], radius=4, fill=c)
    stroke(d, [(cx - 0.7 * s, cy - 0.1 * s), (cx - 0.1 * s, cy - 0.5 * s), (cx + 0.2 * s, cy - 0.3 * s), (cx + 0.8 * s, cy - 0.9 * s)], c, int(w * 0.8))

def i_search(d, N, c):
    cx, cy, s, w = _g(N)
    r = 0.6 * s
    d.ellipse([cx - 0.9 * s, cy - 0.9 * s, cx - 0.9 * s + 2 * r, cy - 0.9 * s + 2 * r], outline=c, width=w)
    stroke(d, [(cx - 0.9 * s + 2 * r * 0.78, cy - 0.9 * s + 2 * r * 0.78), (cx + 0.85 * s, cy + 0.85 * s)], c, int(w * 1.2))

def i_shield(d, N, c):
    cx, cy, s, w = _g(N)
    d.line([(cx, cy - s), (cx - 0.8 * s, cy - 0.6 * s), (cx - 0.8 * s, cy + 0.15 * s),
            (cx, cy + s), (cx + 0.8 * s, cy + 0.15 * s), (cx + 0.8 * s, cy - 0.6 * s), (cx, cy - s)], fill=c, width=w, joint="curve")
    stroke(d, [(cx - 0.38 * s, cy), (cx - 0.08 * s, cy + 0.32 * s), (cx + 0.42 * s, cy - 0.3 * s)], c, w)

def i_lock(d, N, c):
    cx, cy, s, w = _g(N)
    d.rounded_rectangle([cx - 0.7 * s, cy - 0.15 * s, cx + 0.7 * s, cy + 0.9 * s], radius=int(s * 0.18), outline=c, width=w)
    d.arc([cx - 0.45 * s, cy - 0.8 * s, cx + 0.45 * s, cy + 0.1 * s], 180, 360, fill=c, width=w)
    d.ellipse([cx - 0.1 * s, cy + 0.25 * s, cx + 0.1 * s, cy + 0.45 * s], fill=c)

def i_key(d, N, c):
    cx, cy, s, w = _g(N)
    d.ellipse([cx - 0.95 * s, cy - 0.45 * s, cx - 0.1 * s, cy + 0.4 * s], outline=c, width=w)
    stroke(d, [(cx - 0.12 * s, cy - 0.02 * s), (cx + 0.9 * s, cy - 0.02 * s)], c, w)
    stroke(d, [(cx + 0.6 * s, cy - 0.02 * s), (cx + 0.6 * s, cy + 0.3 * s)], c, w)
    stroke(d, [(cx + 0.85 * s, cy - 0.02 * s), (cx + 0.85 * s, cy + 0.25 * s)], c, w)

def i_pin(d, N, c):
    cx, cy, s, w = _g(N)
    d.line([(cx, cy + s), (cx - 0.7 * s, cy - 0.2 * s)], fill=c, width=w)
    d.line([(cx, cy + s), (cx + 0.7 * s, cy - 0.2 * s)], fill=c, width=w)
    d.ellipse([cx - 0.7 * s, cy - 0.9 * s, cx + 0.7 * s, cy + 0.1 * s], outline=c, width=w)
    d.ellipse([cx - 0.22 * s, cy - 0.55 * s, cx + 0.22 * s, cy - 0.11 * s], fill=c)

def i_heart(d, N, c):
    cx, cy, s, w = _g(N)
    d.ellipse([cx - 0.75 * s, cy - 0.7 * s, cx - 0.02 * s, cy + 0.05 * s], fill=c)
    d.ellipse([cx + 0.02 * s, cy - 0.7 * s, cx + 0.75 * s, cy + 0.05 * s], fill=c)
    d.polygon([(cx - 0.72 * s, cy - 0.18 * s), (cx + 0.72 * s, cy - 0.18 * s), (cx, cy + 0.85 * s)], fill=c)

def i_bell(d, N, c):
    cx, cy, s, w = _g(N)
    d.line([(cx - 0.7 * s, cy + 0.5 * s), (cx - 0.55 * s, cy + 0.1 * s), (cx - 0.55 * s, cy - 0.2 * s)], fill=c, width=w, joint="curve")
    d.arc([cx - 0.55 * s, cy - 0.85 * s, cx + 0.55 * s, cy + 0.25 * s], 180, 360, fill=c, width=w)
    d.line([(cx + 0.55 * s, cy - 0.2 * s), (cx + 0.55 * s, cy + 0.1 * s), (cx + 0.7 * s, cy + 0.5 * s)], fill=c, width=w, joint="curve")
    stroke(d, [(cx - 0.7 * s, cy + 0.5 * s), (cx + 0.7 * s, cy + 0.5 * s)], c, int(w * 0.8))
    d.ellipse([cx - 0.14 * s, cy + 0.55 * s, cx + 0.14 * s, cy + 0.82 * s], fill=c)

def i_star(d, N, c):
    cx, cy, s, _ = _g(N)
    pts = []
    for i in range(10):
        ang = -math.pi / 2 + i * math.pi / 5
        r = s if i % 2 == 0 else s * 0.42
        pts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
    d.polygon(pts, fill=c)

def i_percent(d, N, c):
    cx, cy, s, w = _g(N)
    stroke(d, [(cx + 0.7 * s, cy - 0.7 * s), (cx - 0.7 * s, cy + 0.7 * s)], c, w)
    d.ellipse([cx - 0.75 * s, cy - 0.75 * s, cx - 0.25 * s, cy - 0.25 * s], outline=c, width=int(w * 0.9))
    d.ellipse([cx + 0.25 * s, cy + 0.25 * s, cx + 0.75 * s, cy + 0.75 * s], outline=c, width=int(w * 0.9))

def i_sliders(d, N, c):
    cx, cy, s, w = _g(N)
    for i, (fy, kx) in enumerate(((-0.55, 0.35), (0.0, -0.3), (0.55, 0.15))):
        stroke(d, [(cx - 0.8 * s, cy + fy * s), (cx + 0.8 * s, cy + fy * s)], c, int(w * 0.7))
        d.ellipse([cx + kx * s - 0.16 * s, cy + fy * s - 0.16 * s, cx + kx * s + 0.16 * s, cy + fy * s + 0.16 * s], fill=c)

def i_users(d, N, c):
    cx, cy, s, w = _g(N)
    for ox in (-0.42, 0.42):
        d.ellipse([cx + ox * s - 0.3 * s, cy - 0.7 * s, cx + ox * s + 0.3 * s, cy - 0.1 * s], outline=c, width=w)
        d.arc([cx + ox * s - 0.5 * s, cy, cx + ox * s + 0.5 * s, cy + 1.2 * s], 180, 360, fill=c, width=w)

def i_rocket(d, N, c):
    cx, cy, s, w = _g(N)
    d.line([(cx, cy - s), (cx - 0.4 * s, cy - 0.1 * s), (cx - 0.4 * s, cy + 0.5 * s),
            (cx + 0.4 * s, cy + 0.5 * s), (cx + 0.4 * s, cy - 0.1 * s), (cx, cy - s)], fill=c, width=w, joint="curve")
    d.ellipse([cx - 0.17 * s, cy - 0.45 * s, cx + 0.17 * s, cy - 0.11 * s], fill=c)
    d.polygon([(cx - 0.4 * s, cy + 0.1 * s), (cx - 0.75 * s, cy + 0.5 * s), (cx - 0.4 * s, cy + 0.5 * s)], fill=c)
    d.polygon([(cx + 0.4 * s, cy + 0.1 * s), (cx + 0.75 * s, cy + 0.5 * s), (cx + 0.4 * s, cy + 0.5 * s)], fill=c)
    stroke(d, [(cx, cy + 0.5 * s), (cx, cy + 0.95 * s)], c, w)

def i_wallet(d, N, c):
    cx, cy, s, w = _g(N)
    d.rounded_rectangle([cx - 0.85 * s, cy - 0.6 * s, cx + 0.85 * s, cy + 0.7 * s], radius=int(s * 0.16), outline=c, width=w)
    d.rounded_rectangle([cx + 0.2 * s, cy - 0.05 * s, cx + 0.95 * s, cy + 0.35 * s], radius=int(s * 0.1), fill=c)

def i_bank(d, N, c):
    cx, cy, s, w = _g(N)
    d.polygon([(cx, cy - 0.9 * s), (cx + 0.9 * s, cy - 0.35 * s), (cx - 0.9 * s, cy - 0.35 * s)], fill=c)
    for rx in (-0.55, 0, 0.55):
        stroke(d, [(cx + rx * s, cy - 0.2 * s), (cx + rx * s, cy + 0.6 * s)], c, w)
    stroke(d, [(cx - 0.9 * s, cy + 0.8 * s), (cx + 0.9 * s, cy + 0.8 * s)], c, w)

ICONS = {
    "clock": i_clock, "bolt": i_bolt, "phone": i_phone, "chat": i_chat, "doc": i_doc,
    "pen": i_pen, "id": i_id, "check": i_check, "refresh": i_refresh, "calendar": i_calendar,
    "coins": i_coins, "chart": i_chart, "search": i_search, "shield": i_shield, "lock": i_lock,
    "key": i_key, "pin": i_pin, "heart": i_heart, "bell": i_bell, "star": i_star,
    "percent": i_percent, "sliders": i_sliders, "users": i_users, "rocket": i_rocket,
    "wallet": i_wallet, "bank": i_bank,
}

def make_icon(name, disp, col):
    ss = 2; N = disp * ss
    im = Image.new("RGBA", (N, N), (0, 0, 0, 0))
    ICONS[name](ImageDraw.Draw(im), N, col)
    return im.resize((disp, disp), Image.LANCZOS)


# ── shared drawing helpers ───────────────────────────────────────────────────
def circle_gradient(size, top, bot):
    grad = vgradient(size, size, top, bot).convert("RGB")
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).ellipse([0, 0, size - 1, size - 1], fill=255)
    out = Image.new("RGBA", (size, size), (0, 0, 0, 0)); out.paste(grad, (0, 0), mask)
    return out

def wrap_px(draw, text, f, max_w):
    words, lines, cur = text.split(), [], ""
    for wd in words:
        t = (cur + " " + wd).strip()
        if draw.textlength(t, font=f) <= max_w:
            cur = t
        else:
            lines.append(cur); cur = wd
    if cur:
        lines.append(cur)
    return lines

def ctext(draw, cx, y, text, f, fill, tracking=0):
    if tracking:
        total = sum(draw.textlength(ch, font=f) for ch in text) + tracking * (len(text) - 1)
        x = cx - total / 2
        for ch in text:
            draw.text((x, y), ch, font=f, fill=fill); x += draw.textlength(ch, font=f) + tracking
    else:
        draw.text((cx - draw.textlength(text, font=f) / 2, y), text, font=f, fill=fill)


# ── static background (navy + decorative graphics), built once ───────────────
def _base():
    img = vgradient(S, S, BG_TOP, BG_BOT)
    deco = Image.new("RGBA", (S, S), (0, 0, 0, 0)); dd = ImageDraw.Draw(deco)
    dd.ellipse([-260, -320, 360, 300], fill=(70, 110, 180, 55))        # top-left blue glow
    dd.ellipse([720, 760, 1340, 1380], fill=(233, 170, 70, 40))        # bottom-right gold glow
    img.alpha_composite(deco.filter(ImageFilter.GaussianBlur(120)))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, S, 8], fill=GOLD)
    d.rectangle([0, S - 8, S, S], fill=GOLD)
    # faint accent dots, kept clear of the text/medallion zone
    for (x, y, r) in [(92, 150, 9), (150, 205, 6), (986, 158, 9), (928, 212, 6)]:
        d.ellipse([x - r, y - r, x + r, y + r], fill=(233, 170, 70, 110))
    return img

BASE = _base()
BADGE = ic_badge(80)


# ── one feature post ─────────────────────────────────────────────────────────
def make_post(i, icon, label, headline, support):
    img = BASE.copy(); d = ImageDraw.Draw(img)

    # header: IC badge + wordmark
    wm = font(46, bold=True)
    ww = d.textlength("Impulso ", font=wm) + d.textlength("Crédito", font=wm)
    gw = BADGE.width + 16 + ww; x0 = (S - gw) / 2
    img.alpha_composite(BADGE, (int(x0), 74))
    d.text((x0 + BADGE.width + 16, 92), "Impulso ", font=wm, fill=BLUE)
    d.text((x0 + BADGE.width + 16 + d.textlength("Impulso ", font=wm), 92), "Crédito", font=wm, fill=GOLD)
    d.text((S - 96, 74), f"{i:02d}/30", font=font(26, bold=True), fill=SLATE)

    # focal medallion with the icon
    mc, mr = 405, 128
    glow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse([S / 2 - mr - 30, mc - mr - 30, S / 2 + mr + 30, mc + mr + 30], fill=(233, 170, 70, 60))
    img.alpha_composite(glow.filter(ImageFilter.GaussianBlur(50)))
    med = circle_gradient(mr * 2, (56, 90, 152), (26, 44, 84))
    img.alpha_composite(med, (int(S / 2 - mr), mc - mr))
    d.ellipse([S / 2 - mr, mc - mr, S / 2 + mr, mc + mr], outline=GOLD, width=7)
    ic = make_icon(icon, 148, WHITE)
    img.alpha_composite(ic, (int(S / 2 - ic.width / 2), int(mc - ic.height / 2)))
    # orbiting accent dots
    for ang, rr in ((-40, 7), (20, 10), (210, 8), (140, 6)):
        ax = S / 2 + (mr + 24) * math.cos(math.radians(ang)); ay = mc + (mr + 24) * math.sin(math.radians(ang))
        d.ellipse([ax - rr, ay - rr, ax + rr, ay + rr], fill=GOLD)

    # label pill + headline + support, centred as one block between medallion & footer
    lf, hf, sf = font(30, bold=True), font(64, bold=True), font(35)
    hlines = wrap_px(d, headline, hf, S - 140)
    slines = wrap_px(d, support, sf, S - 150)
    lh, slh = 74, 47
    block_h = 52 + 24 + len(hlines) * lh + 14 + len(slines) * slh
    region_top, region_bot = 565, 898
    y = region_top + max(0, (region_bot - region_top - block_h) / 2)

    lw = d.textlength(label.upper(), font=lf) + 54
    d.rounded_rectangle([S / 2 - lw / 2, y, S / 2 + lw / 2, y + 52], radius=26, outline=GOLD, width=3)
    ctext(d, S / 2, y + 9, label.upper(), lf, GOLD, tracking=2)
    y += 52 + 24
    for ln in hlines:
        ctext(d, S / 2, y, ln, hf, WHITE); y += lh
    y += 14
    for ln in slines:
        ctext(d, S / 2, y, ln, sf, SLATE); y += slh

    # footer CTA (no divider line — it was overlapping the text)
    ctext(d, S / 2, S - 150, "Consultá por WhatsApp", font(36, bold=True), WHITE)
    ctext(d, S / 2, S - 100, "@impulso.credito.ars   ·   Zona Norte + AMBA", font(29), SLATE)

    img.convert("RGB").save(HERE / f"feat_{i:02d}.png", "PNG")


# ── the 30 features (icon, label, headline, support, emoji) ──────────────────
FEATURES = [
    ("clock",   "Rapidez",      "Aprobación en minutos",    "Respuesta al toque, no en días.",               "⏱️"),
    ("bolt",    "Al instante",  "El dinero, hoy mismo",     "Aprobado hoy, transferido hoy.",                "⚡"),
    ("phone",   "Online",       "100% desde el celular",    "Pedí desde donde estés, sin moverte.",          "📱"),
    ("id",      "Fácil",        "Solo tu correo y DNI",     "Empezás tu solicitud en segundos.",             "🪪"),
    ("check",   "Simple",       "Cero papeleo",             "Sin trámites, sin filas, sin vueltas.",         "✅"),
    ("chat",    "Atención",     "Te atendemos por WhatsApp","Hablás con una persona real, al instante.",     "💬"),
    ("refresh", "Recurrentes",  "La segunda vez vuela",     "Si ya sos cliente, cargamos tus datos por vos.","🔁"),
    ("calendar","Plazos",       "Plazos flexibles",         "Elegí el plazo que mejor te quede.",            "📅"),
    ("coins",   "Cuotas",       "Pagá en cuotas",           "Montos grandes en cuotas manejables.",          "🪙"),
    ("chart",   "Niveles",      "Subí de nivel",            "Pagá a tiempo y accedé a montos más altos.",    "📈"),
    ("chart",   "Crecimiento",  "Tu límite crece",          "Cada préstamo bien pagado suma a tu historial.","🔝"),
    ("search",  "Claridad",     "Sin letra chica",          "Sabés cuánto y cuándo vas a pagar, siempre.",   "🔎"),
    ("wallet",  "Mercado Pago", "Cobrá por Mercado Pago",   "El dinero llega a tu cuenta al instante.",      "💳"),
    ("bank",    "Transferencia","O por transferencia",      "Recibí en tu cuenta bancaria, como prefieras.", "🏦"),
    ("pen",     "Contrato",     "Firma digital",            "Firmás tu contrato online, sin imprimir nada.", "✍️"),
    ("shield",  "Seguridad",    "Tus datos, protegidos",    "Un proceso seguro y confidencial.",             "🛡️"),
    ("pin",     "Cobertura",    "Zona Norte + AMBA",        "Estamos cerca tuyo, en Buenos Aires.",          "📍"),
    ("heart",   "Cercanía",     "Trato humano",             "Te tratamos por tu nombre, no por un número.",  "❤️"),
    ("bell",    "Avisos",       "Te avisamos tu cuota",     "Recordatorios para que nunca te atrases.",      "🔔"),
    ("star",    "Primera vez",  "Tu primer préstamo",       "También te acompañamos cuando empezás.",        "⭐"),
    ("percent", "Interés",      "Interés claro",            "Se define según el monto y el plazo elegido.",  "📊"),
    ("sliders", "A tu medida",  "Préstamos a tu medida",    "El plazo y las cuotas se ajustan a lo que pedís.","🎚️"),
    ("check",   "Proceso",      "6 pasos y listo",          "Una solicitud simple, guiada y rápida.",        "🧭"),
    ("shield",  "Confianza",    "Cumplimos lo que prometemos","Un prestamista en el que podés confiar.",     "🤝"),
    ("users",   "Referencias",  "Tus referencias cuentan",  "Dos contactos de confianza y listo.",           "👥"),
    ("lock",    "Privacidad",   "Datos confidenciales",     "Tu información está siempre protegida.",         "🔒"),
    ("coins",   "Montos",       "Montos que crecen",        "Desde montos chicos hasta más grandes.",        "💰"),
    ("key",     "Acceso",       "Tu préstamo, a un paso",   "Acceso simple, sin barreras ni demoras.",       "🔑"),
    ("rocket",  "Impulso",      "Dale impulso a tus planes","Hacé realidad eso que necesitás hoy.",          "🚀"),
    ("rocket",  "Empezá",       "Pedí el tuyo hoy",         "Tu próximo impulso está a un mensaje.",         "🚀"),
]


def main():
    md = ["# Impulso Crédito — posts de features (descripciones)\n"]
    for i, (icon, label, head, sup, emoji) in enumerate(FEATURES, start=1):
        make_post(i, icon, label, head, sup)
        body = f"{emoji} {head}. {sup}"
        caption = f"{body}\n\n{CONTACT}\n\n{HASHTAGS}\n"
        (HERE / f"feat_{i:02d}.txt").write_text(caption, encoding="utf-8")
        md.append(f"## feat_{i:02d}.png\n\n```\n{caption}```\n")
        print(f"feat_{i:02d}.png  ({icon})")
    (HERE / "feature_captions.md").write_text("\n".join(md), encoding="utf-8")
    print(f"Done — {len(FEATURES)} posts + captions in {HERE}")


if __name__ == "__main__":
    main()
