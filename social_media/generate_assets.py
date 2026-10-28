#!/usr/bin/env python3
"""
Impulso Crédito — social media asset generator.

Creates a polished profile image (profile.png) and 10 branded "fact" posts
(post_01.png … post_10.png), 1080×1080, in the Impulso Crédito look:
deep-navy background, blue "Impulso" + gold "Crédito", rounded IC monogram.

Run:  python3 generate_assets.py
"""

import textwrap
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

HERE = Path(__file__).parent
S = 1080  # square canvas

# ── Brand palette (from the app's look) ──────────────────────────────────────
BG_TOP   = (30, 48, 88)     # deep navy, top
BG_BOT   = (12, 19, 38)     # deeper navy, bottom
BADGE_TOP = (70, 110, 180)  # IC badge gradient (light)
BADGE_BOT = (34, 56, 100)   # IC badge gradient (dark)
BLUE  = (99, 150, 220)      # "Impulso"
GOLD  = (233, 170, 70)      # "Crédito" / accents
WHITE = (255, 255, 255)
SLATE = (150, 164, 189)     # muted text

FONT_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
FONT_REG  = "/System/Library/Fonts/Supplemental/Arial.ttf"
FALLBACK_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
FALLBACK_REG  = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"


def font(size, bold=False):
    for p in ([FONT_BOLD, FALLBACK_BOLD] if bold else [FONT_REG, FALLBACK_REG]):
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


# ── Helpers ──────────────────────────────────────────────────────────────────

def vgradient(w, h, top, bot):
    """Vertical gradient as an RGBA image."""
    base = Image.new("RGB", (w, h), top)
    draw = ImageDraw.Draw(base)
    for y in range(h):
        t = y / max(1, h - 1)
        c = tuple(int(top[i] + (bot[i] - top[i]) * t) for i in range(3))
        draw.line([(0, y), (w, y)], fill=c)
    return base.convert("RGBA")


def rounded_gradient(w, h, radius, top, bot):
    """Rounded-rectangle tile filled with a vertical gradient (RGBA)."""
    grad = vgradient(w, h, top, bot).convert("RGB")
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, w - 1, h - 1], radius, fill=255)
    out = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    out.paste(grad, (0, 0), mask)
    return out


def ic_badge(size, ring=True):
    """Polished rounded 'IC' monogram badge as an RGBA image of (size,size)."""
    pad = int(size * 0.14)                 # room for shadow/glow
    canvas = Image.new("RGBA", (size + pad * 2, size + pad * 2), (0, 0, 0, 0))

    # soft drop shadow
    sh = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle(
        [pad, pad + int(size * 0.05), pad + size, pad + size + int(size * 0.05)],
        radius=int(size * 0.24), fill=(0, 0, 0, 150))
    sh = sh.filter(ImageFilter.GaussianBlur(int(size * 0.05)))
    canvas.alpha_composite(sh)

    # badge body (gradient)
    badge = rounded_gradient(size, size, int(size * 0.24), BADGE_TOP, BADGE_BOT)
    # top gloss
    gloss = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ImageDraw.Draw(gloss).rounded_rectangle(
        [int(size * 0.08), int(size * 0.06), int(size * 0.92), int(size * 0.48)],
        radius=int(size * 0.2), fill=(255, 255, 255, 26))
    gloss = gloss.filter(ImageFilter.GaussianBlur(int(size * 0.03)))
    badge.alpha_composite(gloss)
    # gold ring
    if ring:
        ImageDraw.Draw(badge).rounded_rectangle(
            [2, 2, size - 3, size - 3], radius=int(size * 0.23),
            outline=GOLD, width=max(2, int(size * 0.02)))
    # "IC"
    d = ImageDraw.Draw(badge)
    f = font(int(size * 0.5), bold=True)
    bb = d.textbbox((0, 0), "IC", font=f)
    d.text(((size - (bb[2] - bb[0])) / 2 - bb[0],
            (size - (bb[3] - bb[1])) / 2 - bb[1]), "IC", font=f, fill=WHITE)

    canvas.alpha_composite(badge, (pad, pad))
    return canvas


def wrap_px(draw, text, f, max_w):
    words, lines, cur = text.split(), [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if draw.textlength(t, font=f) <= max_w:
            cur = t
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def ctext(draw, cx, y, text, f, fill, tracking=0):
    """Draw horizontally-centered text; returns the line height used."""
    if tracking:
        total = sum(draw.textlength(ch, font=f) for ch in text) + tracking * (len(text) - 1)
        x = cx - total / 2
        for ch in text:
            draw.text((x, y), ch, font=f, fill=fill)
            x += draw.textlength(ch, font=f) + tracking
    else:
        w = draw.textlength(text, font=f)
        draw.text((cx - w / 2, y), text, font=f, fill=fill)
    bb = draw.textbbox((0, 0), "Hg", font=f)
    return bb[3] - bb[1]


def wordmark(draw, cx, y, size):
    """Centered 'Impulso Crédito' wordmark (blue + gold)."""
    f = font(size, bold=True)
    w1 = draw.textlength("Impulso ", font=f)
    w2 = draw.textlength("Crédito", font=f)
    x = cx - (w1 + w2) / 2
    draw.text((x, y), "Impulso ", font=f, fill=BLUE)
    draw.text((x + w1, y), "Crédito", font=f, fill=GOLD)


def base_canvas():
    img = vgradient(S, S, BG_TOP, BG_BOT)
    d = ImageDraw.Draw(img)
    # thin gold accent bars, top & bottom
    d.rectangle([0, 0, S, 7], fill=GOLD)
    d.rectangle([0, S - 7, S, S], fill=GOLD)
    return img, d


# ── Profile image ────────────────────────────────────────────────────────────

def make_profile():
    img = vgradient(S, S, BG_TOP, BG_BOT)
    # subtle radial glow behind the badge
    glow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse([S * 0.18, S * 0.12, S * 0.82, S * 0.76],
                                 fill=(70, 110, 180, 70))
    img.alpha_composite(glow.filter(ImageFilter.GaussianBlur(70)))

    badge = ic_badge(int(S * 0.46))
    img.alpha_composite(badge, ((S - badge.width) // 2, int(S * 0.20)))

    d = ImageDraw.Draw(img)
    wordmark(d, S / 2, int(S * 0.70), 86)
    ctext(d, S / 2, int(S * 0.70) + 110, "Préstamos rápidos y confiables",
          font(34), SLATE, tracking=1)

    img.convert("RGB").save(HERE / "profile.png", "PNG")
    print("profile.png")


# ── Fact posts ───────────────────────────────────────────────────────────────

POSTS = [
    ("Velocidad",     "Dinero en 24 horas",            "Aprobamos tu préstamo rápido y te transferimos el mismo día."),
    ("100% online",   "Pedí desde tu celular",         "Solo necesitás tu correo y tu DNI. Sin filas ni papeleos."),
    ("Atención",      "Te atendemos por WhatsApp",     "Consultá sin compromiso y resolvé todo por chat."),
    ("Clientes",      "La segunda vez, más fácil",     "Si ya sos cliente, cargamos tus datos por vos."),
    ("Flexibilidad",  "Plazos a tu medida",            "Elegí el plazo que mejor se adapte a tu bolsillo."),
    ("Confianza",     "Buen pagador, más beneficios",  "Pagá en fecha y accedé a montos cada vez mayores."),
    ("Cobertura",     "Estamos cerca tuyo",            "Acompañamos a Zona Norte y el AMBA, Buenos Aires."),
    ("Transparencia", "Sin letra chica",               "Sabés desde el primer día cuánto y cuándo vas a pagar."),
    ("Pagos",         "Cobrá como te quede cómodo",    "Recibí tu préstamo por Mercado Pago o transferencia."),
    ("Impulso",       "Tu impulso, cuando lo necesitás", "Un préstamo simple y confiable, a un mensaje de distancia."),
]


def make_post(i, label, headline, support):
    img, d = base_canvas()

    # header: small badge + wordmark, centered
    b = ic_badge(96)
    gap = 18
    wm_f = font(46, bold=True)
    wm_w = d.textlength("Impulso ", font=wm_f) + d.textlength("Crédito", font=wm_f)
    group_w = b.width + gap + wm_w
    x0 = (S - group_w) / 2
    img.alpha_composite(b, (int(x0), 70 - (b.height - 96) // 2))
    d.text((x0 + b.width + gap, 92), "Impulso ", font=wm_f, fill=BLUE)
    d.text((x0 + b.width + gap + d.textlength("Impulso ", font=wm_f), 92), "Crédito", font=wm_f, fill=GOLD)

    # series index, top-right
    d.text((S - 90, 70), f"{i:02d}/10", font=font(28, bold=True), fill=SLATE)

    # ── middle block, vertically centered ──
    lab_f  = font(32, bold=True)
    head_f = font(78, bold=True)
    sup_f  = font(40)
    head_lines = wrap_px(d, headline, head_f, S - 200)
    sup_lines  = wrap_px(d, support, sup_f, S - 220)

    head_lh, sup_lh = 94, 56
    block_h = 60 + len(head_lines) * head_lh + 36 + len(sup_lines) * sup_lh
    y = (S - block_h) / 2 - 20

    # gold accent pill above the label
    d.rounded_rectangle([S / 2 - 34, y, S / 2 + 34, y + 8], radius=4, fill=GOLD)
    y += 30
    ctext(d, S / 2, y, label.upper(), lab_f, GOLD, tracking=4)
    y += 70
    for ln in head_lines:
        ctext(d, S / 2, y, ln, head_f, WHITE)
        y += head_lh
    y += 24
    for ln in sup_lines:
        ctext(d, S / 2, y, ln, sup_f, SLATE)
        y += sup_lh

    # ── footer ──
    d.line([S / 2 - 90, S - 190, S / 2 + 90, S - 190], fill=GOLD, width=3)
    ctext(d, S / 2, S - 165, "Consultá por WhatsApp", font(38, bold=True), WHITE)
    ctext(d, S / 2, S - 110, "@impulso.credito.ars   ·   Zona Norte + AMBA", font(30), SLATE)

    out = HERE / f"post_{i:02d}.png"
    img.convert("RGB").save(out, "PNG")
    print(out.name)


# ── Story-highlight covers: one letter of IMPULSO each ───────────────────────

def make_highlight(i, letter):
    """Circular story-highlight cover (content centred for the circle crop)."""
    img = vgradient(S, S, BG_TOP, BG_BOT)
    # soft centre glow
    glow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse([S * 0.24, S * 0.24, S * 0.76, S * 0.76],
                                 fill=(70, 110, 180, 80))
    img.alpha_composite(glow.filter(ImageFilter.GaussianBlur(80)))

    d = ImageDraw.Draw(img)
    r = int(S * 0.40)                               # thin gold ring, inside the safe circle
    d.ellipse([S / 2 - r, S / 2 - r, S / 2 + r, S / 2 + r], outline=GOLD, width=8)

    f = font(int(S * 0.5), bold=True)               # big centred letter
    bb = d.textbbox((0, 0), letter, font=f)
    d.text(((S - (bb[2] - bb[0])) / 2 - bb[0],
            (S - (bb[3] - bb[1])) / 2 - bb[1]), letter, font=f, fill=WHITE)

    out = HERE / f"highlight_{i}_{letter}.png"
    img.convert("RGB").save(out, "PNG")
    print(out.name)


if __name__ == "__main__":
    make_profile()
    for idx, (lab, head, sup) in enumerate(POSTS, start=1):
        make_post(idx, lab, head, sup)
    for idx, letter in enumerate("IMPULSO", start=1):
        make_highlight(idx, letter)
    print("Done — assets in", HERE)
