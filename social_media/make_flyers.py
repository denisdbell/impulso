#!/usr/bin/env python3
"""
Impulso Crédito — 3 flyers explicativos, en dos tamaños de Instagram.

Mismo tema que los feat_XX (fondo navy, barras doradas, insignia IC + marca,
medallones navy con aro dorado e icono blanco, pastilla dorada de etiqueta y
pie "Consultá por WhatsApp"). Texto súper simple, en castellano rioplatense,
pensado para que lo entienda un chico de 12 años.

Se generan DOS tamaños de cada flyer:
  • 4:5 vertical  (1080×1350)  →  flyer_<tema>.png       (máximo tamaño del feed)
  • 1:1 cuadrado  (1080×1080)  →  flyer_<tema>_1x1.png   (post clásico del feed)

Temas:
  requisitos   — qué necesitás para pedir el préstamo (paso a paso)
  niveles      — cómo funciona la escalera de niveles (graduación)
  recordatorios— cómo te avisamos la cuota antes de cada pago

Cifras tomadas del sistema: solicitud (Validaciones.gs), niveles V-32 /
flyer-niveles-es.html, avisos 30/15/7 días (make_reminder_post.py).

Correr:  python3 make_flyers.py
"""

import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

from generate_assets import (font, vgradient, ic_badge, BG_TOP, BG_BOT,
                             BLUE, GOLD, WHITE, SLATE, wrap_px, ctext)
from generate_feature_posts import make_icon, circle_gradient
from generate_captions import CONTACT, HASHTAGS

HERE = Path(__file__).parent
W = 1080
# Instagram recorta los lados cuando se mezclan proporciones (1:1 vs 4:5) en un
# carrusel. Mantenemos todo el contenido importante dentro de una "zona segura"
# centrada, con márgenes amplios a izquierda y derecha, para que nunca se corte.
SAFE = 150
GREEN = (61, 190, 107)
RED = (226, 96, 96)
MED_TOP, MED_BOT = (56, 90, 152), (26, 44, 84)   # medallón navy (igual que feat_XX)

BADGE = ic_badge(80)

# Tamaños a generar: (sufijo de archivo, alto en px). Mismo ancho (1080) siempre.
SIZES = [("", 1350), ("_1x1", 1080)]


def money(v):
    return "$" + f"{v:,.0f}".replace(",", ".")


def L(H):
    """Constantes de maquetación dependientes del alto (vertical vs cuadrado)."""
    sq = H <= 1100
    return {
        "sq": sq,
        "title_y": 190 if sq else 210,
        "title_sz": 58 if sq else 70,
        "title_lh": 68 if sq else 82,
        "sub_sz": 31 if sq else 35,
        "sub_lh": 43 if sq else 47,
        "foot1": 118 if sq else 152,
        "foot2": 74 if sq else 102,
        "band_pad": 20 if sq else 24,
    }


# ── Fondo navy (barras doradas + glows + puntitos), idéntico en estilo ─────────
def base_canvas(H):
    img = vgradient(W, H, BG_TOP, BG_BOT)
    deco = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    dd = ImageDraw.Draw(deco)
    dd.ellipse([-260, -340, 360, 300], fill=(70, 110, 180, 55))          # glow azul arriba-izq
    dd.ellipse([720, H - 310, 1360, H + 330], fill=(233, 170, 70, 40))   # glow dorado abajo-der
    img.alpha_composite(deco.filter(ImageFilter.GaussianBlur(120)))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, 8], fill=GOLD)
    d.rectangle([0, H - 8, W, H], fill=GOLD)
    for (x, y, r) in [(92, 150, 9), (150, 205, 6), (988, 158, 9), (930, 212, 6)]:
        d.ellipse([x - r, y - r, x + r, y + r], fill=(233, 170, 70, 110))
    return img


def header(img, d):
    wm = font(46, bold=True)
    ww = d.textlength("Impulso ", font=wm) + d.textlength("Crédito", font=wm)
    gw = BADGE.width + 16 + ww
    x0 = (W - gw) / 2
    img.alpha_composite(BADGE, (int(x0), 70))
    tx = x0 + BADGE.width + 16
    d.text((tx, 88), "Impulso ", font=wm, fill=BLUE)
    d.text((tx + d.textlength("Impulso ", font=wm), 88), "Crédito", font=wm, fill=GOLD)


def title_block(img, d, label, title, subtitle, cfg):
    y = cfg["title_y"]
    lf = font(30, bold=True)
    lw = d.textlength(label.upper(), font=lf) + 54
    d.rounded_rectangle([W / 2 - lw / 2, y, W / 2 + lw / 2, y + 52], radius=26, outline=GOLD, width=3)
    ctext(d, W / 2, y + 9, label.upper(), lf, GOLD, tracking=2)
    y += 52 + 26
    hf = font(cfg["title_sz"], bold=True)
    for ln in wrap_px(d, title, hf, W - 130):
        ctext(d, W / 2, y, ln, hf, WHITE)
        y += cfg["title_lh"]
    y += 6
    sf = font(cfg["sub_sz"])
    for ln in wrap_px(d, subtitle, sf, W - 170):
        ctext(d, W / 2, y, ln, sf, SLATE)
        y += cfg["sub_lh"]
    return y


def footer(d, H, cfg):
    ctext(d, W / 2, H - cfg["foot1"], "Consultá por WhatsApp", font(36, bold=True), WHITE)
    ctext(d, W / 2, H - cfg["foot2"], "@impulso.credito.ars   ·   Zona Norte + AMBA", font(29), SLATE)


def mini_medallion(disp, icon):
    med = circle_gradient(disp, MED_TOP, MED_BOT)
    ImageDraw.Draw(med).ellipse([2, 2, disp - 3, disp - 3], outline=GOLD, width=max(3, int(disp * 0.045)))
    ic = make_icon(icon, int(disp * 0.54), WHITE)
    med.alpha_composite(ic, (int((disp - ic.width) / 2), int((disp - ic.height) / 2)))
    return med


def big_medallion(img, d, icon, cx, cy, mr):
    glow = Image.new("RGBA", img.size, (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse([cx - mr - 30, cy - mr - 30, cx + mr + 30, cy + mr + 30], fill=(233, 170, 70, 60))
    img.alpha_composite(glow.filter(ImageFilter.GaussianBlur(50)))
    med = circle_gradient(mr * 2, MED_TOP, MED_BOT)
    img.alpha_composite(med, (int(cx - mr), int(cy - mr)))
    d.ellipse([cx - mr, cy - mr, cx + mr, cy + mr], outline=GOLD, width=7)
    ic = make_icon(icon, int(mr * 1.15), WHITE)
    img.alpha_composite(ic, (int(cx - ic.width / 2), int(cy - ic.height / 2)))
    for ang, rr in ((-40, 7), (20, 10), (210, 8), (140, 6)):
        ax = cx + (mr + 24) * math.cos(math.radians(ang))
        ay = cy + (mr + 24) * math.sin(math.radians(ang))
        d.ellipse([ax - rr, ay - rr, ax + rr, ay + rr], fill=GOLD)


def seg_center(d, cx, y, segs):
    """Dibuja una línea centrada con tramos de distinto color (mismo font)."""
    total = sum(d.textlength(t, font=f) for t, f, _ in segs)
    x = cx - total / 2
    for t, f, c in segs:
        d.text((x, y), t, font=f, fill=c)
        x += d.textlength(t, font=f)


def list_row(img, d, icon, title, sub, y, h):
    disp = max(64, min(112, int(h - 14)))
    k = disp / 112.0
    med = mini_medallion(disp, icon)
    img.alpha_composite(med, (SAFE, int(y + (h - disp) / 2)))
    tx = SAFE + disp + max(22, int(disp * 0.30))
    tf = font(max(26, int(40 * k)), bold=True)
    sf = font(max(20, int(29 * k)))
    cy = y + h / 2
    d.text((tx, cy - int(44 * k)), title, font=tf, fill=WHITE)
    d.text((tx, cy + int(7 * k)), sub, font=sf, fill=SLATE)


# ── Flyer 1 — requisitos (qué necesitás para pedir) ───────────────────────────
ROWS = [
    ("id",    "Tu correo y tu DNI",            "Es lo primero que te pedimos."),
    ("doc",   "Tus datos",                     "Nombre, CUIL, teléfono y dirección."),
    ("phone", "Dos fotos del DNI",             "Una de adelante y una de atrás."),
    ("users", "Dos personas que te conozcan",  "Su nombre y su teléfono, nada más."),
    ("coins", "Cuánta plata querés",           "Vos elegís el monto que necesitás."),
]


def flyer_requisitos(H, suffix):
    cfg = L(H)
    img = base_canvas(H)
    d = ImageDraw.Draw(img)
    header(img, d)
    y = title_block(img, d, "Requisitos", "¿Qué necesitás para pedir?",
                    "Es 100% online y te lleva unos minutos.", cfg)
    rows_top = y + cfg["band_pad"]
    rows_bot = H - cfg["foot1"] - 44
    slot = (rows_bot - rows_top) / len(ROWS)
    for i, (icon, t, s) in enumerate(ROWS):
        list_row(img, d, icon, t, s, rows_top + i * slot, slot)
    footer(d, H, cfg)
    img.convert("RGB").save(HERE / f"flyer_requisitos{suffix}.png", "PNG")


# ── Flyer 2 — niveles (cómo subís de nivel) ───────────────────────────────────
def coin_sprite(disp):
    pad = 12
    s = disp
    im = Image.new("RGBA", (s + 2 * pad, s + 2 * pad), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.ellipse([pad, pad, pad + s, pad + s], fill=GOLD)
    d.ellipse([pad + 5, pad + 5, pad + s - 5, pad + s - 5], outline=(255, 255, 255, 160), width=3)
    f = font(int(s * 0.62), bold=True)
    bb = d.textbbox((0, 0), "$", font=f)
    d.text((pad + (s - (bb[2] - bb[0])) / 2 - bb[0], pad + (s - (bb[3] - bb[1])) / 2 - bb[1]),
           "$", font=f, fill=(44, 34, 10))
    return im


def star_sprite(size, col):
    im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    cx = cy = size / 2
    R, r = size * 0.48, size * 0.2
    pts = []
    for i in range(10):
        ang = -math.pi / 2 + i * math.pi / 5
        rad = R if i % 2 == 0 else r
        pts.append((cx + rad * math.cos(ang), cy + rad * math.sin(ang)))
    ImageDraw.Draw(im).polygon(pts, fill=col)
    return im


def rule_chip(d, cx, y, icon_draw, text, sz):
    tf = font(sz, bold=True)
    tw = d.textlength(text, font=tf)
    total = 44 + 14 + tw
    x = cx - total / 2
    icon_draw(d, x + 22, y + 22)
    d.text((x + 44 + 14, y + int(sz * 0.12)), text, font=tf, fill=WHITE)


def draw_up(d, cx, cy):      # flecha hacia arriba verde en círculo
    d.ellipse([cx - 22, cy - 22, cx + 22, cy + 22], fill=GREEN)
    d.line([(cx, cy + 10), (cx, cy - 11)], fill=WHITE, width=6)
    d.line([(cx - 11, cy - 1), (cx, cy - 12), (cx + 11, cy - 1)], fill=WHITE, width=6, joint="curve")


def draw_down(d, cx, cy):    # flecha hacia abajo roja en círculo
    d.ellipse([cx - 22, cy - 22, cx + 22, cy + 22], fill=RED)
    d.line([(cx, cy - 10), (cx, cy + 11)], fill=WHITE, width=6)
    d.line([(cx - 11, cy + 1), (cx, cy + 12), (cx + 11, cy + 1)], fill=WHITE, width=6, joint="curve")


def flyer_niveles(H, suffix):
    cfg = L(H)
    sq = cfg["sq"]
    img = base_canvas(H)
    d = ImageDraw.Draw(img)
    header(img, d)
    tb = title_block(img, d, "Niveles", "Subí de nivel",
                     "Es como un juego: pagá a tiempo y pedí cada vez más.", cfg)

    # Las barras y TODO su contenido quedan dentro de la zona segura [SAFE, W-SAFE].
    n = 5
    x0, x1 = SAFE + 10, W - SAFE - 10
    gap = 22
    bw = (x1 - x0 - gap * (n - 1)) / n
    mid1 = H - cfg["foot1"] - 44
    base_y = mid1 - (200 if sq else 190)
    h_max = 205 if sq else 320
    h_min = 80 if sq else 106
    for i in range(n):
        bx = x0 + i * (bw + gap)
        bh = h_min + (h_max - h_min) * (i / (n - 1))
        top = base_y - bh
        lit = i == n - 1
        d.rounded_rectangle([bx, top, bx + bw, base_y], radius=20, fill=GOLD if lit else (46, 70, 120))
        if not lit:
            d.rounded_rectangle([bx, top, bx + bw, base_y], radius=20, outline=(233, 170, 70, 120), width=2)
    cxl = x0 + (n - 1) * (bw + gap) + bw / 2
    top_last = base_y - h_max
    coin = coin_sprite(58 if sq else 68)
    img.alpha_composite(coin, (int(cxl - coin.width / 2), int(top_last - coin.height - 14)))
    for dx, dy, ssz in ((-70, 16, 24 if sq else 28), (70, 4, 28 if sq else 34), (6, -24, 20 if sq else 24)):
        st = star_sprite(ssz, GOLD)
        img.alpha_composite(st, (int(cxl + dx - ssz / 2), int(top_last - 52 + dy)))

    # cifras clave CENTRADAS (no en las esquinas) → nunca se cortan al recortar los lados
    amt = font(46 if sq else 50, bold=True)
    seg_center(d, W / 2, base_y + 20, [(money(150000), amt, WHITE),
                                       ("   →   ", amt, SLATE),
                                       (money(500000), amt, GOLD)])
    ctext(d, W / 2, base_y + (72 if sq else 78), "del nivel inicial al tope", font(28), SLATE)
    ctext(d, W / 2, (tb + 14) if sq else (top_last - 150), "+ $50.000 por nivel",
          font(33, bold=True), SLATE)

    rsz = 30 if sq else 33
    r1 = base_y + (118 if sq else 132)
    r2 = r1 + (50 if sq else 54)
    rule_chip(d, W / 2, r1, draw_up, "Pagás en fecha → subís y pedís más", rsz)
    rule_chip(d, W / 2, r2, draw_down, "Si te atrasás, bajás de nivel", rsz)

    footer(d, H, cfg)
    img.convert("RGB").save(HERE / f"flyer_niveles{suffix}.png", "PNG")


# ── Flyer 3 — recordatorios (te avisamos la cuota) ────────────────────────────
def flyer_recordatorios(H, suffix):
    cfg = L(H)
    sq = cfg["sq"]
    img = base_canvas(H)
    d = ImageDraw.Draw(img)
    header(img, d)
    y = title_block(img, d, "Avisos a tiempo", "Te avisamos tu cuota",
                    "Un mensajito antes, para que no te olvides.", cfg)

    mid0 = y + cfg["band_pad"]
    mid1 = H - cfg["foot1"] - 44
    mr = 90 if sq else 118
    big_medallion(img, d, "bell", W / 2, mid0 + mr + 8, mr)

    ly = mid1 - (156 if sq else 290)
    lx, rx = SAFE + 70, W - SAFE - 70
    d.line([(lx, ly), (rx, ly)], fill=(233, 170, 70, 150), width=5)
    nodes = [("30", "días antes"), ("15", "días antes"), ("7", "días antes"), (None, "día del pago")]
    nn = len(nodes)
    rad = 36 if sq else 40
    nf = font(34 if sq else 38, bold=True)
    lf = font(25 if sq else 27)
    for i, (big, lab) in enumerate(nodes):
        cx = lx + (rx - lx) * i / (nn - 1)
        last = i == nn - 1
        d.ellipse([cx - rad, ly - rad, cx + rad, ly + rad], fill=GREEN if last else GOLD)
        d.ellipse([cx - rad, ly - rad, cx + rad, ly + rad], outline=(12, 19, 38), width=4)
        if last:                                   # tilde dibujada (el glifo ✓ no está en la fuente)
            d.line([(cx - 17, ly - 1), (cx - 4, ly + 14), (cx + 19, ly - 16)], fill=WHITE, width=8, joint="curve")
        else:
            ctext(d, cx, ly - int(rad * 0.62), big, nf, (28, 24, 12))
        ctext(d, cx, ly + rad + 14, lab, lf, SLATE)

    cap1 = ly + rad + (76 if sq else 88)
    ctext(d, W / 2, cap1, "Te escribimos 3 veces antes de tu fecha.",
          font(34 if sq else 37, bold=True), WHITE)
    ctext(d, W / 2, cap1 + (46 if sq else 52), "Así llegás tranquilo, sin apuros ni sustos.",
          font(30 if sq else 32), SLATE)

    footer(d, H, cfg)
    img.convert("RGB").save(HERE / f"flyer_recordatorios{suffix}.png", "PNG")


# ── Pies de foto (mismo bloque CONTACT + HASHTAGS que los feat_XX) ────────────
CAPTIONS = {
    "flyer_requisitos": (
        "📝 ¿Querés pedir tu préstamo? Es fácil y 100% online. Solo necesitás: "
        "tu correo y tu DNI, tus datos (nombre, CUIL, teléfono y dirección), dos "
        "fotos del DNI (frente y dorso), dos personas de confianza con su teléfono "
        "y el monto que querés. ¡Listo en minutos!"
    ),
    "flyer_niveles": (
        "📈 Es como un juego: cada vez que pagás tu préstamo a tiempo, SUBÍS DE "
        "NIVEL y podés pedir más. Arrancás en $150.000 y sumás +$50.000 por nivel, "
        "hasta un tope de $500.000. ¿Te atrasás? Bajás de nivel. Pagá en fecha y "
        "crecé con nosotros."
    ),
    "flyer_recordatorios": (
        "🔔 Nunca más te olvides de tu cuota. Te avisamos 30, 15 y 7 días antes de "
        "tu fecha de pago, para que llegues tranquilo y sin apuros. Un recordatorio "
        "amable, nada de sustos."
    ),
}


def write_captions():
    for name, body in CAPTIONS.items():
        caption = f"{body}\n\n{CONTACT}\n\n{HASHTAGS}\n"
        (HERE / f"{name}.txt").write_text(caption, encoding="utf-8")


def main():
    for suffix, H in SIZES:
        flyer_requisitos(H, suffix)
        flyer_niveles(H, suffix)
        flyer_recordatorios(H, suffix)
        tag = "4:5 vertical" if H > 1100 else "1:1 cuadrado"
        print(f"{tag} ({W}×{H}):  flyer_requisitos{suffix}.png  flyer_niveles{suffix}.png  flyer_recordatorios{suffix}.png")
    write_captions()
    print("Listo — 3 flyers × 2 tamaños + pies en", HERE)


if __name__ == "__main__":
    main()
