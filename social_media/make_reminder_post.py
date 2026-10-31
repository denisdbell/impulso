#!/usr/bin/env python3
"""
Impulso Crédito — post único de "Avisos a tiempo" (recordatorios, sólo en clave positiva).

Reusa el mismo tema de los feat_XX (medallón navy/oro, icono campana, pie y hashtags),
pero sin el contador NN/30 y con un mensaje 100% positivo: avisamos con anticipación
para que llegues tranquilo al pago. Nada de mora, atraso ni penalidades.

Escribe:  feat_reminders.png  +  feat_reminders.txt

Run:  python3 make_reminder_post.py
"""

import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

from generate_assets import font, BLUE, GOLD, WHITE, SLATE
from generate_feature_posts import (BASE, BADGE, make_icon, circle_gradient,
                                     wrap_px, ctext, S)
from generate_captions import CONTACT, HASHTAGS

HERE = Path(__file__).parent

# icono, etiqueta, titular, bajada, emoji — sólo lado positivo de los recordatorios
ICON = "bell"
LABEL = "A tiempo"
HEADLINE = "Te avisamos con tiempo"
SUPPORT = "Aviso 30, 15 y 7 días antes de tu pago, para que llegues tranquilo."
EMOJI = "🔔"


def render():
    img = BASE.copy()
    d = ImageDraw.Draw(img)

    # cabecera: insignia IC + wordmark (igual que los feat_XX, sin contador NN/30)
    wm = font(46, bold=True)
    ww = d.textlength("Impulso ", font=wm) + d.textlength("Crédito", font=wm)
    gw = BADGE.width + 16 + ww
    x0 = (S - gw) / 2
    img.alpha_composite(BADGE, (int(x0), 74))
    d.text((x0 + BADGE.width + 16, 92), "Impulso ", font=wm, fill=BLUE)
    d.text((x0 + BADGE.width + 16 + d.textlength("Impulso ", font=wm), 92), "Crédito", font=wm, fill=GOLD)

    # medallón focal con el icono
    mc, mr = 405, 128
    glow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse([S / 2 - mr - 30, mc - mr - 30, S / 2 + mr + 30, mc + mr + 30], fill=(233, 170, 70, 60))
    img.alpha_composite(glow.filter(ImageFilter.GaussianBlur(50)))
    med = circle_gradient(mr * 2, (56, 90, 152), (26, 44, 84))
    img.alpha_composite(med, (int(S / 2 - mr), mc - mr))
    d.ellipse([S / 2 - mr, mc - mr, S / 2 + mr, mc + mr], outline=GOLD, width=7)
    ic = make_icon(ICON, 148, WHITE)
    img.alpha_composite(ic, (int(S / 2 - ic.width / 2), int(mc - ic.height / 2)))
    for ang, rr in ((-40, 7), (20, 10), (210, 8), (140, 6)):
        ax = S / 2 + (mr + 24) * math.cos(math.radians(ang))
        ay = mc + (mr + 24) * math.sin(math.radians(ang))
        d.ellipse([ax - rr, ay - rr, ax + rr, ay + rr], fill=GOLD)

    # pastilla de etiqueta + titular + bajada, centrado como bloque
    lf, hf, sf = font(30, bold=True), font(64, bold=True), font(35)
    hlines = wrap_px(d, HEADLINE, hf, S - 140)
    slines = wrap_px(d, SUPPORT, sf, S - 150)
    lh, slh = 74, 47
    block_h = 52 + 24 + len(hlines) * lh + 14 + len(slines) * slh
    region_top, region_bot = 565, 898
    y = region_top + max(0, (region_bot - region_top - block_h) / 2)

    lw = d.textlength(LABEL.upper(), font=lf) + 54
    d.rounded_rectangle([S / 2 - lw / 2, y, S / 2 + lw / 2, y + 52], radius=26, outline=GOLD, width=3)
    ctext(d, S / 2, y + 9, LABEL.upper(), lf, GOLD, tracking=2)
    y += 52 + 24
    for ln in hlines:
        ctext(d, S / 2, y, ln, hf, WHITE)
        y += lh
    y += 14
    for ln in slines:
        ctext(d, S / 2, y, ln, sf, SLATE)
        y += slh

    # pie CTA (idéntico a los feat_XX)
    ctext(d, S / 2, S - 150, "Consultá por WhatsApp", font(36, bold=True), WHITE)
    ctext(d, S / 2, S - 100, "@impulso.credito.ars   ·   Zona Norte + AMBA", font(29), SLATE)

    img.convert("RGB").save(HERE / "feat_reminders.png", "PNG")


def main():
    render()
    body = f"{EMOJI} {HEADLINE}. {SUPPORT}"
    caption = f"{body}\n\n{CONTACT}\n\n{HASHTAGS}\n"
    (HERE / "feat_reminders.txt").write_text(caption, encoding="utf-8")
    print("feat_reminders.png + feat_reminders.txt listos")


if __name__ == "__main__":
    main()
