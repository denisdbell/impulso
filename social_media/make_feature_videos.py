#!/usr/bin/env python3
"""
Impulso Crédito — Instagram Story videos from the feature posts.

Builds polished 1080×1920 (9:16) slideshow videos: each square feature card
floats (rounded + shadow) on the brand background, with dissolve transitions,
a slow Ken-Burns zoom, an IG-style segmented progress bar, a story title and a
persistent CTA. 30 cards → 3 videos of 10 (~30s each), ideal for Story segments.

Run:  python3 make_feature_videos.py
"""

import warnings
from pathlib import Path

warnings.filterwarnings("ignore")

import imageio
import numpy as np
from PIL import Image, ImageDraw, ImageFilter

from generate_assets import font, vgradient, BG_TOP, BG_BOT, BLUE, GOLD, WHITE, SLATE

HERE = Path(__file__).parent
W, H, FPS = 1080, 1920, 30
SLOT, TRANS = 3.0, 0.5          # seconds per card / dissolve length
CARD_W, PAD = 940, 66
CX, CY = 540, 860               # card centre on the story canvas


def clamp(x, a=0.0, b=1.0): return max(a, min(b, x))


def ctext(d, cx, y, text, f, rgb, a=1.0, tracking=0):
    col = (*rgb, int(255 * clamp(a)))
    if tracking:
        total = sum(d.textlength(ch, font=f) for ch in text) + tracking * (len(text) - 1)
        x = cx - total / 2
        for ch in text:
            d.text((x, y), ch, font=f, fill=col); x += d.textlength(ch, font=f) + tracking
    else:
        d.text((cx - d.textlength(text, font=f) / 2, y), text, font=f, fill=col)


# ── static story background (gradient + bars + title + CTA), built once ──────
def _bg():
    img = vgradient(W, H, BG_TOP, BG_BOT)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, 8], fill=GOLD)
    d.rectangle([0, H - 8, W, H], fill=GOLD)
    ctext(d, W / 2, 150, "P O R   Q U É   E L E G I R N O S", font(34, bold=True), GOLD, tracking=2)
    d.line([W / 2 - 60, 1512, W / 2 + 60, 1512], fill=GOLD, width=3)
    ctext(d, W / 2, 1540, "Pedí el tuyo 100% online", font(46, bold=True), WHITE)
    ctext(d, W / 2, 1610, "WhatsApp   ·   @impulso.credito.ars   ·   Zona Norte + AMBA", font(30), SLATE)
    return img

BG = _bg()
NAVY = (*BG_BOT,)


# ── cards ────────────────────────────────────────────────────────────────────
def build_card(path):
    im = Image.open(path).convert("RGB").resize((CARD_W, CARD_W), Image.LANCZOS)
    canvas = Image.new("RGBA", (CARD_W + 2 * PAD, CARD_W + 2 * PAD), (0, 0, 0, 0))
    sh = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle([PAD, PAD + 20, PAD + CARD_W, PAD + CARD_W + 20], radius=46, fill=(0, 0, 0, 160))
    canvas.alpha_composite(sh.filter(ImageFilter.GaussianBlur(28)))
    mask = Image.new("L", (CARD_W, CARD_W), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, CARD_W - 1, CARD_W - 1], radius=40, fill=255)
    card = Image.new("RGBA", (CARD_W, CARD_W), (0, 0, 0, 0)); card.paste(im, (0, 0), mask)
    ImageDraw.Draw(card).rounded_rectangle([1, 1, CARD_W - 2, CARD_W - 2], radius=40, outline=(233, 170, 70, 110), width=2)
    canvas.alpha_composite(card, (PAD, PAD))
    return canvas


def place(frame, card, scale, alpha):
    disp = int(card.width * scale)
    r = card.resize((disp, disp), Image.LANCZOS)
    a = r.split()[3]
    if alpha < 1:
        a = a.point(lambda p: int(p * alpha))
    frame.paste(r, (int(CX - disp / 2), int(CY - disp / 2)), a)


def progress(d, n, idx, local):
    bw, y, gap, h = 1000, 54, 8, 8
    x0 = (W - bw) / 2; sw = (bw - (n - 1) * gap) / n
    for k in range(n):
        x = x0 + k * (sw + gap)
        d.rounded_rectangle([x, y, x + sw, y + h], radius=4, fill=(255, 255, 255, 55))
        fp = 1.0 if k < idx else (clamp(local / SLOT) if k == idx else 0.0)
        if fp > 0:
            d.rounded_rectangle([x, y, x + sw * fp, y + h], radius=4, fill=(*GOLD, 255))


def render(t, cards):
    n = len(cards)
    frame = BG.copy()
    idx = min(int(t / SLOT), n - 1)
    local = t - idx * SLOT
    zoom = lambda lt: 1.0 + 0.045 * clamp(lt / SLOT)
    if local > SLOT - TRANS and idx < n - 1:
        b = (local - (SLOT - TRANS)) / TRANS
        place(frame, cards[idx + 1], zoom(t - (idx + 1) * SLOT), 1.0)   # incoming (underneath)
        place(frame, cards[idx], zoom(local), 1 - b)                    # outgoing dissolves out
    else:
        place(frame, cards[idx], zoom(local), 1.0)
    progress(ImageDraw.Draw(frame), n, idx, local)
    # intro / outro fades
    dur = n * SLOT
    fade = max(clamp((0.5 - t) / 0.5), clamp((t - (dur - 0.5)) / 0.5))
    if fade > 0:
        frame.alpha_composite(Image.new("RGBA", (W, H), (*NAVY, int(255 * fade))))
    return frame


def build_video(group_idx, paths):
    cards = [build_card(p) for p in paths]
    dur = len(cards) * SLOT
    out = HERE / f"story_features_{group_idx}.mp4"
    nframes = int(dur * FPS)
    writer = imageio.get_writer(out, fps=FPS, codec="libx264", macro_block_size=8,
                                pixelformat="yuv420p", ffmpeg_params=["-crf", "20", "-preset", "medium"])
    for i in range(nframes):
        fr = render(i / FPS, cards).convert("RGB")
        if i == int(nframes * 0.18):
            fr.save(HERE / f"vpreview_{group_idx}.png", "PNG")
        writer.append_data(np.asarray(fr))
    writer.close()
    print(f"story_features_{group_idx}.mp4  ({len(cards)} cards, {dur:.0f}s)")


def main():
    feats = sorted(HERE.glob("feat_*.png"))
    if not feats:
        print("No feat_*.png — run generate_feature_posts.py first."); return
    for g, start in enumerate([0, 10, 20], start=1):
        chunk = feats[start:start + 10]
        if chunk:
            build_video(g, chunk)
    print("Done — Story videos in", HERE)


if __name__ == "__main__":
    main()
