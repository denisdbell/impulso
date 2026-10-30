#!/usr/bin/env python3
"""
Impulso Crédito — "Cómo pedir tu préstamo" Instagram Story tutorial.

A high-fidelity 1080×1920 (9:16) vertical video that walks through the 6-step
loan-application form using the real app screenshots (social_media/how-to-use/
1.png … 9.png). The app lives inside a phone mock on the brand background; a
finger taps each field while the data types itself in / reveals (dedo que carga
los datos), with Spanish captions, an IG story progress bar, synthesized
ambient music, soft tap ticks and a success chime.

Pipeline: render frames with Pillow/numpy → encode a silent H.264 mp4 with
imageio → synthesize a WAV → mux the two with imageio-ffmpeg's bundled ffmpeg.

Run:
    python3 make_howto_video.py            # full render + audio  -> story_como_usar.mp4
    python3 make_howto_video.py probe      # one PNG per scene    -> _probe_*.png (fast)
"""

import math
import subprocess
import sys
import wave
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")

import imageio
import imageio_ffmpeg
import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

from generate_assets import (
    font, vgradient, ic_badge,
    BG_TOP, BG_BOT, BLUE, GOLD, WHITE, SLATE,
)

HERE = Path(__file__).parent
SRC = HERE / "how-to-use"
W, H, FPS = 1080, 1920, 30
NAVY = BG_BOT

# ── Phone geometry ───────────────────────────────────────────────────────────
PW, PH = 876, 1392
PX0, PY0 = (W - PW) // 2, 350
BEZEL = 24
ISLAND_H = 20
SX0, SY0 = PX0 + BEZEL, PY0 + BEZEL + ISLAND_H     # inner screen origin
SW, SH = PW - 2 * BEZEL, PH - 2 * BEZEL - ISLAND_H  # inner screen size
SCR_RADIUS = 46
PHONE_RADIUS = 72


def clamp(x, a=0.0, b=1.0): return max(a, min(b, x))
def lerp(a, b, t): return a + (b - a) * t
def ease(t):  # smoothstep ease-in-out
    t = clamp(t); return t * t * (3 - 2 * t)
def ease_out(t): t = clamp(t); return 1 - (1 - t) ** 3


def ctext(d, cx, y, text, f, rgb, a=1.0, tracking=0):
    col = (*rgb, int(255 * clamp(a)))
    if tracking:
        total = sum(d.textlength(c, font=f) for c in text) + tracking * (len(text) - 1)
        x = cx - total / 2
        for c in text:
            d.text((x, y), c, font=f, fill=col); x += d.textlength(c, font=f) + tracking
    else:
        d.text((cx - d.textlength(text, font=f) / 2, y), text, font=f, fill=col)


# ════════════════════════════════════════════════════════════════════════════
#  SCENES — every field box is in the screenshot's own native pixels.
#  kind: type   -> render our own typed text (field is empty in the shot)
#        reveal -> wipe-reveal the value already printed in the shot
#        check  -> draw a checkmark into an empty checkbox
#        file   -> tap + a filename chip floats in
#        tapbtn -> tap a button/radio (highlight + press, no content change)
# ════════════════════════════════════════════════════════════════════════════
SCENES = [
    dict(img="1.png", caption="Ingresá tu correo y tu DNI", kind="form", actions=[
        dict(kind="type",   box=(66, 836, 1010, 904),  tap=(120, 870),  text="alvaro.gomez@gmail.com", dur=1.8),
        dict(kind="type",   box=(66, 976, 1010, 1044), tap=(120, 1010), text="11.223.344", dur=1.3),
        dict(kind="tapbtn", box=(62, 1083, 1012, 1160), tap=(533, 1121), dur=0.9),
    ]),
    dict(img="2.png", caption="¿Ya sos cliente? Cargamos tus datos solos", kind="form",
         edits=[
        dict(box=(205, 458, 316, 481), text="Álvaro!", size=20, bold=True, color=WHITE, pos=(209, 469)),
        dict(box=(70, 696, 928, 742),  text="alvaro.gomez@gmail.com",     size=25, color=(234, 238, 244), pos=(90, 718)),
        dict(box=(70, 796, 928, 844),  text="Álvaro Gómez",               size=25, color=(234, 238, 244), pos=(90, 819)),
        dict(box=(408, 904, 727, 952), text="20112233440",                size=25, color=(234, 238, 244), pos=(424, 927)),
        dict(box=(70, 1014, 928, 1062), text="Av. San Martín 1234, Tigre", size=25, color=(234, 238, 244), pos=(90, 1037)),
         ],
         actions=[
        dict(kind="reveal", box=(66, 690, 933, 744),  tap=(120, 717),  dur=0.62),
        dict(kind="reveal", box=(66, 790, 933, 846),  tap=(120, 818),  dur=0.55),
        dict(kind="reveal", box=(66, 900, 388, 954),  tap=(120, 927),  dur=0.52),
        dict(kind="reveal", box=(402, 900, 731, 954), tap=(440, 927),  dur=0.52),
        dict(kind="reveal", box=(66, 1010, 933, 1064), tap=(120, 1037), dur=0.55),
    ]),
    dict(img="3.png", caption="Sumá dos referencias que te conozcan", kind="form",
         edits=[
        dict(box=(54, 750, 730, 790),  text="Martín Gómez",    size=22, color=(234, 238, 244), pos=(89, 768)),
        dict(box=(54, 970, 730, 1010), text="Lucía Fernández", size=22, color=(234, 238, 244), pos=(89, 988)),
         ],
         actions=[
        dict(kind="reveal", box=(50, 748, 734, 792),    tap=(104, 770),  dur=0.55),
        dict(kind="reveal", box=(50, 858, 388, 900),    tap=(96, 879),   dur=0.5),
        dict(kind="reveal", box=(402, 858, 734, 900),   tap=(440, 879),  dur=0.5),
        dict(kind="reveal", box=(50, 968, 734, 1012),   tap=(104, 990),  dur=0.5),
        dict(kind="reveal", box=(50, 1078, 388, 1120),  tap=(96, 1099),  dur=0.5),
        dict(kind="reveal", box=(402, 1078, 734, 1120), tap=(440, 1099), dur=0.5),
    ]),
    dict(img="4.png", caption="Elegí el monto y la forma de pago", kind="form", actions=[
        dict(kind="type",   box=(32, 634, 547, 690),  tap=(74, 662),  text="150.000", dur=1.4),
        dict(kind="type",   box=(32, 815, 547, 876),  tap=(74, 843),  text="Compra de materiales", dur=1.3, multiline=True),
        dict(kind="reveal", box=(32, 915, 547, 962),  tap=(74, 938),  dur=0.75),
        dict(kind="tapbtn", box=(30, 1078, 300, 1114), tap=(80, 1095), dur=0.75),
    ]),
    dict(img="5.png", caption="Subí las fotos de tu DNI", kind="form",
         # cover the baked 'No file chosen' tooltip AND the whole original dorso label line
         cover=[(40, 497, 360, 540)],
         redraw=[(45, 514, "Foto del dorso del DNI ", 20, True, WHITE),
                 (None, 514, "* ", 20, True, (230, 120, 120)),
                 (None, 514, "— imagen o PDF, nítida", 20, False, SLATE)],
         actions=[
        dict(kind="file", box=(64, 448, 155, 496), tap=(105, 470), text="DNI_frente.jpg", chip=(182, 470), dur=1.2),
        dict(kind="file", box=(64, 537, 155, 566), tap=(105, 551), text="DNI_dorso.jpg",  chip=(182, 551), dur=1.2),
    ]),
    dict(img="6.png", caption="Aceptá los términos y enviá", kind="form", actions=[
        dict(kind="check", box=(81, 652, 100, 671), tap=(91, 661),  dur=0.5),
        dict(kind="check", box=(81, 727, 100, 746), tap=(91, 736),  dur=0.45),
        dict(kind="check", box=(81, 801, 100, 820), tap=(91, 810),  dur=0.45),
        dict(kind="check", box=(81, 876, 100, 895), tap=(91, 885),  dur=0.45),
        dict(kind="check", box=(68, 961, 87, 980),  tap=(78, 970),  dur=0.5),
        dict(kind="tapbtn", box=(50, 1028, 573, 1090), tap=(311, 1059), dur=0.85),
    ]),
    dict(img="7.png", caption="Revisá el resumen y confirmá", kind="form",
         cover=[(376, 744, 566, 790)],   # hide the baked 'No file chosen' tooltip
         edits=[
        dict(box=(266, 363, 374, 391), text="$ 150.000,00",        size=17, bold=True, color=(232, 234, 238), pos=(270, 376)),
        dict(box=(266, 393, 380, 411), text="Ciento cincuenta mil", size=13, color=(150, 164, 189), pos=(270, 401)),
        dict(box=(266, 440, 404, 463), text="$ 37.500,00 (25%)",    size=15, color=(232, 234, 238), pos=(270, 451)),
        dict(box=(266, 471, 374, 495), text="$ 187.500,00",         size=17, bold=True, color=(232, 234, 238), pos=(270, 482)),
        dict(box=(266, 506, 506, 529), text="1 pago de $ 187.500 — vence el", size=14, color=(232, 234, 238), pos=(270, 517)),
        dict(box=(265, 599, 512, 620), text="devolver ($ 9.375 por día) de atraso", size=12, color=(190, 200, 218), pos=(267, 608)),
        dict(box=(320, 674, 420, 696), text="$ 196.875",           size=15, bold=True, color=(232, 234, 238), pos=(324, 685)),
        dict(box=(320, 692, 420, 714), text="$ 215.625",           size=15, bold=True, color=(232, 234, 238), pos=(324, 703)),
        dict(box=(320, 710, 420, 732), text="$ 253.125",           size=15, bold=True, color=(232, 234, 238), pos=(324, 721)),
         ],
         actions=[
        dict(kind="tapbtn", box=(388, 606, 528, 648), tap=(457, 627), dur=1.4),
    ]),
    dict(img="8.png", caption="Enviamos tu solicitud…", kind="progress",
         bar=(105, 256, 546, 292), pct=(250, 300, 420, 348), dur=2.5),
    dict(img="9.png", caption="¡Listo! Tu solicitud quedó registrada", kind="celebrate",
         check=(304, 212), dur=3.1,
         edits=[
        dict(box=(97, 303, 170, 319),  text="Álvaro",                size=15, bold=True, color=WHITE, pos=(100, 310)),
        dict(box=(292, 369, 472, 394), text="Álvaro Gómez (C-0028)", size=16, color=(210, 218, 230), pos=(296, 381)),
        dict(box=(292, 411, 562, 436), text="alvaro.gomez@gmail.com", size=16, color=(210, 218, 230), pos=(296, 423)),
        dict(box=(292, 453, 412, 478), text="11223344",              size=16, color=(210, 218, 230), pos=(296, 465)),
        dict(box=(292, 537, 432, 562), text="$ 150.000,00",          size=16, color=(210, 218, 230), pos=(296, 549)),
        dict(box=(292, 621, 432, 646), text="$ 37.500,00",           size=16, color=(210, 218, 230), pos=(296, 633)),
        dict(box=(292, 663, 432, 688), text="$ 187.500,00",          size=16, bold=True, color=(220, 226, 236), pos=(296, 675)),
        dict(box=(254, 729, 498, 747), text="alvaro.gomez@gmail.com.", size=15, bold=True, color=(205, 214, 230), pos=(256, 737)),
         ]),
    dict(kind="outro", dur=2.8),
]

ENTER = 0.42   # cross-dissolve / entrance window at the top of each scene
HOLD = 0.55    # pause after the last action


# ── Build absolute timeline + tap/type tick schedule ─────────────────────────
def build_timeline():
    t0 = 0.0
    taps, types = [], []        # audio cue times
    success_t = None
    for sc in SCENES:
        sc["start"] = t0
        if sc["kind"] == "form":
            cur = ENTER
            for a in sc["actions"]:
                a["start"], a["end"] = cur, cur + a["dur"]
                contact = t0 + a["start"] + 0.24
                taps.append(contact)
                if a["kind"] == "type":
                    n = len(a["text"])
                    span = a["dur"] - 0.28
                    for k in range(n):
                        tt = t0 + a["start"] + 0.26 + span * (k / max(1, n))
                        if tt < t0 + a["end"]:
                            types.append(tt)
                cur = a["end"]
            sc["body"] = cur
            sc["dur"] = cur + HOLD
        elif sc["kind"] == "celebrate":
            success_t = t0 + 0.3
        t0 += sc["dur"]
    return t0, taps, types, success_t


TOTAL, TAP_TIMES, TYPE_TIMES, SUCCESS_T = build_timeline()


# ════════════════════════════════════════════════════════════════════════════
#  Assets built once
# ════════════════════════════════════════════════════════════════════════════
def _bg():
    img = vgradient(W, H, BG_TOP, BG_BOT)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, 8], fill=GOLD)
    d.rectangle([0, H - 8, W, H], fill=GOLD)
    # footer brand
    d.line([W / 2 - 54, 1756, W / 2 + 54, 1756], fill=(*GOLD, 180), width=3)
    ctext(d, W / 2, 1778, "Pedí tu préstamo 100% online", font(40, bold=True), WHITE)
    ctext(d, W / 2, 1832, "WhatsApp   ·   @impulso.credito.ars   ·   Zona Norte + AMBA", font(26), SLATE)
    return img


def _phone():
    """Transparent phone shell (bezel + island + screen cutout) over the bg."""
    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    # drop shadow
    sh = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle(
        [PX0 - 6, PY0 + 18, PX0 + PW + 6, PY0 + PH + 26], radius=PHONE_RADIUS, fill=(0, 0, 0, 150))
    layer.alpha_composite(sh.filter(ImageFilter.GaussianBlur(26)))
    d = ImageDraw.Draw(layer)
    # titanium bezel (subtle gradient via two rounded rects)
    d.rounded_rectangle([PX0, PY0, PX0 + PW, PY0 + PH], radius=PHONE_RADIUS, fill=(20, 24, 34, 255))
    d.rounded_rectangle([PX0 + 3, PY0 + 3, PX0 + PW - 3, PY0 + PH - 3], radius=PHONE_RADIUS - 3,
                        outline=(70, 80, 98, 255), width=2)
    d.rounded_rectangle([PX0 + BEZEL - 3, PY0 + BEZEL - 3, PX0 + PW - BEZEL + 3, PY0 + PH - BEZEL + 3],
                        radius=SCR_RADIUS + 3, fill=(8, 12, 22, 255))
    return layer


def _island(d):
    cx, y = W / 2, SY0 + 4
    d.rounded_rectangle([cx - 66, y, cx + 66, y + ISLAND_H + 4], radius=14, fill=(0, 0, 0, 255))
    d.ellipse([cx + 40, y + 4, cx + 56, y + 20], fill=(18, 22, 30, 255))


SCREEN_MASK = Image.new("L", (SW, SH), 0)
ImageDraw.Draw(SCREEN_MASK).rounded_rectangle([0, 0, SW - 1, SH - 1], radius=SCR_RADIUS, fill=255)


def _finger():
    """Cropped 👆 emoji sprite + fingertip offset (tip = top-centre)."""
    sz = 160
    f = ImageFont.truetype("/System/Library/Fonts/Apple Color Emoji.ttc", sz)
    big = Image.new("RGBA", (sz * 2, sz * 2), (0, 0, 0, 0))
    ImageDraw.Draw(big).text((sz * 0.5, sz * 0.4), "\U0001F446", font=f, embedded_color=True)
    spr = big.crop(big.getbbox())
    scale = 150 / spr.height
    spr = spr.resize((int(spr.width * scale), int(spr.height * scale)), Image.LANCZOS)
    spr = spr.rotate(-16, expand=True, resample=Image.BICUBIC)
    tip = (spr.width * 0.40, spr.height * 0.03)   # index-tip inside the rotated sprite
    return spr, tip


BG = _bg().convert("RGBA")
PHONE = _phone()
FINGER, FINGER_TIP = _finger()


# ════════════════════════════════════════════════════════════════════════════
#  Per-scene native-resolution helpers
# ════════════════════════════════════════════════════════════════════════════
def modal_color(arr, box):
    x0, y0, x1, y1 = box
    reg = arr[y0:y1, x0:x1].reshape(-1, 3)
    if len(reg) == 0:
        return (24, 40, 72)
    q = (reg // 4 * 4)
    vals, counts = np.unique(q, axis=0, return_counts=True)
    return tuple(int(v) for v in vals[counts.argmax()])


def prep_scene(sc):
    """Load the screenshot, retouch baked artifacts, compute a uniform
    full-width placement (+ vertical pan when the form is taller than the
    screen), and build the 'empty' base (fields blanked with their own fill)."""
    if sc["kind"] == "outro":
        return
    im = Image.open(SRC / sc["img"]).convert("RGB")
    arr0 = np.asarray(im)
    # retouch baked browser artifacts (e.g. the grey 'No file chosen' tooltip)
    if sc.get("cover"):
        cd = ImageDraw.Draw(im)
        for (x0, y0, x1, y1) in sc["cover"]:
            col = modal_color(arr0, (max(0, x0 - 24), y0, max(1, x0 - 6), y1))
            cd.rectangle([x0, y0, x1, y1], fill=col)
    # redraw any label text that the cover wiped out (x=None continues the previous run)
    if sc.get("redraw"):
        rd = ImageDraw.Draw(im)
        cx = 0
        for (x, y, text, size, bold, color) in sc["redraw"]:
            xx = cx if x is None else x
            rd.text((xx, y), text, font=font(size, bold=bold), fill=color, anchor="lm")
            cx = xx + rd.textlength(text, font=font(size, bold=bold))
    # replace baked values (name / email / DNI / amounts): cover each value box
    # with its own sampled background colour, then draw the new text on top.
    if sc.get("edits"):
        ed2 = ImageDraw.Draw(im)
        for e in sc["edits"]:
            x0, y0, x1, y1 = e["box"]
            col = e.get("bg") or modal_color(arr0, (x0, y0, x1, y1))
            ed2.rectangle([x0, y0, x1, y1], fill=col)
            anchor = e.get("anchor", "lm")
            ax = x0 + e.get("pad", 0) if anchor[0] == "l" else (x1 if anchor[0] == "r" else (x0 + x1) / 2)
            pos = e.get("pos") or (ax, (y0 + y1) / 2)
            ed2.text(pos, e["text"], font=font(e.get("size", 22), bold=e.get("bold", False)),
                     fill=e.get("color", WHITE), anchor=anchor)
    sc["orig"] = im
    arr = np.asarray(im)
    iw, ih = im.size
    sc["bg"] = tuple(int(c) for c in arr[3, 3])
    empty = im.copy()
    ed = ImageDraw.Draw(empty)
    for a in sc.get("actions", []):
        if a["kind"] in ("type", "reveal"):
            x0, y0, x1, y1 = a["box"]
            col = modal_color(arr, (x0 + 4, y0 + 3, x1 - 4, y1 - 3))
            ed.rectangle([x0 + 2, y0 + 2, x1 - 2, y1 - 2], fill=col)
            a["fill"] = col
    sc["empty"] = empty
    # ── uniform full-width placement ──
    scale = SW / iw
    disp_h = int(round(ih * scale))
    sc["scale"] = scale
    sc["disp_h"] = disp_h
    if disp_h <= SH:                       # fits: centre vertically, no pan
        sc["pan"] = False
        sc["ctop0"] = (SH - disp_h) // 2
    else:                                  # taller than screen: pan to each field
        sc["pan"] = True
        sc["maxscroll"] = disp_h - SH
        tg = []
        for a in sc.get("actions", []):
            fy = a["tap"][1] * scale
            tg.append(clamp(fy - SH * 0.55, 0, sc["maxscroll"]))
        sc["atargets"] = tg or [0]


def scene_ctop(sc, lt):
    """Vertical offset (screen-space y where native y=0 lands), inc. pan."""
    if not sc.get("pan"):
        return sc.get("ctop0", 0)
    acts = sc.get("actions", [])
    tg = sc["atargets"]
    if not acts or lt < acts[0]["start"]:
        return -0.0
    i = 0
    for j, a in enumerate(acts):
        if lt >= a["start"]:
            i = j
    prevT = tg[i - 1] if i > 0 else 0.0
    g = ease(clamp((lt - acts[i]["start"]) / 0.4))
    return -lerp(prevT, tg[i], g)


def to_disp(sc, nx, ny, ctop):
    """native px -> full-canvas px, given the current vertical offset."""
    return (SX0 + nx * sc["scale"], SY0 + ctop + ny * sc["scale"])


# ── field drawing (native coords, onto a copy of the screenshot) ─────────────
def draw_type(d, a, p, img):
    x0, y0, x1, y1 = a["box"]
    # cap the effective height so a tall textarea (Notas) doesn't get giant text
    fs = int(min(y1 - y0, 58) * 0.44)
    f = font(fs)
    txt = a["text"]
    n = int(round(p * len(txt)))
    shown = txt[:n]
    cx = x0 + 22
    cy = (y0 + y1) / 2 - 2
    if a.get("multiline"):
        cy = y0 + fs * 0.75          # anchor near the textarea's top-left
    d.text((cx, cy), shown, font=f, fill=WHITE, anchor="lm" if not a.get("multiline") else "la")
    if p < 1.0:  # caret
        cw = d.textlength(shown, font=f)
        caret_on = (int(p * len(txt) * 2) % 2 == 0)
        if caret_on:
            cxp = cx + cw + 2
            d.line([cxp, cy - fs * 0.5, cxp, cy + fs * 0.5], fill=GOLD, width=3)


def draw_reveal(a, p, empty, orig):
    x0, y0, x1, y1 = a["box"]
    w = int((x1 - x0) * ease_out(p))
    if w <= 0:
        return
    region = orig.crop((x0, y0, x0 + w, y1))
    empty.paste(region, (x0, y0))
    if p < 1.0:  # glow at the typing edge
        dd = ImageDraw.Draw(empty)
        ex = x0 + w
        dd.line([ex, y0 + 4, ex, y1 - 4], fill=GOLD, width=3)


def draw_check(d, a, p):
    x0, y0, x1, y1 = a["box"]
    r = (x1 - x0) * 0.28
    d.rounded_rectangle([x0, y0, x1, y1], radius=6, fill=(*GOLD, int(255 * clamp(p * 1.4))))
    # check stroke grows
    pts = [(x0 + (x1 - x0) * 0.22, y0 + (y1 - y0) * 0.52),
           (x0 + (x1 - x0) * 0.42, y0 + (y1 - y0) * 0.72),
           (x0 + (x1 - x0) * 0.80, y0 + (y1 - y0) * 0.28)]
    seg = clamp(p)
    if seg <= 0:
        return
    if seg < 0.5:
        q = seg / 0.5
        a2 = (lerp(pts[0][0], pts[1][0], q), lerp(pts[0][1], pts[1][1], q))
        d.line([pts[0], a2], fill=(20, 30, 50), width=5)
    else:
        q = (seg - 0.5) / 0.5
        a2 = (lerp(pts[1][0], pts[2][0], q), lerp(pts[1][1], pts[2][1], q))
        d.line([pts[0], pts[1], a2], fill=(20, 30, 50), width=5, joint="curve")


def draw_filechip(d, a, p):
    cx, cy = a["chip"]
    txt = a["text"]
    f = font(21, bold=True)
    tw = d.textlength(txt, font=f) + 44
    x0 = cx; y0 = cy - 18
    alpha = int(255 * ease_out(p))
    d.rounded_rectangle([x0, y0, x0 + tw, y0 + 36], radius=18, fill=(40, 110, 70, alpha))
    d.text((x0 + 14, cy), "✓", font=font(20, bold=True), fill=(170, 240, 180, alpha), anchor="lm")
    d.text((x0 + 36, cy), txt, font=f, fill=(*WHITE, alpha), anchor="lm")


def render_native(sc, progresses):
    """Build the screenshot content for this scene at the given per-action
    progress list. Returns an RGB PIL image at native resolution."""
    if sc["kind"] in ("progress", "celebrate"):
        img = sc["orig"].copy()
    else:
        img = sc["empty"].copy()
    d = ImageDraw.Draw(img, "RGBA")
    for a, p in zip(sc.get("actions", []), progresses):
        if p <= 0:
            continue
        if a["kind"] == "type":
            draw_type(d, a, p, img)
        elif a["kind"] == "reveal":
            draw_reveal(a, p, img, sc["orig"]); d = ImageDraw.Draw(img, "RGBA")
        elif a["kind"] == "check":
            draw_check(d, a, p)
        elif a["kind"] == "file":
            draw_filechip(d, a, p)
        # tapbtn: no content change
    return img


def progress_bar_scene(sc, lt):
    """Scene 8: redraw the gold progress bar + % counter, filling 0→100%."""
    img = sc["orig"].copy()
    d = ImageDraw.Draw(img, "RGBA")
    bx0, by0, bx1, by1 = sc["bar"]
    px0, py0, px1, py1 = sc["pct"]
    # cover the baked bar + percentage with the LOCAL card background
    arr = np.asarray(sc["orig"])
    strips = np.concatenate([arr[by0 - 30:by0 - 12, bx0:bx1].reshape(-1, 3),
                             arr[by1 + 14:by1 + 30, bx0:bx1].reshape(-1, 3)])
    cov = tuple(int(c) for c in np.median(strips, axis=0))
    iw = sc["orig"].width
    d.rectangle([28, by0 - 12, iw - 28, by1 + 12], fill=cov)
    d.rectangle([px0 - 40, py0 - 8, px1 + 40, py1 + 10], fill=cov)
    p = ease_out(clamp((lt - 0.3) / (sc["dur"] - 0.6)))
    # track
    d.rounded_rectangle([bx0, by0, bx1, by1], radius=(by1 - by0) / 2, fill=(255, 255, 255, 40))
    fw = (bx1 - bx0) * p
    if fw > 6:
        d.rounded_rectangle([bx0, by0, bx0 + fw, by1], radius=(by1 - by0) / 2, fill=(*GOLD, 255))
        # candy stripes
        stripe = Image.new("RGBA", img.size, (0, 0, 0, 0))
        sd = ImageDraw.Draw(stripe)
        for xx in range(int(bx0) - 40, int(bx0 + fw), 22):
            sd.polygon([(xx, by1), (xx + 11, by1), (xx + 11 + 18, by0), (xx + 18, by0)], fill=(255, 255, 255, 46))
        sm = Image.new("L", img.size, 0)
        ImageDraw.Draw(sm).rounded_rectangle([bx0, by0, bx0 + fw, by1], radius=(by1 - by0) / 2, fill=255)
        img.paste(Image.alpha_composite(img.convert("RGBA").crop((0, 0, *img.size)), stripe).convert("RGB"), (0, 0), sm)
        d = ImageDraw.Draw(img, "RGBA")
    pct = int(round(p * 100))
    cx = (px0 + px1) / 2
    d.text((cx, (py0 + py1) / 2), f"{pct}%", font=font(34, bold=True), fill=GOLD, anchor="mm")
    return img


# ════════════════════════════════════════════════════════════════════════════
#  Compositing a full frame
# ════════════════════════════════════════════════════════════════════════════
def place_tile(content_native, sc, ctop):
    """Scale native content to full screen width and place it at vertical
    offset `ctop` (negative = panned up), returning an SW×SH screen tile."""
    tile = Image.new("RGB", (SW, SH), sc["bg"])
    dh = sc["disp_h"]
    tile.paste(content_native.resize((SW, dh), Image.LANCZOS), (0, int(round(ctop))))
    return tile.convert("RGBA")


def scene_index(t):
    for i, sc in enumerate(SCENES):
        if t < sc["start"] + sc["dur"]:
            return i
    return len(SCENES) - 1


def draw_caption(frame, text, alpha):
    if alpha <= 0:
        return
    d = ImageDraw.Draw(frame, "RGBA")
    f = font(46, bold=True)
    # wrap to <= ~26 chars
    words, lines, cur = text.split(), [], ""
    for wd in words:
        if len(cur + " " + wd) <= 27:
            cur = (cur + " " + wd).strip()
        else:
            lines.append(cur); cur = wd
    if cur:
        lines.append(cur)
    y = 300 - (len(lines) - 1) * 28
    ctext(d, W / 2, 196, "C Ó M O   U S A R L O", font(28, bold=True), GOLD, a=alpha, tracking=3)
    for ln in lines:
        ctext(d, W / 2, y, ln, f, WHITE, a=alpha)
        y += 56


def draw_progressbar(frame, idx):
    d = ImageDraw.Draw(frame, "RGBA")
    n = len(SCENES) - 1  # exclude outro from the dots
    bw, y, gap, h = 900, 92, 7, 7
    x0 = (W - bw) / 2
    sw = (bw - (n - 1) * gap) / n
    cur = SCENES[idx]
    local = clamp((_T - cur["start"]) / cur["dur"]) if idx < n else 1.0
    for k in range(n):
        x = x0 + k * (sw + gap)
        d.rounded_rectangle([x, y, x + sw, y + h], radius=4, fill=(255, 255, 255, 60))
        fp = 1.0 if k < idx else (local if k == idx else 0.0)
        if fp > 0:
            d.rounded_rectangle([x, y, x + sw * fp, y + h], radius=4, fill=(*GOLD, 255))


def draw_finger(frame, pos, press, ripple):
    x, y = pos
    scale = 1.0 - 0.06 * press
    spr = FINGER if scale == 1.0 else FINGER.resize(
        (max(1, int(FINGER.width * scale)), max(1, int(FINGER.height * scale))), Image.LANCZOS)
    tipx, tipy = FINGER_TIP[0] * scale, FINGER_TIP[1] * scale
    py = y - 6 * press
    # ripple (subtle, so it never hides the field value underneath)
    if ripple is not None:
        rp, ra = ripple
        rr = 12 + 30 * rp
        rl = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        ImageDraw.Draw(rl).ellipse([x - rr, y - rr, x + rr, y + rr],
                                   outline=(*GOLD, int(95 * ra)), width=3)
        frame.alpha_composite(rl)
    # contact shadow
    shl = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(shl).ellipse([x - 26, py - 10, x + 26, py + 14], fill=(0, 0, 0, 90))
    frame.alpha_composite(shl.filter(ImageFilter.GaussianBlur(7)))
    frame.alpha_composite(spr, (int(x - tipx), int(py - tipy)))


# ── finger / ripple state for a form scene at local time lt ─────────────────
def finger_state(sc, lt, ctop):
    acts = sc["actions"]
    disp_taps = [to_disp(sc, a["tap"][0], a["tap"][1], ctop) for a in acts]
    # find current action by lt
    i = 0
    for j, a in enumerate(acts):
        if lt >= a["start"]:
            i = j
    a = acts[i]
    # entrance glide from below to first tap
    if lt < acts[0]["start"]:
        q = ease(clamp(lt / acts[0]["start"])) if acts[0]["start"] > 0 else 1
        start = (disp_taps[0][0], H + 120)
        pos = (lerp(start[0], disp_taps[0][0], q), lerp(start[1], disp_taps[0][1], q))
        return pos, 0.0, None
    # glide from previous tap to this one over first 0.22s of the action
    g = clamp((lt - a["start"]) / 0.22)
    prev = disp_taps[i - 1] if i > 0 else disp_taps[0]
    pos = (lerp(prev[0], disp_taps[i][0], ease(g)), lerp(prev[1], disp_taps[i][1], ease(g)))
    # press bob right after arrival
    pr = clamp((lt - a["start"] - 0.20) / 0.12) * clamp((a["start"] + 0.40 - lt) / 0.12)
    press = math.sin(clamp(pr) * math.pi)
    # ripple for ~0.5s after contact
    rt = lt - (a["start"] + 0.22)
    ripple = (clamp(rt / 0.5), 1 - clamp(rt / 0.5)) if 0 <= rt <= 0.5 else None
    return pos, press, ripple


# ════════════════════════════════════════════════════════════════════════════
#  Master frame renderer
# ════════════════════════════════════════════════════════════════════════════
_T = 0.0


def render(t):
    global _T
    _T = t
    idx = scene_index(t)
    sc = SCENES[idx]
    lt = t - sc["start"]
    frame = BG.copy()

    if sc["kind"] == "outro":
        _render_outro(frame, lt, sc["dur"])
        draw_progressbar(frame, idx)
        return frame

    # ---- screen content tile (uniform width + vertical pan) ----
    ctop = scene_ctop(sc, lt)
    if sc["kind"] == "form":
        progs = [clamp((lt - a["start"]) / a["dur"]) for a in sc["actions"]]
        content = render_native(sc, progs)
    elif sc["kind"] == "progress":
        content = progress_bar_scene(sc, lt)
    else:  # celebrate
        content = sc["orig"]
    tile = place_tile(content, sc, ctop)

    # ---- transition: fade the screen through the app's navy (no double-text) ----
    scr_a = ease(clamp(lt / 0.24)) * ease(clamp((sc["dur"] - lt) / 0.20))
    if scr_a < 1.0:
        tile = Image.blend(Image.new("RGBA", (SW, SH), (*sc["bg"], 255)), tile, scr_a)

    # phone + screen
    frame.alpha_composite(PHONE)
    frame.paste(tile, (SX0, SY0), SCREEN_MASK)
    _island(ImageDraw.Draw(frame, "RGBA"))
    frame.alpha_composite(_SCREEN_GLOSS)

    # ---- finger (form scenes only, while screen is fully shown) ----
    if sc["kind"] == "form" and 0.26 < lt < sc["dur"] - 0.18:
        pos, press, ripple = finger_state(sc, lt, ctop)
        draw_finger(frame, pos, press, ripple)

    # celebrate sparkles
    if sc["kind"] == "celebrate":
        _celebrate_overlay(frame, sc, lt, ctop)

    # ---- caption + progress dots ----
    cap_a = ease(clamp(lt / 0.35)) * ease(clamp((sc["dur"] - lt) / 0.3))
    draw_caption(frame, sc["caption"], cap_a)
    draw_progressbar(frame, idx)

    # global intro/outro fade
    fade = max(clamp((0.4 - t) / 0.4), clamp((t - (TOTAL - 0.4)) / 0.4))
    if fade > 0:
        frame.alpha_composite(Image.new("RGBA", (W, H), (*NAVY, int(255 * fade))))
    return frame


def _build_gloss():
    g = Image.new("RGBA", (SW, SH), (0, 0, 0, 0))
    gd = ImageDraw.Draw(g)
    gd.rounded_rectangle([0, 0, SW - 1, SH - 1], radius=SCR_RADIUS, outline=(*GOLD, 60), width=2)
    out = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    out.paste(g, (SX0, SY0), SCREEN_MASK)
    return out


_SCREEN_GLOSS = _build_gloss()


# ── celebrate helpers ────────────────────────────────────────────────────────
def _celebrate_overlay(frame, sc, lt, ctop):
    # let the success screen + logo render clean for the first ~0.4s
    if lt < 0.4:
        return
    lt2 = lt - 0.4
    cx, cy = to_disp(sc, sc["check"][0], sc["check"][1], ctop)
    # sparkle burst radiating from the green check
    k = clamp(lt2 / 1.1)
    if k < 1:
        r = 30 + 140 * ease_out(k)
        a = int(190 * (1 - k))
        lay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        ld = ImageDraw.Draw(lay)
        for ang in range(0, 360, 30):
            rad = math.radians(ang)
            x = cx + r * math.cos(rad); y = cy + r * math.sin(rad)
            ld.ellipse([x - 5, y - 5, x + 5, y + 5], fill=(*GOLD, a))
        frame.alpha_composite(lay)
    # confetti falling (sparse, fades in, clears the logo band up top)
    conf = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    cd = ImageDraw.Draw(conf)
    fade = int(150 * clamp(lt2 / 0.4))
    cols = [GOLD, BLUE, (120, 220, 150), WHITE]
    for i in range(38):
        seed = (i * 97) % 1000
        x = PX0 + 40 + (seed * 7) % (PW - 80)
        fall = ((lt2 * (150 + seed % 90)) + seed) % (PH + 200)
        y = PY0 - 40 + fall
        if PY0 + 150 < y < PY0 + PH:          # keep clear of the logo band up top
            c = cols[i % 4]
            s = 5 + seed % 4
            cd.rectangle([x, y, x + s, y + s + 2], fill=(*c, fade))
    # clip confetti to screen
    cm = Image.new("L", (W, H), 0)
    ImageDraw.Draw(cm).rounded_rectangle([SX0, SY0, SX0 + SW, SY0 + SH], radius=SCR_RADIUS, fill=255)
    frame.paste(Image.alpha_composite(frame, conf), (0, 0), cm)


def _render_outro(frame, lt, dur):
    a = ease(clamp(lt / 0.5)) * ease(clamp((dur - lt) / 0.4))
    badge = ic_badge(190)
    bx = int(W / 2 - badge.width / 2)
    tmp = frame.copy()
    tmp.alpha_composite(badge, (bx, 560))
    d = ImageDraw.Draw(tmp, "RGBA")
    # two-tone title
    f = font(70, bold=True)
    t1, t2 = "Impulso ", "Crédito"
    tot = d.textlength(t1, font=f) + d.textlength(t2, font=f)
    x = W / 2 - tot / 2
    d.text((x, 800), t1, font=f, fill=BLUE); x += d.textlength(t1, font=f)
    d.text((x, 800), t2, font=f, fill=GOLD)
    ctext(d, W / 2, 920, "Pedí tu préstamo en minutos", font(40), SLATE)
    d.rounded_rectangle([W / 2 - 300, 1030, W / 2 + 300, 1120], radius=45, fill=GOLD)
    ctext(d, W / 2, 1052, "Empezá ahora  →", font(42, bold=True), (20, 30, 50))
    ctext(d, W / 2, 1230, "@impulso.credito.ars", font(38, bold=True), WHITE)
    frame.alpha_composite(Image.blend(Image.new("RGBA", (W, H), (0, 0, 0, 0)), tmp, a))


# ════════════════════════════════════════════════════════════════════════════
#  Audio
# ════════════════════════════════════════════════════════════════════════════
SR = 44100


def _adsr(n, a, d, s, r, sl=0.75):
    if n <= 0:
        return np.zeros(max(n, 0))
    ai = min(int(a * SR), n)
    ri = min(int(r * SR), n - ai)
    di = min(int(d * SR), max(0, n - ai - ri))
    env = np.ones(n) * sl
    if ai: env[:ai] = np.linspace(0, 1, ai)
    if di: env[ai:ai + di] = np.linspace(1, sl, di)
    if ri: env[n - ri:] = np.linspace(sl, 0, ri)
    return env


def synth_audio(total, taps, types, success_t):
    n = int(total * SR)
    L = np.zeros(n); R = np.zeros(n)
    t = np.arange(n) / SR

    def note(freq, t0, dur, amp, partials=(1, 0.5, 0.25), a=0.02, d=0.1, s=0.8, r=0.3, vib=0.0):
        i0 = int(t0 * SR); i1 = min(n, int((t0 + dur) * SR))
        if i1 <= i0: return None, 0, 0
        m = i1 - i0
        tt = np.arange(m) / SR
        sig = np.zeros(m)
        vibr = (1 + vib * np.sin(2 * np.pi * 5 * tt)) if vib else 1.0
        for k, pa in enumerate(partials, start=1):
            sig += pa * np.sin(2 * np.pi * freq * k * tt * vibr)
        sig *= _adsr(m, a, d, s, r, 0.75) * amp
        return sig, i0, i1

    def add(freq, t0, dur, amp, **kw):
        pan = kw.pop("pan", 0.0)
        sig, i0, i1 = note(freq, t0, dur, amp, **kw)
        if sig is None: return
        L[i0:i1] += sig * (0.5 - 0.5 * pan)
        R[i0:i1] += sig * (0.5 + 0.5 * pan)

    # ── warm pad: I–V–vi–IV in A major, 2s per chord, looped ──
    A = 220.0
    def nt(semi): return A * 2 ** (semi / 12)
    chords = [[nt(0), nt(4), nt(7)],       # A
              [nt(-5), nt(-1), nt(2)],     # E
              [nt(-3), nt(0), nt(4)],      # F#m-ish
              [nt(-7), nt(-3), nt(0)]]     # D
    bar = 2.0
    tpos = 0.0; ci = 0
    while tpos < total:
        for f in chords[ci % 4]:
            add(f, tpos, bar + 0.4, 0.05, partials=(1, 0.4, 0.2, 0.12),
                a=0.5, d=0.4, s=0.75, r=0.6, vib=0.004, pan=0.0)
            add(f * 2, tpos, bar + 0.4, 0.018, a=0.6, d=0.4, s=0.6, r=0.6)
        tpos += bar; ci += 1

    # ── soft marimba/bell arpeggio (A major pentatonic), light ──
    penta = [nt(0), nt(4), nt(7), nt(9), nt(12), nt(9), nt(7), nt(4)]
    step = 0.5
    for i, tt in enumerate(np.arange(0.5, total - 0.5, step)):
        f = penta[i % len(penta)]
        add(f * 2, tt, 0.45, 0.035, partials=(1, 0.6, 0.3), a=0.005, d=0.25, s=0.0, r=0.2,
            pan=0.3 * math.sin(i * 0.7))

    # ── soft UI tap ticks ──
    for tt in taps:
        add(1650, tt, 0.05, 0.05, partials=(1, 0.5), a=0.001, d=0.04, s=0.0, r=0.01)
    for tt in types:
        add(2200, tt, 0.025, 0.018, partials=(1,), a=0.001, d=0.02, s=0.0, r=0.005)

    # ── success chime ──
    if success_t is not None:
        for k, semi in enumerate([12, 16, 19, 24]):
            add(nt(semi) * 1.0, success_t + k * 0.09, 1.2, 0.07,
                partials=(1, 0.5, 0.25), a=0.005, d=0.5, s=0.2, r=0.6, pan=0.0)

    # ── master: gentle FIR low-pass (numpy), soft limit, fades ──
    def smooth(x, k=24):
        win = np.hanning(k); win /= win.sum()
        return np.convolve(x, win, mode="same")
    L = smooth(L); R = smooth(R)

    mx = max(np.max(np.abs(L)), np.max(np.abs(R)), 1e-6)
    L = np.tanh(L / mx * 1.4) * 0.6
    R = np.tanh(R / mx * 1.4) * 0.6
    # fades
    fi = int(1.2 * SR); fo = int(1.6 * SR)
    L[:fi] *= np.linspace(0, 1, fi); R[:fi] *= np.linspace(0, 1, fi)
    L[-fo:] *= np.linspace(1, 0, fo); R[-fo:] *= np.linspace(1, 0, fo)

    stereo = np.stack([L, R], axis=1)
    pcm = (stereo * 32767).astype(np.int16)
    wav = HERE / "_howto_audio.wav"
    with wave.open(str(wav), "w") as w:
        w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR)
        w.writeframes(pcm.tobytes())
    return wav


# ════════════════════════════════════════════════════════════════════════════
#  Drivers
# ════════════════════════════════════════════════════════════════════════════
def prep_all():
    for sc in SCENES:
        prep_scene(sc)


def probe():
    prep_all()
    for i, sc in enumerate(SCENES):
        # pick a representative time: near the end of the scene body
        if sc["kind"] == "form":
            tt = sc["start"] + sc["body"] * 0.85
        elif sc["kind"] == "progress":
            tt = sc["start"] + sc["dur"] * 0.7
        else:
            tt = sc["start"] + sc["dur"] * 0.55
        fr = render(tt).convert("RGB")
        fr.save(HERE / f"_probe_{i+1}.png")
        print(f"_probe_{i+1}.png  (scene {sc.get('img','outro')}, t={tt:.1f}s)")
    print(f"TOTAL {TOTAL:.1f}s")


def main():
    prep_all()
    print(f"Rendering {TOTAL:.1f}s  ({int(TOTAL*FPS)} frames)…")
    silent = HERE / "_howto_silent.mp4"
    writer = imageio.get_writer(silent, fps=FPS, codec="libx264", macro_block_size=8,
                                pixelformat="yuv420p",
                                ffmpeg_params=["-crf", "19", "-preset", "medium"])
    nframes = int(TOTAL * FPS)
    for i in range(nframes):
        writer.append_data(np.asarray(render(i / FPS).convert("RGB")))
        if i % 60 == 0:
            print(f"  {i}/{nframes}")
    writer.close()

    print("Synthesizing audio…")
    wav = synth_audio(TOTAL, TAP_TIMES, TYPE_TIMES, SUCCESS_T)

    print("Muxing…")
    ff = imageio_ffmpeg.get_ffmpeg_exe()
    out = HERE / "story_como_usar.mp4"
    subprocess.run([ff, "-y", "-i", str(silent), "-i", str(wav),
                    "-c:v", "copy", "-c:a", "aac", "-b:a", "192k", "-shortest", str(out)],
                   check=True, capture_output=True)
    print("Done →", out)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "probe":
        probe()
    else:
        main()
