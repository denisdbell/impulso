#!/usr/bin/env python3
"""
Impulso Crédito — caption generator for the social posts.

For every post_XX.png in this folder it writes a matching caption:
a tailored Spanish hook + a fixed contact block + hashtags.
Outputs one post_XX.txt per post (ready to paste) and a combined captions.md.

Run:  python3 generate_captions.py
"""

from pathlib import Path

HERE = Path(__file__).parent

# ── Fixed block included in EVERY caption ────────────────────────────────────
CONTACT = (
    "📲 WhatsApp: +541139255583\n"
    "👇 Info y solicitudes acá\n"
    "https://tinyurl.com/Impulso-prospecto"
)

HASHTAGS = (
    "#ImpulsoCredito #Prestamos #PrestamosPersonales #PrestamosRapidos "
    "#PrestamosOnline #CreditoRapido #DineroYa #ZonaNorte #AMBA "
    "#BuenosAires #Argentina #Finanzas #FinanzasPersonales "
    "#PrestamosArgentina #CreditosPersonales"
)

# ── Tailored caption bodies, in the same order as post_01 … post_10 ──────────
BODIES = [
    # 01 · Velocidad
    "⏱️ ¿Necesitás plata ya? En Impulso Crédito aprobamos rápido y te "
    "transferimos el MISMO DÍA. Sin vueltas ni esperas eternas.",
    # 02 · 100% online
    "📱 Pedí tu préstamo 100% online, desde donde estés. Solo necesitás tu "
    "correo y tu DNI. Sin filas, sin papeleos.",
    # 03 · Atención
    "💬 Te escuchamos. Consultá SIN compromiso y resolvé todo por WhatsApp, "
    "con atención real y cercana.",
    # 04 · Clientes
    "🔁 ¿Ya sos cliente? La próxima es más fácil: cargamos tus datos por vos "
    "y pedís tu préstamo en minutos.",
    # 05 · Flexibilidad
    "🎯 Elegí el plazo que mejor se adapte a tu bolsillo. Préstamos pensados "
    "para vos, no al revés.",
    # 06 · Confianza
    "⭐ Pagá en fecha y crecé con nosotros: accedé a montos cada vez más altos. "
    "La buena conducta se premia.",
    # 07 · Cobertura
    "📍 Acompañamos a los vecinos de Zona Norte y el AMBA. Estamos cerca "
    "cuando más lo necesitás.",
    # 08 · Transparencia
    "🔎 Nada de sorpresas. Desde el primer día sabés cuánto y cuándo vas a "
    "pagar. Transparencia total, sin letra chica.",
    # 09 · Pagos
    "💳 Recibí tu préstamo por Mercado Pago o transferencia bancaria. "
    "Vos elegís cómo te queda más cómodo.",
    # 10 · Impulso
    "🚀 Tu próximo impulso está a un mensaje de distancia. Simple, rápido y "
    "confiable: así es Impulso Crédito.",
]


def caption_for(body):
    return f"{body}\n\n{CONTACT}\n\n{HASHTAGS}\n"


def main():
    posts = sorted(HERE.glob("post_*.png"))
    if not posts:
        print("No post_*.png files found in", HERE)
        return
    if len(posts) != len(BODIES):
        print(f"⚠ {len(posts)} posts but {len(BODIES)} caption bodies — "
              "matching by order up to the shorter list.")

    md = ["# Impulso Crédito — descripciones de los posts\n"]
    for post, body in zip(posts, BODIES):
        caption = caption_for(body)
        txt = post.with_suffix(".txt")
        txt.write_text(caption, encoding="utf-8")
        md.append(f"## {post.name}\n\n```\n{caption}```\n")
        print("wrote", txt.name)

    (HERE / "captions.md").write_text("\n".join(md), encoding="utf-8")
    print("wrote captions.md")


if __name__ == "__main__":
    main()
