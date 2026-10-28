#!/usr/bin/env python3
"""
Impulso Crédito — read existing Instagram posts and attach descriptions.

IMPORTANT — Instagram API limitation:
  A published post's CAPTION cannot be edited via the API (media is immutable;
  captions are editable only by hand in the app). The only programmatic write
  on a published post is adding a COMMENT, and comments are limited to
  ≤300 chars, ≤1 URL, ≤4 hashtags, not all-caps.

So this script:
  1. READS your published posts from graph.instagram.com (/me/media).
  2. MATCHES each one to the right local post image (perceptual hash vs
     post_XX.png) and therefore to the right caption.
  3. Writes update_plan.md — every post's permalink + the full caption to
     PASTE MANUALLY (the only way to set the real caption + all hashtags).
  4. With --comment: posts a compliant first comment (WhatsApp + link + a few
     hashtags) via the API. Needs the instagram_business_manage_comments
     permission on the token; it skips posts that already have our link.

Run:   python3 update_posts.py            # read + build the paste plan
       python3 update_posts.py --comment  # also post the CTA as a comment
"""

import io
import os
import sys
from pathlib import Path

import numpy as np
import requests
from dotenv import load_dotenv
from PIL import Image

from generate_captions import BODIES, CONTACT, caption_for

HERE = Path(__file__).parent
# The Instagram token lives in the instagram/ project; fall back to a local .env.
load_dotenv(HERE.parent / "instagram" / ".env")
load_dotenv(HERE / ".env")
IG_API = "https://graph.instagram.com/v21.0"
TOKEN = os.environ.get("IG_ACCESS_TOKEN", "")

# A comment can hold ≤1 URL and ≤4 hashtags — so just the CTA + 3 tags.
COMMENT_TEXT = CONTACT + "\n\n#ImpulsoCredito #PrestamosRapidos #ZonaNorte"
LINK_MARKER = "tinyurl.com/Impulso-prospecto"


# ── perceptual (average) hash ────────────────────────────────────────────────
def ahash(img):
    g = np.asarray(img.convert("L").resize((16, 16), Image.LANCZOS), dtype=np.float32)
    return (g > g.mean()).flatten()

def hamming(a, b):
    return int(np.count_nonzero(a != b))


# ── Instagram reads ──────────────────────────────────────────────────────────
def fetch_media():
    out, url = [], f"{IG_API}/me/media"
    params = {"fields": "id,caption,media_type,permalink,timestamp,media_url",
              "limit": 50, "access_token": TOKEN}
    while url:
        r = requests.get(url, params=params, timeout=30)
        if not r.ok:
            print("ERROR reading media:", r.status_code, r.text[:400]); sys.exit(1)
        j = r.json()
        out.extend(j.get("data", []))
        url = j.get("paging", {}).get("next")
        params = None  # the 'next' URL already carries params + token
    return out

def existing_comment_has_link(media_id):
    r = requests.get(f"{IG_API}/{media_id}/comments",
                     params={"fields": "text", "access_token": TOKEN}, timeout=30)
    if not r.ok:
        return False  # can't read (likely missing permission) → don't block
    return any(LINK_MARKER in (c.get("text") or "") for c in r.json().get("data", []))

def post_comment(media_id, text):
    r = requests.post(f"{IG_API}/{media_id}/comments",
                      data={"message": text, "access_token": TOKEN}, timeout=30)
    if not r.ok:
        print("   ✗ comment failed:", r.status_code, r.text[:300]); return False
    return True


# ── main ─────────────────────────────────────────────────────────────────────
def main(do_comment):
    if not TOKEN:
        print("No IG_ACCESS_TOKEN in .env"); return

    # local post hashes → caption index
    locals_ = sorted(HERE.glob("post_*.png"))
    if not locals_:
        print("No local post_*.png to match against."); return
    local_hash = {p: ahash(Image.open(p)) for p in locals_}

    media = [m for m in fetch_media() if m.get("media_type") == "IMAGE" and m.get("media_url")]
    print(f"Found {len(media)} published image post(s).\n")

    rows, plan = [], ["# Impulso Crédito — plan para actualizar descripciones\n",
                      "> La API de Instagram **no** permite editar la descripción de un post "
                      "ya publicado. Copiá y pegá a mano la descripción en cada post "
                      "(Editar ▸ pegar). El número en el nombre del archivo indica a qué "
                      "imagen corresponde.\n"]
    for m in media:
        try:
            img = Image.open(io.BytesIO(requests.get(m["media_url"], timeout=30).content))
        except Exception as e:
            print("  skip (download failed):", m["id"], e); continue
        h = ahash(img)
        best = min(local_hash, key=lambda p: hamming(h, local_hash[p]))
        dist = hamming(h, local_hash[best])
        idx = int(best.stem.split("_")[1])
        caption = caption_for(BODIES[idx - 1])
        tag = "match" if dist <= 40 else f"WEAK match (d={dist})"
        rows.append((m, best, dist, idx, caption))
        print(f"  {m['permalink']}  →  {best.name}  [{tag}]")

        plan.append(f"## {best.name}  ·  {m['timestamp'][:10]}\n")
        plan.append(f"- Post: {m['permalink']}")
        plan.append(f"- Match: `{best.name}` (dist {dist})")
        cur = (m.get('caption') or '').strip().replace('\n', ' ')
        plan.append(f"- Descripción actual: {'(vacía)' if not cur else cur[:80] + '…'}\n")
        plan.append("**Pegar esta descripción:**\n\n```\n" + caption + "```\n")

    (HERE / "update_plan.md").write_text("\n".join(plan), encoding="utf-8")
    print(f"\nWrote update_plan.md ({len(rows)} post(s) matched).")

    if do_comment:
        print("\n--comment: posting CTA as first comment (needs manage_comments)…")
        done = 0
        for m, best, dist, idx, caption in rows:
            if existing_comment_has_link(m["id"]):
                print(f"   · {best.name}: ya tiene comentario con el enlace — salteado")
                continue
            if post_comment(m["id"], COMMENT_TEXT):
                print(f"   ✓ {best.name}: comentario publicado"); done += 1
        print(f"Comentarios publicados: {done}")
    else:
        print("Tip: run with --comment to also post the WhatsApp+link as a first comment.")


if __name__ == "__main__":
    main(do_comment="--comment" in sys.argv[1:])
