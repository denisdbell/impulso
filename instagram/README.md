# Impulso Crédito — Instagram auto-poster

`post_story.py` builds a branded **1080×1920** image, writes the day's caption with
Claude, hosts it on ImgBB, and publishes it as an Instagram **Story** through the
Graph API — one story per run. Point a daily scheduler at it and the account stays
active on its own.

> **Note — it posts a Story.** `post_to_instagram()` sends `media_type=STORIES`,
> which matches the 9:16 (1080×1920) image. A **feed** post would need a 4:5…1.91:1
> image and `media_type` omitted — the 9:16 render is rejected by the feed.

---

## 1. Prerequisites

- **Python 3.9+** (`python3 --version`)
- A font the renderer can find. macOS ships Arial at the paths the script uses; on
  Linux install DejaVu/Liberation (`sudo apt install fonts-dejavu` or
  `fonts-liberation`) — the script falls back to them automatically.
- Four credentials (see step 3).

## 2. Install

```bash
cd instagram
python3 -m venv .venv            # optional but recommended
source .venv/bin/activate
pip install -r requirements.txt
```

## 3. Configure credentials

Copy the template and fill in the four keys:

```bash
cp .env.example .env
```

```
ANTHROPIC_API_KEY=   # console.anthropic.com → API keys
IG_ACCESS_TOKEN=     # long-lived Instagram token (see below)
IG_USER_ID=          # Instagram Business account numeric id (see below)
IMGBB_API_KEY=       # free key from api.imgbb.com (click "Get API key")
```

**`IMGBB_API_KEY`** — the image is uploaded to [ImgBB](https://imgbb.com) so
Instagram has a public URL to fetch. Log in at imgbb.com → **https://api.imgbb.com**
→ **Get API key** → copy it into `.env`. Free, no credit card.

**`IG_ACCESS_TOKEN` + `IG_USER_ID`** — follow the two `curl` calls in the
[`commands`](./commands) file:

1. Get a short-lived token from the Graph API Explorer, then exchange it for a
   **long-lived** one (first `curl`). Paste the result into `IG_ACCESS_TOKEN`.
2. Call `me/accounts` with that token (second `curl`) and copy the numeric
   `instagram_business_account.id` into `IG_USER_ID`.

The Instagram account must be a **Business/Creator** account linked to a Facebook
Page, and the app needs the `instagram_basic` + `instagram_content_publish`
permissions. Long-lived tokens last ~60 days — refresh before they expire.

> `.env` and `commands` hold real secrets and are already in `.gitignore`. Never
> commit them. If a token or client secret leaks, rotate it.

---

## 4. Test it with a single story

Do this **before** scheduling anything — it produces exactly one story and never
touches your Instagram account.

**a) Safe dry run** — generates today's caption with Claude and renders the image,
but **skips ImgBB and Instagram** entirely. The result is saved as
`story_preview.png` so you can open and inspect it:

```bash
python3 post_story.py --dry-run
open story_preview.png            # macOS  (Linux: xdg-open)
```

**b) Fully offline render** — also skip the Claude call by passing your own text.
No API keys needed; pure layout check:

```bash
python3 post_story.py --dry-run --message "Préstamo en 24 h por WhatsApp 💪"
```

**c) One real post** — when the preview looks right, run it once with no flags to
generate, host, and publish a single post for real:

```bash
python3 post_story.py
```

Watch the log: it prints the generated `Message`, the hosted `i.ibb.co` URL, the
container id, and the published media id. On success the temp image is deleted; on
failure it raises with the API error so you can see what went wrong.

---

## 5. Schedule it (daily)

Once the single-post test works, automate it with `cron`. Use absolute paths and
the venv's Python so `cron`'s bare environment finds everything:

```cron
# Every day at 10:00 — one post
0 10 * * * cd /Users/denisbell/Documents/repos/impulso/instagram && /Users/denisbell/Documents/repos/impulso/instagram/.venv/bin/python post_story.py >> cron.log 2>&1
```

(`crontab -e` to edit. On macOS, give `cron`/Terminal Full Disk Access if it can't
read the files.)

---

## 6. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `KeyError: 'ANTHROPIC_API_KEY'` (or any key) | A `.env` value is missing or `.env` isn't in this folder. |
| Boxes/wrong glyphs in the image | No usable font found — install DejaVu/Liberation (see Prerequisites). |
| `400` from `graph.facebook.com` | Token expired, wrong `IG_USER_ID`, or the account isn't a linked Business account. |
| ImgBB `400`/`invalid API key` | `IMGBB_API_KEY` missing or wrong in `.env`; grab a fresh one at api.imgbb.com. |
| Date reads e.g. "Tuesday 27 de October" | Locale isn't set to `es_AR`; cosmetic only. |

Flags: `--dry-run` (render only, keep `story_preview.png`), `--message "<text>"`
(use fixed text instead of calling Claude).
