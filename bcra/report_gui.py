#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Impulso Crédito — Credit Dashboard (desktop GUI).

A friendly graphical front-end over the same analyzers the CLI uses
(analyze_bcra.py, analyze_impulso.py, analyze_combined.py), so every number on
screen matches `report.py` exactly. Built with Tkinter/ttk — Python's standard,
zero-install desktop toolkit — so it runs anywhere Python does, no pip needed.

    python3 bcra/report_gui.py

Features
  • Overview with summary cards + a decision bar chart + best/worst lists.
  • Combined / BCRA / Impulso ranking tabs: searchable, sortable, color-coded by
    risk; double-click a row to open that person.
  • Person view: combined / BCRA / Impulso score gauges, a decision badge, the
    Impulso loan strip, the BCRA 24-month situación timeline, and the detail tables.
  • Reload (and launch) re-downloads a fresh workbook from the live Google Sheet
    (fetch_impulso path), falling back to the on-disk copy if it can't; missing
    BCRA files download automatically in the background (fetch_bcra path). Both
    toggle off under Options. Export any ranking or a person dossier to Markdown.

Everything runs off the Impulso workbook at bcra/out/_impulso.xlsx (download it
with fetch_impulso.py, or File ▸ Open workbook…).
"""

import datetime
import queue
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, font as tkfont, filedialog, messagebox

import analyze_bcra
import analyze_impulso
import analyze_combined
import fetch_impulso  # re-download the workbook from the live Google Sheet on reload
import report  # reuse the polite auto-fetch (_ensure_bcra) + weights

HERE = Path(__file__).resolve().parent
DEFAULT_XLSX = HERE / "out" / "_impulso.xlsx"
DEFAULT_OUT = HERE / "out"

# --------------------------------------------------------------------------- #
# Palette
# --------------------------------------------------------------------------- #
BG = "#eef1f5"          # app background
SURFACE = "#ffffff"     # cards / tables
SIDEBAR = "#111827"     # dark rail
SIDEBAR_HOVER = "#1f2937"
SIDEBAR_ACTIVE = "#2563eb"
INK = "#111827"         # primary text
MUTED = "#6b7280"       # secondary text
LINE = "#e5e7eb"        # hairlines
ACCENT = "#2563eb"

RATING_COLOR = {"A": "#15803d", "B": "#65a30d", "C": "#d97706", "D": "#ea580c", "E": "#dc2626", "—": "#9ca3af"}
DECISION_COLOR = {"OK": "#15803d", "REVIEW": "#d97706", "REJECT": "#dc2626", "NO RECORD": "#9ca3af"}
DECISION_ROW_BG = {"OK": "#e9f7ef", "REVIEW": "#fdf3e3", "REJECT": "#fcebea", "NO RECORD": "#f3f4f6"}
# Impulso per-loan status → colour
LOAN_COLOR = {
    "repaid_ontime": "#15803d", "repaid_untimed": "#22c55e", "current": "#2563eb",
    "repaid_late": "#d97706", "overdue": "#dc2626", "in_plan": "#7c3aed", "frozen": "#0891b2",
}
# BCRA situación level (0 none … 5/6 loss) → colour
SIT_COLOR = {0: "#d1fae5", 1: "#16a34a", 2: "#a3e635", 3: "#f59e0b", 4: "#f97316", 5: "#dc2626", 6: "#7f1d1d"}


def rating_letter(rating):
    return rating[0] if rating and rating[0] in "ABCDE" else "—"


def money(x):
    return analyze_impulso.money(x)


# --------------------------------------------------------------------------- #
# Data layer (no Tk — unit-testable on its own)
# --------------------------------------------------------------------------- #
class DataService:
    """Loads and scores the portfolio, and assembles a single person's dossier,
    reusing the analyzer modules so results match the CLI."""

    def __init__(self, xlsx, out_dir, iw=analyze_combined.IMPULSO_WEIGHT):
        self.xlsx = Path(xlsx)
        self.out_dir = Path(out_dir)
        self.iw = iw
        self.bw = 1.0 - iw
        # snapshot
        self.combined = []
        self.icohort = []
        self.bcohort = []
        self.clients = {}
        self.loaded_at = None

    # ---- whole portfolio ----
    def refresh_workbook(self, timeout=60):
        """Re-download out/_impulso.xlsx from the live Google Sheet via the
        service account (fetch_impulso). Returns (ok, message); never raises —
        on any problem we keep whatever copy is already on disk."""
        if self.xlsx.resolve() != DEFAULT_XLSX.resolve():
            return False, "custom workbook open — not overwriting"
        try:
            file_id = fetch_impulso.resolve_file_id(None, HERE)
            cred = fetch_impulso.resolve_credentials(None, HERE)
            n = fetch_impulso.fetch(file_id, cred, self.xlsx, timeout)
            return True, "Downloaded latest workbook (%.0f KB)" % (n / 1024.0)
        except (SystemExit, Exception) as exc:  # SystemExit: fetch_impulso's own failures
            return False, (str(exc).splitlines() or ["download failed"])[0]

    def load_all(self, fetch=True, insecure=False, refresh=False, progress=lambda m: None):
        if refresh:
            progress("Downloading latest workbook…")
            ok, msg = self.refresh_workbook()
            progress(msg if ok else "Using on-disk workbook (%s)" % msg)
        if not self.xlsx.exists():
            raise FileNotFoundError(str(self.xlsx))
        today = datetime.date.today()
        progress("Reading workbook…")
        tables = analyze_impulso.load_tables(str(self.xlsx))
        self.clients = analyze_impulso.build_clients(tables, today)
        self.icohort = analyze_impulso.build_cohort(str(self.xlsx), today)
        if fetch:
            cuils = [c["cuil"] for c in self.icohort if c.get("cuil")]
            progress("Downloading missing BCRA files…")
            report._ensure_bcra(cuils, str(self.out_dir), insecure)
        progress("Scoring BCRA…")
        self.bcohort = analyze_bcra.build_cohort(str(self.out_dir)) if self.out_dir.exists() else []
        progress("Combining…")
        self.combined = analyze_combined.build_combined(str(self.out_dir), str(self.xlsx), today, self.iw, self.bw)
        self.loaded_at = datetime.datetime.now()

    # ---- table-ready rows ----
    def combined_rows(self):
        rows = sorted(self.combined, key=lambda r: (-(r["combined"] if r["combined"] is not None else -1), r["name"]))
        out = []
        for i, r in enumerate(rows, 1):
            out.append({
                "rank": i, "cuil": r["cuil"] or "—", "name": r["name"] or "—",
                "combined": r["combined"] if r["combined"] is not None else "",
                "rating": rating_letter(r["rating"]),
                "bcra": r["bcra_score"] if r["bcra_score"] is not None else "—",
                "impulso": r["imp_score"] if r["imp_score"] is not None else "—",
                "decision": r["decision"], "source": r["source"],
                "token": r["cuil"] or r["name"],
            })
        return out

    def impulso_rows(self):
        rows = sorted(self.icohort, key=lambda c: (-c["score"], c["name"]))
        return [{
            "rank": i, "cuil": c["cuil"] or "—", "name": c["name"] or "—",
            "score": c["score"], "rating": rating_letter(c["rating"]),
            "category": c["category"], "decision": c["decision"], "loans": c["n_loans"],
            "token": c["cuil"] or c["dni"] or c["key"],
        } for i, c in enumerate(rows, 1)]

    def bcra_rows(self):
        rows = sorted(self.bcohort, key=lambda c: (-c["score"], c["cuil"]))
        return [{
            "rank": i, "cuil": c["cuil"] or "—", "name": c["name"] or "—",
            "score": c["score"], "rating": rating_letter(c["rating"]),
            "category": c["category"], "decision": c["decision"],
            "token": c["cuil"],
        } for i, c in enumerate(rows, 1)]

    def summary(self):
        rows = self.combined
        vals = [r["combined"] for r in rows if r["combined"] is not None]
        dec = {"OK": 0, "REVIEW": 0, "REJECT": 0, "NO RECORD": 0}
        for r in rows:
            dec[r["decision"]] = dec.get(r["decision"], 0) + 1
        both = sum(1 for r in rows if r["source"] == "both")
        return {
            "people": len(rows), "avg": (sum(vals) / len(vals)) if vals else 0,
            "decisions": dec, "both": both, "single": len(rows) - both,
            "borrowers": len(self.icohort), "bureau": len(self.bcohort),
            "loaded_at": self.loaded_at,
        }

    def people_choices(self):
        """(label, token) for the person search box."""
        out = []
        for c in self.clients.values():
            if not c["loans"]:
                continue
            tag = c["cuil"] or c["dni"] or c["client_id"]
            if not tag:
                continue
            out.append(("%s — %s" % (c["name"] or "(no name)", tag), tag))
        out.sort(key=lambda t: t[0].lower())
        return out

    # ---- one person ----
    def assemble_person(self, token, fetch=True, insecure=False):
        today = datetime.date.today()
        if not self.clients:
            tables = analyze_impulso.load_tables(str(self.xlsx))
            self.clients = analyze_impulso.build_clients(tables, today)
        cl = analyze_impulso.find_client(self.clients, token)
        import re
        digits = re.sub(r"\D", "", token)
        cuil = (cl["cuil"] if cl and cl["cuil"] else "") or (digits if len(digits) == 11 else "")

        if cuil and fetch:
            report._ensure_bcra([cuil], str(self.out_dir), insecure)
            self.bcohort = analyze_bcra.build_cohort(str(self.out_dir)) if self.out_dir.exists() else self.bcohort

        a_bcra = bcra_rank = None
        if cuil:
            dest = self.out_dir / ("%s.json" % cuil)
            if dest.exists():
                import json
                try:
                    a_bcra = analyze_bcra.analyze_cuil(json.loads(dest.read_text(encoding="utf-8")))
                    bcra_rank = analyze_bcra.rank_of(self.bcohort, a_bcra["cuil"] or cuil, a_bcra["score"], a_bcra["category"])
                except Exception:
                    a_bcra = None

        a_imp = imp_rank = None
        if cl:
            a_imp = analyze_impulso.analyze_client(cl, today)
            if a_imp["has_record"]:
                key = a_imp["cuil"] or ("DNI:" + a_imp["dni"] if a_imp["dni"] else "CID:" + a_imp["client_id"])
                imp_rank = analyze_impulso.rank_of(self.icohort, key, a_imp["score"], a_imp["category"])

        bcra_light = None
        if a_bcra:
            bcra_light = {"cuil": a_bcra["cuil"] or cuil, "name": a_bcra["name"], "score": a_bcra["score"],
                          "rating": a_bcra["rating"], "category": a_bcra["category"], "decision": a_bcra["decision"]}
        imp_light = None
        if a_imp and a_imp["has_record"]:
            imp_light = {"cuil": a_imp["cuil"], "name": a_imp["name"], "score": a_imp["score"],
                         "rating": a_imp["rating"], "category": a_imp["category"], "decision": a_imp["decision"],
                         "n_loans": a_imp["n_loans"]}
        rec = analyze_combined.combine(bcra_light, imp_light, self.iw, self.bw) if (bcra_light or imp_light) else None

        name = (a_imp["name"] if a_imp and a_imp.get("name") else "") or (a_bcra["name"] if a_bcra else "") or token
        return {"token": token, "name": name, "cuil": cuil,
                "a_bcra": a_bcra, "bcra_rank": bcra_rank,
                "a_imp": a_imp, "imp_rank": imp_rank, "rec": rec}


# --------------------------------------------------------------------------- #
# Small UI helpers
# --------------------------------------------------------------------------- #
def draw_gauge(cv, cx, cy, radius, value, title, color, fonts):
    """Semicircle 0-100 gauge."""
    cv.create_arc(cx - radius, cy - radius, cx + radius, cy + radius,
                  start=0, extent=180, style="arc", outline=LINE, width=14)
    if isinstance(value, (int, float)):
        ext = max(0.0, min(100.0, value)) / 100.0 * 180.0
        if ext > 0:
            cv.create_arc(cx - radius, cy - radius, cx + radius, cy + radius,
                          start=180, extent=-ext, style="arc", outline=color, width=14)
        cv.create_text(cx, cy - 4, text=str(value), font=fonts["gauge"], fill=color)
    else:
        cv.create_text(cx, cy - 4, text="—", font=fonts["gauge"], fill=MUTED)
    cv.create_text(cx, cy + 20, text=title, font=fonts["small"], fill=MUTED)


def draw_badge(cv, x, y, text, color, fonts, pad=10):
    tid = cv.create_text(x + pad, y + 5, text=text, font=fonts["badge"], fill="#ffffff", anchor="nw")
    bb = cv.bbox(tid)
    rid = cv.create_rectangle(bb[0] - pad, bb[1] - 4, bb[2] + pad, bb[3] + 4, fill=color, outline=color)
    cv.tag_lower(rid, tid)
    return bb[2] + pad


class ScrollFrame(ttk.Frame):
    """A vertically scrollable container; put content in `.body`."""

    def __init__(self, parent, bg=BG):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.body = tk.Frame(self.canvas, bg=bg)
        self._win = self.canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.body.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        for w in (self.canvas, self.body):
            w.bind("<Enter>", lambda e: self._bind_wheel())
            w.bind("<Leave>", lambda e: self._unbind_wheel())

    def _bind_wheel(self):
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)
        self.canvas.bind_all("<Button-4>", self._on_wheel)
        self.canvas.bind_all("<Button-5>", self._on_wheel)

    def _unbind_wheel(self):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_wheel(self, e):
        if getattr(e, "num", None) == 4:
            self.canvas.yview_scroll(-1, "units")
        elif getattr(e, "num", None) == 5:
            self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")


class RankingPanel(ttk.Frame):
    """Search box + sortable, colour-coded Treeview. `columns` = list of
    (key, heading, width, anchor, numeric)."""

    def __init__(self, parent, columns, on_open, fonts):
        super().__init__(parent, style="Card.TFrame")
        self.columns = columns
        self.on_open = on_open
        self._rows = []
        self._sort = (None, False)

        bar = ttk.Frame(self, style="Card.TFrame")
        bar.pack(fill="x", padx=12, pady=(12, 6))
        ttk.Label(bar, text="🔍", style="Card.TLabel").pack(side="left")
        self.q = tk.StringVar()
        e = ttk.Entry(bar, textvariable=self.q)
        e.pack(side="left", fill="x", expand=True, padx=(6, 0))
        e.insert(0, "")
        self.q.trace_add("write", lambda *a: self._refill())
        ttk.Label(bar, text="double-click a row to open the person", style="Hint.TLabel").pack(side="right")

        keys = [c[0] for c in columns]
        self.tree = ttk.Treeview(self, columns=keys, show="headings", selectmode="browse", height=18)
        for key, head, width, anchor, _num in columns:
            self.tree.heading(key, text=head, command=lambda k=key: self._sort_by(k))
            self.tree.column(key, width=width, anchor=anchor, stretch=(key == "name"))
        for dec, bg in DECISION_ROW_BG.items():
            self.tree.tag_configure("dec_" + dec.replace(" ", "_"), background=bg)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(12, 0), pady=(0, 12))
        vsb.pack(side="left", fill="y", padx=(0, 12), pady=(0, 12))
        self.tree.bind("<Double-1>", self._open)
        self.tree.bind("<Return>", self._open)

    def set_rows(self, rows):
        self._rows = rows
        self._refill()

    def _sort_by(self, key):
        asc = not (self._sort[0] == key and self._sort[1])
        self._sort = (key, asc)
        self._refill()

    def _refill(self):
        self.tree.delete(*self.tree.get_children())
        q = self.q.get().strip().lower()
        rows = [r for r in self._rows
                if not q or q in str(r.get("name", "")).lower() or q in str(r.get("cuil", "")).lower()]
        key, asc = self._sort
        if key:
            num = dict((c[0], c[4]) for c in self.columns).get(key, False)
            def sk(r):
                v = r.get(key, "")
                if num:
                    try:
                        return (0, float(v))
                    except (TypeError, ValueError):
                        return (1, 0.0)
                return (0, str(v).lower())
            rows = sorted(rows, key=sk, reverse=not asc)
        for r in rows:
            vals = [r.get(c[0], "") for c in self.columns]
            tag = "dec_" + str(r.get("decision", "")).replace(" ", "_")
            self.tree.insert("", "end", values=vals, tags=(tag, r.get("token", "")))

    def _open(self, _e):
        sel = self.tree.selection()
        if not sel:
            return
        tags = self.tree.item(sel[0], "tags")
        token = tags[1] if len(tags) > 1 else ""
        if token:
            self.on_open(token)


# --------------------------------------------------------------------------- #
# Application
# --------------------------------------------------------------------------- #
class App(tk.Tk):
    def __init__(self, xlsx, out_dir):
        super().__init__()
        self.title("Impulso Crédito — Credit Dashboard")
        self.geometry("1180x760")
        self.minsize(960, 620)
        self.configure(bg=BG)

        self.data = DataService(xlsx, out_dir)
        self.fetch_var = tk.BooleanVar(value=True)
        self.insecure_var = tk.BooleanVar(value=False)
        self.refresh_var = tk.BooleanVar(value=True)  # re-download workbook on reload
        self.q = queue.Queue()

        self._init_fonts()
        self._init_style()
        self._build_menu()
        self._build_layout()
        self.after(100, self._poll)
        self.reload()

    # ---- fonts / styles ----
    def _init_fonts(self):
        base = "Helvetica"
        for fam in ("SF Pro Text", "Segoe UI", "Helvetica Neue", "Helvetica", "Arial"):
            if fam in tkfont.families():
                base = fam
                break
        self.fonts = {
            "base": tkfont.Font(family=base, size=12),
            "small": tkfont.Font(family=base, size=10),
            "hint": tkfont.Font(family=base, size=9, slant="italic"),
            "bold": tkfont.Font(family=base, size=12, weight="bold"),
            "h1": tkfont.Font(family=base, size=22, weight="bold"),
            "h2": tkfont.Font(family=base, size=15, weight="bold"),
            "card_num": tkfont.Font(family=base, size=26, weight="bold"),
            "gauge": tkfont.Font(family=base, size=22, weight="bold"),
            "badge": tkfont.Font(family=base, size=10, weight="bold"),
            "nav": tkfont.Font(family=base, size=13),
        }

    def _init_style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure(".", font=self.fonts["base"], background=BG, foreground=INK)
        st.configure("TFrame", background=BG)
        st.configure("Card.TFrame", background=SURFACE)
        st.configure("Card.TLabel", background=SURFACE, foreground=INK)
        st.configure("Hint.TLabel", background=SURFACE, foreground=MUTED, font=self.fonts["hint"])
        st.configure("Muted.TLabel", background=BG, foreground=MUTED)
        st.configure("H1.TLabel", background=BG, foreground=INK, font=self.fonts["h1"])
        st.configure("H2.TLabel", background=SURFACE, foreground=INK, font=self.fonts["h2"])
        st.configure("TButton", font=self.fonts["base"], padding=6)
        st.configure("Accent.TButton", foreground="#ffffff", background=ACCENT, padding=8)
        st.map("Accent.TButton", background=[("active", "#1d4ed8")])
        st.configure("Treeview", rowheight=28, fieldbackground=SURFACE, background=SURFACE,
                     foreground=INK, font=self.fonts["base"], borderwidth=0)
        st.configure("Treeview.Heading", font=self.fonts["bold"], padding=6)
        st.map("Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", INK)])
        st.configure("Status.TLabel", background="#0b1220", foreground="#cbd5e1", font=self.fonts["small"])

    # ---- menu ----
    def _build_menu(self):
        m = tk.Menu(self)
        filem = tk.Menu(m, tearoff=0)
        filem.add_command(label="Open workbook…", command=self._open_workbook)
        filem.add_command(label="Reload", accelerator="Ctrl+R", command=self.reload)
        filem.add_separator()
        filem.add_command(label="Export current ranking…", command=self._export_ranking)
        filem.add_separator()
        filem.add_command(label="Quit", command=self.destroy)
        m.add_cascade(label="File", menu=filem)
        opt = tk.Menu(m, tearoff=0)
        opt.add_checkbutton(label="Download latest workbook on reload", variable=self.refresh_var)
        opt.add_checkbutton(label="Download missing BCRA files", variable=self.fetch_var)
        opt.add_checkbutton(label="Insecure TLS (BCRA cert fallback)", variable=self.insecure_var)
        m.add_cascade(label="Options", menu=opt)
        helpm = tk.Menu(m, tearoff=0)
        helpm.add_command(label="About", command=self._about)
        m.add_cascade(label="Help", menu=helpm)
        self.config(menu=m)
        self.bind_all("<Control-r>", lambda e: self.reload())

    # ---- layout ----
    def _build_layout(self):
        # sidebar
        self.side = tk.Frame(self, bg=SIDEBAR, width=210)
        self.side.pack(side="left", fill="y")
        self.side.pack_propagate(False)
        tk.Label(self.side, text="  Impulso", bg=SIDEBAR, fg="#ffffff",
                 font=self.fonts["h2"], anchor="w").pack(fill="x", pady=(18, 2), padx=10)
        tk.Label(self.side, text="  Credit Dashboard", bg=SIDEBAR, fg="#93a4bd",
                 font=self.fonts["small"], anchor="w").pack(fill="x", padx=10)
        tk.Frame(self.side, bg="#1f2937", height=1).pack(fill="x", pady=14, padx=10)

        self.nav_buttons = {}
        self.pages = {}
        for key, label in [("overview", "📊  Overview"), ("combined", "🔀  Combined"),
                           ("bcra", "🏦  BCRA bureau"), ("impulso", "💳  Impulso"),
                           ("person", "🔍  Person")]:
            b = tk.Label(self.side, text="   " + label, bg=SIDEBAR, fg="#e5e7eb",
                         font=self.fonts["nav"], anchor="w", padx=8, pady=10, cursor="hand2")
            b.pack(fill="x", padx=8, pady=1)
            b.bind("<Button-1>", lambda e, k=key: self.show_page(k))
            b.bind("<Enter>", lambda e, bb=b, k=key: bb.configure(bg=SIDEBAR_HOVER) if self._active != k else None)
            b.bind("<Leave>", lambda e, bb=b, k=key: bb.configure(bg=SIDEBAR if self._active != k else SIDEBAR_ACTIVE))
            self.nav_buttons[key] = b

        reload_btn = tk.Label(self.side, text="   ⟳  Reload data", bg=SIDEBAR, fg="#93c5fd",
                              font=self.fonts["nav"], anchor="w", padx=8, pady=10, cursor="hand2")
        reload_btn.pack(side="bottom", fill="x", padx=8, pady=10)
        reload_btn.bind("<Button-1>", lambda e: self.reload())

        # main area
        main = ttk.Frame(self, style="TFrame")
        main.pack(side="left", fill="both", expand=True)
        self.content = ttk.Frame(main, style="TFrame")
        self.content.pack(fill="both", expand=True, padx=16, pady=16)
        self.status = ttk.Label(main, text="", style="Status.TLabel", anchor="w", padding=(12, 5))
        self.status.pack(fill="x", side="bottom")

        self._active = None
        self._build_overview()
        self._build_rankings()
        self._build_person()
        self.show_page("overview")

    def _place(self, frame):
        for f in self.pages.values():
            f.pack_forget()
        frame.pack(fill="both", expand=True)

    def show_page(self, key):
        self._active = key
        for k, b in self.nav_buttons.items():
            b.configure(bg=SIDEBAR_ACTIVE if k == key else SIDEBAR,
                        fg="#ffffff" if k == key else "#e5e7eb")
        self._place(self.pages[key])

    # ---- pages ----
    def _build_overview(self):
        page = ttk.Frame(self.content, style="TFrame")
        self.pages["overview"] = page
        ttk.Label(page, text="Portfolio overview", style="H1.TLabel").pack(anchor="w")
        self.ov_sub = ttk.Label(page, text="", style="Muted.TLabel")
        self.ov_sub.pack(anchor="w", pady=(2, 12))
        self.cards = ttk.Frame(page, style="TFrame")
        self.cards.pack(fill="x")
        lower = ttk.Frame(page, style="TFrame")
        lower.pack(fill="both", expand=True, pady=(14, 0))
        self.chart_card = ttk.Frame(lower, style="Card.TFrame")
        self.chart_card.pack(side="left", fill="both", expand=True)
        ttk.Label(self.chart_card, text="Decisions (combined)", style="H2.TLabel").pack(anchor="w", padx=14, pady=(12, 0))
        self.chart = tk.Canvas(self.chart_card, bg=SURFACE, highlightthickness=0, height=260)
        self.chart.pack(fill="both", expand=True, padx=14, pady=12)
        self.chart.bind("<Configure>", lambda e: self._draw_overview_chart())
        self.lists_card = ttk.Frame(lower, style="Card.TFrame", width=360)
        self.lists_card.pack(side="left", fill="y", padx=(14, 0))
        self.lists_card.pack_propagate(False)
        self.lists_inner = ttk.Frame(self.lists_card, style="Card.TFrame")
        self.lists_inner.pack(fill="both", expand=True, padx=14, pady=12)

    def _card(self, parent, title, value, color=INK):
        c = ttk.Frame(parent, style="Card.TFrame")
        c.pack(side="left", fill="both", expand=True, padx=(0, 12))
        ttk.Label(c, text=title, style="Card.TLabel", foreground=MUTED,
                  font=self.fonts["small"]).pack(anchor="w", padx=14, pady=(12, 0))
        lbl = tk.Label(c, text=value, bg=SURFACE, fg=color, font=self.fonts["card_num"], anchor="w")
        lbl.pack(anchor="w", padx=14, pady=(0, 12))
        return c

    def _refresh_overview(self):
        s = self.data.summary()
        for w in self.cards.winfo_children():
            w.destroy()
        d = s["decisions"]
        self._card(self.cards, "People", str(s["people"]))
        self._card(self.cards, "Avg combined", "%.0f" % s["avg"], ACCENT)
        self._card(self.cards, "OK", str(d.get("OK", 0)), DECISION_COLOR["OK"])
        self._card(self.cards, "Review", str(d.get("REVIEW", 0)), DECISION_COLOR["REVIEW"])
        self._card(self.cards, "Reject", str(d.get("REJECT", 0)), DECISION_COLOR["REJECT"])
        when = s["loaded_at"].strftime("%Y-%m-%d %H:%M") if s["loaded_at"] else "—"
        self.ov_sub.configure(text="%d borrowers · %d bureau records · %d joined on both · %d single-source · loaded %s"
                              % (s["borrowers"], s["bureau"], s["both"], s["single"], when))
        self._draw_overview_chart()
        self._refresh_lists()

    def _draw_overview_chart(self):
        cv = self.chart
        cv.delete("all")
        if not self.data.combined:
            return
        d = self.data.summary()["decisions"]
        order = ["OK", "REVIEW", "REJECT", "NO RECORD"]
        data = [(k, d.get(k, 0)) for k in order]
        w = cv.winfo_width() or 600
        h = cv.winfo_height() or 240
        mx = max((v for _, v in data), default=1) or 1
        pad = 40
        bw = (w - 2 * pad) / len(data)
        base = h - 34
        for i, (k, v) in enumerate(data):
            x = pad + i * bw
            bh = (base - 20) * (v / mx)
            cv.create_rectangle(x + 16, base - bh, x + bw - 16, base,
                                fill=DECISION_COLOR.get(k, MUTED), outline="")
            cv.create_text(x + bw / 2, base - bh - 12, text=str(v), font=self.fonts["bold"], fill=INK)
            cv.create_text(x + bw / 2, base + 16, text=k, font=self.fonts["small"], fill=MUTED)

    def _refresh_lists(self):
        for w in self.lists_inner.winfo_children():
            w.destroy()
        rows = self.data.combined_rows()
        def mini(title, subset):
            ttk.Label(self.lists_inner, text=title, style="H2.TLabel").pack(anchor="w", pady=(4, 2))
            for r in subset:
                line = ttk.Frame(self.lists_inner, style="Card.TFrame")
                line.pack(fill="x", pady=1)
                dot = tk.Canvas(line, width=12, height=12, bg=SURFACE, highlightthickness=0)
                dot.create_oval(2, 2, 11, 11, fill=DECISION_COLOR.get(r["decision"], MUTED), outline="")
                dot.pack(side="left", padx=(0, 8))
                ttk.Label(line, text=("%s" % r["name"])[:26], style="Card.TLabel").pack(side="left")
                ttk.Label(line, text=str(r["combined"]), style="Card.TLabel",
                          foreground=RATING_COLOR.get(r["rating"], INK), font=self.fonts["bold"]).pack(side="right")
        mini("Best payers", rows[:6])
        ttk.Frame(self.lists_inner, style="Card.TFrame", height=10).pack()
        mini("Needs attention", rows[-6:][::-1])

    def _build_rankings(self):
        self.pages["combined"] = self._ranking_page(
            "Combined ranking", "BCRA + Impulso, blended (Impulso %.0f%% · BCRA %.0f%%)"
            % (self.data.iw * 100, self.data.bw * 100),
            [("rank", "#", 44, "center", True), ("cuil", "CUIL", 120, "center", False),
             ("name", "Name", 230, "w", False), ("combined", "Combined", 90, "center", True),
             ("rating", "Rt", 44, "center", False), ("bcra", "BCRA", 70, "center", True),
             ("impulso", "Impulso", 80, "center", True), ("decision", "Decision", 90, "center", False),
             ("source", "Source", 90, "center", False)],
            lambda: self.data.combined_rows(), "combined")
        self.pages["bcra"] = self._ranking_page(
            "BCRA bureau ranking", "External Central de Deudores standing",
            [("rank", "#", 44, "center", True), ("cuil", "CUIL", 120, "center", False),
             ("name", "Name", 240, "w", False), ("score", "Score", 80, "center", True),
             ("rating", "Rt", 44, "center", False), ("category", "Payment profile", 220, "w", False),
             ("decision", "Decision", 90, "center", False)],
            lambda: self.data.bcra_rows(), "bcra")
        self.pages["impulso"] = self._ranking_page(
            "Impulso ranking", "First-party repayment behaviour with us",
            [("rank", "#", 44, "center", True), ("cuil", "CUIL", 120, "center", False),
             ("name", "Name", 230, "w", False), ("score", "Score", 80, "center", True),
             ("rating", "Rt", 44, "center", False), ("category", "Payment profile", 210, "w", False),
             ("loans", "Loans", 60, "center", True), ("decision", "Decision", 90, "center", False)],
            lambda: self.data.impulso_rows(), "impulso")

    def _ranking_page(self, title, subtitle, columns, rows_fn, key):
        page = ttk.Frame(self.content, style="TFrame")
        head = ttk.Frame(page, style="TFrame")
        head.pack(fill="x")
        ttk.Label(head, text=title, style="H1.TLabel").pack(side="left")
        ttk.Button(head, text="Export .md", command=lambda k=key: self._export_named(k)).pack(side="right")
        ttk.Label(page, text=subtitle, style="Muted.TLabel").pack(anchor="w", pady=(2, 12))
        panel = RankingPanel(page, columns, self.open_person, self.fonts)
        panel.pack(fill="both", expand=True)
        panel._rows_fn = rows_fn
        page._panel = panel
        return page

    def _build_person(self):
        page = ttk.Frame(self.content, style="TFrame")
        self.pages["person"] = page
        bar = ttk.Frame(page, style="TFrame")
        bar.pack(fill="x")
        ttk.Label(bar, text="Person", style="H1.TLabel").pack(side="left")
        ttk.Label(bar, text="  search:", style="Muted.TLabel").pack(side="left", padx=(12, 4))
        self.person_cb = ttk.Combobox(bar, width=42, state="normal")
        self.person_cb.pack(side="left")
        self.person_cb.bind("<Return>", lambda e: self._search_person())
        self.person_cb.bind("<<ComboboxSelected>>", lambda e: self._search_person())
        ttk.Button(bar, text="Open", style="Accent.TButton", command=self._search_person).pack(side="left", padx=6)
        self.person_export = ttk.Button(bar, text="Export dossier .md", command=self._export_person, state="disabled")
        self.person_export.pack(side="right")
        self.person_scroll = ScrollFrame(page)
        self.person_scroll.pack(fill="both", expand=True, pady=(12, 0))
        self.person_body = self.person_scroll.body
        self._person_payload = None
        self._cb_map = {}  # label → token; filled on load (guard early searches)
        self._placeholder(self.person_body, "Search a person by name, CUIL, or DNI — or double-click anyone in a ranking.")

    def _placeholder(self, parent, text):
        for w in parent.winfo_children():
            w.destroy()
        tk.Label(parent, text=text, bg=BG, fg=MUTED, font=self.fonts["base"]).pack(pady=40)

    # ---- actions ----
    def open_person(self, token):
        self.show_page("person")
        self.person_cb.set(token)
        self._load_person(token)

    def _search_person(self):
        val = self.person_cb.get().strip()
        if not val:
            return
        token = self._cb_map.get(val, val)
        self._load_person(token)

    def _load_person(self, token):
        self._placeholder(self.person_body, "Loading…")
        self.set_status("Loading person %s…" % token)
        # Read Tk variables on the MAIN thread; Tkinter isn't thread-safe.
        fetch, insecure = self.fetch_var.get(), self.insecure_var.get()

        def work():
            try:
                payload = self.data.assemble_person(token, fetch=fetch, insecure=insecure)
                self.q.put(("person", payload))
            except Exception as exc:
                self.q.put(("error", ("person", exc, traceback.format_exc())))
        threading.Thread(target=work, daemon=True).start()

    def reload(self):
        self.set_status("Loading…")
        # Read Tk variables on the MAIN thread; Tkinter isn't thread-safe.
        fetch, insecure, refresh = self.fetch_var.get(), self.insecure_var.get(), self.refresh_var.get()

        def work():
            try:
                self.data.load_all(fetch=fetch, insecure=insecure, refresh=refresh,
                                   progress=lambda m: self.q.put(("progress", m)))
                self.q.put(("loaded", None))
            except FileNotFoundError as fnf:
                self.q.put(("error", ("load", fnf, "Workbook not found: %s" % fnf)))
            except Exception as exc:
                self.q.put(("error", ("load", exc, traceback.format_exc())))
        threading.Thread(target=work, daemon=True).start()

    # ---- queue pump ----
    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "progress":
                    self.set_status(payload)
                elif kind == "loaded":
                    self._on_loaded()
                elif kind == "person":
                    self._render_person(payload)
                elif kind == "error":
                    where, exc, tb = payload
                    self._on_error(where, exc, tb)
        except queue.Empty:
            pass
        self.after(120, self._poll)

    def _on_loaded(self):
        for key in ("combined", "bcra", "impulso"):
            panel = self.pages[key]._panel
            panel.set_rows(panel._rows_fn())
        self._refresh_overview()
        # person search choices
        choices = self.data.people_choices()
        self._cb_map = {label: token for label, token in choices}
        self.person_cb["values"] = [label for label, _ in choices]
        s = self.data.summary()
        self.set_status("Ready · %d people · avg combined %.0f · loaded %s"
                        % (s["people"], s["avg"], s["loaded_at"].strftime("%H:%M:%S")))

    def _on_error(self, where, exc, tb):
        if where == "load" and isinstance(exc, FileNotFoundError):
            self.set_status("Workbook not found — File ▸ Open workbook…")
            self._placeholder(self.person_body, "Workbook not found.")
            messagebox.showwarning(
                "Workbook not found",
                "Couldn't find the Impulso workbook:\n%s\n\nDownload it with fetch_impulso.py, "
                "or use File ▸ Open workbook… to pick an exported .xlsx." % self.data.xlsx)
            return
        self.set_status("Error: %s" % exc)
        messagebox.showerror("Error", "%s\n\n%s" % (exc, tb))

    # ---- person rendering ----
    def _render_person(self, payload):
        self._person_payload = payload
        self.person_export.configure(state="normal")
        body = self.person_body
        for w in body.winfo_children():
            w.destroy()
        f = self.fonts
        rec, a_bcra, a_imp = payload["rec"], payload["a_bcra"], payload["a_imp"]

        # header card
        head = tk.Frame(body, bg=SURFACE)
        head.pack(fill="x", pady=(0, 12))
        tk.Label(head, text=payload["name"], bg=SURFACE, fg=INK, font=f["h1"], anchor="w").pack(anchor="w", padx=16, pady=(14, 0))
        sub = "CUIL %s" % (payload["cuil"] or "—")
        if a_imp:
            sub += "    ·    DNI %s    ·    client %s" % (a_imp["dni"] or "—", a_imp["client_id"])
        tk.Label(head, text=sub, bg=SURFACE, fg=MUTED, font=f["small"], anchor="w").pack(anchor="w", padx=16, pady=(0, 14))

        # gauges row
        gcard = tk.Frame(body, bg=SURFACE)
        gcard.pack(fill="x", pady=(0, 12))
        gc = tk.Canvas(gcard, bg=SURFACE, highlightthickness=0, height=168)
        gc.pack(fill="x", padx=16, pady=14)
        gc.bind("<Configure>", lambda e: self._draw_person_gauges(gc, rec, a_bcra, a_imp))
        self.after(30, lambda: self._draw_person_gauges(gc, rec, a_bcra, a_imp))

        if rec:
            cap = tk.Frame(body, bg=SURFACE)
            cap.pack(fill="x", pady=(0, 12))
            tk.Label(cap, text="Decision", bg=SURFACE, fg=MUTED, font=f["small"]).pack(side="left", padx=(16, 8), pady=12)
            badge = tk.Canvas(cap, bg=SURFACE, highlightthickness=0, height=28, width=120)
            badge.pack(side="left", pady=12)
            draw_badge(badge, 0, 2, rec["decision"], DECISION_COLOR.get(rec["decision"], MUTED), f)
            tk.Label(cap, text="  (Impulso %.0f%% · BCRA %.0f%%)" % (self.data.iw * 100, self.data.bw * 100),
                     bg=SURFACE, fg=MUTED, font=f["small"]).pack(side="left")

        # Impulso section
        self._section_impulso(body, a_imp, payload["imp_rank"])
        # BCRA section
        self._section_bcra(body, a_bcra, payload["bcra_rank"])
        self.set_status("Showing %s%s" % (payload["name"], "  ·  CUIL " + payload["cuil"] if payload["cuil"] else ""))

    def _draw_person_gauges(self, cv, rec, a_bcra, a_imp):
        cv.delete("all")
        w = cv.winfo_width() or 760
        third = w / 3
        combined = rec["combined"] if rec else "—"
        ccolor = RATING_COLOR.get(rating_letter(rec["rating"]) if rec else "—", MUTED)
        draw_gauge(cv, third * 0.5, 96, 64, combined, "COMBINED", ccolor, self.fonts)
        bval = a_bcra["score"] if a_bcra else "—"
        bcolor = RATING_COLOR.get(rating_letter(a_bcra["rating"]) if a_bcra else "—", MUTED)
        draw_gauge(cv, third * 1.5, 96, 64, bval, "BCRA BUREAU", bcolor, self.fonts)
        ival = a_imp["score"] if (a_imp and a_imp["has_record"]) else "—"
        icolor = RATING_COLOR.get(rating_letter(a_imp["rating"]) if (a_imp and a_imp["has_record"]) else "—", MUTED)
        draw_gauge(cv, third * 2.5, 96, 64, ival, "IMPULSO", icolor, self.fonts)

    def _card_frame(self, parent, title):
        c = tk.Frame(parent, bg=SURFACE)
        c.pack(fill="x", pady=(0, 12))
        tk.Label(c, text=title, bg=SURFACE, fg=INK, font=self.fonts["h2"], anchor="w").pack(anchor="w", padx=16, pady=(12, 4))
        return c

    def _section_impulso(self, body, a_imp, rank):
        c = self._card_frame(body, "💳  Impulso — repayment with us")
        f = self.fonts
        if not a_imp or not a_imp["has_record"]:
            tk.Label(c, text="No loan history at Impulso Crédito.", bg=SURFACE, fg=MUTED,
                     font=f["base"]).pack(anchor="w", padx=16, pady=(0, 14))
            return
        prof = "%s  —  %s" % (a_imp["category"], a_imp["category_why"])
        tk.Label(c, text=prof, bg=SURFACE, fg=INK, font=f["base"], anchor="w", wraplength=760, justify="left").pack(anchor="w", padx=16)
        meta = "On-time %d%%   ·   %s   ·   borrowed %s · repaid %s · outstanding %s" % (
            a_imp["on_time_rate"], a_imp["worst_now"], money(a_imp["total_borrowed"]),
            money(a_imp["total_repaid"]), money(a_imp["outstanding"]))
        tk.Label(c, text=meta, bg=SURFACE, fg=MUTED, font=f["small"], anchor="w").pack(anchor="w", padx=16, pady=(2, 8))
        if rank:
            tk.Label(c, text="Portfolio rank %d of %d (best payer = 1)" % (rank["rank"], rank["of"]),
                     bg=SURFACE, fg=MUTED, font=f["small"]).pack(anchor="w", padx=16)
        # loan strip
        strip = tk.Canvas(c, bg=SURFACE, highlightthickness=0, height=56)
        strip.pack(fill="x", padx=16, pady=(8, 4))
        strip.bind("<Configure>", lambda e, cv=strip, loans=a_imp["loans"]: self._draw_loan_strip(cv, loans))
        self.after(30, lambda cv=strip, loans=a_imp["loans"]: self._draw_loan_strip(cv, loans))
        # loan table
        cols = [("loan", "Loan", 90, "w"), ("date", "Date", 100, "center"), ("capital", "Capital", 110, "e"),
                ("term", "Term", 60, "center"), ("status", "Status", 150, "w"), ("late", "Late", 70, "center")]
        tv = ttk.Treeview(c, columns=[x[0] for x in cols], show="headings", height=min(len(a_imp["loans"]), 8))
        for k, h, w_, anc in cols:
            tv.heading(k, text=h)
            tv.column(k, width=w_, anchor=anc)
        for l in a_imp["loans"]:
            tv.insert("", "end", values=(
                l["loan_id"], l["fecha"].isoformat() if l["fecha"] else "—", money(l["capital"]),
                "%dd" % int(l["term"]), analyze_impulso._STATUS_LABEL.get(l["status"], l["status"]),
                ("%dd" % l["days_late"]) if l["days_late"] else "—"))
        tv.pack(fill="x", padx=16, pady=(4, 14))

    def _draw_loan_strip(self, cv, loans):
        cv.delete("all")
        if not loans:
            return
        w = cv.winfo_width() or 760
        n = len(loans)
        pad = 4
        bw = min(90, (w - 2 * 16) / max(n, 1)) - pad
        x = 16
        for l in loans:
            col = LOAN_COLOR.get(l["status"], MUTED)
            cv.create_rectangle(x, 8, x + bw, 34, fill=col, outline="")
            cv.create_text(x + bw / 2, 46, text=l["loan_id"].replace("L-", ""), font=self.fonts["small"], fill=MUTED)
            x += bw + pad

    def _section_bcra(self, body, a_bcra, rank):
        c = self._card_frame(body, "🏦  BCRA — external bureau")
        f = self.fonts
        if not a_bcra:
            tk.Label(c, text="No BCRA record on disk (enable Options ▸ Download, or it has no debt history).",
                     bg=SURFACE, fg=MUTED, font=f["base"]).pack(anchor="w", padx=16, pady=(0, 14))
            return
        if not a_bcra["has_record"]:
            tk.Label(c, text="No debt found in the BCRA Central de Deudores (clean / no record).",
                     bg=SURFACE, fg=MUTED, font=f["base"]).pack(anchor="w", padx=16, pady=(0, 14))
            return
        prof = "%s  —  %s" % (a_bcra["category"], a_bcra["category_why"])
        tk.Label(c, text=prof, bg=SURFACE, fg=INK, font=f["base"], anchor="w", wraplength=760, justify="left").pack(anchor="w", padx=16)
        meta = "Current sit %d · worst24m %d · performing %.0f%% · creditors %d · total debt %s · checks %d" % (
            a_bcra["current_worst"], a_bcra["worst_24m"], a_bcra["pct_performante"],
            a_bcra["entidades"], money(a_bcra["total_debt"]), a_bcra["cheques"])
        tk.Label(c, text=meta, bg=SURFACE, fg=MUTED, font=f["small"], anchor="w").pack(anchor="w", padx=16, pady=(2, 6))
        if rank:
            tk.Label(c, text="Cohort rank %d of %d (lowest risk = 1)" % (rank["rank"], rank["of"]),
                     bg=SURFACE, fg=MUTED, font=f["small"]).pack(anchor="w", padx=16)
        tk.Label(c, text="24-month situación timeline (old → new):", bg=SURFACE, fg=MUTED,
                 font=f["small"]).pack(anchor="w", padx=16, pady=(8, 2))
        tl = tk.Canvas(c, bg=SURFACE, highlightthickness=0, height=44)
        tl.pack(fill="x", padx=16, pady=(0, 14))
        tl.bind("<Configure>", lambda e, cv=tl, s=a_bcra["worst_series"]: self._draw_sit_timeline(cv, s))
        self.after(30, lambda cv=tl, s=a_bcra["worst_series"]: self._draw_sit_timeline(cv, s))

    def _draw_sit_timeline(self, cv, series):
        cv.delete("all")
        if not series:
            return
        w = cv.winfo_width() or 760
        n = len(series)
        pad = 2
        cw = (w - 2 * 16) / max(n, 1) - pad
        x = 16
        for lvl in series:
            cv.create_rectangle(x, 6, x + cw, 30, fill=SIT_COLOR.get(lvl, MUTED), outline="")
            cv.create_text(x + cw / 2, 38, text=str(lvl), font=self.fonts["small"], fill=MUTED)
            x += cw + pad

    # ---- status / exports ----
    def set_status(self, text):
        self.status.configure(text="  " + text)

    def _open_workbook(self):
        path = filedialog.askopenfilename(title="Open Impulso workbook (.xlsx)",
                                          filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
        if path:
            self.data.xlsx = Path(path)
            self.data.clients = {}
            self.reload()

    def _export_ranking(self):
        self._export_named(self._active if self._active in ("combined", "bcra", "impulso") else "combined")

    def _export_named(self, key):
        if not self.data.loaded_at:
            return
        if key == "combined":
            text = analyze_combined.render_md(self.data.combined)
            default = "_combined_ranking.md"
        elif key == "bcra":
            text = analyze_bcra.render_leaderboard_md(self.data.bcohort)
            default = "_ranking.md"
        else:
            text = analyze_impulso.render_leaderboard_md(self.data.icohort)
            default = "_impulso_ranking.md"
        path = filedialog.asksaveasfilename(title="Export ranking", defaultextension=".md", initialfile=default)
        if path:
            Path(path).write_text(text + "\n", encoding="utf-8")
            self.set_status("Saved %s" % path)

    def _export_person(self):
        p = self._person_payload
        if not p:
            return
        parts = ["# Full Credit Dossier — %s" % p["name"], ""]
        if p["rec"]:
            parts += ["## Combined verdict", "", "```", analyze_combined.render_one(p["rec"]), "```"]
        if p["a_bcra"]:
            parts += ["", "---", "", analyze_bcra.render_md(p["a_bcra"], p["bcra_rank"])]
        if p["a_imp"]:
            parts += ["", "---", "", analyze_impulso.render_client_md(p["a_imp"], p["imp_rank"])]
        stem = p["cuil"] or (p["a_imp"]["dni"] if p["a_imp"] else "") or "person"
        path = filedialog.asksaveasfilename(title="Export dossier", defaultextension=".md",
                                            initialfile="%s.full.report.md" % stem)
        if path:
            Path(path).write_text("\n".join(parts) + "\n", encoding="utf-8")
            self.set_status("Saved %s" % path)

    def _about(self):
        messagebox.showinfo(
            "About",
            "Impulso Crédito — Credit Dashboard\n\n"
            "A Tkinter front-end over analyze_bcra / analyze_impulso / analyze_combined.\n"
            "Every figure matches the report.py CLI. Workbook: %s" % self.data.xlsx)


def _reexec_with_modern_tk():
    """Apple's system Tk 8.5 paints blank windows on modern macOS, so if we were
    launched on an old Tk, transparently re-exec into a Python that ships Tk ≥ 8.6
    (Homebrew python3.14, etc.). A one-shot env flag prevents any re-exec loop."""
    import os, sys, shutil, subprocess
    try:
        tkver = float(tk.TkVersion)  # tkinter is imported as `tk` at module top
    except (TypeError, ValueError):
        tkver = 0.0
    if tkver >= 8.6 or os.environ.get("IMPULSO_GUI_REEXEC") == "1":
        return
    names = ["python3.14", "python3.13", "python3.12", "python3.11", "python3.10"]
    candidates = (["/opt/homebrew/bin/" + n for n in names]
                  + ["/usr/local/bin/" + n for n in names]
                  + [shutil.which(n) for n in names])
    tried = set()
    for py in candidates:
        if not py or py in tried or not os.path.exists(py):
            continue
        tried.add(py)
        try:
            probe = subprocess.run([py, "-c", "import tkinter;print(tkinter.TkVersion)"],
                                   capture_output=True, text=True, timeout=10)
            if probe.returncode == 0 and float(probe.stdout.strip() or 0) >= 8.6:
                sys.stderr.write("report_gui: system Tk %.1f paints blank windows — "
                                 "switching to %s (Tk %s).\n" % (tkver, py, probe.stdout.strip()))
                env = dict(os.environ, IMPULSO_GUI_REEXEC="1")
                os.execve(py, [py, os.path.abspath(__file__)] + sys.argv[1:], env)
        except (OSError, ValueError, subprocess.SubprocessError):
            continue
    sys.stderr.write("report_gui: WARNING — running on Tk %.1f, which paints blank windows on "
                     "macOS, and no Tk≥8.6 Python was found. Install one:\n"
                     "    brew install python-tk@3.14\n" % tkver)


def main():
    import argparse
    _reexec_with_modern_tk()
    ap = argparse.ArgumentParser(description="Impulso credit dashboard (GUI).")
    ap.add_argument("--xlsx", default=str(DEFAULT_XLSX))
    ap.add_argument("--out", default=str(DEFAULT_OUT))
    args = ap.parse_args()
    App(args.xlsx, args.out).mainloop()


if __name__ == "__main__":
    main()
