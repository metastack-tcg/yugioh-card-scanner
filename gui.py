"""metastack card scanner — pick binder photos, read the cards, export to Excel.

    python gui.py

Styled to the editorial house style (paper/ink/one accent, hairlines, serif
display). Vision + database work lives in scan_cards.py; this file is only the
window. Needs ANTHROPIC_API_KEY set to scan.
"""
import ctypes
import os
import threading
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import filedialog, messagebox

import scan_cards

# --- house style tokens (light) ---------------------------------------------
PAPER, INK = "#FAF8F2", "#15130E"
MUTED, FAINT, HAIRLINE = "#6B6353", "#9A8E78", "#E4E0D6"
ACCENT, ON_ACCENT, ACCENT_TINT = "#C2410C", "#FAF8F2", "#F6ECE4"

FONTS = Path(__file__).parent / "fonts"
ICON = Path(__file__).parent / "assets" / "brand" / "app.ico"
APP_ID = "metastack.cardscanner.1"
CONFIG = Path(os.environ.get("APPDATA", Path.home())) / "metastack card scanner" / "config.json"
LOG = CONFIG.parent / "scan.log"


def log(msg):
    import datetime
    try:
        CONFIG.parent.mkdir(parents=True, exist_ok=True)
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now():%H:%M:%S} {msg}\n")
    except OSError:
        pass


def load_key():
    try:
        import json
        return json.loads(CONFIG.read_text())["api_key"]
    except Exception:
        return os.environ.get("ANTHROPIC_API_KEY", "")


def save_key(key):
    import json
    CONFIG.parent.mkdir(parents=True, exist_ok=True)
    CONFIG.write_text(json.dumps({"api_key": key}))


def load_fonts():
    """Register Fraunces for this process only — shared from the envelope printer."""
    FR_PRIVATE = 0x10
    for f in ("Fraunces-Regular.ttf", "Fraunces-Italic.ttf", "Fraunces-Bold.ttf"):
        p = FONTS / f
        if p.exists():
            ctypes.windll.gdi32.AddFontResourceExW(str(p), FR_PRIVATE, 0)


def claim_taskbar_identity():
    """Own taskbar identity instead of inheriting the host's. Run pre-window."""
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)


def set_window_icon(root):
    """iconbitmap only sets the window-class icon; the taskbar reads the
    per-window icon via WM_SETICON — set both sizes explicitly."""
    if not ICON.exists():
        return
    u = ctypes.windll.user32
    IMAGE_ICON, LR_LOADFROMFILE, WM_SETICON = 1, 0x0010, 0x0080
    root.update_idletasks()
    hwnd = u.GetParent(root.winfo_id()) or root.winfo_id()
    for which, metric in ((1, 11), (0, 49)):  # ICON_BIG/SM_CXICON, ICON_SMALL/SM_CXSMICON
        px = u.GetSystemMetrics(metric)
        h = u.LoadImageW(None, str(ICON), IMAGE_ICON, px, px, LR_LOADFROMFILE)
        if h:
            u.SendMessageW(hwnd, WM_SETICON, which, h)


class BrandMark(tk.Canvas):
    """The app mark, drawn from the icon spec — no asset needed."""

    def __init__(self, parent, px=14):
        super().__init__(parent, width=px, height=px, bg=PAPER, highlightthickness=0, bd=0)
        s = px / 512
        for x in (112, 216, 320):
            for y in (66, 200, 334):
                self.create_rectangle(x * s, y * s, (x + 80) * s, (y + 112) * s,
                                      fill=ACCENT if (x, y) == (216, 200) else INK,
                                      outline="")


def serif(size, weight="normal"):
    name = "Fraunces" if "Fraunces" in tkfont.families() else "Georgia"
    return (name, size, weight)


def sans(size, weight="normal"):
    return ("Segoe UI", size, weight)


def hairline(parent):
    return tk.Frame(parent, height=1, bg=HAIRLINE)


def kicker(parent, text):
    return tk.Label(parent, text=" ".join(text.upper()), font=sans(8),
                    fg=FAINT, bg=PAPER, anchor="w")


class UnderlineAction(tk.Frame):
    """Default action: text + accent underline."""

    def __init__(self, parent, text, command):
        super().__init__(parent, bg=PAPER)
        self.command = command
        self.lbl = tk.Label(self, text=text, font=sans(9), fg=ACCENT, bg=PAPER, cursor="hand2")
        self.lbl.pack(anchor="w")
        self.rule = tk.Frame(self, height=2, bg=ACCENT)
        self.rule.pack(fill="x")
        self.on = True
        self.lbl.bind("<Button-1>", lambda e: self.on and self.command())

    def enable(self, on):
        self.on = on
        self.lbl.config(fg=ACCENT if on else FAINT, cursor="hand2" if on else "arrow")
        self.rule.config(bg=ACCENT if on else HAIRLINE)

    def text(self, t):
        self.lbl.config(text=t)


def resolve_set(card, cand):
    """Fill a card's set fields from a chosen candidate printing."""
    card["set_code"] = cand["set_code"]
    card["set_name"] = cand["set_name"]
    card["rarities"] = cand["rarities"]
    card["rarity"] = default_rarity(cand["rarities"], card["guess"])


def default_rarity(rarities, guess):
    """Pick a rarity only when there's evidence: a single known printing, or a
    foil guess that matches one. Otherwise return "" — an unbacked pick would
    show up in the seller's Excel looking like a confident answer."""
    if not rarities:
        return guess or ""
    if len(rarities) == 1:
        return rarities[0]
    if guess:
        g = guess.lower()
        for r in rarities:
            if r.lower() == g:
                return r
        for r in rarities:
            if g in r.lower() or r.lower() in g:
                return r
    return ""


COLS = [("name", "Name", 3), ("set_code", "Set code", 1),
        ("set_name", "Set", 3), ("rarity", "Rarity", 2)]


class App:
    def __init__(self, root):
        self.root = root
        self.photos, self.cards = [], []
        self.scanning = self.working = False
        root.configure(bg=PAPER)
        root.title("metastack card scanner")
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        outer = tk.Frame(root, bg=PAPER)
        outer.pack(fill="both", expand=True, padx=24, pady=24)

        kicker(outer, "metastack").pack(fill="x")
        tk.Label(outer, text="Card scanner", font=serif(22), fg=INK, bg=PAPER,
                 anchor="w").pack(fill="x", pady=(2, 8))
        tk.Frame(outer, height=2, bg=INK).pack(fill="x")

        body = tk.Frame(outer, bg=PAPER)
        body.pack(fill="both", expand=True, pady=(24, 0))
        left = tk.Frame(body, bg=PAPER, width=280)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        right = tk.Frame(body, bg=PAPER)
        right.pack(side="left", fill="both", expand=True, padx=(32, 0))

        # --- bottom stack, packed FIRST so it always keeps its space (pack
        # gives earlier widgets priority; the top content clips, not this) ---
        self.key = tk.StringVar(value=load_key())
        entry = tk.Entry(left, textvariable=self.key, show="•", font=sans(9),
                         fg=INK, bg=PAPER, relief="solid", bd=1, highlightthickness=1,
                         highlightbackground=HAIRLINE, highlightcolor=ACCENT,
                         insertbackground=INK)
        entry.pack(side="bottom", fill="x", ipady=3)
        tk.Label(left, text="from console.anthropic.com — stored on this computer only",
                 font=sans(8), fg=FAINT, bg=PAPER, anchor="w",
                 wraplength=270, justify="left").pack(side="bottom", fill="x", pady=(2, 4))
        kicker(left, "anthropic api key").pack(side="bottom", fill="x", pady=(16, 0))
        self.refresh_act = UnderlineAction(left, "Re-price a sent workbook →", self.refresh)
        self.refresh_act.pack(side="bottom", anchor="w", pady=(8, 0))
        self.export_act = UnderlineAction(left, "Export to Excel →", self.export)
        self.export_act.pack(side="bottom", anchor="w", pady=(4, 0))
        self.export_act.enable(False)
        kicker(left, "export").pack(side="bottom", fill="x", pady=(16, 0))

        # --- left: photos, scan, export --------------------------------------
        kicker(left, "photos").pack(fill="x")
        self.src = tk.Label(left, text="No photos selected", font=sans(10), fg=MUTED,
                            bg=PAPER, anchor="w", wraplength=270, justify="left")
        self.src.pack(fill="x", pady=(4, 8))
        UnderlineAction(left, "Choose photos →", self.browse).pack(anchor="w")

        self.scan_btn = tk.Label(left, text="Scan", font=sans(10), fg=FAINT,
                                 bg=HAIRLINE, pady=8, cursor="arrow")
        self.scan_btn.pack(fill="x", pady=(16, 0))
        self.scan_btn.bind("<Button-1>", lambda e: self.scan())

        self.hero = tk.Label(left, text="", font=serif(44), fg=ACCENT, bg=PAPER, anchor="w")
        self.hero.pack(fill="x", pady=(20, 0))
        self.hero_sub = tk.Label(left, text="Nothing scanned yet", font=sans(9),
                                 fg=MUTED, bg=PAPER, anchor="w")
        self.hero_sub.pack(fill="x")

        # --- right: results table --------------------------------------------
        head = tk.Frame(right, bg=PAPER)
        head.pack(fill="x")
        kicker(head, "cards").pack(side="left")
        tk.Label(head, text="orange = couldn't be read for certain — click to choose",
                 font=sans(8), fg=FAINT, bg=PAPER).pack(side="right")

        cols = tk.Frame(right, bg=PAPER)
        cols.pack(fill="x", pady=(6, 0))
        cols.grid_columnconfigure(0, minsize=10)  # accent gutter
        for i, (_, label, weight) in enumerate(COLS, start=1):
            cols.grid_columnconfigure(i, weight=weight, uniform="c")
            kicker(cols, label).grid(row=0, column=i, sticky="w", padx=(0, 8))
        tk.Frame(right, height=1, bg=HAIRLINE).pack(fill="x", pady=(4, 0))

        wrap = tk.Frame(right, bg=PAPER)
        wrap.pack(fill="both", expand=True)
        self.rc = tk.Canvas(wrap, bg=PAPER, highlightthickness=0, bd=0)
        sb = tk.Scrollbar(wrap, orient="vertical", command=self.rc.yview, width=10)
        self.rc.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.rc.pack(side="left", fill="both", expand=True)
        self.rows = tk.Frame(self.rc, bg=PAPER)
        self._rows_win = self.rc.create_window((0, 0), window=self.rows, anchor="nw")
        self.rows.bind("<Configure>",
                       lambda e: self.rc.configure(scrollregion=self.rc.bbox("all")))
        self.rc.bind("<Configure>",
                     lambda e: self.rc.itemconfig(self._rows_win, width=e.width))
        self.rc.bind_all("<MouseWheel>", self._wheel)

        # --- brand footer -----------------------------------------------------
        foot = tk.Frame(outer, bg=PAPER)
        foot.pack(fill="x", pady=(20, 0))
        BrandMark(foot).pack(side="left")
        tk.Label(foot, text="© 2026 metastack.", font=sans(8), fg=FAINT,
                 bg=PAPER).pack(side="left", padx=6)
        self.status = tk.Label(foot, text="", font=sans(8), fg=FAINT, bg=PAPER, anchor="e")
        self.status.pack(side="right")

    def on_close(self):
        if self.scanning or self.working:
            if not messagebox.askyesno(
                    "Still working",
                    "A scan or export is still running — closing now loses it.\n"
                    "Close anyway?"):
                return
        self.root.destroy()

    def _wheel(self, e):
        if self.rows.winfo_height() > self.rc.winfo_height():
            self.rc.yview_scroll(-1 if e.delta > 0 else 1, "units")

    def say(self, msg):
        self.status.config(text=msg)

    # a visible heartbeat: the phase text plus a ticking elapsed counter, so
    # a slow network call never looks like a dead app
    def phase(self, text):
        import time
        self._phase, self._phase_t0 = text, time.time()

    def _tick(self):
        import time
        if self.scanning:
            if getattr(self, "_phase", None):
                self.say(f"{self._phase} — {int(time.time() - self._phase_t0)}s")
            self.root.after(1000, self._tick)

    # --- flow ----------------------------------------------------------------

    def browse(self):
        paths = filedialog.askopenfilenames(
            title="Binder photos",
            filetypes=[("Photos", "*.jpg *.jpeg *.png *.webp"), ("All files", "*.*")])
        if not paths:
            return
        import hashlib
        seen, unique, dupes = set(), [], 0
        for p in paths:
            h = hashlib.md5(Path(p).read_bytes()).hexdigest()
            if h in seen:
                dupes += 1
            else:
                seen.add(h)
                unique.append(p)
        self.photos = unique
        if dupes:
            self.say(f"Dropped {dupes} duplicate photo{'s' if dupes != 1 else ''}.")
        n = len(self.photos)
        names = [Path(p).name for p in self.photos]
        if len(names) > 6:
            names = names[:6] + [f"+ {len(names) - 6} more"]
        self.src.config(text="\n".join(names), fg=INK)
        self.scan_btn.config(bg=ACCENT, fg=ON_ACCENT, cursor="hand2",
                             text=f"Scan {n} photo{'s' if n != 1 else ''}")

    def scan(self):
        if not self.photos or self.scanning:
            return
        key = self.key.get().strip()
        if not key:
            messagebox.showerror(
                "No API key",
                "Paste your Anthropic API key in the field at the bottom left.\n"
                "Get one at console.anthropic.com → API keys.")
            return
        save_key(key)
        self.scanning = True
        self.cards = []
        self._page_sigs = {}
        self.phase("Starting scan")
        self._tick()
        self.build_rows()
        self.export_act.enable(False)
        self.scan_btn.config(bg=HAIRLINE, fg=FAINT, cursor="arrow", text="Scanning…")
        threading.Thread(target=self._scan_worker, args=(key,), daemon=True).start()

    def _scan_worker(self, key):
        # done() must run no matter what dies in here — a stray exception
        # otherwise kills the thread silently and the GUI hangs on "Scanning…"
        import anthropic
        client = anthropic.Anthropic(api_key=key)
        import time
        try:
            for i, path in enumerate(self.photos, 1):
                name = Path(path).name
                self.root.after(0, self.phase,
                                f"Reading {name} ({i} of {len(self.photos)})")
                t0 = time.time()
                try:
                    found = scan_cards.read_photo(client, path)
                except Exception as e:
                    log(f"{name}: vision FAILED after {time.time() - t0:.0f}s: {e}")
                    self.root.after(0, messagebox.showerror, "Scan failed",
                                    f"{name}:\n{e}")
                    continue
                log(f"{name}: vision {time.time() - t0:.0f}s, {len(found)} cards")
                sig = tuple((c["name"], c["set_code"]) for c in found)
                if found and sig in self._page_sigs:
                    if not self._ask(
                            "Duplicate page?",
                            f"{name} shows exactly the same cards as "
                            f"{self._page_sigs[sig]}.\n\nInclude it anyway?"):
                        continue
                self._page_sigs[sig] = name
                t0 = time.time()
                for c in found:
                    self.root.after(0, self.phase,
                                    f"Looking up {c['name']} ({i} of {len(self.photos)})")
                    try:
                        cname, cands = scan_cards.db_lookup(
                            c["name"], c["set_code"], c.get("text_snippet"))
                    except Exception:
                        cname, cands = c["name"], []
                    if not cands or (c["set_code"] and not any(
                            scan_cards.code_matches(c["set_code"], k["set_code"])
                            for k in cands)):
                        # ygoprodeck's set list is missing this printing (or is
                        # empty for a brand-new set) — trust the read code, or
                        # failing that the name, against the TCGPlayer catalog
                        try:
                            import buylist
                            self.root.after(0, self.phase,
                                            f"Checking the TCGPlayer catalog for {c['name']}")
                            cands = buylist.tcgp_candidates(
                                cname, c["set_code"] or "") or cands
                        except Exception as e:
                            log(f"  catalog fallback failed for {c['name']}: {e}")
                    card = {"photo": name, "name": cname,
                            "read_code": c["set_code"] or "",
                            "sets": cands, "guess": c["rarity_guess"] or "",
                            "set_code": "", "set_name": "", "rarities": [], "rarity": ""}
                    if len(cands) == 1:
                        resolve_set(card, cands[0])
                    self.root.after(0, self.add_card, card)
                log(f"{name}: lookups {time.time() - t0:.0f}s")
        finally:
            self.root.after(0, self.done)

    def _ask(self, title, msg):
        """askyesno from a worker thread — runs the dialog on the UI thread."""
        result, done = {}, threading.Event()

        def ask():
            result["ok"] = messagebox.askyesno(title, msg)
            done.set()

        self.root.after(0, ask)
        done.wait(timeout=300)  # if the dialog somehow never resolves, keep the page
        return result.get("ok", True)

    def add_card(self, card):
        # no auto-merging: identical vision reads can still be two different
        # printings, so every physical card keeps its own row — the seller
        # consolidates via the editable Qty column in Excel if they want to
        self.cards.append(card)
        self.build_rows()
        n = len(self.cards)
        self.hero.config(text=str(n))
        self.hero_sub.config(text=f"card{'s' if n != 1 else ''} found")

    def done(self):
        self.scanning = False
        self._phase = None
        self.scan_btn.config(bg=ACCENT, fg=ON_ACCENT, cursor="hand2",
                             text=f"Scan {len(self.photos)} photo{'s' if len(self.photos) != 1 else ''}")
        self.export_act.enable(bool(self.cards))
        ambiguous = sum(1 for c in self.cards
                        if len(c["rarities"]) > 1 or (len(c["sets"]) > 1 and not c["set_code"]))
        self.say(f"{ambiguous} card{'s need' if ambiguous != 1 else ' needs'} a choice."
                 if ambiguous else "Done.")

    # --- table ----------------------------------------------------------------

    def build_rows(self):
        for w in self.rows.winfo_children():
            w.destroy()
        for idx, c in enumerate(self.cards):
            row = tk.Frame(self.rows, bg=PAPER)
            row.pack(fill="x")
            row.grid_columnconfigure(0, minsize=10)
            for i, (key, _, weight) in enumerate(COLS, start=1):
                row.grid_columnconfigure(i, weight=weight, uniform="c")
            set_open = len(c["sets"]) > 1 and not c["set_code"]
            vals = {"name": (c["name"], INK)}
            for i, (key, _, _) in enumerate(COLS, start=1):
                click = None
                if key == "rarity":
                    multi = len(c["rarities"]) > 1
                    text = c["rarity"] or ("—" if set_open else c["guess"] or "—")
                    fg, click = (ACCENT, self.pick_rarity) if multi else (INK, None)
                elif key == "set_code":
                    if set_open:
                        text = f"{c['read_code'] or '?'} · choose"
                        fg, click = ACCENT, self.pick_set
                    else:
                        text, fg = c["set_code"] or c["read_code"] or "—", FAINT
                elif key == "set_name":
                    text = c["set_name"] or (f"{len(c['sets'])} possible" if set_open else "—")
                    fg = MUTED
                else:
                    text, fg = vals[key]
                lbl = tk.Label(row, text=text, font=sans(9), fg=fg, bg=PAPER,
                               anchor="w", pady=5, cursor="hand2" if click else "arrow")
                if click:
                    lbl.bind("<Button-1>", lambda e, i=idx, w=lbl, f=click: f(i, w))
                lbl.grid(row=0, column=i, sticky="w", padx=(0, 8))
            hairline(self.rows).pack(fill="x")

    def pick_set(self, idx, widget):
        c = self.cards[idx]
        m = tk.Menu(self.root, tearoff=0, font=sans(9), bg=PAPER, fg=INK,
                    activebackground=ACCENT_TINT, activeforeground=ACCENT, bd=0)
        for cand in c["sets"]:
            m.add_command(label=f"{cand['set_code']} — {cand['set_name']}",
                          command=lambda cand=cand: self.set_set(idx, cand))
        m.tk_popup(widget.winfo_rootx(), widget.winfo_rooty() + widget.winfo_height())

    def set_set(self, idx, cand):
        resolve_set(self.cards[idx], cand)
        self.build_rows()

    def pick_rarity(self, idx, widget):
        c = self.cards[idx]
        m = tk.Menu(self.root, tearoff=0, font=sans(9), bg=PAPER, fg=INK,
                    activebackground=ACCENT_TINT, activeforeground=ACCENT, bd=0)
        for r in c["rarities"]:
            m.add_command(label=r, command=lambda r=r: self.set_rarity(idx, r))
        m.tk_popup(widget.winfo_rootx(), widget.winfo_rooty() + widget.winfo_height())

    def set_rarity(self, idx, rarity):
        self.cards[idx]["rarity"] = rarity
        self.cards[idx]["rarities"] = [rarity]  # resolved — drop the accent affordance
        self.build_rows()

    # --- export ---------------------------------------------------------------

    def export(self):
        if not self.cards or self.scanning:
            return
        import datetime
        path = filedialog.asksaveasfilename(
            title="Export buylist", defaultextension=".xlsx",
            initialfile=f"buylist-{datetime.date.today().isoformat()}.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self.export_act.enable(False)
        self.working = True
        threading.Thread(target=self._export_worker, args=(path,), daemon=True).start()

    def _export_worker(self, path):
        import datetime
        import buylist

        def note(t):
            self.root.after(0, self.export_act.text, t)

        try:
            note("Looking up printings…")
            buylist.build_options(
                self.cards,
                progress=lambda d, t: note(f"Pricing… {d} of {t}"))
            note("Writing workbook…")
            buylist.write_workbook(path, self.cards,
                                   datetime.date.today().isoformat())
        except Exception as e:
            self.root.after(0, messagebox.showerror, "Export failed", str(e))
        else:
            self.root.after(0, self._export_done, path)
        self.working = False
        note("Export to Excel →")
        self.root.after(0, self.export_act.enable, True)

    def _export_done(self, path):
        self.say(f"Exported {len(self.cards)} cards.")
        if messagebox.askyesno("Exported",
                               f"Saved {Path(path).name}.\n\nOpen its folder?"):
            os.startfile(Path(path).parent)

    def refresh(self):
        path = filedialog.askopenfilename(
            title="Re-price a buylist workbook", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self.refresh_act.enable(False)
        self.working = True
        threading.Thread(target=self._refresh_worker, args=(path,), daemon=True).start()

    def _refresh_worker(self, path):
        import datetime
        import buylist
        try:
            n = buylist.refresh_workbook(
                path, datetime.date.today().isoformat(),
                progress=lambda d, t: self.root.after(
                    0, self.refresh_act.text, f"Re-pricing… {d} of {t}"))
        except PermissionError:
            self.root.after(0, messagebox.showerror, "File is open",
                            "Close the workbook in Excel first, then try again.")
            self.root.after(0, self.say, "")
        except Exception as e:
            self.root.after(0, messagebox.showerror, "Re-price failed", str(e))
            self.root.after(0, self.say, "")
        else:
            self.root.after(0, self.say,
                            f"Re-priced {n} printings in {Path(path).name}.")
        self.working = False
        self.root.after(0, self.refresh_act.text, "Re-price a sent workbook →")
        self.root.after(0, self.refresh_act.enable, True)


def selftest():
    assert default_rarity(["Ultra Rare"], None) == "Ultra Rare"
    assert default_rarity(["Secret Rare", "Quarter Century Secret Rare"],
                          "quarter century secret rare") == "Quarter Century Secret Rare"
    assert default_rarity(["Secret Rare", "Ultra Rare"], "secret") == "Secret Rare"
    assert default_rarity(["Secret Rare", "Ultra Rare"], None) == ""  # no evidence, no guess
    assert default_rarity(["Secret Rare"], None) == "Secret Rare"     # only one printing
    assert default_rarity([], "Common") == "Common"
    print("selftest OK")  # workbook writing is tested in buylist.py --selftest


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--selftest":
        selftest()
        sys.exit()
    claim_taskbar_identity()
    load_fonts()
    root = tk.Tk()
    root.geometry("1080x680")
    root.minsize(920, 560)
    try:
        root.iconbitmap(default=str(ICON))  # default= so dialogs inherit it
    except tk.TclError:
        pass
    set_window_icon(root)
    App(root)
    root.mainloop()
