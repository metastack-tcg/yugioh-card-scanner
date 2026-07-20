"""Seller-facing buylist workbook: price every candidate printing of each
scanned card (via the TCGP Scraper package), then write an .xlsx where the
seller picks the printing from a dropdown and Set / Price / Verified / Proof
update by formula. No app needed on their end — just Excel.

    python buylist.py --selftest   # offline check of the workbook writer
"""
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

# dev runs import the scraper from its sibling repo; the frozen exe has the
# package bundled by PyInstaller (--paths), so the sibling won't exist there
_SCRAPER = Path(__file__).parent.parent / "TCGP Scraper"
if _SCRAPER.exists():
    sys.path.insert(0, str(_SCRAPER))
from ygo_tcgplayer_pricer import tcgplayer_catalog, tcgplayer_pricing  # noqa: E402

CONDITIONS = tcgplayer_pricing.CONDITIONS  # NM..Damaged — seller picks per row in Excel
DEFAULT_CONDITION = tcgplayer_pricing.DEFAULT_CONDITION

INDEX_CACHE = Path(os.environ.get("APPDATA", Path.home())) / "metastack card scanner" / "tcgp_index_v2.json"
INDEX_MAX_AGE = 7 * 86400  # ponytail: weekly rebuild; new sets mostly match by group name anyway
_IDX = None


def _global_index():
    """(set_code, rarity) -> product across the ENTIRE TCGPlayer Yu-Gi-Oh
    catalog. Fallback for printings whose set can't be matched by group name —
    e.g. video-game promos TCGPlayer files under catch-all groups. Set codes
    are unique game-wide, so this is safe. ~600 tcgcsv fetches on first build,
    then cached to disk."""
    global _IDX
    if _IDX is not None:
        return _IDX
    if INDEX_CACHE.exists() and time.time() - INDEX_CACHE.stat().st_mtime < INDEX_MAX_AGE:
        try:
            raw = json.loads(INDEX_CACHE.read_text())
            _IDX = tuple({tuple(k.split("|", 1)): v for k, v in part.items()}
                         for part in (raw["base"], raw["ea"]))
            return _IDX
        except Exception:
            pass  # corrupt cache — rebuild below

    base, ea = {}, {}
    gname = {g["groupId"]: g["name"] for g in tcgplayer_catalog._all_groups()}
    failed = 0
    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = {pool.submit(tcgplayer_catalog.get_product_lookup, gid): gid
                   for gid in gname}
        for f in as_completed(futures):
            gid = futures[f]
            try:
                by_r, ea_r, _ = f.result()
            except Exception:
                failed += 1  # one throttled group must not sink the other 650
                continue
            for k, v in by_r.items():
                base.setdefault(k, dict(v, group=gname[gid]))
            for k, v in ea_r.items():
                ea.setdefault(k, dict(v, group=gname[gid]))
    if failed and failed >= len(gname) // 2:
        raise RuntimeError(f"TCGPlayer catalog unreachable ({failed} groups failed)")
    INDEX_CACHE.parent.mkdir(parents=True, exist_ok=True)
    INDEX_CACHE.write_text(json.dumps(
        {"base": {"|".join(k): v for k, v in base.items()},
         "ea": {"|".join(k): v for k, v in ea.items()}}))
    _IDX = (base, ea)
    return _IDX


def _price_jobs(jobs, progress=None):
    """{(product_id, condition): priced_or_None} — concurrent, deduped by caller."""
    prices = {}
    with ThreadPoolExecutor(max_workers=tcgplayer_pricing.MAX_WORKERS) as pool:
        futures = {pool.submit(tcgplayer_pricing.get_lowest_price_safe,
                               pid, condition=cond): (pid, cond)
                   for pid, cond in jobs}
        for i, f in enumerate(as_completed(futures), 1):
            prices[futures[f]] = f.result()
            if progress:
                progress(i, len(jobs))
    return prices


def tcgp_candidates(name, read_code):
    """Candidate printings straight from the TCGPlayer catalog, keyed by the
    set code read off the physical card. Fallback for when ygoprodeck's
    per-card set list is missing the printing in hand (their lists lag —
    e.g. Dark Magician's omits Maximum Gold: El Dorado entirely). The code
    pattern narrows the field; a name-similarity floor keeps out neighbours
    on the same page."""
    import difflib

    import scan_cards

    gbase, _ = _global_index()
    out = {}
    for (number, rarity), prod in gbase.items():
        # an empty read_code matches every number — name-only mode, for cards
        # ygoprodeck knows but has no printings for (brand-new sets)
        if read_code and not scan_cards.code_matches(read_code, number):
            continue
        pname = prod.get("name", "").split(" (")[0]  # strip "(Extended Art)" etc.
        m = difflib.SequenceMatcher(None, name.lower(), pname.lower())
        if m.quick_ratio() < 0.75 or m.ratio() < 0.75:  # quick_ratio gates the slow call
            continue
        e = out.setdefault(number, {"set_code": number,
                                    "set_name": prod.get("group", ""), "rarities": []})
        if rarity.title() not in e["rarities"]:
            e["rarities"].append(rarity.title())
    return [dict(e, rarities=sorted(e["rarities"])) for _, e in sorted(out.items())]


def build_options(cards, progress=None):
    """Attach card["options"] — one entry per candidate (printing, rarity),
    priced. Options span every candidate set so the seller can still correct
    a wrong auto-resolution in Excel."""
    for card in cards:
        card["options"] = []
        for cand in card["sets"]:
            gid = tcgplayer_catalog.find_group_id(cand["set_name"])
            by_rarity, ea_by_rarity, _ = (
                tcgplayer_catalog.get_product_lookup(gid) if gid else ({}, {}, {}))
            for rarity in cand["rarities"]:
                key = (cand["set_code"], rarity.strip().lower())
                base = by_rarity.get(key)
                ea_key = ea_by_rarity.get(key)
                if base is None and ea_key is None:
                    gbase, gea = _global_index()
                    base, ea_key = gbase.get(key), gea.get(key)
                label = f"{cand['set_code']} · {rarity}"
                if base and base.get("is_extended_art"):
                    label += " (Extended Art)"
                card["options"].append(
                    {"label": label, "set_name": cand["set_name"], "product": base})
                if ea_key:
                    card["options"].append(
                        {"label": f"{cand['set_code']} · {rarity} (Extended Art)",
                         "set_name": cand["set_name"], "product": ea_key})

    # one request per (product, condition) — 5x the calls of NM-only, so a
    # binder page lands around 100-250 requests; dedup + the scraper's thread
    # pool keep it to about a minute
    jobs = {(o["product"]["productId"], cond)
            for c in cards for o in c["options"] if o["product"]
            for cond in CONDITIONS}
    prices = _price_jobs(jobs, progress)

    for c in cards:
        for o in c["options"]:
            o["url"] = o["product"]["url"] if o["product"] else None
            o["prices"], o["verified"] = {}, {}
            for cond in CONDITIONS:
                priced = prices.get((o["product"]["productId"], cond)) if o["product"] else None
                o["prices"][cond] = priced["price"] if priced else None
                o["verified"][cond] = priced["verified"] if priced else None


HEADERS = ["Card Name", "Printing (choose)", "Set", "Condition", "Qty",
           "Price/Unit", "Price", "Verified", "Link/Proof", "Photo"]
TINT = "F6ECE4"  # house accent tint — marks rows still needing a choice


def write_workbook(path, cards, timestamp=""):
    """cards must already carry ["options"] from build_options."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()
    ws = wb.active
    ws.title = "Cards"
    opts = wb.create_sheet("Options")
    opts.sheet_state = "hidden"
    # A label · B set · C url · D-H price per condition · I-M verified per condition
    opts.append(["Label", "Set", "URL"] + CONDITIONS + CONDITIONS)

    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    if timestamp:
        ws.append([])  # spacer under a note row keeps the table clean
        ws["A2"] = ("Prices: lowest verified TCGPlayer listing in the chosen "
                    f"condition, {timestamp}. A blank price means no verified "
                    "listing right now.")
        ws["A2"].font = Font(italic=True, size=9)

    cond_dv = DataValidation(
        type="list", formula1='"' + ",".join(CONDITIONS) + '"',
        allow_blank=False, showErrorMessage=True)
    ws.add_data_validation(cond_dv)

    r = 3 if timestamp else 2
    orow = 2
    for c in cards:
        ws.cell(r, 1, c["name"])
        ws.cell(r, 5, c.get("qty", 1))
        ws.cell(r, 10, c["photo"])
        if c["options"]:
            lo, hi = orow, orow + len(c["options"]) - 1
            for o in c["options"]:
                def vtext(cond):
                    if o["prices"][cond] is None:
                        return ""
                    return "yes" if o["verified"][cond] else "no"
                opts.append([o["label"], o["set_name"], o["url"] or ""]
                            + [o["prices"][cond] for cond in CONDITIONS]
                            + [vtext(cond) for cond in CONDITIONS])
            orow = hi + 1
            labels = f"Options!$A${lo}:$A${hi}"
            conds = "Options!$D$1:$H$1"
            row_m = f"MATCH($B{r},{labels},0)"
            col_m = f"MATCH($D{r},{conds},0)"
            dv = DataValidation(type="list", formula1=labels,
                                allow_blank=True, showErrorMessage=True)
            ws.add_data_validation(dv)
            dv.add(ws.cell(r, 2).coordinate)
            cond_dv.add(ws.cell(r, 4).coordinate)
            # IF(...="","",...) wraps: INDEX/VLOOKUP render an empty cell as 0,
            # and a missing listing must read as blank, not "$0.00"
            price_ix = f"INDEX(Options!$D${lo}:$H${hi},{row_m},{col_m})"
            verif_ix = f"INDEX(Options!$I${lo}:$M${hi},{row_m},{col_m})"
            url_vl = f"VLOOKUP($B{r},Options!$A${lo}:$C${hi},3,FALSE)"
            ws.cell(r, 3).value = (f'=IFERROR(VLOOKUP($B{r},Options!$A${lo}:$B${hi},2,FALSE),"")')
            ws.cell(r, 4).value = DEFAULT_CONDITION
            ws.cell(r, 6).value = f'=IFERROR(IF({price_ix}="","",{price_ix}),"")'
            ws.cell(r, 6).number_format = "$0.00"
            ws.cell(r, 7).value = f'=IFERROR($E{r}*$F{r},"")'
            ws.cell(r, 7).number_format = "$0.00"
            ws.cell(r, 8).value = f'=IFERROR(IF({verif_ix}="","",{verif_ix}),"")'
            ws.cell(r, 9).value = f'=IFERROR(IF({url_vl}="","",HYPERLINK({url_vl})),"")'

            labels = [o["label"] for o in c["options"]]
            pre = f"{c['set_code']} · {c['rarity']}" if c["set_code"] and c["rarity"] else None
            # an Extended Art foil read outranks the plain printing of the same rarity
            tries = (pre, pre and pre + " (Extended Art)")
            if pre and "extended art" in (c.get("guess") or "").lower():
                tries = (pre + " (Extended Art)", pre)
            pick = next((v for v in tries if v in labels), None)
            if pick:
                ws.cell(r, 2).value = pick
            else:
                ws.cell(r, 2).fill = PatternFill("solid", fgColor=TINT)
        r += 1

    first = 3 if timestamp else 2
    if r > first:
        ws.cell(r + 1, 1, "Total").font = Font(bold=True)
        ws.cell(r + 1, 2).value = (
            f'=COUNTA(B{first}:B{r - 1})&" of "&COUNTA(A{first}:A{r - 1})&" cards chosen"')
        ws.cell(r + 1, 5).value = f"=SUM(E{first}:E{r - 1})"  # total physical cards
        total = ws.cell(r + 1, 7)
        total.value = f"=SUM(G{first}:G{r - 1})"
        total.number_format = "$0.00"
        total.font = Font(bold=True)

    for col, width in zip("ABCDEFGHIJ", (32, 34, 34, 13, 5, 10, 10, 9, 40, 16)):
        ws.column_dimensions[col].width = width
    wb.save(path)


PRODUCT_URL_RE = re.compile(r"/product/(\d+)")


def refresh_workbook(path, timestamp="", progress=None):
    """Re-price a workbook the seller sent back. Only the hidden Options
    sheet's price/verified cells and the note row are touched — the seller's
    dropdown choices, added columns, and everything else stay as-is. The
    Cards-sheet formulas pick up the new numbers when the file next opens.
    Returns the number of printings re-priced."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    opts = wb["Options"]
    rows = []  # (options_row, product_id)
    for r in range(2, opts.max_row + 1):
        m = PRODUCT_URL_RE.search(opts.cell(r, 3).value or "")
        if m:
            rows.append((r, int(m.group(1))))

    prices = _price_jobs({(pid, cond) for _, pid in rows for cond in CONDITIONS},
                         progress)

    for r, pid in rows:
        for i, cond in enumerate(CONDITIONS):
            priced = prices.get((pid, cond))
            opts.cell(r, 4 + i).value = priced["price"] if priced else None
            opts.cell(r, 9 + i).value = (
                "" if priced is None else ("yes" if priced["verified"] else "no"))

    ws = wb["Cards"]
    if timestamp and str(ws["A2"].value or "").startswith("Prices:"):
        ws["A2"] = ("Prices: lowest verified TCGPlayer listing in the chosen "
                    f"condition, {timestamp}. A blank price means no verified "
                    "listing right now.")
    wb.save(path)
    return len(rows)


def selftest():
    import os
    import tempfile
    from openpyxl import load_workbook

    global _IDX
    _IDX = ({("MZMU-EN003", "ultra rare"):
             {"productId": 1, "url": "u", "name": "Stare of the Snake Hair", "group": "Set A"},
             ("SBLS-EN026", "common"):
             {"productId": 2, "url": "u", "name": "The Snake Hair", "group": "Set B"}}, {})
    try:
        cands = tcgp_candidates("Stare of the Snake-Hair", "")   # name-only mode
        assert cands == [{"set_code": "MZMU-EN003", "set_name": "Set A",
                          "rarities": ["Ultra Rare"]}], cands
        assert tcgp_candidates("Stare of the Snake-Hair", "SBLS-EN026") == []  # code excludes
    finally:
        _IDX = None

    def opt(label, nm_price, pid=0):
        return {"label": label, "set_name": "25th Anniversary Rarity Collection",
                "url": f"https://www.tcgplayer.com/product/{pid}/x" if nm_price else None,
                "prices": {c: (nm_price if c == "Near Mint" else None) for c in CONDITIONS},
                "verified": {c: (True if c == "Near Mint" and nm_price else None)
                             for c in CONDITIONS}}

    cards = [
        {"name": "Fallen of Albaz", "photo": "a.jpg", "set_code": "RA01-EN021",
         "rarity": "Secret Rare", "sets": [],
         "options": [opt("RA01-EN021 · Secret Rare", 1.23, 524540),
                     opt("RA01-EN021 · Ultra Rare", None)]},
        {"name": "Mystery Card", "photo": "a.jpg", "set_code": "", "rarity": "",
         "sets": [], "options": []},
        {"name": "Witness of the Ancient", "photo": "b.jpg", "set_code": "CORI-EN012",
         "rarity": "Ultra Rare", "guess": "ultra rare (extended art)", "qty": 3,
         "sets": [],
         "options": [opt("CORI-EN012 · Ultra Rare", 0.50, 611001),
                     opt("CORI-EN012 · Ultra Rare (Extended Art)", 4.00, 611002)]},
    ]
    with tempfile.TemporaryDirectory() as d:
        p = os.path.join(d, "t.xlsx")
        write_workbook(p, cards, "2026-07-19")
        wb = load_workbook(p)
        ws = wb["Cards"]
        assert wb["Options"].sheet_state == "hidden"
        assert wb["Options"]["D1"].value == "Near Mint"           # MATCH header row
        assert ws["B3"].value == "RA01-EN021 · Secret Rare"       # preselected
        assert ws["D3"].value == "Near Mint"                      # default condition
        assert ws["E3"].value == 1 and ws["E5"].value == 3        # qty column
        assert "IF(INDEX" in ws["F3"].value and "MATCH($D3" in ws["F3"].value, ws["F3"].value
        assert ws["G3"].value == '=IFERROR($E3*$F3,"")'           # qty x unit
        # one printing dropdown per card with options + the shared condition list
        assert len(ws.data_validations.dataValidation) == 3
        assert ws["B4"].value is None                             # unresolved left blank
        assert ws["B5"].value == "CORI-EN012 · Ultra Rare (Extended Art)"  # EA guess wins

        # refresh: seller edits survive, prices change, no network (stubbed)
        ws["B3"] = "RA01-EN021 · Ultra Rare"  # seller corrected the printing
        ws["D3"] = "Lightly Played"           # and the condition
        wb.save(p)
        real = tcgplayer_pricing.get_lowest_price_safe
        tcgplayer_pricing.get_lowest_price_safe = (
            lambda pid, condition=None, **kw: {"price": 9.99, "seller": "s", "verified": True})
        try:
            n = refresh_workbook(p, "2026-08-01")
        finally:
            tcgplayer_pricing.get_lowest_price_safe = real
        wb2 = load_workbook(p)
        assert n == 3, n                                          # only rows with URLs
        assert wb2["Options"]["D2"].value == 9.99                 # re-priced
        assert wb2["Cards"]["B3"].value == "RA01-EN021 · Ultra Rare"   # seller choice kept
        assert wb2["Cards"]["D3"].value == "Lightly Played"
        assert "2026-08-01" in wb2["Cards"]["A2"].value
        assert wb2["Cards"]["A7"].value == "Total"
        assert wb2["Cards"]["G7"].value == "=SUM(G3:G5)", wb2["Cards"]["G7"].value
    print("selftest OK")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--selftest":
        selftest()
