# metastack Card Scanner

A buylist tool for Yu-Gi-Oh: sellers send binder photos, you scan them, and
out comes an Excel workbook the *seller* corrects — printing and condition
dropdowns per card, live price formulas, and a computed offer. No app needed
on their end, just Excel.

## The workflow

1. **Scan** — pick the seller's binder photos, hit Scan. Claude vision reads
   each card's name, set code, and foil treatment (it handles glare and foil
   text that break normal OCR). Cards resolve against two databases; anything
   uncertain shows in orange for a click.
2. **Export to Excel** — every candidate printing of every card is priced
   (lowest verified TCGPlayer listing, all five conditions) and written into
   a workbook with per-row dropdowns. Send it to the seller.
3. **Seller corrects in Excel** — picking a different printing or condition
   updates the price instantly (formulas, no internet needed). Wrong card
   entirely? They type the right name or set code into the row.
4. **Re-price a sent workbook** — run it on the returned file: typed
   corrections re-resolve, broken rows self-heal, every price refreshes to
   current market, and the seller's choices survive. The offer block
   (bulk tiers + discount rate) recomputes live.

## The offer math

Two editable cells in the workbook drive it: a **bulk pricepoint** (default
$1) and an **offer rate** (default 65%). Cards priced under the pricepoint
total as bulk — commons/rares and bulk-binder foils at $5 per thousand,
everything else at $30 per thousand — and cards at or above it total at
market × rate.

## Setup

```
pip install anthropic openpyxl requests pillow
```

Needs the sibling [yugioh-tcgplayer-pricer](https://github.com/matthewbraner/yugioh-tcgplayer-pricer)
repo checked out next to this one (as `TCGP Scraper/`) — pricing and catalog
matching import from it.

Run with `python gui.py`. Scanning needs an Anthropic API key (paste it into
the field bottom-left; stored in `%APPDATA%\metastack card scanner\`, about
3¢ per binder page on Sonnet). Pricing and re-pricing need no key.

## Building the exe

```
python -m PyInstaller --noconfirm --onefile --windowed --name "metastack Card Scanner" --icon "assets\brand\app.ico" --add-data "fonts;fonts" --add-data "assets;assets" --paths "..\TCGP Scraper" gui.py
```

**Use `python -m PyInstaller`, not bare `pyinstaller`** — a PATH pyinstaller
from a different Python ships an exe silently missing packages. Verify any
build with `"dist\metastack Card Scanner.exe" --check`, which writes
per-dependency results to `%APPDATA%\metastack card scanner\scan.log` (the
same file scan timings and errors go to).

## Data sources and the quirks handled

- **[YGOPRODeck](https://db.ygoprodeck.com/api-guide/)** — card names, sets,
  printings. Free, no key. Its per-card set lists lag reality (missing
  reprints, empty lists for new sets) and ship placeholder rarities ("New")
  and HTML entities — all worked around below.
- **[tcgcsv.com](https://tcgcsv.com)** — a free mirror of TCGPlayer's
  catalog. Built into a full (set code, rarity) → product index, cached a
  week in `%APPDATA%`. This is the fallback source of truth whenever
  YGOPRODeck's data is missing or wrong, and set codes are unique game-wide.
- **TCGPlayer's internal listings API** — prices, via the scraper repo.
  Unofficial; can break without notice.

Resolution is layered: exact set code → partial code (`SKE-EN0??` wildcards)
→ garbage code degraded to name search → name-merge across both databases.
Fuzzy name rescues need 0.8 similarity unless the card's printed text
corroborates (a blank row beats a confidently wrong one on a buylist).
Rarity names reconcile across databases (apostrophe variants, TCGPlayer's
"Prismatic Collector's Rare" for YGOPRODeck's "Collector's Rare").

## Files

| File | Role |
|---|---|
| `gui.py` | The window (tkinter, editorial house style) |
| `scan_cards.py` | Vision reads + YGOPRODeck lookups |
| `buylist.py` | TCGPlayer catalog index, pricing, workbook writer, re-price |

Each has an offline `--selftest`. The icon spec lives in the sibling
`App Icons/` gallery (`card-scanner`).
