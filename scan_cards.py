"""Read Yu-Gi-Oh card names/sets/rarities from binder photos.

The engine behind gui.py: Claude vision reads each photo (handles foils/glare
that break normal OCR), then each card is cross-referenced against the free
YGOPRODeck database to resolve printings from the set code.

    python scan_cards.py --selftest    # live check of the DB lookup
"""
import base64
import json
import sys
import urllib.parse
import urllib.request

MODEL = "claude-sonnet-5"  # high-res vision at ~3c/page; opus-4-8 if accuracy ever disappoints

SCHEMA = {
    "type": "object",
    "properties": {
        "cards": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string"},
                    "set_code": {
                        "type": ["string", "null"],
                        "description": "Printed set code e.g. DOOD-EN039. Use ? for any "
                                       "character you can't read (e.g. SKE-EN0??). "
                                       "Null only if fully unreadable.",
                    },
                    "rarity_guess": {
                        "type": ["string", "null"],
                        "description": "Rarity judged from foil pattern/name color, null if "
                                       "unsure. Note special variants when visible: a "
                                       "'Quarter Century' foil stamp means Quarter Century "
                                       "Secret Rare; artwork overflowing its frame means "
                                       "Extended Art (append '(Extended Art)').",
                    },
                    "text_snippet": {
                        "type": ["string", "null"],
                        "description": "Usually null. ONLY when the card name is partially "
                                       "obscured or you are not fully confident in your "
                                       "reading of it: give the first ~12 words of the "
                                       "effect/flavor text exactly as printed — the matte "
                                       "text box is more legible than a foil name and is "
                                       "used to verify the match.",
                    },
                },
                "required": ["name", "set_code", "rarity_guess", "text_snippet"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["cards"],
    "additionalProperties": False,
}

PROMPT = (
    "List every Yu-Gi-Oh card visible in this photo, in reading order "
    "(left-to-right, top-to-bottom). For each card give its exact printed name, "
    "the set code printed near the artwork (e.g. DOOD-EN039) if legible, and "
    "your best guess at the rarity from the foil treatment and name lettering. "
    "Collectors often dedicate pages to premium variants (Quarter Century "
    "Secret Rare, Collector's Rare, Extended Art), so the same card name may "
    "legitimately appear more than once across pages — list every physical "
    "card, and look closely at foil stamps and treatments to tell variants apart."
)

MEDIA = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "webp": "image/webp"}


def _prep_image(path):
    """Downscale to the API's effective max (2576px long edge) and re-encode —
    a 12MP phone photo uploads ~5-10x faster with no accuracy loss, and
    exif_transpose fixes sideways phone shots the model would misread."""
    try:
        import io
        from PIL import Image, ImageOps
    except ImportError:
        with open(path, "rb") as f:
            return f.read(), MEDIA.get(path.rsplit(".", 1)[-1].lower(), "image/jpeg")
    try:
        from pillow_heif import register_heif_opener  # iPhone HEIC photos
        register_heif_opener()
    except ImportError:
        pass
    im = ImageOps.exif_transpose(Image.open(path))
    if max(im.size) > 2576:
        im.thumbnail((2576, 2576), Image.LANCZOS)
    buf = io.BytesIO()
    im.convert("RGB").save(buf, "JPEG", quality=88)
    return buf.getvalue(), "image/jpeg"


def read_photo(client, path):
    raw, media = _prep_image(path)
    data = base64.standard_b64encode(raw).decode()
    resp = client.messages.create(
        model=MODEL,
        max_tokens=16000,
        output_config={"format": {"type": "json_schema", "schema": SCHEMA}},
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media, "data": data}},
                {"type": "text", "text": PROMPT},
            ],
        }],
    )
    if resp.stop_reason == "refusal":
        raise RuntimeError(f"model declined to process {path}")
    text = next(b.text for b in resp.content if b.type == "text")
    return json.loads(text)["cards"]


# Dragon Shield's rarity abbreviations → the names the databases use.
# Unknown abbreviations pass through as None — the dropdown handles it.
RARITY_ABBR = {"C": "Common", "R": "Rare", "SP": "Short Print",
               "SR": "Super Rare", "UR": "Ultra Rare", "ScR": "Secret Rare",
               "UtR": "Ultimate Rare", "GR": "Gold Rare", "GhR": "Ghost Rare",
               "PScR": "Prismatic Secret Rare", "PlScR": "Platinum Secret Rare",
               "QCScR": "Quarter Century Secret Rare", "StR": "Starlight Rare",
               "CR": "Collector's Rare", "PG": "Premium Gold Rare",
               "SFR": "Starfoil Rare", "MSR": "Mosaic Rare"}


def _ds_condition(s):
    """DragonShield "NearMint" → TCGPlayer "Near Mint"; unknowns → NM."""
    import re
    spaced = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", s.strip())
    known = ("Near Mint", "Lightly Played", "Moderately Played",
             "Heavily Played", "Damaged")
    return spaced if spaced in known else "Near Mint"


def parse_dragonshield_csv(path):
    """A Dragon Shield portfolio export → the same card rows a scan produces,
    but with exact set codes, quantities, and per-card conditions."""
    import csv
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        if not f.readline().lstrip('"').startswith("sep="):
            f.seek(0)  # some exports lack the (possibly quoted) separator hint
        for row in csv.DictReader(f):
            qty = int(row.get("Quantity") or 0)
            name = (row.get("Card Name") or "").strip()
            if not qty or not name:
                continue
            rows.append({
                "name": name,
                "set_code": (row.get("Card Number") or "").strip() or None,
                "qty": qty,
                "condition": _ds_condition(row.get("Condition") or ""),
                "rarity_guess": RARITY_ABBR.get((row.get("Rarity") or "").strip()),
            })
    return rows


def code_matches(guess, actual):
    """Partial-read set codes match on their readable characters only —
    '?' and other junk the model emits for glare are wildcards; a short
    guess is a prefix match."""
    g, a = guess.upper(), actual.upper()
    if len(g) > len(a):
        return False
    return all(gc == ac for gc, ac in zip(g, a) if gc.isalnum())


def _fetch(param, value):
    url = "https://db.ygoprodeck.com/api/v7/cardinfo.php?" + urllib.parse.urlencode({param: value})
    req = urllib.request.Request(url, headers={"User-Agent": "scan_cards/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            return json.load(r)["data"]
    except Exception:
        return None


def _find_card(name, snippet=None):
    """Exact match, then substring, then a misread-tolerant fallback: drop
    leading words until a substring search hits (glare mangles the start of a
    name more often than the whole thing), then pick the candidate whose name
    — and, when read, whose printed card text — best matches. The text box is
    matte, so the snippet is often the most reliable thing in the photo."""
    import difflib

    for param in ("name", "fname"):
        data = _fetch(param, name)
        if data:
            return data[0]

    def score(cand):
        s = difflib.SequenceMatcher(None, name.lower(), cand["name"].lower()).ratio()
        desc = cand.get("desc", "")
        if snippet and desc:
            t = difflib.SequenceMatcher(
                None, snippet.lower(), desc[:2 * len(snippet)].lower()).ratio()
            return (s + 2 * t) / 3  # the matte text outweighs the foil name
        return s

    words = name.split()
    for i in range(1, len(words)):  # len(words), not -1: two-word names get a try
        data = _fetch("fname", " ".join(words[i:]))
        if data:
            best = max(data, key=score)
            # printed-text corroboration earns a looser bar; a lone name
            # match must be close ("Final Fusion" ≠ a misread "Evil Fusion")
            return best if score(best) >= (0.6 if snippet else 0.8) else None
    return None


def db_lookup(name, set_code, snippet=None):
    """Return (canonical_name, candidates): the card's printings grouped by
    set code, narrowed by whatever part of the code was readable.
    One candidate = resolved; several = the user picks."""
    import html

    card = _find_card(name, snippet)
    if not card:
        return name, []
    groups = {}
    for s in card.get("card_sets", []):
        # ygoprodeck leaks HTML entities ("The Fallen &amp; The Virtuous")
        g = groups.setdefault(s["set_code"], {"set_code": s["set_code"],
                                              "set_name": html.unescape(s["set_name"]),
                                              "rarities": set()})
        g["rarities"].add(s["set_rarity"])
    cands = [dict(g, rarities=sorted(g["rarities"])) for _, g in sorted(groups.items())]
    cname = html.unescape(card["name"])
    if set_code:
        matched = [c for c in cands if code_matches(set_code, c["set_code"])]
        if matched:
            return cname, matched
    return cname, cands


def selftest():
    try:  # iPhone HEIC in → JPEG out (skipped if pillow_heif absent)
        import os
        import tempfile
        import pillow_heif
        from PIL import Image
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "t.heic")
            pillow_heif.from_pillow(Image.new("RGB", (100, 80), "red")).save(p)
            raw, media = _prep_image(p)
            assert media == "image/jpeg" and raw[:2] == b"\xff\xd8", media
    except ImportError:
        pass

    import os
    import tempfile
    with tempfile.TemporaryDirectory() as d:  # Dragon Shield CSV parsing
        p = os.path.join(d, "ds.csv")
        with open(p, "w", encoding="utf-8") as f:
            f.write('"sep=,"\n'
                    "Folder Name,Quantity,Trade Quantity,Card Name,Set Code,"
                    "Set Name,Card Number,Rarity,Condition,Printing,Language,"
                    "Price Bought,Date Bought,LOW,MID,MARKET\n"
                    "Ryan,2,0,Abaki,CDIP,Cyberdark Impact,CDIP-EN013,C,NearMint,"
                    "1st Edition,English,0.19,2026-05-16,0.01,0.22,0.16\n"
                    "Ryan,0,0,Skipped,X,X,X-001,C,NearMint,,,,,,,\n"
                    "Ryan,1,0,Ally of Justice Nullfier,HA02,Hidden Arsenal 2,"
                    "HA02-EN018,SR,LightlyPlayed,Unlimited,English,,,,,\n")
        rows = parse_dragonshield_csv(p)
        assert len(rows) == 2, rows  # qty-0 row dropped
        assert rows[0] == {"name": "Abaki", "set_code": "CDIP-EN013", "qty": 2,
                           "condition": "Near Mint", "rarity_guess": "Common"}, rows[0]
        assert rows[1]["condition"] == "Lightly Played"
        assert rows[1]["rarity_guess"] == "Super Rare"

    assert code_matches("SKE-EN0??", "SKE-EN020")
    assert code_matches("SKE-EN..?", "SKE-EN020")
    assert code_matches("SKE-EN", "SKE-EN020")          # prefix
    assert not code_matches("BLC-EN0??", "SKE-EN020")   # readable chars differ
    assert not code_matches("SKE-EN0201", "SKE-EN020")  # longer than actual

    name, cands = db_lookup("Fallen of Albaz", "RA01-EN021")
    assert name == "Fallen of Albaz", name
    assert len(cands) == 1 and cands[0]["set_name"] == "25th Anniversary Rarity Collection", cands
    assert "Secret Rare" in cands[0]["rarities"], cands

    _, cands = db_lookup("Fallen of Albaz", "RA01-EN0??")  # partial read resolves
    assert len(cands) == 1 and cands[0]["set_code"] == "RA01-EN021", cands

    _, cands = db_lookup("fallen of albaz", None)  # no code → all printings offered
    assert len(cands) > 5, len(cands)

    # misread leading word still finds the real card
    name, cands = db_lookup("Yusia and the Dark Dragon", None)
    assert name == "Ecclesia and the Dark Dragon", name
    assert cands, "expected printings for the fuzzy-matched card"
    print("selftest OK")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--selftest":
        selftest()
    else:
        sys.exit(__doc__)
