"""Read Yu-Gi-Oh card names/sets/rarities from binder photos.

Usage:
    python scan_cards.py photo1.jpg [photo2.jpg ...]   # needs ANTHROPIC_API_KEY
    python scan_cards.py --selftest                    # offline check of the DB lookup

Claude vision reads each photo (handles foils/glare that break normal OCR),
then each card is cross-referenced against the free YGOPRODeck database to
resolve the exact set name and printed rarity from the set code.
"""
import base64
import csv
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


def read_photo(client, path):
    with open(path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode()
    media = MEDIA.get(path.rsplit(".", 1)[-1].lower(), "image/jpeg")
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
    for i in range(1, len(words) - 1):
        data = _fetch("fname", " ".join(words[i:]))
        if data:
            best = max(data, key=score)
            return best if score(best) >= 0.6 else None
    return None


def db_lookup(name, set_code, snippet=None):
    """Return (canonical_name, candidates): the card's printings grouped by
    set code, narrowed by whatever part of the code was readable.
    One candidate = resolved; several = the user picks."""
    card = _find_card(name, snippet)
    if not card:
        return name, []
    groups = {}
    for s in card.get("card_sets", []):
        g = groups.setdefault(s["set_code"], {"set_code": s["set_code"],
                                              "set_name": s["set_name"], "rarities": set()})
        g["rarities"].add(s["set_rarity"])
    cands = [dict(g, rarities=sorted(g["rarities"])) for _, g in sorted(groups.items())]
    if set_code:
        matched = [c for c in cands if code_matches(set_code, c["set_code"])]
        if matched:
            return card["name"], matched
    return card["name"], cands


def main(paths):
    import anthropic
    client = anthropic.Anthropic()
    writer = csv.writer(sys.stdout)
    writer.writerow(["photo", "name", "set_code", "set_name", "rarity", "rarity_guess"])
    for path in paths:
        for c in read_photo(client, path):
            name, cands = db_lookup(c["name"], c["set_code"])
            if len(cands) == 1:
                code, set_name, rarities = cands[0]["set_code"], cands[0]["set_name"], cands[0]["rarities"]
            else:
                code = c["set_code"] or ""
                set_name = f"ambiguous ({len(cands)} printings)" if cands else ""
                rarities = []
            writer.writerow([path, name, code, set_name,
                             " / ".join(rarities), c["rarity_guess"] or ""])


def selftest():
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
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    if sys.argv[1] == "--selftest":
        selftest()
    else:
        main(sys.argv[1:])
