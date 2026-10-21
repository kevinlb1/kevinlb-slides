"""Replace fonts the decks ask for but nobody can install.

The decks reference Aptos (the theme's body and title face), Big Caslon Medium
and a few CSS-style font stacks. Aptos ships only with Microsoft 365, Big Caslon
only with macOS, and neither may be redistributed -- so on Office 2021 and on
the Linux build runner the text falls back and renders as garbled, overlapping
glyphs. The replacements are OFL-licensed, committed under fonts/, and
on both.

Roboto, which the title slides used for the lecture title, is retired in favour
of a heavier weight of the same family as the body text: Inter ExtraBold. Weight
lives in the family name there, so the bold flag is dropped from any run that
asked for Roboto -- otherwise PowerPoint would synthesise a fake bold on top of
the ExtraBold face.

Edits the packed XML by string substitution and rezips with each part's original
compression, changing nothing else. Deliberately does not use python-pptx: its
save() re-serialises every part it has parsed, which is far more churn than this
needs.

    python python/retheme_fonts.py --check     # report only
    python python/retheme_fonts.py             # rewrite the decks
"""

from __future__ import annotations

import argparse
import re
import shutil
import zipfile
from pathlib import Path

DECKS = Path("lectures")

# Order matters: the CSS stacks must be collapsed before the bare-name rules,
# so "Arial,Sans-Serif" does not become "Arial,sans-serif" and stay malformed.
REPLACEMENTS = [
    # A typeface attribute takes one font name; a comma-separated list is a CSS
    # font stack and is not valid OOXML.
    ("Courier New,monospace", "Courier New"),
    ("Wingdings,sans-serif", "Wingdings"),
    ("Arial,Sans-Serif", "Arial"),
    # Microsoft 365 only, and not redistributable.
    ("Aptos Display", "Inter"),
    ("Aptos", "Inter"),
    # macOS only, and not redistributable.
    ("Big Caslon Medium", "Libre Caslon Display"),
    ("Big Caslon", "Libre Caslon Display"),
    # Retired: the lecture title now uses a heavier weight of the body family.
    ("Roboto", "Inter ExtraBold"),
    # Came in with pasted Google Slides text; never installed anywhere here.
    ("Lato", "Inter"),
]

# Families whose weight is in their name. A bold flag on top of these makes
# PowerPoint fake a bold, so the flag comes off wherever they are set.
WEIGHT_NAMED = {"Inter ExtraBold"}


def rewrite(xml: str) -> tuple[str, dict[str, int]]:
    """Substitute inside typeface="..." attributes only."""
    counts: dict[str, int] = {}

    def swap(match: re.Match[str]) -> str:
        value = match.group(1)
        for old, new in REPLACEMENTS:
            if value == old:
                counts[f"{old} -> {new}"] = counts.get(f"{old} -> {new}", 0) + 1
                return f'typeface="{new}"'
        return match.group(0)

    xml = re.sub(r'typeface="([^"]*)"', swap, xml)

    def unbold(match: re.Match[str]) -> str:
        tag, attrs, body = match.group(1), match.group(2), match.group(3)
        if ' b="1"' in attrs and any(f'<a:latin typeface="{f}"' in body for f in WEIGHT_NAMED):
            counts["bold flag dropped"] = counts.get("bold flag dropped", 0) + 1
            attrs = attrs.replace(' b="1"', "")
        return f"<a:{tag}{attrs}>{body}</a:{tag}>"

    # Run properties are never nested, and a self-closing <a:rPr .../> has no
    # body to carry a typeface, so the non-greedy match is safe.
    xml = re.sub(r"<a:(rPr|defRPr|endParaRPr)([^>/]*)>(.*?)</a:\1>", unbold, xml, flags=re.S)
    return xml, counts


def process(deck: Path, check_only: bool) -> dict[str, int]:
    zin = zipfile.ZipFile(deck)
    totals: dict[str, int] = {}
    edited: dict[str, bytes] = {}

    for name in zin.namelist():
        if not name.endswith(".xml"):
            continue
        text = zin.read(name).decode("utf-8")
        new_text, counts = rewrite(text)
        if counts:
            for key, value in counts.items():
                totals[key] = totals.get(key, 0) + value
            edited[name] = new_text.encode("utf-8")

    if edited and not check_only:
        tmp = deck.with_suffix(".pptx.tmp")
        with zipfile.ZipFile(tmp, "w") as zout:
            for item in zin.infolist():
                data = edited.get(item.filename) or zin.read(item.filename)
                info = zipfile.ZipInfo(item.filename, date_time=item.date_time)
                info.compress_type = item.compress_type
                info.external_attr = item.external_attr
                zout.writestr(info, data)
        zin.close()
        shutil.move(tmp, deck)
    else:
        zin.close()

    return totals


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="report without writing")
    parser.add_argument("--skip", action="append", default=[], metavar="DECK",
                        help="leave this deck untouched (repeatable); the run is idempotent, so do it later")
    args = parser.parse_args(argv)

    grand: dict[str, int] = {}
    for deck in sorted(DECKS.glob("*.pptx")):
        if deck.name in args.skip:
            print(f"{deck.name}: skipped")
            continue
        totals = process(deck, args.check)
        if totals:
            print(f"{deck.name}:")
            for key, value in sorted(totals.items()):
                print(f"    {value:4d}  {key}")
                grand[key] = grand.get(key, 0) + value

    if not grand:
        print("No font references needed changing.")
        return 0

    print("\ntotal:")
    for key, value in sorted(grand.items()):
        print(f"    {value:4d}  {key}")
    if args.check:
        print("\n(--check: nothing written)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
