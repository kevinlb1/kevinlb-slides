"""Tidy the emphasis and whitespace of a deck's slide text to STYLE.md.

Four mechanical fixes, all safe to re-run:

- Alert weights and tints. An alert is a run coloured Accent 1 or Accent 2
  through a theme reference, as the AlertText add-in writes it. Its weight
  must be the look's weight for the outline level: alert 1 is the plain family
  with the bold flag on level 1, SemiBold on level 2 and Medium below; alert 2
  is SemiBold, Medium, then the plain family. Its colour is the accent at full
  strength on level 1, 15 % lighter on level 2, 30 % on level 3 and 50 % below.
  Text boxes, tables and titles count as level 1. A run at any other weight or
  tint is put right, bold flag included. Runs with an explicit size inside an
  outline placeholder are left alone: they were made smaller on purpose and
  follow the add-in's Smaller ladder, not this table.
- Trailing whitespace. Spaces, tabs and no-break spaces at the end of a
  paragraph are removed from every text body on the slides (titles and tables
  included, notes excluded), and whitespace-only runs at the end go with them.
- Terminal periods. A bullet never ends in a period, so the last one goes;
  periods inside the bullet stay, as do question and exclamation marks and an
  ellipsis. Titles and table cells are not bullets and are left alone.
- Maths. A symbol set as a Cambria Math run is turned into the equation zone
  the rule asks for, and the shape that holds it into the
  mc:AlternateContent pair PowerPoint uses (see deck_text.py). The weights
  above are not applied to a symbol: it wears Cambria Math italic by design.
- Notes. A slide's notes keep one record of the text a human wrote before the
  editing started, not a log of every pass, so a second and later block is
  removed.

    python python/tidy_slide_text.py --check decks/*.pptx   # report only
    python python/tidy_slide_text.py "decks/Lecture01.pptx"               # fix in place, after a backup
    python python/tidy_slide_text.py --only alerts "decks/Lecture01.pptx" # just the emphasis

--only picks which of the four run. A lecture that has already been given can
take the alert pass, which touches nothing but a run's weight and colour,
without any of the three that edit characters.

In place, the deck is first copied into a backups/ folder beside it, and a
deck that is open in PowerPoint is skipped. Edits the packed XML directly and
rezips with each part's original compression (see retheme_fonts.py for why
not a python-pptx round trip).
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import shutil
import sys
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from add_title_picture_layout import Package  # noqa: E402
from deck_text import (A14, MATH, NS, PANOSE, Master, alert_accent, in_fallback, is_math,
                       look_props, wrap_for_math,
                       notes_body, notes_part, para_level, parse, records,  # noqa: E402
                       trim_extra_records,
                       placeholder, plain_text, q, serialize, slide_parts)

OUTLINE_TYPES = {None, "body", "obj", "subTitle"}     # where the paragraph's level is the look's level
AFTER_LATIN = {"ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"}


def snippet(p: etree._Element) -> str:
    text = " ".join(plain_text(p).split())
    return text if len(text) <= 60 else text[:57] + "..."


def run_context(run: etree._Element) -> tuple[int, bool]:
    """The outline level that decides a run's look, and whether the run sits in
    an outline placeholder: the paragraph's level there, level 1 in a title,
    text box or table cell."""
    p = run.getparent()
    sp = p
    while sp is not None and etree.QName(sp).localname != "sp":
        sp = sp.getparent()
    if sp is None:                                 # a table cell
        return 1, False
    ph = placeholder(sp)
    if ph is None:
        # a placeholder detached from its layout keeps its name and list style
        # and is edited as a body (see deck_text.SlideView)
        name = sp.find(".//p:cNvPr", NS)
        detached = (name is not None and name.get("name", "").startswith("Content Placeholder")
                    and sp.find("p:txBody/a:lstStyle/a:lvl1pPr", NS) is not None)
        return (para_level(p), True) if detached else (1, False)
    if ph.get("type") not in OUTLINE_TYPES:
        return 1, False
    return para_level(p), True


def math_runs_to_zones(root: etree._Element) -> list[str]:
    """Turn a symbol set as a Cambria Math run into an equation zone.

    The rule is that maths is an equation object, which is what PowerPoint's
    editor inserts and what deck_text.py writes now; a run in the same font
    looks the same but is not one. Runs like that were written here on
    17 September 2026, before the rule was settled."""
    changes = []
    for paragraph in root.iter(q("a:p")):
        if in_fallback(paragraph):
            continue
        children = list(paragraph)
        for run in children:
            if run.tag != q("a:r") or not is_math(run.find("a:rPr", NS)):
                continue
            if run.getparent() is None:
                continue                                  # already taken as a subscript
            following = children[children.index(run) + 1:]
            sub = None
            if following and following[0].tag == q("a:r") \
                    and is_math(following[0].find("a:rPr", NS)) \
                    and following[0].find("a:rPr", NS).get("baseline"):
                sub = following[0]
            text = (run.find("a:t", NS).text or "")
            zone = math_zone(run, sub)
            paragraph.insert(paragraph.index(run), zone)
            paragraph.remove(run)
            if sub is not None:
                text += "_" + (sub.find("a:t", NS).text or "")
                paragraph.remove(sub)
            changes.append(f'"{text}" set as an equation')
        if changes:
            wrap_for_math(paragraph.getparent().getparent())
    return changes


def math_zone(run: etree._Element, sub: etree._Element | None) -> etree._Element:
    zone = etree.Element(q("a14:m"), nsmap={"a14": A14})
    omath = etree.SubElement(zone, q("m:oMath"), nsmap={"m": MATH})
    if sub is None:
        omath.append(math_r(run))
        return zone
    ssub = etree.SubElement(omath, q("m:sSub"))
    ctrl = etree.SubElement(etree.SubElement(ssub, q("m:sSubPr")), q("m:ctrlPr"))
    ctrl.append(copy.deepcopy(run.find("a:rPr", NS)))
    etree.SubElement(ssub, q("m:e")).append(math_r(run))
    etree.SubElement(ssub, q("m:sub")).append(math_r(sub))
    return zone


def math_r(run: etree._Element) -> etree._Element:
    r = etree.Element(q("m:r"))
    rpr = copy.deepcopy(run.find("a:rPr", NS))
    rpr.attrib.pop("baseline", None)
    r.append(rpr)
    etree.SubElement(r, q("m:t")).text = run.find("a:t", NS).text or ""
    return r


def fix_alert_weights(root: etree._Element, master: Master) -> list[str]:
    changes = []
    for rpr in root.iter(q("a:rPr")):
        run = rpr.getparent()
        level, outline = run_context(run)
        accent = alert_accent(rpr, level, master)
        if accent is None:
            continue
        look = "A1" if accent == "accent1" else "A2"
        if outline and rpr.get("sz") is not None:
            continue          # made smaller on purpose; follows the Smaller ladder, not this table
        wanted_colour = master.alert(accent, level)
        fill = rpr.find("a:solidFill", NS)
        srgb = fill.find("a:srgbClr", NS) if fill is not None else None
        current_colour = (srgb.get("val") or "").upper() if srgb is not None else "a theme colour"
        text = plain_text(run.getparent()) if run.getparent() is not None else ""
        word = ((run.find("a:t", NS) if run.tag == q("a:r") else run.find("m:t", NS)) is not None
                and (run.find("a:t", NS) if run.tag == q("a:r") else run.find("m:t", NS)).text or "").strip() or "(space)"
        where = f'(alert {look[1]}, level {level}) in "{" ".join(text.split())[:50]}"'
        if current_colour != wanted_colour:
            for child in list(fill):
                fill.remove(child)
            etree.SubElement(fill, q("a:srgbClr"), val=wanted_colour)
            changes.append(f'"{word}" colour {current_colour} -> {wanted_colour} {where}')
        if run.tag != q("a:r") or is_math(rpr):
            # a symbol set as maths wears Cambria Math italic by design; only
            # its colour follows the level
            continue
        family, bold, _ = look_props(look, level, master.base)
        latin = rpr.find("a:latin", NS)
        current = latin.get("typeface") if latin is not None else master.family.get(level, master.base)
        if current.startswith("+"):
            current = master.family.get(level, master.base)
        is_bold = rpr.get("b") == "1"
        if (current, is_bold) == (family, bold):
            continue
        if latin is None:
            latin = etree.Element(q("a:latin"))
            for child in rpr:
                if etree.QName(child).localname in AFTER_LATIN:
                    child.addprevious(latin)
                    break
            else:
                rpr.append(latin)
        latin.set("typeface", family)
        if PANOSE.get(family):
            latin.set("panose", PANOSE[family])
        elif "panose" in latin.attrib:
            del latin.attrib["panose"]
        latin.set("pitchFamily", "2")
        latin.set("charset", "0")
        if bold:
            rpr.set("b", "1")
        elif "b" in rpr.attrib:
            del rpr.attrib["b"]
        changes.append(f'"{word}" {current}{" bold" if is_bold else ""} -> '
                       f'{family}{" bold" if bold else ""} {where}')
    return changes


def describe_tint(tint: tuple[int, int | None]) -> str:
    return f"lighter {tint[1] // 1000}%" if tint[1] else "full strength"


def strip_trailing_whitespace(root: etree._Element) -> list[str]:
    changes = []
    for p in root.iter(q("a:p")):
        trimmed = False
        while True:
            content = [c for c in p if etree.QName(c).localname in ("r", "br", "fld", "m")]
            if not content:
                break
            last = content[-1]
            if etree.QName(last).localname != "r":
                break                      # a line break, a field, or a maths zone last
            t = last.find("a:t", NS)
            text = t.text or ""
            if not text.strip():
                p.remove(last)
                trimmed = True
                continue
            if text != text.rstrip():
                t.text = text.rstrip()
                trimmed = True
            break
        if trimmed:
            changes.append(f'trailing whitespace removed from "{snippet(p)}"')
    return changes


def strip_terminal_periods(root: etree._Element) -> list[str]:
    """A bullet never ends in a period (STYLE.md rule 1).

    Only the period at the very end goes. The ones inside the bullet stay, and
    so do a question mark, an exclamation mark and an ellipsis, whether it is
    written as one character or as three dots. Titles and table cells are not
    bullets and are left alone.
    """
    changes = []
    for sp in root.iter(q("p:sp")):
        ph = placeholder(sp)
        if ph is not None and ph.get("type") not in OUTLINE_TYPES:
            continue
        for p in sp.iter(q("a:p")):
            content = [c for c in p if etree.QName(c).localname in ("r", "br", "fld", "m")]
            if not content or etree.QName(content[-1]).localname != "r":
                continue
            t = content[-1].find("a:t", NS)
            text = t.text if t is not None else None
            if not text or not text.endswith(".") or text.endswith(".."):
                continue
            t.text = text[:-1]
            if not t.text:
                p.remove(content[-1])      # the period had a run to itself
            changes.append(f'terminal period removed from "{snippet(p)}"')
    return changes


PASSES = ("maths", "alerts", "whitespace", "periods", "notes")


def tidy(deck: Path, check: bool, passes: tuple[str, ...] = PASSES) -> list[str]:
    pkg = Package(deck)
    master = Master(pkg)
    report = []
    for index, part in enumerate(slide_parts(pkg), 1):
        root = parse(pkg, part)
        changes = []
        if "maths" in passes:
            changes += math_runs_to_zones(root)
        if "alerts" in passes:
            changes += fix_alert_weights(root, master)
        if "whitespace" in passes:
            changes += strip_trailing_whitespace(root)
        if "periods" in passes:
            changes += strip_terminal_periods(root)
        if changes:
            report.extend(f"slide {index}: {change}" for change in dict.fromkeys(changes))
            if not check:
                pkg.write_bytes(part, serialize(root))
        if "notes" not in passes:
            continue
        notes = notes_part(pkg, part)
        if notes is None:
            continue
        notes_root = parse(pkg, notes)
        body = notes_body(notes_root)
        if body is None or len(records(body)) < 2:
            continue
        report.append(f"slide {index}: notes cut back to the first original-text record")
        if not check:
            trim_extra_records(body)
            pkg.write_bytes(notes, serialize(notes_root))
    if check:
        pkg.zin.close()
    else:
        pkg.save()
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("decks", nargs="+", type=Path)
    parser.add_argument("--check", action="store_true", help="report what would change and write nothing")
    parser.add_argument("--no-backup", action="store_true", help="edit in place without first copying the deck into backups/")
    parser.add_argument("--only", action="append", choices=PASSES, metavar="PASS",
                        help=f"run only this pass ({', '.join(PASSES)}); repeatable, all of them by default")
    args = parser.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")
    for deck in args.decks:
        if not args.check and (deck.parent / f"~${deck.name}").exists():
            print(f"{deck.name}: open in PowerPoint; skipped")
            continue
        if not args.check and not args.no_backup:
            backups = deck.parent / "backups"
            backups.mkdir(exist_ok=True)
            copy = backups / f"{deck.stem} - before tidy {dt.datetime.now():%Y-%m-%d %H%M}{deck.suffix}"
            if not copy.exists():
                shutil.copy2(deck, copy)
        report = tidy(deck, args.check, tuple(args.only) if args.only else PASSES)
        verb = "would change" if args.check else "changed"
        print(f"{deck.name}: {len(report)} {verb}" if report else f"{deck.name}: nothing to tidy")
        for line in report:
            print(f"  {line}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
