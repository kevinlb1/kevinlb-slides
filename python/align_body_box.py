"""Put a text box moved around a picture back on the template's margins.

A body placeholder makes room for a picture by giving up one edge. Every other
edge should stay where the text-only template has it, so the text block does
not drift inwards, downwards or in from the far side, and the first line starts
where it starts on the slides that have no picture at all. Which edge gives way
follows from where the picture is:

- **beside the text** (a picture to the right): the box keeps the layout's
  top-left and bottom-left corners and only its right edge moves;
- **below the text** (a picture under the bullets): the box keeps the layout's
  two top corners, so the first line starts where it always does, and only its
  bottom edge moves;
- **above the text** (a table over the bullets): the box keeps the layout's
  width and its bottom edge, and only its top edge moves.

This reads each slide's body placeholder (or the text box PowerPoint still
calls "Content Placeholder N" after it was detached from the layout), works out
which case it is from the other shapes on the slide, and rewrites its position
accordingly. A box that ends up exactly where the layout puts it loses its
explicit position altogether, so it follows the template again. A box with an
explicit position on a slide with nothing else on it is reported, not moved:
the rules say it should not have one, but why it has one is a judgment call.

    python python/align_body_box.py --check "decks/Lecture04.pptx"
    python python/align_body_box.py "decks/Lecture04.pptx" --slides 9 11 15

Edits the packed XML directly and rezips with each part's original compression
(see retheme_fonts.py for why not a python-pptx round trip).
"""

from __future__ import annotations

import argparse
import datetime as dt
import shutil
import sys
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from add_title_picture_layout import Package  # noqa: E402
from deck_text import NS, parse, placeholder, resolve, serialize, shape_name, slide_parts  # noqa: E402

EMU = 914400
RT_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
TOLERANCE = int(0.05 * EMU)      # a picture may overlap the text box by this much and still count as beside it


def box(sp: etree._Element) -> tuple[int, int, int, int] | None:
    """The shape's own position, or None when it takes the layout's."""
    xfrm = sp.find("p:spPr/a:xfrm", NS)
    if xfrm is None:
        xfrm = sp.find("p:xfrm", NS)          # a table or a chart carries its own
    if xfrm is None:
        return None
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    return int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))


def inches(b: tuple[int, int, int, int]) -> str:
    return "{:.2f} {:.2f} {:.2f} x {:.2f}".format(*(v / EMU for v in b))


def is_body(sp: etree._Element) -> bool:
    ph = placeholder(sp)
    if ph is not None:
        return ph.get("type") in (None, "body", "obj")
    return shape_name(sp).startswith("Content Placeholder") \
        and sp.find("p:txBody/a:lstStyle/a:lvl1pPr", NS) is not None


def layout_box(pkg: Package, slide: str, sp: etree._Element) -> tuple[int, int, int, int] | None:
    """Where the slide's layout puts this placeholder."""
    rel = pkg.relationship(slide, RT_LAYOUT)
    if rel is None:
        return None
    root = parse(pkg, resolve(slide, rel[1]))
    ph = placeholder(sp)
    idx = ph.get("idx") if ph is not None else None
    fallback = None
    for shape in root.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}sp"):
        lph = placeholder(shape)
        if lph is None or lph.get("type") not in (None, "body", "obj"):
            continue
        if idx is not None and lph.get("idx") == idx:
            return box(shape)
        if fallback is None:
            fallback = box(shape)
    return fallback


def others(tree: etree._Element, sp: etree._Element) -> list[tuple[int, int, int, int]]:
    """Every other shape on the slide that carries a position of its own."""
    out = []
    for other in tree:
        if other is sp or etree.QName(other).localname in ("nvGrpSpPr", "grpSpPr"):
            continue
        b = box(other)
        if b is not None:
            out.append(b)
    return out


def kept_above(wanted: tuple[int, int, int, int], current: tuple[int, int, int, int],
               rest: list[tuple[int, int, int, int]]) -> tuple[int, int, int, int]:
    """Keep a bottom edge that was stopping short of a picture underneath.

    The template's bottom is the right one where the space below the text is
    empty, and reaching it is most of the point of this. But a box can also
    have been pulled up to leave room for something directly beneath the
    bullets -- the pixel grids on Lecture 5 slide 6 -- and pushing its bottom
    back down through those puts the last line on top of them. So the bottom
    only descends into space nothing else is using.
    """
    bottom, was = wanted[1] + wanted[3], current[1] + current[3]
    if bottom <= was:
        return wanted
    left, right = wanted[0], wanted[0] + wanted[2]
    for b in rest:
        if b[0] + TOLERANCE >= right or b[0] + b[2] - TOLERANCE <= left:
            continue                      # beside the box, not under it
        if was - TOLERANCE <= b[1] <= bottom:
            return (wanted[0], wanted[1], wanted[2], was - wanted[1])
    return wanted


def align(pkg: Package, indices: set[int]) -> list[str]:
    report = []
    for number, part in enumerate(slide_parts(pkg), start=1):
        if indices and number not in indices:
            continue
        root = parse(pkg, part)
        tree = root.find("p:cSld/p:spTree", NS)
        changed = False
        for sp in tree:
            if etree.QName(sp).localname != "sp" or not is_body(sp):
                continue
            current = box(sp)
            if current is None:
                continue
            frame = layout_box(pkg, part, sp)
            if frame is None:
                report.append(f"slide {number}: the layout does not place this placeholder; skipped")
                continue
            right = current[0] + current[2]
            rest = others(tree, sp)
            if any(b[0] + TOLERANCE >= right for b in rest):
                # a picture beside the text: the right edge is the one that gave way
                wanted = (frame[0], frame[1], right - frame[0], frame[3])
            elif rest:
                # a picture under the bullets leaves the top corners alone and
                # takes the bottom; something over them pushes the top down
                above = [b for b in rest if b[1] + b[3] <= current[1] + TOLERANCE
                         and b[0] < frame[0] + frame[2] and b[0] + b[2] > frame[0]]
                top = current[1] if above else frame[1]
                wanted = (frame[0], top, frame[2], frame[1] + frame[3] - top)
            else:
                report.append(f"slide {number}: {shape_name(sp)} is placed by hand with nothing "
                              f"else on the slide; left as it is ({inches(current)})")
                continue
            wanted = kept_above(wanted, current, rest)
            if wanted == current:
                continue
            xfrm = sp.find("p:spPr/a:xfrm", NS)
            if wanted == frame:                       # back on the layout's own box
                xfrm.getparent().remove(xfrm)
                note = "follows the layout again"
            else:
                xfrm.find("a:off", NS).set("x", str(wanted[0]))
                xfrm.find("a:off", NS).set("y", str(wanted[1]))
                xfrm.find("a:ext", NS).set("cx", str(wanted[2]))
                xfrm.find("a:ext", NS).set("cy", str(wanted[3]))
                note = inches(wanted)
            report.append(f"slide {number}: {shape_name(sp)} {inches(current)} -> {note}")
            changed = True
        if changed:
            pkg.write_bytes(part, serialize(root))
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("deck", type=Path)
    parser.add_argument("--slides", type=int, nargs="*", default=[], help="only these slide numbers")
    parser.add_argument("--check", action="store_true", help="report what would change and write nothing")
    parser.add_argument("-o", "--output", type=Path, help="write the result here instead of in place")
    parser.add_argument("--no-backup", action="store_true", help="in place without first copying the deck into backups/")
    args = parser.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")

    target = args.deck
    if args.output and not args.check:
        shutil.copyfile(args.deck, args.output)
        target = args.output
    elif not args.check:
        if (target.parent / f"~${target.name}").exists():
            print(f"{target.name}: open in PowerPoint; close it first")
            return 1
        if not args.no_backup:
            backups = target.parent / "backups"
            backups.mkdir(exist_ok=True)
            copy = backups / f"{target.stem} - before align_body_box {dt.datetime.now():%Y-%m-%d %H%M}{target.suffix}"
            if not copy.exists():
                shutil.copy2(target, copy)
                print(f"backed up to {copy}")
    pkg = Package(target)
    report = align(pkg, set(args.slides))
    if args.check:
        pkg.zin.close()
    else:
        pkg.save()
    print(f"{target.name}: " + ("\n  " + "\n  ".join(report) if report else "nothing to align"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
