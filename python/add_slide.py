"""Add a slide to a deck, copied from one the deck already has.

`deck_text.py` rewrites the text of slides that exist; this makes a new one,
for the times the house style asks for a slide that is not there (a lecture's
conclusion, say). It copies a model slide already in the deck -- its layout,
its footer and slide-number placeholders, its language -- keeps only the title
and the first content placeholder, drops everything else that was on it
(pictures, tables, loose boxes, explicit placeholder positions), and writes
the title and bullets given to it. The bullets are `deck_text.py`'s markup, so
the emphasis looks and the levels are the same:

    python python/add_slide.py "decks/Lecture04.pptx" --after 19 \
        --title Conclusion --body conclusion.txt

`--after 0` puts the slide first and leaving `--after` out puts it last. The
model slide is the last one with a title, one content placeholder and nothing
else, unless `--like N` names another. A new slide has no previous text, so
nothing is written to its notes; editing in place first copies the deck into
`backups/` beside it and refuses a deck that is open in PowerPoint, as
`deck_text.py` does.

Edits the packed XML directly and rezips with each part's original compression
(see retheme_fonts.py for why not a python-pptx round trip).
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import shutil
import sys
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from add_title_picture_layout import Package  # noqa: E402
from deck_text import (  # noqa: E402
    NS, RT_SLIDE, Master, SlideView, lang_of, parse, parse_spec, placeholder, q,
    rewrite_body, serialize, set_title, slide_parts,
)

RT_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
RELS_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{RT_LAYOUT}" Target="{{layout}}"/>'
    "</Relationships>"
)
CREATION_ID = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"
SLIDE_ID = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"


def is_plain(pkg: Package, part: str, master: Master) -> bool:
    """A title, one content placeholder, and nothing else to copy."""
    view = SlideView(pkg, part, master)
    return view.title_body is not None and len(view.bodies) == 1 and not view.readonly


def pick_model(pkg: Package, parts: list[str], master: Master) -> str:
    for part in reversed(parts):
        if is_plain(pkg, part, master):
            return part
    raise ValueError("no slide in the deck has just a title and one content placeholder; "
                     "name one to copy with --like")


def clone(pkg: Package, model: str, title: str, paragraphs: list[tuple[int, str]],
          master: Master) -> etree._Element:
    """The model slide with its own content stripped and the new text written."""
    root = parse(pkg, model)
    tree = root.find("p:cSld/p:spTree", NS)
    title_body = body = None
    for sp in list(tree):
        tag = etree.QName(sp).localname
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        ph = placeholder(sp) if tag == "sp" else None
        kind = ph.get("type") if ph is not None else None
        if ph is None or kind not in (None, "title", "ctrTitle", "body", "obj", "ftr", "sldNum", "dt"):
            tree.remove(sp)                       # a picture, a table, a loose box
            continue
        if kind in ("title", "ctrTitle"):
            title_body = sp.find("p:txBody", NS)
        elif kind in (None, "body", "obj"):
            if body is not None:
                tree.remove(sp)
                continue
            body = sp.find("p:txBody", NS)
        else:
            continue                              # the footer and the slide number, as they are
        xfrm = sp.find("p:spPr/a:xfrm", NS)        # let the layout place it again
        if xfrm is not None:
            xfrm.getparent().remove(xfrm)
    if title_body is None or body is None:
        raise ValueError(f"{model} has no title or no content placeholder to copy")
    for ext in [e for e in root.iter(q("a:ext"), q("p:ext"))   # PowerPoint's own shape and slide ids
                if e.get("uri") in (CREATION_ID, SLIDE_ID)]:
        parent = ext.getparent()
        parent.remove(ext)
        if not len(parent):
            parent.getparent().remove(parent)
    set_title(title_body, title, lang_of(title_body, "en-US"))
    rewrite_body(body, paragraphs, master, lang_of(body, "en-US"))
    return root


def next_part(pkg: Package) -> str:
    used = [int(m.group(1)) for n in [*pkg.names, *pkg.added]
            if (m := re.match(r"ppt/slides/slide(\d+)\.xml$", n))]
    return f"ppt/slides/slide{max(used, default=0) + 1}.xml"


def register(pkg: Package, part: str, rid: str, after: int) -> None:
    """Put the new slide into the slide list, after the given position."""
    presentation = pkg.read("ppt/presentation.xml")
    ids = [int(m) for m in re.findall(r'<p:sldId id="(\d+)"', presentation)]
    entry = f'<p:sldId id="{max(ids, default=255) + 1}" r:id="{rid}"/>'
    matches = list(re.finditer(r'<p:sldId [^>]*/>', presentation))
    if after <= 0:
        at = presentation.index("<p:sldIdLst>") + len("<p:sldIdLst>")
    else:
        at = matches[after - 1].end()
    pkg.write("ppt/presentation.xml", presentation[:at] + entry + presentation[at:])
    pkg.add_override(part, CT_SLIDE)


def update_properties(pkg: Package, title: str, after: int) -> None:
    """Keep docProps/app.xml's slide count and title list the right length."""
    if "docProps/app.xml" not in pkg.names:
        return
    app = pkg.read("docProps/app.xml")
    m = re.search(r"<Slides>(\d+)</Slides>", app)
    if m is None:
        return
    slides = int(m.group(1))
    app = app[:m.start()] + f"<Slides>{slides + 1}</Slides>" + app[m.end():]
    m = re.search(r"(<vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)(</vt:i4>)", app)
    if m is None:
        return
    app = app[:m.start()] + m.group(1) + str(int(m.group(2)) + 1) + m.group(3) + app[m.end():]
    m = re.search(r'(<TitlesOfParts><vt:vector size=")(\d+)(")', app)
    parts = int(m.group(2))
    app = app[:m.start()] + m.group(1) + str(parts + 1) + m.group(3) + app[m.end():]
    entries = list(re.finditer(r"<vt:lpstr>[^<]*</vt:lpstr>", app[app.index("<TitlesOfParts>"):]))
    first = app.index("<TitlesOfParts>")
    at = first + entries[parts - slides + after - 1].end()   # the fonts and theme come first
    escaped = title.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    pkg.write("docProps/app.xml", app[:at] + f"<vt:lpstr>{escaped}</vt:lpstr>" + app[at:])


def add(pkg: Package, title: str, paragraphs: list[tuple[int, str]],
        after: int | None, like: int | None) -> str:
    master = Master(pkg)
    parts = slide_parts(pkg)
    position = len(parts) if after is None else after
    if not 0 <= position <= len(parts):
        raise ValueError(f"--after {after}: the deck has {len(parts)} slides")
    if like is not None:
        if not 1 <= like <= len(parts):
            raise ValueError(f"--like {like}: the deck has {len(parts)} slides")
        model = parts[like - 1]
    else:
        model = pick_model(pkg, parts, master)
    root = clone(pkg, model, title, paragraphs, master)
    part = next_part(pkg)
    layout = pkg.relationship(model, RT_LAYOUT)
    if layout is None:
        raise ValueError(f"{model} has no layout relationship")
    pkg.write_bytes(part, serialize(root))
    pkg.write(pkg.rels_name(part), RELS_TEMPLATE.format(layout=layout[1]))
    rid = pkg.add_relationship("ppt/presentation.xml", RT_SLIDE, f"slides/{part.rsplit('/', 1)[1]}")
    register(pkg, part, rid, position)
    update_properties(pkg, title, position)
    return (f'slide {position + 1}: "{title}", {len(paragraphs)} paragraphs, '
            f"copied from slide {parts.index(model) + 1}, written as {part}")


def read_body(source: Path | None) -> list[tuple[int, str]]:
    if source is None:
        return []
    text = sys.stdin.read() if str(source) == "-" else source.read_text(encoding="utf-8")
    if not text.lstrip().startswith("## Slide"):
        text = "## Slide 1: \n" + text
    spec = parse_spec(text)
    slide = spec[next(iter(spec))]
    return slide.bodies.get(None, [])


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("deck", type=Path)
    parser.add_argument("--title", required=True, help="the new slide's title")
    parser.add_argument("--body", type=Path, help="a file of '- ' lines in deck_text.py's markup ('-' for stdin)")
    parser.add_argument("--after", type=int, help="the slide to put it after (0 first, the end by default)")
    parser.add_argument("--like", type=int, help="the slide to copy (the last plain one by default)")
    parser.add_argument("-o", "--output", type=Path, help="write the result here instead of in place")
    parser.add_argument("--no-backup", action="store_true", help="in place without first copying the deck into backups/")
    args = parser.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")

    paragraphs = read_body(args.body)
    target = args.deck
    if args.output:
        if args.output.resolve() == args.deck.resolve():
            parser.error("-o must not be the deck itself; leave it out to edit in place")
        shutil.copyfile(args.deck, args.output)
        target = args.output
    else:
        if (target.parent / f"~${target.name}").exists():
            print(f"{target.name}: open in PowerPoint; close it first")
            return 1
        if not args.no_backup:
            backups = target.parent / "backups"
            backups.mkdir(exist_ok=True)
            copy = backups / f"{target.stem} - before add_slide {dt.datetime.now():%Y-%m-%d %H%M}{target.suffix}"
            if not copy.exists():
                shutil.copy2(target, copy)
                print(f"backed up to {copy}")
    pkg = Package(target)
    report = add(pkg, args.title, paragraphs, args.after, args.like)
    pkg.save()
    print(report)
    return 0


if __name__ == "__main__":
    sys.exit(main())
