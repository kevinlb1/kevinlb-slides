"""Slide text as an editable text file: dump a deck's bullets, edit them, apply them back.

The lecture decks follow the house style in STYLE.md, which is about
words and emphasis, so this tool exposes exactly that and nothing else:

    python python/deck_text.py dump  "decks/Lecture03.pptx" > lecture03.txt
    ... edit lecture03.txt ...
    python python/deck_text.py apply "decks/Lecture03.pptx" lecture03.txt -o candidate.pptx
    python python/deck_text.py apply "decks/Lecture03.pptx" lecture03.txt   # in place, after a backup

The file has one block per slide:

    ## Slide 5: Applications of search
    - Search can be applied in a few different scenarios:
        - Find [[A1:any solution]]
            - you just want some way to get from Vancouver to Antarctica
    -
    > #activity-QR

A "- " line is a paragraph of the slide's content placeholder, with four more
leading spaces per outline level; a bare "-" is an empty paragraph; "<br>" is
a line break inside a paragraph. Emphasis is written [[A1:text]] (alert 1),
[[A2:text]] (alert 2) or [[B:text]] (Bolder), the three looks of the AlertText
add-in; everything else takes the level's default look. [[M:text]] is a symbol set
as maths -- an equation zone, as PowerPoint's editor inserts -- with [[M:x_1]]
for a subscript; it may sit inside an alert ([[A1:[[M:k]]-nearest neighbours]])
and then wears the alert's colour. A shape holding one is wrapped in
mc:AlternateContent, with a plain-text copy in the fallback for readers that
do not do OOXML maths: LibreOffice makes the published PDFs and is one of
them (STYLE.md rule 12). Lines starting with
">" are the notes (shown so the build's tags are visible, never applied),
"[read-only ...]" blocks with "|" lines are tables and text boxes, and
"[body: NAME]" introduces each content placeholder on a slide that has more
than one. Lines starting with "#" (other than the slide headers) are comments.

A hyperlinked run is [[URL:text]]; apply keeps the link the same words carried
before (or the body's only link if the words changed), in the plain look.

A dump also flags emphasis that is none of the three looks: [[BOLD:text]] for
PowerPoint's bold flag and [[FMT:text]] for any other explicit run formatting
(a size, italics, a hand-picked colour, a hyperlink). "apply" refuses to write
a paragraph that still carries those, so decide what each should be first.

"apply" rewrites only the content placeholders whose text differs from the
file. Each run is the level's default plus the requested look and nothing
else (no explicit sizes, fonts or colours, so the text keeps following the
template), the stale shrink-to-fit result is dropped so PowerPoint recomputes
it, and the slide's previous text is appended to the bottom of its notes
between marker lines. Tables, text boxes and pictures are left alone. A title
that differs from the file identifies a slide that moved, so the slide is
skipped, unless --retitle says the file's titles are the intended new ones
(written as one run in the title's existing formatting) or --force says to
edit the body regardless. Editing in place first copies the deck into a
backups/ folder beside it and refuses a deck that is open in PowerPoint.

Edits the packed XML directly and rezips with each part's original compression
(see retheme_fonts.py for why not a python-pptx round trip).
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import re
import shutil
import sys
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from add_title_picture_layout import Package  # noqa: E402

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"      # PowerPoint's <a14:m> math zone
MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"a": A, "p": P, "r": R, "a14": A14, "m": MATH, "mc": MC}
# What PowerPoint writes for a symbol typed into an equation: Cambria Math,
# italic, inside a math zone of its own. Whole equations are beyond this tool;
# it carries the single letters and operators the decks use ([[M:k]]).
MATH_FAMILY = "Cambria Math"
MATH_PANOSE = "02040503050406030204"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
RT_SLIDE = R + "/slide"
RT_NOTES = R + "/notesSlide"
RT_NOTES_MASTER = R + "/notesMaster"
RT_THEME = R + "/theme"
CT_NOTES = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

# The add-in's looks (macros/AlertText.bas): weight is a family name,
# stepped by outline level; the colour a theme accent tinted by level.
LOOKS = {  # look: (weight suffix on levels 1, 2, deeper), deepest level taking the bold flag, accent
    "A1": (("", "", "SemiBold"), 2, "accent1"),
    "A2": (("SemiBold", "SemiBold", "Medium"), 0, "accent2"),
    "B": (("SemiBold", "SemiBold", "Medium"), 0, None),
}
# How much of the accent is left at each level, the rest being white: an alert
# on a sub-bullet is the same colour, fainter. PowerPoint's own "Lighter 15%"
# is a lightness change in HSL, which holds HSL saturation and so *raises* the
# real chroma of a dark colour as it lightens -- Accent 2 at level 3 came out a
# brighter turquoise than the level-1 teal it was meant to be a faded version
# of. Blending toward white is what "the same colour, lighter" looks like, and
# writing the result as a plain colour rather than as transparency keeps it the
# same over a picture or the activity slides' tinted background.
# As whole percents, and mixed with integer arithmetic: the macro computes the
# same colours in VBA, and 70 % of 180 lands either side of 202.5 depending on
# how the two languages round a float.
BLEND = {1: 100, 2: 85, 3: 70}
DEEP_BLEND = 50
WEIGHT = {"Thin": 100, "ExtraLight": 200, "Light": 300, "": 400, "Medium": 500,
          "SemiBold": 600, "Bold": 700, "ExtraBold": 800, "Black": 900}
SUFFIXES = sorted((s for s in WEIGHT if s), key=len, reverse=True)   # ExtraLight before Light
PANOSE = {  # what PowerPoint records beside these families; optional but tidy
    "Inter": "02000503000000020004", "Inter Light": "02000403000000020004",
    "Inter ExtraLight": "02000303000000020004", "Inter Medium": "02000603000000020004",
    "Inter SemiBold": "02000703000000020004",
}

TOKEN = re.compile(r"\[\[(A1|A2|B|M|URL|BOLD|FMT):(.*?)\]\]")
TOKEN_START = re.compile(r"\[\[(A1|A2|B|M|URL|BOLD|FMT):")
UNRESOLVED = re.compile(r"\[\[(BOLD|FMT):")
BR = "<br>"
HEADER = re.compile(r"^## Slide (\d+)(?::\s*(.*))?$")
BODY_HEADER = re.compile(r"^\[body: (.+)\]$")
BULLET = re.compile(r"^( *)-(?: (.*))?$")
BACKUP_START = "--- Original slide text before the "
BACKUP_END = "--- End of original slide text ---"

NOTES_TEMPLATE = (
    f'<p:notes xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"><p:cSld><p:spTree>'
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/>'
    '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>'
    '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
    '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/>'
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr>'
    '</p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/></p:txBody></p:sp>'
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
)
RELS_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{RT_NOTES_MASTER}" Target="../notesMasters/{{master}}"/>'
    f'<Relationship Id="rId2" Type="{RT_SLIDE}" Target="../slides/{{slide}}"/>'
    "</Relationships>"
)


def q(tag: str) -> str:
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def split_family(name: str) -> tuple[str, str]:
    """('Inter', 'SemiBold') for 'Inter SemiBold'; ('Inter', '') for 'Inter'."""
    for suffix in SUFFIXES:
        if name.endswith(" " + suffix):
            return name[: -len(suffix) - 1], suffix
    return name, ""


def look_props(look: str, level: int, base: str) -> tuple[str, bool, str | None]:
    """Family, bold flag and accent for a look on an outline level.

    Weight 700 has no family of its own, so the levels that wear Bold reach it
    through the plain family name plus the flag."""
    suffixes, bold_through, accent = LOOKS[look]
    suffix = suffixes[min(level, 3) - 1]
    return (f"{base} {suffix}" if suffix else base), level <= bold_through, accent


class Master:
    """The body text style of the deck's first slide master: per level, the
    family and ink colour that text inherits when it says nothing itself."""

    def __init__(self, pkg: Package):
        name = min(n for n in pkg.names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n))
        root = parse(pkg, name)
        self.family: dict[int, str] = {}
        self.ink: dict[int, str | None] = {}
        for level in range(1, 10):
            ppr = root.find(f"p:txStyles/p:bodyStyle/a:lvl{level}pPr", NS)
            rpr = ppr.find("a:defRPr", NS) if ppr is not None else None
            latin = rpr.find("a:latin", NS) if rpr is not None else None
            family = latin.get("typeface") if latin is not None else None
            if not family or family.startswith("+"):
                family = self.family.get(level - 1, "Inter")
            self.family[level] = family
            srgb = rpr.find("a:solidFill/a:srgbClr", NS) if rpr is not None else None
            self.ink[level] = srgb.get("val").upper() if srgb is not None else self.ink.get(level - 1)
        self.base = split_family(self.family[1])[0]
        self.body_style = root.find("p:txStyles/p:bodyStyle", NS)
        self.accent = theme_accents(pkg, name)

    def alert(self, accent: str, level: int) -> str:
        """The colour an alert wears at this level: the accent, blended toward
        white by the level's share, as six hex digits."""
        keep = BLEND.get(level, DEEP_BLEND)
        base = self.accent[accent]
        channels = (int(base[i:i + 2], 16) for i in (0, 2, 4))
        return "".join(f"{(c * keep + 255 * (100 - keep) + 50) // 100:02X}" for c in channels)


def theme_accents(pkg: Package, master: str) -> dict[str, str]:
    """accent1 and accent2 of the theme this master uses."""
    rel = pkg.relationship(master, RT_THEME)
    root = parse(pkg, resolve(master, rel[1])) if rel else None
    out = {"accent1": "2F77B4", "accent2": "1E6663"}
    if root is None:
        return out
    for name in out:
        clr = root.find(f"a:themeElements/a:clrScheme/a:{name}/a:srgbClr", NS)
        if clr is not None:
            out[name] = clr.get("val").upper()
    return out


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------

def parse(pkg: Package, name: str) -> etree._Element:
    return etree.fromstring(pkg.read(name).encode("utf-8"))


def serialize(root: etree._Element) -> bytes:
    return DECL + etree.tostring(root, encoding="UTF-8")


def resolve(base: str, target: str) -> str:
    parts = (base.rsplit("/", 1)[0] + "/" + target).split("/")
    out: list[str] = []
    for part in parts:
        if part == "..":
            out.pop()
        elif part != ".":
            out.append(part)
    return "/".join(out)


def slide_parts(pkg: Package) -> list[str]:
    """The slide parts in presentation order."""
    presentation = pkg.read("ppt/presentation.xml")
    rels = pkg.read("ppt/_rels/presentation.xml.rels")
    parts = []
    for m in re.finditer(r'<p:sldId [^>]*r:id="([^"]+)"', presentation):
        rid = m.group(1)
        t = (re.search(rf'<Relationship [^>]*?Id="{rid}"[^>]*?Target="([^"]+)"', rels)
             or re.search(rf'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="{rid}"', rels))
        parts.append(resolve("ppt/presentation.xml", t.group(1)))
    return parts


def notes_part(pkg: Package, slide: str) -> str | None:
    rel = pkg.relationship(slide, RT_NOTES)
    return resolve(slide, rel[1]) if rel else None


def shape_name(sp: etree._Element) -> str:
    return sp.find(".//p:cNvPr", NS).get("name")


def placeholder(sp: etree._Element) -> etree._Element | None:
    return sp.find(".//p:nvPr/p:ph", NS)


def para_level(p: etree._Element) -> int:
    ppr = p.find("a:pPr", NS)
    return int(ppr.get("lvl", "0")) + 1 if ppr is not None else 1


def math_text(child: etree._Element) -> str:
    """The characters inside a math zone; a subscript comes back as x_1."""
    omath = child.find("m:oMath", NS)
    out = []
    for el in (omath if omath is not None else []):
        if etree.QName(el).localname == "sSub":
            base, sub = el.find("m:e", NS), el.find("m:sub", NS)
            out.append("_".join("".join(t.text or "" for t in part.iter(q("m:t")))
                                for part in (base, sub) if part is not None))
        else:
            out.append("".join(t.text or "" for t in el.iter(q("m:t"))))
    return "".join(out)


def math_zone_look(zone: etree._Element) -> str:
    return math_look(zone)


def in_fallback(el: etree._Element) -> bool:
    """True for the mc:Fallback copy of a shape, which is not the slide's text."""
    return any(a.tag == q("mc:Fallback") for a in el.iterancestors())


def alert_accent(rpr: etree._Element | None, level: int, master: Master) -> str | None:
    """Which accent an alert run wears, by colour: the level's alert colour, or
    the theme reference the add-in used to write before 18 September 2026."""
    if rpr is None:
        return None
    fill = rpr.find("a:solidFill", NS)
    if fill is None:
        return None
    scheme = fill.find("a:schemeClr", NS)
    if scheme is not None and scheme.get("val") in ("accent1", "accent2"):
        return scheme.get("val")
    srgb = fill.find("a:srgbClr", NS)
    if srgb is None:
        return None
    val = (srgb.get("val") or "").upper()
    for accent in ("accent1", "accent2"):
        if any(val == master.alert(accent, n) for n in (1, 2, 3, 4)):
            return accent
    return None


def math_look(child: etree._Element, level: int, master: Master) -> str:
    """Which look a symbol wears: only its colour can say."""
    rpr = child.find(".//a:rPr", NS)
    accent = alert_accent(rpr, level, master)
    return {"accent1": "A1", "accent2": "A2"}.get(accent, "plain")


def is_math(rpr: etree._Element | None) -> bool:
    latin = rpr.find("a:latin", NS) if rpr is not None else None
    return latin is not None and latin.get("typeface") == MATH_FAMILY


def plain_text(p: etree._Element) -> str:
    out = ""
    for child in p:
        tag = etree.QName(child).localname
        if tag in ("r", "fld"):
            t = child.find("a:t", NS)
            out += (t.text or "") if t is not None else ""
        elif tag == "m" and child.tag == q("a14:m"):
            out += math_text(child)
        elif tag == "br":
            out += " "
    return out


def classify(rpr: etree._Element | None, level: int, master: Master) -> str:
    """Which look a run wears: plain, A1, A2, B, or the flags BOLD and FMT for
    emphasis that matches no look."""
    if rpr is None:
        return "plain"
    fill = rpr.find("a:solidFill", NS)
    scheme = fill.find("a:schemeClr", NS) if fill is not None else None
    srgb = fill.find("a:srgbClr", NS) if fill is not None else None
    accent = alert_accent(rpr, level, master)
    if accent and not is_math(rpr):
        return "A1" if accent == "accent1" else "A2"
    if is_math(rpr):
        return "M"
    default_family = master.family.get(level) or master.family[1]
    latin = rpr.find("a:latin", NS)
    family = latin.get("typeface") if latin is not None else None
    if family is not None and family.startswith("+"):
        family = None
    bolder_family = look_props("B", level, master.base)[0]
    odd = (
        rpr.get("i") == "1"
        or rpr.get("u") not in (None, "none")
        or rpr.get("strike") not in (None, "noStrike")
        or rpr.get("sz") is not None
        or rpr.find("a:hlinkClick", NS) is not None
        or rpr.find("a:highlight", NS) is not None
        or scheme is not None
        or (srgb is not None and srgb.get("val").upper() != (master.ink.get(level) or "")
            and not accent)
        or (family not in (None, default_family, bolder_family))
    )
    bold = rpr.get("b") == "1"
    link = rpr.find("a:hlinkClick", NS) is not None
    only_link = link and not (
        rpr.get("i") == "1" or rpr.get("u") not in (None, "none") or rpr.get("sz") is not None
        or scheme is not None or srgb is not None or bold
        or family not in (None, default_family))
    if only_link:
        return "URL"
    if odd:
        return "FMT"
    if bold:
        return "BOLD"
    if family == bolder_family and family != default_family:
        return "B"
    return "plain"


def para_markup(p: etree._Element, level: int, master: Master) -> str:
    pieces: list[tuple[str, str]] = []
    for child in p:
        tag = etree.QName(child).localname
        if tag == "m" and child.tag == q("a14:m"):
            pieces.append((f"[[M:{math_text(child)}]]", math_look(child, level, master)))
            continue
        if tag == "r":
            text = child.find("a:t", NS).text or ""
            rpr = child.find("a:rPr", NS)
            look = classify(rpr, level, master)
            if look == "M":
                # already marked up, so a symbol inside an alert comes back as
                # [[A1:[[M:k]]-nearest neighbours]] and keeps its colour
                if rpr is not None and rpr.get("baseline") and pieces and pieces[-1][0].endswith("]]"):
                    head, sub = pieces[-1][0][:-2], text        # x and 1 are one symbol
                    pieces[-1] = (f"{head}_{sub}]]", pieces[-1][1])
                else:
                    pieces.append((f"[[M:{text}]]", math_look(child)))
                continue
            pieces.append((text, look))
        elif tag == "m" and child.tag == q("a14:m"):
            # already marked up, so that a symbol inside an alert comes back as
            # [[A1:[[M:k]]-nearest neighbours]] and keeps its colour on apply
            pieces.append((f"[[M:{math_text(child)}]]", math_look(child, level, master)))
        elif tag == "br":
            pieces.append((BR, "plain"))
        elif tag == "fld":
            t = child.find("a:t", NS)
            pieces.append(((t.text or "") if t is not None else "", "plain"))
    merged: list[tuple[str, str]] = []
    for text, look in pieces:
        if merged and merged[-1][1] == look:
            merged[-1] = (merged[-1][0] + text, look)
        else:
            merged.append((text, look))
    out = ""
    for text, look in merged:
        if look == "plain" or not text.strip():
            out += text
        else:
            lead = text[: len(text) - len(text.lstrip())]
            trail = text[len(text.rstrip()):]
            out += f"{lead}[[{look}:{text.strip()}]]{trail}"
    return out


def line_for(level: int, text: str) -> str:
    indent = "    " * (level - 1)
    return f"{indent}- {text}" if text else f"{indent}-"


def body_lines(txBody: etree._Element, master: Master) -> list[str]:
    return [line_for(para_level(p), para_markup(p, para_level(p), master).strip())
            for p in txBody.findall("a:p", NS)]


class SlideView:
    """A slide's text: its title, its editable content placeholders, and the
    read-only rest (tables, text boxes, subtitles)."""

    def __init__(self, pkg: Package, part: str, master: Master):
        self.part = part
        self.root = parse(pkg, part)
        self.title: str | None = None
        self.title_body: etree._Element | None = None
        self.bodies: list[tuple[str, etree._Element]] = []
        self.readonly: list[tuple[str, str, list[str]]] = []
        tree = self.root.find("p:cSld/p:spTree", NS)
        for sp in tree.iter(q("p:sp"), q("p:graphicFrame")):
            if in_fallback(sp):
                continue          # the copy for readers without maths, not the slide's text
            ph = placeholder(sp)
            kind = ph.get("type") if ph is not None else None
            if etree.QName(sp).localname == "graphicFrame":
                tbl = sp.find(".//a:tbl", NS)
                if tbl is not None:
                    rows = ["| " + " | ".join(
                        " / ".join(plain_text(p).strip() for p in tc.findall("a:txBody/a:p", NS))
                        for tc in tr.findall("a:tc", NS)) + " |" for tr in tbl.findall("a:tr", NS)]
                    self.readonly.append(("table", shape_name(sp), rows))
                continue
            txBody = sp.find("p:txBody", NS)
            if txBody is None or (ph is not None and kind in ("ftr", "sldNum", "dt")):
                continue
            if ph is not None and kind in ("title", "ctrTitle"):
                self.title = " ".join(plain_text(p) for p in txBody.findall("a:p", NS)).strip()
                self.title_body = txBody
                continue
            # a content placeholder, or one that was detached from its layout
            # (PowerPoint keeps the "Content Placeholder N" name and copies the
            # body list style into the box)
            detached = ph is None and shape_name(sp).startswith("Content Placeholder") \
                and txBody.find("a:lstStyle/a:lvl1pPr", NS) is not None
            if (ph is not None and kind in (None, "body", "obj")) or detached:
                name = shape_name(sp)
                taken = {n for n, _ in self.bodies}
                k = 2
                while name in taken:          # two shapes may share a name
                    name = f"{shape_name(sp)} ({k})"
                    k += 1
                self.bodies.append((name, txBody))
                continue
            lines = [plain_text(p).strip() for p in txBody.findall("a:p", NS)]
            if any(lines):
                self.readonly.append(("text box" if ph is None else kind, shape_name(sp), lines))


def notes_lines(pkg: Package, notes: str) -> list[str]:
    root = parse(pkg, notes)
    txBody = notes_body(root)
    if txBody is None:
        return []
    out, inside = [], False
    for p in txBody.findall("a:p", NS):
        text = plain_text(p).strip()
        if text.startswith(BACKUP_START):
            inside = True
            out.append("> [copy of the slide's earlier text omitted]")
        elif inside:
            inside = text != BACKUP_END
        elif text:
            out.append(f"> {text}")
    return out


def dump(pkg: Package, indices: set[int]) -> str:
    master = Master(pkg)
    out = [
        f"# {pkg.path.name}: slide text written by python/deck_text.py on {dt.date.today().isoformat()}.",
        "# Edit the '- ' lines, then: python python/deck_text.py apply DECK THIS-FILE [-o out.pptx]",
        "# [[A1:..]] alert 1, [[A2:..]] alert 2, [[B:..]] bolder; [[BOLD:..]] and [[FMT:..]] flag",
        "# formatting that matches no look and must be resolved before apply will write the line.",
        "# Lines starting with '>' (notes), '|' (read-only shapes) and '#' are ignored on apply.",
    ]
    for index, part in enumerate(slide_parts(pkg), 1):
        if indices and index not in indices:
            continue
        view = SlideView(pkg, part, master)
        out.append("")
        out.append(f"## Slide {index}: {view.title}" if view.title else f"## Slide {index}:")
        for name, txBody in view.bodies:
            if len(view.bodies) > 1:
                out.append(f"[body: {name}]")
            out.extend(body_lines(txBody, master))
        for kind, name, lines in view.readonly:
            out.append(f"[read-only {kind}: {name}]")
            out.extend(f"| {line}" for line in lines)
        notes = notes_part(pkg, part)
        if notes:
            out.extend(notes_lines(pkg, notes))
    return "\n".join(out) + "\n"


# ---------------------------------------------------------------------------
# The text file
# ---------------------------------------------------------------------------

class SpecSlide:
    def __init__(self, index: int, title: str):
        self.index = index
        self.title = title
        self.bodies: dict[str | None, list[tuple[int, str]]] = {}


def parse_spec(text: str) -> dict[int, SpecSlide]:
    slides: dict[int, SpecSlide] = {}
    current: SpecSlide | None = None
    body: str | None = None
    for number, raw in enumerate(text.splitlines(), 1):
        line = raw.rstrip()
        if not line.strip():
            continue
        m = HEADER.match(line)
        if m:
            current = SpecSlide(int(m.group(1)), (m.group(2) or "").strip())
            slides[current.index] = current
            body = None
            continue
        if line[0] in "#>|" or line.startswith("[read-only"):
            continue
        m = BODY_HEADER.match(line)
        if m and current is not None:
            body = m.group(1)
            current.bodies.setdefault(body, [])
            continue
        m = BULLET.match(line)
        if m and current is not None:
            spaces = len(m.group(1))
            if spaces % 4:
                raise ValueError(f"line {number}: indent by four spaces per level, not {spaces}")
            current.bodies.setdefault(body, []).append((spaces // 4 + 1, (m.group(2) or "").strip()))
            continue
        raise ValueError(f"line {number}: cannot read {raw!r}")
    return slides


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------

def lang_of(txBody: etree._Element, default: str) -> str:
    for el in txBody.iter(q("a:rPr"), q("a:endParaRPr")):
        if el.get("lang"):
            return el.get("lang")
    return default


def make_rpr(parent: etree._Element, look: str, level: int, master: Master, lang: str,
             links: dict[str, etree._Element] | None = None, text: str = "") -> None:
    rpr = etree.SubElement(parent, q("a:rPr"))
    rpr.set("lang", lang)
    if look == "plain":
        rpr.set("dirty", "0")
        return
    if look == "URL":
        # keep the hyperlink the old text carried on the same words, or the
        # body's only hyperlink when the words changed
        rpr.set("dirty", "0")
        links = links or {}
        link = links.get(text.strip())
        if link is None and len(links) == 1:
            link = next(iter(links.values()))
        if link is None:
            raise ValueError(f"[[URL:{text}]]: no matching hyperlink in the slide's previous text")
        rpr.append(copy.deepcopy(link))
        return
    family, bold, accent = look_props(look, level, master.base)
    if bold:
        rpr.set("b", "1")
    rpr.set("dirty", "0")
    if accent:
        fill = etree.SubElement(rpr, q("a:solidFill"))
        etree.SubElement(fill, q("a:srgbClr"), val=master.alert(accent, level))
    latin = etree.SubElement(rpr, q("a:latin"), typeface=family)
    if PANOSE.get(family):
        latin.set("panose", PANOSE[family])
    latin.set("pitchFamily", "2")
    latin.set("charset", "0")


def make_math(p: etree._Element, text: str, lang: str, look: str, level: int,
              master: Master) -> None:
    """A symbol as PowerPoint's equation editor writes it: a math zone in
    Cambria Math italic, carrying the surrounding look's colour and weight so a
    symbol inside a term is part of the term. x_1 becomes a subscript.

    The shape that holds one is wrapped in mc:AlternateContent by
    wrap_for_math() below, which is what keeps the PDF right."""
    zone = etree.SubElement(p, q("a14:m"), nsmap={"a14": A14})
    omath = etree.SubElement(zone, q("m:oMath"), nsmap={"m": MATH})
    base, sep, sub = text.partition("_")
    if not sep:
        math_run(omath, base, lang, look, level, master)
        return
    ssub = etree.SubElement(omath, q("m:sSub"))
    math_rpr(etree.SubElement(etree.SubElement(ssub, q("m:sSubPr")), q("m:ctrlPr")),
             lang, look, level, master)
    math_run(etree.SubElement(ssub, q("m:e")), base, lang, look, level, master)
    math_run(etree.SubElement(ssub, q("m:sub")), sub, lang, look, level, master)


def math_rpr(parent: etree._Element, lang: str, look: str, level: int, master: Master) -> None:
    """The run properties of a symbol: the look's colour, Cambria Math italic."""
    _, bold, accent = look_props(look, level, master.base) if look in LOOKS else ("", False, None)
    rpr = etree.SubElement(parent, q("a:rPr"), lang=lang)
    if bold:
        rpr.set("b", "1")
    rpr.set("i", "1")
    rpr.set("dirty", "0")
    if accent:
        fill = etree.SubElement(rpr, q("a:solidFill"))
        etree.SubElement(fill, q("a:srgbClr"), val=master.alert(accent, level))
    etree.SubElement(rpr, q("a:latin"), typeface=MATH_FAMILY, panose=MATH_PANOSE,
                     pitchFamily="18", charset="0")


def math_run(parent: etree._Element, text: str, lang: str, look: str, level: int,
             master: Master) -> None:
    r = etree.SubElement(parent, q("m:r"))
    math_rpr(r, lang, look, level, master)
    etree.SubElement(r, q("m:t")).text = text


def math_symbols(zone: etree._Element):
    """(text, run properties, is a subscript) for each run of a math zone."""
    for r in zone.iter(q("m:r")):
        sub = any(a.tag == q("m:sub") for a in r.iterancestors())
        t = r.find("m:t", NS)
        yield (t.text or "" if t is not None else ""), r.find("a:rPr", NS), sub


def flattened(sp: etree._Element) -> etree._Element:
    """A copy of the shape with its math zones as ordinary Cambria Math runs.

    This is the mc:Fallback: what a reader that does not do OOXML maths shows.
    LibreOffice, which converts the decks to the published PDFs, is one, and
    without a fallback it drops the symbols from the slide altogether.
    PowerPoint writes a picture of the shape here instead, and will replace
    this the next time the deck is saved."""
    clone = copy.deepcopy(sp)
    for zone in list(clone.iter(q("a14:m"))):
        parent = zone.getparent()
        at = parent.index(zone)
        for offset, (text, rpr, sub) in enumerate(math_symbols(zone)):
            run = etree.Element(q("a:r"))
            props = copy.deepcopy(rpr) if rpr is not None else etree.Element(q("a:rPr"))
            props.attrib.pop("smtClean", None)
            if sub:
                props.set("baseline", "-25000")
            run.append(props)
            etree.SubElement(run, q("a:t")).text = text
            parent.insert(at + offset, run)
        parent.remove(zone)
    return clone


def wrap_for_math(sp: etree._Element) -> None:
    """Put a shape that holds maths into the mc:AlternateContent pair
    PowerPoint uses, or take it out again when the maths has gone."""
    parent = sp.getparent()
    wrapped = parent.tag == q("mc:Choice")
    has_math = sp.find(".//" + q("a14:m")) is not None
    if not wrapped and not has_math:
        return
    if wrapped:
        alt = parent.getparent()
        tree, at = alt.getparent(), alt.getparent().index(alt)
        parent.remove(sp)
        tree.remove(alt)
        tree.insert(at, sp)
        if not has_math:
            return
        parent = tree
    at = parent.index(sp)
    alt = etree.Element(q("mc:AlternateContent"), nsmap={"mc": MC})
    choice = etree.SubElement(alt, q("mc:Choice"), nsmap={"a14": A14})
    choice.set("Requires", "a14")
    fallback = etree.SubElement(alt, q("mc:Fallback"))
    parent.remove(sp)
    fallback.append(flattened(sp))
    choice.append(sp)
    parent.insert(at, alt)


def scan_tokens(text: str):
    """The outermost [[LOOK:..]] tokens, as (start, end, look, inner text).

    A regular expression cannot do this: [[A1:[[M:k]]-nearest neighbours]] has a
    token inside a token, and a non-greedy match stops at the first "]]"."""
    i = 0
    while (m := TOKEN_START.search(text, i)) is not None:
        depth, j = 1, m.end()
        while j < len(text) and depth:
            if text.startswith("[[", j):
                depth, j = depth + 1, j + 2
            elif text.startswith("]]", j):
                depth, j = depth - 1, j + 2
            else:
                j += 1
        yield m.start(), j, m.group(1), text[m.end():j - 2]
        i = j


def emit_marked(p: etree._Element, text: str, look: str, level: int, master: Master, lang: str,
                links: dict[str, etree._Element] | None = None) -> None:
    """Emit a run of text and everything marked up inside it."""
    pos = 0
    for start, end, inner, chunk in scan_tokens(text):
        emit_runs(p, text[pos:start], look, level, master, lang, links)
        if inner == "M":
            make_math(p, chunk, lang, look, level, master)
        else:
            emit_marked(p, chunk, inner, level, master, lang, links)
        pos = end
    emit_runs(p, text[pos:], look, level, master, lang, links)


def emit_runs(p: etree._Element, text: str, look: str, level: int, master: Master, lang: str,
              links: dict[str, etree._Element] | None = None) -> None:
    for k, chunk in enumerate(text.split(BR)):
        if k:
            make_rpr(etree.SubElement(p, q("a:br")), "plain", level, master, lang)
        if chunk:
            r = etree.SubElement(p, q("a:r"))
            make_rpr(r, look, level, master, lang, links, chunk)
            t = etree.SubElement(r, q("a:t"))
            t.text = chunk
            if chunk != chunk.strip():
                t.set(XML_SPACE, "preserve")


def make_paragraph(level: int, text: str, master: Master, lang: str,
                   links: dict[str, etree._Element] | None = None) -> etree._Element:
    p = etree.Element(q("a:p"))
    if level > 1:
        etree.SubElement(p, q("a:pPr"), lvl=str(level - 1))
    emit_marked(p, text, "plain", level, master, lang, links)
    etree.SubElement(p, q("a:endParaRPr"), lang=lang, dirty="0")
    return p


def rewrite_body(txBody: etree._Element, paragraphs: list[tuple[int, str]], master: Master, lang: str) -> None:
    sp = txBody.getparent()
    lst = txBody.find("a:lstStyle", NS)
    if placeholder(sp) is None and lst is not None and len(lst) and master.body_style is not None:
        # a detached placeholder carries a copy of some old body style; give it
        # the master's, so its levels look like every other body's
        for child in list(lst):
            lst.remove(child)
        for level in master.body_style:
            lst.append(copy.deepcopy(level))
    links: dict[str, etree._Element] = {}
    for r in txBody.iter(q("a:r")):
        link = r.find("a:rPr/a:hlinkClick", NS)
        if link is not None:
            links[(r.find("a:t", NS).text or "").strip()] = link
    for p in txBody.findall("a:p", NS):
        txBody.remove(p)
    for level, text in paragraphs:
        txBody.append(make_paragraph(level, text, master, lang, links))
    autofit = txBody.find("a:bodyPr/a:normAutofit", NS)
    if autofit is not None:       # the recorded shrink belongs to the old text
        autofit.attrib.clear()


def set_title(txBody: etree._Element, title: str, lang: str) -> None:
    """Replace a title's text with one run in the formatting of its first run."""
    paragraphs = txBody.findall("a:p", NS)
    first = paragraphs[0]
    for extra in paragraphs[1:]:
        txBody.remove(extra)
    template = None
    for child in list(first):
        tag = etree.QName(child).localname
        if tag == "r" and template is None:
            rpr = child.find("a:rPr", NS)
            template = copy.deepcopy(rpr) if rpr is not None else None
        if tag in ("r", "br", "fld"):
            first.remove(child)
    r = etree.Element(q("a:r"))
    if template is not None:
        r.append(template)
    else:
        rpr = etree.SubElement(r, q("a:rPr"))
        rpr.set("lang", lang)
        rpr.set("dirty", "0")
    t = etree.SubElement(r, q("a:t"))
    t.text = title
    end = first.find("a:endParaRPr", NS)
    if end is not None:
        end.addprevious(r)
    else:
        first.append(r)


def notes_body(root: etree._Element) -> etree._Element | None:
    for sp in root.iter(q("p:sp")):
        ph = placeholder(sp)
        if ph is not None and ph.get("type") == "body":
            return sp.find("p:txBody", NS)
    return None


def note_paragraph(text: str, lang: str) -> etree._Element:
    p = etree.Element(q("a:p"))
    if not text:
        etree.SubElement(p, q("a:endParaRPr"), lang=lang, dirty="0")
        return p
    r = etree.SubElement(p, q("a:r"))
    etree.SubElement(r, q("a:rPr"), lang=lang, dirty="0")
    t = etree.SubElement(r, q("a:t"))
    t.text = text
    if text != text.strip():
        t.set(XML_SPACE, "preserve")
    return p


def records(txBody: etree._Element) -> list[tuple[int, int]]:
    """Where each original-text block starts and ends among the paragraphs."""
    texts = [plain_text(p).strip() for p in txBody.findall("a:p", NS)]
    out, start = [], None
    for i, text in enumerate(texts):
        if text.startswith(BACKUP_START):
            start = i
        elif text == BACKUP_END and start is not None:
            out.append((start, i))
            start = None
    if start is not None:
        out.append((start, len(texts) - 1))
    return out


def trim_extra_records(txBody: etree._Element) -> bool:
    """Keep the first original-text block and drop any later ones."""
    blocks = records(txBody)
    if len(blocks) < 2:
        return False
    paragraphs = txBody.findall("a:p", NS)
    for start, end in reversed(blocks[1:]):
        while start and not plain_text(paragraphs[start - 1]).strip():
            start -= 1                     # the blank line that separates it
        for p in paragraphs[start:end + 1]:
            txBody.remove(p)
    return True


def add_notes_backup(pkg: Package, slide: str, lines: list[str], date: str) -> bool:
    """Append the slide's previous text to its notes (creating the notes part
    if the slide has none); False when that exact block is already there."""
    block = [f"{BACKUP_START}{date} edit; [[A1:..]] alert 1, [[A2:..]] alert 2, "
             "[[B:..]] bolder, [[BOLD:..]] plain bold ---", *lines, BACKUP_END]
    name = notes_part(pkg, slide)
    if name is None:
        masters = sorted(n for n in pkg.names if re.match(r"ppt/notesMasters/notesMaster\d+\.xml$", n))
        if not masters:
            raise ValueError("the deck has no notes master; add a note to any slide in PowerPoint first")
        used = [int(m.group(1)) for n in [*pkg.names, *pkg.added]
                if (m := re.match(r"ppt/notesSlides/notesSlide(\d+)\.xml$", n))]
        name = f"ppt/notesSlides/notesSlide{max(used, default=0) + 1}.xml"
        short = name.rsplit("/", 1)[1]
        pkg.write(name, NOTES_TEMPLATE)
        pkg.write(f"ppt/notesSlides/_rels/{short}.rels", RELS_TEMPLATE.format(
            master=masters[0].rsplit("/", 1)[1], slide=slide.rsplit("/", 1)[1]))
        pkg.add_relationship(slide, RT_NOTES, f"../notesSlides/{short}")
        pkg.add_override(name, CT_NOTES)
    root = parse(pkg, name)
    txBody = notes_body(root)
    if txBody is None:
        raise ValueError(f"{name} has no notes placeholder")
    existing = txBody.findall("a:p", NS)
    texts = [plain_text(p).strip() for p in existing]
    if any(text.startswith(BACKUP_START) for text in texts):
        # The notes record what the slide said before anyone started editing it,
        # once. A second block would make them a changelog, which is not what
        # they are for: the earlier text is the one worth keeping.
        if trim_extra_records(txBody):
            pkg.write_bytes(name, serialize(root))
            return True
        return False
    lang = lang_of(txBody, "en-CA")
    if any(texts):
        txBody.append(note_paragraph("", lang))
    else:
        for p in existing:
            txBody.remove(p)
    for line in block:
        txBody.append(note_paragraph(line, lang))
    pkg.write_bytes(name, serialize(root))
    return True


def normalised(title: str | None) -> str:
    return " ".join((title or "").split())


def apply(pkg: Package, spec: dict[int, SpecSlide], date: str, force: bool, dry_run: bool,
          retitle: bool = False) -> list[str]:
    master = Master(pkg)
    parts = slide_parts(pkg)
    report = []
    for index in sorted(spec):
        entry = spec[index]
        if not 1 <= index <= len(parts):
            report.append(f"slide {index}: the deck has {len(parts)} slides; skipped")
            continue
        view = SlideView(pkg, parts[index - 1], master)
        new_title = None
        if normalised(entry.title) != normalised(view.title):
            if retitle and view.title_body is not None and entry.title:
                new_title = entry.title
            elif not force:
                report.append(f'slide {index}: the file says "{entry.title}" but the deck\'s title is '
                              f'"{view.title}"; skipped (--retitle to change titles, --force to edit anyway)')
                continue
        if not entry.bodies and new_title is None:
            continue
        names = [name for name, _ in view.bodies]
        mapping: dict[str, list[tuple[int, str]]] = {}
        if None in entry.bodies:
            if len(view.bodies) != 1:
                report.append(f"slide {index}: {len(view.bodies)} content placeholders, so label the file's "
                              f"paragraphs with [body: NAME] ({', '.join(names) or 'none'}); skipped")
                continue
            mapping[names[0]] = entry.bodies[None]
        missing = [name for name in entry.bodies if name is not None and name not in names]
        if missing:
            report.append(f'slide {index}: no content placeholder named "{missing[0]}"; skipped')
            continue
        mapping.update({name: paragraphs for name, paragraphs in entry.bodies.items() if name is not None})
        changed, backup, blocked = [], ([f"Title: {view.title}"] if view.title else []), False
        for name, txBody in view.bodies:
            if name not in mapping:
                continue
            new = mapping[name]
            old_lines = body_lines(txBody, master)
            if [line_for(level, text) for level, text in new] == old_lines:
                continue
            if any(UNRESOLVED.search(text) for _, text in new):
                report.append(f'slide {index}: "{name}" still has [[BOLD:..]] or [[FMT:..]] tokens; '
                              "give each a look first; skipped")
                blocked = True
                break
            if len(view.bodies) > 1:
                backup.append(f"[body: {name}]")
            backup.extend(old_lines)
            if not dry_run:
                rewrite_body(txBody, new, master, lang_of(txBody, "en-US"))
                wrap_for_math(txBody.getparent())
            changed.append(name)
        if blocked or (not changed and new_title is None):
            continue
        if not dry_run:
            if new_title is not None and view.title_body is not None:
                set_title(view.title_body, new_title, lang_of(view.title_body, "en-US"))
            pkg.write_bytes(view.part, serialize(view.root))
            add_notes_backup(pkg, view.part, backup, date)
        done = []
        if new_title is not None:
            done.append(f'title "{view.title}" -> "{new_title}"')
        if changed:
            done.append(f"rewrote {', '.join(changed)}")
        report.append(f"slide {index}: {'would: ' if dry_run else ''}{'; '.join(done)}")
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    d = sub.add_parser("dump", help="write the deck's slide text as an editable file")
    d.add_argument("deck", type=Path)
    d.add_argument("--slides", type=int, nargs="*", default=[], help="only these slide numbers")
    d.add_argument("-o", "--output", type=Path, help="write here instead of standard output")
    a = sub.add_parser("apply", help="apply an edited text file to the deck")
    a.add_argument("deck", type=Path)
    a.add_argument("spec", type=Path, help="the text file, as written by dump and edited")
    a.add_argument("-o", "--output", type=Path, help="write the result here instead of in place")
    a.add_argument("--dry-run", action="store_true", help="report what would change and write nothing")
    a.add_argument("--force", action="store_true", help="edit slides whose title differs from the file")
    a.add_argument("--retitle", action="store_true", help="write the file's titles into the deck instead of treating a different title as a moved slide")
    a.add_argument("--no-backup", action="store_true", help="in place without first copying the deck into backups/")
    a.add_argument("--date", default=dt.date.today().isoformat(), help="date written into the notes-pane copies")
    args = parser.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")

    if args.command == "dump":
        text = dump(Package(args.deck), set(args.slides))
        if args.output:
            args.output.write_text(text, encoding="utf-8")
        else:
            sys.stdout.write(text)
        return 0

    spec = parse_spec(args.spec.read_text(encoding="utf-8"))
    target = args.deck
    if args.output and not args.dry_run:
        if args.output.resolve() == args.deck.resolve():
            parser.error("-o must not be the deck itself; leave it out to edit in place")
        shutil.copyfile(args.deck, args.output)
        target = args.output
    elif not args.dry_run:
        if (target.parent / f"~${target.name}").exists():
            print(f"{target.name}: open in PowerPoint; close it first")
            return 1
        if not args.no_backup:
            backups = target.parent / "backups"
            backups.mkdir(exist_ok=True)
            copy = backups / f"{target.stem} - before deck_text {dt.datetime.now():%Y-%m-%d %H%M}{target.suffix}"
            if not copy.exists():
                shutil.copy2(target, copy)
                print(f"backed up to {copy}")
    pkg = Package(target)
    report = apply(pkg, spec, args.date, args.force, args.dry_run, args.retitle)
    if args.dry_run:
        pkg.zin.close()
    else:
        pkg.save()
    print("\n".join(report) if report else "nothing to change")
    return 0


if __name__ == "__main__":
    sys.exit(main())
