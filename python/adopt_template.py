"""Finish moving a lecture deck onto the teaching template.

Applying `templates/Presentation-template-Inter-teaching.potx` (Design > Browse
for Themes in PowerPoint, or `ApplyTemplate` through COM) swaps in the new
master and layouts, but PowerPoint leaves two kinds of debris behind:

1. **Merged layout decoration.** Where an old layout and a new one share a name,
   PowerPoint keeps the old layout's user-drawn shapes on top of the new one: the
   gold underline connector, the old activity icon, the gradient course bar. It
   also keeps any old layout it could not match, renamed ("1_Title Slide").
   With `--template`, this script deletes layout shapes the template does not
   have and unused layouts the template does not define, plus the media only
   they referenced.

2. **Run-level overrides on the slides.** The title slide's explicit Roboto and
   gold "Lecture" colour hide the layout's Black title and Accent 1 lecture
   number; `<a:ea typeface="Roboto"/>` sits on most titles; pasted web text
   carries hard colours, an explicit body family and explicit point sizes that
   block the level styles; old-theme bullet overrides hide the template's
   bullets; and the ad-hoc emphasis of bold plus 0070C0 blue predates alert
   text, which is Bold on the first two outline levels and
   SemiBold below, in theme Accent 1. All of these are rewritten to inherit.

Edits the packed XML by string substitution and rezips with each part's
original compression (see retheme_fonts.py for why not python-pptx). Follow
with `retheme_fonts.py` for the Lato and Roboto font names themselves.

    python python/adopt_template.py --template templates/Presentation-template-Inter-teaching.potx decks/Lecture1.pptx
    python python/adopt_template.py --check decks/Lecture1.pptx
"""

from __future__ import annotations

import argparse
import posixpath
import re
import shutil
import zipfile
from pathlib import Path

BODY_FAMILY = "Inter"
# Alert 1 is Bold on levels 1 and 2 and SemiBold below. It was a weight
# lighter on each level below the first until 16 September 2026, which the
# tint left too faint to carry the eye. Weight 700 has no family of its own --
# Regular and Bold share the plain name -- so a Bold level is the plain family
# with the bold flag, which selects the genuine Bold face. Keep in step with
# macros/AlertText.bas and python/deck_text.py.
ALERT_BY_LEVEL = {1: ("Inter", True), 2: ("Inter", True)}
ALERT_DEFAULT = ("Inter SemiBold", False)
ALERT_FAMILIES = {"Inter SemiBold", "Inter Medium"}   # weight-named alert families
# Alert colours lighten with the level so an alert sits at about the lightness
# of the text around it. This tool writes the form the decks used until
# 18 September 2026, a theme reference with lumMod/lumOff; the colour is now
# the accent blended toward white and written plainly (STYLE.md rule
# on emphasis, python/deck_text.py's BLEND). Run python/tidy_slide_text.py
# after this to bring a converted deck the rest of the way.
ALERT_TINTS = {1: 0, 2: 15, 3: 30}
ALERT_TINT_DEEP = 50
ACCENT = re.compile(r'<a:schemeClr val="(accent[12])"(?:/>|>.*?</a:schemeClr>)', re.S)


def tinted(scheme: str, level: int) -> str:
    pct = ALERT_TINTS.get(level, ALERT_TINT_DEEP)
    if not pct:
        return f'<a:schemeClr val="{scheme}"/>'
    return (f'<a:schemeClr val="{scheme}"><a:lumMod val="{(100 - pct) * 1000}"/>'
            f'<a:lumOff val="{pct * 1000}"/></a:schemeClr>')

LEGACY_ALERT_COLOURS = {"0070C0"}
LEGACY_TITLE_FONTS = {"Roboto", "Libre Caslon Display"}   # the old theme's title faces
LEGACY_TITLE_COLOURS = {"AA7509", "C38602"}   # the old gold "Lecture N"
# Hard colours that came in with text pasted from the course website: its body
# grey and its link green. On the slide they hide the level colours, and any
# new bullet typed beside them inherits them.
PASTED_TEXT_COLOURS = {"212529", "0A5C36"}

RUN = re.compile(r"<a:(rPr|endParaRPr)([^>/]*)>(.*?)</a:\1>", re.S)
EMPTY_RUN = re.compile(r"<a:rPr([^>]*)/>")
PARA_PROPS = re.compile(r"<a:pPr([^>/]*)>(.*?)</a:pPr>", re.S)
PARA = re.compile(r"<a:p>.*?</a:p>", re.S)
SHAPE = re.compile(r"<p:(sp|pic|cxnSp|grpSp|graphicFrame)>.*?</p:\1>", re.S)
BODY_PLACEHOLDER = re.compile(r"<p:sp>(?:(?!</p:sp>).)*?<p:ph(?: type=\"(?:body|obj)\")?(?: sz=\"[^\"]*\")? idx=\"\d+\"/>.*?</p:sp>", re.S)
LAYOUT_PART = re.compile(r"ppt/slideLayouts/slideLayout\d+\.xml$")
SLIDE_PART = re.compile(r"ppt/slides/slide\d+\.xml$")


def alert_spec(level: int) -> tuple[str, bool]:
    return ALERT_BY_LEVEL.get(level, ALERT_DEFAULT)


def _with_bold(attrs: str, bold: bool) -> str:
    attrs = attrs.replace(' b="1"', "").replace(' b="0"', "")
    return attrs + ' b="1"' if bold else attrs


def _set_family(body: str, family: str) -> str:
    body = re.sub(r'<a:(?:latin|ea|cs) typeface="[^"]*"[^>]*/>', "", body)
    return body + f'<a:latin typeface="{family}"/>'


# --------------------------------------------------------------------------
# Slide text
# --------------------------------------------------------------------------
def rewrite_runs(xml: str) -> tuple[str, dict[str, int]]:
    counts: dict[str, int] = {}
    level = 1          # outline level of the paragraph being rewritten

    def bump(key: str) -> None:
        counts[key] = counts.get(key, 0) + 1

    def fix(match: re.Match[str]) -> str:
        tag, attrs, body = match.group(1), match.group(2), match.group(3)
        new_body = body
        family, bold = alert_spec(level)

        for face in LEGACY_TITLE_FONTS:
            pat = r'<a:(?:latin|ea|cs) typeface="' + re.escape(face) + r'"[^>]*/>'
            if re.search(pat, new_body):
                new_body = re.sub(pat, "", new_body)
                bump(f"{face} override dropped")
        for colour in LEGACY_TITLE_COLOURS:
            fill = f'<a:solidFill><a:srgbClr val="{colour}"/></a:solidFill>'
            if fill in new_body:
                new_body = new_body.replace(fill, "")
                bump(f"{colour} colour dropped")

        # Legacy emphasis, bold plus hard blue, becomes alert text.
        for colour in LEGACY_ALERT_COLOURS:
            fill = f'<a:solidFill><a:srgbClr val="{colour}"/></a:solidFill>'
            if ' b="1"' in attrs and fill in new_body:
                attrs = _with_bold(attrs, bold)
                new_body = new_body.replace(fill, f'<a:solidFill>{tinted("accent1", level)}</a:solidFill>')
                new_body = _set_family(new_body, family)
                bump("alert text")

        # Alert text set at another level's weight (earlier rules used one
        # weight everywhere) moves to this level's alert weight.
        for fam in ALERT_FAMILIES - {family}:
            if f'typeface="{fam}"' in new_body:
                new_body = _set_family(new_body, family)
                attrs = _with_bold(attrs, bold)
                bump("alert weight re-levelled")

        # Alert colours carry the level's tint.
        def retint(m: re.Match[str]) -> str:
            want = tinted(m.group(1), level)
            if m.group(0) != want:
                bump("alert tint set")
            return want
        new_body = ACCENT.sub(retint, new_body)

        # Anything in Accent 1 is alert text by convention; make sure it has
        # this level's alert family and bold state, whatever produced it.
        if tag == "rPr" and '<a:schemeClr val="accent1"' in new_body:
            cur = re.search(r'<a:latin typeface="([^"]*)"', new_body)
            effective = cur.group(1) if cur else (BODY_FAMILY if level == 1 else None)
            if effective != family or (' b="1"' in attrs) != bold:
                new_body = _set_family(new_body, family)
                attrs = _with_bold(attrs, bold)
                bump("alert weight re-levelled")

        # Plain bold is only ever real on level 1, whose family has a Bold
        # face. Below that the base families are Light and ExtraLight, so
        # PowerPoint would synthesise it; use the level's alert weight instead,
        # keeping whatever colour the run has.
        if ' b="1"' in attrs and tag == "rPr" and not bold:
            attrs = _with_bold(attrs, False)
            new_body = _set_family(new_body, family)
            bump("bold made alert weight")

        # Web-paste residue. Naming the body family explicitly blocks the
        # lighter weights the deeper levels inherit; the hard colours hide the
        # level colours and the theme's link colour.
        pat = r'<a:(?:latin|ea|cs) typeface="' + re.escape(BODY_FAMILY) + r'"[^>]*/>'
        if re.search(pat, new_body):
            new_body = re.sub(pat, "", new_body)
            bump("explicit body family dropped")
        for colour in PASTED_TEXT_COLOURS:
            fill = f'<a:solidFill><a:srgbClr val="{colour}"/></a:solidFill>'
            if fill in new_body:
                new_body = new_body.replace(fill, "")
                bump(f"{colour} colour dropped")
        if "<a:effectLst/>" in new_body:
            new_body = new_body.replace("<a:effectLst/>", "")

        if new_body != body:
            new_body = _reorder(new_body)
        return f"<a:{tag}{attrs}>{new_body}</a:{tag}>"

    # The same for runs whose properties are a self-closing tag.
    def fix_empty(match: re.Match[str]) -> str:
        attrs = match.group(1)
        family, bold = alert_spec(level)
        if ' b="1"' not in attrs or bold:
            return match.group(0)      # level 1 plain bold is the real Bold face
        bump("bold made alert weight")
        return f'<a:rPr{_with_bold(attrs, False)}><a:latin typeface="{family}"/></a:rPr>'

    # Paragraph by paragraph, so the alert weight can follow the outline level.
    def paragraph(match: re.Match[str]) -> str:
        nonlocal level
        lvl = re.search(r'<a:pPr[^>]*lvl="(\d)"', match.group(0))
        level = int(lvl.group(1)) + 1 if lvl else 1
        text = RUN.sub(fix, match.group(0))
        return EMPTY_RUN.sub(fix_empty, text)

    xml = PARA.sub(paragraph, xml)

    # Old-theme bullets ("o" in Courier New, Wingdings squares) set on the
    # paragraphs override the template's dash and square. buNone is a
    # deliberate choice and stays.
    def bullets(match: re.Match[str]) -> str:
        attrs, body = match.group(1), match.group(2)
        new_body = re.sub(r"<a:buFont [^>]*/>|<a:buChar [^>]*/>|<a:buClr>.*?</a:buClr>|<a:buSzPct [^>]*/>", "", body, flags=re.S)
        if new_body != body:
            bump("bullet override dropped")
        return f"<a:pPr{attrs}/>" if not new_body.strip() else f"<a:pPr{attrs}>{new_body}</a:pPr>"

    xml = PARA_PROPS.sub(bullets, xml)

    # Explicit point sizes on runs in the body placeholder override the
    # master's 28/24/20 ladder, and any text typed beside them inherits the
    # override, so levels drift out of step. Only the body placeholder is
    # touched: sizes in free text boxes and tables are deliberate. A stale
    # autofit scale is reset with them; PowerPoint recomputes it on edit.
    def body_sizes(match: re.Match[str]) -> str:
        block = match.group(0)
        new_block, n = re.subn(r'(<a:(?:rPr|endParaRPr)\b[^>]*?) sz="\d+"', r"\1", block)
        if n:
            bump("explicit size dropped")
        new_block = re.sub(r"<a:normAutofit [^>]*/>", "<a:normAutofit/>", new_block)
        return new_block

    xml = BODY_PLACEHOLDER.sub(body_sizes, xml)
    return xml, counts


_ORDER = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst", "effectDag",
          "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick",
          "hlinkMouseOver", "rtl", "extLst"]


def _reorder(body: str) -> str:
    """Sort a run's property children into schema order."""
    chunks = re.findall(r"<a:\w+(?:\s[^>]*)?(?:/>|>.*?</a:\w+>)", body, re.S)
    if "".join(chunks) != body:
        return body
    keyed = sorted(chunks, key=lambda c: _ORDER.index(re.match(r"<a:(\w+)", c).group(1))
                   if re.match(r"<a:(\w+)", c).group(1) in _ORDER else 99)
    return "".join(keyed)


# --------------------------------------------------------------------------
# Layouts
# --------------------------------------------------------------------------
def _layout_name(xml: str) -> str:
    m = re.search(r'<p:cSld name="([^"]*)"', xml)
    return m.group(1) if m else ""


def _shape_names(xml: str) -> set[str]:
    return set(re.findall(r'<p:cNvPr id="\d+" name="([^"]*)"', xml))


def template_layouts(template: Path) -> dict[str, set[str]]:
    """Layout name -> names of every shape the template's layout carries."""
    z = zipfile.ZipFile(template)
    out: dict[str, set[str]] = {}
    for name in z.namelist():
        if LAYOUT_PART.match(name):
            xml = z.read(name).decode("utf-8")
            out[_layout_name(xml)] = _shape_names(xml)
    return out


def strip_merged_shapes(xml: str, keep: set[str]) -> tuple[str, list[str], list[str]]:
    """Drop non-placeholder shapes the template's layout does not have.

    Returns the new XML, the names removed, and the r:embed ids they used.
    """
    removed: list[str] = []
    embeds: list[str] = []

    def drop(match: re.Match[str]) -> str:
        block = match.group(0)
        head = block[:block.find("</p:nv")] if "</p:nv" in block else block[:400]
        if "<p:ph" in head:
            return block
        name = re.search(r'<p:cNvPr id="\d+" name="([^"]*)"', block)
        if not name or name.group(1) in keep:
            return block
        removed.append(name.group(1))
        embeds.extend(re.findall(r'r:embed="([^"]+)"', block))
        return ""

    tree = re.search(r"<p:spTree>.*</p:spTree>", xml, re.S)
    if not tree:
        return xml, removed, embeds
    new_tree = SHAPE.sub(drop, tree.group(0))
    return xml.replace(tree.group(0), new_tree), removed, embeds


# --------------------------------------------------------------------------
# Package
# --------------------------------------------------------------------------
def _rels_path(part: str) -> str:
    d, f = posixpath.split(part)
    return f"{d}/_rels/{f}.rels"


def _targets(rels_xml: str, base: str) -> set[str]:
    out = set()
    for m in re.finditer(r'<Relationship [^>]*Target="([^"]+)"[^>]*/>', rels_xml):
        if 'TargetMode="External"' in m.group(0):
            continue
        out.add(posixpath.normpath(posixpath.join(base, m.group(1))))
    return out


def process(deck: Path, check_only: bool, template: Path | None) -> dict[str, int]:
    zin = zipfile.ZipFile(deck)
    parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    infos = {i.filename: i for i in zin.infolist()}
    zin.close()
    totals: dict[str, int] = {}
    changed: set[str] = set()
    deleted: set[str] = set()

    def bump(key: str, n: int = 1) -> None:
        totals[key] = totals.get(key, 0) + n

    def text(name: str) -> str:
        return parts[name].decode("utf-8")

    def put(name: str, s: str) -> None:
        parts[name] = s.encode("utf-8")
        changed.add(name)

    # -- 1. slide runs -------------------------------------------------------
    for name in list(parts):
        if SLIDE_PART.match(name):
            new_text, counts = rewrite_runs(text(name))
            if counts:
                for k, v in counts.items():
                    bump(k, v)
                put(name, new_text)

    if template is not None:
        wanted = template_layouts(template)
        master = "ppt/slideMasters/slideMaster1.xml"
        master_rels = _rels_path(master)

        # -- 2. shapes PowerPoint merged into same-named layouts ---------------
        for name in list(parts):
            if not LAYOUT_PART.match(name):
                continue
            xml = text(name)
            lname = _layout_name(xml)
            if lname not in wanted:
                continue
            new_xml, removed, embeds = strip_merged_shapes(xml, wanted[lname])
            if removed:
                put(name, new_xml)
                bump("merged layout shape removed", len(removed))
                rels_name = _rels_path(name)
                rels = text(rels_name)
                for rid in embeds:
                    rels = re.sub(r'<Relationship Id="%s"[^>]*/>' % re.escape(rid), "", rels)
                put(rels_name, rels)

        # -- 3. unused layouts the template does not define ---------------------
        used_layouts: set[str] = set()
        for name in parts:
            if SLIDE_PART.match(name):
                used_layouts |= {t for t in _targets(text(_rels_path(name)), posixpath.dirname(name))
                                 if LAYOUT_PART.match(t)}
        for name in list(parts):
            if not LAYOUT_PART.match(name) or name in used_layouts:
                continue
            if _layout_name(text(name)) in wanted:
                continue
            mrels = text(master_rels)
            rel = re.search(r'<Relationship Id="(rId\d+)"[^>]*Target="\.\./slideLayouts/%s"[^>]*/>'
                            % re.escape(posixpath.basename(name)), mrels)
            if not rel:
                continue
            put(master_rels, mrels.replace(rel.group(0), ""))
            put(master, re.sub(r'<p:sldLayoutId id="\d+" r:id="%s"/>' % rel.group(1), "", text(master)))
            deleted.add(name)
            deleted.add(_rels_path(name))
            bump("unused layout removed")

        # -- 4. media nothing references any more ---------------------------------
        referenced: set[str] = set()
        for name in parts:
            if name.endswith(".rels") and name not in deleted:
                src = name.replace("/_rels/", "/")[:-5]
                if src in deleted:
                    continue
                referenced |= _targets(text(name), posixpath.dirname(src))
        for name in list(parts):
            if name.startswith("ppt/media/") and name not in referenced:
                deleted.add(name)
                bump("orphaned media removed")

        # -- 5. content types for deleted parts ----------------------------------
        ct = text("[Content_Types].xml")
        for name in deleted:
            ct = re.sub(r'<Override PartName="/%s"[^>]*/>' % re.escape(name), "", ct)
        if ct != text("[Content_Types].xml"):
            put("[Content_Types].xml", ct)

    if (changed or deleted) and not check_only:
        tmp = deck.with_suffix(".pptx.tmp")
        with zipfile.ZipFile(tmp, "w") as zout:
            for name, info in infos.items():
                if name in deleted:
                    continue
                new = zipfile.ZipInfo(name, date_time=info.date_time)
                new.compress_type = info.compress_type
                new.external_attr = info.external_attr
                zout.writestr(new, parts[name])
        shutil.move(tmp, deck)
    return totals


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("decks", nargs="+", type=Path)
    parser.add_argument("--template", type=Path, help="the .potx that was applied; enables the layout cleanup")
    parser.add_argument("--check", action="store_true", help="report without writing")
    args = parser.parse_args(argv)
    for deck in args.decks:
        totals = process(deck, args.check, args.template)
        print(f"{deck.name}:" if totals else f"{deck.name}: nothing to change")
        for key, value in sorted(totals.items()):
            print(f"    {value:4d}  {key}")
    if args.check:
        print("(--check: nothing written)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
