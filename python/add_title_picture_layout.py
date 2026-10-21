"""Add the "Title Slide with Picture" layout and put lecture title slides on it.

The teaching template's Title Slide centres the lecture title and number. This
layout keeps that look but moves them into the left two-thirds and reserves a
square picture placeholder on the right, the arrangement Lecture 2's title
slide arrived at by hand, so every lecture can carry a colourful image beside
its title. In PowerPoint the empty placeholder shows "Click icon to add
picture"; in a slideshow and in the published PDF it draws nothing.

`--layout-only` adds the layout to a file (the template, or a deck) without
touching its slides. Otherwise each deck's first slide is moved onto the new
layout: explicit positions on its title and subtitle are dropped so they take
the layout's, and a loose picture on the slide is moved into the placeholder,
cropped to the square. Running again is harmless.

Edits the packed XML directly and rezips with each part's original compression
(see retheme_fonts.py for why not a python-pptx round trip).

    python python/add_title_picture_layout.py --layout-only templates/Presentation-template-Inter-teaching.potx
    python python/add_title_picture_layout.py decks/*.pptx
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

LAYOUT_NAME = "Title Slide with Picture"
SOURCE_LAYOUT = "Title Slide"
PIC_IDX = "13"                       # placeholder index unused by the template's layouts
EMU = 914400
TITLE_BOX = (int(0.5 * EMU), int(2.0 * EMU), int(8.4 * EMU), int(2.6 * EMU))
SUBTITLE_BOX = (int(0.5 * EMU), int(4.8 * EMU), int(8.4 * EMU), int(0.8 * EMU))
PICTURE_BOX = (int(9.2 * EMU), int(2.05 * EMU), int(3.5 * EMU), int(3.5 * EMU))
CREATION_ID = "3415926535"

R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RT_LAYOUT = R + "/slideLayout"
RT_MASTER = R + "/slideMaster"
CT_LAYOUT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"

PICTURE_PLACEHOLDER = (
    '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Picture Placeholder 3"/>'
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
    f'<p:nvPr><p:ph type="pic" sz="half" idx="{PIC_IDX}"/></p:nvPr></p:nvSpPr>'
    '<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>'
    "<a:t>Click icon to add picture</a:t></a:r></a:p></p:txBody></p:sp>"
)


def xfrm(box) -> str:
    x, y, cx, cy = box
    return f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'


class Package:
    """A pptx/potx opened for string edits, rezipped with compression kept."""

    def __init__(self, path: Path):
        self.path = path
        self.zin = zipfile.ZipFile(path)
        self.names = self.zin.namelist()
        self.edited: dict[str, bytes] = {}
        self.added: dict[str, bytes] = {}

    def read_bytes(self, name: str) -> bytes:
        if name in self.edited:
            return self.edited[name]
        if name in self.added:
            return self.added[name]
        return self.zin.read(name)

    def read(self, name: str) -> str:
        return self.read_bytes(name).decode("utf-8")

    def write(self, name: str, text: str) -> None:
        (self.added if name not in self.names else self.edited)[name] = text.encode("utf-8")

    def rels_name(self, part: str) -> str:
        folder, name = part.rsplit("/", 1)
        return f"{folder}/_rels/{name}.rels"

    def relationship(self, part: str, reltype: str) -> tuple[str, str] | None:
        rels = self.read(self.rels_name(part))
        m = re.search(rf'<Relationship [^>]*?Id="([^"]+)"[^>]*?Type="{re.escape(reltype)}"[^>]*?Target="([^"]+)"', rels)
        if m is None:
            m = re.search(rf'<Relationship [^>]*?Type="{re.escape(reltype)}"[^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)
            return (m.group(2), m.group(1)) if m else None
        return (m.group(1), m.group(2))

    def add_relationship(self, part: str, reltype: str, target: str) -> str:
        name = self.rels_name(part)
        rels = self.read(name)
        rid = "rId" + str(max([int(m) for m in re.findall(r'Id="rId(\d+)"', rels)] or [0]) + 1)
        rels = rels.replace("</Relationships>",
                            f'<Relationship Id="{rid}" Type="{reltype}" Target="{target}"/></Relationships>')
        self.write(name, rels)
        return rid

    def write_bytes(self, name: str, data: bytes) -> None:
        (self.added if name not in self.names else self.edited)[name] = data

    def remove_relationship(self, part: str, rid: str) -> None:
        name = self.rels_name(part)
        rels = self.read(name)
        self.write(name, re.sub(rf'<Relationship [^>]*?Id="{re.escape(rid)}"[^>]*?/>', "", rels, count=1))

    def add_default(self, extension: str, content_type: str) -> None:
        types = self.read("[Content_Types].xml")
        if f'Extension="{extension}"' not in types:
            types = types.replace("<Default ", f'<Default Extension="{extension}" ContentType="{content_type}"/><Default ', 1)
            self.write("[Content_Types].xml", types)

    def add_override(self, part: str, content_type: str) -> None:
        types = self.read("[Content_Types].xml")
        if f'PartName="/{part}"' not in types:
            types = types.replace("</Types>", f'<Override PartName="/{part}" ContentType="{content_type}"/></Types>')
            self.write("[Content_Types].xml", types)

    def save(self) -> None:
        # An edit that writes back what was already there is not an edit. Saying
        # so here keeps a rewrite that changes nothing from touching the file at
        # all: the stamper re-generates its shapes on every build, and a deck
        # rewritten with identical bytes still went back to Dropbox through
        # rclone --update and came down again as a new revision, which is how a
        # deck open in PowerPoint ended up with conflicted copies.
        self.edited = {name: data for name, data in self.edited.items()
                       if data != self.zin.read(name)}
        if not self.edited and not self.added:
            self.zin.close()
            return
        tmp = self.path.with_suffix(self.path.suffix + ".tmp")
        with zipfile.ZipFile(tmp, "w") as zout:
            for item in self.zin.infolist():
                info = zipfile.ZipInfo(item.filename, date_time=item.date_time)
                info.compress_type = item.compress_type
                info.external_attr = item.external_attr
                zout.writestr(info, self.edited[item.filename] if item.filename in self.edited
                              else self.zin.read(item.filename))
            for name, data in self.added.items():
                zout.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
        self.zin.close()
        shutil.move(tmp, self.path)


def layouts(pkg: Package) -> dict[str, str]:
    """Layout name -> part name."""
    found = {}
    for name in pkg.names + list(pkg.added):
        if re.match(r"ppt/slideLayouts/slideLayout\d+[.]xml$", name):
            m = re.search(r'<p:cSld name="([^"]*)"', pkg.read(name))
            if m:
                found[m.group(1)] = name
    return found


def replace_xfrm(shape_xml: str, box) -> str:
    """Give the placeholder shape exactly one xfrm, the given box."""
    if re.search(r"<a:xfrm[ >]", shape_xml):
        return re.sub(r"<a:xfrm[^>]*>.*?</a:xfrm>", xfrm(box), shape_xml, count=1, flags=re.S)
    if "<p:spPr/>" in shape_xml:
        return shape_xml.replace("<p:spPr/>", f"<p:spPr>{xfrm(box)}</p:spPr>", 1)
    return re.sub(r"(<p:spPr(?: [^>]*)?>)", rf"\g<1>{xfrm(box)}", shape_xml, count=1)


def placeholder_shape(xml: str, ph_pattern: str) -> re.Match | None:
    """The <p:sp> whose placeholder tag matches ph_pattern."""
    for m in re.finditer(r"<p:sp>.*?</p:sp>", xml, re.S):
        if re.search(ph_pattern, m.group(0)):
            return m
    return None


def add_layout(pkg: Package) -> str:
    existing = layouts(pkg)
    if LAYOUT_NAME in existing:
        return existing[LAYOUT_NAME]
    source = existing[SOURCE_LAYOUT]
    xml = pkg.read(source)

    xml = xml.replace(f'<p:cSld name="{SOURCE_LAYOUT}">', f'<p:cSld name="{LAYOUT_NAME}">', 1)
    xml = re.sub(r'(<p14:creationId [^>]*val=")\d+(")', rf"\g<1>{CREATION_ID}\2", xml)
    title = placeholder_shape(xml, r'<p:ph type="ctrTitle"')
    xml = xml.replace(title.group(0), replace_xfrm(title.group(0), TITLE_BOX), 1)
    subtitle = placeholder_shape(xml, r'<p:ph type="subTitle"')
    xml = xml.replace(subtitle.group(0), replace_xfrm(subtitle.group(0), SUBTITLE_BOX), 1)
    x, y, cx, cy = PICTURE_BOX
    picture = PICTURE_PLACEHOLDER.format(x=x, y=y, cx=cx, cy=cy)
    xml = xml.replace("</p:spTree>", picture + "</p:spTree>", 1)

    numbers = [int(re.search(r"(\d+)", n.rsplit("/", 1)[1]).group(1)) for n in existing.values()]
    part = f"ppt/slideLayouts/slideLayout{max(numbers) + 1}.xml"
    pkg.write(part, xml)
    master_rel = pkg.relationship(source, RT_MASTER)
    pkg.write(pkg.rels_name(part),
              '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              f'<Relationship Id="rId1" Type="{RT_MASTER}" Target="{master_rel[1]}"/></Relationships>')
    pkg.add_override(part, CT_LAYOUT)

    master = "ppt/slideMasters/" + master_rel[1].rsplit("/", 1)[1]
    rid = pkg.add_relationship(master, RT_LAYOUT, "../slideLayouts/" + part.rsplit("/", 1)[1])
    master_xml = pkg.read(master)
    ids = [int(i) for i in re.findall(r'<p:sldLayoutId id="(\d+)"', master_xml)]
    entry = f'<p:sldLayoutId id="{max(ids) + 1}" r:id="{rid}"/>'
    # Right after the Title Slide, so it sits second in the layout gallery.
    rels = pkg.read(pkg.rels_name(master))
    source_rid = re.search(rf'<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="../slideLayouts/{source.rsplit("/", 1)[1]}"', rels)
    if source_rid is None:
        source_rid = re.search(rf'<Relationship [^>]*?Target="../slideLayouts/{source.rsplit("/", 1)[1]}"[^>]*?Id="([^"]+)"', rels)
    anchor = re.search(rf'<p:sldLayoutId id="\d+" r:id="{source_rid.group(1)}"/>', master_xml)
    if anchor:
        master_xml = master_xml.replace(anchor.group(0), anchor.group(0) + entry, 1)
    else:
        master_xml = master_xml.replace("</p:sldLayoutIdLst>", entry + "</p:sldLayoutIdLst>", 1)
    pkg.write(master, master_xml)
    return part


def first_slide(pkg: Package) -> str:
    presentation = pkg.read("ppt/presentation.xml")
    rid = re.search(r'<p:sldId [^>]*r:id="([^"]+)"', presentation).group(1)
    rels = pkg.read("ppt/_rels/presentation.xml.rels")
    m = re.search(rf'<Relationship [^>]*?Id="{rid}"[^>]*?Target="([^"]+)"', rels) or \
        re.search(rf'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="{rid}"', rels)
    return "ppt/" + m.group(1).removeprefix("/ppt/")


def fit_crop(pic_xml: str, box) -> str:
    """Crop the picture (on top of any crop it has) so it fills the box with
    its proportions kept, then give it the box as its frame."""
    m = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', pic_xml)
    if m is None:
        return pic_xml
    shown_w, shown_h = int(m.group(1)), int(m.group(2))
    src = re.search(r"<a:srcRect([^>]*)/>", pic_xml)
    crop = {k: int(v) / 100000 for k, v in re.findall(r'\b([lrtb])="(-?\d+)"', src.group(1) if src else "")}
    l, r, t, b = (crop.get(k, 0.0) for k in "lrtb")
    if shown_w <= 0 or shown_h <= 0:
        return pic_xml
    shown_aspect = shown_w / shown_h
    box_aspect = box[2] / box[3]
    if shown_aspect > box_aspect:            # too wide: trim the sides
        keep = box_aspect / shown_aspect
        trim = (1 - keep) / 2 * (1 - l - r)
        l, r = l + trim, r + trim
    else:                                    # too tall: trim top and bottom
        keep = shown_aspect / box_aspect
        trim = (1 - keep) / 2 * (1 - t - b)
        t, b = t + trim, b + trim
    rect = "".join(f' {k}="{round(v * 100000)}"' for k, v in zip("lrtb", (l, r, t, b)) if round(v * 100000))
    new_src = f"<a:srcRect{rect}/>" if rect else ""
    if src:
        pic_xml = pic_xml.replace(src.group(0), new_src, 1)
    elif new_src:
        pic_xml = re.sub(r"(</a:blip>|<a:blip [^>]*/>)", rf"\1{new_src}", pic_xml, count=1)
    return re.sub(r"<a:xfrm[^>]*>.*?</a:xfrm>", xfrm(box), pic_xml, count=1, flags=re.S)


def apply_to_title_slide(pkg: Package, layout_part: str) -> str:
    slide = first_slide(pkg)
    xml = pkg.read(slide)
    if '<p:ph type="ctrTitle"' not in xml:
        return "first slide has no centred title; left alone"
    rels_name = pkg.rels_name(slide)
    rels = pkg.read(rels_name)
    target = "../slideLayouts/" + layout_part.rsplit("/", 1)[1]
    rels, n = re.subn(rf'(<Relationship [^>]*?Type="{re.escape(RT_LAYOUT)}"[^>]*?Target=")[^"]+(")',
                      rf"\g<1>{target}\2", rels)
    if n == 0:
        rels, n = re.subn(rf'(<Relationship [^>]*?Target=")[^"]+("[^>]*?Type="{re.escape(RT_LAYOUT)}")',
                          rf"\g<1>{target}\2", rels)
    pkg.write(rels_name, rels)

    # Inherit the layout's positions.
    for ph in ("ctrTitle", "subTitle"):
        m = placeholder_shape(xml, rf'<p:ph type="{ph}"')
        if m:
            shape = re.sub(r"<a:xfrm[^>]*>.*?</a:xfrm>", "", m.group(0), count=1, flags=re.S)
            shape = shape.replace("<p:spPr></p:spPr>", "<p:spPr/>")
            xml = xml.replace(m.group(0), shape, 1)

    # A loose picture becomes the placeholder's picture.
    outcome = "moved onto the layout"
    if f'idx="{PIC_IDX}"' not in xml:
        pics = re.findall(r"<p:pic>.*?</p:pic>", xml, re.S)
        loose = [p for p in pics if "<p:ph " not in p]
        if loose:
            pic = loose[0]
            new = pic.replace("<p:nvPr/>", f'<p:nvPr><p:ph type="pic" sz="half" idx="{PIC_IDX}"/></p:nvPr>', 1)
            if new == pic:
                new = re.sub(r"<p:nvPr>", f'<p:nvPr><p:ph type="pic" sz="half" idx="{PIC_IDX}"/>', pic, count=1)
            new = fit_crop(new, PICTURE_BOX)
            xml = xml.replace(pic, new, 1)
            outcome += "; picture placed in the placeholder"
    pkg.write(slide, xml)
    return outcome


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("files", nargs="+", type=Path)
    parser.add_argument("--layout-only", action="store_true", help="add the layout; leave slides alone")
    args = parser.parse_args(argv)
    for path in args.files:
        if (path.parent / f"~${path.name}").exists():
            print(f"{path.name}: open in PowerPoint; skipped")
            continue
        pkg = Package(path)
        layout = add_layout(pkg)
        outcome = "layout present"
        if not args.layout_only:
            outcome = apply_to_title_slide(pkg, layout)
        pkg.save()
        print(f"{path.name}: {outcome}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
