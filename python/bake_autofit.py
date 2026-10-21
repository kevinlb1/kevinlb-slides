"""Freeze PowerPoint's shrink-to-fit decisions into explicit font sizes.

A text body with "shrink text on overflow" records what PowerPoint settled on:
`<a:normAutofit fontScale="90000" lnSpcReduction="10000"/>` means the text is
drawn at 90% of its nominal sizes with line spacing reduced by 10%. LibreOffice
reads that flag but fits the text again with its own algorithm and its own
text metrics, so a title that PowerPoint kept on one line at 90% can come out
wrapped at a different size in the PDF. Multiplying the recorded scale into
every run's size, and dropping the autofit flag, makes LibreOffice draw exactly
the sizes PowerPoint chose.

Most runs carry no size of their own; it comes from the shape's list style, the
layout's placeholder, the master's placeholder or the master's text styles, in
that order, and finally the presentation's default text style. `inherited_size`
walks that chain.
"""

from __future__ import annotations

from lxml import etree
from pptx.oxml.ns import qn

DEFAULT_SIZE = 1800          # hundredths of a point: PowerPoint's last resort
# Which master text style a placeholder type inherits from.
TITLE_TYPES = {"title", "ctrTitle"}
BODY_TYPES = {"body", "subTitle", "obj", None}      # None: a bare idx placeholder


def placeholder_key(sp):
    """(type, idx) of a shape's placeholder element, or None for plain shapes."""
    ph = sp.find(qn("p:nvSpPr") + "/" + qn("p:nvPr") + "/" + qn("p:ph"))
    if ph is None:
        return None
    return ph.get("type"), ph.get("idx")


def find_placeholder(spTree, key):
    """The placeholder in a layout or master matching a slide placeholder:
    same idx when it has one, else same type."""
    kind, idx = key
    by_type = None
    for sp in spTree.iter(qn("p:sp")):
        other = placeholder_key(sp)
        if other is None:
            continue
        if idx is not None and other[1] == idx:
            return sp
        if kind is not None and other[0] == kind and by_type is None:
            by_type = sp
    return by_type


def level_size(lst_style, level: int):
    """defRPr@sz for a level in an a:lstStyle-like element, or None."""
    if lst_style is None:
        return None
    ppr = lst_style.find(qn(f"a:lvl{level + 1}pPr"))
    if ppr is None:
        return None
    defrpr = ppr.find(qn("a:defRPr"))
    return int(defrpr.get("sz")) if defrpr is not None and defrpr.get("sz") else None


def shape_list_style(sp):
    body = sp.find(qn("p:txBody"))
    return None if body is None else body.find(qn("a:lstStyle"))


def inherited_size(prs, slide, sp, level: int) -> int:
    """The size a run at this paragraph level draws at when nothing on the run
    or paragraph says otherwise."""
    size = level_size(shape_list_style(sp), level)
    if size:
        return size
    key = placeholder_key(sp)
    layout = slide.slide_layout
    master = layout.slide_master
    if key is not None:
        for part in (layout, master):
            match = find_placeholder(part._element.find(qn("p:cSld")).find(qn("p:spTree")), key)
            size = level_size(shape_list_style(match), level) if match is not None else None
            if size:
                return size
        styles = master._element.find(qn("p:txStyles"))
        if styles is not None:
            kind = key[0]
            style_tag = "p:titleStyle" if kind in TITLE_TYPES else \
                        "p:bodyStyle" if kind in BODY_TYPES else "p:otherStyle"
            size = level_size(styles.find(qn(style_tag)), level)
            if size:
                return size
    else:
        styles = master._element.find(qn("p:txStyles"))
        size = level_size(styles.find(qn("p:otherStyle")), level) if styles is not None else None
        if size:
            return size
    default = prs.part._element.find(qn("p:defaultTextStyle"))
    return level_size(default, level) or DEFAULT_SIZE


def run_properties(run):
    """The rPr of a run or field, created when missing."""
    rpr = run.find(qn("a:rPr"))
    if rpr is None:
        rpr = etree.Element(qn("a:rPr"))
        run.insert(0, rpr)
    return rpr


def bake_shape(prs, slide, sp) -> bool:
    body = sp.find(qn("p:txBody"))
    if body is None:
        return False
    bodypr = body.find(qn("a:bodyPr"))
    autofit = None if bodypr is None else bodypr.find(qn("a:normAutofit"))
    if autofit is None:
        return False
    scale = int(autofit.get("fontScale") or 100000) / 100000
    reduction = int(autofit.get("lnSpcReduction") or 0)
    if scale == 1 and reduction == 0:
        return False
    for paragraph in body.findall(qn("a:p")):
        ppr = paragraph.find(qn("a:pPr"))
        level = int(ppr.get("lvl")) if ppr is not None and ppr.get("lvl") else 0
        for run in list(paragraph.findall(qn("a:r"))) + list(paragraph.findall(qn("a:fld"))) + \
                list(paragraph.findall(qn("a:endParaRPr"))):
            rpr = run if run.tag == qn("a:endParaRPr") else run_properties(run)
            nominal = int(rpr.get("sz")) if rpr.get("sz") else None
            if nominal is None:
                para_default = None if ppr is None else ppr.find(qn("a:defRPr"))
                nominal = int(para_default.get("sz")) if para_default is not None and para_default.get("sz") \
                    else inherited_size(prs, slide, sp, level)
            rpr.set("sz", str(max(100, round(nominal * scale))))
        if reduction and (ppr is None or ppr.find(qn("a:lnSpc")) is None):
            if ppr is None:
                ppr = etree.Element(qn("a:pPr"))
                paragraph.insert(0, ppr)
            spacing = etree.SubElement(ppr, qn("a:lnSpc"))
            ppr.remove(spacing)
            ppr.insert(0, spacing)
            etree.SubElement(spacing, qn("a:spcPct")).set("val", str(100000 - reduction))
    # The sizes are now the fitted ones; nothing is left for a renderer to fit.
    bodypr.remove(autofit)
    return True


def bake_autofit(prs) -> int:
    """Bake every shrunk text body in the deck; returns how many."""
    baked = 0
    for slide in prs.slides:
        for sp in slide._element.iter(qn("p:sp")):
            baked += bake_shape(prs, slide, sp)
    return baked
