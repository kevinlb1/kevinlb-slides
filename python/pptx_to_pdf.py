"""Convert a PowerPoint deck to a PDF that looks like what PowerPoint shows.

LibreOffice does the conversion, but a deck handed to it as-is comes out
wrong in four ways, each fixed on a temporary copy before LibreOffice sees it:

- speaker notes, comments and hidden slides can leak into the PDF, so the
  copy has them removed (a slide tagged `#omit-slide` in its notes goes too);
- shrink-to-fit text is refitted with LibreOffice's own metrics, so the scale
  PowerPoint settled on is baked into explicit sizes (`bake_autofit.py`);
- equation zones have no text in LibreOffice, so PowerPoint's picture of the
  shape is drawn instead (`animation_steps.show_picture_fallbacks`);
- animations are not played, so a slide on which something disappears or
  covers other content becomes one page per click (`animation_steps.py`).
  `#animate-PDF` in a slide's notes forces that, `#static-PDF` forbids it.

Tags are read from every slide's notes except slide 1's, which is where a deck
can explain them. The fonts in `fonts/` beside this repository are made
visible to LibreOffice through a private fontconfig tree (Linux and macOS;
on Windows install them with `fonts/install-fonts.ps1`), with
`fonts/fonts.conf` supplying the fallbacks for fonts nobody has.

    python python/pptx_to_pdf.py deck.pptx                 # deck.pdf beside it
    python python/pptx_to_pdf.py deck.pptx -o out/         # into a directory
    python python/pptx_to_pdf.py deck.pptx --prepare-only  # just the cleaned copy

Needs LibreOffice (`soffice` on the PATH, or --soffice). The deck itself is
never modified.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import animation_steps  # noqa: E402
import bake_autofit  # noqa: E402

HERE = Path(__file__).resolve().parent
FONTS = HERE.parent / "fonts"

TAG_ANIMATE = "#animate-pdf"
TAG_STATIC = "#static-pdf"
TAG_OMIT = "#omit-slide"
TAG = re.compile(r"(#animate-pdf|#static-pdf|#omit-slide)\b", re.I)

# Relationship types, by their last path segment, of the slide parts that must
# never reach a PDF. LibreOffice leaves notes and comments out by default, but
# a default is not a guarantee; removing the parts is.
PRIVATE_SLIDE_PARTS = {"notesSlide", "comments"}


def notes_text(slide) -> str:
    if not slide.has_notes_slide:
        return ""
    return slide.notes_slide.notes_text_frame.text


def tags(index: int, notes: str) -> set[str]:
    """The tags in a slide's notes. Slide 1 never carries tags."""
    if index == 1:
        return set()
    return {m.group(1).lower() for m in TAG.finditer(notes)}


def strip_private_content(prs) -> tuple[int, int]:
    """Drop speaker notes, comments and hidden slides from the deck in memory.
    A part nothing refers to any more is not written on save."""
    stripped = 0
    for slide in prs.slides:
        private = [rel.rId for rel in slide.part.rels.values()
                   if rel.reltype.rsplit("/", 1)[-1] in PRIVATE_SLIDE_PARTS]
        for rId in private:
            slide.part.rels.pop(rId)
        stripped += bool(private)
    hidden = remove_slides(prs, {sldId.id for sldId in prs.slides._sldIdLst
                                 if prs.slides.get(sldId.id)._element.get("show") == "0"})
    return stripped, hidden


def remove_slides(prs, slide_ids: set[int]) -> int:
    removed = 0
    for sldId in list(prs.slides._sldIdLst):
        if sldId.id in slide_ids:
            prs.part.drop_rel(sldId.rId)
            prs.slides._sldIdLst.remove(sldId)
            removed += 1
    return removed


def expand_animations(prs, decisions: dict[int, str]) -> None:
    numbers = {slide.slide_id: number for number, slide in enumerate(prs.slides, 1)}
    for slide in list(prs.slides):
        number = numbers[slide.slide_id]
        decision = decisions.get(slide.slide_id)
        if decision == "static":
            continue
        reason = "#animate-PDF" if decision == "force" else animation_steps.reason_for_steps(prs, slide)
        if reason is None:
            continue
        pages = animation_steps.expand(prs, slide)
        if pages == 1:
            print(f"  slide {number}: {reason}, but no animation clicks to expand")
        else:
            print(f"  slide {number}: {pages} pages ({reason})")


def prepare(deck: Path, destination: Path) -> None:
    """Write the copy of `deck` that LibreOffice should convert."""
    prs = Presentation(deck)
    decisions: dict[int, str] = {}
    omitted: set[int] = set()
    for index, slide in enumerate(prs.slides, start=1):
        found = tags(index, notes_text(slide))
        if TAG_STATIC in found:
            decisions[slide.slide_id] = "static"
        elif TAG_ANIMATE in found:
            decisions[slide.slide_id] = "force"
        if TAG_OMIT in found:
            omitted.add(slide.slide_id)
            print(f"  slide {index}: #omit-slide, kept out of the PDF")

    # Numbers first, while every slide is still in place and hidden slides
    # still count, as they do in PowerPoint's own numbering.
    animation_steps.literalise_slide_numbers(prs)
    stripped, hidden = strip_private_content(prs)
    hidden += remove_slides(prs, omitted)
    if stripped or hidden:
        print(f"  kept out of the PDF: notes or comments on {stripped} slide(s), "
              f"{hidden} hidden or omitted slide(s)")
    baked = bake_autofit.bake_autofit(prs)
    if baked:
        print(f"  froze PowerPoint's shrink-to-fit sizes on {baked} text box(es)")
    fallbacks = animation_steps.show_picture_fallbacks(prs)
    if fallbacks:
        print(f"  drew PowerPoint's picture of the text on {fallbacks} equation box(es)")
    expand_animations(prs, decisions)
    destination.parent.mkdir(parents=True, exist_ok=True)
    prs.save(destination)


def soffice(explicit: str | None) -> str | None:
    if explicit:
        return explicit
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                      "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if Path(candidate).exists():
            return candidate
    return None


def font_environment(fonts: Path, scratch: Path) -> dict[str, str]:
    """An environment in which fontconfig sees the repository's fonts and
    fallback rules, without installing anything. Windows LibreOffice does not
    use fontconfig, so there this changes nothing."""
    env = dict(os.environ)
    if os.name == "nt" or not fonts.is_dir():
        return env
    xdg = scratch / "xdg"
    (xdg / "fonts").mkdir(parents=True)
    for ttf in fonts.glob("*.ttf"):
        shutil.copyfile(ttf, xdg / "fonts" / ttf.name)
    conf = fonts / "fonts.conf"
    if conf.is_file():
        (xdg / "fontconfig").mkdir()
        shutil.copyfile(conf, xdg / "fontconfig" / "fonts.conf")
    env["XDG_DATA_HOME"] = str(xdg)
    env["XDG_CONFIG_HOME"] = str(xdg)
    if shutil.which("fc-cache"):
        subprocess.run(["fc-cache", "-f", str(xdg / "fonts")], env=env,
                       capture_output=True)
    return env


def to_pdf(binary: str, deck: Path, out_dir: Path, env: dict[str, str]) -> Path | None:
    # A private profile keeps repeated runs from fighting over the default
    # one, which is a common cause of silent hangs.
    with tempfile.TemporaryDirectory() as profile:
        result = subprocess.run(
            [binary, "--headless", "--norestore", "--invisible",
             f"-env:UserInstallation=file://{Path(profile).as_posix()}",
             "--convert-to", "pdf", "--outdir", str(out_dir), str(deck)],
            capture_output=True, text=True, timeout=600, env=env)
    pdf = out_dir / f"{deck.stem}.pdf"
    if result.returncode != 0 or not pdf.exists():
        print(f"  LibreOffice failed for {deck.name}: "
              f"{(result.stderr or result.stdout).strip()[:300]}", file=sys.stderr)
        return None
    return pdf


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("decks", nargs="+", type=Path)
    parser.add_argument("-o", "--out", type=Path,
                        help="directory for the PDFs (default: beside each deck)")
    parser.add_argument("--prepare-only", action="store_true",
                        help="write the cleaned .pptx copy and stop, without LibreOffice")
    parser.add_argument("--soffice", help="path to the LibreOffice binary")
    parser.add_argument("--fonts", type=Path, default=FONTS,
                        help=f"font directory to expose to LibreOffice (default: {FONTS})")
    args = parser.parse_args(argv)

    binary = None if args.prepare_only else soffice(args.soffice)
    if not args.prepare_only and binary is None:
        print("LibreOffice not found: put soffice on the PATH or pass --soffice", file=sys.stderr)
        return 2

    failures = 0
    with tempfile.TemporaryDirectory() as scratch_name:
        scratch = Path(scratch_name)
        env = font_environment(args.fonts, scratch) if binary else dict(os.environ)
        for deck in args.decks:
            out_dir = args.out or deck.parent
            out_dir.mkdir(parents=True, exist_ok=True)
            print(f"{deck.name}:")
            if args.prepare_only:
                target = out_dir / f"{deck.stem}-for-pdf.pptx"
                prepare(deck, target)
                print(f"  -> {target}")
                continue
            work = scratch / deck.stem
            work.mkdir()
            prepared = work / deck.name
            prepare(deck, prepared)
            pdf = to_pdf(binary, prepared, work, env)
            if pdf is None:
                failures += 1
                continue
            final = out_dir / pdf.name
            shutil.copyfile(pdf, final)
            print(f"  -> {final}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
