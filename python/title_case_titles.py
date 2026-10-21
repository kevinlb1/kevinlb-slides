"""Put each lecture deck's title (slide 1) in title case.

Title case here follows the usual editorial rule: every word capitalised
except articles, coordinating conjunctions and short prepositions inside the
title; the first and last words are always capitalised; both halves of a
hyphenated word are capitalised; words already in capitals (AI, GPU) are kept.
Only the centred title on the first slide is touched, run by run, so line
breaks and per-run formatting survive. The same package-level editing as the
other deck scripts, so nothing else in the file changes.

    python python/title_case_titles.py decks/*.pptx
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from xml.sax.saxutils import escape, unescape

sys.path.insert(0, str(Path(__file__).resolve().parent))
from add_title_picture_layout import Package, first_slide, placeholder_shape  # noqa: E402

SMALL = {"a", "an", "the", "and", "but", "or", "nor", "for", "so", "yet", "as", "at",
         "by", "in", "of", "on", "per", "to", "up", "via", "vs", "vs.", "v.", "from",
         "into", "onto", "with", "over", "than"}
WORD = re.compile(r"[A-Za-z][A-Za-z'’]*")


def cap(word: str) -> str:
    if word.isupper():
        return word
    return word[0].upper() + word[1:]


def title_case(text: str) -> str:
    words = list(WORD.finditer(text))
    if not words:
        return text
    out, last = [], 0
    for k, m in enumerate(words):
        word = m.group(0)
        first_or_last = k == 0 or k == len(words) - 1
        keep_small = word.lower() in SMALL and not first_or_last
        out.append(text[last:m.start()])
        out.append(word if keep_small else cap(word))
        last = m.end()
    out.append(text[last:])
    return "".join(out)


def title_case_runs(texts: list[str]) -> list[str]:
    """Title-case a title split across runs, judging first/last words and
    hyphen halves against the whole title, not each run."""
    joined = "".join(texts)
    whole = title_case_hyphens(title_case(joined))
    result, pos = [], 0
    for text in texts:
        result.append(whole[pos:pos + len(text)])
        pos += len(text)
    return result


def title_case_hyphens(text: str) -> str:
    return re.sub(r"-([a-z])", lambda m: "-" + m.group(1).upper(), text)


def process(path: Path) -> str:
    pkg = Package(path)
    slide = first_slide(pkg)
    xml = pkg.read(slide)
    shape = placeholder_shape(xml, r'<p:ph type="ctrTitle"')
    if shape is None:
        pkg.save()
        return "no centred title on slide 1"
    body = shape.group(0)
    texts = [unescape(t) for t in re.findall(r"<a:t>([^<]*)</a:t>", body)]
    if not texts:
        pkg.save()
        return "empty title"
    new_texts = title_case_runs(texts)
    if new_texts == texts:
        pkg.save()
        return f"already title case: {''.join(texts)!r}"
    parts = re.split(r"(<a:t>[^<]*</a:t>)", body)
    k = 0
    for i, part in enumerate(parts):
        if part.startswith("<a:t>"):
            parts[i] = f"<a:t>{escape(new_texts[k])}</a:t>"
            k += 1
    pkg.write(slide, xml.replace(body, "".join(parts), 1))
    pkg.save()
    return f"{''.join(texts)!r} -> {''.join(new_texts)!r}"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("decks", nargs="+", type=Path)
    args = parser.parse_args(argv)
    for deck in args.decks:
        if (deck.parent / f"~${deck.name}").exists():
            print(f"{deck.name}: open in PowerPoint; skipped")
            continue
        print(f"{deck.name}: {process(deck)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
