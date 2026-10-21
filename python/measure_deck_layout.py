"""Ask PowerPoint how a deck's text lays out: lines per paragraph, frames that
overflow, and optionally a PNG of every slide.

The house style (STYLE.md) wants one line per bullet and no
shrink-to-fit, and only PowerPoint knows where a line wraps, so this opens a
scratch copy of the deck in a hidden PowerPoint through COM, reads each
paragraph's laid-out line count and each text frame's height against its
shape, and quits. Windows with PowerPoint only; the deck itself is never
opened, so it can be open in PowerPoint meanwhile.

    python python/measure_deck_layout.py candidate.pptx                  # every text paragraph, as TSV
    python python/measure_deck_layout.py candidate.pptx --problems       # only bullets over one line or frames that overflow
    python python/measure_deck_layout.py candidate.pptx --png qa         # also qa/slide-NN.png renders

Columns: slide, kind (title, body, subtitle, footer, slide number, date,
placeholder, text box), shape, paragraph, level, lines, frame ("fits" or
"OVERFLOW bound=... avail=..." in points, then the frame's width and height),
text. With --problems the exit status is 1 when anything is listed, so a
script can stop on it.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

POWERSHELL = r'''
param([string]$Deck, [string]$OutDir, [string]$Tsv, [switch]$Png)
$ErrorActionPreference = "Stop"
$rows = New-Object System.Collections.Generic.List[string]
$app = New-Object -ComObject PowerPoint.Application
try {
    $pres = $app.Presentations.Open($Deck, -1, 0, 0)     # ReadOnly, not Untitled, no window
    $rows.Add("slide`tkind`tshape`tpara`tlevel`tlines`tframe`ttext")
    foreach ($slide in $pres.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -ne -1) { continue }
            $tr = $shape.TextFrame.TextRange
            if ($tr.Length -eq 0) { continue }
            $kind = "text box"
            if ($shape.Type -eq 14) {
                $kind = switch ($shape.PlaceholderFormat.Type) {
                    1 { "title" } 3 { "title" } 2 { "body" } 7 { "body" } 4 { "subtitle" }
                    13 { "slide number" } 15 { "footer" } 16 { "date" } default { "placeholder" }
                }
            }
            $tf2 = $shape.TextFrame2
            $avail = $shape.Height - $tf2.MarginTop - $tf2.MarginBottom
            $bound = $tf2.TextRange.BoundHeight
            $frame = "fits"
            if ($bound -gt $avail + 0.5) {
                $frame = "OVERFLOW bound=" + [math]::Round($bound, 1) + " avail=" + [math]::Round($avail, 1)
            }
            $frame += " w=" + [math]::Round($shape.Width, 1) + " h=" + [math]::Round($shape.Height, 1)
            $n = $tr.Paragraphs().Count
            for ($i = 1; $i -le $n; $i++) {
                $p = $tr.Paragraphs($i)
                $t = $p.Text.TrimEnd([char]13, [char]10).Replace([string][char]11, "<br>").Replace("`t", " ")
                $lc = 0
                if ($t.Length -gt 0) { $lc = $p.Lines().Count }
                $rows.Add("$($slide.SlideIndex)`t$kind`t$($shape.Name)`t$i`t$($p.IndentLevel)`t$lc`t$frame`t$t")
            }
        }
        if ($Png) {
            $slide.Export((Join-Path $OutDir ("slide-{0:D2}.png" -f $slide.SlideIndex)), "PNG", 1600, 900)
        }
    }
    $pres.Close()
}
finally {
    $app.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
[System.IO.File]::WriteAllLines($Tsv, $rows, (New-Object System.Text.UTF8Encoding $false))
'''

PROBLEM_KINDS = {"body", "text box", "placeholder"}


def measure(deck: Path, png_dir: Path | None) -> list[list[str]]:
    work = Path(tempfile.mkdtemp(prefix="measure-deck-"))
    try:
        scratch = work / deck.name
        shutil.copyfile(deck, scratch)
        script = work / "measure.ps1"
        script.write_text(POWERSHELL, encoding="utf-8")
        tsv = work / "measure.tsv"
        if png_dir:
            png_dir.mkdir(parents=True, exist_ok=True)
        command = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
                   "-Deck", str(scratch), "-OutDir", str(png_dir.resolve() if png_dir else work), "-Tsv", str(tsv)]
        if png_dir:
            command.append("-Png")
        result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace")
        if result.returncode or not tsv.exists():
            raise RuntimeError(f"PowerPoint measurement failed:\n{result.stdout}\n{result.stderr}")
        # split on newlines only: str.splitlines() would also split at the
        # vertical tabs PowerPoint uses for line breaks inside a paragraph
        return [line.split("\t") for line in tsv.read_text(encoding="utf-8").rstrip("\n").split("\n")]
    finally:
        shutil.rmtree(work, ignore_errors=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("deck", type=Path)
    parser.add_argument("--problems", action="store_true",
                        help="list only body text over one line or in an overflowing frame; exit 1 if any")
    parser.add_argument("--png", type=Path, metavar="DIR", help="export a PNG of every slide into DIR")
    parser.add_argument("--slides", type=int, nargs="*", default=[], help="only these slide numbers")
    args = parser.parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")
    if os.name != "nt":
        print("measure_deck_layout.py needs Windows with PowerPoint installed", file=sys.stderr)
        return 2
    header, *rows = measure(args.deck, args.png)
    if args.slides:
        rows = [r for r in rows if int(r[0]) in set(args.slides)]
    if args.problems:
        rows = [r for r in rows if r[1] in PROBLEM_KINDS and (int(r[5]) > 1 or r[6].startswith("OVERFLOW"))]
    print("\t".join(header))
    for row in rows:
        print("\t".join(row))
    if args.problems:
        print(f"{len(rows)} problem paragraph(s)" if rows else "no problems: every body paragraph is one line and every frame fits")
        return 1 if rows else 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
