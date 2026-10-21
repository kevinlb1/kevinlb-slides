# kevinlb-slides

PowerPoint templates, an emphasis add-in, the fonts they need, a house style
for slide text, and the scripts that edit, measure and convert decks. Extracted
from the tooling built for UBC's AI 100 course so that it can be used for any
deck.

| Directory | What is there |
|---|---|
| `templates/` | The `.potx` templates: the same design in Inter and in Hanken Grotesk, plus a teaching variant with footers and an activity layout |
| `macros/` | The AlertText add-in (prebuilt `.ppam`, VBA source, ribbon XML, installers) |
| `fonts/` | Every font the templates name, as static per-weight files, with an installer and fontconfig fallbacks |
| `python/` | Scripts that dump and apply slide text, tidy emphasis, measure layout, add slides, move decks onto the template and convert to PDF |
| `STYLE.md` | The house style for slide text, and the workflow for bringing a deck into line |
| `AGENTS.md` | Instructions for an AI coding agent asked to edit a deck (`CLAUDE.md` is a copy) |

## Setting up a machine

1. **Fonts.** Close PowerPoint, then right-click `fonts/install-fonts.ps1` and
   choose Run with PowerShell (Windows; it asks for Administrator rights), or
   open the `.ttf` files and click Install (macOS), or copy them into
   `~/.local/share/fonts` and run `fc-cache -f` (Linux). `fonts/README.md`
   explains why every weight is a separate family and what each family replaces.
2. **Add-in.** On Windows run `macros/deploy-to-new-machine.ps1` once; on a Mac
   follow `macros/README-mac.md`. PowerPoint's Home tab then ends with a group
   holding Alert, Bolder, Smaller, Watch and Fix Alerts.
3. **Python**, for the scripts: `uv sync` (or `pip install python-pptx lxml`).
   Measuring layout needs Windows with PowerPoint; converting to PDF needs
   LibreOffice; everything else runs anywhere.

## The design

A full-width navy title bar with a white ExtraBold title, dark-navy body text
and lighter weights on each deeper outline level, so hierarchy is carried by
weight and tint rather than size alone. Weight is always a family name (Inter
Light, Inter SemiBold), never PowerPoint's bold flag, because PowerPoint knows
only four faces per family and would synthesise the rest.

Emphasis is three looks written by the add-in and never by hand: **Alert 1**
(three weights up, in Accent 1 blue) for the key term of a top-level bullet,
**Alert 2** (two weights up, in Accent 2 teal) for a quieter second kind of
emphasis and for sub-bullets, and **Bolder** (two weights up, colour unchanged)
where bold is wanted without a colour. Both accents lighten with the outline
level so an alert sits at about the lightness of the text around it. `STYLE.md`
has the exact weights and colours per level, and the rules for the words.

## The scripts

All of them take deck paths as arguments, refuse a deck PowerPoint has open
(a `~$` lock file beside it), and edit the packed XML part by part so that
everything they do not touch stays byte for byte as PowerPoint saved it.
`python/add_title_picture_layout.py` holds the `Package` helper the others
share.

| Script | Does |
|---|---|
| `deck_text.py dump / apply` | A deck's bullets as an editable text file and back; `[[A1:..]]`, `[[A2:..]]`, `[[B:..]]` are the looks, `[[M:..]]` an equation zone |
| `tidy_slide_text.py` | Mechanical fixes: alerts at the right weight and tint for their level, trailing whitespace, terminal periods, maths runs, notes trimmed; `--check` reports, `--only` picks one pass |
| `measure_deck_layout.py` | Asks a hidden PowerPoint how many lines each paragraph takes and which frames overflow; `--png` renders every slide (Windows only) |
| `add_slide.py` | A new slide copied from one the deck has, with a title and bullets in the same markup |
| `align_body_box.py` | Puts a text box narrowed around a picture back on the template's margins |
| `adopt_template.py`, `retheme_fonts.py` | Finish moving an old deck onto the template: strip merged layout debris, run-level overrides and retired font names |
| `title_case_titles.py` | Slide 1's title in title case |
| `add_title_picture_layout.py` | Adds the Title Slide with Picture layout and moves a deck's title slide onto it |
| `pptx_to_pdf.py` | A PDF that matches PowerPoint: notes, comments and hidden slides removed, shrink-to-fit baked in, equation fallbacks drawn, animated slides expanded one page per click, then LibreOffice |

`STYLE.md` walks through the editing workflow that ties them together.

## Converting to PDF

```bash
python python/pptx_to_pdf.py deck.pptx -o out/
```

Tags in a slide's speaker notes steer it: `#animate-PDF` forces one page per
click, `#static-PDF` forbids it, `#omit-slide` leaves the slide out. Slide 1's
notes are never read as tags, so they can hold the instructions. On Linux and
macOS the script shows LibreOffice the fonts in `fonts/` through a private
fontconfig tree, with `fonts/fonts.conf` mapping fonts nobody has (Arial,
Calibri, Helvetica) to their metric-compatible free equivalents; on Windows,
install the fonts first.

## The add-in

`macros/AlertText.ppam` is prebuilt; `macros/AlertText.bas` is its source and
`macros/customUI14.xml` its ribbon. Change the source on a Windows machine with
"Trust access to the VBA project object model" enabled and run
`macros/update-macro-on-change.ps1`, which compiles, embeds the ribbon,
installs the result and copies it back as the `.ppam`; commit both. The ribbon
group is still labelled "AI 100", the course it was written for; relabel it in
`customUI14.xml` and `GroupLabel` in the `.bas`, rebuild, and it is yours.

An installed add-in looks in a `macros` folder beside any deck it opens and
replaces itself with a newer `.ppam` found there, so a team that keeps this
repository beside its decks stays current without reinstalling.

## Licence

Code, templates and documentation are MIT (see `LICENSE`). The fonts are under
the SIL Open Font License 1.1, each with its licence text beside it.
