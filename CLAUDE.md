# Instructions for agents

These mirror `AGENTS.md`, which other agent tools read. Keep the two in sync.

This repository holds slide templates, the AlertText add-in, fonts and the
scripts that edit PowerPoint decks. Decks themselves usually live elsewhere
(a Dropbox folder, another repository) and are passed to the scripts by path.

## Editing slide text

- **Read `STYLE.md` before changing any slide's text.** It is the whole house
  style: emphasis through the add-in's three looks and never the bold flag,
  bullets with no terminal period, asides as sub-bullets, titles in Title
  Case, everything else inherited from the template. The checklist at its end
  is what a finished edit must satisfy.
- **Back up first.** Copy the deck into a `backups/` folder beside it under a
  different name, and keep a `deck_text.py dump` beside the copy.
  `deck_text.py apply` in place makes the copy itself.
- **Keep the old words, once.** The first time a slide's text changes, its
  previous text goes at the bottom of that slide's notes between the marker
  lines `STYLE.md` gives. `deck_text.py apply` does this.
- **Never edit a deck PowerPoint has open** (a `~$` lock file beside it). Ask
  for it to be closed, or work on a copy and say so.
- **Never round-trip a deck through python-pptx's `save()`** or anything
  that re-serialises the whole package. Rewrite only the parts you change,
  through the `Package` helper in `python/add_title_picture_layout.py`, as
  every script here does. `python/pptx_to_pdf.py` is the one exception, and
  it only ever writes a temporary copy.
- **Change nothing the rules do not require.** No moving content between
  slides, reordering, retitling, touching pictures, layouts or shape
  positions, or adding highlights unless asked. Conventions you notice that
  are not rules go in your report, not into the deck.
- **Measure, do not count.** Only PowerPoint knows where a line wraps:
  `python/measure_deck_layout.py --problems --png qa` on a candidate deck, and
  look at every changed slide's PNG. On a machine without PowerPoint, say that
  the layout is unverified.
- **Report** which slides changed, every judgment call, and the conventions
  the deck departs from that you left alone.

## The tools

```bash
python python/tidy_slide_text.py --check deck.pptx            # what is off
python python/deck_text.py dump deck.pptx > deck.txt           # text out
python python/deck_text.py apply deck.pptx deck.txt -o candidate.pptx --retitle
python python/measure_deck_layout.py candidate.pptx --problems --png qa
python python/deck_text.py apply deck.pptx deck.txt --retitle  # in place, after backup
python python/tidy_slide_text.py deck.pptx
python python/pptx_to_pdf.py deck.pptx -o out/
```

`deck_text.py`'s markup: `- ` lines are paragraphs, four more spaces per
level, `[[A1:..]]` alert 1, `[[A2:..]]` alert 2, `[[B:..]]` bolder,
`[[M:..]]` maths, `[[URL:..]]` a kept hyperlink. A dump flags `[[BOLD:..]]`
and `[[FMT:..]]` where formatting matches no look; `apply` refuses to write a
line still carrying either, so resolve each one.

## The add-in and the templates

- `macros/AlertText.bas` and `python/deck_text.py`'s `LOOKS` must agree on
  what each look writes (weight per level, colour per level). Change one,
  change the other, and update the tables in `STYLE.md`.
- Rebuilding `AlertText.ppam` needs Windows, PowerPoint and VBA project
  access; `macros/update-macro-on-change.ps1` does it. Commit the source and
  the rebuilt `.ppam` together.
- Weight is always a family name, never the bold flag, except that level-1
  alert 1 uses the bold flag on the plain family to reach the real Bold face.
  Do not add fonts that cannot be redistributed (Aptos, Big Caslon, the
  Microsoft core fonts); free replacements and the reasons are in
  `fonts/README.md`.
- `.pptx`, `.potx`, `.ppam` and `.ttf` are binary in `.gitattributes`; keep
  them so.

## Commits

At the end of every turn that changes this repository, make a local commit
with a descriptive message. Do not push unless asked.
