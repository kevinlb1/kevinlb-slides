# Course fonts

The lecture decks name fonts that neither Windows nor Office LTSC 2021 ships.
Without them PowerPoint falls back and the slides render as garbled, overlapping
glyphs -- and the Linux build runner substitutes differently again, so the
generated PDFs would not match what you see.

These are the actual font files, committed so that your machine and the build
render identically.

## Installing

These sit beside the decks so that anyone who clones this repository has
everything needed to open them correctly. **Close PowerPoint first** -- Office
reads the font list at startup.

### Windows

Either select the `.ttf` files in Explorer, right-click and choose **Install**
(or **Install for all users**), or right-click `install-fonts.ps1` and choose
**Run with PowerShell**; it asks for Administrator rights itself and installs
for every user, or pass `-PerUser` to install for your account without them.
`-Force` reinstalls over existing entries. Either way the fonts work as soon as
PowerPoint is started again.

Until 25 September 2026 the script did not have that last property. It copied
each file into the font folder and listed it in the registry, which only makes
Windows load the font at the next sign-in, and never loaded it into the running
session (`AddFontResource`, then a `WM_FONTCHANGE` broadcast), which is what
Explorer's Install does. So a run of it changed nothing until the next
sign-out, while right-click Install worked at once -- which is what happened on
Kevin's second machine. Run unelevated, it also refused and printed a command
in a window that "Run with PowerShell" closes immediately.

This page used to say a per-user install is invisible to Office -- Word
reporting "This font isn't available on the device you are currently using",
PowerPoint drawing every glyph on top of the last. Fonts installed per user
from Explorer render correctly, so that observation was most likely the same
missing load step rather than anything about per-user installs. Word's warning
is still a good test that a font is really there: type `Inter` into its font
box, and no dialog means it is. PowerPoint's font dropdown is not -- it lists
whatever the open document names, installed or not -- and the AlertText
add-in now says so itself when a deck opens and a font it needs is missing.

### macOS

Select the `.ttf` files, open them, and click **Install Font**; or copy them
in:

```bash
cp fonts/*.ttf ~/Library/Fonts/
```

Per-user is fine here, as it is on Windows.

### Linux

```bash
mkdir -p ~/.local/share/fonts && cp fonts/*.ttf ~/.local/share/fonts/ && fc-cache -f
```

This is what CI does. On newcastle, the build also synchronizes the committed
fonts into its private font directory whenever that Git tree changes, including
new weights and removed faces, before either profile converts a deck.

## What is here, and why these

| Font | Licence | Replaces |
|---|---|---|
| Inter (ExtraLight to Black) | SIL OFL 1.1 | Aptos, Aptos Display; Roboto |
| Hanken Grotesk (ExtraLight to Black) | SIL OFL 1.1 | (candidate for the new template) |
| Libre Caslon Display | SIL OFL 1.1 | Big Caslon Medium |
| Carlito (Regular, Bold, Italic, BoldItalic) | SIL OFL 1.1 | Calibri, metric for metric: the older diagrams (Lecture 1 slide 14, for one) still name it, and without Carlito fontconfig fell back to the wider DejaVu Sans and wrapped labels mid-word |
| Fira Sans (Light, Regular, Bold) | SIL OFL 1.1 | itself: the Lecture 1 chart SVGs name it, and Office fetches it from Microsoft's cloud font service, which the build has no access to |

All of them are freely redistributable, which is the whole reason they can live
here. The fonts they replace cannot:

- **Aptos** ships only with Microsoft 365. This machine runs Office LTSC
  ProPlus 2021, which predates it and receives no cloud fonts, so Aptos will
  never appear. It has no metric-compatible substitute on Linux either, so it
  was breaking the PDFs as well. Inter is the closest free match -- both are
  neo-grotesques from the same Helvetica lineage.
- **Big Caslon Medium** ships with macOS. Libre Caslon Display is an OFL
  Caslon revival with the same high-contrast old-style feel.

Do not commit Aptos, Big Caslon, or the Microsoft core fonts here. They are not
redistributable, private repository or not.

Roboto was retired in September 2026. It was the title font on the lecture
title slides, alongside Inter everywhere else; the titles now use Inter
ExtraBold, a heavier weight of the same family, so the decks name one sans
family instead of two.

## Weights are separate families

PowerPoint knows four faces per family: Regular, Bold, Italic and Bold Italic.
It has no notion of Light or ExtraBold, and it cannot pick an instance out of a
variable font. So every weight the templates use is a **static** file whose
family name carries the weight -- "Inter Light", "Inter SemiBold", "Inter
ExtraBold" -- and text asks for that family by name with the bold flag off. The
Regular and Bold files share the plain family name ("Inter", "Hanken Grotesk"),
which is what makes Ctrl+B produce the real Bold face there.

That is why the folder holds several files per family, and why the install
script composes each registry entry from the file's own weight and style (every
one of these files reports "Regular" as its face name, so registering them by
face would leave one weight per family).

The templates use, per family: Black (cover title), ExtraBold (slide titles),
Regular (first-level text), Light (second level and cover affiliations) and
ExtraLight (deeper levels). Alert text is three weights above the level's
base: the genuine Bold face on level 1 (reached as "Inter" plus the bold
flag), SemiBold on level 2 and Medium on level 3 and below; SemiBold is also
the speaker's name on the cover. Italic and Bold Italic ship for the Regular
weight and Light Italic for the second level; other italics were left out to
keep the folder small. Medium was added on 7 September 2026, so a machine set
up before then needs the installer run again.

## Changing a deck's fonts

`python/retheme_fonts.py` rewrites `typeface="..."` attributes across
`decks/*.pptx` in place, preserving each part's original compression and
changing nothing else. Run `--check` first to see what it would do. It also
collapses CSS-style font stacks like `Courier New,monospace`, which are invalid
in OOXML -- a typeface attribute takes a single font name -- and which appear to
have come in via a Google Slides or macOS round-trip.

It also retires Roboto: runs that asked for it get `Inter ExtraBold`, and the
bold flag on those runs is dropped, since a bold flag on a weight-named family
only makes PowerPoint synthesise a fake bold on top of the real face. Pass
`--skip Lecture01.pptx` to leave a deck someone else is editing untouched; the
rewrite is idempotent, so the deck can be done later.

`docProps/app.xml` keeps a "fonts used" list that still names the old fonts.
That is informational only; PowerPoint rewrites it on save.
