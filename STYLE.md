# Slide text house style

How the text on the slides is written and formatted, in enough detail that a
person or an agent can bring a deck into line without asking. The decks sit on
one of the templates in `templates/` and are usually rendered to PDF by
`python/pptx_to_pdf.py`, so what matters is how PowerPoint lays the text out,
not how the XML reads.

## Before touching a deck

1. **Back up.** Copy the deck into a `backups/` folder beside it under a
   different name, e.g. `backups/Deck - original before formatting cleanup
   2026-09-15.pptx`, and keep a plain-text dump of the slides beside it.
   `python/deck_text.py apply` makes the copy itself when it edits in place.
2. **Keep the old words, once.** The first time a slide's text is changed, the
   text it had goes at the bottom of that slide's notes pane between two marker
   lines (format below). One record per slide, of what a human wrote before any
   of this started, not a log of every pass: later edits add nothing, and
   `tidy_slide_text.py` cuts an existing log back to its first block. Notes
   never reach a PDF made with `pptx_to_pdf.py`, which strips them.
3. **Never edit a deck PowerPoint has open** (a `~$Deck.pptx` lock file sits
   beside it): PowerPoint's next save would overwrite the edit. Ask for it to
   be closed, or work on a copy and say so.
4. **Never round-trip a deck through python-pptx** or anything else that
   re-serialises the whole package; rewrite the parts you change and keep
   every other part byte for byte (`python/deck_text.py` does, through the
   `Package` helper in `python/add_title_picture_layout.py`). python-pptx's
   `save()` rewrites every part it parsed and drops parts it does not model,
   and has garbled decks that way.
5. **Change nothing the rules do not require.** The rules are about wording and
   emphasis. They do not license moving content between slides, reordering
   points, changing titles, pictures, layouts or shape positions, or adding
   highlights. Conventions you notice that are not rules yet go in the report,
   not into the deck.

## The template's look

Everything here is what the master defines. Text that "follows the template"
carries none of it explicitly, only its outline level, so it re-styles itself
when the level or the template changes. The Inter template:

| Level | Size | Family | Ink | Bullet |
|---|---|---|---|---|
| 1 | 28 pt | Inter | navy `002040` | `•` |
| 2 | 24 pt | Inter Light | navy at 70 % | `–` |
| 3 | 20 pt | Inter ExtraLight | navy at 55 % | small square |
| 4 and deeper | 16 pt, then 11 pt | Inter ExtraLight | navy at 30 % | `–`, then `»` |

The Hanken Grotesk template is the same ladder in that family. Titles are
ExtraBold, white, in the navy title bar. Weight is always a family name (Inter
Light, Inter SemiBold), never PowerPoint's bold flag, except where the next
table says Bold: weight 700 has no family of its own, so the flag is how the
real Bold face is reached. Body placeholders shrink text on overflow
(`normAutofit`); a slide that needs that has too much on it.

## Emphasis: three looks, applied with the add-in

`macros/AlertText.ppam` puts the buttons in a group at the end of the Home tab
(`macros/AlertText.bas` is the source). Use them, or write exactly what they
write. Never the B button, never a hand-picked colour, weight or size.

| Look | Level 1 | Level 2 | Level 3 and deeper | Colour |
|---|---|---|---|---|
| Alert 1 (press Alert once) | Inter with the bold flag, which selects the real Bold face | the same | Inter SemiBold | Accent 1, blue `2F77B4` |
| Alert 2 (press Alert twice) | Inter SemiBold | Inter SemiBold | Inter Medium | Accent 2, teal `1E6663` |
| Bolder | Inter SemiBold | Inter SemiBold | Inter Medium | unchanged |

Level 2 wears level 1's weight and is told apart by its size and its tint: an
alert a weight lighter, with the tint on top, is too faint on a sub-bullet to
carry the eye.

Both accents lighten with the level so an alert sits at about the lightness of
the text around it: the accent blended toward white, keeping 100 % of it on
level 1, 85 % on level 2, 70 % on level 3 and 50 % deeper, and written as a
plain colour.

| Level | Alert 1, from Accent 1 | Alert 2, from Accent 2 |
|---|---|---|
| 1 | `2F77B4` | `1E6663` |
| 2 | `4E8BBF` | `407D7A` |
| 3 | `6DA0CB` | `629492` |
| 4 and deeper | `97BBDA` | `8FB3B1` |

Those are the values for the Inter templates; the add-in and the tools compute
them from the deck's own theme when they write, so a template with other
accents gets its own. Blending toward white (rather than PowerPoint's "Accent
N, Lighter X%", which holds HSL saturation and makes a dark accent more vivid
as it lightens) fades the chroma with the lightness, as the body ink's own
ladder does. The blend is baked rather than left as transparency so an alert
looks the same over a tinted background or a picture. In tables and text boxes
the level-1 look applies. Bullets keep their own colour whatever the text does.

When to use which:

- **Alert 1 on every top-level bullet's key term**: the term the bullet
  introduces or defines, or the label of a definition-list bullet (`**State:**
  each different situation the environment can be in`). A question or a bare
  example line may have no key term; every other top-level bullet has one. It
  replaces every use of plain bold.
- **Sub-bullets carry alert 2 or nothing**, never alert 1: a defined term in a
  sub-bullet (`maintain a **frontier**: a list of states under consideration`)
  is alert 2, and most sub-bullets need no emphasis.
- **Alert 2 on level 1** when alert 1 would make the slide busy: a second kind
  of highlighted thing on the same slide.
- **Bolder** where bold is really called for without a colour drawing the
  eye: table headings, a column header, a name in a list.
- **The alert covers the term or label only**: `**Initial state:** how we
  start`, with the space after the colon plain. `**Instructor:** Kevin
  Leyton-Brown`, never `**Instructor: **Kevin Leyton-Brown`.
- **A text box that PowerPoint no longer counts as a placeholder**, one it
  still names "Content Placeholder N" because the layout's own placeholder is
  taken by a table, is still an outline body: its levels keep the master's
  sizes and weights, so its alerts take the level's look like anywhere else.
  The add-in cannot see this and writes the level-1 look on every line;
  `python/tidy_slide_text.py` puts those runs right.
- **The weight and colour are the look's for the level**, from the tables
  above; an alert at any other weight or colour is a defect.
  `python/tidy_slide_text.py` finds and fixes those, including a colour left
  behind by promoting or demoting the bullet.
- On a cleanup pass, give each top-level bullet its alert-1 key term where it
  lacks one, demote alert 1 on sub-bullets to alert 2, convert plain bold, and
  add nothing beyond that.

What the add-in writes, for tools that generate runs (this is also how
`python/deck_text.py` recognises the looks when it dumps a deck):

```xml
<!-- alert 1 on level 1 -->
<a:rPr lang="en-US" b="1" dirty="0">
  <a:solidFill><a:srgbClr val="2F77B4"/></a:solidFill>
  <a:latin typeface="Inter" panose="02000503000000020004" pitchFamily="2" charset="0"/>
</a:rPr>
<!-- alert 1 on level 2: the same face, the level's colour -->
<a:rPr lang="en-US" b="1" dirty="0">
  <a:solidFill><a:srgbClr val="4E8BBF"/></a:solidFill>
  <a:latin typeface="Inter" panose="02000503000000020004" pitchFamily="2" charset="0"/>
</a:rPr>
<!-- alert 2 on level 3 -->
<a:rPr lang="en-US" dirty="0">
  <a:solidFill><a:srgbClr val="629492"/></a:solidFill>
  <a:latin typeface="Inter Medium" panose="02000603000000020004" pitchFamily="2" charset="0"/>
</a:rPr>
```

Plain text is `<a:rPr lang="en-US" dirty="0"/>`: no font, size or colour.

## Writing the bullets

A bullet keeps the sentence its author wrote. The rules below shape and tidy
it; they do not license rewording unless the request says so.

1. **No period.** A bullet never ends in a period. Question marks and
   exclamation marks stay at the end of the line: `What would the output
   be?`, `Make sure your instructions are clear!`. `tidy_slide_text.py --only
   periods` takes the period off the end of every bullet in a deck, and leaves
   the periods inside one alone.
2. **One line, where that costs nothing.** *(Apply only when asked to
   rewrite.)* Each bullet is phrased without redundancy and normally occupies
   one line as PowerPoint lays it out. Three things outrank the single line,
   and each goes in the report:
   - **Meaning.** Before shortening a bullet that ran to two lines, read the
     one-line version against the original and ask whether anything has been
     changed, dropped or blurred. If it has, keep the longer text and let it
     wrap.
   - **A slide with little on it.** On a sparse slide nothing is gained by
     trimming a sentence to fit one line.
   - **A narrowed box.** Beside a picture a line holds only a few words, so
     bullets wrap there whatever they say. What must not happen is the frame
     overflowing.
3. **Asides become sub-bullets.** Parenthetical remarks, lists and sub-points
   move one level down, one item per bullet. A two-item pair may stay on one
   line when splitting it only adds lines (`**depth-first search (DFS)** and
   **breadth-first search (BFS)**`). This is the rule that shapes a paragraph
   into a slide, and it does it without touching a word: a bullet of several
   sentences keeps its first and puts the others a level below, and a clause
   after a spaced dash becomes a bullet of its own. Four things are left whole,
   because splitting them breaks the sentence:
   - two dashes in one sentence, which are a pair around an interjection
     rather than a tail;
   - a tail that begins *and*, *or*, *but*, *that*, *which* or *who*, which
     finishes the sentence rather than commenting on it;
   - a piece of three words or fewer, which belongs to the sentence before it
     (`Why?` is not a bullet);
   - a period that ends an abbreviation or a decimal, or that sits inside
     quotation marks.

   A new bullet that sits beside sentences takes a capital by rule 9; the
   author's own bullets keep the case they had.
4. **Emphasis as above.** Every top-level bullet with a key term carries it
   in alert 1; sub-bullets carry alert 2 or nothing; Bolder where bold is
   really wanted; never the bold flag, never an explicit colour; on label
   bullets the alert covers the label only, not the space after it.
5. **Everything else follows the template.** No explicit sizes, fonts,
   colours, line spacing or indents on runs or paragraphs; no resized
   placeholders except around a picture; no trailing spaces or tabs at the end
   of a paragraph; trailing empty paragraphs removed (they push the text into
   shrink-to-fit). One deliberate empty paragraph as a spacer between groups
   is fine.
6. **Content stays.** Rephrase only as far as the rules force. Splitting a
   sentence into a bullet and sub-bullets, dropping "we can see that", turning
   "In X, we do Y" into `**X**: Y` is paraphrase; dropping a fact, an example
   or a qualifier is not. When a one-line rewrite would lose a nuance, keep the
   nuance in a sub-bullet or accept the second line, and name the spot in the
   report either way.
7. **Titles in Title Case.** `The Story So Far`, `Depth- and Breadth-First
   Search`, `Wrap-Up`. Capitalise every word except articles, coordinating
   conjunctions and short prepositions inside the title; adverbs such as `So`
   are words, not prepositions; the first and last word and both halves of a
   hyphenated word are always capitalised. Slide 1's title is the deck title
   (`python/title_case_titles.py` does that one); `deck_text.py apply
   --retitle` writes the rest from the text file.
8. **Sub-bullet capitalisation.** A sub-bullet that continues its parent's
   phrase starts lowercase (`states are circles`), as does any other fragment;
   a standalone question or imperative starts with a capital (`How long was
   the flight?`). Top-level bullets start with a capital unless they are items
   of one list.
9. **One list, one capitalisation.** The sibling bullets under a parent (or
   the top-level bullets of a slide) all start alike. If any item must be
   capitalised, capitalise every item in that list.
10. **Titles do not wrap.** A title fits its bar on one line, as PowerPoint
    lays it out (`measure_deck_layout.py` reports the title's line count).
    Rephrase first; only if the title cannot be shortened, let the title
    placeholder's shrink-to-fit reduce the size. As a first guess the bar
    holds about 40 characters.
11. **A box moved around a picture gives up one edge, not four.** Every other
    edge stays where the text-only template has it, so the text does not drift
    inwards or downwards and the first line starts where it starts on the
    slides with no picture at all:
    - **beside the text**, a picture on the right: the box keeps the layout's
      top-left and bottom-left corners and only its right edge moves;
    - **below the text**, a picture under the bullets: the box keeps the
      layout's two top corners and only its bottom edge moves;
    - **above the text**, a table over the bullets: the box keeps the layout's
      width and its bottom edge, and its top edge moves down to clear it.
    `python/align_body_box.py --check` reports the departures and fixes them.
    A box with a position of its own on a slide with nothing else on it is
    reported and left: rule 5 says it should not have one, but why it does is
    a judgment call.
12. **Maths is an equation object.** A letter standing for a quantity, an
    operator or a short expression goes in a math zone, what PowerPoint's
    equation editor inserts (Cambria Math italic), and not in the body font:
    the *k* of *k*-nearest neighbours, the inputs x1, x2, x3 of a neuron, 3x3,
    8/10 = 80%. `deck_text.py` writes one from `[[M:k]]`, takes a subscript as
    `[[M:x_1]]`, and gives it the alert's colour and weight when the symbol
    sits inside a term (`[[A1:[[M:k]]-nearest neighbours]]`). Numbers that are
    just prose stay prose (`1.5 billion parameters`), and so do letters that
    name a thing rather than stand for a quantity (the `Q:` and `A:` of an
    example prompt).

    The shape that holds one is wrapped, as PowerPoint wraps it: an
    `mc:AlternateContent` whose `mc:Choice` holds the shape with its equations
    and whose `mc:Fallback` holds a copy for readers that do not do OOXML
    maths. LibreOffice is one of those, and a zone with no fallback loses its
    symbols in the PDF. PowerPoint's fallback is a picture of the shape, which
    renders correctly but puts no text in the PDF; `deck_text.py` writes the
    same text in Cambria Math runs instead, which renders and stays
    searchable, and PowerPoint replaces it with its picture the next time the
    deck is saved. Either way the slide looks right.

## Fitting the line

PowerPoint decides where a line wraps, so measure rather than count
(`python/measure_deck_layout.py`). As a first guess for the full-width content
placeholder of the Title and Content layout, 12.77 in wide:

| Level | Fits on one line | Height of one line with its spacing |
|---|---|---|
| 1 | about 65 characters | about 43 pt |
| 2 | about 75 characters | about 34 pt |
| 3 | about 90 characters | about 29 pt |

Alert text is wider than regular text, so a level-1 line with a long bold
label fits fewer characters. The placeholder is about 406 pt tall inside its
margins: nine level-1 lines, or a mix adding up to that. A slide that measures
over is shrunk by PowerPoint (line spacing first, then font size) and
`pptx_to_pdf.py` bakes that shrink into the PDF. Tolerate a 10 % line spacing
reduction on a dense slide and nothing more. The value stored in the file can
be stale, so trust the measurement, not the XML. `deck_text.py apply` clears
the stored value so PowerPoint recomputes it.

## Workflow

Measuring and rendering need Windows with PowerPoint; everything else runs
anywhere the Python environment does.

```bash
DECK=decks/Deck.pptx
mkdir -p decks/backups
cp -p "$DECK" "decks/backups/Deck - original before formatting cleanup $(date +%F).pptx"
python python/tidy_slide_text.py --check "$DECK"                         # 0. alert weights and trailing whitespace
python python/deck_text.py dump "$DECK" > deck.txt                       # 1. the deck as text
#   ... edit the "- " lines and the "## Slide N:" titles of deck.txt to the rules ...
python python/deck_text.py apply "$DECK" deck.txt -o candidate.pptx --retitle   # 2. a candidate deck
python python/measure_deck_layout.py candidate.pptx --problems --png qa   # 3. lines, overflow, a PNG per slide
#   ... look at every changed slide's PNG, fix the text file, repeat 2 and 3 ...
python python/deck_text.py apply "$DECK" deck.txt --retitle               # 4. in place: backup, notes copies, lock check
python python/tidy_slide_text.py "$DECK"                                  # 5. the mechanical fixes from step 0
#   ... open the deck in PowerPoint once and save it: step 6 ...
python python/tidy_slide_text.py "$DECK"                                  # 7. PowerPoint puts empty runs back
```

Step 6 is not optional when the text has grown. A placeholder set to shrink on
overflow records what PowerPoint settled on (`<a:normAutofit fontScale="85000"
lnSpcReduction="20000"/>`) and `pptx_to_pdf.py` bakes that scale into explicit
sizes so that LibreOffice draws the sizes PowerPoint chose. Text written into
the XML from outside leaves the old scale, or none, and LibreOffice then fits
the text its own way. PowerPoint only measures a slide it lays out, so a deck
opened and saved by automation is not enough on its own: go to each changed
slide and set the placeholder's AutoSize again.

`python/align_body_box.py` puts a text box narrowed around a picture back on
the template's margins (rule 11), and `python/add_slide.py` writes a slide the
deck does not have yet, copying one it does for the layout, the footers and
the language; its `--body` file is the same markup, and the result goes
through steps 3 and 5 like any other edit:

```bash
python python/add_slide.py "$DECK" --after 19 --title Conclusion --body conclusion.txt
```

Before step 4, check that no `~$` lock file exists and that the deck on disk
still matches your backup: someone may have saved it meanwhile. Then report
which slides changed, every judgment call (two-line exceptions, added labels,
nuances dropped or moved), and the conventions above the deck departs from.

### The text file

`deck_text.py dump` writes one block per slide. Only the `- ` lines are
applied; four more leading spaces per outline level; a bare `-` is an empty
paragraph; `<br>` is a line break inside a paragraph. Emphasis is
`[[A1:text]]`, `[[A2:text]]` or `[[B:text]]`; `[[M:text]]` is a math zone
(`[[M:x_1]]` for a subscript), which may sit inside an alert; a hyperlinked run
is `[[URL:text]]` and keeps its link on apply; a detached content placeholder
counts as a body and is given the master's body style when rewritten; a dump
flags non-conforming emphasis as `[[BOLD:text]]` (the bold flag) or
`[[FMT:text]]` (any other explicit run formatting), and `apply` refuses to
write a paragraph that still carries either, so decide each one first. Lines
starting with `>` are the notes, `[read-only …]` blocks with `|` lines are
tables and text boxes, and `[body: NAME]` labels the placeholders on a slide
that has more than one. `apply` rewrites only placeholders whose text differs
from the file and leaves tables, text boxes and pictures alone: those are
edited in PowerPoint, to the same rules. A title that differs from the file
marks a slide that moved, so the slide is skipped, unless `--retitle` says the
file's titles are the intended new ones; then the title is written as one run
in its existing formatting.

```text
## Slide 5: Applications of Search
- Search can be applied in a few different scenarios:
    - Find [[A1:any solution]]
        - you just want some way to get from Vancouver to Antarctica
        - you want to complete a Sudoku
    - Find [[A1:the best solution]]
```

### The notes-pane copy

```text
--- Original slide text before the 2026-09-15 edit; [[A1:..]] alert 1, [[A2:..]] alert 2, [[B:..]] bolder, [[BOLD:..]] plain bold ---
Title: The story so far
- In the last lecture, we tried getting an agent to perform a task by [[A1:explicitly writing a set of instructions]].
- However, beyond very simple environments
    - it can be hard to come up with such instructions by hand
--- End of original slide text ---
```

Existing notes (including the `#animate-PDF`, `#static-PDF` and `#omit-slide`
tags `pptx_to_pdf.py` reads) stay above the block, separated by a blank line;
where a block is already there, nothing is added and any later blocks are
removed. Never put a tag word inside a block.

## Checklist before installing

- backup copy in `backups/`, dump beside it, deck not open in PowerPoint
- every edited slide: the author's own words, with no terminal periods and
  asides a level down; no bold flag; no explicit sizes, fonts or colours; no
  trailing whitespace or empty paragraphs; no frame overflowing
- every top-level bullet's key term in alert 1, sub-bullets in alert 2 or
  plain; alerts at the level's weight and tint (`tidy_slide_text.py --check`
  reports nothing) and covering the term only
- titles in Title Case; sub-bullets lowercase except standalone questions and
  imperatives
- slide 1, pictures, tables and shape positions untouched unless asked
- notes carry the previous text; tags intact
- renders of every changed slide reviewed
- report lists changes, judgment calls and conventions not applied
