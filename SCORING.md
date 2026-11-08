# Scoring

How we measure our PDFs against Microsoft Word's.

Every test fixture is a folder in `tests/fixtures/<group>/<case>/` holding an
`input.docx` and a `reference.pdf` that Word exported from it. The test suite
converts the DOCX with our library and compares our PDF with Word's. Each
comparison gives one or more scores per fixture.

## At a glance

| Score | What it asks | Range | Fails the suite when |
|---|---|---|---|
| Jaccard | Is the ink in the same places as Word's? | 0–100%, higher is better | it drops more than 2 points |
| SSIM | Do the shapes look the same, allowing a little vertical drift? | 0–100%, higher is better | it drops more than 2 points |
| Page count | Same number of pages as Word? | yes/no | never (reported) |
| Visual hash | Did any pixel of our output change? | changed/unchanged | never (flagged for review) |
| TxtBnd | Do lines start and end on the same words as Word's? | 0–100% | not run at the moment |
| `ua_fail` | How many accessibility rules does our PDF break? | count, lower is better | it goes up at all |
| `ua_deficit` | How many accessibility rules do we break that Word gets right? | count, lower is better | it goes up at all |
| `a11y_struct` | Is the document's tagged structure the same as Word's? | 0–100% | it drops more than 2 points |
| `a11y_text` | Does a screen reader get the same text, in the same order, as from Word's PDF? | 0–100% | it drops more than 2 points |

There are also some reports that don't fail the suite, except fonts in
`cases/` (see [Other checks](#other-checks)).

## How a run works

```bash
./tools/run-tests.sh                        # everything, compact report
./tools/run-tests.sh --test accessibility   # one suite
./tools/run-tests.sh --case case5           # one fixture
```

1. Each DOCX is converted to `tests/output/<group>/<case>/generated.pdf`.
2. Both PDFs are rendered to PNG images at 150 DPI with `mutool`, one image
   per page.
3. The scores are written to `tests/output/latest_scores.json`.
4. They are compared with the accepted scores in `tests/baselines.json`. A
   score that got worse by more than the allowed margin fails the run.
5. `tools/compact_report.py` prints only what changed by more than the margin
   (below): percentage scores that moved more than 2 points, and counts that
   moved at all.

Fixtures listed in `tests/fixtures/SKIPLIST` are left out.

**Baselines.** `tests/baselines.json` holds the last accepted score for each
fixture. It is committed. A better score does not update it automatically;
`tools/target/debug/accept-baselines` copies the latest scores and visual
hashes into it (`--dry-run` shows what would change first). We only accept
baselines after a human has looked at the changes.

**Why a margin of 2 points.** "Points" means percentage points: 80% → 77.5%
is a drop of 2.5 points. Rendering has tiny, harmless jitter. A 2-point
margin keeps that from failing the suite while still catching real
regressions. Counts (`ua_fail`, `ua_deficit`) have no margin: one more broken
rule is a regression.

---

## Visual scores

These compare the page images. They only look at pixels, not at the PDF's
internals.

### Jaccard (ink overlap)

**The idea:** paint every dark pixel on both pages, then ask how much of the
painted area is shared.

- A pixel counts as **ink** when it is darker than about 200 on a 0–255 scale
  (brightness weighted the usual way: green counts most, blue least). Light
  backgrounds and pale shading don't count; text, lines and dark images do.
- For each page: *pixels inked in both* ÷ *pixels inked in either*.
- The fixture's score is the average over its pages.

**Example:** Word inks 1,000 pixels, we ink 1,000 pixels, and 600 of them are
the same. Shared = 600, either = 1,400, so Jaccard = 600 ÷ 1,400 ≈ 43%.

**Why the numbers look low.** Text strokes are only a few pixels wide. If a
line of text is one pixel off, most of its ink no longer overlaps, even though
a person would call the pages identical. Across our fixtures the median is
about 68% (the middle half fall between 53% and 78%), and the "pass" mark is
only **20.5%**. The pass mark is a label (`Y`/`N` in
the full output); it doesn't fail the suite. Only a drop against the baseline
does.

**Diff images** in `tests/output/<group>/<case>/diff/` show the overlap:
gray = ink in both, blue = only in Word's, red = only in ours.

**Limits:**
- Only pages both PDFs have are compared. If we produce 3 pages and Word
  produces 4, page 4 isn't scored (the page count report catches that). A
  page whose size differs from Word's by more than 2 pixels isn't scored
  either.
- No tolerance for movement: a shift of a few pixels in any direction costs a
  lot.

### SSIM (structural similarity)

**The idea:** look at the page in small squares and ask whether each square
*looks* the same (same brightness, same contrast, same pattern), forgiving
small vertical shifts.

- The page is cut into 8×8-pixel squares.
- Only squares with some ink in **Word's** page are scored; blank squares are
  skipped.
- For each square, we try our page at the same spot and up to 8 pixels above
  or below it, and keep the best match. This forgives lines that drift a
  little vertically, which is common between two renderers.
- The scores of all squares are averaged, then averaged over pages.

SSIM is more forgiving than Jaccard, so its "pass" mark is **75%**. Across
our fixtures the median is about 88%. As with Jaccard, the pass mark is only
a label.

**Limits:**
- It only searches up and down. A horizontal shift gets no forgiveness.
- Because blank squares in Word's page are skipped, extra content we draw on
  an otherwise empty area isn't penalised here. Jaccard does penalise it.

### Page count

`ref_pages` and `gen_pages` record how many pages each PDF has. The summary
line says, for example, "228/234 page counts match". It's reported, not gated.

### Visual hashes (did anything change?)

Not a quality score. For every generated page we store a fingerprint (SHA-256)
of its pixels. If a code change alters even one pixel anywhere, the report
lists the fixture under "Visual changes".

This answers "did my change move anything?". Jaccard and SSIM answer "is it
closer to Word?". A change can alter pixels without moving those scores. For
work that shouldn't affect rendering at all, such as accessibility tagging,
the hashes must stay identical.

The accepted fingerprints are in `tests/visual_hashes.json`. They're updated
with `accept-baselines` or the case browser's "Acknowledge" button.

### TxtBnd (text boundaries) — currently not run

For each page, the text is split into lines. Each of our lines is compared
with Word's line at the same position: does it start and end with the same
word? The score is the share of lines that do. Pages whose line counts differ
by more than 15% are skipped. It also records how far page breaks have moved,
in words (`max_break_drift`).

The test lost its `#[test]` attribute in `fb9373b`, so it doesn't run and its
`text_boundary` baselines are out of date.

---

## Accessibility scores

These check what a screen reader user gets. They look inside the PDF, not at
the pixels. Two tools do the work:

- **veraPDF** checks a PDF against **PDF/UA-1**, the ISO standard for
  accessible PDF. It runs 106 automatic rules. Examples: "every image has
  alternative text", "the document declares its language", "every piece of
  content is either tagged or marked as decoration".
- **Poppler's `pdfinfo -struct-text`** prints the PDF's **structure tree**:
  the hidden outline that tells a screen reader "this is a heading, this is a
  list item, this is a table cell, this text belongs to it", in reading
  order.

Run them with `./tools/run-tests.sh --test accessibility`. This needs
`brew install verapdf poppler`; without them the test skips with a notice.

### `ua_fail` — rules our PDF breaks

The number of PDF/UA-1 rules veraPDF says our PDF fails. It is judged on its
own; Word doesn't come into it. Scored for every fixture.

Some failures come from the DOCX itself, and we leave them alone on purpose.
If the document has no title, it fails rule 7.1-9. If a picture has no
description, it fails 7.3-1. If the headings skip a level, it fails 7.4.2-1.
We don't make up a title or alt text. The baseline simply absorbs these. What
matters is that the count never goes up.

Rule 5-1 ("declares PDF/UA conformance") fails unless we claim conformance.
We claim it only when everything the document controls checks out: it has a
title, every picture has alt text, the headings are in order, every font is
embedded and no character is missing from its font. A PDF that claims
PDF/UA must then fail no rule at all, or the suite fails.

### `ua_deficit` — rules where we do worse than Word

The number of PDF/UA-1 rules where we are worse than Word's own export:

- we fail a rule that Word passes, or
- we both fail it, but we fail a bigger share of its checks than Word, by
  more than 1 percentage point.

The second case matters for rules like "all content is tagged". Word often
fails it on one stray item. We shouldn't be able to fail it on every page and
still count as "no worse".

0 means "at least as good as Word" on the automatic checks.

### `a11y_struct` — same structure as Word?

We list the structure elements of each PDF in reading order, for example
`H1, P, P, L, LI, Lbl, LBody, Table, TR, TD, …`, and measure how many edits
(insert, delete, replace one element) turn our list into Word's.

Score = 1 − *edits needed* ÷ *length of the longer list*. One wrong element
out of 100 gives 99%.

- `Span` elements are ignored, because Word splits them at every formatting
  change.
- A `Figure` with alt text and a `Figure` without it count as different
  elements.

### `a11y_text` — does a screen reader get the same text?

We collect the text of each block (paragraph, heading, list item, table cell,
…) in structure order, with spacing normalised. Then we count the characters
in blocks that match Word's **exactly** and appear in the **same order**,
divided by the total characters of the longer side.

Symbol glyphs count as one character whatever they map to: symbol-font
private-use codes, arrows, geometric shapes, dingbats and pictographs. Word
extracts the same Wingdings checkbox as raw U+F0A8 in one document and as
Unicode in another, so which symbol it is says nothing about reading order.
Math operators and • are still compared. The cost: a checked and an
unchecked box (☑/☐) look the same to this score.

Any difference in a block costs the whole block. This catches:
- text hidden from screen readers (drawn but not tagged),
- text read in the wrong order,
- headers or footers read as body text,
- words run together because no space was written ("Helloworld").

### Fixtures without a tagged Word reference

`a11y_struct`, `a11y_text` and `ua_deficit` need a Word reference that is
itself tagged. 4 of our 244 references aren't, on purpose: case63, case64,
door_air_cooling_unit_spec and missing_font_substitution are local
print-path exports, the only way Word shows the comment pane. Those fixtures
get only `ua_fail`; the other 240 get all four scores.

### Where to look

- `tests/output/<group>/<case>/generated.deficit.json`: the failing rules
  (with check counts for our PDF and Word's) and the struct and text scores.
- `*.a11y.json` next to it: cached tool output. It's reused while it's newer
  than the PDF. Delete these files after upgrading veraPDF or Poppler.
- `DOCXIDE_A11Y_GEN=libreoffice.pdf` scores another engine's PDF the same
  way, for comparison. It doesn't write baselines.

---

## Other checks

These print a report. Only the font check can fail the suite, and only for
the handcrafted `cases/` fixtures.

| Check | What it reports |
|---|---|
| Font validation | Fonts the DOCX asks for vs fonts in our PDF. A font that isn't installed is tolerated. It fails when the Helvetica fallback shows up although every requested font was found. |
| File size | Our PDF's size ÷ Word's. Labelled a pass up to 10×. We also check size by hand on every change, because we aim to be no bigger than Word. |
| Image count | Number of images per page, ours vs Word's. |
| Page geometry | Page size, ours vs Word's, within 1 point. |
| Speed (`convert_ms`) | Conversion time per fixture. Opt-in (`#[ignore]`). Flags a fixture 1.5× slower than its baseline, or a total 1.3× slower. |
| LibreOffice comparison | Scores LibreOffice's PDF with the same Jaccard and SSIM, for calibration. Opt-in: `./tools/run-tests.sh --libreoffice`. |

For a side-by-side look at Word, us and other converters, see
`tools/engine_compare.py` (described in `CLAUDE.md`).
