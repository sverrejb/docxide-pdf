# Progress for ralph/plan_vaccine_heading_fix.md

## Step 1: Add textbox content height estimation helper — REVERTED

- Was added but is no longer needed for the actual root cause
- The `estimate_textbox_height` function was removed along with the Step 2 revert
- May still be useful if the caption textbox (PARA[2] TB[0]) needs to contribute later

## Root cause re-investigation — COMPLETED (2026-03-12)

Added temporary debug prints to trace per-paragraph layout values. Key findings:

### PARA[0] is the problem, not PARA[2]

The original plan blamed the SmartArt paragraph (PARA[2]) and its caption textbox. The actual
culprit is **PARA[0]**, which contains a TopAndBottom textbox (TB[3]) positioned relative to the
page margin (`v_relative_from=Margin`).

**TB[3] properties**: `wrap=TopAndBottom, w=540, h=180, v_off=-41, v_rel=Margin, dist_b=36`

The `_ =>` branch in the textbox reservation loop (added in commit 74f19a6) does
`content_h += tb_bottom` for non-Paragraph textboxes. This adds 175pt to PARA[0]'s content
height (from 14.6 → 189.6), pushing everything below it 175pt down.

**Key insight**: Margin-relative textboxes are absolutely positioned on the page. They should
NOT contribute to the paragraph's vertical space. The `_ =>` branch should either be a no-op
or use a different calculation.

### Complication: removing 175pt is too much

Simply setting content_h to 14.6 for PARA[0] would move "The Beginning:" heading up by ~175pt,
far more than the 24.5pt needed. This suggests PARA[0] DOES need some vertical space — just
not the 175pt from the margin-relative textbox.

The plan has been updated with recommended next steps for further investigation.

## Recommended next step 1: Measure reference positions — COMPLETED (2026-03-12)

Used mutool to extract ALL text positions from both reference and generated PDFs on page 1.

### Key findings

| Text | Ref y_top | Gen y_top | Diff (gen-ref) |
|------|-----------|-----------|----------------|
| "Chapter 1" | 62.26 | 60.30 | -1.96 |
| "History of Vaccines and How..." | 128.30 | 138.99 | +10.69 |
| "The History of the Vaccine" | 290.44 | 289.49 | -0.95 |
| "(Figure 1.1) A time line..." | 449.87 | 449.26 | -0.61 |
| **"The Beginning:"** | **499.03** | **523.52** | **+24.49** |
| "Vaccines are an important..." | 525.19 | 550.08 | +24.89 |
| "The Polio Vaccine:" | 645.91 | 670.47 | +24.56 |

### Analysis

1. **Content above "The Beginning:" is well-positioned** (within ~1pt): the heading, SmartArt
   labels, and figure caption all match the reference closely.

2. **From "The Beginning:" onwards, consistent ~24.5pt offset**: all body text after the spacer
   paragraphs is shifted down by exactly the same amount.

3. **The gap is introduced between the figure caption (y≈449) and "The Beginning:" (y≈499 ref
   vs 523 gen)**. This is where PARA[3-5] (spacer paragraphs) live.
   - Reference gap: 499.03 - 449.87 = 49.16pt
   - Generated gap: 523.52 - 449.26 = 74.26pt
   - Extra: 74.26 - 49.16 = 25.1pt ≈ the 24.5pt bug

4. **PARA[0]'s 175pt inflation does NOT push "The Beginning:" down by 175pt** — the heading and
   SmartArt area above absorb or compensate for most of it. Only ~24.5pt leaks through to the
   text below the SmartArt.

5. **Reference shows spacer characters at y=456.24 and y=480.96** between the figure caption
   and "The Beginning:" — these are the spacer paragraphs (PARA[3-5]) which in the generated
   PDF don't produce visible text but still consume vertical space.

### Next step

Proceed to recommended next step 2: understand PARA[0]'s role and why 175pt inflation only
produces a 24.5pt shift. The SmartArt area (PARA[1-2]) likely has a fixed or minimum height
that absorbs most of the extra space from PARA[0].

## Recommended next step 2: Understand PARA[0]'s role — COMPLETED (2026-03-12)

Added temporary debug prints to trace exact per-paragraph layout values for page 1.

### Debug output (section 1 paragraphs)

| Para | slot_top | inter_gap | content_h | line_h | text | textboxes | space_before | space_after |
|------|----------|-----------|-----------|--------|------|-----------|-------------|-------------|
| 0 | 720.0 | 0.0 | 189.6 | 14.6 | *(empty)* | 4 | 0.0 | 10.0 |
| 1 | 530.4 | 24.0 | 24.8 | 24.8 | "The History of the Vaccine" | 0 | 24.0 | 6.0 |
| 2 | 481.5 | 6.0 | 117.8 | 14.6 | *(empty, SmartArt)* | 1 | 0.0 | 10.0 |
| 3 | 357.7 | 10.0 | 14.6 | 14.6 | *(empty spacer)* | 0 | 0.0 | 10.0 |
| 4 | 333.1 | 10.0 | 14.6 | 14.6 | *(empty spacer)* | 0 | 0.0 | 10.0 |
| 5 | 308.4 | 10.0 | 17.6 | 17.6 | *(empty spacer, sz=24)* | 0 | 0.0 | 10.0 |

Section 2 starts (continuous section break):

| Para | slot_top | inter_gap | content_h | line_h | text |
|------|----------|-----------|-----------|--------|------|
| 0 | 280.8 | 10.0 | 16.6 | 16.6 | "The Beginning: " |

### PARA[0] textboxes detail

| TB | wrap | v_rel | v_off | height | width | dist_b |
|----|------|-------|-------|--------|-------|--------|
| 0 | None | Paragraph | 94.9 | 103.0 | 104.0 | 0.0 |
| 1 | None | Paragraph | 94.9 | 106.0 | 104.0 | 0.0 |
| 2 | None | Paragraph | 91.9 | 106.0 | 107.0 | 0.0 |
| 3 | TopAndBottom | Margin | -41.0 | 180.0 | 540.0 | 36.0 |

### Key findings

1. **PARA[0]'s 189.6pt content_h is approximately CORRECT.** It gives the right position for
   PARA[1] ("The History of the Vaccine") — generated at y=289.49, reference at y=290.44 (0.95pt diff).
   Removing it would place the heading at ~y=135 from top, which is catastrophically wrong.

2. **The `_ =>` branch (`content_h += tb_bottom`) is conceptually wrong but accidentally right.**
   TB[3]'s margin-relative position (`v_off=-41 from margin`) maps to a region starting at
   y=31 from page top. The banner (h=180, dist_b=36) clears at y=247 from top. The paragraph
   at y=72 (margin top) needs content_h ≈ 175pt to push content past this banner — which is
   almost exactly what `+= tb_bottom` gives.

3. **The 24.5pt error is NOT from PARA[0]'s height calculation.** All content through the figure
   caption (PARA[2]) matches the reference within ~1pt. The error is introduced entirely in the
   spacer paragraphs (PARA[3-5]) between the SmartArt and "The Beginning:".

4. **Spacer paragraph analysis:**
   - PARA[3-5] + section 2 PARA[0] inter_gap consume 76.9pt (from 357.7 to 280.8)
   - Reference equivalent gap: ~49.16pt (from caption to "The Beginning:")
   - Excess: ~27.7pt (accounts for the 24.5pt offset plus font metrics differences)

5. **Why 175pt inflation → only 24.5pt shift:** The 175pt inflation IS correct for PARA[0].
   It does NOT "leak through" — it's accurately consumed by the heading+SmartArt area above.
   The 24.5pt error is a separate issue in the spacer paragraphs.

6. **Document structure at the spacers:**
   - PARA[3]: empty, Normal style (sz=20 = 10pt, spacing after="200" = 10pt, line="288" auto)
   - PARA[4]: empty, same as PARA[3]
   - PARA[5]: empty, sz=24 (12pt), contains `w:sectPr` (continuous section break)
   - All three have only tab stop definitions, no text runs

### Conclusion

**PARA[0]'s role is correct** — it anchors the timeline banner and its 189.6pt content_h properly
positions the heading and SmartArt below it. No changes to PARA[0]'s height calculation are needed.

**The fix should target the spacer paragraphs (PARA[3-5])**, which consume ~28pt more vertical
space than in the reference. Possible causes:
- Empty paragraph height calculation may differ from Word's handling
- The `w:sectPr` paragraph (PARA[5]) may be handled differently in Word (e.g., zero height)
- Inter-paragraph spacing between SmartArt and spacers may collapse differently in Word

### Recommended next step

Investigate why PARA[3-5] consume 76.9pt vs ~49pt in reference. The most promising avenue is
checking whether the section-break paragraph (PARA[5]) should have reduced or zero height in
Word's layout model, and whether inter-gap collapsing differs for empty paragraphs after
SmartArt content.
