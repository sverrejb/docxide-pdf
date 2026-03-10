# Fix stem_partnerships_guide Rendering Issues

## Context

The `scraped/stem_partnerships_guide` fixture has three visible rendering issues vs the Word reference:
1. Logo rendered as 3 duplicates + black rectangle (page 1)
2. Missing white "Opportunity through learning" text on the colored shape (page 1)
3. ~1 line text drift by "How do I establish..." header (page 2)

Current scores: Jaccard 25.84%, SSIM 41.15%. Fixing these should significantly improve both.

---

## Fix 1: Grayscale JPEG Color Space (highest impact)

**Root cause**: `image1.jpeg` is a **grayscale JPEG** (1 component, 542x135px). The embed_image closure in `src/pdf/mod.rs:739` hardcodes `device_rgb()` for all JPEGs. The PDF viewer misinterprets 1-byte-per-pixel grayscale data as 3-byte RGB, causing the image to appear as 3 squished copies + a black rectangle.

**Changes**:

1. **`src/model.rs:130`** (EmbeddedImage struct) — Add `pub jpeg_components: u8` field (default 3)

2. **`src/docx/images.rs:34`** (image_dimensions) — The SOF marker already contains the component count at byte `data[i+9]`. Change return type to `Option<(u32, u32, ImageFormat, u8)>`, extract and return component count. PNG branch returns 3 (irrelevant, PNG goes through decode path).

3. **`src/docx/images.rs:83`** (read_image_from_zip) — Destructure the 4-tuple, set `jpeg_components` on `EmbeddedImage`.

4. **`src/pdf/mod.rs:733-742`** (embed_image closure, JPEG branch) — Replace `xobj.color_space().device_rgb()` with:
   ```
   match img.jpeg_components {
       1 => device_gray(),
       4 => device_cmyk(),
       _ => device_rgb(),
   }
   ```

5. Update all other `EmbeddedImage` construction sites to include `jpeg_components: 3` (or the parsed value).

---

## Fix 2: Mid-paragraph page break renders text on wrong page

**Root cause**: The "Opportunity through learning" paragraph contains both text runs AND a `w:br type="page"` in the same paragraph. Currently, both `w:pageBreakBefore` (paragraph property) and `w:br type="page"` (run-level element) set the same `has_page_break` flag, which becomes `page_break_before: true`. This causes the renderer to break to a new page BEFORE rendering the text, so the white text ends up invisible on page 2 (white on white).

The color parsing itself is correct: `w:color w:val="FFFFFF"` → `[255, 255, 255]`.

**Changes**:

1. **`src/docx/runs.rs:37`** (ParsedRuns) — Rename `has_page_break` to `has_page_break_before` and add `has_page_break_after: bool`.

2. **`src/docx/runs.rs:524-525`** — `w:br type="page"` sets `has_page_break_after = true` (NOT `has_page_break`).

3. **`src/docx/runs.rs:604-608`** — `w:pageBreakBefore` sets `has_page_break_before = true`.

4. **`src/docx/runs.rs:613`** — Guard condition uses `has_page_break_before` (not the old combined flag).

5. **`src/model.rs:264`** (Paragraph) — Add `pub page_break_after: bool`.

6. **`src/docx/mod.rs:584-585`** — Wire up: `page_break_before` from `parsed.has_page_break_before || style.page_break_before`; `page_break_after` from `parsed.has_page_break_after`.

7. **`src/pdf/mod.rs`** — After paragraph rendering completes (around line 1400, after `slot_top -= content_h + bdr_top_pad`), add handling for `para.page_break_after`:
   ```
   if para.page_break_after {
       // flush page, same pattern as the page_break_before code at line 1048-1072
       all_contents.push(mem::replace(&mut current_content, Content::new()));
       // ... push links, footnotes, alpha states, etc.
       slot_top = effective_slot_top(...);
       effective_margin_bottom = compute_effective_margin_bottom(...);
       is_first_page_of_section = false;
   }
   ```

---

## Fix 3: Text drift on page 2 (investigate after fixes 1 & 2)

**Likely cause**: The paragraph containing the floating anchor image has `w:pPr/w:rPr/w:sz val="460"` (230pt paragraph mark). Since its only run produces a floating image (not a text run), the `runs` vector is empty and the synthetic run code (`runs.rs:613-632`) creates a 230pt placeholder. This makes the paragraph ~230pt tall, which may differ from Word's handling.

**Plan**: Re-evaluate after fixes 1 and 2 are applied, as the page break fix changes content placement. If drift persists, investigate whether the synthetic run should be suppressed or sized differently when the paragraph only contains floating images.

---

## Verification

```bash
# Test this specific fixture
DOCXIDE_CASE=stem_partnerships_guide cargo test -- --nocapture

# Check for regressions across all cases
cargo test -- --nocapture 2>&1 | grep "REGRESSION"

# Visual comparison
# Check tests/output/scraped/stem_partnerships_guide/generated/page_001.png
# Check tests/output/scraped/stem_partnerships_guide/diff/page_001.png
```

## Critical Files
- `src/model.rs` — EmbeddedImage (add jpeg_components), Paragraph (add page_break_after)
- `src/docx/images.rs` — image_dimensions() return type, read_image_from_zip()
- `src/docx/runs.rs` — ParsedRuns (split page break flags), w:br handling
- `src/docx/mod.rs` — Paragraph construction wiring
- `src/pdf/mod.rs` — embed_image JPEG color space, page_break_after rendering
