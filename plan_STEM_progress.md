# Progress for plan_STEM.md

## Fix 1: Grayscale JPEG Color Space — COMPLETED
- Added `jpeg_components: u8` field to `EmbeddedImage` in `src/model.rs:137`
- Changed `image_dimensions()` return type to `Option<(u32, u32, ImageFormat, u8)>` — extracts component count from JPEG SOF marker byte at `data[i+9]` in `src/docx/images.rs:34`
- Updated `read_image_from_zip()` to pass `jpeg_components` through in `src/docx/images.rs:83`
- Updated `embed_single_image()` JPEG branch to use `device_gray()` for 1-component, `device_cmyk()` for 4-component, `device_rgb()` otherwise in `src/pdf/mod.rs:601`
- Result: Logo on page 1 now renders correctly as a single grayscale image (was 3 duplicates + black rectangle)
- Scores: Jaccard 25.84% → 26.1% (+0.3pp), no regressions across all cases

## Fix 2: Mid-paragraph page break renders text on wrong page — COMPLETED
- Split `has_page_break` in `ParsedRuns` into `has_page_break_before` and `has_page_break_after` in `src/docx/runs.rs:38-40`
- `w:br type="page"` (run-level) now sets `has_page_break_after` instead of `has_page_break` in `src/docx/runs.rs:526`
- `w:pageBreakBefore` (paragraph property) sets `has_page_break_before` in `src/docx/runs.rs:607`
- Synthetic run guard condition updated to use `has_page_break_before` in `src/docx/runs.rs:615`
- Added `page_break_after: bool` field to `Paragraph` in `src/model.rs:350`
- Wired up both flags in paragraph construction in `src/docx/mod.rs:609-611`
- Added `page_break_after` handling in renderer: flushes page after paragraph rendering completes in `src/pdf/mod.rs:2085-2096`
- Result: White "Opportunity through learning" text now renders on page 1 (on the purple shape) instead of being invisible on page 2 (white on white)
- Scores: Jaccard 26.1% → 25.5% (-0.6pp), no regressions across all cases. Small score drop expected as page content distribution changed.

## Fix 3: Text drift on page 2 — INVESTIGATED / NO CODE CHANGE
- Investigated the 230pt synthetic run hypothesis: the paragraph with `w:pPr/w:rPr/w:sz val="460"` and a floating behind-doc image is correctly creating a 230pt placeholder run. The author intentionally set this size to push "Opportunity through learning" down page 1 to align with the cover shape.
- Confirmed page 1 renders correctly: logo, titles, 230pt gap, and "Opportunity through learning" all match reference positioning.
- The page 2+ drift is NOT caused by the 230pt paragraph (which is in section 1, before the section break). Verified by visual comparison of page 1 reference vs generated.
- Analyzed page 2–4 diff images: drift accumulates gradually across body text paragraphs in sections 2–3, consistent with font metrics/line spacing differences rather than a structural layout bug.
- The document uses Calibri (via theme and explicit style declarations) with docDefaults line spacing of 276 auto (1.15x). Small per-paragraph height differences from font metric approximations accumulate over many paragraphs.
- Conclusion: no actionable code change for this plan. The 230pt synthetic run is correct. Remaining drift is a general font metrics accuracy issue.
- Scores: unchanged at Jaccard 25.5%, SSIM 40.8%
