# Progress for Marge — OOXML spec coverage audit

Marge walks the OOXML / WordprocessingML spec (ECMA-376 Part 1 / ISO-IEC
29500-1) one chapter per iteration, checking whether docxside-pdf covers each
chapter's features. WordprocessingML reference material is Clause 17; each 17.x
subclause is treated as one per-run "chapter".

## Chapters processed

- **17.1 Fundamentals / clause preamble** — 2026-06-18 22:45 — no implementable
  XML elements (reference-material intro + spec reading conventions). No audit
  needed.
- **17.2 Main Document Story** — 2026-06-18 22:45 — elements: `background`
  (17.2.1), `body` (17.2.2), `document` (17.2.3). `body` parsed at
  `src/docx/mod.rs:702`; `document` is the parsed XML root; final block-level
  `sectPr` in body handled (sections supported). Gaps below.

## Coverage gaps found

### 17.2.1 `w:background` (Document Background) — NOT handled
- **Spec**: §17.2.1. Specifies a page background painted behind all content on
  every page. Three flavors:
  - solid RGB via `@w:color` (ST_HexColor, §17.18.38)
  - theme color via `@w:themeColor` + optional `@w:themeTint` / `@w:themeShade`
    (resolved against the theme part)
  - gradient/image fill via a child `w:drawing` (§17.3.3.9, DrawingML)
- **Today**: no parsing anywhere in `src/docx/` (grep for `"background"` finds
  only run/cell `w:shd` shading — `src/pdf/layout.rs`, `src/pdf/table.rs`). No
  page-background field in the model, no rendering.
- **Real-world**: fixture `tests/fixtures/scraped/traditional_skills_job_form`
  (currently FAILING, Jaccard 15.5%) contains `<w:background w:color="FFFFFF"/>`
  — white, so no visual impact there, but confirms the element appears in the
  wild. A colored/themed/image background would render wrong.
- **Likely owner**: parse in `src/docx/sections.rs` (or new doc-level field in
  `src/model/`), render as a full-page fill in the existing behind-doc z-order
  layer in `src/pdf/mod.rs`. Theme-color resolution reuses theme parsing in
  `src/docx/styles.rs`; image/gradient fill reuses DrawingML fill code in
  `src/docx/textbox.rs`.
- **Caveat**: Word's PDF export paints page backgrounds only when "print
  background colors and images" is enabled — confirm against a reference PDF
  with a non-white background before implementing, or the generated output may
  diverge from Word's actual export.

### 17.15.1.26 `w:displayBackgroundShape` (settings flag) — NOT handled
- **Spec**: §17.15.1.26. Settings-part flag (`w:settings`) gating whether the
  `w:background` (above) is shown in print-layout view. Belongs to Clause 17.15
  (Settings) but directly gates 17.2.1, so noted here.
- **Today**: not parsed (`src/docx/settings.rs`). The job-form fixture sets
  `<w:displayBackgroundShape/>`.
- **Likely owner**: `src/docx/settings.rs`; consume alongside the 17.2.1 work to
  decide whether to paint the background at all.

### 17.2.3 `w:document/@w:conformance` (Document Conformance Class) — NOT read
- **Spec**: §17.2.3. Attribute on root `w:document`; ST_ConformanceClass
  (§22.9.2.2) = `strict` | `transitional`, default `transitional`.
- **Today**: attribute not read (grep `conformance` in `src/` → none); 0
  fixtures currently use it.
- **Impact**: mostly informational, BUT `strict`-conformance documents use the
  Strict (transitional-free) namespace URIs. If docxside matches only the
  Transitional `w:` namespace URI, a Strict DOCX could fail to parse. Low
  priority (rare in practice).
- **Likely owner**: root/namespace handling in `src/docx/mod.rs`.
