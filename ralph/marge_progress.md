# Progress for Marge — OOXML spec coverage audit

Marge walks the OOXML spec (ECMA-376 Part 1 / ISO/IEC 29500-1) one chapter per
iteration, checking whether docxside-pdf covers each chapter's features.

## Scope & master checklist

A DOCX renderer touches more than WordprocessingML. The audit covers the
clauses whose markup actually appears in a .docx and that this codebase renders.
Walk top to bottom; do the first `[ ]` item each run; mark `[x]` when audited.
When every box is `[x]`, the audit is complete — output COMPLETE and stop.

**Clause 17 — WordprocessingML** (core: `src/docx/*`, `src/pdf/*`)
- [x] 17.2 Main Document Story
- [x] 17.3 Paragraphs and Rich Formatting
- [x] 17.4 Tables
- [x] 17.5 Custom Markup
- [x] 17.6 Sections
- [x] 17.7 Styles
- [x] 17.8 Fonts
- [x] 17.9 Numbering
- [x] 17.10 Headers and Footers
- [x] 17.11 Footnotes and Endnotes
- [x] 17.12 Glossary Document
- [x] 17.13 Annotations (comments, revisions / tracked changes)
- [x] 17.14 Mail Merge
- [x] 17.15 Settings
- [x] 17.16 Fields and Hyperlinks
- [x] 17.18 Simple Types (ST_* value domains — reference only, low audit value)

**Clause 20 — DrawingML Framework** (`src/geometry/*`, `docx/textbox.rs`, `docx/images.rs`)
- [x] 20.1 DrawingML Main (shapes, presetGeometry, fills, effects, text body — `a:`)
- [x] 20.2 DrawingML Picture (`pic:` picture container)
- [x] 20.4 DrawingML WordprocessingML Drawing (`wp:` inline/anchor positioning, wrap)

**Clause 21 — DrawingML Components** (`docx/charts.rs`, `docx/smartart.rs`, `pdf/charts*.rs`, `pdf/smartart.rs`)
- [x] 21.2 DrawingML Charts (`c:` chartSpace)
- [x] 21.4 DrawingML Diagrams (`dgm:`/`dsp:` SmartArt)

**Clause 22 — Shared MLs**
- [x] 22.1 Math (OMML `m:` namespace)
- [x] 22.6 Bibliography (`b:` citation sources)

**Out of scope** (no print impact / not in a DOCX): clause 18 SpreadsheetML,
clause 19 PresentationML, 20.3 Locked Canvas, 20.5 SpreadsheetML Drawing,
21.1/21.3 chart-internal drawings, 22.2–22.5 / 22.7–22.9 (doc properties,
variant types, custom-XML props, additional characteristics, simple-type
domains). VML (legacy `v:` drawings) lives in ECMA-376 Part 4, not this spec
PDF — audit it separately against the codebase's VML fallback if needed.

## Chapters processed

- **17.1 Fundamentals / clause preamble** — 2026-06-18 22:45 — no implementable
  XML elements (reference-material intro + spec reading conventions). No audit
  needed.
- **17.2 Main Document Story** — 2026-06-18 22:45 — elements: `background`
  (17.2.1), `body` (17.2.2), `document` (17.2.3). `body` parsed at
  `src/docx/mod.rs:702`; `document` is the parsed XML root; final block-level
  `sectPr` in body handled (sections supported). Gaps below.
- **17.3 Paragraphs and Rich Formatting** — 2026-06-18 23:01 — paragraph
  properties (17.3.1) parsed in `src/docx/paragraph.rs`, run properties (17.3.2)
  in `src/docx/runs.rs`. WELL COVERED: `jc`, `ind`, `spacing`, `tabs`, `shd`,
  `pStyle`, `numPr`, `keepNext`, `keepLines`, `widowControl`, `contextualSpacing`,
  `snapToGrid`, `autoSpaceDE/DN`, `outlineLvl`, `framePr`, `pBdr`
  (top/bottom/left/right/between) at paragraph level; `b/i/u/strike/dstrike/caps/
  smallCaps/vanish/color/sz/rFonts/highlight/shd/bdr/vertAlign/kern/spacing(char)/
  w(text scale)` at run level (all consumed in `src/pdf/`). Gaps below, grouped by
  priority.
- **17.4 Tables** — 2026-06-18 23:06 — table model in `src/model/table.rs`,
  parsing in `src/docx/tables.rs` (1004 lines), layout/render in
  `src/pdf/table.rs` + `src/pdf/table_layout.rs`. WELL COVERED: `tbl`, `tblPr`,
  `tblGrid`/`gridCol`, `tr`/`tc`/`trPr`/`tcPr`; `tblW` (type-aware auto/pct/dxa);
  `tblBorders`/`tcBorders` incl. §17.4.66 conflict resolution with weight + style
  precedence; `insideH`/`insideV`; `start`/`end` + legacy `left`/`right` borders &
  margins; `tblCellMar`/`tcMar`; `tblInd`; `tblLayout` (fixed/autofit); `tblLook`
  (named attrs + legacy hex bitmask); `tblpPr` floating positioning (all attrs);
  `tblStyle` + conditional formatting (wholeTable / first-last row-col / corners /
  band1-2 Horz-Vert); `gridSpan`; `vMerge`; `vAlign`; `textDirection`; `trHeight`
  (+ exact); `cantSplit`; `tblHeader`; `hideMark` (consumed in
  `src/pdf/table_layout.rs:565`); `shd` (solid + hatch/stripe patterns). Gaps
  below, grouped by priority.
- **17.5 Custom Markup** — 2026-06-18 23:12 — three sub-forms: smart tags
  (17.5.1.1–17.5.1.10), custom XML markup (17.5.1.x), structured document tags /
  content controls (17.5.2.x). For a static PDF these are mostly *transparent
  wrappers* — the deliverable is descending into the wrapper and rendering the
  content inside; the property bags (`sdtPr`, `smartTagPr`, `customXmlPr`) are
  metadata with no print impact and are correctly ignorable. HANDLED:
  block-level `w:sdt` descends into `w:sdtContent` (`src/docx/mod.rs:536`,
  `collect_block_nodes`); run-level `w:sdt` + `w:smartTag` descend into their
  content (`src/docx/runs.rs:615,619`, `collect_run_nodes`); row-level `w:sdt`
  around a `w:tr` works transitively because table-row collection runs through
  `collect_block_nodes` (`src/docx/tables.rs:326`). Corpus: `w:sdt` in 4
  fixtures, `w:smartTag` in 5 — both render. Gaps below.
- **17.6 Sections** — 2026-06-18 23:20 — section properties parsed in
  `src/docx/sections.rs` (`parse_section_properties`), model in `src/model/mod.rs`
  (`SectionProperties`, `ColumnsConfig`/`ColumnDef`, `DocGridType`,
  `SectionBreakType`). Full 17.6 element list: 17.6.1 bidi, 17.6.2 bottom,
  17.6.3 col, 17.6.4 cols, 17.6.5 docGrid, 17.6.6 formProt, 17.6.7 left,
  17.6.8 lnNumType, 17.6.9 paperSrc, 17.6.10 pgBorders, 17.6.11 pgMar,
  17.6.12 pgNumType, 17.6.13 pgSz, 17.6.14 printerSettings, 17.6.15 right,
  17.6.16 rtlGutter, 17.6.17/.18/.19 sectPr (final/section/previous),
  17.6.20 textDirection, 17.6.21 top, 17.6.22 type, 17.6.23 vAlign.
  WELL COVERED: `cols`/`col` (num, equalWidth, space, sep, child col w/space —
  `sections.rs:73`); `docGrid` (linePitch, type — `:38`); `pgMar`
  (top/bottom/left/right/header/footer/gutter — `:21`); `pgSz` (w, h — `:19`);
  `rtlGutter` (gutter side — `:32`); `type` (continuous/oddPage/evenPage/nextPage
  — `:62`); `pgNumType` start+fmt (`:54`); the three `sectPr` placements all parse
  (final in body, per-paragraph, and `sectPrChange/sectPr` harmlessly ignored).
  Gaps below, grouped by priority. NOTE: every section-level instance of the
  unhandled props that DOES appear in the corpus carries a default/no-op value
  (`textDirection=lrTb`, `bidi=0`, `formProt=false`, `paperSrc`=tray) — so these
  gaps are *latent*: output is correct today only because no fixture sets a
  non-default value.
- **17.7 Styles** — 2026-06-18 23:40 — style definitions parsed in
  `src/docx/styles.rs` (1087 lines). WELL COVERED: `docDefaults` (`rPrDefault`
  17.7.5.3 + `pPrDefault` 17.7.5.2, parsed `styles.rs:525`); `style` (17.7.4.17)
  for `type="paragraph"` / `"character"` / `"table"` (`styles.rs:603`); `styleId`,
  `default="1"` (default-paragraph-style detection), `name` (17.7.4.9 → STYLEREF
  map), `basedOn` (17.7.4.3) **for paragraph styles** with full
  inheritance-chain resolution (`resolve_based_on`, `styles.rs:916`); paragraph-
  style `pPr` (spacing/borders/shd/jc/keep*/ind/tabs/numPr/outlineLvl/snapToGrid/
  autoSpace*) + `rPr` (b/i/u/strike/caps/color/sz/rFonts/kern/spacing/w14 effects);
  character-style `rPr`; table-style base `tblBorders` + base `rPr`
  (sz/font/b/i) + `tblStylePr` conditional formatting (firstRow/lastRow/firstCol/
  lastCol/band1-2Horz/Vert/corners — `styles.rs:838`). Gaps below, grouped by
  priority. NOTE: the chapter's UI/editing-only metadata (latent styles, qFormat,
  semiHidden, uiPriority, hidden, locked, aliases, autoRedefine, personal*, rsid,
  `next`, `link`) is correctly absent — none of it has a print representation; see
  the record-only list.
- **17.8 Fonts** — 2026-06-18 23:54 — font table parsed in
  `src/docx/embedded_fonts.rs` (`parse_font_table`), model in `src/model/mod.rs`
  (`FontTable`/`FontTableEntry`/`FontFamily`), substitution + embedded-font use in
  `src/fonts/mod.rs`. Full 17.8 element list: 17.8.1 Embedded Font Obfuscation
  (algorithm), 17.8.2 Font Substitution (descriptive), 17.8.3.1 altName,
  17.8.3.2 charset, 17.8.3.3 embedBold, 17.8.3.4 embedBoldItalic,
  17.8.3.5 embedItalic, 17.8.3.6 embedRegular, 17.8.3.7 embedSystemFonts (settings),
  17.8.3.8 embedTrueTypeFonts (settings), 17.8.3.9 family, 17.8.3.10 font,
  17.8.3.11 fonts, 17.8.3.12 notTrueType, 17.8.3.13 panose1, 17.8.3.14 pitch,
  17.8.3.15 saveSubsetFonts (settings), 17.8.3.16 sig.
  WELL COVERED: `fonts` root + `font @name` (`embedded_fonts.rs:88`); `altName`
  parsed (`:98`) AND consumed first in substitution (`fonts/mod.rs:418`); `family`
  (roman/swiss/modern/script/decorative/auto) parsed (`:101`) AND consumed via
  `family_fallback`; `embedRegular`/`embedBold`/`embedItalic`/`embedBoldItalic`
  (`r:id` + `@fontKey`) parsed, font bytes extracted, deobfuscated, and used with
  priority over system fonts (`fonts/mod.rs:173`); 17.8.1 obfuscation algorithm
  implemented (`deobfuscate_font` + `parse_guid_to_bytes`, ST_Guid §22.9.2.4, XOR
  first 32 bytes with reversed GUID); 17.8.2 substitution chain (altName → name
  list → alias → family/CJK fallback). Gaps below. NOTE: `analyze-fixtures --grep`
  only searches `document.xml`+`styles.xml`, so it cannot confirm corpus presence
  of any `fontTable.xml`/`settings.xml` element below — corpus impact unverified by
  that tool.
- **17.9 Numbering** — 2026-06-18 23:58 — numbering definitions parsed in
  `src/docx/numbering.rs` (594 lines), model structs (`LevelDef`, `NumberingInfo`,
  `ListLabelInfo`) local to that file, label resolution in `parse_list_info`,
  rendered in `src/pdf/layout.rs`. Style→numbering link via paragraph-style `numPr`
  (`src/docx/styles.rs:679`, `num_id`/`num_ilvl` on `ParagraphStyle`). Full 17.9
  element list: 17.9.1 abstractNum, .2 abstractNumId, .3 ilvl, .4 isLgl, .5 lvl
  (override def), .6 lvl (def), .7 lvlJc, .8 lvlOverride, .9 lvlPicBulletId,
  .10 lvlRestart, .11 lvlText, .12 multiLevelType, .13 name, .14 nsid, .15 num,
  .16 numbering, .17 numFmt, .18 numId, .19 numIdMacAtCleanup, .20 numPicBullet,
  .21 numStyleLink, .22 pPr, .23 pStyle, .24 rPr, .25 start, .26 startOverride,
  .27 styleLink, .28 suff, .29 tmpl. WELL COVERED: `abstractNum`/`abstractNumId`/
  `num`/`numId`/`numbering` structure; `ilvl` (via `numPr`); `lvl` definitions +
  `lvlOverride`/`lvl` full level redefinition (`numbering.rs:162`); `start` +
  `startOverride` (counter reset, `:357`); `lvlText` with `%N` placeholder
  substitution; `numFmt` (8 of ~60 formats — see gap); `suff` (tab/space/nothing,
  consumed `:421`); `numStyleLink`→`styleLink` chain resolution (`:194`); partial
  `pPr` (only `ind` start/left/hanging + `tabs`) and partial `rPr` (only
  rFonts/sz/b/color) within `lvl`; default level-restart (deeper counters reset on
  return to a higher level, `:348`). Gaps below, grouped by priority. NOTE:
  `analyze-fixtures --grep` searches only `document.xml`+`styles.xml`, NOT
  `numbering.xml`, so it cannot confirm corpus presence of any `numbering.xml`
  element below (same caveat as 17.8); `w:numPr` itself appears in 29 fixtures.
- **17.10 Headers and Footers** — 2026-06-19 — full element list: 17.10.1
  evenAndOddHeaders, .2 footerReference, .3 ftr, .4 hdr, .5 headerReference,
  .6 titlePg. WELL COVERED: `headerReference`/`footerReference` all three
  `@w:type` values (default/first/even) parsed with `r:id` relationship
  resolution (`src/docx/sections.rs:123`, `resolve_hf` `:155`); `hdr`/`ftr` part
  content parsed as body-like block-level markup (`parse_header_footer_xml`,
  `src/docx/headers_footers.rs:30`); `titlePg` → `SectionProperties.different_first_page`
  (`sections.rs:52`); `evenAndOddHeaders` settings flag parsed
  (`src/docx/settings.rs:55`) into `Document.even_and_odd_headers`. Per-page
  selection (`resolve_header_for_page`/`resolve_footer_for_page`,
  `src/pdf/header_footer.rs:984`/`1012`) chooses first/even/default and even/odd
  parity uses the LOGICAL page number (honors `pgNumType @start`, computed at
  `src/pdf/mod.rs:3086`) — spec-correct. Corpus: `headerReference`/
  `footerReference` in ~18 fixtures, `titlePg` in 9, `evenAndOddHeaders` present
  — all render. Gaps below are all LOW/latent facets of one §17.10.1 rule.
- **17.11 Footnotes and Endnotes** — 2026-06-19 00:07 — full element list: 17.11.1
  continuationSeparator, .2 endnote (content), .3 endnote (special list in
  endnotePr), .4 endnotePr (doc-wide), .5 endnotePr (section-wide), .6 endnoteRef
  (mark), .7 endnoteReference, .8 endnotes (part root), .9 footnote (special list
  in footnotePr), .10 footnote (content), .11 footnotePr (section-wide), .12
  footnotePr (doc-wide), .13 footnoteRef (mark), .14 footnoteReference, .15
  footnotes (part root), .16 noEndnote, .17 numFmt (endnote), .18 numFmt
  (footnote), .19 numRestart, .20 numStart, .21 pos (footnote, ST_FtnPos), .22 pos
  (endnote, ST_EdnPos), .23 separator. WELL COVERED: note *content* parsed from
  `word/footnotes.xml` (`parse_footnotes`, simple builder) and `word/endnotes.xml`
  (`parse_endnotes`, rich builder) keyed by `w:id`, skipping any note carrying a
  `w:type` attribute (`src/docx/headers_footers.rs:74-275`); `footnoteReference`/
  `endnoteReference` + `footnoteRef`/`endnoteRef` marks parsed into runs
  (`src/docx/runs.rs:1161-1198`) and stored on `Run.footnote_id`/`endnote_id`/
  `is_footnote_ref_mark`/`is_endnote_ref_mark` (`src/model/mod.rs:343`); footnotes
  rendered at page bottom with a hardcoded separator line, endnotes rendered once
  at document end (`src/pdf/footnotes.rs`, `src/pdf/mod.rs:3004-3047`); display
  numbers assigned sequentially in body encounter order
  (`src/pdf/mod.rs:2651-2686`) and substituted into the ref marks. The hardcoded
  footnote position (pageBottom), endnote position (docEnd), and decimal numbering
  all match the spec/Word defaults, so every gap below is LATENT in the current
  corpus: all 6 footnote fixtures ship EMPTY `<w:footnotePr/>`/`<w:endnotePr/>`
  (verified via `docx-inspect … word/settings.xml`) and no fixture uses endnotes
  (`analyze-fixtures --grep "endnoteReference"` → 0). Gaps below, grouped by
  priority.
- **17.12 Glossary Document** — 2026-06-19 00:10 — full element list: 17.12.1
  behavior, .2 behaviors, .3 category, .4 description, .5 docPart, .6 docPartBody,
  .7 docPartPr, .8 docParts, .9 gallery, .10 glossaryDocument (root), .11 guid,
  .12 name (category name), .13 name (entry name), .14 style, .15 type, .16 types.
  The glossary document (`word/glossary/document.xml`) is a **supplementary
  document story** — a store of building blocks / AutoText / Quick Parts / cover
  pages / SDT placeholder text that is NOT part of the main document flow and is
  inserted into the body only when a user/app explicitly drops a building block.
  Word's own PDF export does not render glossary content directly. **CORRECTLY
  IGNORED**: `grep -rni "glossary\|docPart" src/` → 0 hits — docxside never opens
  `word/glossary/*`, so all 16 elements are unparsed, which is the right behavior
  for a static render. The parser only ever reads `word/document.xml` as the body
  story. No implementable rendering gap exists for this chapter; see the single
  LATENT cross-reference note below. Corpus: 3 fixtures ship a glossary part
  (`education_consultant_posting`, `zimbabwe_gold_mining_judgment`,
  `alpharetta_school_governance_council`) and all three contain only `bbPlcHdr`
  (SDT-placeholder) building blocks like `[Type the company name]` /
  `[Type the document title]` — none is rendered, correctly.
- **17.13 Annotations** — 2026-06-19 00:25 — three annotation storage methods
  (Inline / "Cross Structure" / Properties; §17.13.1/.2/.3 are intro prose with no
  standalone elements). Element-bearing subclauses: 17.13.4 Comments, 17.13.5
  Revisions, 17.13.6 Bookmarks, 17.13.7 Range Permissions. WELL COVERED:
  **Comments (17.13.4)** parsed from `word/comments.xml` (`src/docx/comments.rs`
  `parse_comments`: id/author/initials/text), anchored via `commentRangeStart`/
  `commentRangeEnd` + `commentReference`-mark detection (`src/docx/runs.rs:553`,
  `:569-593`), modelled as `Comment` + `Run.comment_ids` (`src/model/mod.rs:153,360`),
  and RENDERED as Word's right-margin comment pane with a pink highlight on the
  anchored span and ~76% body scaling (`src/pdf/comments.rs`, `src/pdf/assembly.rs`,
  `src/docx/mod.rs:816`) — matches Word's PDF export; tested by handcrafted
  case63/case64. **Revisions (17.13.5)** in Final/accept-all view: inline `w:ins`
  descended-into (shown) and `w:del` skipped (`src/docx/runs.rs:615-618`);
  property-change revisions (`pPrChange`/`rPrChange`/`sectPrChange`/`tcPrChange`/…)
  are nested in property bags the parsers never read → correctly ignored for the
  final view. **Bookmarks (17.13.6)** `bookmarkStart @name` parsed into
  `Paragraph.bookmarks` (`src/docx/paragraph.rs:370-374`) and resolved to page
  positions for PAGEREF / internal-hyperlink targets (`src/pdf/mod.rs:684`,
  `compute_bookmark_positions`); bookmarks are invisible anchors so no glyph
  rendering is needed — complete. Corpus (via `analyze-fixtures --grep`, which scans
  only `document.xml`+`styles.xml`, NOT `word/comments.xml`): literal `<w:ins ` in 1
  fixture, `w:del` in 1, `w:bookmarkStart` in 36, comments in handcrafted
  case63/case64; `w:moveFrom`/`w:moveTo`/`w:permStart`/`w:cellDel`/`w:cellIns`/
  `w:cellMerge`/`w:*Change` in 0 — so every gap below is LATENT in the current
  corpus (NOTE: a bare `--grep "w:ins"` returns 28 because it substring-matches
  `w:instrText`; the real tracked-insertion count is 1). Gaps below, grouped by
  priority.
- **17.14 Mail Merge** — 2026-06-19 00:45 — full element list (all children of
  `w:mailMerge` §17.14.20, stored in the Document Settings part
  `word/settings.xml`; nested groups under `w:odso` §17.14.25 /
  `w:fieldMapData` §17.14.15 / `w:recipients` §17.14.29 / `w:recipientData`
  §17.14.27): 17.14.1 active, .2 activeRecord, .3 addressFieldName, .4
  checkErrors, .5 colDelim, .6 column (mapped), .7 column (unique-value), .8
  connectString, .9 dataSource, .10 dataType, .11 destination, .12
  doNotSuppressBlankLines, .13 dynamicAddress, .14 fHdr, .15 fieldMapData, .16
  headerSource, .17 lid, .18 linkToQuery, .19 mailAsAttachment, .20 mailMerge,
  .21 mailSubject, .22 mainDocumentType, .23 mappedName, .24 name, .25 odso, .26
  query, .27 recipientData (single record), .28 recipientData (incl/excl ref),
  .29 recipients, .30 src, .31 table, .32 type (ODSO data-source type), .33 type
  (merge-field-mapping), .34 udl. **CORRECTLY IGNORED — no implementable
  rendering gap.** Every 17.14 element is mail-merge *connection / configuration
  metadata* (which external data source, how its columns map to MERGEFIELD field
  names, which records to include, what to do with the merged output). It lives
  in `word/settings.xml`, never in the body story, and has **no print
  representation**: the merge is an authoring/runtime operation performed by the
  host app, after which the populated field values are stored as ordinary cached
  field results in `word/document.xml`. Verified: `grep -rniE
  "mailmerge|odso|fieldmapdata|mergefield|mappedname|recipientdata|maindocumenttype"
  src/` → 0 hits (nothing parsed); `parse_settings`
  (`src/docx/settings.rs:30`) never reads `w:mailMerge`. This is the right
  behavior for a static render — Word's own PDF export shows the cached field
  result, not the connection settings. The single cross-reference note (the
  print-relevant MERGEFIELD field *result*, owned by Clause 17.16) is below;
  there are no latent 17.14 gaps because none of these elements can ever change
  rendered output.
- **17.15 Settings** — 2026-06-19 — Document Settings part (`word/settings.xml`),
  parsed in `src/docx/settings.rs` (`parse_settings` → `DocumentSettings`, 64
  lines). The chapter has three groups: **17.15.1 Document Settings** (~94
  elements, .1–.94), **17.15.2 Web Page Settings** (~45 elements, HTML-save only —
  out of scope for PDF), **17.15.3 Language Compatibility Settings** (`compat`
  legacy toggles + `compatSetting`). PARSED TODAY (7 elements only): `defaultTabStop`
  (17.15.1.25 → consumed in tab layout, `src/pdf/layout.rs`/`mod.rs`);
  `evenAndOddHeaders` (~17.15.1.39 → consumed in header/footer selection,
  `src/pdf/header_footer.rs`/`mod.rs`); `gutterAtTop` (17.15.1.49 → consumed in
  `parse_section_properties`); `themeFontLang @eastAsia` (17.15.1.88 → consumed
  in theme East-Asian font resolution, `src/docx/mod.rs:683`); `mirrorMargins`
  (17.15.1.57 → parsed but `#[allow(dead_code)]`, NEVER consumed);
  `autoHyphenation` (17.15.1.10 → stored on `Document.auto_hyphenation` but NEVER
  consumed — no hyphenation engine); `themeFontLang @val` → `Document.default_lang`
  (stored, never consumed). Verified via `grep -rl … src/`: `documentProtection`,
  `writeProtection`, `decimalSymbol`, `listSeparator`, `characterSpacingControl`,
  `compat`/`compatSetting`, `hyphenationZone`, `consecutiveHyphenLimit`,
  `doNotHyphenateCaps`, `noPunctuationKerning`, `printFractionalCharacterWidth`,
  `doNotExpandShiftReturn`, `printTwoOnOne`, `bookFold*`, `drawingGrid*`,
  `defaultTableStyle`, `trackChanges`, `savePreviewPicture`, `zoom`, `proofState`
  → **0 hits each**. Corpus presence confirmed by inspecting `word/settings.xml`
  in all 53 scraped fixtures with `docx-inspect`: `w:compat` in **all 53**;
  `characterSpacingControl`+`decimalSymbol` in ~51; `hyphenationZone` in ~25;
  `drawingGrid*` in ~10; `noPunctuationKerning` in ~6; `autoHyphenation`,
  `doNotHyphenateCaps`, `doNotExpandShiftReturn` in ~3 each; `documentProtection`
  in 2; `mirrorMargins`, `bookFoldPrinting`, `displayBackgroundShape`,
  `savePreviewPicture` in 1 each (`traditional_skills_job_form` carries the rare
  ones). Gaps below, grouped by priority. NOTE: the vast majority of 17.15.1's 94
  elements are authoring/UI/round-trip metadata with no print representation
  (spell/grammar state, doc variables, style-pane filters, save options, e-mail
  envelope, etc.) and are correctly ignorable — see the record-only list.
- **17.16 Fields and Hyperlinks** — 2026-06-19 06:43 — three families:
  (a) **complex fields** `fldChar` (§17.16.18) / `instrText` (§17.16.23) and
  **simple fields** `fldSimple` (§17.16.19); (b) the **hyperlink** element
  (§17.16.22); (c) **form fields** `ffData` (§17.16.17) + its ~25 child
  property elements (§17.16.7–.34) used by FORMTEXT/FORMCHECKBOX/FORMDROPDOWN.
  Field parsing lives in `src/docx/runs.rs` (`collect_run_nodes` §623 unwraps
  `fldSimple`/`hyperlink`/`sdt`; the main run-builder §1000–1063 maintains a
  `FieldFrame` stack for `fldChar`/`instrText`/`separate`/`end`), the
  `FieldCode` model enum is in `src/model/mod.rs:445`, and dynamic-field
  re-evaluation + hyperlink link-annotation emission are in `src/pdf/` (mod.rs,
  layout.rs `LinkAnnotation`, assembly.rs §63–103). WELL COVERED: complex+simple
  field STRUCTURE (begin/separate/end nesting, instruction-region vs result-region
  visibility, nested fields as arguments — `is_dynamic_field`/`FieldFrame`
  `visible`/`seen_sep`); **all non-dynamic field codes render their cached result
  text** (TOC body, REF/cross-ref text, DATE/TIME values, MERGEFIELD merged
  values, etc.) because result-region `w:t` after `separate` flows to
  `pending_text` — correct for a static PDF; the 4 **dynamic** field codes
  PAGE, NUMPAGES, STYLEREF, PAGEREF (all under §17.16.5) are re-evaluated at render
  time against docxside's own pagination (`FieldCode::Page/NumPages/StyleRef/
  PageRef`, consumed in `src/pdf/header_footer.rs:31`, `src/pdf/mod.rs:1090`);
  the **`w:hyperlink` element** (§17.16.22) is fully handled — `r:id`→external
  URL and `@anchor`→`#bookmark` both parsed (`runs.rs:594`) and emitted as PDF
  Link annotations (external = URI action, internal = GoTo to the resolved
  bookmark position, `assembly.rs:75–99`); underline styling forced for linked
  runs (`layout.rs:1839`). Corpus: complex fields (`fldChar`) in 6 fixtures, TOC
  in 3, PAGEREF/REF cross-refs very heavy (croatian 47, uk_lease 89+288),
  `w:hyperlink` element in 14 fixtures, HYPERLINK field code in 1
  (`traditional_skills_job_form`); form fields (`ffData`/FORMTEXT/FORMCHECKBOX/
  FORMDROPDOWN/MERGEFIELD/`fldData`) in **0** fixtures (verified via
  `analyze-fixtures --grep`, which scans only `document.xml`+`styles.xml`).
  Several field-heavy fixtures are among the worst failing (croatian 7.6%,
  uk_commercial_lease 9.4%, traditional_skills_job_form 15.5%) — fields are a
  plausible contributor but not isolated here. Gaps below, grouped by priority.
- **17.18 Simple Types** — 2026-06-19 06:48 — §17.18 is the *complete catalog of
  the ~110 simple types dedicated to WordprocessingML* (chunk 4402: "This is the
  complete list of simple types dedicated to WordprocessingML"). It defines **no
  renderable elements** — every entry is a value domain (an enumeration, or a
  scalar/measure restriction) referenced by an attribute that belongs to an
  element already audited in 17.2–17.16. So this pass is a *value-domain* audit,
  not an element audit: for each enumeration whose owning element docxside
  actually parses, does the code accept the FULL set of allowed values, or does
  it truncate to a subset? **Audit method**: spot-checked the high-traffic
  enumerations against their parse sites in `src/docx/`. SCALAR/MEASURE simple
  types need no enumeration audit — `ST_DecimalNumber`, `ST_UnsignedDecimalNumber`,
  `ST_TwipsMeasure`, `ST_SignedTwipsMeasure`, `ST_HpsMeasure`, `ST_SignedHpsMeasure`,
  `ST_HexColor`, `ST_LongHexNumber`/`ST_ShortHexNumber`/`ST_UcharHexNumber`,
  `ST_String`, `ST_Guid`, `ST_OnOff`, `ST_DateTime`, `ST_PixelsMeasure`,
  `ST_MeasurementOrPercent` etc. are parsed inline at each consumption site
  (`twips_to_pts`, `parse_hex_color`, `wml_bool`, the §17.8.1 GUID decode) and
  carry no "which-values-are-supported" question. ENUMERATION CONFIRMED COMPLETE:
  `ST_HighlightColor` (§17.18.40) — all 16 named colors mapped in `highlight_color`
  (`src/docx/mod.rs:208-228`), unknown/`none` → `None` (correct). NEW
  value-domain truncations found (3, below); several more value-domain gaps were
  already captured under earlier chapters and are cross-referenced (record-only)
  rather than duplicated.
- **20.1 DrawingML Main** — 2026-06-19 — the `a:` DrawingML core, consumed by
  shapes/textboxes (`src/docx/textbox.rs`), inline/floating pictures
  (`src/docx/images.rs`), grouped shapes (`src/docx/group.rs`), WordArt
  (`src/docx/wordart.rs`), SmartArt (`src/docx/smartart.rs`), and the preset/custom
  geometry engine (`src/geometry/*`). Color resolution + transforms centralised in
  `src/docx/color.rs`. WELL COVERED: transform `a:xfrm` (off/ext/rot/flipH/flipV +
  group `chOff`/`chExt`); preset geometry `a:prstGeom`/`a:avLst`/`a:gd` and custom
  geometry `a:custGeom` (full path command set via the geometry engine); solid fill
  `a:solidFill` and linear gradient `a:gradFill` (`a:gsLst`/`a:gs`/`a:lin`); color
  model `a:srgbClr`/`a:schemeClr`/`a:prstClr` (+ `a:sysClr` in theme parsing,
  `styles.rs:319`) with transforms `tint`/`shade`/`alpha`/`lumMod`/`lumOff`/`satMod`/
  `hueOff`/`satOff`; line `a:ln` (width + solid/noFill color) with arrowheads
  `a:headEnd`/`a:tailEnd` (5 types); effects `a:effectLst` → `outerShdw`/`innerShdw`/
  `glow`/`reflection`/`softEdge` (on pictures); text body `a:bodyPr`
  (anchor/wrap/insets) + `a:prstTxWarp` WordArt warps + `a:normAutofit`/`a:spAutoFit`;
  blip `a:blip @r:embed`. Gaps below, grouped by priority. NOTE: the DrawingML *text*
  sub-elements (`a:bodyPr`, `a:lstStyle`, `a:p`/`a:pPr`/`a:r`/`a:rPr`/`a:t`) are
  formally defined in clause **21.1.2** (DrawingML — Text) but are audited here per
  this file's scope grouping; `analyze-fixtures --grep` searches `document.xml`, where
  inline/anchored DrawingML lives, so corpus counts below ARE reliable for `a:`
  markup (unlike the chart/numbering parts).
- **20.2 DrawingML Picture** — 2026-06-19 — the `pic:` picture container
  (CT_Picture, §20.2.2.5) and its three children: `pic:nvPicPr` (§20.2.2.4 =
  `cNvPr` §20.2.2.3 + `cNvPicPr` §20.2.2.2), `pic:blipFill` (§20.2.2.1,
  CT_BlipFillProperties), `pic:spPr` (§20.2.2.6, CT_ShapeProperties). Located by
  descendant search in `src/docx/images.rs` (`find_pic_sp_pr` `:106`,
  `find_blip_embed` `:323`) and consumed in all four picture paths: inline
  (`images.rs:628`), floating/anchored (`:540`), in-group (`src/docx/group.rs:306`
  `emit_pic`), and paragraph-height measurement (`images.rs:713`); image bytes
  pulled by `read_image_from_zip`/`_extra` (`:271`/`:281`, WMF→raster + EMF
  supported). WELL COVERED: `pic:pic` container (18 corpus fixtures);
  `pic:blipFill/a:blip @r:embed` (blip resolved to image bytes via `find_blip_embed`);
  `pic:spPr` (§20.2.2.6) — `a:xfrm @rot` rotation incl. scene3d-camera fallback
  (`parse_image_rotation` `:120`), outline `a:ln` (`parse_pic_outline` `:141`),
  effects `a:effectLst` → outerShdw/softEdge/glow/innerShdw/reflection
  (`parse_pic_effects` `:214`), and preset/custom-geometry clip path
  (`clip_geometry`, reuses `textbox::parse_shape_geometry`). Gaps below, grouped by
  priority. NOTE: `analyze-fixtures --grep` scans `document.xml`, where inline
  DrawingML pictures live, so the corpus counts below ARE reliable for `pic:`/`a:`
  picture markup.
- **20.4 DrawingML WordprocessingML Drawing** — 2026-06-19 09:02 — the `wp:`
  namespace that anchors/positions DrawingML objects inside the run-level
  `w:drawing` (§17.3.3.9). Two containers: `wp:inline` (CT_Inline, §20.4.2.8) and
  `wp:anchor` (CT_Anchor, §20.4.2.3). Parsing lives in `src/docx/images.rs`:
  inline vs anchor detected at `:470` (`is_wpd_drawing`), positioning in
  `parse_anchor_position` (`:336`), wrap in `parse_wrap_type` (`:384`) +
  `parse_wrap_polygon` (`:419`), sizing in `extent_dimensions` (`:41`), inline
  layout space in `inline_extra_height` (`:32`); floating-image fields populated at
  `:497`/`:557`; group/canvas anchors via `src/docx/group.rs`; floating textboxes
  via `src/docx/textbox.rs:920`. Model enums `HorizontalPosition`/`VerticalPosition`/
  `HRelativeFrom`/`VRelativeFrom`/`WrapType`/`WrapText` in `src/model/mod.rs:170-217`;
  wrap-polygon vertices on the floating model in `src/model/drawing.rs:94`.
  Render/placement in `src/pdf/positioning.rs` + `src/pdf/mod.rs` (z-order by
  `relativeHeight`), wrap geometry in `src/pdf/mod.rs:285-1450`. WELL COVERED:
  `wp:inline` + `wp:anchor`; `wp:extent @cx/@cy` (EMU→pt); `wp:effectExtent
  @l/t/r/b` (inline line-height contribution); `wp:positionH`/`wp:positionV` with
  `@relativeFrom` (page/margin/column for H; page/margin/topMargin/paragraph for V)
  and children `wp:align` (left/center/right; top/center/bottom) + `wp:posOffset`
  (absolute EMU offset); distance-from-text `@distT/@distB/@distL/@distR`;
  `@behindDoc` (behind-text z-layer); `@relativeHeight` (z-order among floats);
  `wp:docPr @hidden` (drawing suppressed, `:478`); all five wrap modes —
  `wp:wrapNone`/`wp:wrapSquare`/`wp:wrapTight`/`wp:wrapThrough`/`wp:wrapTopAndBottom`
  — with `@wrapText` bothSides/left/right/largest (largest = larger-free-side,
  `src/pdf/mod.rs:1417`) and `wp:wrapPolygon`/`wp:start`/`wp:lineTo` vertices for
  tight/through. Corpus (`analyze-fixtures --grep`, scans `document.xml` where
  `wp:` lives — counts reliable): `wp:anchor` 15, `wp:inline` 13, `wp:wrapNone` 11,
  `wp:wrapSquare` 6, `wp:wrapTight`/`wp:wrapTopAndBottom`/`wp:wrapPolygon` 1 each,
  `wp:wrapThrough` 0; `relativeFrom` only `column`(9)/`paragraph`(10)/page/margin in
  the wild; `behindDoc="1"` 4. Gaps below, grouped by priority. NOTE: most gaps are
  LATENT — every exotic `@relativeFrom`, `inside`/`outside` align, and `simplePos="1"`
  value is absent from the corpus (0 fixtures); the one non-default value actually
  present is `layoutInCell="0"` (2 fixtures).
- **21.2 DrawingML Charts** — 2026-06-19 — the `c:` chart namespace
  (`http://schemas.openxmlformats.org/drawingml/2006/chart`), root `c:chartSpace`.
  Parsing in `src/docx/charts.rs` (`parse_chart_space` `:225`, `detect_chart_type`
  `:263`), model in `src/model/chart.rs` (`Chart`/`ChartType`/`ChartSeries`/
  `ChartAxis`/`ChartLegend`/`MarkerSymbol`/`BarGrouping`/`InlineChart`), triggered
  from `src/docx/images.rs:650` (`find_chart_ref` matches `a:graphicData @uri` =
  the chart URI, resolves the `c:chart @r:id` → `word/charts/chartN.xml`).
  Rendering: cartesian (bar/line/area/scatter/bubble) + radar in `src/pdf/charts.rs`
  (1357 lines), pie/doughnut in `src/pdf/charts_radial.rs`, shared legend in
  `src/pdf/chart_legend.rs`. WELL COVERED: `chartSpace`/`chart`/`plotArea` skeleton;
  8 chart-type elements — `barChart` (`@barDir` col/bar, `@grouping`
  clustered/stacked/percentStacked **all three RENDERED** incl. stacked geometry,
  `@gapWidth`), `lineChart`, `pieChart`, `pie3DChart`→Pie, `areaChart`,
  `doughnutChart` (`@holeSize`), `radarChart`, `scatterChart`, `bubbleChart`;
  `c:ser` with `c:tx`→`strRef`/`strCache` series name, `c:spPr` solid-fill + line
  color + `a:alpha`, `c:marker @symbol` (9 symbols), `c:val`/`c:yVal`/`c:xVal`/
  `c:bubbleSize`→`numRef`/`numCache` (idx-indexed `c:pt`/`c:v`) and `c:cat`→`strRef`/
  `strCache` category labels; `c:numCache/c:formatCode` → `val_format_code` consumed
  in tick-label formatting (`format_tick_label` `pdf/charts.rs:47`); `c:catAx`/
  `c:valAx` `@delete`, `c:majorGridlines` color, axis line color, two `c:valAx` for
  scatter/bubble; `c:legend` `c:legendPos` r/b/t/l. Corpus: charts in 6 fixtures
  (cases 29/30/31/52/53, sample500kB); feature elements actually present —
  `c:scaling`+`c:orientation`, `c:tickLblPos`, `c:minor/majorTickMark`, `c:crossesAt`,
  `c:overlap` (case52), `c:varyColors`, `c:holeSize`. NO `c:title`, `c:dLbls`,
  explicit `c:min`/`c:max` scaling, 3D, trendlines, or combo charts appear in the
  corpus — so most gaps below are LATENT. (NOTE: `analyze-fixtures --grep` scans only
  `document.xml`+`styles.xml`, NOT `word/charts/*`, so corpus presence here was
  verified by direct `unzip -p … word/charts/chartN.xml` scans, not that tool.)
  Gaps below, grouped by priority.
- **22.1 Math (OMML)** — 2026-06-19 — the `m:` namespace
  (`http://schemas.openxmlformats.org/officeDocument/2006/math`, §22.1), parsed in
  `src/docx/runs.rs`. Entry points: `parse_runs` descends `m:oMathPara` →
  `m:oMath` (`runs.rs:633-640`); each `m:oMath` zone is fed to `omath_to_runs`
  (`runs.rs:787-903`), which **linearizes** the math tree into ordinary text runs
  (reusing the normal font/sub-superscript pipeline) — there is NO two-dimensional
  math layout. Display-math justification handled by `display_math_alignment`
  (`runs.rs:764`, reads `m:oMathPara/m:oMathParaPr/m:jc`, defaults Center =
  centerGroup). Synthesized runs carry `Run.is_math=true` (`src/model/mod.rs:364`),
  which clamps line-height/ascent so a tall math font doesn't balloon the line
  (`src/pdf/layout.rs:2014`, `src/pdf/table_layout.rs:457`). HANDLED structurally:
  `m:r`/`m:t` (text, reads `w:rPr` for fonts/size); `m:sSub`/`m:sSup`/`m:sSubSup`
  (true sub/superscript via VertAlign); `m:f` fraction (→ `num/den`); `m:rad`
  radical (→ `√(e)`); `m:d` delimiter (→ `(e)`); `m:nary` (honors `m:naryPr/m:chr`,
  default `∫`, with sub/sup limits). All other math objects (`m:acc`, `m:bar`,
  `m:groupChr`, `m:limLow`/`m:limUpp`, `m:func`, `m:m` matrix, `m:eqArr`, `m:sPre`,
  `m:box`, `m:borderBox`, `m:phant`) hit the catch-all `_ => omath_to_runs(...)`
  branch (`runs.rs:900`) — their *text* is collected in document order but all
  positioning/decoration semantics are dropped. Document-level `m:mathPr`
  (§22.1.2.61, in `word/settings.xml`) is NOT parsed (`src/docx/settings.rs` has no
  math handling). Corpus: `analyze-fixtures --grep "oMath"` → **0 fixtures** (it
  unzips and scans `document.xml`, where OMML lives, so this count IS reliable) —
  so EVERY gap below is LATENT today; the math path exists for real-world docs but
  no committed fixture exercises it. Gaps below, grouped by priority.
- **22.6 Bibliography** — 2026-06-19 — the `b:` namespace
  (`http://schemas.openxmlformats.org/officeDocument/2006/bibliography`, §22.6),
  stored in the **Bibliography part** (`customXml/itemN.xml`, related from the Main
  Document via `relationships/customXml`, §22.6 intro / part-type table). Full
  element list (all §22.6.2.x): `Sources` (.60, root: `@SelectedStyle`/`@StyleName`/
  `@URI`/`@Version`), `Source` (.59, one bibliography entry), `SourceType` (.61),
  `Tag` (.65), `Guid` (.34), `LCID` (.44), `RefOrder` (.58), and the data-bearing
  children `Title`/`Year`/`Month`/`Day`/`City`/`Publisher`/`Pages`/`Volume`/`Edition`/
  `Comments`/`StandardNumber`/`DOI`/`URL`/etc. (.1–.66), plus the contributor block
  `Author` (.4, container) → role wrappers `Author`/`Editor`/`Translator`/`Artist`/
  `Compiler`/`Composer`/`Conductor`/`Counsel`/`Director`/`Interviewee`/`Interviewer`/
  `Inventor`/`Performer`/`ProducerName`/`Writer`/`BookAuthor` → `NameList` (.49) →
  `Person` (.55) → `Last`/`First`/`Middle` (.43/.31/.46), and the corporate-author
  form `Corporate` (.21). **CORRECTLY IGNORED — no implementable rendering gap.**
  The Bibliography part is pure *citation-source data storage* with **no direct
  print representation** (directly analogous to 17.14 Mail Merge and 17.12
  Glossary): it is the database an authoring app reads to *generate* citations and
  the bibliography list. The actual rendered output comes from the **CITATION**
  (§17.16.5.8) and **BIBLIOGRAPHY** field codes, whose results an app computes from
  the `b:Source` data + the XSL named by `@SelectedStyle` and then caches as
  ordinary `w:p`/`w:t` runs in the body story — and docxside already renders those
  cached results via Clause 17.16's non-dynamic-field path (CITATION/BIBLIOGRAPHY
  are NOT in `is_dynamic_field`, `src/docx/runs.rs:22`, so their post-`separate`
  result text flows to `pending_text` like any static field). Verified: the part is
  never opened — `grep -rniE "bibliograph|citation|sourceType|SelectedStyle|
  RefOrder|b:Source" src/` → **0 hits**; the parser only reads `word/document.xml`
  as the body story, never `customXml/*`. Corpus: 125 fixtures ship a `b:Sources`
  part but in 123 it is the EMPTY `<b:Sources SelectedStyle="…"/>` placeholder
  python-docx/Word writes by default; only **2** carry real `<b:Source>` entries —
  `tests/fixtures/new/covid_insomnia_mental_health` (68 sources + ONE `BIBLIOGRAPHY`
  field whose formatted result is stored as ~70 cached `<w:p>` reference paragraphs
  after `separate`, which render correctly; Jaccard 0.2148, no in-text CITATION
  fields) and `tests/fixtures/new/japanese_land_development_sign_form` (1 orphan
  source, NO citing field → correctly renders nothing). The single LATENT gap is
  below. NOTE: `analyze-fixtures --grep` reported 0 CITATION/0 BIBLIOGRAPHY because
  it scans only `tests/fixtures/cases` + `scraped` (NOT `tests/fixtures/new/`, where
  the covid fixture lives) AND scans only `document.xml`+`styles.xml` (NOT
  `customXml/*`); corpus presence here was confirmed by direct
  `unzip -p … customXml/item*.xml` / `word/document.xml` scans, not that tool.

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

---

## 17.3 Paragraphs and Rich Formatting — gaps

### HIGH priority (affect common Western docs)

#### 17.3.2.24 `w:position` (Vertically Raised or Lowered Text) — NOT handled
- **Spec**: §17.3.2.24. Run property; `@w:val` is a signed half-point amount by
  which the run's text is raised (+) or lowered (−) relative to the surrounding
  baseline, **without** changing font size. Distinct from `w:vertAlign`
  (super/subscript, which also shrinks the glyph). Common in Western docs for
  fine baseline nudging, chemical formulae, logotypes.
- **Today**: not parsed. `grep "\"position\"" src/` → none. `RunFormat`/`Run`
  carry `vertical_align` (super/sub) but no raise/lower offset field.
- **What exists vs missing**: super/subscript via `vertAlign` is fully handled
  (`src/docx/runs.rs:394`, rendered in `src/pdf/layout.rs`); the independent
  `w:position` baseline shift is entirely absent.
- **Likely owner**: add `position_pt: f32` to `Run` (`src/model/mod.rs`), parse
  `w:position @val / 2.0` in `RunFormat::resolve_run_format`
  (`src/docx/runs.rs`), apply as a baseline Y offset where runs are placed in
  `src/pdf/layout.rs`.

#### 17.3.1 `w:textAlignment` (Vertical Char Alignment On Line) — NOT handled
- **Spec**: paragraph property in 17.3.1; `@w:val` = ST_TextAlignment
  (§17.18.91) = `top` | `center` | `baseline` | `bottom` | `auto`. Controls how
  characters of differing font sizes on the same line are aligned vertically
  (e.g. align cap-tops vs. baselines when a big and small run share a line).
- **Today**: not parsed (`grep "\"textAlignment\"" src/` → none; only
  `textDirection` appears, and only in `src/docx/tables.rs`). Line building in
  `src/pdf/layout.rs` always aligns runs on the baseline.
- **What exists vs missing**: baseline alignment is the implicit default and
  works; the other four modes (notably `top`/`center`/`bottom`) are ignored.
- **Likely owner**: add a field to `Paragraph` (`src/model/mod.rs`), parse in
  `src/docx/paragraph.rs`, honor in the line-box vertical placement in
  `src/pdf/layout.rs`.

#### Run text-effect toggles — `w:outline` / `w:shadow` / `w:emboss` / `w:imprint` / `w:effect` — NOT handled
- **Spec**: §17.3.2.23 `outline` (Display Character Outline), §17.3.2.31
  `shadow`, §17.3.2.13 `emboss`, §17.3.2.18 `imprint`, plus `w:effect` (animated
  text effect, ST_TextEffect). The first four are boolean toggle properties
  (§17.17.4); mutually exclusive per spec.
- **Today**: none parsed. IMPORTANT — `src/docx/wordart.rs:108` reads
  `find_w14(rpr, "shadow")` which is the **w14 DrawingML** glow/shadow/outline
  effect set (`Run.text_shadow/text_outline/...`), NOT these simple `w:` run
  toggles. So a plain `<w:rPr><w:shadow/></w:rPr>` or `<w:outline/>` renders as
  normal text today.
- **What exists vs missing**: fancy w14 text effects exist; the legacy
  Word 97-style outline/shadow/emboss/imprint toggles do not. `w:effect`
  (blinking/sparkle text) has no print representation — Word's PDF export
  renders the base text only, so it can likely be parsed-and-ignored.
- **Likely owner**: add booleans to `Run` (`src/model/mod.rs`), parse via
  `wml_bool` in `src/docx/runs.rs`, render in `src/pdf/layout.rs` (outline =
  stroke text render mode; emboss/imprint = light/dark offset shadow).

### MEDIUM priority (RTL / complex-script cluster — entirely absent)

#### 17.3.1.6 `w:bidi` + 17.3.2.30 `w:rtl` + complex-script run props — NOT handled
- **Spec**: §17.3.1.6 `bidi` (paragraph right-to-left layout: affects `ind`,
  `jc`, `tab`, `textDirection`); §17.3.2.30 `rtl` (run-level RTL base direction,
  drives Unicode bidi reordering, UAX #9); complex-script run twins
  §17.3.2.2 `bCs`, §17.3.2.17 `iCs`, §17.3.2.35 `szCs`, §17.3.2.6 `cs`
  (bold/italic/size/force-complex applied to Arabic/Hebrew/etc. text).
- **Today**: none parsed (`grep` for each → none). No bidi reordering anywhere;
  text is laid out strictly left-to-right.
- **What exists vs missing**: LTR layout works; Arabic/Hebrew/Persian documents
  would render with wrong glyph order, wrong indentation side, and ignore the
  `*Cs` bold/italic/size for complex-script runs.
- **Likely owner**: large, cross-cutting — `Run`/`Paragraph` model fields, parse
  in `src/docx/runs.rs` + `src/docx/paragraph.rs`, bidi reordering + RTL line
  placement in `src/pdf/layout.rs`. Consider a dedicated effort; not a one-liner.

### LOW priority (East Asian typography — rare in current corpus)

- **17.3.2.12 `w:em` (Emphasis Mark)** — NOT handled. Marks (dot/comma/circle
  above/below each non-space glyph), ST_Em §17.18.24. `grep "\"em\""` only hits
  `src/docx/alt_chunk.rs` (HTML `<em>`, unrelated). Owner: `src/docx/runs.rs` +
  `src/pdf/layout.rs`.
- **17.3.2.10 `w:eastAsianLayout`** (two-lines-in-one / horizontal-in-vertical) —
  NOT handled. Owner: `src/docx/runs.rs`.
- **17.3.2.15 `w:fitText`** (fit run text to a fixed width) — NOT handled.
  Owner: `src/docx/runs.rs` + `src/pdf/layout.rs`.
- **17.3.1.1 `w:adjustRightInd`** (auto right-indent vs. doc grid) — NOT parsed;
  default true. Minor unless a docGrid char-pitch is set. Owner:
  `src/docx/paragraph.rs`.
- **East Asian line-break rules**: `w:kinsoku`, `w:overflowPunct`,
  `w:topLinePunct`, `w:wordWrap` — NONE parsed. Affect where lines may break in
  CJK text. (`autoSpaceDE`/`autoSpaceDN` ARE already parsed + consumed.) Owner:
  `src/docx/paragraph.rs` + line breaking in `src/pdf/layout.rs`.

### LOW priority (other / niche)

- **17.3.1.5 `w:cnfStyle` (Paragraph Conditional Formatting)** — NOT parsed
  anywhere (`grep "\"cnfStyle\"" src/` → none). Carries the conditional-format
  bitmask (first/last row/col, banding) that selects which `tblStylePr`
  overrides apply to a paragraph/cell. Without it, table-style band/header
  formatting driven through `cnfStyle` is not applied. Owner: `src/docx/tables.rs`
  / `src/docx/paragraph.rs`; pairs with table-style work in 17.4/17.7.
- **17.3.1.16 `w:mirrorIndents`** (swap left/right indent on facing pages) —
  NOT handled. Owner: `src/docx/paragraph.rs` + duplex layout in `src/pdf/`.
- **17.3.1.4 `w:bar`** (paragraph border between facing pages) — NOT modeled;
  `ParagraphBorders` has top/bottom/left/right/between but no `bar`. Niche.
  Owner: `src/model/mod.rs` + `src/docx/mod.rs` border parsing.
- **`w:suppressLineNumbers`, `w:suppressOverlap`, `w:textboxTightWrap`,
  `w:divId`** — NOT parsed; mostly no visual impact for a static PDF (line
  numbering, frame overlap, textbox wrap tightness, HTML div id). Low value.
- **`w:suppressAutoHyphens` (17.3.1.34)** — parsed into
  `Paragraph.suppress_auto_hyphens` (marked `#[allow(dead_code)]`) but never
  consumed in `src/pdf/`; moot until automatic hyphenation itself is implemented
  (`Document.auto_hyphenation` exists but no hyphenation engine).
- **Run effects with no print impact** — `w:noProof` (spell-check),
  `w:webHidden`, `w:specVanish` — safe to leave unparsed for PDF output.

---

## 17.5 Custom Markup — gaps

### HIGH priority — content loss

#### 17.5.1.1/.3/.5/.6 `w:customXml` (block / cell / row / inline custom XML element) — NOT handled → silently DROPS wrapped content
- **Spec**: §17.5.1. `w:customXml` is a transparent wrapper (like `w:sdt`) that
  surrounds content to attach schema-based semantics. Four placements, by where
  the element sits: block-level (around paragraphs/tables), row-level (around a
  `w:tr` inside `w:tbl`), cell-level (around a `w:tc` inside `w:tr`), and
  inline-level (around runs inside a `w:p`). Carries `@w:element` (ST_XmlName,
  §22.9.2.21) + optional `@w:uri`, and an optional child `w:customXmlPr`.
- **Today**: the string `customXml` appears NOWHERE in `src/` (verified by
  `grep -rc customXml src/` → 0). The two descent helpers only special-case
  `w:sdt` / `w:smartTag` / `mc:AlternateContent`:
  - `collect_block_nodes` (`src/docx/mod.rs:531`) pushes any non-`sdt`,
    non-`AlternateContent` child verbatim, then callers filter for `w:p`/`w:tbl`
    (or `w:tr` in `src/docx/tables.rs:326`, `w:tc` at `:338`). A `w:customXml`
    node is pushed as-is and then dropped by those filters → **all content
    wrapped solely in `customXml` disappears** (paragraphs, tables, rows, cells).
  - `collect_run_nodes` (`src/docx/runs.rs:~590`) handles `ins`/`smartTag`/`sdt`/
    `fldSimple`/`AlternateContent`/math but NOT `customXml` → inline-level
    custom-XML-wrapped runs are dropped.
- **What exists vs missing**: the analogous `w:sdt` descent is already
  implemented in both helpers — `customXml` needs the exact same transparent
  descent (recurse into the `customXml` node's children rather than into a
  `sdtContent` child, since `customXml` has no content wrapper; its children ARE
  the content). Ignore `@w:element`/`@w:uri`/`w:customXmlPr` (metadata only).
- **Corpus**: 0 of the current fixtures contain `w:customXml`
  (`analyze-fixtures --grep "w:customXml"` → 0), so no test currently exercises
  or regresses on this — but it is a conformant, in-spec wrapper and any document
  using the Word 2007 "custom XML mapping" feature would lose content.
- **Likely owner**: `collect_block_nodes` in `src/docx/mod.rs` (block + the
  row/cell cases, since table collection routes through it) and
  `collect_run_nodes` in `src/docx/runs.rs` (inline). Mirror the existing `sdt`
  branches.

### LOW priority — no print impact (record-only)

- **17.5.2 `w:sdtPr` sub-properties — parsed-and-ignored, currently fine.** SDT
  *content* renders (via the `sdtContent` descent above), and Word stores the
  resolved display text/picture inside `sdtContent`, so for a static PDF the
  property bag is correctly ignorable. Elements not read (all in `src/docx/`,
  none referenced): `alias` (17.5.2.1), `tag`, `id` (17.5.2.18), `lock`,
  `rPr`, `dataBinding` (17.5.2.6), `comboBox` (17.5.2.5) / `dropDownList`
  (17.5.2.15) / `listItem` (17.5.2.21), `date` (17.5.2.7) / `dateFormat`
  (17.5.2.8) / `lid` / `storeMappedDataAs` (17.5.2.40), `docPartObj` /
  `docPartList` (17.5.2.12) / `docPartGallery`, `picture` (17.5.2.24), `text`,
  `equation` (17.5.2.16), `group` (17.5.2.17), `citation` (17.5.2.4),
  `bibliography`, `placeholder` (17.5.2.25), `temporary`, `sdtEndPr`. No action
  needed unless a future feature wants to re-evaluate a data-bound SDT against
  its custom XML data part (rare; would belong in `src/docx/mod.rs`/`runs.rs`
  next to the existing `sdtContent` descent).
- **17.5.2.39 `w:showingPlcHdr` (contents are placeholder text) — ONE subtlety
  worth a future check.** When set, `sdtContent` holds prompt/placeholder text
  (styled `PlaceholderText`, e.g. "Click here to enter a date…"). The code
  renders `sdtContent` verbatim, so placeholder prompts print. Word's own PDF
  export DOES render content-control placeholder text (it is visible/printed
  gray text), so current behavior is likely correct — but confirm against a
  reference PDF containing an unfilled content control before assuming so.
  Owner if a change is needed: `src/docx/runs.rs` / `src/docx/mod.rs` SDT descent.
- **17.5.1.2/.4/.8/.10 `attr` / `customXmlPr` / `smartTagPr`** — property bags
  for custom XML / smart tags; pure metadata, no print representation. Safe to
  leave unparsed (current state).
- **Custom-markup revision range markers** — `customXmlInsRangeStart/End`,
  `customXmlDelRangeStart/End`, `customXmlMoveFromRangeStart/End`,
  `customXmlMoveToRangeStart/End` (17.5.1.x): mark tracked-change ranges around
  custom XML tags. No visual impact in the accepted/final view docxside renders
  (it already drops `w:del` content and shows `w:ins` content). Leave unparsed.

---

## 17.6 Sections — gaps

### HIGH priority (real visual features; absent from corpus but common in the wild)

#### 17.6.10 `w:pgBorders` + 17.6.2/.7/.15/.21 `w:bottom`/`w:left`/`w:right`/`w:top` (Page Borders) — NOT handled
- **Spec**: §17.6.10 (`pgBorders`) with four child border elements
  §17.6.2 `bottom`, §17.6.7 `left`, §17.6.15 `right`, §17.6.21 `top`. Draws a
  border box around each page. `pgBorders` attributes: `@offsetFrom`
  (ST_PageBorderOffset §17.18.63 = `page` | `text`), `@display`
  (ST_PageBorderDisplay §17.18.62 = `allPages` | `firstPage` | `notFirstPage`),
  `@zOrder` (ST_PageBorderZOrder = `front` | `back`). Each child border is a
  CT_Border: `@val` (ST_Border §17.18.2 — includes plain line styles AND ~160
  "art" image borders like `seattle`, `apples`), `@sz` (eighths of a point),
  `@space` (points, 0–31), `@color`, `@themeColor`, plus art-border `@frame`.
- **Today**: `grep "pgBorders" src/` → 0 hits. No model field, no parse, no
  render. CT_Border parsing already exists for paragraph/table borders
  (`src/docx/mod.rs:335` start/end fallback, `parse_paragraph_borders`) so the
  per-edge attribute decode is reusable; only the page-level container + page
  placement are missing. Plain-line borders are achievable; the ~160 "art"
  image borders would need bitmap assets and are a separate, larger effort.
- **Corpus**: `analyze-fixtures --grep "w:pgBorders"` → 0 fixtures, so no test
  regresses today — but page borders are common on title pages, certificates,
  and form templates.
- **Likely owner**: add `page_borders` to `SectionProperties`
  (`src/model/mod.rs`), parse in `src/docx/sections.rs`, render as stroked
  rectangle(s) in the behind-doc/foreground z-layer in `src/pdf/mod.rs` honoring
  `@offsetFrom` (page-edge vs text-extent inset) and `@display` (which pages).

#### 17.6.23 `w:vAlign` (Vertical Text Alignment on Page) — NOT handled (section level)
- **Spec**: §17.6.23. `sectPr` child; `@w:val` = ST_VerticalJc §17.18.102 =
  `top` (default) | `center` | `both` (justify) | `bottom`. Positions the
  section's body text vertically between the top and bottom margins on each page
  — `center` vertically centers content (classic title-page layout), `both`
  vertically justifies, `bottom` aligns to the bottom margin.
- **Today**: `grep "vAlign" src/` hits only the **table cell** `vAlign`
  (`src/docx/tables.rs:436`) and an unrelated comment in `src/pdf/table.rs:274`.
  The section-level element is not parsed; body text is always top-aligned.
  Distinct element/owner from the cell vAlign that already works.
- **Corpus**: element appears in 15 fixtures but all are cell-level; 0 fixtures
  use it inside `w:sectPr` (verified by scanning `sectPr` segments). Latent gap.
- **Likely owner**: add `vertical_align` to `SectionProperties`
  (`src/model/mod.rs`), parse in `src/docx/sections.rs`, apply as a vertical
  offset / extra-leading distribution of the page's body block in
  `src/pdf/mod.rs` (needs measured body height vs. available text height).

#### 17.6.8 `w:lnNumType` (Line Numbering Settings) — NOT handled
- **Spec**: §17.6.8. `sectPr` child adding line numbers in the margin (legal
  documents, contracts, code listings). Attributes: `@countBy` (number every Nth
  line), `@start` (first number), `@distance` (twips from text to the number),
  `@restart` (ST_LineNumberRestart = `newPage` | `newSection` | `continuous`).
  Interacts with the paragraph property `w:suppressLineNumbers` (17.3.1.35),
  which is also not parsed.
- **Today**: `grep "lnNumType" src/` → 0 hits. Not modeled, parsed, or rendered.
- **Corpus**: 0 fixtures (`analyze-fixtures --grep "w:lnNumType"`).
- **Likely owner**: model field on `SectionProperties` (`src/model/mod.rs`),
  parse in `src/docx/sections.rs`, render numbers in the left (or right, if RTL)
  margin during paragraph/line layout in `src/pdf/layout.rs`; honor per-paragraph
  `w:suppressLineNumbers`.

### LOW priority (latent — correct today only because corpus values are defaults)

#### 17.6.1 `w:bidi` (Right-to-Left Section Layout) — NOT handled
- **Spec**: §17.6.1. Boolean toggle (§17.17.4). When true, section-level layout
  is RTL: text columns populate right-to-left and page numbers/section furniture
  mirror to the right. Affects section geometry only, NOT in-paragraph text
  order (that is the separate `w:bidi`/`w:rtl` 17.3 cluster).
- **Today**: `grep "bidi" src/` hits only border-naming comments
  (`src/docx/mod.rs:335,340`); the section toggle is not parsed.
- **Corpus**: appears at section level in 2 fixtures
  (`traditional_skills_job_form`, `slovak_fuel_pricing_report`) but both set
  `w:val="0"` (off), so no current impact. Pairs with the open 17.3 bidi/RTL
  gap — best tackled together.
- **Likely owner**: parse in `src/docx/sections.rs` (model bool on
  `SectionProperties`); apply to multi-column ordering + margin/gutter mirroring
  in `src/pdf/mod.rs` / `src/pdf/layout.rs`.

#### 17.6.20 `w:textDirection` (Section Text Flow Direction) — NOT handled (section level)
- **Spec**: §17.6.20. `sectPr` child; `@w:val` = ST_TextDirection §17.18.93
  (`lrTb` default, `tbRl`, `btLr`, `lrTbV`, `tbRlV`, `tbLrV`). Non-default values
  rotate/stack the whole section's text flow (vertical CJK layouts).
- **Today**: `grep "textDirection" src/` hits only **table cell** textDirection
  (`src/docx/tables.rs:442`); the section-level element is not parsed. Body text
  always flows `lrTb`.
- **Corpus**: appears at section level in 4 fixtures
  (`polish_building_procurement_spec`, `german_mezzo_soprano_bio`,
  `double-underline`, `sample500kB`) but every instance is `w:val="lrTb"` (the
  default) → currently a no-op. Latent.
- **Likely owner**: `SectionProperties` field (`src/model/mod.rs`), parse in
  `src/docx/sections.rs`, honor in line/glyph placement in `src/pdf/layout.rs`
  (large effort; only needed for genuinely vertical sections).

#### 17.6.12 `w:pgNumType` — PARTIALLY handled (`@chapSep` / `@chapStyle` missing)
- **Spec**: §17.6.12. `@start` and `@fmt` ARE parsed (`src/docx/sections.rs:54`).
  Not parsed: `@chapSep` (ST_ChapterSep — separator glyph between chapter and
  page number) and `@chapStyle` (heading style level that supplies the chapter
  number), used to render numbers like "3-12". Niche.
- **Likely owner**: extend the `pgNumType` parse in `src/docx/sections.rs` and
  the page-number formatting in `src/pdf/mod.rs:3082`.

### NO PRINT IMPACT (record-only — safe to leave unparsed)

- **17.6.6 `w:formProt` (Only Allow Editing of Form Fields)** — editing-protection
  flag; no bearing on a static PDF. Corpus instances all `w:val="false"`. None.
- **17.6.9 `w:paperSrc` (Paper Source Information)** — printer tray selection
  (`@first`/`@other`); no visual impact. Present in `dutch_government_budget_letter`.
- **17.6.14 `w:printerSettings` (Reference to Printer Settings Data)** — r:id to a
  Printer Settings part; no print/layout representation. Leave unparsed.
- **`w:pgSz @orient`** — NOT read, and need not be: per §17.6.13 the `@w` / `@h`
  attributes already carry the post-orientation dimensions (landscape stores
  `w > h`), and `sections.rs:19` uses them directly, so page size is correct
  without consulting `@orient`. Correctly ignorable.
- **Section-level `w:endnotePr` / `w:footnotePr` (17.11.5 / 17.11.11)** — may
  appear as `sectPr` children but are documented under Clause 17.11; defer their
  audit (per-section footnote/endnote numbering + position) to the 17.11 pass.

---

## 17.8 Fonts — gaps

The recommended substitution priority (§17.8.2) is, in descending order:
`sig` → `charset` → `panose1` → `pitch` → `family` → `altName` → `notTrueType`.
docxside currently keys substitution off `altName` first, then `family`/CJK
fallback (`src/fonts/mod.rs:418`) — i.e. only the two *lowest*-priority criteria.
`altName`-first is a deliberate pragmatic choice (it is the document's explicit
mapping; see the comment at `fonts/mod.rs:417`) and works well, but the four
higher-fidelity matchers (`sig`/`charset`/`panose1`/`pitch`) that Word uses to
pick a *visually closest* substitute for an unknown font do not exist. All four
gaps are latent: every fixture either ships the named font, embeds it, or has an
`altName` — none currently force a metrics-based match. They matter when a
document references a font that is absent, un-embedded, and has no `altName`.

### MEDIUM priority (font-substitution fidelity — better matches for missing fonts)

#### 17.8.3.16 `w:sig` (Supported Unicode Subranges and Code Pages) — NOT parsed
- **Spec**: §17.8.3.16. Child of `w:font`; attributes `@usb0`..`@usb3` (Unicode
  subrange bitfield, the OS/2 `ulUnicodeRange1-4`) and `@csb0`/`@csb1` (code-page
  bitfield, OS/2 `ulCodePageRange1-2`), each a hex `ST_LongHexNumber`. This is the
  spec's **#1** (highest-priority) substitution signal — it lets an app pick a
  substitute that covers the same scripts/code pages as the missing font.
- **Today**: `grep "\"sig\"\|usb0\|csb0" src/` → none. Not parsed or stored.
- **What exists vs missing**: nothing today; `FontTableEntry` has no signature
  field. A substitute is chosen purely by name/family, so a missing font could be
  replaced by one lacking the needed scripts (e.g. a Latin-only substitute for a
  font whose `sig` declares Cyrillic/Greek coverage).
- **Likely owner**: add a `sig` field to `FontTableEntry` (`src/model/mod.rs`),
  parse the six hex attrs in `parse_font_table` (`src/docx/embedded_fonts.rs`),
  and add a coverage-aware tie-break to the substitution chain in
  `src/fonts/mod.rs` (compare against candidate fonts' OS/2 ranges, available via
  ttf-parser).

#### 17.8.3.13 `w:panose1` (Panose-1 Typeface Classification) — NOT parsed
- **Spec**: §17.8.3.13. Child of `w:font`; `@w:val` is the 10-byte PANOSE-1
  classification number (ISO/IEC 14496-22 §5.2.7.17) as a 20-hex-digit string.
  Spec's **#3** substitution signal — encodes family kind, serif style, weight,
  proportion, contrast, etc., enabling a "nearest typeface" match.
- **Today**: `grep "panose" src/` hits only an unrelated comment
  (`src/fonts/mod.rs:217`). Not parsed or stored.
- **What exists vs missing**: nothing; no PANOSE field on `FontTableEntry`,
  no PANOSE-distance matcher. (PDF FontDescriptors carry a Panose entry too — the
  PDF-spec hits at `pdf-spec.pdf` chunk 2222–2223 — so the byte layout is the same
  one Word reads from the OS/2 table.)
- **Likely owner**: `FontTableEntry` field (`src/model/mod.rs`), parse in
  `src/docx/embedded_fonts.rs`, nearest-PANOSE matcher in `src/fonts/mod.rs`
  (compare the doc's 10 bytes against installed fonts' OS/2 PANOSE via ttf-parser).

#### 17.8.3.2 `w:charset` (Character Set) — NOT parsed
- **Spec**: §17.8.3.2. Child of `w:font`; `@w:val` = ST_UcharHexNumber (a single
  byte, e.g. `00`=ANSI/Latin, `CC`=Cyrillic, `EE`=Eastern European, `80`=Shift-JIS,
  `A1`=Greek). Spec's **#2** substitution signal; also the legacy hint for which
  code page a non-Unicode font expects.
- **Today**: `grep "charset" src/` → none. Not parsed or stored.
- **Likely owner**: `FontTableEntry` field (`src/model/mod.rs`), parse in
  `src/docx/embedded_fonts.rs`, factor into the substitution tie-break in
  `src/fonts/mod.rs`.

#### 17.8.3.14 `w:pitch` (Font Pitch) — PARTIALLY handled (parsed, never consumed)
- **Spec**: §17.8.3.14. Child of `w:font`; `@w:val` = ST_Pitch §17.18.74
  (`fixed` | `variable` | `default`). Spec's **#4** substitution signal — a
  fixed-pitch (monospace) font should be replaced by another fixed-pitch font.
- **Today**: parsed into `FontTableEntry.pitch_fixed`
  (`src/docx/embedded_fonts.rs:105`) but the field is `#[allow(dead_code)]`
  (`src/model/mod.rs:140`) — never read. So pitch is recorded but does not steer
  substitution.
- **What exists vs missing**: capture exists; consumption does not. A monospace
  font that is missing and has no `altName` could be replaced by a proportional
  family fallback.
- **Likely owner**: consume `pitch_fixed` in the family-fallback branch of
  `src/fonts/mod.rs` (prefer a monospace fallback like Courier/Consolas when
  `pitch_fixed` is set), and drop the `#[allow(dead_code)]`.

#### 17.8.3.12 `w:notTrueType` (Not a TrueType Outline Font) — NOT parsed
- **Spec**: §17.8.3.12. Boolean toggle child of `w:font` (§17.17.4) flagging that
  the font is not a TrueType-outline font (e.g. a bitmap/Type1 font). Spec's
  lowest (**#7**) substitution signal.
- **Today**: `grep "notTrueType" src/` → none. Not parsed.
- **Impact**: very low — affects only the final tie-break of font substitution.
- **Likely owner**: optional `FontTableEntry` bool (`src/model/mod.rs`), parse via
  `wml_bool` in `src/docx/embedded_fonts.rs`; only worth doing alongside the
  sig/panose/charset matcher work above.

### NO PRINT IMPACT (record-only — safe to leave unparsed)

- **`@w:subsetted` on `embedRegular`/`embedBold`/`embedItalic`/`embedBoldItalic`
  (17.8.3.3–.6)** — NOT read (`grep "subsetted" src/` → none). Declares that the
  embedded font is already a subset. Informational only: the embedded bytes are
  complete enough for ttf-parser to register and render regardless, so reading the
  flag changes nothing for output. Embedded fonts are already extracted,
  deobfuscated, and used (`src/docx/embedded_fonts.rs`, `src/fonts/mod.rs:173`).
- **17.8.3.8 `w:embedTrueTypeFonts` / 17.8.3.7 `w:embedSystemFonts` /
  17.8.3.15 `w:saveSubsetFonts`** — settings-part (`w:settings`) flags; NOT parsed
  (`grep` in `src/docx/settings.rs` → none). These are *save-time / round-trip
  authoring hints* (whether the app should embed fonts, embed system fonts too, and
  subset them when next saving). They have no bearing on reading + rendering: an
  embedded font present in `word/fonts/*` is used whenever it exists, irrespective
  of these flags. Only relevant if docxside ever *writes* DOCX. Leave unparsed.
- **17.8.1 obfuscation / 17.8.2 substitution** — prose/algorithm sections, no
  standalone elements; the algorithm (17.8.1) is implemented and the substitution
  process (17.8.2) exists (see WELL COVERED above and the priority note at the top
  of this 17.8 section).

---

## 17.9 Numbering — gaps

### HIGH priority (affect common Western numbered/heading documents)

#### 17.9.7 `w:lvlJc` (Numbering Symbol Justification) — NOT handled
- **Spec**: §17.9.7. Child of `w:lvl`; `@w:val` = ST_Jc (`start`/`left`,
  `center`, `end`/`right`). Justifies the number/bullet symbol *relative to the
  text margin* of the parent numbered paragraph. Default = left (LTR). `end`/
  `right` justification is how Word right-aligns the numbers in legal,
  outline, and TOC-style lists (e.g. ` 1.`, ` 10.`, `100.` with right-aligned
  dots), and `center` centers the symbol in the hanging-indent gutter.
- **Today**: `grep lvlJc src/` → 0 hits. Not parsed in `parse_level_def`
  (`src/docx/numbering.rs:45`); `LevelDef` has no justification field. The label
  is always placed left-justified at `indent_left - indent_hanging` in
  `src/pdf/layout.rs`. Right- and center-justified numbers render left-aligned →
  visibly misplaced in multi-digit numbered lists.
- **What exists vs missing**: left/default justification works; `center`/`end`
  modes are ignored. The hanging-indent geometry (`indent_left`,
  `indent_hanging`, `tab_stop`) needed to compute the symbol's right edge is
  already parsed — only the justification decision and the right-/center-anchor
  placement are missing.
- **Likely owner**: add `lvl_jc: String` (or enum) to `LevelDef`
  (`src/docx/numbering.rs`), parse `wml_attr(lvl, "lvlJc")`, carry it onto
  `ListLabelInfo`, and anchor the label's x-position at the right/center of the
  indent gutter in the label-rendering path in `src/pdf/layout.rs`.

#### 17.9.4 `w:isLgl` (Display All Levels Using Arabic Numerals) — NOT handled
- **Spec**: §17.9.4. Boolean toggle (§17.17.4) child of `w:lvl` — the "legal
  numbering" rule. When present, every level reference inside this level's
  `lvlText` (`%1`, `%2`, …) is rendered using **decimal** numerals regardless of
  that referenced level's own `numFmt`. So a level whose `lvlText` is `%1.%2.%3`
  where the sub-levels are `upperRoman`/`lowerLetter` still prints `1.2.3`.
- **Today**: `grep isLgl src/` → 0 hits. Not parsed; no field on `LevelDef`.
  The placeholder substitution loop (`src/docx/numbering.rs:396`) always formats
  each `%N` with *that referenced level's* `num_fmt` (`:409`), so a legal-numbered
  level mixing formats renders the wrong glyphs (e.g. `1.II.c` instead of
  `1.2.3`).
- **What exists vs missing**: per-level placeholder substitution works; the
  `isLgl` override that forces all referenced levels to decimal does not exist.
- **Likely owner**: add `is_lgl: bool` to `LevelDef` (`src/docx/numbering.rs`),
  parse via `wml_bool(lvl, "isLgl")`, and in the substitution loop force
  `lvl_fmt = "decimal"` for every `%N` when the *current* level's `is_lgl` is set.

#### 17.9.10 `w:lvlRestart` (Restart Numbering Level) — NOT handled
- **Spec**: §17.9.10. Child of `w:lvl`; `@w:val` (decimal) names the 1-based
  level whose advance triggers this level to restart at its `start` value.
  `@w:val="0"` means **never restart** (continuous numbering across the whole
  document). Omitted = default restart behavior (restart whenever any
  higher/more-significant level advances).
- **Today**: `grep lvlRestart src/` → 0 hits. Not parsed. The counter logic in
  `parse_list_info` (`src/docx/numbering.rs:347`) implements only the *default*
  rule — it clears deeper counters when returning to a higher level. It honors
  neither `lvlRestart=0` (a level that should keep counting across the document
  still gets reset when a higher level reappears) nor `lvlRestart=K` (restart
  gated on a *specific* ancestor level rather than any higher level).
- **What exists vs missing**: default cascade-restart approximated; explicit
  per-level restart control absent → continuous lists restart incorrectly and
  custom restart points are ignored.
- **Likely owner**: add `lvl_restart: Option<u32>` to `LevelDef`
  (`src/docx/numbering.rs`), parse `wml_attr(lvl, "lvlRestart")`, and gate the
  deeper-counter reset in `parse_list_info` on it (skip reset entirely when 0;
  only reset when the named level advances otherwise).

#### 17.9.23 `w:pStyle` (Paragraph Style's Associated Numbering Level) — NOT handled
- **Spec**: §17.9.23. Child of `w:lvl`; `@w:val` is a `styleId` linking this
  numbering level to a paragraph style. Combined with §17.3.1.19: when a
  paragraph-style's `pPr` carries `numPr`, the `ilvl` is **ignored** and the
  level is instead chosen by matching the paragraph's style against the
  abstractNum level whose `pStyle` names it. This is the standard mechanism for
  multilevel **heading numbering** (`Heading 1`→level 0, `Heading 2`→level 1, …).
- **Today**: `grep pStyle src/docx/numbering.rs` → 0 hits. The level is taken
  straight from the style's own `numPr` `ilvl` (`src/docx/styles.rs:679`,
  `num_ilvl`), defaulting to 0 when absent. A document that relies on `pStyle`
  level-mapping (style `numPr` without an explicit `ilvl`) collapses every heading
  to level 0 → wrong number format/indent per heading level.
- **What exists vs missing**: explicit-`ilvl` style numbering works; the
  `pStyle`-driven reverse lookup (style → level) does not. Most heading-numbered
  docs DO write an explicit `ilvl` so this is latent for many, but breaks the
  ones that omit it.
- **Likely owner**: parse `pStyle` per level into `NumberingInfo` (a
  `style_id → (abs_id, ilvl)` map) in `src/docx/numbering.rs`; when a paragraph's
  effective style matches, resolve the level via that map in `parse_list_info` /
  the caller in `src/docx/mod.rs:747`.

### MEDIUM priority (visible but rarer)

#### 17.9.9 `w:lvlPicBulletId` + 17.9.20 `w:numPicBullet` (Picture Bullets) — NOT handled
- **Spec**: §17.9.9 (`lvlPicBulletId @val`, child of `w:lvl`) references a
  §17.9.20 `numPicBullet @numPicBulletId` (child of `w:numbering`) which wraps a
  `w:drawing` holding the bullet image. Each char of `lvlText` is replaced by one
  instance of that image.
- **Today**: `grep "numPicBullet\|lvlPicBullet" src/` → 0 hits. Neither the
  `numPicBullet` registry nor the per-level reference is parsed. A picture-bulleted
  list falls through to the `bullet` `numFmt` path and renders whatever `lvlText`
  holds (often a placeholder glyph like `` in a symbol font) instead of the image.
- **What exists vs missing**: nothing for picture bullets; the inline-image
  drawing/extraction code already exists in `src/docx/images.rs` and could supply
  the bitmap.
- **Likely owner**: parse `numPicBullet` (id → relationship/image) at the
  `numbering` root in `src/docx/numbering.rs`; add `pic_bullet_id: Option<u32>` to
  `LevelDef`; render the referenced image in place of the bullet glyph in
  `src/pdf/layout.rs` (reuse the inline-image rendering path).

#### 17.9.17 `w:numFmt` — PARTIALLY handled (8 of ~60 ST_NumberFormat values)
- **Spec**: §17.9.17, values enumerated by ST_NumberFormat (§17.18.59 — ~60
  values). `format_number` (`src/docx/numbering.rs:261`) implements `decimal`,
  `decimalZero`, `lowerLetter`, `upperLetter`, `lowerRoman`, `upperRoman`, `none`,
  plus the `bullet` short-circuit — covering essentially all common Western lists.
  Unknown formats fall back to `decimal` (`:270`).
- **Missing (Western, occasionally seen)**: `ordinal` (1st, 2nd, 3rd),
  `cardinalText` (One, Two, Three), `ordinalText` (First, Second), `hex`,
  `chicago` (*, †, ‡, §), `decimalEnclosedCircle`/`decimalEnclosedParen`/
  `decimalEnclosedFullstop`. **Missing (CJK/other scripts, rare in current
  corpus)**: the ~45 East-Asian/Hebrew/Arabic-alpha/Thai/Russian formats
  (`japaneseCounting`, `chineseCounting`, `koreanDigital`, `ideographDigital`,
  `hebrew1/2`, `arabicAlpha`, `arabicAbjad`, `thaiLetters`, `russianLower/Upper`,
  `aiueo`, `iroha`, `ganada`, `chosung`, …) and the text-substitution formats
  `bahtText`/`dollarText`. Also `@w:format` (custom XSLT-style format string) on
  `w:numFmt` is not read.
- **Impact**: a missing-format level renders as plain decimal — wrong glyphs but
  not catastrophic (numbers still increment). Low for the CJK set (absent from
  corpus), medium for `ordinal`/`cardinalText` which appear in English docs.
- **Likely owner**: extend `format_number` (`src/docx/numbering.rs:261`) with the
  Western formats first; the CJK formats are a larger, separate effort.

#### 17.9.22 `w:pPr` (Level-Associated Paragraph Properties) — PARTIALLY handled
- **Spec**: §17.9.22. The `pPr` inside a `w:lvl` carries paragraph properties
  (indent, justification, spacing, tabs, …) that are merged into every paragraph
  using that level.
- **Today**: only `ind` (`start`/`left`/`hanging`) and the `num`-type `tab` are
  read (`src/docx/numbering.rs:64`). Other level `pPr` children — `jc`, `spacing`,
  `keepNext`, etc. — are ignored, so a level that (e.g.) sets its own line spacing
  or justification on the numbered paragraph won't apply it.
- **Likely owner**: broaden the `pPr` parse in `parse_level_def`
  (`src/docx/numbering.rs`) and merge the result into the paragraph's effective
  properties at the call site (`src/docx/mod.rs:747` / `src/docx/paragraph.rs`).

#### 17.9.24 `w:rPr` (Numbering Symbol Run Properties) — PARTIALLY handled
- **Spec**: §17.9.24. Run properties applied to the number/bullet symbol itself.
- **Today**: only `rFonts` (ascii/hAnsi), `sz`, `b`, and `color` are read
  (`src/docx/numbering.rs:75`). Missing: `i` (italic), `u` (underline), `strike`,
  `vertAlign`, `highlight`, etc. on the label. So an italic or underlined list
  number renders upright/plain.
- **Likely owner**: extend the `rPr` parse in `parse_level_def`
  (`src/docx/numbering.rs`) and honor the extra attrs when emitting the synthetic
  label run in `src/pdf/layout.rs`.

### NO PRINT IMPACT (record-only — safe to leave unparsed)

- **17.9.12 `w:multiLevelType`** (singleLevel/multilevel/hybridMultilevel) — the
  spec (§17.9.12) explicitly states this is a UI-only hint that "shall not be used
  to limit the behavior of the list." No rendering impact. Not parsed; correct.
- **17.9.13 `w:name`** (Abstract Numbering Definition Name) — UI label only. Not
  parsed; correct.
- **17.9.14 `w:nsid`** (Abstract Numbering Definition Identifier) — copy/paste
  merge-identity hex id; no print representation. Not parsed; correct. (Grep hits
  for "nsid" in `src/` are all `inside`/`consider` substrings, not this element.)
- **17.9.19 `w:numIdMacAtCleanup`** (Last Reviewed Abstract Numbering Definition)
  — round-trip/cleanup bookkeeping; no print impact. Not parsed; correct.
- **17.9.29 `w:tmpl`** (Numbering Template Code) — authoring/round-trip template
  GUID; no print impact. Not parsed; correct.
- **`w:lvl @w:tplc` / `@w:tentative`** — per-level template-code and
  "tentative" (auto-generated, not yet used) flags; authoring metadata, no print
  impact. Not parsed; correct.

---

## 17.10 Headers and Footers — gaps

All gaps below are facets of the single §17.10.1 rule on what to show when a
header/footer *type* is missing for a page: the appropriate type "should be
inherited from the previous section; if this is the first section in the
document, then a blank header/footer shall be created as needed (**another
header/footer type shall not be used in its place**)." docxside instead falls
back to the section's `default` (odd) header in every missing-type case, which
both (a) substitutes a different type, and (b) does not inherit the *same* type
from an earlier section. All are LOW/latent: most real documents that turn on
these features also define every part they need, so a fixture only diverges
when a type is deliberately omitted. No current fixture is known to fail on
this; the symptom would be an unexpected header/footer printing on a page that
Word leaves blank.

### LOW priority — §17.10.1 type-preservation / cross-section inheritance

#### 17.10.1 evenAndOddHeaders=true but the even part is omitted → wrong type shown
- **Spec**: §17.10.1. When `evenAndOddHeaders` is true and the even
  header/footer is omitted for a section, the even part must be inherited from
  the previous section's *even* part, or be blank if none — the odd/default part
  shall NOT be shown on even pages in its place.
- **Today**: `resolve_header_for_page` / `resolve_footer_for_page`
  (`src/pdf/header_footer.rs:984`/`1012`) test
  `doc.even_and_odd_headers && page_num % 2 == 0 && s.header_even.is_some()`; when
  `header_even`/`footer_even` is `None` they fall straight through to
  `s.header_default` (lines 995–998 / 1023–1026). So an even page with no even
  part defined prints the odd/default header — the substitution the spec forbids.
- **What exists vs missing**: the even-vs-odd *selection* works when all parts
  exist; the "blank when even omitted" behaviour does not.
- **Likely owner**: `src/pdf/header_footer.rs` — in the even-page branch, when
  `header_even`/`footer_even` is `None`, return blank (`None`) for the current
  section rather than falling to `header_default`, before the cross-section walk.

#### 17.10.1 cross-section inheritance only carries the `default` type
- **Spec**: §17.10.1. A missing header/footer type for a section is inherited
  from the previous section **of the same type** (first→first, even→even,
  odd→odd).
- **Today**: the backward walk in `resolve_header_for_page` /
  `resolve_footer_for_page` (`src/pdf/header_footer.rs:990`/`1018`) only ever
  consults `s.header_default` / `s.footer_default` for prior sections (the
  `else` arm, lines 1000–1002 / 1028–1030) — it never looks at a prior section's
  `header_first`/`header_even` (or footer equivalents). So first-page and
  even-page parts are not inherited across a section boundary; only the
  odd/default part is.
- **What exists vs missing**: default-type inheritance works; first/even-type
  inheritance is absent.
- **Likely owner**: `src/pdf/header_footer.rs` — in the `idx != section_idx` arm
  of both resolve functions, pick the prior section's same-type part
  (first/even/default) according to the page's role, not always `*_default`.

#### 17.10.6 titlePg set but no first-page part defined → default shown instead of blank
- **Spec**: §17.10.1 / §17.10.6. With `titlePg` on, the first page of the
  section uses the first-page header/footer; if none is defined it should be
  blank (or inherit a prior section's *first* part), not the default.
- **Today**: when `different_first_page` is true and `header_first`/`footer_first`
  is `None`, the `is_first && different_first_page` branch
  (`src/pdf/header_footer.rs:993`/`1021`) yields `None`, so the backward walk
  proceeds and paints a prior section's `header_default` on the title page.
- **What exists vs missing**: first-page selection works when a first part
  exists; the blank-when-absent behaviour does not.
- **Likely owner**: `src/pdf/header_footer.rs` — same fix as the two gaps above
  (treat a missing first-page part as blank for the current section, and inherit
  the *first* type, not default, across sections).

### NO PRINT IMPACT / fully covered (record-only)
- **17.10.2 footerReference / 17.10.5 headerReference** — fully parsed incl. all
  `@w:type` values and `r:id` resolution (`src/docx/sections.rs:123`). Done.
- **17.10.3 ftr / 17.10.4 hdr** — block-level content parsed like the body via
  `parse_header_footer_xml` (`src/docx/headers_footers.rs:30`). Done.
- **17.10.1 evenAndOddHeaders / 17.10.6 titlePg** — flags parsed and consumed for
  the common (all-parts-present) case; only the omitted-part edge cases above
  diverge.

---

## 17.11 Footnotes and Endnotes — gaps

All gaps are LATENT in the current corpus: docxside hardcodes the spec/Word
defaults (footnotes at page bottom in decimal, endnotes at document end in
decimal, starting at 1, never restarting), and every corpus fixture that uses
footnotes ships an EMPTY `<w:footnotePr/>`/`<w:endnotePr/>` so no fixture forces
a non-default value. They become real divergences the moment a document sets a
footnote/endnote property — and Word documents authored through the UI commonly
DO (e.g. Word's own endnote default written into settings is
`<w:numFmt w:val="lowerRoman"/>`, which docxside would currently ignore and
print as `1, 2, 3` instead of `i, ii, iii`).

### HIGH priority (footnote/endnote property bags entirely unparsed)

#### 17.11.12/.4 `w:footnotePr` / `w:endnotePr` (Document-Wide) — NOT parsed
- **Spec**: §17.11.12 (`footnotePr`) and §17.11.4 (`endnotePr`) are children of
  `w:settings` (the Settings part) carrying the document-wide note properties:
  child `w:pos` (§17.11.21/.22), `w:numFmt` (§17.11.18/.17), `w:numStart`
  (§17.11.20), `w:numRestart` (§17.11.19), plus the special-note ID lists
  `w:footnote`/`w:endnote` (§17.11.9/.3).
- **Today**: `src/docx/settings.rs` parses NO footnote/endnote element (grep for
  `footnote`/`endnote` in `settings.rs` → none). `Document` has no note-property
  field. So position, numbering format, start value, and restart rule are never
  read; rendering uses the hardcoded defaults in `src/pdf/footnotes.rs` /
  `src/pdf/mod.rs`.
- **What exists vs missing**: note *content* parsing and *rendering* exist; the
  property bag that controls *how* they number and *where* they sit does not.
- **Likely owner**: parse `footnotePr`/`endnotePr` in `src/docx/settings.rs` into
  new fields on `Document` (`src/model/mod.rs`), e.g. a small
  `NoteProperties { pos, num_fmt, num_start, num_restart }` struct for each;
  consume in the numbering pre-pass (`src/pdf/mod.rs:2651`) and placement logic
  (`src/pdf/footnotes.rs` / `src/pdf/mod.rs:3004`). This single parse unblocks the
  four numbering/position gaps below.

#### 17.11.11/.5 `w:footnotePr` / `w:endnotePr` (Section-Wide) — NOT parsed
- **Spec**: §17.11.11 (`footnotePr`) and §17.11.5 (`endnotePr`) are children of
  `w:sectPr`, overriding the document-wide note properties for that section. Same
  child elements as the doc-wide form. Notably, `pos="sectEnd"` endnotes (a very
  common setup — "endnotes at end of each section") are configured here.
- **Today**: `src/docx/sections.rs` parses no footnote/endnote element (grep →
  none). `SectionProperties` has no note-property field.
- **Likely owner**: parse in `src/docx/sections.rs` into `SectionProperties`
  (`src/model/mod.rs`); these override the doc-wide values from the gap above for
  notes referenced within the section. Pairs with — and is harder than — the
  doc-wide parse because `sectEnd`/per-section restart need section-aware
  placement, not just a global render pass.

#### 17.11.18/.17 `w:numFmt` (Footnote / Endnote Numbering Format) — NOT handled
- **Spec**: §17.11.18 (footnote) and §17.11.17 (endnote). `@w:val` =
  ST_NumberFormat (§17.18.59); selects the glyph set for the auto-numbered
  reference mark (decimal, lowerRoman, upperRoman, lowerLetter, upperLetter,
  `chicago` = *, †, ‡, §, symbol, etc.). Default per spec = decimal.
- **Today**: display numbers are emitted as plain decimal — the sequential `u32`
  is rendered with `display_num.to_string()` in `substitute_ref_marks`
  (`src/pdf/footnotes.rs:14-26`) and in the body-side substitution
  (`src/pdf/mod.rs:1080-1086`). No format field exists anywhere.
- **What exists vs missing**: sequential decimal numbering works; every other
  ST_NumberFormat is ignored. The existing `format_number` helper in
  `src/docx/numbering.rs:261` already implements decimal/roman/letter and could be
  reused to format the mark.
- **Likely owner**: after parsing `numFmt` (doc/section gaps above), format the
  display number through `numbering::format_number` at the two substitution sites
  (`src/pdf/footnotes.rs`, `src/pdf/mod.rs:1080`); the `chicago`/symbol set would
  need a small extra mapping.

#### 17.11.21/.22 `w:pos` (Footnote / Endnote Placement) — NOT handled
- **Spec**: §17.11.21 footnote `pos`, `@w:val` = ST_FtnPos (§17.18.34) =
  `pageBottom` (default) | `beneathText` | `sectEnd` | `docEnd`. §17.11.22 endnote
  `pos`, `@w:val` = ST_EdnPos (§17.18.22) = `sectEnd` | `docEnd` (default).
- **Today**: footnotes are always painted at the page bottom margin
  (`render_page_footnotes`, `src/pdf/footnotes.rs:96`, called with `eff_bottom` at
  `src/pdf/mod.rs:3015`); endnotes are always painted once on the final page
  (`src/pdf/mod.rs:3026`, the `endnote_ids` comment at `:464` states this
  assumption). `beneathText`/`sectEnd`/`docEnd` for footnotes and `sectEnd` for
  endnotes are not implemented.
- **What exists vs missing**: only the two default positions (footnote
  pageBottom, endnote docEnd) work. `sectEnd` (endnotes per section) and
  `beneathText` (footnotes directly under their referencing text) need new
  placement logic.
- **Likely owner**: consume the parsed `pos` (doc/section gaps) in
  `src/pdf/mod.rs:3004-3047`; `sectEnd` needs the note flush to key off section
  boundaries rather than the final page, `beneathText` needs the footnote block
  to be emitted right after the line that references it instead of at the margin.

### MEDIUM priority (numbering control)

#### 17.11.20 `w:numStart` (Numbering Starting Value) — NOT handled
- **Spec**: §17.11.20. `@w:val` (ST_DecimalNumber) = the first auto-number for
  notes (and the value used after each restart). Default = 1.
- **Today**: the pre-pass hardcodes `next_fn_num = 1` / `next_en_num = 1`
  (`src/pdf/mod.rs:2656-2657`). A document specifying a different start renders
  off-by-(start−1).
- **Likely owner**: seed the counters in `src/pdf/mod.rs:2651` from the parsed
  `numStart` instead of the literal `1`.

#### 17.11.19 `w:numRestart` (Numbering Restart Location) — NOT handled
- **Spec**: §17.11.19. `@w:val` = ST_RestartNumber (§17.18.74) = `continuous`
  (default) | `eachSect` | `eachPage`. When restarting, numbering returns to the
  `numStart` value.
- **Today**: numbering is always continuous — the encounter-order pre-pass
  (`src/pdf/mod.rs:2651-2686`) never resets. `eachSect`/`eachPage` not honored.
- **What exists vs missing**: `continuous` works; per-section/per-page restart
  absent. `eachPage` is awkward because the display-order pre-pass runs before
  pagination is known, so this likely has to move into (or coordinate with) the
  page-layout loop.
- **Likely owner**: `src/pdf/mod.rs` numbering pre-pass + the page/section layout
  loop; reset the counter at section/page boundaries per the parsed value.

### LOW priority (separators, suppression, custom marks)

#### 17.11.23 `w:separator` / 17.11.1 `w:continuationSeparator` — document-defined separators IGNORED
- **Spec**: §17.11.23 `separator` is the short horizontal rule between body text
  and the page's footnotes; §17.11.1 `continuationSeparator` is the full-width
  rule shown when a note continues from a previous page. Both are authored as the
  content of special notes (`<w:footnote w:type="separator">` /
  `w:type="continuationSeparator"`).
- **Today**: both parsers skip any note with a `w:type` attribute
  (`src/docx/headers_footers.rs:128-130` and `:240-242`), so the document's
  separator content is discarded. `render_page_footnotes` instead draws a
  HARDCODED separator: 0.5pt black, `min(144pt, text_width)` wide
  (`src/pdf/footnotes.rs:117-128`). The `continuationSeparator` is never drawn at
  all, but that is moot because docxside does not split a single footnote across
  pages.
- **Impact**: low — the hardcoded line matches Word's default separator closely
  (confirmed: `uk_commercial_lease_template` ships exactly
  `<w:footnote w:type="separator"><w:separator/></w:footnote>`, i.e. the default).
  A document with a customized separator (different width/style, or text) would
  diverge.
- **Likely owner**: only if customization matters — parse the `separator`-typed
  note's content in `src/docx/headers_footers.rs` and render it instead of the
  hardcoded rule in `src/pdf/footnotes.rs`.

#### 17.11.16 `w:noEndnote` (Suppress Endnotes In Document) — NOT handled
- **Spec**: §17.11.16. Settings-part boolean toggle; when set, all endnotes in
  the document shall not be displayed or printed.
- **Today**: not parsed (`src/docx/settings.rs` → none). Endnotes are always
  rendered (`src/pdf/mod.rs:3026`). A document with `w:noEndnote` would wrongly
  print endnotes Word omits.
- **Likely owner**: parse via `wml_bool` in `src/docx/settings.rs` into a
  `Document` flag; gate the endnote render block (`src/pdf/mod.rs:3026`) on it.

#### `customMarkFollows` attribute on `w:footnoteReference`/`w:endnoteReference` (17.11.14/.7) — NOT read
- **Spec**: §17.11.14/.7. Boolean attribute; when true the note uses a *custom*
  reference mark (a literal symbol authored in the adjacent run) rather than an
  auto-number, and the note must NOT increment the auto-numbering sequence.
- **Today**: the attribute is not read (`src/docx/runs.rs:1161-1198` parses only
  `@w:id`); every referenced note is assigned a sequential auto-number in
  `src/pdf/mod.rs:2651`. A custom-marked note both prints a spurious auto-number
  and shifts every following note's number by one.
- **Likely owner**: read `customMarkFollows` in `src/docx/runs.rs`, store it on
  the `Run`, and skip such IDs when assigning display numbers in the pre-pass
  (`src/pdf/mod.rs:2651`).

### NO PRINT IMPACT / by-design (record-only)

- **17.11.9/.3 `w:footnote`/`w:endnote` (Special Note ID Lists)** — these lists
  inside the settings `footnotePr`/`endnotePr` merely enumerate which note IDs are
  separator/continuation/special. docxside identifies special notes directly by
  the `w:type` attribute on each note in `word/footnotes.xml`/`endnotes.xml`
  (`src/docx/headers_footers.rs:128,240`), which is equivalent, so the ID list is
  not needed. By design; no action.
- **17.11.8 `w:endnotes` / 17.11.15 `w:footnotes` (Part Roots)** — both part roots
  are read and iterated (`src/docx/headers_footers.rs:112,222`). Done.
- **17.11.2/.10 `w:endnote`/`w:footnote` (Content)** — content parsed for normal
  (non-typed) notes. Done. (`w:footnote @w:type` on a normal note is absent by
  definition; the `@w:id` is parsed; the rarely-used `w:type="normal"` explicit
  value would be skipped by the current type filter — a latent edge case, but no
  corpus fixture writes an explicit `type="normal"`.)
- **17.11.6/.13 `w:endnoteRef`/`w:footnoteRef` (Reference Marks)** — parsed and
  substituted with the display number. Done.
- **17.11.7/.14 `w:endnoteReference`/`w:footnoteReference`** — parsed (`@w:id`),
  rendered as superscript marks. Done except the `customMarkFollows` attribute
  (LOW gap above).

---

## 17.12 Glossary Document — gaps

The whole chapter is a supplementary document story (the building-block /
AutoText / Quick-Part store in `word/glossary/document.xml`). Its content is
*not* part of the main document flow — Word inserts a `docPart` into the body
only on explicit user action, and Word's own PDF export does not paint glossary
content directly. docxside never opens `word/glossary/*` (verified: `grep -rni
"glossary\|docPart" src/` → 0), which is the correct behavior. **There is no
rendering gap to implement for 17.12.** All 16 elements (`behavior`, `behaviors`,
`category`, `description`, `docPart`, `docPartBody`, `docPartPr`, `docParts`,
`gallery`, `glossaryDocument`, `guid`, `name`×2, `style`, `type`, `types`) are
pure building-block metadata + content that only matters inside an editor, not in
a static PDF. The records below exist so a future agent does NOT mistake "glossary
unparsed" for a bug.

### NO PRINT IMPACT / by-design (record-only)

- **17.12.1–.16 entire glossary document** — correctly unparsed. The 16 elements
  define named building blocks (`docPart`), their insertion behaviors
  (`behavior`/`behaviors` — ST_DocPartBehavior §17.18.15), gallery/category
  classification (`gallery` — ST_DocPartGallery §17.18.16, `category`, `name`),
  entry types (`type`/`types` — ST_DocPartType §17.18.17, e.g. `bbPlcHdr`,
  `autoExp`, `formFld`), descriptive metadata (`description`, `guid`, `style`),
  and the entry content (`docPartBody`, analogous to `w:body`). None of this is
  emitted into the main story unless an editor inserts it, so a static
  render must skip it. Owner if a build-time "expand all AutoText" feature is
  ever wanted (it is not, for fidelity-to-Word's-export): a new glossary-part
  reader, not in any existing file. Do not add unless a concrete need appears.

### LATENT — one cross-reference to 17.5 (SDT placeholder resolution)

- **`bbPlcHdr` glossary entries as content-control placeholder text** — the ONLY
  way glossary content can legitimately surface in output. A `docPart` whose
  `types` is `bbPlcHdr` (§17.12.15, ST_DocPartType) holds the prompt text shown
  inside an *unfilled* structured document tag (content control), referenced from
  the body SDT's `sdtPr/docPartObj`/`docPartGallery` (the §17.5.2 property bag).
  - **Today**: docxside renders the body SDT's `sdtContent` verbatim
    (`src/docx/runs.rs` / `src/docx/mod.rs` SDT descent — see 17.5 entry) and does
    NOT resolve placeholders out of the glossary. This is fine in practice because
    **Word writes the resolved placeholder text directly into the body
    `sdtContent`** — so the prompt prints from the body copy, the glossary lookup
    is never needed. Confirmed against all 3 corpus fixtures with a glossary part:
    each glossary entry is a `bbPlcHdr` like `[Type the company name]` /
    `[Type the document title]`, and the same text lives in the body's
    `sdtContent` (so it renders without touching the glossary).
  - **When it would matter**: a (rare/malformed) document with an empty body
    `sdtContent` but a non-empty glossary `bbPlcHdr` entry — then Word would show
    the glossary prompt and docxside would show nothing. No corpus fixture
    exercises this. This is the SAME open item already recorded under
    17.5.2.39 `w:showingPlcHdr` (LOW, record-only) — see the 17.5 gaps section.
    Owner if ever needed: SDT descent in `src/docx/runs.rs` / `src/docx/mod.rs`,
    plus a new glossary-part reader keyed by `docPart` name/guid.

---

## 17.13 Annotations — gaps

The big picture: docxside renders the **Final / "accept-all" / No-Markup** view of
revisions (insertions shown, deletions hidden) and a faithful right-margin
**comment pane**. That is the correct default and the most common case, so every
gap below is LATENT in the current corpus (0 fixtures exercise moves, cell
revisions, permissions, or markup-display mode; 1 fixture each uses a plain
`<w:ins>`/`<w:del>`). Two gaps cause silent CONTENT LOSS today (moved text,
deleted-paragraph-mark merge); the rest are fidelity refinements.

### MEDIUM priority (content loss / wrong layout in common revision docs)

#### 17.13.5.21 `w:moveFrom` + 17.13.5.25 `w:moveTo` (Move Source / Destination) — NOT handled → moved text DROPPED
- **Spec**: §17.13.5.21 `moveFrom` wraps the run content at a move's *source*
  location; §17.13.5.25 `moveTo` wraps the *same* content at its *destination*.
  They are run-level revision wrappers exactly analogous to `del`/`ins`
  (CT_RunTrackChange). In the Final/accept view, `moveTo` content (the
  destination = the text's final position) is SHOWN and `moveFrom` content (the
  vacated source) is DROPPED.
- **Today**: `collect_run_nodes` (`src/docx/runs.rs:615-642`) has branches for
  `ins` (descend → shown) and `del` (skip → dropped) but NONE for `moveFrom` /
  `moveTo`. Any node not matched by a branch is silently not emitted, so BOTH the
  source and the destination are dropped. Result: moved text disappears entirely
  — including the `moveTo` copy that must appear in the final document. (The only
  `moveTo` hit in `src/` is `src/docx/textbox.rs:226`, an unrelated DrawingML
  animation build step — not this revision element.)
- **What exists vs missing**: ins/del handled; the move twins are not. The fix
  mirrors the existing ins/del branches: add `name == "moveTo"` →
  `collect_run_nodes(child, …)` (descend, shown) and `name == "moveFrom"` → skip
  (dropped).
- **Corpus**: 0 fixtures (`analyze-fixtures --grep "w:moveTo"` / `"w:moveFrom"`
  → 0). Latent. Note also the empty milestone markers
  `moveFromRangeStart`/`moveFromRangeEnd` (§17.13.5.23/.24) and
  `moveToRangeStart`/`moveToRangeEnd` (§17.13.5.26/.28) carry NO content (they
  only name/scope the move) and need no handling.
- **Likely owner**: `collect_run_nodes` in `src/docx/runs.rs` (add the two
  branches next to the `ins`/`del` arms at lines 615-618).

#### 17.13.5.15 `w:del` + 17.13.5.20 `w:ins` on a paragraph mark (in `pPr/rPr`) — NOT handled → paragraphs not merged
- **Spec**: §17.13.5.15 a `del` inside `w:pPr/w:rPr` marks the paragraph-ending
  mark itself as deleted; in the Final view the paragraph's content merges into
  the FOLLOWING paragraph (the pilcrow is gone). §17.13.5.20 `ins` on a paragraph
  mark is the inverse (a tracked paragraph split). These are distinct from the
  run-level `del`/`ins` (17.13.5.14/.18) that wrap run content.
- **Today**: `del`/`ins` are detected only at run level (`src/docx/runs.rs:617`).
  The paragraph-properties parser (`src/docx/paragraph.rs`) does not inspect
  `pPr/rPr/del`, so a tracked paragraph-mark deletion still renders as TWO
  separate paragraphs (an extra paragraph break vs. Word's single merged
  paragraph). Minor layout divergence (one spurious break) but visible.
- **What exists vs missing**: run-content ins/del handled; paragraph-mark
  ins/del entirely unhandled → no paragraph merge/split in final view.
- **Corpus**: 0 fixtures known to use it. Latent.
- **Likely owner**: detect `del` inside `pPr/rPr` in `src/docx/paragraph.rs`
  (set a flag on `Paragraph`), then in the block-assembly loop (`src/docx/mod.rs`
  around `:740`) or line layout (`src/pdf/layout.rs`) suppress the paragraph break
  / concatenate the flagged paragraph with the next.

#### 17.13.5.1 `w:cellDel` / 17.13.5.2 `w:cellIns` / 17.13.5.3 `w:cellMerge` (Tracked Table-Cell Revisions) — NOT handled
- **Spec**: §17.13.5.1 `cellDel` (cell deleted), §17.13.5.2 `cellIns` (cell
  inserted), §17.13.5.3 `cellMerge` (`@vMerge`/`@vMergeOrig`, ST_AnnotationVMerge
  §17.18.1 — cell vertically merged/split) — all children of `tcPr`/`trPr`. In
  the Final view a `cellDel` cell is removed, a `cellIns` cell kept, and
  `cellMerge` applies the *revised* vertical-merge state.
- **Today**: `grep "cellDel\|cellIns\|cellMerge" src/` → 0. Not parsed in
  `src/docx/tables.rs`. A tracked cell deletion still renders the cell; a tracked
  merge is not applied → wrong table geometry for revision-marked tables.
- **Corpus**: 0 fixtures. Latent.
- **Likely owner**: `src/docx/tables.rs` — read these in `tcPr`/`trPr`, drop
  `cellDel` cells and apply `cellMerge`'s `@vMerge` in the final view (the
  existing `vMerge` machinery can be reused).

### LOW priority (deliberate default / niche)

#### 17.13.5 revision MARKUP-display mode (strikethrough/underline/change-bars/balloons) — NOT implemented (Final view hardcoded)
- **Spec**: §17.13.5. A WordprocessingML document carries revisions that Word can
  display either as the Final document (accept-all) OR AS MARKUP (deletions with
  strikethrough, insertions underlined in the author's colour, vertical change
  bars in the margin, moves/format-changes in balloons). Word's PDF export honors
  the document's *display-for-review* state; the visibility is gated by
  `w:revisionView` (§17.15.1.69, Settings part: `@markup`/`@comments`/`@insDel`/
  `@formatting`/`@inkAnnotations`).
- **Today**: docxside ALWAYS renders the Final view — `ins` shown, `del` dropped
  (`src/docx/runs.rs:615-618`) — and never renders revision marks. `w:revisionView`
  is not parsed (`grep revisionView src/` → 0; `src/docx/settings.rs`). So an
  "All Markup"-configured document that Word would export WITH visible tracked
  changes is exported clean by docxside.
- **What exists vs missing**: Final view works and is the sensible default;
  markup view (and the setting that selects it) is absent.
- **Likely owner**: only if markup-view fidelity is wanted — parse `revisionView`
  / display-for-review in `src/docx/settings.rs`, retain (rather than drop) `del`
  positions, and add a strike+underline+colour+change-bar render path in
  `src/pdf/layout.rs`. Large; low value for fidelity-to-Word's-default-export.

#### 17.13.4 Comment-pane fidelity refinements (pane EXISTS; these are polish)
- **Comment text flattened to plain string.** `collect_text`
  (`src/docx/comments.rs:56-79`) gathers only `t`/`tab`/`br`/`p` into a plain
  `String`, discarding run formatting (bold/italic/colour) and inter-paragraph
  styling inside the comment. Word's pane preserves comment formatting; docxside
  renders it plain. Owner: `src/docx/comments.rs` (model richer comment content)
  + `src/pdf/comments.rs` (render it).
- **`w:comment @w:date` (17.13.4.2) not parsed** — only `@author`/`@initials`/
  text are read. Word's pane label is "Commented [initials N]:" and does not show
  the date, so likely no visual impact — record-only.
- **17.13.4.1 `annotationRef`** — not parsed as an element; the pane label is
  synthesized from `initials + display_index` (`src/pdf/comments.rs` `format_label`),
  which matches Word's mark. No action.

### NO PRINT IMPACT (record-only — safe to leave unparsed)

- **17.13.7 `w:permStart` / `w:permEnd` (Range Permissions)** — NOT parsed
  (`grep "permStart\|permEnd" src/` → 0). Invisible editing-region markers gated
  by `documentProtection` (§17.15.1.29); no visual representation in a static PDF.
  `@edGrp` (ST_EdGrp §17.18.21) / `@ed` select which user group may edit. 0
  corpus fixtures. Correctly ignorable.
- **17.13.5.29–.38 property-change revisions** — `pPrChange` (.29),
  `rPrChange` on the paragraph mark (.30) and on a run (.31), `sectPrChange` (.32),
  `tblPrChange`/`tblPrExChange`, `tblGridChange`, `trPrChange` (.37),
  `tcPrChange` (.36), `numberingChange`. Each stores the FORMER property set; the
  Final view applies the *current* properties and ignores the `*Change` child.
  The property parsers never read these nested children → correctly ignored
  (`sectPrChange` was already noted "harmlessly ignored" under 17.6). No action.
- **17.13.5.4–.13 custom-XML tracked-change range markers** —
  `customXmlInsRangeStart`/`End`, `customXmlDelRangeStart`/`End`,
  `customXmlMoveFromRangeStart`/`End`, `customXmlMoveToRangeStart`/`End`: empty
  milestone markers scoping a tracked change around custom XML; no content, no
  print impact (already recorded under 17.5). Leave unparsed.
- **17.13.6.1 `w:bookmarkEnd`** — intentionally not read; only the start position
  (name → location, `src/docx/paragraph.rs:370`) is needed for PAGEREF/hyperlink
  targets, and bookmarks render no glyphs. Complete by design.
- **17.13.5 `delText` / `delInstrText`** — the deleted-text payload inside a
  `del`; never reached because `del` is skipped wholesale in the final view.
  Correct.

---

## 17.14 Mail Merge — gaps

**No implementable rendering gap.** The entire chapter (17.14.1–17.14.34, all
children of `w:mailMerge` §17.14.20 in `word/settings.xml`, with nested groups
under `w:odso` / `w:fieldMapData` / `w:recipients` / `w:recipientData`) is
mail-merge *connection / configuration metadata*: which external data source to
connect to (`dataSource` §17.14.9 r:id, `connectString` §17.14.8, `src`
§17.14.30, `udl` §17.14.34, `query` §17.14.26, `table` §17.14.31, `type`
§17.14.32, `dataType` §17.14.10, `colDelim` §17.14.5, `fHdr` §17.14.14,
`headerSource` §17.14.16, `linkToQuery` §17.14.18), how its columns map to
predefined MERGEFIELD names (`fieldMapData` §17.14.15, `column`/`name`/
`mappedName`/`type`/`lid` §17.14.6/.24/.23/.33/.17, `addressFieldName` §17.14.3,
`dynamicAddress` §17.14.13), which records to include (`recipients`/
`recipientData`/`active`/`activeRecord` §17.14.29/.27/.1/.2), and what to do
with the merged output (`mainDocumentType` §17.14.22, `destination` §17.14.11,
`mailSubject` §17.14.21, `mailAsAttachment` §17.14.19, `checkErrors` §17.14.4,
`doNotSuppressBlankLines` §17.14.12). None of these has a print representation —
the merge is an authoring/runtime operation performed by the host app; once
done, populated values are stored as ordinary cached field results in the body.
A static renderer correctly shows the cached result and ignores the settings.

- **Verified absent**: `grep -rniE
  "mailmerge|odso|fieldmapdata|mergefield|mappedname|recipientdata|maindocumenttype"
  src/` → 0 hits. `parse_settings` (`src/docx/settings.rs:30`) does not read
  `w:mailMerge`. This is correct; do NOT add parsing — it would have no effect
  on output.
- **Corpus**: 0 fixtures contain `MERGEFIELD` in `document.xml`
  (`analyze-fixtures --grep "MERGEFIELD"` → 0). CAVEAT: `analyze-fixtures --grep`
  scans only `document.xml`+`styles.xml`, NOT `word/settings.xml`, so it cannot
  confirm whether any fixture carries a `w:mailMerge` block in its settings part
  (same tool limitation noted under 17.8/17.9/17.11). Impact is moot regardless,
  since the element has no print effect.

### Cross-reference (the only print-relevant facet — owned by Clause 17.16, NOT 17.14)

- **MERGEFIELD / ADDRESSBLOCK / GREETINGLINE / NEXT / MERGEREC / MERGESEQ /
  SKIPIF / NEXTIF field RESULTS** — these are the field *instruction codes* that
  consume mail-merge data (§17.16.5.35 MERGEFIELD, etc.), and they DO produce
  visible text. They are field machinery (Clause 17.16 Fields and Hyperlinks),
  not 17.14. Today `src/docx/runs.rs` parses `fldSimple` (`:623`), the
  `fldChar` begin/separate/end run (`:1001`), and `instrText` (`:1053`), and
  renders the **cached field result** for any field whose keyword is not in
  `is_dynamic_field` (`runs.rs:22`) — MERGEFIELD is not dynamic, so its cached
  result (e.g. `«FirstName»` or already-merged text) renders as-is, which is
  spec-correct for a static export. Whether the `\f`/`\b`/`\c` ADDRESSBLOCK
  composite-formatting switches and the un-cached `dynamicAddress`/`mappedName`
  reordering need explicit evaluation is a question for the **17.16** audit pass
  (owner: `src/docx/runs.rs` field handling); it is deliberately out of scope
  here.

---

## 17.15 Settings — gaps

All gaps below are in the **Document Settings part** (`word/settings.xml`),
owner `src/docx/settings.rs` (`parse_settings` → `DocumentSettings`, threaded
onto the `Document` model in `src/docx/mod.rs:830`+ and consumed in `src/pdf/`).
The Web Page Settings group (17.15.2) is entirely HTML-save metadata and is
correctly out of scope (record-only note at the bottom).

### HIGH priority (real layout impact; present in the corpus today)

#### 17.15.1.10 `w:autoHyphenation` (+ 17.15.1.53 `hyphenationZone`, 17.15.1.22 `consecutiveHyphenLimit`, 17.15.1.37 `doNotHyphenateCaps`) — hyphenation cluster: PARTIALLY parsed, NEVER applied
- **Spec**: §17.15.1.10 `autoHyphenation` (bool toggle §17.17.4) turns on
  automatic hyphenation for the whole document; §17.15.1.53 `hyphenationZone`
  (`@val` twips, default 360 = 0.25") = max trailing whitespace allowed before
  the next word is hyphenated; §17.15.1.22 `consecutiveHyphenLimit` (`@val`
  decimal, 0 = unlimited) caps consecutive hyphenated lines; §17.15.1.37
  `doNotHyphenateCaps` (bool) excludes ALL-CAPS words. Per-paragraph opt-out is
  `w:suppressAutoHyphens` (17.3.1.34). Honoring these changes where lines break,
  which reflows text and can shift page counts vs. Word's reference PDF.
- **Today**: `autoHyphenation` IS parsed (`settings.rs:60`) into
  `Document.auto_hyphenation` but is NEVER consumed in `src/pdf/` (no hyphenation
  engine exists — `grep auto_hyphenation src/pdf/` → 0). `hyphenationZone`,
  `consecutiveHyphenLimit`, `doNotHyphenateCaps` are NOT parsed at all
  (`grep -rl … src/` → 0). `suppress_auto_hyphens` is parsed onto `Paragraph`
  but also `#[allow(dead_code)]` (noted in the 17.3 gaps). So a document that
  asks for hyphenation gets none — lines break only at spaces, diverging from
  Word.
- **What exists vs missing**: the on/off flag is captured; the actual
  hyphenation algorithm (dictionary/pattern-based break insertion in the
  line-breaker), the zone threshold, the consecutive-line cap, and the caps
  exclusion are all missing.
- **Corpus**: `hyphenationZone` in ~25 of 53 scraped fixtures; `autoHyphenation`
  (true) in 3 (`german_mezzo_soprano_bio`, `slovak_misdemeanor_amendment`,
  `traditional_skills_job_form`); `doNotHyphenateCaps` in ~3. The 3
  auto-hyphenated fixtures are the ones whose body text actively reflows wrong
  today.
- **Likely owner**: parse the four elements in `src/docx/settings.rs` (add fields
  to `DocumentSettings`/`Document`); implement hyphenation in the line-breaker
  `src/pdf/layout.rs` (`build_lines_with_float`), gated on `auto_hyphenation`,
  honoring `hyphenation_zone`, `consecutive_hyphen_limit`, `doNotHyphenateCaps`,
  and per-paragraph `suppress_auto_hyphens`. Large effort (needs a hyphenation
  dictionary/pattern set, e.g. Liang/TeX patterns per language from
  `default_lang`).

#### 17.15.1.21 `w:compat` + 17.15.3.4 `w:compatSetting` (Compatibility Settings) — NOT parsed; gate layout-engine behavior
- **Spec**: §17.15.1.21 `compat` is a container of legacy boolean compatibility
  toggles (Word 6/95/97/2000/2002… layout quirks); §17.15.3.4 `compatSetting`
  (children `@w:name`/`@w:uri`/`@w:val`) carries the modern key/value
  compatibility flags, most importantly `compatibilityMode` (15 = Word 2013+
  layout rules, 14 = Word 2010, 12 = Word 2007). `compatibilityMode` selects
  which line-spacing / table / kerning rules the layout engine uses, so it is the
  single highest-leverage flag for matching Word's geometry.
- **Today**: not parsed anywhere (`grep -rl "compat" src/` → 0). docxside
  implicitly assumes one fixed set of layout rules.
- **What exists vs missing**: nothing read. The engine has no notion of
  compatibility mode, so it cannot switch between Word-2007 and Word-2013 line
  metrics, table-style font/justification rules, etc.
- **Corpus**: `w:compat` present in **all 53** fixtures; every one carries
  `<w:compatSetting w:name="compatibilityMode" w:val="15"/>` plus a handful of
  `w:val="1"` toggles (`overrideTableStyleFontSizeAndJustification`,
  `enableOpenTypeFeatures`, `doNotFlipMirrorIndents`,
  `differentiateMultirowTableHeaders`, `useWord2013TrackBottomHyphenation`).
- **Likely owner**: parse `compat`/`compatSetting` in `src/docx/settings.rs` into
  a small `CompatFlags` struct on `Document`; the highest-value first step is to
  read `compatibilityMode` and `overrideTableStyleFontSizeAndJustification` and
  thread them into the relevant layout decisions (`src/pdf/layout.rs`,
  `src/pdf/table*.rs`). Most legacy `compat` toggles are Word-95/97-era quirks
  unlikely to appear in the corpus and can be deferred. SCOPE NOTE: this is a
  research-heavy gap — each flag that proves to matter needs a Word-vs-generated
  diff to confirm the exact rule change before implementing; do NOT bulk-parse
  all ~70 toggles blindly.

#### 17.15.3.5 `w:doNotExpandShiftReturn` (Don't Justify Lines Ending in Soft Line Break) — NOT handled
- **Spec**: §17.15.3.5 (a `w:compat` child). Bool toggle. Default: a fully
  justified paragraph (`w:jc val="both"`) justifies every line EXCEPT the last;
  with this flag set, any line ending in a **soft line break** (`w:br` without
  type) is ALSO left unjustified (ragged). Affects horizontal glyph spacing on
  those lines.
- **Today**: not parsed (`grep -rl doNotExpandShiftReturn src/` → 0). The
  justification code in `src/pdf/layout.rs` justifies all non-final lines
  uniformly, with no soft-break exception.
- **Corpus**: present in 3 fixtures (`east_asia_conference_form`,
  `japanese_interlibrary_loan`, `korean_japanese_conference_form`).
- **Likely owner**: parse the flag in `src/docx/settings.rs`; in the justified
  line-spacing path of `src/pdf/layout.rs`, skip justification for lines that
  terminate at a soft `w:br` when the flag is set (requires the line-breaker to
  track whether a line ended on a soft break).

### MEDIUM priority (latent in corpus / smaller visual scope)

#### 17.15.1.57 `w:mirrorMargins` — PARSED but DEAD (never applied)
- **Spec**: §17.15.1.57. Bool toggle. When set, left/right page margins swap on
  facing (odd/even) pages so the inner (binding) margin is consistent — the page
  geometry differs between odd and even pages.
- **Today**: parsed into `DocumentSettings.mirror_margins` (`settings.rs:57`) but
  the field is `#[allow(dead_code)]` (`settings.rs:8`) and is NOT carried onto
  `Document` or consumed in `src/pdf/` — odd/even pages render with identical
  margins.
- **Corpus**: present in 1 fixture (`traditional_skills_job_form`).
- **Likely owner**: thread `mirror_margins` onto `Document` in `src/docx/mod.rs`,
  then in `src/pdf/mod.rs`/`header_footer.rs` swap `margin_left`/`margin_right`
  for even pages when set. Pairs with the duplex-layout work needed for 17.3.1.16
  `mirrorIndents` (already noted in 17.3 gaps) and `compatSetting`
  `doNotFlipMirrorIndents`.

#### 17.15.1.64 `w:printTwoOnOne` + 17.15.1.11/.12/.13 `w:bookFoldPrinting`/`bookFoldPrintingSheets`/`bookFoldRevPrinting` (Page Imposition) — NOT handled
- **Spec**: §17.15.1.64 `printTwoOnOne` halves each section page and prints two
  logical pages per physical sheet; §17.15.1.11–.13 book-fold printing imposes
  pages as folded signatures (landscape sheet split vertically, page order
  reordered for binding, `bookFoldPrintingSheets @val` = sheets per signature).
  Both change the physical page geometry/ordering of the output PDF.
- **Today**: not parsed (`grep -rl bookFold src/` → 0; `printTwoOnOne` → 0). Each
  logical page is emitted 1:1 to a PDF page at full section size.
- **Corpus**: `bookFoldPrinting` in 1 fixture (`traditional_skills_job_form`).
  NOTE: confirm against the reference PDF whether Word's PDF *export* actually
  applies book-fold imposition (it may export 1:1 and only impose at print time)
  before implementing — otherwise matching the reference means doing NOTHING here.
- **Likely owner**: page-emission loop in `src/pdf/mod.rs`. Significant geometry
  work; low value unless a reference PDF proves Word imposes on export.

#### 17.15.1.60 `w:noPunctuationKerning` — NOT handled
- **Spec**: §17.15.1.60. Bool toggle. When kerning is enabled on a run
  (`w:kern`, 17.3.2.19), Latin text always kerns; this flag additionally
  suppresses kerning of punctuation glyphs. Subtle horizontal spacing effect.
- **Today**: not parsed (`grep -rl noPunctuationKerning src/` → 0). `w:kern` is
  consumed in layout, but there is no punctuation-specific kerning carve-out.
- **Corpus**: present in ~6 fixtures.
- **Likely owner**: parse in `src/docx/settings.rs`; in the kerning path of
  `src/pdf/layout.rs`, skip kern pairs where one glyph is punctuation when set.
  Low visual magnitude.

#### 17.15.1.24 `w:defaultTableStyle` — NOT handled
- **Spec**: §17.15.1.24. `@val` = `styleId` of the table style auto-applied to
  newly inserted tables that don't reference their own style. Affects rendering
  only for tables that genuinely carry no `tblStyle`.
- **Today**: not parsed (`grep -rl defaultTableStyle src/` → 0). Tables without
  an explicit `tblStyle` fall back to docxside's built-in default rather than the
  document-declared default table style.
- **Likely owner**: parse in `src/docx/settings.rs`; apply as the fallback
  `tblStyle` in table-style resolution (`src/docx/tables.rs` / `src/pdf/table.rs`).

### LOW priority (East-Asian / niche layout; latent)

- **17.15.3.1 `w:adjustLineHeightInTable`** (compat child) — NOT parsed. When set,
  a `docGrid` *line* pitch (17.6.5) is also added to lines inside table cells
  (normally it is not). Only matters when a section uses a line-grid AND has
  tables. Owner: `src/docx/settings.rs` + line-height calc in
  `src/pdf/table_layout.rs`/`layout.rs`.
- **17.15.1.18 `w:characterSpacingControl`** (ST_CharacterSpacing §17.18.7:
  `doNotCompress` default / `compressPunctuation` / `compressPunctuationAndJapaneseKana`)
  — NOT parsed. Controls full-width (CJK) whitespace compression at line layout.
  Present in ~51 fixtures but the value is the `doNotCompress` default in Latin
  docs → latent; matters only for CJK justification. Owner: `src/docx/settings.rs`
  + CJK spacing in `src/pdf/layout.rs`.
- **17.15.1.62 `w:printFractionalCharacterWidth`** — NOT parsed. Legacy flag
  choosing fractional vs. integer (rounded-to-pixel) advance widths; docxside
  already uses exact font metrics (effectively fractional), so honoring it is
  likely a no-op for PDF output. Owner: `src/docx/settings.rs` (record-only
  unless a reference diff shows Word rounding).
- **Drawing grid family — 17.15.1.27 `displayHorizontalDrawingGridEvery`,
  .42 `doNotUseMarginsForDrawingGridOrigin`, .44 `drawingGridHorizontalOrigin`,
  .45 `drawingGridHorizontalSpacing`, .46 `drawingGridVerticalOrigin`,
  .47 `drawingGridVerticalSpacing`** — NOT parsed. The drawing grid is an
  authoring aid for snapping floating objects; present in ~10 fixtures but has no
  bearing on the final rendered position of already-placed objects (the object's
  resolved offset is stored on the drawing). Effectively record-only for a static
  render. Owner: `src/docx/settings.rs` only if floating-object snapping is ever
  emulated.
- **17.15.1.88 `w:themeFontLang @val` → `Document.default_lang`** — PARSED but
  never consumed (`grep default_lang src/pdf/` → 0). Only relevant once
  language-aware behavior exists (hyphenation dictionary selection, lang-tagged
  font fallback). Pair with the hyphenation gap above.

### NO PRINT IMPACT (record-only — safe to leave unparsed)

- **17.15.1.29 `w:documentProtection` / 17.15.1.93 `w:writeProtection`** —
  editing-protection (read-only / forms-only / comments-only) + write password;
  no bearing on a static PDF. In corpus: `documentProtection` in 2 fixtures
  (`centrifugal_water_chillers`, `education_consultant_posting`), both
  `w:enforcement="0"` (not even enforced). None.
- **17.15.1.23 `w:decimalSymbol` / 17.15.1.55 `w:listSeparator`** — radix point
  and argument separator used only when *evaluating field code formulas* (= and
  IF fields). docxside renders cached field results (per the 17.14/17.16 notes),
  so these never affect output today. Present in ~51 fixtures. Owner (only if
  live field evaluation is ever added): `src/docx/runs.rs` field handling.
- **Spell/grammar & UI state** — 17.15.1.1 `activeWritingStyle`, .51
  `hideGrammaticalErrors`, .52 `hideSpellingErrors`, .65 `proofState`,
  `noProofState`, `clickAndTypeStyle` (.19), `stylePaneFormatFilter` (.85),
  `stylePaneSortMethod`, `defaultTabStop`-adjacent UI knobs, `zoom` (.94),
  `view`, `showEnvelope` (.79), `doNotTrackMoves`/`doNotTrackFormatting`: all
  authoring/UI state with no print representation. Leave unparsed.
- **Save / round-trip options** — `saveSubsetFonts`, `embedTrueTypeFonts`,
  `embedSystemFonts` (covered in 17.8), `doNotAutoCompressPictures` (.33),
  `saveXmlDataOnly`, `saveThroughXslt`, `savePreviewPicture`, `forceUpgrade`
  (.48), `doNotEmbedSmartTags`, `doNotIncludeSubdocsInStats`: only relevant when
  *writing* a DOCX. `savePreviewPicture` in 1 fixture. Leave unparsed.
- **`w:docVars` / `w:docVar` (17.15.1.31/.32)** — name/value storage for the
  DOCVARIABLE field; pure data, only visible if a DOCVARIABLE field references it
  (Clause 17.16). Cached field result already renders. Leave unparsed.
- **17.15.2 Web Page Settings (entire group, ~45 elements:
  `allowPNG`, `frameset`, `doNotUseLongFileNames`, `pixelsPerInch`,
  `targetScreenSz`, `optimizeForBrowser`, …)** — govern HTML export only; OOXML
  itself says saving as HTML is out of scope. Zero PDF relevance. Leave unparsed.
- **Tracked-change *settings*** — `w:trackChanges` (17.15.1.90),
  `rsids`/`rsid` (revision-save IDs): `trackChanges` is the editor toggle, not
  content; rsids are change-tracking bookkeeping. docxside renders the final
  (accept-all) view regardless (see 17.13). Leave unparsed.

---

## 17.16 Fields and Hyperlinks — gaps

### HIGH priority (real content/interactivity loss, present in corpus)

#### 17.16.17 `w:ffData` + form fields (FORMTEXT / FORMCHECKBOX / FORMDROPDOWN) — NOT handled → form-field glyphs/values missing
- **Spec**: §17.16.17 `ffData` (Form Field Properties) is a child of the `begin`
  `fldChar` of a legacy form field; its field code (in `instrText`) is FORMTEXT
  (§17.16.5.22), FORMCHECKBOX (§17.16.5.20), or FORMDROPDOWN (§17.16.5.21).
  Child property elements: `name`/`label` (§17.16.27/.24), `enabled`
  (§17.16.14), `calcOnExit`, `entryMacro`/`exitMacro` (§17.16.15/.16),
  `helpText`/`statusText` (§17.16.21/.31); for text boxes `textInput`
  (§17.16.32) → `type` (§17.16.34, ST_FFTextType), `default` (§17.16.10),
  `maxLength` (§17.16.26), `format` (§17.16.20); for checkboxes `checkBox`
  (§17.16.7) → `size`/`sizeAuto` (§17.16.29/.30), `default` (§17.16.12),
  `checked` (§17.16.8); for drop-downs `ddList` (§17.16.9) → `listEntry`
  (§17.16.25), `result` (§17.16.28), `default` (§17.16.11).
- **Today**: the string `ffData` / `checkBox` / `ddList` / `FORMCHECKBOX` etc.
  appears NOWHERE in `src/` (verified `grep -rniE "ffData|formcheckbox|formdropdown|ddList" src/` → 0).
  The `begin` `fldChar`'s children are not inspected, so `ffData` is silently
  dropped. Consequences for a static PDF:
  - **FORMCHECKBOX**: renders **nothing** — a checkbox has no result-region run,
    so the checked/unchecked box glyph (driven by `checkBox/checked` vs
    `default`, drawn at `checkBox/size` points) never appears.
  - **FORMTEXT**: the entered/cached value usually IS a result-region run, so it
    renders as text; but an *empty* FORMTEXT with only a `textInput/default`
    string shows nothing (default value lives in `ffData`, not as a result run).
  - **FORMDROPDOWN**: the selected entry (`result` index into `listEntry`) is
    normally cached as a result run and renders; an unselected one shows nothing.
- **Corpus**: **0** fixtures (`analyze-fixtures --grep "w:ffData"` → 0; same for
  FORMTEXT/FORMCHECKBOX/FORMDROPDOWN). Latent gap — no current test regresses,
  but legacy protected forms are common in the wild (the audit's
  `traditional_skills_job_form` is a form built from underscores + HYPERLINK
  fields, NOT `ffData`, so it does not exercise this).
- **Likely owner**: parse `ffData` from the `begin` `fldChar` in the run-builder
  in `src/docx/runs.rs` (where the `FieldFrame` is pushed, §1011); model a
  form-field variant on `Run` or a new `FieldCode` arm in `src/model/mod.rs`;
  render the checkbox box + selected/default text in `src/pdf/layout.rs`. The
  checkbox glyph can reuse the existing symbol/rect drawing primitives.

#### 17.16.5.X HYPERLINK field code — NOT recognized → clickable link annotation lost
- **Spec**: §17.16.5 HYPERLINK field. A complex field whose `instrText` is
  `HYPERLINK "target" [switches]` produces a clickable hyperlink whose display
  text is the result-region runs — the field-code form of the `w:hyperlink`
  element (§17.16.22). Switches: `\l anchor` (sub-location/bookmark), `\o tooltip`,
  `\t target-frame`, `\m`/`\n`.
- **Today**: in the `fldChar` `end` handler (`src/docx/runs.rs:1026–1039`) the
  keyword switch recognizes only PAGE / NUMPAGES / STYLEREF / PAGEREF; HYPERLINK
  falls to `fc = None`, so the result-region runs are emitted as ordinary text
  with **`hyperlink_url = None`** → **no PDF Link annotation** is generated (the
  external-URI/GoTo emission in `src/pdf/assembly.rs:75–99` only fires for runs
  carrying a `hyperlink_url`). Visual styling is usually preserved because the
  display runs carry `rStyle="Hyperlink"` (blue+underline from the char style),
  so the **pixel/Jaccard impact is ~nil** — the loss is the *clickable link*.
- **Corpus**: HYPERLINK field code in 1 fixture — `traditional_skills_job_form`
  (FAILING 15.5%): `<w:instrText> HYPERLINK "mailto:info@glasgowheritage.org.uk"</w:instrText>`
  (verified via `docx-inspect … word/document.xml`). Its low score is unrelated
  to this (styling renders); only the mailto link is non-clickable.
- **Likely owner**: in the `fldChar` `end` arm in `src/docx/runs.rs`, when
  keyword is HYPERLINK, parse the quoted target (+ `\l` anchor) from `f.instr`
  and set `hyperlink_url` on the emitted result runs (mirror the `w:hyperlink`
  element path at `runs.rs:594`), so the existing `LinkAnnotation` machinery
  picks it up. No new rendering code needed.

### MEDIUM priority (cached-result staleness — correct only if Word's cache matches docxside layout)

#### Field results are taken verbatim from Word's cached result region (no re-evaluation except 4 codes)
- **Spec**: §17.16.18/.19 — a field's *result* is the content between `separate`
  and `end` (complex) or the `fldSimple` children (simple); a conforming reader
  MAY recalculate it. Word stores the last-computed result; `@w:dirty`
  (§17.16.18) flags it stale.
- **Today**: only PAGE / NUMPAGES / STYLEREF / PAGEREF are re-evaluated
  (`is_dynamic_field`, `src/docx/runs.rs:22`); every other field code renders
  Word's cached result text verbatim. This is the right default for a static PDF
  and matches Word's own PDF export for most codes, BUT diverges when the cached
  result depends on layout that docxside computes differently:
  - **TOC (§17.16.5.71)**: the cached TOC body renders, and its per-entry page
    numbers come from nested PAGEREF sub-fields which ARE re-evaluated (they sit
    in the TOC result region, so `visible=true`) — so page numbers track
    docxside pagination. However the TOC entry *text/leader/indent layout* is the
    cached run content; if docxside's heading set or pagination differs, the TOC
    can list wrong/missing entries (it does not regenerate the entry list).
    Corpus: TOC in 3 fixtures (feminist_voice 66.5%, transition_to_work 40.1%,
    uk_commercial_lease 9.4%).
  - **`@w:dirty="true"` fields**: docxside ignores `dirty` and shows the cached
    result even when Word would recompute on open. No model field for it.
- **Likely owner**: `src/docx/runs.rs` (field handling) if live evaluation of
  more codes is ever added; for TOC regeneration, a new pass in `src/pdf/` that
  walks heading entries (the `HeadingEntry`/`heading_entries` infrastructure
  already exists for PDF outline/bookmarks). Large; low ROI vs. cached results.

#### 17.16.19 `w:fldSimple` dynamic codes not re-evaluated; `@w:instr` ignored
- **Spec**: §17.16.19. `fldSimple @w:instr` carries the field code; children are
  the cached result. A `<w:fldSimple w:instr=" PAGE ">` should re-evaluate like
  the complex PAGE field.
- **Today**: `collect_run_nodes` (`src/docx/runs.rs:623`) recurses into
  `fldSimple` children and **ignores `@w:instr`** — the cached child runs are
  shown as static text. So a simple-field PAGE/NUMPAGES/HYPERLINK shows a stale
  number / loses its link, unlike the complex-field path which handles those.
- **Corpus**: **0** fixtures use `w:fldSimple` (`analyze-fixtures --grep` → 0) —
  fully latent.
- **Likely owner**: `collect_run_nodes` / run-builder in `src/docx/runs.rs` —
  read `@w:instr`, and for dynamic/HYPERLINK keywords apply the same treatment as
  the complex-field `end` handler instead of blindly emitting cached children.

### LOW priority (no print impact / niche — record-only)

- **17.16.22 `w:hyperlink` secondary attributes** — `@w:tooltip`,
  `@w:tgtFrame`, `@w:docLocation`, `@w:history` are NOT read (only `r:id` +
  `@anchor` are). `tooltip` could map to the Link annotation's `/Contents`
  (accessibility/hover text) but has no visual/pixel impact; `tgtFrame`/
  `history`/`docLocation` are viewer/browser hints with no print representation.
  Owner if tooltip alt-text is ever wanted: `src/docx/runs.rs:594` +
  `LinkAnnotation` in `src/pdf/layout.rs`/`assembly.rs`.
- **17.16.18 `w:fldChar` `@w:fldLock`** — "do not recalculate" flag. docxside
  shows cached results for all but 4 codes anyway, so honoring `fldLock` would
  only matter once broader live evaluation exists. No action.
- **17.16.13 `w:delInstrText` (Deleted Field Code)** — instruction text inside a
  tracked-deletion. docxside renders the final/accept-all view and already drops
  `w:del` content (`src/docx/runs.rs:617`), so deleted field instructions are
  correctly ignored. No action.
- **17.16.14 `enabled`, 17.16.15/.16 `entryMacro`/`exitMacro`, `calcOnExit`** —
  form-field behavior/scripting metadata; no print representation. Subsumed by
  the `ffData` gap above (parse-and-ignore these even if form fields are
  implemented). No standalone action.
- **17.16.20 `format` / 17.16.34 `type` / 17.16.26 `maxLength`** — text-box
  form-field input constraints (numeric/date formatting, length cap) applied only
  during interactive editing; for a static render the cached value already
  reflects them. Subsumed by the `ffData` gap. No standalone action.
- **Field codes with no static representation** — e.g. ASK / FILLIN
  (§17.16.5.3/§17.16.5.19, interactive prompts), MACROBUTTON, GOTOBUTTON,
  AUTOTEXT triggers: Word's PDF export renders only their cached result (or
  nothing), which is exactly docxside's current behavior. Leave as-is.

---

## 17.18 Simple Types — gaps

17.18 defines value domains, not elements. The gaps below are *value-domain
truncations*: docxside parses the owning element but accepts only a subset of
the enumeration values the spec allows; values outside the subset silently
collapse to a default. None of these is a missing element — the rendering path
already exists, it just ignores the distinction between enum values.

### MEDIUM priority

#### 17.18.99 `ST_Underline` (Underline Patterns) — value domain truncated to 2 of 17
- **Spec**: §17.18.99. The `w:u @w:val` underline style (§17.3.2.40, applied via
  the `w:u` run property) has **17 enumeration values**: `single`, `words`
  (non-space chars only), `double`, `thick`, `dotted`, `dottedHeavy`, `dash`,
  `dashedHeavy`, `dashLong`, `dashLongHeavy`, `dotDash`, `dashDotHeavy`,
  `dotDotDash`, `dashDotDotHeavy`, `wave`, `wavyHeavy`, `wavyDouble`, plus `none`.
- **Today**: `RunFormat` models underline as just two booleans — `underline: bool`
  + `double_underline: bool` (`src/model/mod.rs:318`, parsed
  `src/docx/runs.rs:352-359`): `underline = (val != "none")` and
  `double_underline = (val == "double")`. So `single` and `double` render
  correctly, but the other **15 styles** (`thick`, all `dotted`/`dash*`/
  `dotDash*`/`wave*` variants, and `words`) all collapse to a **plain single
  solid line**. Rendered as a thin filled rectangle in the decoration pass in
  `src/pdf/layout.rs` (underline at `y − 0.12·fs`; the double case draws a
  second rule).
- **What exists vs missing**: single + double lines exist; dotted/dashed/wavy/
  thick dash patterns and the `words` (skip-spaces) variant do not. Visually
  noticeable for `wave` (grammar/spell decorations), `dotted`, and `thick`.
- **Likely owner**: replace the two booleans with an underline-style enum on
  `Run` (`src/model/mod.rs`), map the full `ST_Underline` value set in
  `RunFormat::resolve_run_format` (`src/docx/runs.rs`), and emit the
  corresponding stroke pattern (dash array / double rule / sine wave) in the
  decoration pass of `src/pdf/layout.rs`. `words` additionally requires
  suppressing the rule under space glyphs (the existing
  `pending_space_underline` logic at `src/pdf/layout.rs:580` is the hook).

#### 17.18.2 `ST_Border` — plain-line styles truncated to `Single`
- **Spec**: §17.18.2. The `@w:val` of every CT_Border (paragraph `w:pBdr`, table
  `w:tblBorders`/`w:tcBorders`, page `w:pgBorders`) is `ST_Border`, which lists
  ~26 plain-line styles (`single`, `thick`, `double`, `triple`, `dotted`,
  `dashed`, `dotDash`, `dotDotDash`, `dashSmallGap`, `dashDotStroked`,
  `thinThickSmallGap`/`thickThinSmallGap`/`thinThickThinSmallGap` and their
  Medium/Large-gap siblings, `wave`, `doubleWave`, `threeDEmboss`,
  `threeDEngrave`, `outset`, `inset`, `nil`, `none`) **plus ~160 "art" image
  borders**.
- **Today**: `parse_border_style` (`src/docx/mod.rs:311-317`) maps only
  `dotted`, `dashed`, `dashSmallGap`, `dashDotStroked`/`dashDot`, `dashDotDot`,
  and `double`; **everything else falls through to `BorderStyle::Single`**
  (`:317`). So `thick`, `triple`, all `thinThick*`/`thickThin*` multi-line
  styles, `wave`/`doubleWave`, `threeDEmboss`/`threeDEngrave`, and `outset`/
  `inset` collapse to a single hairline.
- **What exists vs missing**: single/dotted/dashed/double exist; the
  multi-stroke, wave, 3-D, and weighted-line plain styles do not. (The ~160 art
  image borders are a separate, larger gap already recorded under §17.6.10
  `w:pgBorders` — this entry is the finer *plain-line* truncation, which affects
  paragraph and table borders too, not just page borders.)
- **Likely owner**: extend the `BorderStyle` enum (`src/model/mod.rs`) and
  `parse_border_style` (`src/docx/mod.rs:308`), then render the extra stroke
  geometries where borders are drawn (`src/pdf/layout.rs` for paragraph borders,
  `src/pdf/table.rs` for table borders).

### LOW priority

#### 17.18.44 `ST_Jc` — `distribute` and kashida values collapse to left-align
- **Spec**: §17.18.44. The paragraph `w:jc @w:val` (§17.3.1.13) domain is
  `start`, `center`, `end`, `both`, `distribute`, `mediumKashida`,
  `highKashida`, `lowKashida`, `thaiDistribute`, `numTab`, plus the legacy
  `left`/`right`. `both` is normal justification; `distribute` additionally
  justifies the *last* line and adds inter-character spacing (East Asian);
  the kashida values stretch Arabic via tatweel.
- **Today**: the IR `Alignment` enum has only `Left`/`Center`/`Right`/`Justify`
  (`src/model/mod.rs:12-18`); `parse_alignment` (`src/docx/styles.rs:221-224`)
  maps `center`→Center, `right`/`end`→Right, `both`→Justify, and **everything
  else → Left**. `start`/`left`→Left and `end`/`right`→Right are correct for LTR
  text, but `distribute`, `thaiDistribute`, `mediumKashida`, `highKashida`,
  `lowKashida`, and `numTab` all silently become left-aligned.
- **What exists vs missing**: the four common alignments work; full-justify
  distribute and the Arabic/Thai kashida modes do not. Niche outside CJK/Arabic
  documents (pairs with the open 17.3 bidi/RTL cluster).
- **Likely owner**: add `Distribute` (and optionally kashida) to `Alignment`
  (`src/model/mod.rs`), extend `parse_alignment` (`src/docx/styles.rs`), and
  honor last-line + inter-glyph distribution in line layout
  (`src/pdf/layout.rs`).

### Value-domain gaps already captured under earlier chapters (record-only — do NOT duplicate)

These are enumeration truncations that the relevant element-chapter audit
already documented; listed here only so the 17.18 catalog cross-references them:
- **17.18.59 `ST_NumberFormat`** — ~8 of ~60 numbering formats handled (decimal,
  roman, letter, bullet, …); the rest fall back. See **17.9 Numbering — gaps**.
- **17.18.93 `ST_TextDirection`** — only `lrTb` honored; `tbRl`/`btLr`/vertical
  variants ignored at both section and cell level. See **17.4 / 17.6**.
- **17.18.102 `ST_VerticalJc`** — handled for table cells, NOT for `w:sectPr`
  (page vertical alignment). See **17.6 Sections — gaps** (`w:vAlign`).
- **17.18.91 `ST_TextAlignment`** — only the implicit `baseline` mode;
  `top`/`center`/`bottom`/`auto` ignored. See **17.3 — gaps** (`w:textAlignment`).
- **17.18.24 `ST_Em`** (emphasis marks) — not handled at all. See **17.3 — gaps**.
- **17.18.62/.63 page-border display/offset + `ST_Border` art borders** — see
  **17.6.10 `w:pgBorders`**.

### No audit question (record-only)

- **Scalar / measure / hex / boolean simple types** carry no enumeration to
  truncate — they are parsed inline at each consumption site and are correct by
  construction: `ST_TwipsMeasure`/`ST_SignedTwipsMeasure` (→ `twips_to_pts`),
  `ST_HpsMeasure`/`ST_SignedHpsMeasure` (half-points → pts), `ST_DecimalNumber`/
  `ST_UnsignedDecimalNumber`, `ST_HexColor`/`ST_LongHexNumber`/`ST_ShortHexNumber`/
  `ST_UcharHexNumber` (→ `parse_hex_color` / hex parse), `ST_OnOff` (→ `wml_bool`,
  handles `0`/`false`/`off`/`1`/`true`/`on`), `ST_Guid` (§17.8.1 font-obfuscation
  decode), `ST_String`, `ST_DateTime`, `ST_PixelsMeasure`,
  `ST_MeasurementOrPercent`. Nothing to do.
- **Enumerations on elements docxside intentionally does not render** (per the
  earlier chapters' record-only lists) — e.g. `ST_Proof`/`ST_ProofErr`
  (§17.18.69/.70, spell/grammar state), `ST_InfoTextType` (§17.18.43, form-field
  help text), `ST_DocProtect`/`ST_Mailmerge*` families — have no print
  representation and need no value-domain coverage.

---

## 20.1 DrawingML Main — gaps

### HIGH priority (corpus-present, real visual impact)

#### 20.1.8.55 `a:srcRect` (Source Rectangle / image crop) — NOT handled
- **Spec**: §20.1.8.55. Child of a blip fill; defines the rectangular window of the
  source image that is actually shown — i.e. image **cropping**. `@l`/`@t`/`@r`/`@b`
  are each a percentage (ST_Percentage, 1000ths of a percent / actually 1/1000 %)
  of the image edge to clip off that side. With a `srcRect`, only the inner window
  is scaled into the picture frame; without it the whole image is shown.
- **Today**: the string `srcRect` appears NOWHERE in `src/` (`grep -rln srcRect src/`
  → 0). Inline/floating pictures parse `a:blip @r:embed` (`src/docx/images.rs:323`)
  and scale the FULL image into the `a:ext` box. A cropped image therefore renders
  with the cropped-away regions still visible and the wrong content scaling.
- **Corpus**: **13 fixtures** contain `a:srcRect` (`analyze-fixtures --grep
  "a:srcRect"`) — by far the most common unhandled DrawingML feature. Any of those
  with a non-zero crop renders wrong.
- **Likely owner**: parse `a:srcRect` (and the sibling `a:stretch`/`a:fillRect`
  vs `a:tile` fill mode, §20.1.8.56/.30/.58) in `src/docx/images.rs` blip-fill
  parsing; apply the crop when drawing the image in `src/pdf/` (either pre-crop the
  decoded bitmap, or emit a PDF clip path around the scaled image so only the
  source window shows). Note: PDF has no native blip-crop, so the renderer must
  translate the percentages into an image-space clip / sub-image.

#### 20.1.8.58 `a:tile` (Tile fill mode) — NOT handled
- **Spec**: §20.1.8.58. Alternative blip fill mode to `a:stretch`: the source image
  is repeated (tiled) across the fill bounds at its native size with `@tx`/`@ty`
  offset, `@sx`/`@sy` scale, `@flip`, `@algn`. Used for textured/wallpaper fills.
- **Today**: not parsed (`grep "\"tile\"" src/docx/images.rs` → none). Images are
  always stretched to the frame.
- **Corpus**: 1 fixture contains `a:tile`. Lower impact than `srcRect`.
- **Likely owner**: `src/docx/images.rs` blip-fill parse + tiled draw in `src/pdf/`.

#### 20.1.8.48 `a:prstDash` (Preset line dash) on shape lines — NOT handled
- **Spec**: §20.1.8.48 (value domain ST_PresetLineDashVal §20.1.10.49: `solid`,
  `dot`, `dash`, `dashDot`, `lgDash`, `lgDashDot`, `lgDashDotDot`, `sysDash`,
  `sysDot`, `sysDashDot`, `sysDashDotDot`). Child of `a:ln`; makes a shape/connector
  outline dashed instead of solid.
- **Today**: shape/textbox line parsing (`src/docx/textbox.rs`, `src/docx/images.rs`
  `a:ln`) reads only width + color; no dash. (`dash`/`Dash` in `src/` is confined to
  table/paragraph **border** styles and **chart** line styles — `docx/tables.rs`,
  `docx/charts.rs`, `pdf/charts.rs` — none of which is the DrawingML shape `a:ln`.)
  `a:custDash` (§20.1.8.21, explicit dash-stop list) likewise absent.
- **Corpus**: 2 fixtures contain `a:prstDash` (`brazilian_logistics_study`,
  `classroom_weekly_newsletter`). A dashed shape outline currently draws solid.
- **Likely owner**: extend the `a:ln` parse in `src/docx/textbox.rs` (+ the picture
  border `a:ln` in `src/docx/images.rs`) to carry a dash pattern on the shape's
  line model (`src/model/drawing.rs`), and emit a PDF dash array (`d` operator) when
  stroking in `src/pdf/`.

### MEDIUM priority (no corpus instance today, but in-spec & visible when present)

#### 20.1.2.2.24 `a:ln` line styling beyond width+color — `@cap`/`@cmpd`/`@algn` + joins — PARTIAL
- **Spec**: §20.1.2.2.24. `a:ln` also carries `@cap` (ST_LineCap: `rnd`/`sq`/`flat`),
  `@cmpd` (ST_CompoundLine: `sng`/`dbl`/`thickThin`/`thinThick`/`tri` — double/triple
  parallel lines), `@algn` (center vs inset stroke), and a line-join child
  `a:round`/`a:bevel`/`a:miter` (§20.1.8.52/.9/.43). It can also be filled with a
  gradient or pattern (`a:gradFill`/`a:pattFill` inside `a:ln`), not just solid.
- **Today**: only `@w` (width) + a solid/noFill color are read (`textbox.rs:384`,
  `images.rs:149`). Cap/compound/join/align and gradient/pattern line fills ignored
  → compound (double) borders draw as a single line; non-default caps/joins use PDF
  defaults.
- **Corpus**: `a:cap=` → 0 fixtures; compound lines untested. Latent.
- **Likely owner**: `src/docx/textbox.rs` + `src/model/drawing.rs` line model; map to
  PDF `J` (cap) / `j` (join) / `M` (miter limit) operators and multi-stroke for
  `cmpd` in `src/pdf/`.

#### `a:bodyPr` text rotation / vertical text — `@rot` and `@vert` — NOT handled
- **Spec**: text-body properties (clause 21.1.2, `a:bodyPr`): `@rot` (text block
  rotation in 60000ths of a degree) and `@vert` (ST_TextVerticalType: `horz`,
  `vert`, `vert270`, `wordArtVert`, `eaVert`, `mongolianVert`, `wordArtVertRtl`) —
  rotate or stack the text inside a shape/textbox.
- **Today**: `a:bodyPr` parsing (`src/docx/textbox.rs:406`) reads `@anchor`, `@wrap`,
  and the `lIns/tIns/rIns/bIns` insets, plus `@anchorCtr` is NOT read. `@rot`/`@vert`
  are absent → vertical-text and rotated-text textboxes render horizontally.
- **Corpus**: `a:vert=` → 0 fixtures (note: shape `bodyPr` in the corpus is usually
  `wps:bodyPr`, not `a:bodyPr`, so the `a:bodyPr` grep shows 0; the same attributes
  apply). Latent.
- **Likely owner**: `src/docx/textbox.rs` body-properties parse; apply a text-matrix
  rotation / vertical glyph advance in `src/pdf/` text-in-shape layout.

#### 20.1.8.47 `a:pattFill` (Pattern fill) — NOT handled
- **Spec**: §20.1.8.47. Two-color preset hatch/pattern fill: `@prst`
  (ST_PresetPatternVal §20.1.10.51 — `pct5`, `dkHorizontal`, `cross`, `diagBrick`,
  …) with `a:fgClr` (§20.1.8.27) + `a:bgClr` (§20.1.8.10).
- **Today**: not parsed for DrawingML shapes (`grep "pattFill" src/` → 0). (Word
  table/cell `w:shd` hatch patterns ARE handled separately in `src/pdf/table.rs`,
  but that is the WML shading path, not DrawingML `a:pattFill`.) A pattern-filled
  shape renders with no fill.
- **Corpus**: 0 fixtures. Latent.
- **Likely owner**: `src/docx/textbox.rs` fill parse + a hatch-pattern tiling paint
  in `src/pdf/` (could reuse the WML hatch logic).

#### 20.1.8.35 `a:grpFill` (Group fill) — NOT handled
- **Spec**: §20.1.8.35. "Inherit the parent group shape's fill." Rare.
- **Today**: not parsed (`grep "grpFill" src/` → 0); a `grpFill` shape gets no fill.
- **Corpus**: 0 fixtures. Latent. Owner: `src/docx/textbox.rs`/`src/docx/group.rs`.

### LOW priority (rare / niche / no-op today)

- **20.1.2.3 alternate color models `a:scrgbClr` / `a:hslClr`** — NOT handled
  (`grep` → 0). `a:srgbClr`/`a:schemeClr`/`a:prstClr`/`a:sysClr` cover essentially
  all real documents; linear-RGB (`scrgbClr`) and HSL (`hslClr`) color specs would
  fall through to a default. 0 corpus fixtures. Owner: `src/docx/color.rs`.
- **20.1.4.1.4 `a:fmtScheme` + 20.1.6.3 `a:clrMap` (explicit theme parsing)** —
  PARTIAL. Theme **color** and **font** schemes are consumed indirectly via
  `resolve_theme_color_key` / major-minor font lookup in `src/docx/styles.rs`, but
  the format scheme (`fmtScheme`: theme fill/line/effect style lists referenced by
  a shape's `<wps:style>`/`a:lnRef`/`a:fillRef`/`a:effectRef` indices) and the
  color map (`clrMap`, which remaps `bg1/tx1/bg2/tx2` slots) are not parsed. A shape
  that styles itself purely by theme style-matrix reference (no explicit `spPr`
  fill/line) therefore gets default styling. Owner: theme parse in
  `src/docx/styles.rs` + style-ref resolution in `src/docx/textbox.rs`.
- **20.1.5 3D — `a:sp3d` / `a:bevelT`-`a:bevelB` / `a:lightRig` / full `a:scene3d`**
  — NOT handled (`grep` → 0; only a `a:camera/a:rot @rev` spin is read for image
  rotation in `src/docx/images.rs:131-137`). True extruded/beveled 3D shapes and
  lighting are unmodeled. 0 corpus fixtures; very large effort with little Word-PDF
  payoff (Word flattens most 3D in export). Owner: `src/geometry/*` + `src/pdf/`.
- **20.1.8 effects on shapes/text (vs pictures)** — `a:effectLst` →
  `outerShdw`/`glow`/`reflection`/`softEdge` is parsed for **pictures**
  (`src/docx/images.rs:217-269`) and a w14 effect set for **WordArt text**
  (`src/docx/wordart.rs`), but the same `a:effectLst` on a plain autoshape/textbox
  `a:spPr` is not wired through `src/docx/textbox.rs`. Also absent: `a:effectDag`
  (§20.1.8.25, DAG of effects), `a:prstShdw` (§20.1.8.49), standalone `a:blur`
  (§20.1.8.15). 2 fixtures use `a:effectLst` (both on pictures, already handled).
  Owner: route `a:effectLst` through `src/docx/textbox.rs` shape parsing.
- **21.1.2 `a:lstStyle` (text list/default paragraph styles in a text body)** —
  NOT parsed (`grep` → 0; 0 fixtures). Defines per-level default run/paragraph
  formatting for shape text; absent means shape paragraphs use only their explicit
  `a:pPr`/`a:rPr`. Owner: `src/docx/textbox.rs` text-body parse.
- **`a:bodyPr @anchorCtr` and `a:noAutofit`** — NOT read (0 corpus). `anchorCtr`
  centers the text block horizontally within the shape; `noAutofit` disables
  shrink-to-fit. Minor. Owner: `src/docx/textbox.rs`.
- **20.1.9.x gradient path fill `a:path` (radial/rectangular gradients)** — only the
  **linear** `a:lin` gradient is implemented (`src/docx/textbox.rs:88`); a
  path/shape gradient (`<a:gradFill><a:path path="circle|rect|shape">`) falls back
  to (or ignores) the gradient. 1 fixture has `a:gradFill` (linear). Owner:
  `src/docx/textbox.rs` gradient parse + `src/pdf/` radial-shading emit.
- **`a:custGeom` guide formulas** — custom geometry path commands are evaluated, but
  per the geometry engine only simple `<a:gd fmla="val N">` guides are read; complex
  guide formulas (`*/ +- sin cos tan at2 …`) inside a `custGeom`'s own `a:gdLst`
  are not (the preset-shape formula interpreter in `src/geometry/formulas.rs` is not
  invoked for inline custom geometry). 0 corpus `a:custGeom` fixtures. Owner:
  `src/docx/textbox.rs` custom-geometry parse → `src/geometry/formulas.rs`.

---

## 20.2 DrawingML Picture — gaps

### HIGH priority (common, real visual impact, present in failing fixtures)

#### 20.2.2.1 `a:srcRect` (Source Rectangle / picture crop) — NOT handled
- **Spec**: §20.2.2.1 `blipFill` is CT_BlipFillProperties, whose `a:srcRect`
  child crops the source image before it is drawn. Attributes `@l`/`@t`/`@r`/`@b`
  (ST_Percentage, in 1000ths of a percent) give the fraction of the image to clip
  off each edge; the remaining sub-rectangle is then scaled into the display
  extent. E.g. `<a:srcRect l="10000" r="20000"/>` drops the left 10% and right 20%
  of the bitmap.
- **Today**: `grep "srcRect" src/` → 0 hits. `EmbeddedImage`
  (`src/model/drawing.rs:21`) has NO crop field; `read_image_from_zip` /
  `read_image_from_zip_extra` (`src/docx/images.rs:271`/`:281`) load the full
  bitmap and the renderer draws the whole image stretched to the extent. A cropped
  picture therefore renders **uncropped** — wrong content shown and wrong
  apparent scale (the kept region is squashed to fit while the cropped-away
  borders are still painted).
- **Corpus**: `analyze-fixtures --grep "srcRect"` → **13 fixtures**, several of
  them currently FAILING — `russian_sports_ranking_decree` (10.0%),
  `mongolian_human_rights_law` (14.1%), `brazilian_logistics_study` (16.8%, 14
  occurrences) — so this is a live contributor to low scores, not latent.
- **Likely owner**: parse `a:srcRect @l/@t/@r/@b` in the blipFill where
  `find_blip_embed` runs (`src/docx/images.rs`), add a crop field (e.g.
  `crop: Option<[f32;4]>` as fractional l/t/r/b) to `EmbeddedImage`
  (`src/model/drawing.rs`), and apply it when drawing the image XObject in
  `src/pdf/` — either by clipping + offset/scaling the placed image so only the
  source sub-rectangle shows, or by pre-cropping the decoded bitmap. NOTE:
  `srcRect` also appears in shape `a:blipFill` (picture-fill of an autoshape), a
  separate path in `src/docx/textbox.rs`; the crop semantics are identical and
  could share a helper.

### MEDIUM priority (recolor / transform — wrong colors or orientation)

#### 20.2.2.1 blip recolor & adjust ops (`a:lum`/`a:clrChange`/`a:duotone`/`a:grayscl`/`a:biLevel`/`a:alphaModFix`) — NOT handled
- **Spec**: §20.2.2.1 `a:blip` may carry child image-transform effects (defined in
  DrawingML §20.1.8): `a:lum` (§20.1.8.42, brightness `@bright` + contrast
  `@contrast`), `a:clrChange` (§20.1.8.16, replace one color with another / set a
  transparent color via `clrFrom`→`clrTo`), `a:duotone` (§20.1.8.23, map shadows
  and highlights to two colors), `a:grayscl` (§20.1.8.34, desaturate), `a:biLevel`
  (§20.1.8.11, threshold to black/white), `a:alphaModFix` (§20.1.8.4, uniform
  opacity). These modify the rendered pixels of the picture.
- **Today**: none parsed (`grep` for each → 0 hits). `find_blip_embed` reads only
  `a:blip @r:embed`; any child effects are ignored, so a picture that Word shows
  recolored / grayscaled / brightened / partially transparent renders at its
  **original** colors and full opacity.
- **Corpus**: `a:lum` in **2 fixtures**; `clrChange`/`duotone`/`grayscl`/`biLevel`/
  `alphaModFix` in **0** (latent). So the brightness/contrast (`a:lum`) case is the
  only one with current corpus impact; the rest are conformant-but-unexercised.
- **Likely owner**: parse the blip child effects alongside `find_blip_embed`
  (`src/docx/images.rs`), model them on `EmbeddedImage` (`src/model/drawing.rs`),
  and apply in `src/pdf/` image emission. `a:alphaModFix` (uniform alpha) and
  `a:lum` (brightness/contrast as a PDF transfer/`gs` adjustment) are the cheapest;
  `clrChange`/`duotone`/`grayscl`/`biLevel` need per-pixel bitmap manipulation
  before encoding the XObject.

#### 20.2.2.6 `pic:spPr/a:xfrm @flipH`/`@flipV` (picture flip) — NOT applied to raster pictures
- **Spec**: §20.2.2.6 picture shape properties include the standard `a:xfrm`
  transform, whose `@flipH`/`@flipV` booleans mirror the picture horizontally /
  vertically before placement.
- **Today**: `grep "flipH\|flipV" src/` shows flip is honored only for **connector
  lines** (`src/docx/textbox.rs:510`, `src/pdf/positioning.rs:249`). The picture
  transform reader `parse_image_rotation` (`src/docx/images.rs:120`) reads `@rot`
  (and a scene3d-camera fallback) but NOT `@flipH`/`@flipV`, and `EmbeddedImage`
  has no flip field — so a mirrored picture renders un-mirrored. (Shared with the
  20.1 `a:xfrm` transform; called out here because the picture path specifically
  drops it.)
- **Corpus**: not separately counted; flips are uncommon on pictures but trivially
  authored ("Rotate → Flip Horizontal" in Word).
- **Likely owner**: extend `parse_image_rotation` (or a sibling) in
  `src/docx/images.rs` to read `@flipH`/`@flipV`, carry on `EmbeddedImage`
  (`src/model/drawing.rs`), and negate the X/Y scale of the image matrix in
  `src/pdf/` when placing the XObject.

### LOW priority (rare / latent in current corpus)

- **20.2.2.1 `a:blip @r:link` (linked / external image)** — NOT handled. The blip
  may reference an image **outside** the package by a `r:link` relationship
  (External target) instead of `r:embed`. `find_blip_embed`
  (`src/docx/images.rs:323`) reads only `@r:embed`, so a link-only picture yields
  no image. 1 corpus fixture has `r:link`. A portable render usually cannot fetch
  the external file, so the realistic fix is parse-and-skip gracefully (no panic /
  no zero-size box). Owner: `src/docx/images.rs`.
- **20.2.2.1 fill mode `a:tile` vs `a:stretch`/`a:fillRect`** — NOT handled. The
  blipFill fill mode is ignored; the image is always drawn stretched into the
  extent. `a:tile` (repeat the blip to fill) renders as a single stretched copy.
  1 corpus fixture has `a:tile`. Low impact for inline pictures (stretch is the
  norm) — matters mainly for picture-*fill* of shapes. Owner: `src/docx/images.rs`
  blipFill parse + `src/pdf/` tiled-pattern emit.
- **20.2.2.3 `pic:cNvPr/a:hlinkClick @r:id` (clickable picture hyperlink)** — NOT
  handled. `pic:nvPicPr` is never descended into, so a picture wrapped in a
  hyperlink emits no PDF Link annotation. 0 corpus fixtures (`hlinkClick` → 0),
  latent. The text-run hyperlink → Link-annotation machinery already exists
  (`LinkAnnotation` in `src/pdf/layout.rs`, emitted in `src/pdf/assembly.rs`); this
  would parse `cNvPr/a:hlinkClick @r:id`, resolve the relationship, and attach a
  GoTo/URI Link annotation over the picture's bounding box. Owner:
  `src/docx/images.rs` (parse) + `src/pdf/` (emit over image rect).
- **20.2.2.3 `pic:cNvPr @descr` (alt text) / @title** — NOT read. Accessibility /
  PDF-tagging metadata only; no visual representation in an untagged PDF. Safe to
  leave unparsed unless tagged-PDF / `/Alt` text output is later wanted (owner
  `src/docx/images.rs` → PDF structure tree). Record-only.
- **20.2.2.1 `@cstate` (compression-state hint) / 20.2.2.2 `cNvPicPr` +
  `a:picLocks` (edit locks) / `@preferRelativeResize`** — authoring/UI metadata
  with no print representation. Correctly ignored (0 hits). Record-only.

---

## 20.4 DrawingML WordprocessingML Drawing — gaps

### MEDIUM priority (one is exercised by the corpus; rest are common-in-the-wild)

#### 20.4.2.3 `wp:anchor @layoutInCell` (Lay Out In Table Cell) — NOT parsed → corpus-exercised
- **Spec**: §20.4.2.3, attribute on `wp:anchor`, ST_OnOff, default `true`. When
  `1`, a floating object anchored inside a table cell is positioned **relative to
  the cell** (its `relativeFrom` page/margin references are clamped to the cell's
  extents). When `0`, the object positions relative to the page/margin and may
  overflow the cell — Word lets it float free of the cell box.
- **Today**: `grep "layoutInCell" src/` → 0 hits. Never parsed; floating objects
  inside a cell are always laid out with the in-cell behavior implied by the
  existing positioning code. No model field.
- **Corpus**: `analyze-fixtures --grep 'layoutInCell="0"'` → **2 fixtures** set the
  non-default value (only 20.4 gap with current corpus impact). All 15 anchored
  fixtures carry the attribute (default `1`).
- **Likely owner**: parse `@layoutInCell` on the anchor in
  `src/docx/images.rs` (`parse_anchor_position`/the `:493` anchor branch), carry a
  bool on `FloatingImage` (`src/model/drawing.rs`), and in
  `src/pdf/positioning.rs` / `src/pdf/table.rs` decide whether the float's
  page/margin reference is clamped to the cell rect or to the page.

#### 20.4.2.9/.10 `wp:positionH`/`wp:positionV @relativeFrom` — enum truncated
- **Spec**: §20.4.2.9 (`positionH @relativeFrom`, ST_RelFromH §20.4.3.4) =
  `margin`|`page`|`column`|`character`|`leftMargin`|`rightMargin`|`insideMargin`|
  `outsideMargin`; §20.4.2.10 (`positionV @relativeFrom`, ST_RelFromV §20.4.3.5) =
  `margin`|`page`|`paragraph`|`line`|`topMargin`|`bottomMargin`|`insideMargin`|
  `outsideMargin`.
- **Today** (`src/docx/images.rs:345-368`): H accepts only `page`/`margin`,
  everything else (incl. `column`) falls to `Column`; V accepts only
  `page`/`margin`/`topMargin`, everything else (incl. `paragraph`) falls to
  `Paragraph`. Missing H values `character`,`leftMargin`,`rightMargin`,
  `insideMargin`,`outsideMargin` and V values `line`,`bottomMargin`,`insideMargin`,
  `outsideMargin` silently collapse to the column/paragraph default — wrong origin
  for any object that uses them. Model enums `HRelativeFrom`/`VRelativeFrom`
  (`src/model/mod.rs:188-200`) only have 3 / 4 variants.
- **Corpus**: only `column`(9)/`paragraph`(10)/page/margin appear — the exotic
  values are 0 fixtures, so this is LATENT today.
- **Likely owner**: extend the two enums in `src/model/mod.rs`, the two `match`
  arms in `src/docx/images.rs:345`/`:363`, and the origin-resolution in
  `src/pdf/positioning.rs` (`leftMargin`/`rightMargin`/`insideMargin`/
  `outsideMargin` need the gutter/facing-page geometry; `character`/`line` need the
  anchor run's x / line-top, which positioning code does not currently track).

#### 20.4.2.1/.2 `wp:align` value `inside`/`outside` — NOT handled
- **Spec**: §20.4.2.1 (`align` in positionH, ST_AlignH) = `left`|`right`|`center`|
  `inside`|`outside`; §20.4.2.2 (`align` in positionV, ST_AlignV) = `top`|`bottom`|
  `center`|`inside`|`outside`. `inside`/`outside` align to the inner / outer margin
  on facing (mirrored) pages — i.e. left on odd pages, right on even (or vice
  versa).
- **Today** (`src/docx/images.rs:350-379`): H maps only center/right/(else left);
  V maps only bottom/center/(else top). `inside`/`outside` fall to the left/top
  default, ignoring facing-page mirroring.
- **Corpus**: `<wp:align>inside`/`outside` → 0 fixtures. LATENT.
- **Likely owner**: `parse_anchor_position` in `src/docx/images.rs` (needs new
  `HorizontalPosition`/`VerticalPosition` variants in `src/model/mod.rs:172-185`)
  + facing-page resolution in `src/pdf/positioning.rs` keyed on logical page parity
  (the same parity already computed for even/odd headers, `src/pdf/mod.rs:3086`).

### LOW priority (latent — absent or default-valued in current corpus)

- **20.4.2.3 `wp:anchor @simplePos` (attribute) + `wp:simplePos` (element,
  §20.4.2.12)** — NOT parsed. When `@simplePos="1"`, the object ignores
  `positionH`/`positionV` and is placed at the absolute `wp:simplePos @x/@y` (EMU)
  coordinates. `grep "simplePos" src/` → 0 hits; the `wp:simplePos` element is
  present in 15 fixtures but always with `@simplePos="0"` on the anchor (0 fixtures
  set `="1"`), so positionH/V is always used and output is correct today. Owner:
  `src/docx/images.rs` (read the attr in the `:493` anchor branch; if set, parse
  `wp:simplePos @x/@y` instead of `parse_anchor_position`) + a Page/Page absolute
  placement in `src/pdf/positioning.rs`.
- **20.4.2.3 `wp:anchor @allowOverlap` — NOT parsed.** ST_OnOff, default `true`.
  When `0`, Word must displace this float so it does not overlap other floats.
  `grep` → 0 hits; `allowOverlap="0"` in 0 corpus fixtures (all default `1`).
  Latent. Owner: `src/docx/images.rs` (parse) + overlap-displacement logic in
  `src/pdf/positioning.rs` (none exists today — floats may currently overlap
  regardless).
- **20.4.2.3 `wp:anchor @locked` — NOT parsed; no print impact.** Editing-only
  anchor lock (prevents the anchor moving during edits). Correctly ignorable for a
  static render. Record-only.
- **20.4.2.5 `wp:docPr/a:hlinkClick @r:id` (clickable drawing hyperlink)** — NOT
  handled. The WordprocessingML-drawing-level hyperlink (distinct element from the
  picture-level `pic:cNvPr/a:hlinkClick` already noted under **20.2 — gaps**, but
  same machinery): `wp:docPr` is read only for `@hidden` (`src/docx/images.rs:478`),
  never for a child `a:hlinkClick`, so a hyperlinked drawing emits no PDF Link
  annotation. `analyze-fixtures --grep "hlinkClick"` → 0 fixtures, LATENT. The
  text-run hyperlink → Link-annotation path already exists (`LinkAnnotation` in
  `src/pdf/layout.rs`, emitted in `src/pdf/assembly.rs`). Owner: parse
  `wp:docPr/a:hlinkClick @r:id` in `src/docx/images.rs`, resolve the relationship,
  attach a GoTo/URI Link annotation over the drawing's bounding box in `src/pdf/`.
- **20.4.2.5 `wp:docPr @name`/`@descr`/`@title`** — NOT read. Accessibility / PDF-
  tagging metadata (alt text); no visual representation in an untagged PDF. Safe to
  leave unparsed unless `/Alt`-text / tagged-PDF output is later wanted. Record-only.
- **20.4.2.4 `wp:cNvGraphicFramePr` + `a:graphicFrameLocks`** — editing locks /
  non-visual frame properties; no print representation. Correctly ignored.
  Record-only.

---

## 21.2 DrawingML Charts — gaps

### HIGH priority (real content loss / output divergence)

#### Unsupported chart-type elements render NOTHING (chart silently vanishes) — `c:bar3DChart` / `c:line3DChart` / `c:area3DChart` / `c:surfaceChart` / `c:surface3DChart` / `c:stockChart` / `c:ofPieChart`
- **Spec**: §21.2.2 chart-group elements — `bar3DChart` (§21.2.2.17), `line3DChart`
  (§21.2.2.97), `area3DChart` (§21.2.2.4), `surfaceChart` (§21.2.2.204),
  `surface3DChart` (§21.2.2.203), `stockChart` (§21.2.2.198), `ofPieChart`
  (§21.2.2.130, pie-of-pie / bar-of-pie). Each is a sibling of the
  `barChart`/`lineChart`/… elements docxside already detects.
- **Today**: `detect_chart_type` (`src/docx/charts.rs:263`) matches only
  `barChart`, `doughnutChart`, `lineChart`, `pieChart`, `pie3DChart`, `areaChart`,
  `radarChart`, `scatterChart`, `bubbleChart`. For any other group element it
  returns `None` → `parse_chart_space` returns `None` → `parse_chart_from_zip`
  returns `None` → `find_chart_ref`'s caller in `src/docx/images.rs:654` produces
  **no `RunDrawingResult` at all**. The chart (and the paragraph height it should
  occupy) silently disappears — no placeholder box, no fallback image.
- **What exists vs missing**: the 2-D analogues are fully rendered; the 3-D and
  stock/of-pie variants have no detection branch and no renderer. Cheapest partial
  fix: map `bar3DChart`/`line3DChart`/`area3DChart` onto the existing 2-D renderers
  (drop the depth) so *something* draws; `surfaceChart`/`stockChart`/`ofPieChart`
  need new renderers. At minimum, draw a placeholder rectangle so the chart never
  vanishes.
- **Corpus**: 0 fixtures (all 6 corpus charts are 2-D bar/line/pie/area/doughnut),
  so no test regresses — but a 3-D or stock chart in the wild produces a blank gap.
- **Likely owner**: `detect_chart_type` + `ChartType` enum (`src/model/chart.rs`);
  renderers in `src/pdf/charts.rs` / `src/pdf/charts_radial.rs`.

#### Combo charts (multiple chart-group elements in one `c:plotArea`) — only the FIRST group renders
- **Spec**: §21.2.2.145 `plotArea` may contain more than one chart-group child
  (e.g. a `c:barChart` **and** a `c:lineChart` sharing axes = a combo chart). Each
  group has its own `c:ser` list.
- **Today**: `detect_chart_type` returns the FIRST matching group node and
  `collect_series` (`src/docx/charts.rs:189`) iterates `c:ser` only within that one
  node. A second chart-group element and all its series are dropped.
- **What exists vs missing**: single-group charts work; multi-group plot areas lose
  every series after the first group. No model representation for per-series chart
  type.
- **Corpus**: 0 fixtures. Latent.
- **Likely owner**: model would need per-series type (`src/model/chart.rs`);
  parse loop in `src/docx/charts.rs`; render dispatch in `src/pdf/charts.rs`.

#### `c:overlap` (Bar Overlap) — NOT parsed
- **Spec**: §21.2.2.132. Child of `c:barChart`/`c:bar3DChart`; signed percentage
  (−100..100) of how much adjacent bars in a category overlap (negative = gap,
  positive = overlap; clustered Word default is −27 in the UI but the element is
  written explicitly). Governs bar width/placement within a category cluster.
- **Today**: not parsed (`grep "\"overlap\"" src/` → 0). Bars are laid out with a
  fixed internal spacing in `src/pdf/charts.rs` (bar geometry near `:379`); the
  `@val` is ignored, so clustered multi-series bars get the wrong width/gap.
- **Corpus**: **present in case52** (all 4 charts carry `c:overlap`) — the one
  non-latent axis/bar gap. case52's score is a candidate to improve once honored.
- **Likely owner**: parse `gapWidth`-adjacent in `detect_chart_type`
  (`src/docx/charts.rs`), add `overlap_pct` to `Chart`/`ChartType::Bar`
  (`src/model/chart.rs`), consume in the bar-placement math in `src/pdf/charts.rs`.

#### `c:title` + `c:autoTitleDeleted` (Chart Title) — NOT parsed or rendered
- **Spec**: §21.2.2.210 `title` (chart title, holds a `c:tx` rich-text or string
  ref + `c:overlay` + `c:spPr`/`c:txPr`), §21.2.2.27 `autoTitleDeleted` (suppress
  the auto-generated title). A shown title occupies vertical space above the plot
  area and shifts the plot down.
- **Today**: no `c:title` parsing (`grep` for chart `title` → none); the renderer
  reserves no title band, so the plot area fills the whole chart box. A titled
  chart renders with the plot mis-sized and the title text absent.
- **Corpus**: 0 corpus charts carry `c:title`. Latent.
- **Likely owner**: add `title: Option<String>` to `Chart` (`src/model/chart.rs`),
  parse in `src/docx/charts.rs` (`c:title/c:tx/...` mirroring series-name decode),
  reserve a top band + draw centered in `src/pdf/charts.rs`.

#### `c:dLbls` / `c:dLbl` (Data Labels) — NOT parsed or rendered
- **Spec**: §21.2.2.49 `dLbls` (group-level) / §21.2.2.47 `dLbl` (per-point), with
  toggles `c:showVal` (§21.2.2.190), `c:showCatName` (§21.2.2.187), `c:showSerName`
  (§21.2.2.189), `c:showPercent` (§21.2.2.188), `c:showLegendKey` (§21.2.2.186),
  `c:numFmt`, `c:dLblPos` (§21.2.2.48). Draws value/category/percent text on/near
  each data point.
- **Today**: not parsed (`grep "\"dLbls\"" src/` → 0); no labels drawn on bars,
  slices, points.
- **Corpus**: 0 corpus charts carry `c:dLbls`. Latent.
- **Likely owner**: model fields on `ChartSeries`/`Chart` (`src/model/chart.rs`),
  parse in `src/docx/charts.rs`, draw in the per-type render loops (`src/pdf/charts.rs`,
  `src/pdf/charts_radial.rs`); reuse `val_format_code` for value formatting.

### MEDIUM priority (axis fidelity — elements present in corpus but treated as default)

#### `c:scaling` (`c:min` / `c:max` / `c:orientation` / `c:logBase`) — NOT parsed; renderer auto-scales
- **Spec**: §21.2.2.176 `scaling` with children §21.2.2.121 `min`, §21.2.2.95
  `max`, §21.2.2.133 `orientation` (ST_Orientation = `minMax` | `maxMin`, i.e.
  reversed axis), §21.2.2.103 `logBase` (logarithmic axis).
- **Today**: not parsed. The value axis range is computed automatically from the
  data in `src/pdf/charts.rs`; an author-fixed `c:min`/`c:max`, a reversed
  (`maxMin`) axis, or a log scale is ignored.
- **Corpus**: `c:scaling`+`c:orientation` present in every corpus chart, but all
  carry the default `orientation=minMax` and NO explicit `c:min`/`c:max`/`c:logBase`
  (verified: the earlier `c:min` grep hit was a false match on `minorTickMark`). So
  output is correct today only because every value is default — LATENT.
- **Likely owner**: parse into `ChartAxis` (`src/model/chart.rs` — add
  `min`/`max`/`reversed`/`log_base`), in `parse_axis` (`src/docx/charts.rs:162`);
  honor in the axis-range + coordinate-mapping code in `src/pdf/charts.rs`.

#### `c:tickLblPos` / `c:majorTickMark` / `c:minorTickMark` — NOT parsed
- **Spec**: §21.2.2.216 `tickLblPos` (ST_TickLblPos = `high` | `low` | `nextTo` |
  `none` — where axis tick labels sit, incl. *hidden*), §21.2.2.97 `majorTickMark` /
  §21.2.2.123 `minorTickMark` (ST_TickMark = `cross` | `in` | `out` | `none` — tick
  dash style).
- **Today**: not parsed (`grep` → 0). The renderer always draws axis labels at its
  default position and draws no tick-mark dashes. `tickLblPos=none` (a common way to
  hide axis labels) would still show labels — a visible divergence.
- **Corpus**: `c:tickLblPos` + `c:minorTickMark` present in every corpus
  cartesian chart; values not audited but Word defaults (`nextTo`/`none`) likely, so
  mostly cosmetic today. LATENT for the `none` case.
- **Likely owner**: `parse_axis` (`src/docx/charts.rs`) + `ChartAxis`
  (`src/model/chart.rs`); consume in axis rendering in `src/pdf/charts.rs`.

#### `c:crosses` / `c:crossesAt` (axis crossing point) — NOT parsed
- **Spec**: §21.2.2.43 `crosses` (ST_Crosses = `autoZero` | `min` | `max`) /
  §21.2.2.44 `crossesAt` (explicit value at which the perpendicular axis crosses).
  Determines where the category axis sits relative to the value axis (e.g. a baseline
  at a non-zero value, or category labels drawn at the bottom vs. crossing through
  the middle for data with negatives).
- **Today**: not parsed (`grep` → 0). The category axis is always drawn at the
  chart's computed zero/baseline.
- **Corpus**: `c:crossesAt` present in every corpus cartesian chart (value not
  audited). LATENT unless a chart sets a non-zero crossing.
- **Likely owner**: `parse_axis` (`src/docx/charts.rs`) + baseline placement in
  `src/pdf/charts.rs`.

#### `c:numFmt` on an axis (vs. only the `numCache/formatCode`) — partial
- **Spec**: §21.2.2.124 `numFmt` is a direct child of `c:valAx`/`c:catAx` carrying
  `@formatCode` + `@sourceLinked` — the axis's own number format, independent of the
  series `c:numCache/c:formatCode`.
- **Today**: only the **series** `c:numCache/c:formatCode` is read (→
  `val_format_code`, `src/docx/charts.rs:63`); the axis-level `c:numFmt` element is
  not read. When they differ (or there is no numCache format), axis tick labels use
  the wrong/General format.
- **Likely owner**: read `c:numFmt @formatCode` in `parse_axis`
  (`src/docx/charts.rs`), prefer it over the series format for that axis's labels in
  `src/pdf/charts.rs`.

### LOW priority (niche decorations / effects — all absent from corpus)

- **`c:varyColors` (§21.2.2.227)** — NOT parsed. Vary each data point's color
  (pie/doughnut/single-series bar). The renderer already colors pie/doughnut slices
  distinctly from the accent palette, so `varyColors=1` is effectively honored; but
  `varyColors=0` (all points one color) is NOT — slices would still be multicolored.
  Owner: `src/docx/charts.rs` + `src/pdf/charts_radial.rs`.
- **`c:dPt` (§21.2.2.52, Data Point)** — NOT parsed. Per-point format override
  (individual slice/bar `c:spPr` fill, `c:explosion`, `c:bubble3D`). Today every
  point in a series shares the series fill. Owner: `src/model/chart.rs` (per-point
  color vec on `ChartSeries`), `src/docx/charts.rs`, render loops.
- **`c:explosion` (§21.2.2.61)** — NOT parsed. Pie/doughnut slice pulled radially
  out from center. Slices always drawn flush. Owner: `src/pdf/charts_radial.rs`.
- **`c:firstSliceAng` (§21.2.2.68)** — NOT parsed. Pie/doughnut start-angle
  rotation (child of `c:pieChart`/`c:doughnutChart`/`c:pie3DChart`). Renderer starts
  at a fixed angle, so a rotated pie is mis-oriented. Owner: `src/docx/charts.rs` +
  `src/pdf/charts_radial.rs`.
- **`c:smooth` (§21.2.2.194)** — NOT parsed. Line-series Bézier smoothing. Lines
  always drawn as straight segments. Owner: `src/docx/charts.rs` + line render in
  `src/pdf/charts.rs`.
- **`c:grouping` on `c:lineChart`/`c:areaChart` (stacked / percentStacked)** — NOT
  read. `detect_chart_type` (`src/docx/charts.rs:268`) reads `@grouping` ONLY for
  `barChart`; a stacked line or stacked/100%-area chart is rendered as
  overlapping/standard instead of accumulated. Owner: `src/docx/charts.rs` +
  `src/pdf/charts.rs`.
- **`c:view3D` (§21.2.2.228) + floor/sideWall/backWall** — NOT parsed. 3-D
  rotation/perspective/depth. Tied to the unsupported 3-D chart types above. Owner:
  `src/docx/charts.rs` + new 3-D render path.
- **`c:trendline` (§21.2.2.211)** — NOT parsed. Per-series regression/trend line
  (linear/exp/log/poly/movingAvg) + optional equation/R². Not drawn. Owner:
  `src/docx/charts.rs` + `src/pdf/charts.rs`.
- **`c:errBars` (§21.2.2.55)** — NOT parsed. Error bars on a series. Not drawn.
  Owner: `src/docx/charts.rs` + `src/pdf/charts.rs`.
- **`c:upDownBars` / `c:hiLowLines` / `c:dropLines` / `c:serLines`** — NOT parsed.
  Line/stock-chart decorations (up-down bars, high-low lines, drop lines, series
  connector lines). Not drawn. Owner: `src/docx/charts.rs` + `src/pdf/charts.rs`.
- **`c:layout` / `c:manualLayout` (§21.2.2.88 / §21.2.2.110)** — NOT parsed. Manual
  pixel/fraction positioning of the plot area, legend, or title. docxside always
  auto-lays-out these regions, so a chart with a manually-moved legend/plot will not
  match. Owner: `src/docx/charts.rs` (legend/plotArea layout) + `src/pdf/charts.rs`
  / `src/pdf/chart_legend.rs`.
- **`c:userShapes` (§21.2.2.220, Chart Drawing part)** — NOT parsed. Free shapes
  (cdr:relSizeAnchor/absSizeAnchor) drawn on top of a chart via the
  `chartUserShapes` relationship. Never opened. Owner: new parse in
  `src/docx/charts.rs` resolving the `userShapes` rel + render via the existing
  shape/geometry engine.
- **`c:plotVisOnly` (§21.2.2.144) / `c:dispBlanksAs` (§21.2.2.46)** — NOT parsed.
  Whether to plot only visible cells, and how to render blank cells (gap / zero /
  span). No hidden/blank-cell handling; all cached points are plotted as-is. Owner:
  `src/docx/charts.rs`. Mostly no impact since only cached values are available.
- **Legend `c:legendEntry` deletions + `c:overlay`** (§21.2.2.94 / §21.2.2.132-adjacent)
  — NOT parsed. Individual legend entries can be deleted, and the legend can overlay
  the plot. Today all series get a legend entry and the legend never overlays.
  Owner: `src/docx/charts.rs` + `src/pdf/chart_legend.rs`.
- **`c:spPr`/`c:txPr` on `chartSpace`/`chart`/`legend`/axis title (frame fills,
  fonts)** — only series/plot/axis-line solid fills are read. Chart-area background
  fill, rounded-corner frame (`c:roundedCorners` §21.2.2.175), and text properties
  (font/size/color via `c:txPr`) are ignored. Owner: `src/docx/charts.rs` +
  `src/pdf/charts.rs`.

### Record-only (no print impact / not in a Word-exported chart)
- `c:idx` / `c:order` (series index/order), `c:f` (the spreadsheet formula string
  inside every `*Ref` — docxside correctly reads the cache, not the formula),
  `c:extLst`/`mc:AlternateContent` (extensions), `c:externalData`/`c:autoUpdate`
  (link to the embedded XLSX — irrelevant since caches are used), `c:date1904`,
  `c:lang`, `c:printSettings`/`c:pageMargins`/`c:pageSetup`/`c:headerFooter` (chart
  print setup — the chart is embedded in a Word page, not printed standalone). None
  affect the rendered chart; correctly unparsed.

---

## 21.4 DrawingML Diagrams — gaps

### Architecture note — docxside renders the *cached drawing*, not the data model

docxside renders SmartArt by reading the **Microsoft `dsp:` "drawing" part**
(`word/diagrams/drawingN.xml`, namespace
`http://schemas.microsoft.com/office/drawing/2008/diagram`), NOT by interpreting
the standardized clause-21.4 `dgm:` parts. The flow (`src/docx/smartart.rs`):
`has_diagram_ref` matches `a:graphicData @uri =
"http://schemas.openxmlformats.org/drawingml/2006/diagram"` (`:72`) →
`find_diagram_drawing` reads `dgm:relIds @r:dm` (the data-part relationship,
`:90`) and derives the drawing path by string-replacing `/data`→`/drawing`
(`:92`), with a fallback that scans rels for any `diagrams/drawing` target
(`:106`) → `parse_smartart_drawing` reads `dsp:drawing/dsp:spTree`, collects each
`dsp:sp`, and parses position/fill/line/text (`:117-124`). Rendering in
`src/pdf/smartart.rs` (`render_smartart` `:165`): solid fill + stroke + text per
shape with rotation.

This is a **pragmatic and correct shortcut whenever the `dsp:` cache exists** —
Word always writes it, with colors/quick-style/layout already resolved and baked
into each `dsp:sp` (explicit `a:xfrm` positions, resolved `a:solidFill`, resolved
`a:fontRef`). Consequently the four standardized parts are deliberately *not*
read, which is fine **only while the cache is present**:
- `dgm:dataModel` (`word/diagrams/dataN.xml`, §21.4 CT_DataModel: `ptLst`/`pt`/
  `cxnLst`/`cxn`/`bg`/`whole`) — the actual diagram content + connections.
- `dgm:layoutDef` (`layoutN.xml`) — the layout algorithm that maps data→shapes.
- `dgm:colorsDef` (`colorsN.xml`) and `dgm:styleDef` (`quickStyleN.xml`) — color +
  quick-style themes.

`grep -rn "dgm(" src/` shows the `dgm:` namespace is touched only to read
`relIds`; the data/layout/colors/style parts are never parsed (record-only — see
last gap). Corpus: 2 fixtures carry SmartArt — `learning_cultures_dissertation`
(2 diagrams, Jaccard 33.3%) and `vaccines_history_chapter` (1 diagram, 42.9%);
both ship the `dsp:` drawing cache (`drawing1.xml`/`drawing2.xml`), so both render
today and both pass the 20% threshold. Gaps below, grouped by priority.

### HIGH priority

#### No fallback when the `dsp:` drawing cache is absent — diagram silently disappears
- **Spec**: §21.4 standardizes `dgm:dataModel` + `dgm:layoutDef` (+ colors/style).
  The `dsp:` drawing part docxside depends on is a **Microsoft extension** (note
  the `schemas.microsoft.com` namespace), NOT part of ECMA-376 Part 1 / ISO/IEC
  29500-1. A spec-conformant producer is only required to emit the four `dgm:`
  parts; the `dsp:` cache is optional.
- **Today**: `parse_smartart_drawing` (`src/docx/smartart.rs:95`) returns a
  `SmartArtDiagram` with `shapes: Vec::new()` whenever no `diagrams/drawing*`
  target is found (`drawing_target` is `None`). Nothing is rendered → the entire
  SmartArt object **vanishes with no placeholder or error**.
- **What exists vs missing**: full rendering exists *only* for the MS cache;
  there is zero interpretation of `dgm:dataModel` (the text points in `pt/t` and
  the parent/child `cxn` connections) or `dgm:layoutDef`. Documents from
  LibreOffice, Google Docs export, python-docx, or any non-Microsoft toolchain —
  and potentially future strict-conformance Word output — would lose the diagram.
- **Likely owner**: `src/docx/smartart.rs`. A full layout engine (interpreting
  `layoutDef` algorithms) is a large effort and out of proportion for a renderer;
  a **minimal fallback** is the lazy win: when no drawing cache exists, read
  `dgm:dataModel` from the `r:dm` target, walk `ptLst/pt[@type="node"]/prSet` +
  `t/a:p/a:r/a:t` for each node's text, and emit them as a simple stacked list of
  boxes (or plain bulleted paragraphs) so content is not silently dropped. Model
  reuse: `SmartArtShape`/`SmartArtPara` in `src/model/drawing.rs` already suffice
  for a box-list rendering.

### MEDIUM priority

#### `a:effectLst` on `dsp:sp` shapes (shadow/glow/reflection/soft-edge) — NOT parsed
- **Spec**: §20.1.8 effects, attached to each diagram shape's `dsp:spPr`.
- **Today**: `parse_dsp_shape` (`src/docx/smartart.rs:137`) reads fill, line, and
  text but never looks at `a:effectLst`. `SmartArtShape` (`src/model/drawing.rs:225`)
  has no effect fields and `render_smartart` draws no shadows.
- **Corpus**: PRESENT — `learning_cultures_dissertation/word/diagrams/drawing1.xml`
  contains **17 `<a:effectLst>`** (one per shape, typically `outerShdw`). The
  diagram boxes render flat instead of with Word's drop shadows → a real visual
  diff in a currently-passing fixture.
- **What exists vs missing**: pictures already parse the same effects
  (`parse_pic_effects`, `src/docx/images.rs:214` → outerShdw/softEdge/glow/
  innerShdw/reflection). SmartArt shapes need the equivalent.
- **Likely owner**: parse in `src/docx/smartart.rs` (reuse/adapt
  `parse_pic_effects`), add effect fields to `SmartArtShape`, render in
  `src/pdf/smartart.rs`.

#### Gradient / pattern shape fills (`a:gradFill` / `a:pattFill`) on `dsp:sp` — NOT parsed
- **Spec**: §20.1.8.33 `gradFill`, §20.1.8.47 `pattFill`.
- **Today**: `parse_dsp_shape` fill resolution is `parse_solid_fill(sp_pr, theme)`
  only (`src/docx/smartart.rs:162`); `SmartArtShape.fill` is `Option<[u8;3]>` (a
  single solid RGB). A gradient- or pattern-filled diagram shape falls through to
  `None` → rendered with **no fill** (or just outline).
- **Corpus**: LATENT — both SmartArt fixtures use only `solidFill`/`noFill` in
  their drawing parts (verified: 22 `solidFill`, 12 `noFill`, 0 `gradFill` in
  drawing1.xml). But gradient fills are common in Word's built-in SmartArt color
  styles, so real-world documents will hit this.
- **Likely owner**: `src/docx/smartart.rs` fill parsing + `src/model/drawing.rs`
  (`SmartArtShape` needs a richer fill enum) + `src/pdf/smartart.rs`. The linear-
  gradient parsing in `src/docx/color.rs`/`textbox.rs` can be reused.

#### Connector shapes (`dsp:cxnSp`) — dropped
- **Spec**: connector/arrow shapes between diagram nodes (org-chart lines, process
  arrows) are emitted in the drawing part as connector shapes.
- **Today**: `parse_smartart_drawing` collects only `is_ns(*n, "sp", DSP_NS)`
  (`src/docx/smartart.rs:122`). Any `dsp:cxnSp` (connector) sibling in the
  `spTree` is filtered out → connecting arrows/lines between boxes are missing.
- **Corpus**: NOT present in the 2 fixtures (both use plain `dsp:sp` only — 0
  `cxnSp` in drawing1.xml), so no current regression; but cycle/process/hierarchy
  layouts that draw explicit connectors would lose them.
- **Caveat / verify first**: confirm whether the MS drawing part actually emits
  connectors as `dsp:cxnSp` vs. as ordinary `dsp:sp` with a connector preset
  geometry before implementing — if the latter, they already flow through and this
  is a non-issue. Inspect an org-chart/process SmartArt's `drawingN.xml`.
- **Likely owner**: `src/docx/smartart.rs` (extend the spTree filter), reusing
  `parse_dsp_shape` (geometry engine already handles connector preset shapes).

### LOW priority

- **Line dash / cap / join / compound on `dsp:sp` (`a:prstDash`, `a:custDash`,
  `@cap`, `a:round`/`a:bevel`)** — NOT parsed. `parse_dsp_shape` reads only solid
  line color + width (`src/docx/smartart.rs:170-183`); dashed/dotted borders on
  diagram shapes render solid. Owner: `src/docx/smartart.rs` + `src/pdf/smartart.rs`.
- **Line gradient fill (`a:gradFill` inside `dsp:sp/a:ln`)** — NOT parsed; line
  color is solid-fill only. Same owner as above.
- **Text `a:bodyPr @vert` (vertical / rotated text) and `@anchorCtr`** — NOT
  honored. `parse_dsp_text` (`src/docx/smartart.rs:254`) reads only insets +
  vertical `anchor` (t/ctr/b); `vert="vert"/"vert270"/"wordArtVert"` and
  horizontal centering `anchorCtr` are ignored. LATENT — corpus uses
  `vert="horz"` everywhere. Owner: `src/docx/smartart.rs` + `src/pdf/smartart.rs`.
- **Per-run baseline/super-subscript beyond raw `@baseline`** — `parse_dsp_text`
  stores `baseline` (`:335`) and `render_smartart` applies `baseline_y_offset`,
  but font-size scaling for super/subscript text in diagrams is not applied (Word
  shrinks raised/lowered glyphs). Minor. Owner: `src/pdf/smartart.rs`.

### Record-only (correct as-is while the `dsp:` cache exists)

- **`dgm:colorsDef` / `dgm:styleDef` / `dgm:layoutDef` parts never read** — by
  design. Their resolved output (colors, quick-style, computed positions) is
  already baked into the `dsp:` drawing cache that docxside reads, so re-parsing
  them would be redundant. These become relevant ONLY if the HIGH-priority
  data-model fallback (above) is built, at which point a fallback renderer would
  need at least `colorsDef` (for node fill colors) and `styleDef`. Owner then:
  `src/docx/smartart.rs`.
- **`dgm:dataModel` `bg`/`whole` (background + container formatting), `pt @cxnId`,
  `cxn @srcOrd`/`@destOrd`/`@parTransId`/`@sibTransId`, `prSet` presentation
  layout hints, `extLst`** — all data-model internals that the layout algorithm
  consumes to *produce* the drawing; irrelevant while reading the resolved cache.
  Would matter only inside a full layout-engine implementation (not recommended).

---

## 22.1 Math (OMML) — gaps

All math handling lives in `src/docx/runs.rs::omath_to_runs` (`:787-903`), a
**linearizer**: it walks the `m:oMath` tree and emits a flat sequence of ordinary
text `Run`s (each marked `is_math=true`), reusing the normal font + sub/superscript
machinery. There is no 2-D box model, so every gap below is a facet of "the linear
approximation drops/misorders something the spec lays out two-dimensionally." NOTE:
the current corpus has **0 fixtures containing `m:oMath`** (`analyze-fixtures --grep
"oMath"` → 0; the tool unzips and scans `document.xml` where OMML lives, so the
count is reliable) — so none of these gaps is exercised or regression-guarded by a
committed test today. A fresh math fixture (handcrafted `caseN` via python-docx, or
a scraped sample) should accompany any fix.

### HIGH priority — cheap fixes inside the existing linearizer (wrong/missing symbols)

#### 22.1.2.86/.26/.27 `m:rad` degree (nth root) — DROPPED
- **Spec**: §22.1.2.86 `rad` (Radical) has children `m:deg` (§22.1.2.26, the
  degree/index) and `m:e` (radicand), plus `m:radPr/m:degHide` (§22.1.2.27,
  hide-degree flag → square root). A cube root is `deg=3`, `√` with a small 3.
- **Today**: `omath_to_runs` `:867-874` emits a literal `√` then `(e)` and **never
  reads `m:deg`** — so `∛x`, `ⁿ√x`, etc. all render as `√(x)`. `degHide` is also
  ignored (square roots happen to look right only because the degree is dropped
  anyway).
- **Fix**: in the `"rad"` arm, if `m:deg` is present and non-empty (and `degHide`
  not set), emit the degree as a superscript-style prefix before `√` (e.g.
  `³√(x)`), else just `√(x)`. Owner: `src/docx/runs.rs:867`.

#### 22.1.2.24 `m:d` delimiter — opening/closing/separator chars HARDCODED to `()`
- **Spec**: §22.1.2.24 `d` (Delimiter) reads `m:dPr/m:begChr` (§22.1.2.10),
  `m:endChr` (§22.1.2.30), `m:sepChr` (§22.1.2.99). Brackets `[x]`, braces `{x}`,
  absolute value `|x|`, floor/ceiling, and angle brackets all live here, as do
  multi-element delimiters like `(x, y)` where `sepChr` (default `,` or `|`)
  separates the `m:e` children. Defaults: begChr `(`, endChr `)`, sepChr `,`.
- **Today**: the `"d"` arm (`runs.rs:875-881`) emits literal `(` … `)` and processes
  only the FIRST `m:e` child — so `[x]`/`{x}`/`|x|` all become `(x)` and any
  multi-element delimiter loses its 2nd+ operands and separators entirely.
- **Fix**: read `begChr`/`endChr`/`sepChr` from `m:dPr` (honor empty-string =
  "no delimiter"); iterate ALL `m:e` children, joining with `sepChr`. Owner:
  `src/docx/runs.rs:875`.

#### 22.1.2.1/.9/.40 accent / bar / group-character marks — DROPPED
- **Spec**: `m:acc` (§22.1.2.1 Accent) carries the combining mark in
  `m:accPr/m:chr` (§22.1.2.16, default U+0302 circumflex) over its `m:e` — e.g.
  x̄, x⃗, x̂. `m:bar` (§22.1.2.9) draws a bar above/below `m:e` (`m:barPr/m:pos`
  top|bot). `m:groupChr` (§22.1.2.40) places a grouping char (`m:groupChrPr/m:chr`,
  e.g. overbrace ⏞ / underbrace ⏟) above or below `m:e` (`m:pos`).
- **Today**: all three fall through to the catch-all `_` arm (`runs.rs:900`), which
  recurses into `m:e` and emits only the base text — the accent/bar/brace mark is
  silently lost (e.g. `x̄` → `x`, an overbraced group → the bare group).
- **Fix**: add explicit arms. For `acc`: emit `e` then append the `chr` combining
  mark (or a spacing approximation). For `bar`/`groupChr`: emit the char before/
  after `e` per `pos` (a full implementation overlays it, but even appending the
  literal preserves the symbol). Owner: `src/docx/runs.rs` (new match arms).

#### 22.1.2.88 `m:rPr` — math italic / style / literal IGNORED (variables render upright)
- **Spec**: §22.1.2.88 `m:rPr` (math run properties) holds `m:sty` (§22.1.2.114,
  style = `p` plain | `b` bold | `i` italic | `bi` bold-italic), `m:nor`
  (§22.1.2.76, force normal/upright text), `m:scr` (§22.1.2.92, math alphabet:
  double-struck/fraktur/script/sans-serif/monospace), `m:lit` (§22.1.2.49, literal
  — suppress auto-italic), `m:brk` (§22.1.2.13, manual line break), `m:aln`
  (§22.1.2.4). Word's default for a single-letter math variable is **italic**
  (`sty="i"` implied); `nor`/`lit` force upright (used for multi-letter function
  names, units).
- **Today**: `omath_to_runs` `:815-816` reads only the run's **`w:rPr`** (WML font/
  size) and explicitly ignores `m:rPr` ("m:rPr math styling is ignored"). A math
  run that carries no explicit `w:i` therefore renders upright, whereas Word
  italicizes variables — and `m:scr` alphabet variants (ℝ, 𝔄, etc.) are never
  applied.
- **Fix**: read `m:rPr`; map `sty`/`nor`/`lit` to `Run.italic` (italic by default
  for letters, upright when `nor`/`lit`/`sty="p"`); optionally map `m:scr` to the
  corresponding Unicode math-alphanumeric block. Owner: `src/docx/runs.rs:815`
  (+ possibly `src/model/mod.rs` if a math-alphabet field is added).

#### 22.1.2.94 `m:sPre` (pre-sub/superscript) — scripts MISORDERED, rendered at baseline
- **Spec**: §22.1.2.94 `sPre` puts the script terms BEFORE the base (e.g. the
  leading subscript/superscript in ₙCₖ or ⁲³X): children order `m:sub`, `m:sup`,
  `m:e` and the sub/sup sit to the *left* of the base, lowered/raised.
- **Today**: `sPre` has no arm, so the catch-all recurses children in document
  order — it emits `sub`, `sup`, then `e` all at baseline (no VertAlign), so a
  pre-script `²X` renders as the flat text `2X` and the raise/lower is lost.
- **Fix**: add an `"sPre"` arm mirroring `sSubSup` but emitting `sub`
  (Subscript) + `sup` (Superscript) BEFORE `e`. Owner: `src/docx/runs.rs`.

### MEDIUM priority — need real 2-D layout to be correct (linearization is lossy)

#### 22.1.2.55 `m:m` matrix + 22.1.2.28 `m:eqArr` equation array — FLATTENED
- **Spec**: `m:m` (§22.1.2.55 Matrix) has `m:mPr` (baseJc, plcHide, `m:mcs/m:mc`
  column justification §22.1.2.54/.53, row/col spacing rSp/cSp/cGp) and `m:mr`
  (§22.1.2.75, matrix rows) each holding `m:e` cells; `m:eqArr` (§22.1.2.28
  Equation Array) stacks multiple `m:e` rows (used for aligned multi-line
  equations / cases). Both are inherently 2-D.
- **Today**: both fall through the catch-all and recurse — every cell/row's text is
  concatenated into ONE line with no row breaks, no column alignment, and (for a
  matrix) no surrounding brackets (those come from a wrapping `m:d`). A 2×2 matrix
  becomes "abcd"; a 3-line equation array becomes one run-on line.
- **Fix**: requires emitting per-row line breaks at minimum (and, for full
  fidelity, a small grid layout). Linear approximation could insert a newline
  between `m:mr`/`m:e` rows. Owner: `src/docx/runs.rs` (+ multi-line handling in
  `src/pdf/layout.rs` since a single math zone would then span multiple lines).

#### 22.1.2.50/.52 `m:limLow` / `m:limUpp` (under/over limits) — limit rendered inline at baseline
- **Spec**: `m:limLow` (§22.1.2.50) places `m:lim` (§22.1.2.48) BELOW `m:e`;
  `m:limUpp` (§22.1.2.52) places it ABOVE. Used for `lim` under arrows, `max`/`min`
  with conditions, etc. (e.g. lim with `x→0` underneath).
- **Today**: catch-all recursion emits `e` then `lim` both at baseline → `lim_{x→0}`
  prints as the flat run "limx→0".
- **Fix**: at minimum render the limit as sub/superscript (`limLow`→Subscript,
  `limUpp`→Superscript) to convey position. Owner: `src/docx/runs.rs`.

#### 22.1.2.32 `m:f` fraction-type & 22.1.2.70 `m:nary` limit-location IGNORED
- **Spec**: `m:fPr/m:type` (§22.1.2.117 ST_FType) = `bar` (stacked, default) |
  `skw` (skewed `a⁄b`) | `lin` (inline `a/b`) | `noBar` (stack with no rule, for
  binomials). `m:naryPr/m:limLoc` (§22.1.2.51) = `undOvr` (limits centered above/
  below the operator) vs `subSup` (limits to the right). Also `m:naryPr/m:subHide`
  (§22.1.2.112) / `m:supHide` (§22.1.2.113) hide a limit; `m:grow` (§22.1.2.39).
- **Today**: `m:f` always emits `num/den` regardless of type (correct only for
  `lin`); `m:nary` always treats sub/sup as sub/superscript (the `subSup` layout)
  and ignores `limLoc=undOvr`, `subHide`/`supHide`, `grow`.
- **Fix**: for `noBar` binomials suppress the `/`; honor `subHide`/`supHide` (skip
  the hidden limit). The `undOvr` vs `subSup` distinction needs 2-D layout to be
  truly correct but sub/superscript is an acceptable linear fallback. Owner:
  `src/docx/runs.rs:858` (`f`) and `:882` (`nary`).

### LOW priority — latent / niche

- **22.1.2.61 `m:mathPr` (document math properties)** — NOT parsed. Lives in
  `word/settings.xml`; children incl. `m:mathFont` (§22.1.2.58, the math typeface,
  normally Cambria Math), `m:defJc` (default display-math justification — docxside
  hardcodes Center, which matches the `centerGroup` default), `m:dispDef`
  (§22.1.2.25), `m:brkBin`/`m:brkBinSub` (binary-operator break placement),
  `m:smallFrac` (§22.1.2.108), `m:intLim`/`m:naryLim` (default limit location for
  integrals / other n-ary ops), `m:lMargin`/`m:rMargin`/`m:wrapIndent`. Owner:
  `src/docx/settings.rs` (parse) — `mathFont` would feed the font chosen for
  `is_math` runs; `intLim`/`naryLim` would set the `m:nary` limit-location default.
  Latent: needs a fixture that sets a non-Cambria-Math math font or non-default
  break/limit behavior to show any divergence.
- **22.1.2.11/.12/.81 `m:borderBox` / `m:box` / `m:phant` (decoration/grouping)** —
  recursed via catch-all, so their content text prints but: `m:borderBox`
  (§22.1.2.11) draws no border (and ignores `borderBoxPr` strike-through flags
  strikeH/strikeV/strikeBLTR/strikeTLBR); `m:phant` (§22.1.2.81 Phantom) should be
  *invisible* spacing (or zero-width/asc/desc per `m:phantPr`) but renders its
  content as visible text — wrong when used to reserve space; `m:box` (§22.1.2.12)
  is pure grouping so recursion is fine. Owner: `src/docx/runs.rs`.
- **22.1.2.96 `m:sSubSup/m:sSubSupPr/m:alnScr` (align scripts)** — the align-scripts
  flag is ignored (scripts are emitted as plain sub/superscript); negligible in a
  linear renderer. Owner: `src/docx/runs.rs:847`.
- **22.1.2.23 `m:ctrlPr` (control properties)** — the per-object run-property
  wrapper (carries the `w:rPr` for an object's normal text, e.g. the fraction bar's
  formatting). Only `m:r` reads `w:rPr` today; object-level `ctrlPr` is unread.
  Mostly affects color/size of generated glyphs (bars, parens). Owner:
  `src/docx/runs.rs`.
- **22.1.2.13 `m:brk` (manual break) / 22.1.2.4 `m:aln` (alignment point)** — math
  line-break and alignment hints inside `m:rPr`; ignored. Only relevant once
  multi-line math (eqArr/matrix) layout exists. Owner: `src/docx/runs.rs` +
  `src/pdf/layout.rs`.

---

## 22.6 Bibliography — gaps

The Bibliography part itself has no rendering gap (it is data, not content — see
the "Chapters processed" note). There is exactly one print-relevant facet, and it
is LATENT in the current corpus.

### LOW priority — latent (uncached citation/bibliography field result)

#### §17.16.5.8 CITATION / BIBLIOGRAPHY field with NO cached result — would render nothing
- **Spec**: §22.6 (Bibliography part data) + §17.16.5.8 (CITATION field) and the
  BIBLIOGRAPHY field. A CITATION field displays the `b:Source` (§22.6.2.59) whose
  `b:Tag` (§22.6.2.65) matches its field-argument, formatted by the XSL named in
  the `b:Sources/@SelectedStyle` (§22.6.2.60); the BIBLIOGRAPHY field emits the
  full list. A conformant producer caches the formatted result as ordinary
  `w:p`/`w:t` runs between the field's `fldChar separate` and `end`.
- **Today**: docxside renders the cached result region of any non-dynamic field
  (CITATION/BIBLIOGRAPHY are not in `is_dynamic_field`, `src/docx/runs.rs:22`), so
  documents whose citation/bibliography fields carry a cached result render
  correctly (verified: `tests/fixtures/new/covid_insomnia_mental_health`'s
  BIBLIOGRAPHY field prints its ~70 cached reference paragraphs). The
  Bibliography part (`customXml/itemN.xml`, `b:Sources`/`b:Source`) is **never
  opened** — `grep -rniE "bibliograph|citation|b:Source|SelectedStyle" src/` → 0.
- **What exists vs missing**: cached-result rendering works (owned by Clause
  17.16). What is missing is *computing* a citation/bibliography from the
  `b:Source` data when the field has NO cached result — i.e. a field marked
  `w:dirty="true"` or one that was inserted but never updated by Word. In that
  case Word recomputes from `b:Sources` + the `@SelectedStyle` XSL on open/print,
  whereas docxside (which cannot run the XSL) would emit nothing for that field.
- **Corpus**: 0 fixtures exercise this — the only fixture with a real
  bibliography field (`covid_insomnia_mental_health`) ships a cached result, and
  the only other fixture with sources (`japanese_land_development_sign_form`, 1
  source) has no citing field at all. So output is correct today purely because no
  fixture relies on field recomputation.
- **Likely owner**: this is a large feature, not a quick fix — it requires (a)
  parsing the Bibliography part `customXml/itemN.xml` into a source model (new
  parser, likely `src/docx/bibliography.rs` + a model struct), (b) parsing
  CITATION/BIBLIOGRAPHY field-arguments and switches (§17.16.5.8: `\l \f \s \p \v
  \n \t \y \m`) in `src/docx/runs.rs`, and (c) a citation-style formatter
  (effectively a built-in renderer for the common APA/MLA/Chicago styles, since
  the actual `@SelectedStyle` XSL is not shippable). Only worth doing if dirty/
  never-updated citation fields show up in real corpus documents; for the normal
  case (cached results present) the existing field path is sufficient and correct.
