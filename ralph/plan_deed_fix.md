# Fix transition_to_work_deed Rendering (pages 1-10)

## Context

The scraped fixture `transition_to_work_deed` scores 6.3% Jaccard / 18.2% SSIM - critically failing. Comparing reference vs generated output (pages 1-10) reveals three main issues:

1. **Missing numbering** (~1700 paragraphs): sections 1., 1.1, 1.2, (a), (b) etc. are all missing labels
2. **TOC rendered blue with underlines**: reference shows black text, we show blue underlined text
3. **Cumulative y-drift**: text positions diverge increasingly across pages

The numbering issue alone affects ~1700 paragraphs and is the dominant cause of low scores. The y-drift will partially self-correct once numbering is fixed (labels change line wrapping).

---

## Fix 1: numStyleLink/styleLink resolution (CRITICAL)

**Files**: `src/docx/numbering.rs`

**Root cause**: The OOXML "linked numbering styles" mechanism is not implemented. The chain:
- StandardClause/StandardSubclause styles → numId 138 → abstractNum 30
- abstractNum 30 has `<w:numStyleLink w:val="Style10"/>` with NO level definitions
- abstractNum 81 has `<w:styleLink w:val="Style10"/>` WITH level definitions (decimal `%1.`, `%1.%2`, `(%3)`, etc.)
- Our parser finds no levels for abstractNum 30 → empty labels

**Changes**:

1. Add `#[derive(Clone)]` to `LevelDef` struct (line 5)

2. In `parse_numbering()`, during the `abstractNum` parsing loop (~line 82), record link attributes:
   ```rust
   let mut num_style_link: HashMap<String, String> = HashMap::new();  // absId → style name
   let mut style_link_target: HashMap<String, String> = HashMap::new(); // style name → absId

   // Inside "abstractNum" arm, after inserting levels:
   if let Some(link) = wml_attr(node, "numStyleLink") {
       num_style_link.insert(abs_id.to_string(), link.to_string());
   }
   if let Some(link) = wml_attr(node, "styleLink") {
       style_link_target.insert(link.to_string(), abs_id.to_string());
   }
   ```

3. After the main loop (after line 96), resolve links:
   ```rust
   for (abs_id, style_name) in &num_style_link {
       if let Some(target_abs_id) = style_link_target.get(style_name) {
           if let Some(source_levels) = abstract_nums.get(target_abs_id).cloned() {
               if !source_levels.is_empty() {
                   let entry = abstract_nums.entry(abs_id.clone()).or_default();
                   if entry.is_empty() {
                       *entry = source_levels;
                   }
               }
           }
       }
   }
   ```

---

## Fix 2: Suppress Hyperlink char style for anchor-only hyperlinks (TOC color)

**Files**: `src/docx/runs.rs`

**Root cause**: TOC entries use `<w:hyperlink w:anchor="...">` with `w:rStyle val="Hyperlink"`. The Hyperlink character style specifies blue color + underline. Word's PDF export suppresses these for internal links. Our code faithfully applies the style → blue underlined TOC entries.

**Changes**:

1. In `collect_run_nodes()`, change the output tuple to carry an `is_anchor_hyperlink` flag:
   - Type: `Vec<(Node, Option<String>, bool)>` (third element = anchor hyperlink flag)
   - Regular runs push `false`
   - For `w:hyperlink` nodes: check `w:anchor` attribute (WML namespace). If anchor is present and `r:id` is absent, push `true`

2. In the RunFormat loop (line 299), destructure the flag:
   ```rust
   for (run_node, hyperlink_url, is_anchor_hyperlink) in run_nodes {
   ```

3. When resolving `char_style` (line 303), suppress it for anchor hyperlinks:
   ```rust
   let char_style = if is_anchor_hyperlink {
       None
   } else {
       char_style_id_str.and_then(|id| styles.character_styles.get(id))
   };
   ```

---

## Not fixing now

- **"Reader's Guide" page position** (page 1 vs 2): Likely resolves itself once spacing changes from the fixes above. Re-check after implementation.
- **lvlOverride / startOverride**: Affects counter restart per-section. Follow-up work.
- **y-drift**: Multiple causes. Will improve significantly once numbering is fixed (1700 paragraphs get labels → different line wrapping → different spacing).

---

## Verification

1. `DOCXIDE_CASE=transition_to_work_deed cargo test -- --nocapture` — check improved Jaccard/SSIM
2. `cargo test -- --nocapture` — full suite, check for "REGRESSION in:" lines
3. Visually compare pages 3-4 (TOC should be black), pages 7-10 (numbering should appear)
4. `./tools/target/debug/analyze-fixtures` — overview of all scores
