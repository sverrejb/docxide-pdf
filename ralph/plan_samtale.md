# Plan: Improve samtale fixture (samples/samtale)

## Context

The samtale fixture scores **Jaccard 9.2%, SSIM 43.6%** (both failing; thresholds are 20%/75%). It's an A4 landscape, two-column document with:
- Page 1: Law text (left col) + evaluation form header (right col)
- Page 2: 13 numbered Q&A items with gray bottom borders, split across two columns

The primary visual issue is **wrong column content distribution on page 2** — "Vurdering" heading and questions 6-13 end up in different positions than Word's reference because paragraph heights are wrong. A secondary issue is a missing Wingdings smiley character.

## Two Fixes

### Fix 1: Numbering Label Font Properties (HIGH IMPACT)

The DOCX numbering definition for questions has `w:sz="40"` (20pt), `w:b` (bold), `w:color="A6A6A6"` on the label `w:rPr`. Currently none of these are parsed — the label renders at the paragraph's text size (10pt) and the line height ignores the label entirely. With 13 questions, the compounded height error shifts column break points.

**Files to modify:**

#### A. `src/docx/numbering.rs` — Parse label rPr properties ✅ COMPLETED

1. Add fields to `LevelDef` (line 6-13):
   - `font_size: Option<f32>` (points, from `w:sz` half-points)
   - `bold: bool` (from `w:b`)
   - `color: Option<[u8; 3]>` (from `w:color @val`)

2. Parse from `w:lvl/w:rPr` (after line 67, where `bullet_font` is parsed):
   ```rust
   let rpr = wml(lvl, "rPr");
   let bullet_font = rpr.and_then(|r| wml(r, "rFonts")).and_then(|rf| ...);
   let label_font_size = rpr
       .and_then(|r| wml(r, "sz"))
       .and_then(|n| n.attribute((WML_NS, "val")))
       .and_then(|v| v.parse::<f32>().ok())
       .map(|hp| hp / 2.0);
   let label_bold = rpr.and_then(|r| wml_bool(r, "b")).unwrap_or(false);
   let label_color = rpr.and_then(|r| wml(r, "color"))
       .and_then(|n| n.attribute((WML_NS, "val")))
       .and_then(parse_hex_color);
   ```
   Add `wml_bool` and `parse_hex_color` to the imports from `super::`.

3. Change `parse_list_info` return type (line 215) from `(f32, f32, String, Option<String>)` to a struct:
   ```rust
   pub(super) struct ListLabelInfo {
       pub(super) indent_left: f32,
       pub(super) indent_hanging: f32,
       pub(super) label: String,
       pub(super) font: Option<String>,
       pub(super) font_size: Option<f32>,
       pub(super) bold: bool,
       pub(super) color: Option<[u8; 3]>,
   }
   ```
   With a `Default` impl (all zeros/None/empty/false). Update early returns (lines 225, 228, 231, 234, 237, 240) to `ListLabelInfo::default()`. Update final return (line 297) to construct the struct, pulling `font_size`, `bold`, `color` from `def`.

#### B. `src/model.rs` — Add label properties to Paragraph ✅ COMPLETED

Add after `list_label_font` (line 341):
```rust
pub list_label_font_size: Option<f32>,
pub list_label_bold: bool,
pub list_label_color: Option<[u8; 3]>,
```
Add to `Default` impl (after line 376): `list_label_font_size: None, list_label_bold: false, list_label_color: None,`

#### C. `src/docx/mod.rs` — Pass label properties to Paragraph (line 498-601) ✅ COMPLETED

Destructure `parse_list_info` result as `ListLabelInfo` struct, then add the three new fields to the `Paragraph` construction at line 589-609.

#### D. `src/docx/tables.rs` — Same pattern (line 253-286) ✅ COMPLETED

Destructure, add three fields to `Paragraph` at line 272.

#### E. `src/docx/textbox.rs` — Same pattern (line 51-101) ✅ COMPLETED

Destructure, add three fields to `Paragraph` at line 84.

#### F. `src/pdf/mod.rs` — Font registration for bold label (line 765-777) ✅ COMPLETED

When building the font key for label chars, include `/B` suffix if `list_label_bold`:
```rust
if !para.list_label.is_empty() {
    let key = if let Some(ref bf) = para.list_label_font {
        let mut k = bf.clone();
        if para.list_label_bold { k.push_str("/B"); }
        k
    } else if let Some(run) = para.runs.first() {
        // Override bold with label_bold
        let mut tmp_run = Run { bold: para.list_label_bold || run.bold, ..run.clone() };
        font_key_buf(&tmp_run, &mut key_buf).to_string()
    } else { continue; };
    used_chars_per_font.entry(key).or_default().extend(para.list_label.chars());
}
```

#### G. `src/pdf/mod.rs` — Update `label_for_paragraph` (line 2453-2470) ✅ COMPLETED

Build font key with bold suffix when `para.list_label_bold`. Currently looks up by bare font name — needs to use bold key:
```rust
fn label_for_paragraph<'a>(para: &Paragraph, seen_fonts: &'a HashMap<String, FontEntry>) -> (&'a str, Vec<u8>) {
    // Build key considering label bold
    let key = if let Some(ref bf) = para.list_label_font {
        let mut k = bf.clone();
        if para.list_label_bold { k.push_str("/B"); }
        k
    } else if let Some(run) = para.runs.first() {
        let key_run = Run { bold: para.list_label_bold || run.bold, ..run.clone() };
        font_key(&key_run)
    } else { return ("", vec![]); };
    let Some(entry) = seen_fonts.get(&key) else { return ("", vec![]); };
    let bytes = match &entry.char_to_gid {
        Some(map) => encode_as_gids(&para.list_label, map),
        None => to_winansi_bytes(&para.list_label),
    };
    (entry.pdf_name.as_str(), bytes)
}
```

#### H. `src/pdf/mod.rs` — Label rendering (3 sites) ✅ COMPLETED

**Main rendering (line 2003-2022):**
- Line 2006-2012: Use `para.list_label_color` with fallback to first run color
- Line 2015: Use `para.list_label_font_size.unwrap_or(font_size)` for `set_font`
- Line 2019-2021: Reset color if label_color was set

**Split paragraph rendering (line 1693-1712):** Same three changes.

**Textbox rendering (line 258-273):** Same pattern.

#### I. `src/pdf/mod.rs` — First-line height for label font size (line 1572-1575) ✅ COMPLETED

This is the critical fix for column distribution. When the label has a larger font size, the first line should be taller:

```rust
} else {
    let min_lines = 1 + para.extra_line_breaks as usize;
    let num_lines = lines.len().max(min_lines);
    // If label has a larger font size, first line is taller
    let first_line_h = if let Some(label_fs) = para.list_label_font_size {
        if label_fs > font_size {
            resolve_line_h(effective_ls, label_fs, tallest_lhr)
        } else { line_h }
    } else { line_h };
    if num_lines <= 1 {
        first_line_h
    } else {
        first_line_h + (num_lines - 1) as f32 * line_h
    }
};
```

#### J. `src/pdf/header_footer.rs` — Label rendering (line 285-304) ✅ COMPLETED

Same color/size changes as main rendering.

---

### Fix 2: `w:sym` Element (MEDIUM IMPACT) ✅ COMPLETED

Missing `<w:sym w:font="Wingdings" w:char="F04A"/>` (smiley in Q12).

**File to modify:** `src/docx/runs.rs`

Add match arm before `_ => {}` at line 604:
```rust
"sym" if !in_field => {
    // Flush pending text so the symbol gets its own Run with the symbol font
    if !pending_text.is_empty() {
        let run = fmt.text_run(std::mem::take(&mut pending_text), hyperlink_url.clone());
        runs.extend(split_run_by_script(run));
    }
    let sym_font = child.attribute((WML_NS, "font")).unwrap_or(&fmt.font_name);
    if let Some(ch) = child.attribute((WML_NS, "char"))
        .and_then(|hex| u32::from_str_radix(hex, 16).ok())
        .and_then(char::from_u32)
    {
        runs.push(Run {
            text: ch.to_string(),
            font_name: sym_font.to_string(),
            font_size: fmt.font_size,
            bold: fmt.bold,
            italic: fmt.italic,
            color: fmt.color,
            underline: fmt.underline,
            strikethrough: fmt.strikethrough,
            char_spacing: fmt.char_spacing,
            ..Run::default()
        });
    }
}
```

The PUA character (U+F04A) maps through Wingdings' cmap to the smiley glyph. The existing font pipeline handles per-run fonts and subsetting.

---

## Verification

1. `cargo build` — must compile cleanly
2. `DOCXIDE_CASE=samtale cargo test -- --nocapture` — check Jaccard/SSIM improvement
3. `cargo test -- --nocapture 2>&1 | grep "REGRESSION"` — verify zero regressions across all fixtures
4. Visual inspection of `tests/output/samples/samtale/diff/page_002.png` — "Vurdering" should now appear within the right column, not at the page bottom
