use std::collections::HashMap;

use crate::fonts::{FontEntry, font_key_buf};

use crate::model::{
    Alignment, Block, CellMargins, HorizontalPosition, Table, TextDirection, VMerge,
    VerticalPosition, WrapType,
};

use super::RenderContext;
use super::header_footer::substitute_hf_runs;
use super::helpers::{effective_space_after, effective_space_before};
use super::layout::{
    TextLine, build_paragraph_lines, build_tabbed_line, east_asian_leading, is_text_empty,
    run_line_metrics, word_width_for_run,
};
use super::resolve_line_h;
use super::table_columns::{apply_nested_pct_width, cell_span_width};
pub(super) use super::table_columns::{apply_pct_width, auto_fit_columns, fixed_columns};

/// Height of a paragraph's text content, matching the layout computation in
/// `compute_row_layouts`. Empty paragraphs (no lines) still occupy one line
/// height unless they carry an explicit `content_height` (e.g. from an image).
pub(super) fn para_block_height(p: &CellParagraphLayout) -> f32 {
    if p.lines.is_empty() {
        if p.paragraph_mark_vanish {
            0.0
        } else {
            let h = if p.content_height > 0.0 {
                p.content_height
            } else {
                p.line_h
            };
            h + p.bottom_band.1
        }
    } else {
        cell_lines_h(p, 0..p.lines.len())
    }
}

/// Height of a cell paragraph's lines `range`, each at its own pitch
/// (`size_lines_by_own_runs`).
pub(super) fn cell_lines_h(p: &CellParagraphLayout, range: std::ops::Range<usize>) -> f32 {
    let bottom_band = if range.end == p.lines.len() {
        p.bottom_band.1
    } else {
        0.0
    };
    p.lines[range]
        .iter()
        .map(|l| super::layout::inline_line_advance(l, p.line_h, p.picture_metrics()))
        .sum::<f32>()
        + bottom_band
}

pub(super) struct CellFloatingImageLayout {
    pub(super) pdf_name: String,
    pub(super) display_width: f32,
    pub(super) display_height: f32,
    pub(super) h_offset: f32,
    pub(super) v_offset: f32,
    /// In-plane rotation in degrees (OOXML clockwise). Cell-anchored floats must
    /// carry this just like body floats, else e.g. a 90°-rotated vertical label
    /// renders horizontally.
    pub(super) rotation_deg: f32,
    /// wp:anchor relativeHeight, so a cell picture can paint over a
    /// connector/textbox anchored in the same paragraph (annotation #241).
    pub(super) z_index: u32,
    pub(super) behind_doc: bool,
    /// The picture's `docPr@descr` and decorative flag, for tagging.
    pub(super) alt: Option<String>,
    pub(super) decorative: bool,
}

impl CellFloatingImageLayout {
    /// (alt, decorative), for `cell_figure`.
    pub(super) fn tagging(&self) -> (Option<&str>, bool) {
        (self.alt.as_deref(), self.decorative)
    }
}

#[derive(Default)]
pub(super) struct CellParagraphLayout {
    pub(super) lines: Vec<TextLine>,
    pub(super) line_h: f32,
    /// Draws the baselines on Word's grid; its page height is the drawing
    /// page's.
    pub(super) line_grid: Option<super::layout::LineGrid>,
    pub(super) font_size: f32,
    pub(super) ascender_ratio: f32,
    /// Picture-line ascent; exact spacing keeps the fixed baseline and pitch.
    pub(super) image_line_ascent: f32,
    pub(super) clip_inline_pictures: bool,
    pub(super) descender_ratio: f32,
    /// What a picture line hangs below its picture (`picture_line_bottom`).
    pub(super) picture_descent: f32,
    pub(super) font_substituted: bool,
    pub(super) alignment: Alignment,
    pub(super) space_before: f32,
    /// Collapsed spacing less the previous page's already consumed after gap.
    pub(super) continuation_space_before: f32,
    pub(super) indent_left: f32,
    pub(super) indent_right: f32,
    pub(super) indent_hanging: f32,
    pub(super) indent_first_line: f32,
    pub(super) text_hanging: f32,
    /// `list_label::label_shift`: where a right/centre-aligned label starts.
    pub(super) label_shift: f32,
    /// Extra left indent from wrapSquare/Tight floating images in this paragraph.
    /// Text lines are laid out narrower and rendered further right to avoid the image.
    pub(super) float_indent_left: f32,
    /// The first `beside_lines` lines start `beside_indent` further right,
    /// level with a left floating nested table.
    pub(super) beside_indent: f32,
    pub(super) beside_lines: usize,
    /// `w:keepLines`: the paragraph never splits over a row break.
    pub(super) keep_lines: bool,
    /// Word 2013+ layout (compatibilityMode 15): a row break keeps two lines
    /// of this paragraph on each side; earlier layout lets a row start with one.
    pub(super) two_line_widow: bool,
    pub(super) list_label: String,
    /// The label by font (`list_label::label_pieces`), empty without one.
    pub(super) label_pieces: Vec<(Option<String>, String)>,
    pub(super) label_font_size: f32,
    /// For L/LI tagging inside the cell.
    pub(super) list_item: Option<crate::model::ListItem>,
    /// Tagged as the Formula itself (`sole_formula`).
    pub(super) sole_formula: bool,
    pub(super) label_color: Option<[u8; 3]>,
    pub(super) image_name: Option<String>,
    /// The picture's `docPr@descr` and decorative flag, for tagging.
    pub(super) image_alt: Option<String>,
    pub(super) image_decorative: bool,
    pub(super) image_width: f32,
    pub(super) image_trim: f32,
    pub(super) image_height: f32,
    pub(super) image_stroke_color: Option<[u8; 3]>,
    pub(super) image_stroke_width: f32,
    pub(super) image_shadow: Option<crate::model::ImageShadow>,
    pub(super) image_glow: Option<crate::model::ImageGlow>,
    pub(super) image_fx: Option<super::images::EffectXObjs>,
    pub(super) image_clip: Option<crate::model::ShapeGeometry>,
    pub(super) content_height: f32,
    pub(super) bottom_border: Option<crate::model::ParagraphBorder>,
    /// The gap above `bottom_border` and the whole band below the lines
    /// (`helpers::between_band` inside a border group).
    pub(super) bottom_band: (f32, f32),
    pub(super) paragraph_mark_vanish: bool,
    pub(super) floating_images: Vec<CellFloatingImageLayout>,
    pub(super) space_after: f32,
    pub(super) has_textboxes: bool,
    pub(super) has_connectors: bool,
}

impl CellParagraphLayout {
    pub(super) fn tag_kind(&self) -> &'static str {
        if self.sole_formula { "Formula" } else { "P" }
    }

    /// (ascent, descent) that size this paragraph's picture lines.
    pub(super) fn picture_metrics(&self) -> (f32, f32) {
        (self.image_line_ascent, self.picture_descent)
    }
}

// Synthetic Word controls switch between 18.75 pt and 18.70 pt of free
// right strip; the transition is unchanged for 8–12 pt marks and no borders.
const MIN_NESTED_FLOAT_STRIP: f32 = 18.75;

pub(super) enum CellContentItem {
    Paragraph(CellParagraphLayout),
    /// A table inside the cell, laid out (and drawn) at `col_widths`. A row
    /// split breaks between its rows (the cursor's `line` counts nested rows)
    /// or inside one that can split (`CellCursor::nested`).
    NestedTable {
        col_widths: Vec<f32>,
        rows: Vec<RowLayout>,
        space_before: f32,
        floating_offset: Option<f32>,
        floating_x_offset: Option<f32>,
        /// Half the table's bottom outer border band, below its last row.
        bottom_band: f32,
    },
}

pub(super) struct CellLayout {
    pub(super) items: Vec<CellContentItem>,
    /// The cell's margins (its own or the table's), border bands included.
    pub(super) cm: CellMargins,
    pub(super) total_height: f32,
    /// The space after the last paragraph that `total_height` includes; the
    /// chunk of a split row that finishes the cell charges it too.
    pub(super) trailing_space_after: f32,
    pub(super) text_direction: TextDirection,
}

pub(super) struct RowLayout {
    pub(super) height: f32,
    pub(super) cells: Vec<CellLayout>,
    /// The room a row needs on a page to break across it: None for cantSplit
    /// and exact-height rows, which Word never breaks; an at-least row's
    /// trHeight with its cell margins (croatian_grant's 15pt rows of five
    /// bullets break at the foot of a page, stem_partnerships' 76.55pt row
    /// and traditional_skills' form rows move whole).
    pub(super) split_min: Option<f32>,
    /// The first cell holds a keepNext paragraph: the row breaks only
    /// together with the next (`table::render_table`'s rule; issue #145's
    /// nested panels).
    pub(super) keeps_next: bool,
}

/// The index after the run of rows from `start` that keep with their
/// successor: a chain that moves as one.
pub(super) fn keep_chain_end(rows: &[RowLayout], start: usize) -> usize {
    let mut end = start + 1;
    while end < rows.len() && rows[end - 1].keeps_next {
        end += 1;
    }
    end
}

/// When provided, field codes in header/footer table runs are substituted with
/// their resolved values before layout.
pub(super) struct HfSubstitution<'a> {
    pub(super) page_num: usize,
    pub(super) total_pages: usize,
    pub(super) section_pages: usize,
    pub(super) styleref_values: &'a HashMap<String, String>,
    pub(super) page_num_format: Option<&'a str>,
}

/// A one-row, text-anchored table floated at the cell's left margin with room
/// for text beside it: its vertical extent below its anchor.
fn registration_table_extent(table: &Table, cell_text_w: f32, ctx: &RenderContext) -> Option<f32> {
    let pos = table.position.as_ref()?;
    if table.rows.len() != 1
        || pos.v_anchor != "text"
        || pos.h_anchor != "margin"
        || table.table_indent != 0.0
        || table.alignment != crate::model::TableAlignment::Left
    {
        return None;
    }
    // A small offset from the margin keeps it a registration table
    // (nested-floats-spaced-text-nonzero-x's 1pt tblpX); the offset eats
    // into the strip beside it.
    let HorizontalPosition::Offset(x) = pos.h_position else {
        return None;
    };
    if x < 0.0 {
        return None;
    }
    let mut widths = auto_fit_columns(table, ctx.fonts, Some(cell_text_w), None, 0.0);
    apply_nested_pct_width(table, &mut widths, cell_text_w);
    let strip =
        cell_text_w - x - widths.iter().sum::<f32>() - pos.left_from_text - pos.right_from_text;
    if strip < MIN_NESTED_FLOAT_STRIP {
        return None;
    }
    let height = compute_row_layouts(table, &widths, ctx, None)
        .iter()
        .map(|r| r.height)
        .sum::<f32>();
    Some(height + pos.v_offset_pt.max(0.0) + pos.top_from_text + pos.bottom_from_text)
}

/// A paragraph whose height is only its lines and spacing.
fn plain_spacer(p: &crate::model::Paragraph) -> bool {
    !(p.frame_props.is_some()
        || p.paragraph_mark_vanish
        || p.image.is_some()
        || p.inline_chart.is_some()
        || !p.smartart.is_empty()
        || !p.floating_images.is_empty()
        || !p.textboxes.is_empty()
        || !p.connectors.is_empty()
        || p.contextual_spacing
        || p.space_before_auto
        || p.space_after_auto
        || p.runs.iter().any(|r| {
            r.is_tab
                || r.is_line_break
                || r.inline_image.is_some()
                || r.field_code.is_some()
                || r.math.is_some()
                || r.vanish
                || r.footnote_id.is_some()
                || r.endnote_id.is_some()
        }))
}

pub(super) fn spacer_line_h(p: &crate::model::Paragraph, ctx: &RenderContext) -> f32 {
    let (fs, lhr, _) = super::layout::tallest_run_metrics(&p.runs, ctx.fonts);
    resolve_line_h(p.line_spacing.unwrap_or(ctx.doc_line_spacing), fs, lhr)
}

// Interleaved empty anchors contribute flow height; their floating tables do
// not, provided that the accumulated spacer lines clear every table's extent.
fn interleaved_spacers_clear_nested_registration(
    cell: &crate::model::TableCell,
    block_idx: usize,
    cell_text_w: f32,
    ctx: &RenderContext,
    hf_sub: Option<&HfSubstitution>,
) -> bool {
    if hf_sub.is_some() || cell.text_direction != TextDirection::LrTb {
        return false;
    }
    let mut flow = 0.0f32;
    let mut occupied = 0.0f32;
    let mut prev_after = 0.0f32;
    let mut tables = 0;
    for (index, block) in cell.content.iter().enumerate() {
        match block {
            Block::Table(table) => {
                tables += 1;
                if tables > 2 || (tables == 2 && flow <= 0.0) {
                    return false;
                }
                let Some(extent) = registration_table_extent(table, cell_text_w, ctx) else {
                    return false;
                };
                occupied = occupied.max(flow + prev_after + extent);
            }
            Block::Paragraph(p) => {
                if !plain_spacer(p) {
                    return false;
                }
                flow += prev_after.max(p.space_before);
                if p.runs.iter().any(|r| !r.text.trim().is_empty()) {
                    return tables == 2 && block_idx < index && flow >= occupied;
                }
                flow += spacer_line_h(p, ctx);
                prev_after = p.space_after;
            }
        }
    }
    false
}

// Leading margin-anchored registration tables may precede visible cell text
// without contributing to flow height when measured empty paragraphs already
// clear their entire extent. Keep the old narrow path for all other layouts.
fn spacers_clear_nested_registration(
    cell: &crate::model::TableCell,
    block_idx: usize,
    cell_text_w: f32,
    ctx: &RenderContext,
    hf_sub: Option<&HfSubstitution>,
) -> bool {
    if hf_sub.is_some() || cell.text_direction != TextDirection::LrTb {
        return false;
    }
    let tables: Vec<&Table> = cell
        .content
        .iter()
        .map_while(|b| match b {
            Block::Table(t) => Some(t),
            Block::Paragraph(_) => None,
        })
        .collect();
    if tables.is_empty() || tables.len() > 2 || block_idx >= tables.len() {
        return false;
    }
    let mut occupied = 0.0f32;
    for table in &tables {
        let Some(extent) = registration_table_extent(table, cell_text_w, ctx) else {
            return false;
        };
        occupied += extent;
    }
    let mut empty_space = 0.0f32;
    let mut prev_after = 0.0f32;
    let mut visible_seen = false;
    for block in &cell.content[tables.len()..] {
        let Block::Paragraph(p) = block else {
            return false;
        };
        if !plain_spacer(p) {
            return false;
        }
        if visible_seen {
            continue;
        }
        empty_space += prev_after.max(p.space_before);
        if p.runs.iter().any(|r| !r.text.trim().is_empty()) {
            if empty_space < occupied {
                return false;
            }
            visible_seen = true;
        } else {
            empty_space += spacer_line_h(p, ctx);
            prev_after = p.space_after;
        }
    }
    visible_seen
}

/// The single-line width of a paragraph of plain visible text with its
/// indents; None for anything else or a missing font.
fn plain_line_width(
    p: &crate::model::Paragraph,
    fonts: &HashMap<String, FontEntry>,
) -> Option<f32> {
    if p.frame_props.is_some()
        || p.list_item.is_some()
        || p.image.is_some()
        || !p.floating_images.is_empty()
        || !p.textboxes.is_empty()
        || !p.runs.iter().any(|r| !r.text.trim().is_empty())
        || p.runs.iter().any(|r| {
            r.is_tab
                || r.is_line_break
                || r.inline_image.is_some()
                || r.field_code.is_some()
                || r.math.is_some()
                || r.vanish
        })
    {
        return None;
    }
    let mut width = p.indent_left + p.indent_right + p.indent_first_line.max(0.0);
    let mut key = String::new();
    for run in &p.runs {
        let font = fonts.get(font_key_buf(run, &mut key))?;
        let text = super::layout::caps_word(run, &run.text);
        width += super::layout::soft_split_width(&text, |t| {
            word_width_for_run(font, run, t, run.font_size)
        });
    }
    Some(width)
}

fn plain_text_fits_nested_strip(
    cell: &crate::model::TableCell,
    idx: usize,
    table: &Table,
    widths: &[f32],
    cell_text_w: f32,
    ctx: &RenderContext,
) -> bool {
    let Some(pos) = &table.position else {
        return false;
    };
    if !matches!(pos.h_anchor, "margin" | "column")
        || !matches!(pos.h_position, HorizontalPosition::Offset(x) if x == 0.0)
        || pos.v_offset_pt < 0.0
        || table.table_indent != 0.0
        || table.alignment != crate::model::TableAlignment::Left
    {
        return false;
    }
    let Some(Block::Paragraph(p)) = cell.content.get(idx + 1) else {
        return false;
    };
    if p.runs.iter().any(|r| r.caps) {
        return false;
    }
    let strip = cell_text_w - widths.iter().sum::<f32>() - pos.right_from_text;
    let Some(line_w) = plain_line_width(p, ctx.fonts) else {
        return false;
    };
    // A cell's sole table lets text of any length, over several paragraphs,
    // wrap beside it (advocacy's picture stories); beside another table only
    // a lone paragraph fitting one line is known to.
    let sole_table = cell
        .content
        .iter()
        .enumerate()
        .all(|(i, b)| i == idx || matches!(b, Block::Paragraph(_)));
    strip >= MIN_NESTED_FLOAT_STRIP
        && (sole_table || (cell.content.len() == idx + 2 && line_w <= strip))
}

// A single cell heading can precede positive-offset floating tables even
// when the tables are serialized before its anchor paragraph.
fn text_fits_above_nested_floats(
    cell: &crate::model::TableCell,
    cell_text_w: f32,
    fonts: &HashMap<String, FontEntry>,
    doc_line_spacing: crate::model::LineSpacing,
    hf_sub: Option<&HfSubstitution>,
) -> bool {
    if hf_sub.is_some()
        || cell.text_direction != TextDirection::LrTb
        || !matches!(cell.v_align, crate::model::CellVAlign::Top)
    {
        return false;
    }
    let Some((Block::Paragraph(p), prefix)) = cell.content.split_last() else {
        return false;
    };
    if prefix.is_empty() {
        return false;
    }
    let Some(width) = plain_line_width(p, fonts) else {
        return false;
    };
    let mut first_top = f32::INFINITY;
    for block in prefix {
        let Block::Table(t) = block else {
            return false;
        };
        let Some(pos) = &t.position else {
            return false;
        };
        if t.rows.len() != 1
            || pos.v_anchor != "text"
            || pos.h_anchor != "margin"
            || pos.v_offset_pt <= 0.0
            || t.table_indent != 0.0
            || !matches!(pos.h_position, HorizontalPosition::Offset(x) if x == 0.0)
            || t.alignment != crate::model::TableAlignment::Left
        {
            return false;
        }
        first_top = first_top.min(pos.v_offset_pt - pos.top_from_text);
    }
    let (fs, lhr, _) = super::layout::tallest_run_metrics(&p.runs, fonts);
    let height = p.space_before.max(0.0)
        + resolve_line_h(p.line_spacing.unwrap_or(doc_line_spacing), fs, lhr)
        + p.space_after.max(0.0);
    width <= cell_text_w && height <= first_top
}

pub(super) fn compute_row_layouts(
    table: &Table,
    col_widths: &[f32],
    ctx: &RenderContext,
    hf_sub: Option<&HfSubstitution>,
) -> Vec<RowLayout> {
    let cm = &table.cell_margins;
    let mut layouts: Vec<RowLayout> = table
        .rows
        .iter()
        .map(|row| {
            let mut max_h: f32 = 0.0;
            let row_keeps_next = row.cells.iter().flat_map(|cell| &cell.content)
                .any(|block| matches!(block, Block::Paragraph(p) if p.keep_next));
            // Word uses the largest top/bottom cell margins across a row,
            // even when the cell carrying the margin has only a hidden mark.
            // Left/right margins remain specific to each cell.
            let (row_top, row_bottom) = row.cells.iter().fold((0.0f32, 0.0f32), |(top, bottom), cell| {
                let margins = cell.cell_margins.as_ref().unwrap_or(cm);
                (top.max(margins.top - cell.borders.top.band() / 2.0),
                 bottom.max(margins.bottom - cell.borders.bottom.band() / 2.0))
            });
            // Every cell gets the row's widest band: master_thesis's arrow
            // cell, unbordered beside a boxed "yes/no", shares its baseline and
            // a two-line row grows by the box's 0.5pt (Word reference).
            let (top_band, bottom_band) = row_half_bands(row);
            let cells: Vec<CellLayout> = row
                .grid_cells()
                .map(|(grid_col, span, cell)| {
                    let span_w = cell_span_width(col_widths, grid_col, span);
                    // The resolved grid width is what the renderer draws borders
                    // and content at, so the layout must use the same width — a
                    // larger tcW preference otherwise wraps text past the drawn
                    // cell border (#115). Fixed layout too: `fixed_columns`
                    // already takes the cells' widths where Word does, and a
                    // dxa-wide table keeps its grid (a corpus redline's 2in
                    // columns under 4800-twip cells break 72pt "aft|er").
                    let col_w = span_w;
                    if cell.v_merge == VMerge::Continue {
                        return CellLayout {
                            items: vec![],
                            cm: *cm,
                            total_height: 14.4,
                            trailing_space_after: 0.0,
                            text_direction: TextDirection::LrTb,
                        };
                    }

                    let effective_margins = CellMargins {
                        top: row_top + top_band,
                        bottom: row_bottom + bottom_band,
                        ..*cell.cell_margins.as_ref().unwrap_or(cm)
                    };
                    let ecm = &effective_margins;
                    let is_rotated = cell.text_direction != TextDirection::LrTb;
                    let cell_text_w = if is_rotated {
                        10000.0
                    } else {
                        (col_w - ecm.left - ecm.right).max(0.0)
                    };
                    let mut total_h: f32 = ecm.top + ecm.bottom;
                    let mut floating_extent: f32 = 0.0;
                    let mut left_float_bottom: f32 = 0.0;
                    let mut left_float_top: f32 = 0.0;
                    let mut left_float_indent: f32 = 0.0;
                    let mut first_positive_float = false;
                    let mut left_float_right_gap: f32 = f32::INFINITY;
                    let mut max_rotated_line_w: f32 = 0.0;
                    let mut items: Vec<CellContentItem> = Vec::new();
                    let mut prev_space_after = 0.0f32;
                    let mut para_idx = 0usize;
                    let mut prev_was_nested_table = false;
                    let mut negative_anchor_shift = 0.0f32;
                    let mut final_float_mark_hidden = false;
                    // Bottom-aligned cells place text-anchored nested floats like
                    // centred ones (nested_table_bottom_positive).
                    let top_aligned = matches!(cell.v_align, crate::model::CellVAlign::Top);
                    let text_above_floats =
                        text_fits_above_nested_floats(cell, cell_text_w, ctx.fonts, ctx.doc_line_spacing, hf_sub);

                    let block_count = cell.content.len();
                    // Modern Word drops the relocated anchor line only when
                    // its empty paragraph has expanded automatic line spacing.
                    // Single/exact/at-least spacing and legacy modes retain it.
                    let suppress_negative_mark = |idx: usize, offset: Option<f32>| {
                        ctx.compat_mode >= 15 && offset.is_some_and(|y| y < 0.0)
                            && idx.checked_sub(1).and_then(|i| cell.content.get(i))
                                .is_some_and(|b| matches!(b, Block::Paragraph(p)
                                    if is_text_empty(&p.runs)
                                        && matches!(p.line_spacing.unwrap_or(ctx.doc_line_spacing),
                                            crate::model::LineSpacing::Auto(factor) if factor > 1.0)))
                    };
                    // Whether the end-of-cell mark after the nested table at `idx` is
                    // hidden: always after an inline one; after a float when the mark
                    // is the lone anchor of a wide one, when a positive offset
                    // suppresses the centered cell's trailing mark, when a negative
                    // offset displaces its anchor line, or when the anchor shifted.
                    let mark_after_nested_hidden = |idx: usize, offset: Option<f32>, final_mark_hidden: bool, anchor_shift: f32| {
                        final_mark_hidden || offset.is_none() || (!top_aligned && offset.is_some_and(|y| y > 0.0))
                            || suppress_negative_mark(idx, offset)
                            || (anchor_shift > 0.0 && offset.is_some())
                    };
                    for (block_idx, block) in cell.content.iter().enumerate() {
                        match block {
                            Block::Paragraph(para) => {
                                // An HTML rule (o:hr) alone in a hideMark cell keeps a
                                // 5.25pt line in Word (advocacy's story separators:
                                // 0.75pt margins, the 0.5pt rule 2.5pt below the top),
                                // where the hidden mark would otherwise collapse it.
                                // ponytail: one measurement; Word's rule for the
                                // shape's line box is unknown.
                                // Word 2013+ layout (compatibilityMode 15) keeps the
                                // mark's full text line instead: maine's separator
                                // rows step one line, 5.25pt once saved as mode 14
                                // (Word probes).
                                let rule_only = para.horizontal_rule.is_some() && is_text_empty(&para.runs);
                                let rule_keeps_line = rule_only && ctx.compat_mode >= 15;
                                let para_content_h = if cell.hide_mark && rule_only && !rule_keeps_line {
                                    5.25
                                } else {
                                    para.content_height
                                };
                                let nested_float_clearance = if left_float_right_gap < MIN_NESTED_FLOAT_STRIP
                                    && para.runs.iter().all(|r| r.text.trim().is_empty() && !r.is_tab && !r.is_line_break)
                                {
                                    let (fs,lhr,_) = super::layout::tallest_run_metrics(&para.runs,ctx.fonts);
                                    let pitch=resolve_line_h(para.line_spacing.unwrap_or(ctx.doc_line_spacing),fs,lhr);
                                    if !first_positive_float || total_h + pitch > left_float_top { (left_float_bottom-total_h).max(0.0) } else { 0.0 }
                                } else { 0.0 };
                                let substituted;
                                let runs = if let Some(sub) = hf_sub {
                                    substituted = substitute_hf_runs(&para.runs, sub);
                                    &substituted
                                } else if let Some(marked) = ctx.with_note_marks(&para.runs) {
                                    // A note reference mark's run is empty until it
                                    // shows its note's number, as in body paragraphs.
                                    substituted = marked;
                                    &substituted
                                } else {
                                    &para.runs
                                };
                                // A cell paragraph containing only ordinary
                                // padding spaces has the paragraph mark's line
                                // metrics in Word, not the spaces' run format.
                                let padding_mark;
                                let space_runs = runs;
                                let padding = para_content_h == 0.0
                                    && runs.iter().any(|run| !run.text.is_empty())
                                    && runs.iter().all(|run| !run.is_tab
                                        && !run.is_line_break && run.math.is_none()
                                        && run.inline_image.is_none()
                                        && run.text.bytes().all(|b| b == b' '));
                                let runs = if padding {
                                    padding_mark = crate::model::Run {
                                        font_size: para.paragraph_mark_font_size.unwrap_or(10.0),
                                        font_name: para.paragraph_mark_font_name.clone().unwrap_or_default(),
                                        ..Default::default()
                                    };
                                    std::slice::from_ref(&padding_mark)
                                } else { runs };
                                // Size the cell line from the tallest run, not the
                                // first — a small leading run (e.g. padding spaces)
                                // must not pull the baseline up. Math runs are
                                // clamped inside tallest_run_metrics so a header
                                // cell leading with math doesn't balloon the row.
                                // Cell metrics come from the first run carrying
                                // real text — a leading padding-spaces run must
                                // not set the baseline (Word sizes the line from
                                // the content run; observed in header cells like
                                // "                g" where g is larger).
                                let metric_run = runs
                                    .iter()
                                    .filter(|r| !r.vanish)
                                    .find(|r| {
                                        r.is_tab
                                            || r.text.is_empty()
                                            || !r.text.trim().is_empty()
                                    })
                                    .or(runs.first());
                                // A tab-only empty cell paragraph is sized by its
                                // paragraph mark, as in body text. Do not let an
                                // inherited tab font enlarge a smaller explicit mark.
                                // So is one of only hidden runs (`unsized_line_font`).
                                let mark_run;
                                let metric_run = if !runs.is_empty()
                                    && (runs.iter().all(|r| r.vanish)
                                        || runs.iter().all(|r| r.is_tab)
                                            && para.tab_stops.iter().all(|tab| tab.leader.is_none()))
                                    && let Some(size) = para.paragraph_mark_font_size
                                    && let Some(name) = &para.paragraph_mark_font_name
                                {
                                    mark_run = crate::model::Run {
                                        font_size: size,
                                        font_name: name.clone(),
                                        ..Default::default()
                                    };
                                    Some(&mark_run)
                                } else {
                                    metric_run
                                };
                                let font_size = metric_run.map_or(12.0, |r| r.font_size);
                                let mut kb0 = String::new();
                                let metric_font = metric_run
                                    .map(|r| font_key_buf(r, &mut kb0).to_owned())
                                    .and_then(|k| ctx.fonts.get(&k));
                                let (tallest_lhr, tallest_ar) = metric_run
                                    .zip(metric_font)
                                    .map_or((None, None), |(r, e)| run_line_metrics(e, &r.text));
                                let effective_ls =
                                    para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                                let line_h =
                                    resolve_line_h(effective_ls, font_size, tallest_lhr);
                                // A negative floating table moves its preceding empty
                                // anchor mark below itself. The mandatory end-of-cell
                                // mark does not replace that anchor's line metrics.
                                let relocated_anchor_line = (block_idx + 1 == block_count
                                    && nested_float_clearance > 0.0
                                    && is_text_empty(runs)
                                    && matches!(items.last(), Some(CellContentItem::NestedTable {
                                        floating_offset: Some(offset), ..
                                    }) if *offset < 0.0)
                                    && block_idx.checked_sub(2).and_then(|i| cell.content.get(i))
                                        .is_some_and(|b| matches!(b, Block::Paragraph(p)
                                            if is_text_empty(&p.runs) && p.content_height == 0.0
                                                && p.image.is_none() && !p.paragraph_mark_vanish)))
                                    .then(|| items.get(items.len().saturating_sub(2)))
                                    .flatten()
                                    .and_then(|item| match item {
                                        CellContentItem::Paragraph(p) => Some(p.line_h),
                                        _ => None,
                                    });
                                let line_h = relocated_anchor_line.unwrap_or(line_h);

                                // Auto-spaced cell lines snap to the section's line grid
                                // like body lines: japanese_medical's ten empty cell
                                // paragraphs step 18pt (the grid), not their 15pt
                                // natural height; chinese_student's at-least-0 cell
                                // lines keep their own height.
                                let grid_pitch = ctx.cell_grid_pitch.get();
                                let grid_snapped = para.snap_to_grid
                                    && grid_pitch > 0.0
                                    && matches!(effective_ls, crate::model::LineSpacing::Auto(_));
                                let line_h = if grid_snapped {
                                    super::layout::grid_snapped_line_h(
                                        runs,
                                        ctx.fonts,
                                        effective_ls,
                                        line_h,
                                        grid_pitch,
                                    )
                                } else if is_text_empty(runs) && para_content_h == 0.0 {
                                    line_h + super::mark_position_stretch(para, effective_ls, false, ctx.fonts)
                                } else {
                                    line_h
                                };

                                // A numbering label taller than the text raises the
                                // first line (see `label_boosted_line_h`); CV's 9pt
                                // Symbol bullets on 9pt Georgia add 0.8pt per item.
                                // ponytail: folded into space_before, which equals the
                                // baseline drop at single spacing; split them if a
                                // multiple-spaced list cell ever drifts.
                                let label_extra = super::label_boosted_line_h(
                                    para,
                                    ctx.fonts,
                                    line_h,
                                    effective_ls,
                                    font_size,
                                    tallest_lhr,
                                    tallest_ar,
                                ) - line_h;
                                // Contextual spacing applies between a cell's
                                // paragraphs as in the body (erasmus_plus: Normal's
                                // 12pt after "Faculty/Department" goes).
                                let space_after = effective_space_after(
                                    para,
                                    super::block_para(&cell.content, block_idx + 1),
                                );
                                // hideMark drops the cell's trailing empty mark with
                                // its spacing; the paragraph above keeps its space
                                // after (Word probes).
                                let mark_hidden = cell.hide_mark
                                    && block_idx == block_count - 1
                                    && para_content_h == 0.0
                                    && is_text_empty(runs)
                                    && !para.paragraph_mark_vanish
                                    && !rule_keeps_line;
                                // A cell drops HTML auto spacing at its edges.
                                let space_before = if mark_hidden {
                                    prev_space_after
                                } else if para_idx > 0 {
                                    let prev = super::block_para(&cell.content, block_idx - 1);
                                    f32::max(prev_space_after, effective_space_before(para, prev))
                                } else if para.space_before_auto {
                                    0.0
                                } else {
                                    para.space_before
                                } + label_extra;
                                let space_before = space_before.max(nested_float_clearance);
                                total_h += space_before;

                                // The first baseline sits this far (per em) below the
                                // cell top: the ascent, line gap included, as in body
                                // text; one em for an East Asian font, whose 1.3×
                                // leading Word keeps out of the cell's top
                                // (japanese_interlibrary: 11pt MS Mincho 11.0).
                                let east_asian = metric_run
                                    .zip(metric_font)
                                    .is_some_and(|(r, e)| east_asian_leading(e, &r.text));
                                let ascender_ratio =
                                    if east_asian { 1.0 } else { tallest_ar.unwrap_or(0.75) };
                                // Win-path metrics identity: line_h_ratio −
                                // ascender_ratio = usWinDescent/units. Fallback
                                // 0.2 pairs with the 1.2 default line ratio so
                                // standard fonts get zero trailing leading.
                                // A picture-only line keeps no descent unless a
                                // picture is underlined (`picture_line_bottom`).
                                let descender_ratio = if runs.iter().any(|r| r.inline_image.is_some())
                                    && runs.iter().all(|r| !r.is_tab && r.text.trim().is_empty())
                                    && !runs.iter().any(|r| !r.vanish && r.underlined_picture()) {
                                    0.0
                                } else {
                                    tallest_lhr.zip(tallest_ar)
                                        .map(|(lh, ar)| (lh - ar).max(0.0)).unwrap_or(0.2)
                                };
                                // Multiple spacing leads a picture line as in the body,
                                // glyphless or not (scottish_fundraising's header logo);
                                // an at-least minimum below the picture adds nothing
                                // (dutch_government's logo cell).
                                let multiple_spacing = matches!(effective_ls, crate::model::LineSpacing::Auto(_));
                                let picture_descent = if multiple_spacing {
                                    super::layout::picture_line_bottom(runs, para, ctx.fonts, effective_ls)
                                } else {
                                    font_size * descender_ratio
                                };
                                let exact_spacing =
                                    matches!(effective_ls, crate::model::LineSpacing::Exact(_));
                                let image_line_ascent = if exact_spacing {
                                    super::layout::runs_max_image_h(runs).max(font_size * ascender_ratio)
                                } else {
                                    font_size * ascender_ratio
                                };
                                let font_substituted =
                                    metric_font.is_some_and(|e| e.is_substituted);

                                // Compute extra left indent from left-aligned
                                // wrapSquare/Tight floating images so text wraps
                                // to the right of the image within the cell.
                                let float_indent_left: f32 = para
                                    .floating_images
                                    .iter()
                                    .filter(|fi| {
                                        matches!(
                                            fi.wrap_type,
                                            WrapType::Square | WrapType::Tight | WrapType::Through
                                        ) && matches!(
                                            fi.h_position,
                                            HorizontalPosition::AlignLeft
                                                | HorizontalPosition::Offset(_)
                                        )
                                    })
                                    .map(|fi| {
                                        let left_edge = match fi.h_position {
                                            HorizontalPosition::Offset(o) => o,
                                            _ => 0.0,
                                        };
                                        left_edge + fi.image.display_width + fi.dist_right
                                    })
                                    .fold(0.0f32, f32::max);
                                // Lines level with a left floating nested table
                                // keep clear of it, the rest run full width, as
                                // beside a wrapping picture.
                                let beside_float = total_h < left_float_bottom
                                    && total_h + line_h > left_float_top
                                    && left_float_right_gap >= MIN_NESTED_FLOAT_STRIP;
                                let mut beside_indent = if beside_float {
                                    (left_float_indent - float_indent_left).max(0.0)
                                } else {
                                    0.0
                                };
                                let mut beside_lines = if beside_indent > 0.0 {
                                    ((left_float_bottom - total_h) / line_h.max(1.0)).ceil().max(1.0) as usize
                                } else {
                                    0
                                };
                                // ponytail: tabbed lines take no per-line widths, so
                                // a tabbed paragraph stays indented throughout.
                                let float_indent_left = if beside_lines > 0 && runs.iter().any(|r| r.is_tab) {
                                    beside_lines = 0;
                                    float_indent_left + std::mem::take(&mut beside_indent)
                                } else {
                                    float_indent_left
                                };

                                // An inline-image paragraph whose only other run
                                // is a trailing line break (Word's logo-in-cell
                                // idiom) is NOT text-empty, but its row height must
                                // come from the image's content_height, not a stray
                                // text line — otherwise the cell collapses and a
                                // tall header logo fails to push the body down.
                                let image_only = para_content_h > 0.0
                                    && runs.iter().all(|r| r.text.is_empty() && !r.is_tab);
                                let mut break_mark_hidden = false;
                                let end_mark_after_table = prev_was_nested_table
                                    && block_idx == block_count - 1
                                    && para_content_h == 0.0
                                    && is_text_empty(runs);
                                // Match the rendering's first_line_hanging: when a
                                // list label is present, the label is drawn separately
                                // and the text starts at indent_left, so the first
                                // line has no extra hanging width.
                                let hanging = super::list_label::text_hanging(para, ctx.default_tab_stop, ctx.fonts);
                                let lines = if !is_text_empty(runs) && !image_only {
                                    let para_text_w = (cell_text_w
                                        - para.indent_left
                                        - para.indent_right
                                        - float_indent_left)
                                        .max(0.0);
                                    let has_tabs = runs.iter().any(|r| r.is_tab);
                                    let inline_images = super::layout::run_image_map(runs, ctx.table_cell_image_names);
                                    let inline_effects = super::layout::run_image_map(runs, ctx.effect_table_names);
                                    let mut lines = if has_tabs {
                                        build_tabbed_line(
                                            runs,
                                            ctx.fonts,
                                            &para.tab_stops,
                                            para.indent_left,
                                            para_text_w,
                                            para.indent_right,
                                            hanging,
                                            &inline_images,
                                            &inline_effects,
                                            ctx.default_tab_stop,
                                            &[],
                                            ctx.compat_mode,
                                            ctx.cjk(false, para),
                                        )
                                    } else {
                                        let beside_widths = vec![(para_text_w - beside_indent).max(1.0); beside_lines];
                                        build_paragraph_lines(
                                            runs,
                                            ctx.fonts,
                                            para_text_w,
                                            hanging,
                                            &inline_images,
                                            &inline_effects,
                                            None,
                                            (beside_lines > 0).then_some(beside_widths.as_slice()),
                                            None,
                                            ctx.cjk(para.auto_space_de || para.auto_space_dn, para),
                                        )
                                    };
                                    if is_rotated {
                                        for line in &lines {
                                            max_rotated_line_w =
                                                max_rotated_line_w.max(line.total_width);
                                        }
                                    }
                                    // Each line is as tall as its own runs, as in body
                                    // text: nabl's "(Mark √ in the" header line, √ a
                                    // w:sym Symbol run, steps 12.24 where Arial gives 11.50.
                                    if !east_asian && !grid_snapped && !matches!(effective_ls, crate::model::LineSpacing::Exact(_)) {
                                        super::layout::size_lines_by_own_runs(
                                            &mut lines,
                                            ctx.fonts,
                                            effective_ls,
                                            line_h,
                                            font_size * ascender_ratio,
                                            super::helpers::auto_ascent_scale(effective_ls),
                                            para.text_alignment == crate::model::TextAlignment::Top,
                                        );
                                        super::size_mark_line(&mut lines, para, runs, effective_ls, tallest_lhr, ctx.fonts);
                                        super::size_tab_mark_line(&mut lines, para, effective_ls, tallest_lhr, ctx.fonts);
                                    }
                                    // hideMark also hides the mark left alone on the
                                    // line after the cell's closing break, with its
                                    // spacing (maine's "…this Bureau.<w:br/>"; Word
                                    // probes: one or two breaks, 0 or 12pt after).
                                    if cell.hide_mark
                                        && block_idx == block_count - 1
                                        && lines.len() > 1
                                        && lines.last().is_some_and(TextLine::is_blank)
                                        && lines[lines.len() - 2].ends_with_break
                                    {
                                        lines.pop();
                                        break_mark_hidden = true;
                                    }
                                    total_h += lines
                                        .iter()
                                        .map(|l| super::layout::inline_line_advance(l, line_h, (image_line_ascent, picture_descent)))
                                        .sum::<f32>();
                                    lines
                                } else {
                                    let mut end_spaces = Vec::new();
                                    if para.paragraph_mark_vanish {
                                        // vanished paragraph mark: zero height
                                    } else if mark_hidden {
                                        // hideMark: last empty paragraph in cell
                                        // contributes no height
                                    } else if end_mark_after_table {
                                        // End-of-cell mark directly after a nested
                                        // table: Word hides it; space_after is
                                        // suppressed in the trailing-space block
                                        // below.
                                    } else if para_content_h > 0.0 {
                                        // Image paragraph: the image is line 1; each
                                        // trailing w:br adds a further blank line
                                        // (matches the non-table header height path).
                                        let br_count = runs
                                            .iter()
                                            .filter(|r| r.is_line_break)
                                            .count();
                                        total_h += para_content_h
                                            + if multiple_spacing { picture_descent } else { 0.0 }
                                            + br_count as f32 * line_h;
                                    } else {
                                        total_h += line_h;
                                        if padding && ctx.ul_trail_space {
                                            let width = (cell_text_w - para.indent_left - para.indent_right).max(0.0);
                                            let lines = build_paragraph_lines(
                                                space_runs,
                                                ctx.fonts,
                                                width,
                                                0.0,
                                                &HashMap::new(),
                                                &HashMap::new(),
                                                None,
                                                None,
                                                None,
                                                ctx.cjk(false, para),
                                            );
                                            if lines.iter().any(|l| !l.end_spaces.is_empty()) {
                                                end_spaces = lines;
                                            }
                                        }
                                    }
                                    end_spaces
                                };

                                let image_name = para.image.as_ref().and_then(|img| {
                                    ctx.table_cell_image_names.get(&img.id).cloned()
                                });
                                let (image_width, image_height, img_stroke_color, img_stroke_width, img_shadow) = para
                                    .image
                                    .as_ref()
                                    .map(|img| (img.display_width, img.display_height, img.stroke_color, img.stroke_width, img.shadow.clone()))
                                    .unwrap_or((0.0, 0.0, None, 0.0, None));
                                let table_fx = para.image.as_ref().and_then(|img| {
                                    ctx.effect_table_names.get(&img.id)
                                });
                                let img_glow = para.image.as_ref().and_then(|img| img.glow.clone());

                                let cell_floats: Vec<CellFloatingImageLayout> = para
                                    .floating_images
                                    .iter()
                                    .filter_map(|fi| {
                                        let pdf_name =
                                            ctx.table_cell_image_names.get(&fi.image.id)?.clone();
                                        let h_offset = match fi.h_position {
                                            HorizontalPosition::Offset(o) => o,
                                            HorizontalPosition::AlignCenter => {
                                                (col_w - fi.image.display_width) / 2.0
                                            }
                                            HorizontalPosition::AlignRight => {
                                                // Modern Word aligns to the cell's right text edge.
                                                // Legacy mode also shifts the origin by the left inset.
                                                let legacy_origin = if ctx.compat_mode < 15 {
                                                    ecm.left
                                                } else {
                                                    0.0
                                                };
                                                col_w + legacy_origin - ecm.right - fi.image.display_width
                                            }
                                            HorizontalPosition::AlignLeft => 0.0,
                                            p @ HorizontalPosition::InGroup { .. } => {
                                                p.place(0.0, col_w, fi.image.display_width)
                                            }
                                        };
                                        let v_offset = match fi.v_position {
                                            VerticalPosition::Offset(o) => o,
                                            _ => 0.0,
                                        };
                                        Some(CellFloatingImageLayout {
                                            pdf_name,
                                            display_width: fi.image.display_width,
                                            display_height: fi.image.display_height,
                                            h_offset,
                                            v_offset,
                                            rotation_deg: fi.image.rotation_deg,
                                            z_index: fi.z_index,
                                            behind_doc: fi.behind_doc,
                                            alt: fi.image.alt.clone(),
                                            decorative: fi.image.decorative,
                                        })
                                    })
                                    .collect();

                                let next = super::block_para(&cell.content, block_idx + 1)
                                    .filter(|n| super::helpers::joins_border_group(para, n));
                                let (bottom_border, bottom_band) = super::helpers::bottom_border(para, next);
                                total_h += bottom_band.1;

                                let first_ascent = grid_snapped
                                    .then(|| super::layout::grid_baseline_offset(runs, ctx.fonts, line_h))
                                    .flatten()
                                    .or_else(|| super::layout::boxed_line_ascent(
                                        effective_ls,
                                        line_h,
                                        font_size,
                                        tallest_lhr,
                                        tallest_ar,
                                        runs,
                                        ctx.fonts,
                                    ))
                                    .unwrap_or(
                                        font_size
                                            * ascender_ratio
                                            * super::helpers::auto_ascent_scale(effective_ls),
                                    );
                                let line_grid = super::layout::LineGrid {
                                    snapped: grid_snapped,
                                    ..super::layout::LineGrid::new(
                                        0.0,
                                        first_ascent,
                                        effective_ls,
                                        line_h,
                                        font_size * tallest_lhr.unwrap_or(1.2),
                                        image_line_ascent,
                                        space_after,
                                    )
                                };
                                items.push(CellContentItem::Paragraph(CellParagraphLayout {
                                    lines,
                                    line_h,
                                    line_grid: Some(line_grid),
                                    font_size,
                                    // Places the first baseline only, which rises
                                    // with a box shrunk below single spacing.
                                    // Exact and winning at-least lines place it as
                                    // in the body: carbon_farming's 8pt footer cells
                                    // sit on their descent under Normal's 13pt
                                    // minimum, arizona's exact 13.2pt lines 80% down.
                                    // Grid-snapped lines centre in their cells as
                                    // in the body (japanese_hr's 9pt lines sit
                                    // 2.9pt into a 17.5pt cell).
                                    ascender_ratio: first_ascent / font_size,
                                    descender_ratio,
                                    picture_descent,
                                    image_line_ascent,
                                    clip_inline_pictures: exact_spacing,
                                    font_substituted,
                                    alignment: para.alignment,
                                    space_before,
                                    continuation_space_before: super::own_space_before(space_before, prev_space_after),
                                    indent_left: para.indent_left,
                                    indent_right: para.indent_right,
                                    indent_hanging: para.indent_hanging,
                                    indent_first_line: para.indent_first_line,
                                    text_hanging: hanging,
                                    label_shift: super::list_label::label_shift(para, ctx.fonts),
                                    float_indent_left,
                                    beside_indent,
                                    beside_lines,
                                    // Word 2010 layout splits a keepLines cell paragraph too.
                                    keep_lines: para.keep_lines && ctx.compat_mode >= 15,
                                    // Without widow control (and nothing keeping the row
                                    // with the next) Word 2013+ also lets one line stand.
                                    two_line_widow: ctx.compat_mode >= 15 && (para.widow_control || row_keeps_next),
                                    list_label: para.list_label.clone(),
                                    sole_formula: super::sole_formula(para),
                                    label_pieces: if para.list_label.is_empty() {
                                        Vec::new()
                                    } else {
                                        super::list_label::label_pieces(para)
                                            .into_iter()
                                            .map(|(k, t)| (k, t.to_string()))
                                            .collect()
                                    },
                                    label_font_size: para.list_label_font_size.unwrap_or(font_size),
                                    list_item: para.list_item,
                                    label_color: para.list_label_color,
                                    image_name,
                                    image_alt: para.image.as_ref().and_then(|i| i.alt.clone()),
                                    image_decorative: para
                                        .image
                                        .as_ref()
                                        .is_some_and(|i| i.decorative),
                                    image_width,
                                    image_trim: para.image.as_ref().map_or(0.0, |i| i.trimmed_width),
                                    image_height,
                                    image_stroke_color: img_stroke_color,
                                    image_stroke_width: img_stroke_width,
                                    image_shadow: img_shadow,
                                    image_glow: img_glow,
                                    image_fx: table_fx.cloned(),
                                    image_clip: para.image.as_ref().and_then(|img| img.clip_geometry.clone()),
                                    content_height: para_content_h,
                                    bottom_border: bottom_border.cloned(),
                                    bottom_band,
                                    paragraph_mark_vanish: para.paragraph_mark_vanish
                                        || end_mark_after_table,
                                    floating_images: cell_floats,
                                    space_after,
                                    has_textboxes: !para.textboxes.is_empty(),
                                    has_connectors: !para.connectors.is_empty(),
                                }));

                                prev_space_after = if mark_hidden
                                    || break_mark_hidden
                                    || (para.space_after_auto && block_idx == block_count - 1)
                                {
                                    0.0
                                } else {
                                    space_after
                                };
                                prev_was_nested_table = false;
                                para_idx += 1;
                            }
                            Block::Table(nested_table) => {
                                // Adjacent text-relative floats share the flow anchor.
                                // A preceding negative float displaces its empty
                                // anchor below the table before the next float's
                                // own offset is applied (not just to the table edge).
                                if nested_table.position.as_ref().is_some_and(|p| p.v_anchor == "text")
                                    && left_float_right_gap < MIN_NESTED_FLOAT_STRIP
                                    // Modern expanded anchors are not relocated (Word 15).
                                    && !(ctx.compat_mode >= 15
                                        && cell.content.get(block_idx.saturating_sub(2))
                                            .is_some_and(|b| matches!(b, Block::Paragraph(p)
                                                if matches!(p.line_spacing.unwrap_or(ctx.doc_line_spacing),
                                                    crate::model::LineSpacing::Auto(factor) if factor > 1.0))))
                                    && matches!(items.last(), Some(CellContentItem::NestedTable {
                                        floating_offset: Some(offset), ..
                                    }) if *offset < 0.0)
                                    && let Some(CellContentItem::Paragraph(anchor)) = items.get(items.len().saturating_sub(2))
                                    && block_idx.checked_sub(2).and_then(|i| cell.content.get(i))
                                        .is_some_and(|b| matches!(b, Block::Paragraph(p) if is_text_empty(&p.runs)))
                                    && anchor.image_name.is_none() && anchor.floating_images.is_empty()
                                {
                                    negative_anchor_shift = (left_float_bottom + anchor.line_h - total_h).max(0.0);
                                }
                                // The widths the nested table is drawn at, so a split's
                                // line cursors index the lines it draws.
                                let mut nested_cw = auto_fit_columns(nested_table, ctx.fonts, Some(cell_text_w), None, 0.0);
                                apply_nested_pct_width(nested_table, &mut nested_cw, cell_text_w);
                                let nested_layouts =
                                    compute_row_layouts(nested_table, &nested_cw, ctx, hf_sub);
                                let nested_h = nested_layouts.iter().map(|rl| rl.height).sum::<f32>();
                                let plain_side_text = plain_text_fits_nested_strip(cell, block_idx, nested_table, &nested_cw, cell_text_w, ctx);
                                // The room a float leaves on its right.
                                let gap_beside = nested_table.position.as_ref().map_or(0.0, |pos| {
                                    let float_x = match pos.h_position { HorizontalPosition::Offset(x) => x, _ => 0.0 };
                                    cell_text_w - float_x - nested_cw.iter().sum::<f32>() - pos.right_from_text
                                });
                                let blank = |p: &crate::model::Paragraph| {
                                    p.runs.iter().all(|r| r.text.trim().is_empty()
                                        && !r.is_tab && !r.is_line_break && r.inline_image.is_none() && r.field_code.is_none())
                                        && p.image.is_none() && p.floating_images.is_empty() && p.textboxes.is_empty()
                                };
                                // A left float leaving a text strip of at least
                                // MIN_NESTED_FLOAT_STRIP beside it may be followed by
                                // text: the text flows beside it (advocacy's picture
                                // tables with their stories to the right).
                                let mut floating_offset = nested_table.position.as_ref()
                                    .filter(|pos| pos.v_anchor == "text" && (nested_table.rows.len() == 1 || plain_side_text) && (top_aligned || pos.v_offset_pt < 0.0 || (block_idx == 0 && pos.v_offset_pt > 0.0)))
                                    .filter(|pos| {
                                        let at_left_edge = matches!(pos.h_position, HorizontalPosition::Offset(v) if v == 0.0);
                                        let next_is_blank = matches!(
                                            cell.content.get(block_idx + 1),
                                            Some(Block::Paragraph(p)) if blank(p) && p.frame_props.is_none() && !p.paragraph_mark_vanish
                                        );
                                        cell.content[block_idx + 1..].iter().enumerate().all(|(i, b)| match b {
                                            Block::Paragraph(p) => blank(p),
                                            Block::Table(t) => t.rows.len() == 1 && (
                                                t.position.as_ref().is_some_and(|p| p.v_anchor == "text")
                                                || (at_left_edge && t.position.is_none() && next_is_blank)
                                                || (i == 0 && t.position.is_none() && t.col_widths.iter().sum::<f32>() <= gap_beside)
                                            ),
                                        }) || plain_side_text
                                            || text_above_floats
                                            || (top_aligned && (spacers_clear_nested_registration(cell, block_idx, cell_text_w, ctx, hf_sub)
                                                || interleaved_spacers_clear_nested_registration(cell, block_idx, cell_text_w, ctx, hf_sub)))
                                    })
                                    // Word keeps a table at the cell's top when its first
                                    // text anchor has a negative vertical offset.
                                    .map(|pos| if block_idx == 0 { pos.v_offset_pt.max(0.0) } else { pos.v_offset_pt })
                                    .map(|offset| offset + negative_anchor_shift);
                                let space_before = prev_space_after;
                                let mut floating_x_offset = None;
                                let auto_left_float = nested_table.position.as_ref().is_some_and(|p|
                                    matches!(p.h_position, crate::model::HorizontalPosition::Offset(v) if (p.h_anchor == "margin" && v + nested_cw.iter().sum::<f32>() > 0.0)
                                        || (p.h_anchor == "column" && plain_side_text && v == 0.0)))
                                    && nested_table.alignment == crate::model::TableAlignment::Left
                                    && nested_table.table_indent == 0.0;
                                if auto_left_float && let Some(offset) = floating_offset {
                                    let top = total_h + space_before + offset;
                                    if left_float_bottom > top {
                                        floating_offset = Some(offset + left_float_bottom - top);
                                        floating_x_offset = Some(nested_table.position.as_ref().unwrap().left_from_text);
                                    }
                                }
                                // An inline table straight after a left float that
                                // leaves it room sits beside the float, level with its
                                // top (nested-float-no-spacer-inline).
                                if nested_table.position.is_none()
                                    && matches!(items.last(), Some(CellContentItem::NestedTable { floating_offset: Some(_), .. }))
                                    && left_float_bottom > total_h
                                    && nested_cw.iter().sum::<f32>() <= left_float_right_gap
                                {
                                    floating_offset = Some(left_float_top - total_h - space_before);
                                    floating_x_offset = Some(cell_text_w - left_float_right_gap);
                                }
                                // An inline nested table's outer border bands sit inside
                                // the cell's flow, as a top-level table's do: Word starts
                                // autofit-cell-widths-nested's inner top band where the
                                // host's ends, and the host's bottom band below its own.
                                let (top_band, bottom_band) = if floating_offset.is_none() {
                                    outer_half_bands(nested_table)
                                } else {
                                    (0.0, 0.0)
                                };
                                if let Some(offset) = floating_offset {
                                    floating_extent = floating_extent.max(total_h + space_before + offset + nested_h);
                                    if auto_left_float {
                                        left_float_bottom = total_h + space_before + offset + nested_h;
                                        left_float_top = total_h + space_before + offset;
                                        left_float_indent = if plain_side_text { nested_cw.iter().sum::<f32>() + nested_table.position.as_ref().unwrap().right_from_text } else { 0.0 };
                                        first_positive_float = block_idx == 0 && nested_table.position.as_ref().is_some_and(|p| p.v_offset_pt >= 0.0);
                                        left_float_right_gap = gap_beside;
                                    }
                                } else {
                                    total_h += space_before + top_band + nested_h + bottom_band;
                                }
                                items.push(CellContentItem::NestedTable {
                                    col_widths: nested_cw,
                                    rows: nested_layouts,
                                    space_before: space_before + top_band,
                                    floating_offset,
                                    floating_x_offset,
                                    bottom_band,
                                });
                                if floating_offset.is_none() { prev_space_after = 0.0; }
                                // A lone end-of-cell mark directly anchoring a
                                // wide float does not add a line below that float.
                                final_float_mark_hidden = block_idx + 2 == block_count
                                    && auto_left_float && left_float_right_gap < MIN_NESTED_FLOAT_STRIP
                                    && floating_offset.is_some_and(|offset| offset >= 0.0)
                                    && matches!(cell.content.get(block_idx+1),Some(Block::Paragraph(p)) if is_text_empty(&p.runs));
                                prev_was_nested_table = mark_after_nested_hidden(block_idx, floating_offset, final_float_mark_hidden, negative_anchor_shift);
                                para_idx += 1;
                            }
                        }
                    }

                    // When a cell ends with a nested table plus the mandatory
                    // end-of-cell paragraph mark (empty, no text), Word does
                    // not count the trailing paragraph's space_after toward
                    // the row height — the mark glyph height and line_h are
                    // already suppressed above via prev_was_nested_table.
                    let trailing_mark_after_table = items.len() >= 2
                        && matches!(items.get(items.len() - 2), Some(CellContentItem::NestedTable { floating_offset, .. })
                            if mark_after_nested_hidden(block_count.saturating_sub(2), *floating_offset, final_float_mark_hidden, negative_anchor_shift))
                        && matches!(items.last(), Some(CellContentItem::Paragraph(p)) if p.lines.is_empty() && p.image_name.is_none() && p.floating_images.is_empty());
                    let trailing_space_after =
                        if trailing_mark_after_table { 0.0 } else { prev_space_after };
                    total_h += trailing_space_after;
                    total_h = total_h.max(floating_extent);
                    if is_rotated {
                        total_h = ecm.top + ecm.bottom + max_rotated_line_w;
                    }
                    if cell.v_merge != VMerge::Restart {
                        max_h = max_h.max(total_h);
                    }
                    CellLayout {
                        items,
                        cm: *ecm,
                        total_height: total_h,
                        trailing_space_after,
                        text_direction: cell.text_direction,
                    }
                })
                .collect();

            // Border bands are already in the cell insets (docx::tables); the
            // 0.5pt once added here was Table Grid's border width.
            let content_h = max_h;
            // An at-least trHeight bounds the row between its border bands,
            // which sit on top of it: 0.5pt Table Grid rows step trHeight + 0.5
            // (belgian_youth 19.85 → 20.5, italian_evaluation 12.0 → 12.48,
            // japanese_interlibrary 21.25 → 21.77). An exact height is the
            // border-to-border pitch (case15: 36.0).
            // The cells' top and bottom margins sit on top of it too, like the
            // bands in their insets: a 30pt row with 5pt margins is 40.5 in
            // Word, and polish_ministry's 33.15pt rows with 1.4pt ones 36.25.
            let insets = cells
                .iter()
                .map(|c| c.cm.top + c.cm.bottom)
                .fold(0.0f32, f32::max);
            let height = match (row.height, row.height_exact) {
                // An exact one gains only the bottom margin: 30pt with 5pt
                // margins is 35.0, with only a 5pt top margin 30.0 (Word probes).
                // The inset's half band is already in the exact pitch.
                (Some(h), true) => {
                    h + cells
                        .iter()
                        .map(|l| l.cm.bottom - bottom_band)
                        .fold(0.0f32, f32::max)
                }
                (Some(h), false) => content_h.max(h + insets),
                _ => content_h,
            };

            RowLayout {
                height,
                cells,
                split_min: (!row.cant_split && !row.height_exact)
                    .then(|| row.height.map_or(0.0, |h| h + insets)),
                keeps_next: row.cells.first().is_some_and(|c| {
                    c.content
                        .iter()
                        .any(|b| matches!(b, Block::Paragraph(p) if p.keep_next))
                }),
            }
        })
        .collect();

    // A merged cell taller than the rows it spans grows the last of them:
    // indonesian_school_admission_checklist's "NO." header (49.7pt of content)
    // makes the second row 29.28pt where its at-least trHeight asks 24.15.
    for (ri, row) in table.rows.iter().enumerate() {
        for (ci, (grid_col, _, cell)) in row.grid_cells().enumerate() {
            if cell.v_merge != VMerge::Restart {
                continue;
            }
            let mut last = ri;
            while table.rows.get(last + 1).is_some_and(|next| {
                next.grid_cells()
                    .any(|(c, _, n)| c == grid_col && n.v_merge == VMerge::Continue)
            }) {
                last += 1;
            }
            let spanned: f32 = layouts[ri..=last].iter().map(|l| l.height).sum();
            let overflow = layouts[ri].cells[ci].total_height - spanned;
            if overflow > 0.0
                && let Some(r) = (ri..=last).rev().find(|&r| !table.rows[r].height_exact)
            {
                layouts[r].height += overflow;
            }
        }
    }
    layouts
}

/// Half of the widest top and bottom border band among a row's cells. A
/// vertical merge's first cell counts the merge's bottom border, except the
/// table's own (polish_municipal's logo cell draws the table's double bottom
/// border under row 2, which must not deepen the title row; covid_insomnia's
/// and geneva's merges ending mid-table do count theirs).
fn row_half_bands(row: &crate::model::TableRow) -> (f32, f32) {
    row.cells.iter().fold((0.0f32, 0.0f32), |(t, b), c| {
        let bottom = c.borders.own_bottom.unwrap_or(c.borders.bottom);
        (
            t.max(c.borders.top.band() / 2.0),
            b.max(bottom.band() / 2.0),
        )
    })
}

/// Half of the widest outer border band at the table's top (first row) and
/// bottom (last row): the part of its outer borders inside its flow height.
pub(super) fn outer_half_bands(table: &Table) -> (f32, f32) {
    (
        table.rows.first().map_or(0.0, |r| row_half_bands(r).0),
        table.rows.last().map_or(0.0, |r| row_half_bands(r).1),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn text_above_float_requires_both_horizontal_and_vertical_room() {
        let mut doc = crate::docx::parse_bytes(
            include_bytes!(
                "../../tests/fixtures/cases/cell-text-above-floats-y1200-n1-long0/input.docx"
            ),
            crate::RevisionView::Final,
        )
        .unwrap();
        let Block::Table(table) = &mut doc.sections[0].blocks[0] else {
            panic!("outer table");
        };
        let cell = &mut table.rows[0].cells[0];
        let fonts = HashMap::from([("Times New Roman".into(), FontEntry::stub())]);
        let fits = |cell: &crate::model::TableCell, width| {
            text_fits_above_nested_floats(
                cell,
                width,
                &fonts,
                crate::model::LineSpacing::Auto(1.0),
                None,
            )
        };
        assert!(fits(cell, 230.0));
        assert!(!fits(cell, 10.0));
        let Block::Table(float) = &mut cell.content[0] else {
            panic!("float");
        };
        float.position.as_mut().unwrap().v_offset_pt = 1.0;
        assert!(!fits(cell, 230.0));
        let Block::Table(float) = &mut cell.content[0] else {
            unreachable!();
        };
        float.position.as_mut().unwrap().v_offset_pt = 60.0;
        let Some(Block::Paragraph(heading)) = cell.content.last_mut() else {
            panic!("heading");
        };
        heading.runs[0].is_line_break = true;
        assert!(!fits(cell, 230.0));
    }
}
