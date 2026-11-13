mod assembly;
mod bands;
mod bidi;
mod chart_legend;
mod charts;
mod charts_ex;
mod charts_radial;
pub(crate) mod color;
mod comments;
mod emf;
mod fonts;
mod footnotes;
mod glyph_grid;
mod header_footer;
mod helpers;
mod hf_fields;
mod images;
mod layout;
mod line_metrics;
mod line_render;
mod line_sizing;
mod link_tagger;
mod list_label;
mod objstm;
mod page_builder;
mod paragraph;
mod positioning;
mod prepass;
mod shape_path;
mod smartart;
mod svg;
mod table;
mod table_layout;
mod tagging;
mod textbox_render;
mod wordart;

use std::collections::{HashMap, HashSet};

use pdf_writer::{Content, Name, Pdf, Ref};

use crate::ConvertOptions;
use crate::error::Error;
use crate::fonts::FontEntry;
use crate::model::{
    Block, Document, FieldCode, FrameProperties, LineNumberRestart, LineSpacing, PageVerticalAlign,
    Paragraph, ParagraphBorder, Run, SectionBreakType, SectionProperties, ShapeFill, ShapeGeometry,
    VRelativeFrom, VerticalPosition, WrapText, WrapType,
};

use crate::fonts::font_key;
use assembly::{HeadingEntry, assemble_pdf_pages};
use bands::{
    Band, FlowFrame, OpenFrame, add_bands, auto_frame_width, block_para, column_geometry,
    empty_frame_table_bands, first_line_reach, first_word_width, floating_table_band, flow_frame,
    follows_floating_table, frame_beside_text, frame_content_height, has_bands, hf_page_bands,
    inter_gap, is_empty_wrapping_frame, lays_out_nothing, leading_para, lifted_frame,
    open_section_first_page, page_anchored, paragraph_line_count, place_first_line,
    place_moved_first_line, prescan_new_page, prescan_page_bands, step_below_bands,
    step_below_existing_bands, step_lines_below_bands, take_page_break_before,
};
use color::{fill_rgb, stroke_border, stroke_segment};
use fonts::collect_and_register_fonts;
use footnotes::{footnote_height, render_endnotes_inline, render_page_footnotes};
use header_footer::{
    HfPageContext, compute_effective_margin_bottom, effective_slot_top, render_header_footer,
    resolve_footer_for_page, resolve_header_for_page,
};
pub(super) use helpers::resolve_line_h;
use helpers::{
    align_offset, auto_ascent_scale, between_space, border_band, border_box_indents, bottom_border,
    collect_paras, draw_horizontal_rule, horizontal_border_ext, joins_border_group,
    side_border_outsets, spacing_leading,
};
use images::{EffectXObjs, EmbeddedImages, embed_all_images};
use layout::{
    CjkLayout, DualRegion, EMPTY_EFFECTS, EMPTY_INLINE_IMAGES, LineNumberArg, LineOpts,
    LinkAnnotation, LinkTagger, PUNCT_PULL_FREE, PUNCT_SQUEEZE, TextLine, build_lines,
    build_paragraph_lines, descender_ratio, grid_baseline_offset, grid_snapped_line_h,
    inserted_picture_descent, is_text_empty, line_h_with_picture, line_max_image_h, lines_height,
    picture_line_bottom, render_paragraph_lines, run_line_metrics, size_lines_by_own_runs,
    tallest_glyph_run_metrics, tallest_run_metrics,
};
use line_metrics::{
    break_font, break_run_lhr, label_boosted_baseline_offset, label_boosted_line_h, mark_font,
    mark_line_h, mark_position_stretch, size_break_lines, size_mark_line, size_tab_mark_line,
    size_unsized_lines, unsized_line_font, unsized_line_metrics,
};
use list_label::{label_font_key, render_list_label};
use page_builder::{
    AnchorKind, BodyPageFields, LayoutState, PageBuilder, ParaAnchor, PendingAnchor,
    last_page_number, page_number_formats, page_numbers, section_break_spacing,
    section_page_counts,
};
use paragraph::render_paragraph_block;
use positioning::{
    FloatTableExtent, FloatZone, FloatingTablePos, MIN_EMPTY_STRIP, MIN_TABLE_SIDE_ROOM,
    blocking_textbox_band, draw_debug_wrap_overlay, first_line_region_x, hangs_below_anchor,
    image_band, image_blocks_band, keep_open, left_text_region, no_strip_beside, para_line_h,
    plain_line_h, poly_scanline, polygon_edges, rectangular_polygon_clearance, render_connector,
    render_floating_images, resolve_fi_x, right_text_region, side_room, side_strip,
    skip_blocked_lines, step_lines_below_floats, textbox_blocks_band, textbox_room_beside,
    textbox_wraps_beside, textbox_zones, wrap_room, wrap_side, wraps_in_column,
};
pub(super) use positioning::{resolve_fi_y_top, resolve_h_position};
use prepass::{
    body_page_field_text, body_runs, cell_grid_pitch, compute_bookmark_positions, document_lang,
    keeps_section_mark_line,
};
use smartart::draw_shape_path;
use table::render_table;
use textbox_render::render_single_textbox;

/// A frame's width: w:w, or its widest block within `avail_w`.
fn frame_width(
    props: &FrameProperties,
    blocks: &[Block],
    ctx: &RenderContext,
    avail_w: f32,
) -> f32 {
    if props.width > 0.0 {
        props.width
    } else {
        auto_frame_width(props, blocks, ctx, avail_w)
    }
}

/// The height a frame takes on the page: w:h if exact, else at least its
/// content at `width`.
fn frame_box_h(props: &FrameProperties, blocks: &[Block], ctx: &RenderContext, width: f32) -> f32 {
    if props.height_exact {
        props.height
    } else {
        props
            .height
            .max(frame_content_height(props, blocks, ctx, width))
    }
}

/// The frame a paragraph opens in the flow (`flow_frame`).
fn flow_frame_of(p: &Paragraph) -> Option<&FrameProperties> {
    p.frame_props.as_ref().filter(|fp| flow_frame(fp))
}

pub(super) struct RenderContext<'a> {
    pub(super) fonts: &'a HashMap<String, FontEntry>,
    /// The document's sections: a section without its own header or footer
    /// lays out around the one it inherits.
    pub(super) sections: &'a [crate::model::Section],
    /// `w:evenAndOddHeaders`: even pages lay out around the even header.
    pub(super) even_and_odd_headers: bool,
    /// Paragraph style names by id, for tagging textbox paragraphs.
    pub(super) style_id_to_name: &'a HashMap<String, String>,
    pub(super) doc_line_spacing: LineSpacing,
    pub(super) note_separator: footnotes::NoteSeparator,
    pub(super) endnote_separator: footnotes::NoteSeparator,
    pub(super) default_tab_stop: f32,
    /// Image names for inline images in table cells, keyed by Arc data pointer address.
    pub(super) table_cell_image_names: &'a HashMap<usize, String>,
    pub(super) effect_table_names: &'a HashMap<usize, EffectXObjs>,
    /// Image names for images inside textbox paragraphs, keyed by Arc data pointer address.
    pub(super) textbox_image_names: &'a HashMap<usize, String>,
    pub(super) chart_font_name: &'a str,
    /// Word's `compressPunctuation` setting (see `docx::settings`).
    pub(super) compress_punctuation: bool,
    /// Display numbers of footnote and endnote reference marks, by note id.
    pub(super) footnote_marks: &'a HashMap<u32, String>,
    pub(super) endnote_marks: &'a HashMap<u32, String>,
    /// Page numbers from a first pass, for PAGE, NUMPAGES and SECTIONPAGES
    /// fields in the body.
    pub(super) body_page_fields: Option<&'a BodyPageFields>,
    /// Each page's w:vAlign displacement from a first pass, which Word applies
    /// before it rounds and stretches the baselines.
    pub(super) valign_offsets: &'a [f32],
    /// What a PAGEREF shows for each page of the bookmark pre-pass.
    pub(super) page_labels: Vec<String>,
    /// Word's `compatibilityMode` (see `docx::settings`).
    pub(super) compat_mode: u32,
    /// Word's `enableOpenTypeFeatures` (see `docx::settings`).
    pub(super) open_type_features: bool,
    /// `Document::east_asian_lang`.
    pub(super) east_asian_lang: bool,
    /// Word's `doNotExpandShiftReturn` (see `docx::settings`).
    pub(super) do_not_expand_shift_return: bool,
    /// Word's `ulTrailSpace` (see `docx::settings`).
    pub(super) ul_trail_space: bool,
    /// Word's `wpJustification` (see `docx::settings`).
    pub(super) wp_justification: bool,
    /// `Document::additive_paragraph_spacing`.
    pub(super) additive_paragraph_spacing: bool,
    /// The current section's docGrid line pitch when its grid snaps lines
    /// and `adjustLineHeightInTable` is set (0 otherwise): table cells then
    /// snap to it like body text; physical_therapy (no flag) keeps natural
    /// cell lines on its 18pt grid.
    pub(super) cell_grid_pitch: std::cell::Cell<f32>,
    /// The current section's page height: Word draws baselines and rules on
    /// a grid from the page top, which the table renderers reach through
    /// `ctx` only.
    pub(super) page_height: std::cell::Cell<f32>,
    /// Comment anchors of table-cell text, which the cell renderers reach
    /// through `ctx` only; the body table code moves them onto the page.
    pub(super) cell_comment_anchors: std::cell::RefCell<Vec<(u32, f32, f32, f32)>>,
    /// Measured header/footer heights (`header_footer::header_height`), which
    /// body pagination asks for per table row and page-top paragraph.
    pub(super) hf_heights: std::cell::RefCell<HashMap<header_footer::HfHeightKey, f32>>,
}

impl RenderContext<'_> {
    /// The text a note reference mark run shows: its note's number. The run's
    /// own text is empty (`docx::runs` `footnoteReference`), unless it is a
    /// custom mark.
    fn note_mark_text(&self, run: &Run) -> Option<String> {
        if run.custom_note_mark {
            return None;
        }
        let (marks, id) = match (run.footnote_id, run.endnote_id) {
            (Some(id), _) => (self.footnote_marks, id),
            (None, Some(id)) => (self.endnote_marks, id),
            (None, None) => return None,
        };
        Some(marks.get(&id).cloned().unwrap_or_default())
    }

    /// `runs` with every note reference mark showing its number, when any has one.
    pub(super) fn with_note_marks(&self, runs: &[Run]) -> Option<Vec<Run>> {
        runs.iter()
            .any(|r| r.footnote_id.is_some() || r.endnote_id.is_some())
            .then(|| {
                runs.iter()
                    .map(|r| match self.note_mark_text(r) {
                        Some(text) => Run { text, ..r.clone() },
                        None => r.clone(),
                    })
                    .collect()
            })
    }
    /// East Asian switches for `build_paragraph_lines`: the paragraph's autospace
    /// choice plus the document-wide punctuation compression.
    fn cjk(&self, auto_space: bool, para: &crate::model::Paragraph) -> CjkLayout {
        CjkLayout {
            auto_space,
            // Word 2013+ squeezes marks only while justifying: the same text
            // set left-aligned wraps where the justified copy compresses; the
            // 2010 layout squeezes any alignment, a little less (Word online
            // probes 2026-10-27, see `PUNCT_SQUEEZE`).
            compress_punct: self.compress_punctuation,
            punct_pull: self.compat_mode < 15
                || matches!(
                    para.alignment,
                    crate::model::Alignment::Justify | crate::model::Alignment::Distribute
                ),
            punct_cap: if self.compat_mode < 15 {
                0.4
            } else {
                PUNCT_SQUEEZE
            },
            punct_pull_free_ratio: if self.compat_mode < 15 {
                0.0
            } else {
                PUNCT_PULL_FREE
            },
            punct_pull_free_em: if self.compat_mode < 15 { 0.5 } else { 0.0 },
            squeeze_spaces: self.compat_mode >= 15
                && matches!(para.alignment, crate::model::Alignment::Justify),
            punct_half_out: self.east_asian_lang,
            squeeze_line: self.wp_justification
                && matches!(para.alignment, crate::model::Alignment::Justify),
            line_reserve: para.kashida_reserve,
            expand_shift_return: !self.do_not_expand_shift_return,
            ul_trail_space: self.ul_trail_space,
            bidi: para.bidi,
            shaped_kerning: self.open_type_features,
            break_words: para.break_words,
        }
    }
}

pub(super) struct GradientSpec {
    pub(super) pattern_name: String,
    pub(super) stops: Vec<([u8; 3], f32)>,
    pub(super) angle_deg: f32,
    pub(super) x: f32,
    pub(super) y: f32,
    pub(super) w: f32,
    pub(super) h: f32,
}

pub(super) fn render_shape_fill(
    content: &mut Content,
    fill: &ShapeFill,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    shape: &ShapeGeometry,
    gradient_specs: &mut Vec<GradientSpec>,
) {
    match fill {
        ShapeFill::Solid(c) => {
            content.save_state();
            fill_rgb(content, *c);
            draw_shape_path(content, x, y, w, h, shape);
            content.fill_nonzero();
            content.restore_state();
        }
        ShapeFill::LinearGradient { stops, angle_deg } => {
            let pat_name = format!("Grd{}", gradient_specs.len());
            content.save_state();
            draw_shape_path(content, x, y, w, h, shape);
            content.clip_nonzero();
            content.end_path();
            content.set_fill_color_space(pdf_writer::types::ColorSpaceOperand::Pattern);
            content.set_fill_pattern([], Name(pat_name.as_bytes()));
            draw_shape_path(content, x, y, w, h, shape);
            content.fill_nonzero();
            content.restore_state();
            gradient_specs.push(GradientSpec {
                pattern_name: pat_name,
                stops: stops.clone(),
                angle_deg: *angle_deg,
                x,
                y,
                w,
                h,
            });
        }
    }
}

/// Record what STYLEREF reads from `para` under the style's id and name.
fn styleref_insert(
    running: &mut HashMap<String, String>,
    page_first: &mut HashMap<String, String>,
    id: &str,
    para: &Paragraph,
    text: &str,
    style_id_to_name: &HashMap<String, String>,
) {
    for name in std::iter::once(id).chain(style_id_to_name.get(id).map(String::as_str)) {
        for (form, value) in para.styleref_values(text) {
            let key = header_footer::styleref_key(name, form);
            if !page_first.contains_key(&key) {
                page_first.insert(key.clone(), value.to_string());
            }
            running.insert(key, value.to_string());
        }
    }
}

fn update_styleref_from_para(
    running: &mut HashMap<String, String>,
    page_first: &mut HashMap<String, String>,
    para: &Paragraph,
    style_id_to_name: &HashMap<String, String>,
) {
    if let Some(ref sid) = para.style_id {
        let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
        if !text.is_empty() {
            styleref_insert(running, page_first, sid, para, &text, style_id_to_name);
        }
    }
    // A character style's text is the whole stretch of runs carrying it.
    for group in para
        .runs
        .chunk_by(|a, b| a.char_style_id == b.char_style_id)
    {
        let Some(ref csid) = group[0].char_style_id else {
            continue;
        };
        let text: String = group.iter().map(|r| r.text.as_str()).collect();
        if !text.is_empty() {
            styleref_insert(running, page_first, csid, para, &text, style_id_to_name);
        }
    }
}

/// The part of a collapsed paragraph gap that is the paragraph's own space
/// before, beyond the previous paragraph's space after.
pub(super) fn own_space_before(gap: f32, prev_space_after: f32) -> f32 {
    (gap - prev_space_after).max(0.0)
}

/// Book footnote `id` into the current page's footnote area (once per page)
/// and shrink the body area by its height, plus the separator for the first.
fn track_page_footnote(
    state: &mut LayoutState,
    doc: &Document,
    ctx: &RenderContext,
    col: (f32, f32),
    id: u32,
) {
    let height = doc
        .footnotes
        .contains_key(&id)
        .then(|| footnote_height(id, &doc.footnotes, ctx, col.1));
    state.pb.book_footnote(
        id,
        col,
        height,
        ctx.note_separator.height,
        &mut state.effective_margin_bottom,
    );
}

/// Footnote ids referenced by `lines`, in reading order, each once.
fn line_footnote_ids(lines: &[TextLine]) -> Vec<u32> {
    let mut seen = HashSet::new();
    lines
        .iter()
        .flat_map(|l| l.chunks.iter())
        .filter_map(|c| c.footnote_id)
        .filter(|id| seen.insert(*id))
        .collect()
}

/// Footnote space each line of a paragraph adds to the page, plus the total.
/// Word charges a footnote to the page carrying its reference mark, so a
/// footnote whose line overflows travels to the next page instead of eating
/// room on this one. `line_refs` holds the ids referenced on each line and
/// `run_refs` every id the paragraph's runs carry: a reference that produced
/// no chunk is charged to the last line, where the split path registers it
/// too. `sep_h` is added with the first footnote the page gets.
fn per_line_footnote_extra(
    line_refs: &[Vec<u32>],
    run_refs: &[u32],
    tracked: &HashSet<u32>,
    sep_h: f32,
    mut footnote_h: impl FnMut(u32) -> f32,
) -> (Vec<f32>, f32) {
    let mut seen = HashSet::new();
    let mut charge = |id: u32| {
        if !tracked.contains(&id) && seen.insert(id) {
            footnote_h(id)
        } else {
            0.0
        }
    };
    let mut per_line: Vec<f32> = line_refs
        .iter()
        .map(|ids| ids.iter().map(|&id| charge(id)).sum())
        .collect();
    let unattributed: f32 = run_refs.iter().map(|&id| charge(id)).sum();
    if let Some(last) = per_line.last_mut() {
        *last += unattributed;
    }
    let mut total: f32 = if per_line.is_empty() {
        unattributed
    } else {
        per_line.iter().sum()
    };
    if total > 0.0 {
        total += sep_h;
        if let Some(first) = per_line.iter_mut().find(|e| **e > 0.0) {
            *first += sep_h;
        }
    }
    (per_line, total)
}

/// How much of a line must fit above the margin for the line to stay on the
/// page: its text height when its trailing leading may hang past the margin,
/// else its whole advance — above a footnote area, and for an exact-spaced
/// line, whose box is fixed (Word for Mac probe p27: 25 lines of 25pt on a
/// 648pt body, the 26th moves on although its 12pt glyphs would end 0.5pt
/// above the margin).
fn line_fit_height(whole_box: bool, own_pitch: Option<f32>, advance: f32, text_h: f32) -> f32 {
    if whole_box {
        advance
    } else {
        own_pitch.map_or(text_h, |p| p.min(text_h))
    }
}

/// `per_line_footnote_extra` for the current page and column: the separator
/// counts only while the column has no footnote area yet.
fn page_footnote_extra(
    state: &LayoutState,
    doc: &Document,
    ctx: &RenderContext,
    col_geometry: &[(f32, f32)],
    line_refs: &[Vec<u32>],
    run_refs: &[u32],
) -> (Vec<f32>, f32) {
    per_line_footnote_extra(
        line_refs,
        run_refs,
        &state.pb.footnote_ids_set,
        if state.pb.col_fn_reserved == 0.0 {
            ctx.note_separator.height
        } else {
            0.0
        },
        |id| footnote_height(id, &doc.footnotes, ctx, col_geometry[state.current_col].1),
    )
}

/// How many of `lines` fit in `room`, `fn_extra[i]` being the footnote room
/// line i adds to the page. A line fits when the advances of the lines above
/// it plus its own text height fit (its trailing leading may hang past the
/// margin), but not past a footnote area: there the whole line must fit
/// (environmental_law_clinic's double-spaced lines stop a line earlier above
/// the footnotes on every page in Word), nor when an at-least minimum sets
/// the line (Word examples: 10pt Arial under 24pt needs all 24pt, on every
/// page of a split), nor for an exact-spaced line (`line_fit_height`).
/// `whole_lines` asks for the whole line in all of these.
fn lines_that_fit(
    lines: &[TextLine],
    fn_extra: &[f32],
    room: f32,
    whole_lines: bool,
    line_h: f32,
    text_h: f32,
    para_metrics: (f32, f32),
) -> usize {
    let mut fit = 0;
    let (mut above, mut fn_acc) = (0.0f32, 0.0f32);
    for (line, fn_extra) in lines.iter().zip(fn_extra) {
        fn_acc += fn_extra;
        let advance = layout::inline_line_advance(line, line_h, para_metrics);
        let own_h = line_fit_height(
            whole_lines || fn_acc > 0.0,
            line.pitch,
            advance - line.gap_after,
            text_h,
        );
        if above + own_h > room - fn_acc {
            break;
        }
        fit += 1;
        above += advance;
    }
    fit
}

/// Lines of an `n`-line paragraph that must share a page with what precedes
/// it: one, or with widow control two — all of them when it has three or
/// fewer, since any split would leave a lone line. `n` is only counted when
/// widow control needs it.
fn lines_kept_together(widow_control: bool, n: impl FnOnce() -> usize) -> usize {
    if !widow_control {
        return 1;
    }
    match n() {
        n if n <= 3 => n.max(1),
        _ => 2,
    }
}

/// Height of `table`'s first row laid out in a column `col_w` wide, as
/// `table::render_table` sizes it.
fn first_row_height(table: &crate::model::Table, ctx: &RenderContext, col_w: f32) -> f32 {
    let mut col_widths = if table.fixed_layout {
        table_layout::fixed_columns(table, col_w, ctx.compat_mode)
    } else {
        table_layout::auto_fit_columns(
            table,
            ctx.fonts,
            Some(col_w),
            Some(col_w),
            table.autofit_overhang(ctx.compat_mode),
        )
    };
    table_layout::apply_pct_width(table, &mut col_widths, col_w, ctx.compat_mode);
    table_layout::compute_row_layouts(table, &col_widths, ctx, None)
        .first()
        .map_or(0.0, |r| r.height)
}

/// About how many lines `para` lays out to in a column `col_w` wide. Every
/// line gets the body measure: a hanging label tabs its first line's text out
/// to the indent anyway (western_australia's "(a)" items).
fn line_count(para: &Paragraph, ctx: &RenderContext, col_w: f32) -> usize {
    let width = (col_w - para.indent_left - para.indent_right).max(1.0);
    build_paragraph_lines(
        &para.runs,
        ctx.fonts,
        width,
        0.0,
        &EMPTY_INLINE_IMAGES,
        &EMPTY_EFFECTS,
        None,
        None,
        None,
        ctx.cjk(para.auto_space_de || para.auto_space_dn, para),
    )
    .len()
}

/// Markup view: the change bars of one page. Word draws them in the left
/// margin (the outside margin with mirrored margins), half an inch from the
/// text but no deeper than halfway into the margin, 0.72pt wide, on its
/// 0.48pt device grid (Word for Mac probes 2026-11-05: margins of 0.5, 1,
/// 1.25 and 2 inches; two columns still bar at the page's left).
fn draw_change_bars(
    content: &mut Content,
    bars: &[(f32, f32)],
    sp: &SectionProperties,
    mirror_margins: bool,
    page_idx: usize,
) {
    if bars.is_empty() {
        return;
    }
    let snap = |v: f32| (v / 0.48).round() * 0.48;
    let x = if mirror_margins && page_idx % 2 == 0 {
        sp.page_width - sp.margin_right + f32::min(36.0, sp.margin_right / 2.0)
    } else {
        sp.margin_left - f32::min(36.0, sp.margin_left / 2.0)
    };
    // x lands on a 0.24pt step, rounded down (27 → 26.88, 585 → 584.88).
    let x = (x / 0.24).floor() * 0.24;
    content.save_state();
    content.set_fill_rgb(0.0, 0.0, 0.0);
    for &(bottom, top) in bars {
        let (bottom, top) = (snap(bottom), snap(top));
        content.rect(x, bottom, 0.72, top - bottom).fill_nonzero();
    }
    content.restore_state();
}

/// Compute effective first-line hanging indent for a paragraph.
fn compute_text_hanging(
    para: &Paragraph,
    default_tab_stop: f32,
    fonts: &HashMap<String, FontEntry>,
) -> f32 {
    list_label::text_hanging(para, default_tab_stop, fonts)
}

/// Structure type of a body paragraph. Word tags outline levels as H1–H6
/// (deeper levels stay H6) and its Title style as Title, role-mapped to H1.
pub(super) fn para_tag_kind(para: &Paragraph, style_name: Option<&String>) -> &'static str {
    if style_name.is_some_and(|n| n.eq_ignore_ascii_case("title")) {
        return "H1";
    }
    match para.outline_level {
        Some(l) => ["H1", "H2", "H3", "H4", "H5", "H6"][usize::from(l).min(5)],
        None if sole_formula(para) => "Formula",
        None => "P",
    }
}

/// A paragraph holding one equation and nothing else (a display equation, or
/// a zone alone on its line) is that Formula, read by its `/Alt` (set when
/// the zone is drawn, see `LinkTagger`), not a P around it.
pub(super) fn sole_formula(para: &Paragraph) -> bool {
    para.image.is_none()
        && para.runs.iter().filter(|r| r.math.is_some()).count() == 1
        && para.runs.iter().all(|r| {
            r.math.is_some()
                || (r.text.trim().is_empty()
                    && r.inline_image.is_none()
                    && r.footnote_id.is_none()
                    && r.endnote_id.is_none())
        })
}

/// Word's SmartArt alt text: the diagrams' text, one line per node.
fn smartart_alt(diagrams: &[crate::model::SmartArtDiagram]) -> Option<String> {
    let lines: Vec<String> = diagrams
        .iter()
        .flat_map(|d| &d.shapes)
        .flat_map(|s| &s.paragraphs)
        .map(|p| p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
        .filter(|t| !t.trim().is_empty())
        .collect();
    (!lines.is_empty()).then(|| lines.join("\n"))
}

pub fn render(doc: &Document, options: &ConvertOptions) -> Result<Vec<u8>, Error> {
    let (pdf, page_fields, valign_offsets) = render_pass(doc, None, &[], options)?;
    // Word shows body PAGE, NUMPAGES and SECTIONPAGES fields live, not their cached result,
    // and their values are known only once the document is paginated.
    let has_page_fields =
        body_runs(doc).any(|r| r.field_code.as_ref().is_some_and(FieldCode::is_page_count));
    if !has_page_fields && !valign_offsets.iter().any(|&off| off > 0.0) {
        return Ok(pdf);
    }
    // The fields keep the first pass's numbers even where their new text
    // moves the pages (Word for Mac: NUMPAGES shows 3 on a page its own
    // shorter result pulled into a two-page document), but the vAlign offsets
    // the lines are rounded with must come from the pass they are drawn in.
    let page_fields = has_page_fields.then_some(&page_fields);
    let (pdf, _, measured) = render_pass(doc, page_fields, &valign_offsets, options)?;
    if measured == valign_offsets {
        return Ok(pdf);
    }
    render_pass(doc, page_fields, &measured, options).map(|(pdf, ..)| pdf)
}

fn render_pass(
    doc: &Document,
    page_fields: Option<&BodyPageFields>,
    valign_offsets: &[f32],
    options: &ConvertOptions,
) -> Result<(Vec<u8>, BodyPageFields, Vec<f32>), Error> {
    let debug_wrap = options.debug_wrap;
    let t0 = std::time::Instant::now();
    // Bars collected by a previous pass (headers, footers) must not leak in.
    layout::take_change_bars();
    let mut pdf = Pdf::new();
    let mut next_id = 1i32;
    let mut alloc = || {
        let r = Ref::new(next_id);
        next_id += 1;
        r
    };

    let catalog_id = alloc();
    let pages_id = alloc();

    let (seen_fonts, font_order) = collect_and_register_fonts(doc, &mut pdf, &mut alloc, options);
    let smartart_font_key = font_order.first().map(|s| s.as_str()).unwrap_or("");
    let t_fonts = t0.elapsed();

    let images = embed_all_images(
        doc,
        &mut pdf,
        &mut alloc,
        &seen_fonts,
        options.resample_pictures,
    );

    let t_images = t0.elapsed();

    // Pre-compute footnote and endnote display order: scan body runs for
    // footnote_id / endnote_id, assign sequential numbers in encounter order.
    let mut footnote_display_order: HashMap<u32, String> = HashMap::new();
    let mut endnote_display_order: HashMap<u32, String> = HashMap::new();
    {
        // §17.11.18/.17 mark numbering format. Word reads it from the SECTION's
        // sectPr footnotePr/endnotePr (NOT the doc-wide settings.xml bag — case74
        // declares upperRoman/lowerLetter there yet renders the defaults 1,2,3 / i,ii).
        // Built-in defaults: footnote decimal, endnote lowerRoman. ponytail: notes are
        // numbered doc-wide, so we take the first section that names a format; per-section
        // formats in multi-section docs are unexercised.
        let fn_fmt = doc
            .sections
            .iter()
            .find_map(|s| s.properties.footnote_num_fmt.as_deref())
            .unwrap_or("decimal");
        let en_fmt = doc
            .sections
            .iter()
            .find_map(|s| s.properties.endnote_num_fmt.as_deref())
            .unwrap_or("lowerRoman");
        let mut next_fn_num = 1u32;
        let mut next_en_num = 1u32;
        for run in body_runs(doc) {
            if let Some(id) = run.footnote_id {
                if run.custom_note_mark {
                    footnote_display_order.entry(id).or_default();
                    continue;
                }
                footnote_display_order.entry(id).or_insert_with(|| {
                    next_fn_num += 1;
                    crate::docx::numbering::format_number(next_fn_num - 1, fn_fmt)
                });
            }
            if let Some(id) = run.endnote_id {
                if run.custom_note_mark {
                    endnote_display_order.entry(id).or_default();
                    continue;
                }
                endnote_display_order.entry(id).or_insert_with(|| {
                    next_en_num += 1;
                    crate::docx::numbering::format_number(next_en_num - 1, en_fmt)
                });
            }
        }
    }

    let mut ctx = RenderContext {
        fonts: &seen_fonts,
        sections: &doc.sections,
        even_and_odd_headers: doc.even_and_odd_headers,
        style_id_to_name: &doc.style_id_to_name,
        doc_line_spacing: doc.line_spacing,
        default_tab_stop: doc.default_tab_stop,
        table_cell_image_names: &images.table_cell_image_names,
        effect_table_names: &images.effect_table_names,
        textbox_image_names: &images.textbox_image_names,
        chart_font_name: &doc.theme_minor_font,
        compress_punctuation: doc.compress_punctuation,
        footnote_marks: &footnote_display_order,
        endnote_marks: &endnote_display_order,
        body_page_fields: page_fields,
        valign_offsets,
        page_labels: Vec::new(),
        compat_mode: doc.compat_mode,
        open_type_features: doc.open_type_features,
        east_asian_lang: doc.east_asian_lang,
        do_not_expand_shift_return: doc.do_not_expand_shift_return,
        ul_trail_space: doc.ul_trail_space,
        wp_justification: doc.wp_justification,
        additive_paragraph_spacing: doc.additive_paragraph_spacing,
        note_separator: footnotes::NoteSeparator::new(
            doc.footnote_separator.as_ref(),
            &seen_fonts,
            doc.line_spacing,
        ),
        endnote_separator: footnotes::NoteSeparator::new(
            doc.endnote_separator.as_ref(),
            &seen_fonts,
            doc.line_spacing,
        ),
        cell_grid_pitch: std::cell::Cell::new(
            doc.sections
                .first()
                .map_or(0.0, |s| cell_grid_pitch(doc, &s.properties)),
        ),
        page_height: std::cell::Cell::new(
            doc.sections
                .first()
                .map_or(0.0, |s| s.properties.page_height),
        ),
        cell_comment_anchors: Default::default(),
        hf_heights: Default::default(),
    };

    let (bookmark_positions, page_labels) = compute_bookmark_positions(doc, &ctx);
    ctx.page_labels = page_labels;

    // Phase 2: build multi-page content streams (section-aware)
    let first_sp = &doc.sections[0].properties;
    let mut cur_sp = first_sp;
    let first_page = last_page_number(&doc.sections, &[(0, true, 0)]);
    let initial_slot_top = effective_slot_top(cur_sp, true, first_page, &ctx);
    let mut state = LayoutState {
        pb: PageBuilder::new(initial_slot_top),
        prev_space_after: 0.0,
        prev_between_below: 0.0,
        effective_margin_bottom: compute_effective_margin_bottom(cur_sp, true, first_page, &ctx),
        current_col: 0,
        global_block_idx: 0,
        heading_entries: Vec::new(),
        bookmark_positions,
        long_keep_chain: false,
        reflowed: false,
        band_prescan: false,
        page_break_taken: false,
        frame_overflow: false,
    };
    state.pb.tags.lang = document_lang(doc);
    state.pb.frame_bands = hf_page_bands(
        cur_sp,
        true,
        first_page,
        state.effective_margin_bottom,
        &ctx,
    );

    for (sect_idx, section) in doc.sections.iter().enumerate() {
        let sp = &section.properties;
        ctx.cell_grid_pitch.set(cell_grid_pitch(doc, sp));
        ctx.page_height.set(sp.page_height);

        // §17.6.8 restart="newPage": set before the break handling so the
        // flush that opens this section's first page already resets the count.
        state.pb.line_numbers_restart_each_page = sp
            .line_numbering
            .as_ref()
            .is_some_and(|ln| ln.restart == LineNumberRestart::NewPage);

        // Section break handling (not for the first section)
        if sect_idx > 0 {
            match sp.break_type {
                SectionBreakType::NextPage
                | SectionBreakType::OddPage
                | SectionBreakType::EvenPage => {
                    // A page break just before the section break leaves an empty
                    // page; Word starts the section there instead of after it.
                    if !state.pb.is_empty_page() {
                        state.pb.flush_page(sect_idx - 1);
                    }
                    state.pb.flush_note_page(sect_idx - 1);

                    // Insert blank page for odd/even page alignment
                    let restart_parity = sp
                        .page_num_start
                        .filter(|_| doc.even_and_odd_headers)
                        .map(|start| start % 2);
                    let need_odd =
                        sp.break_type == SectionBreakType::OddPage || restart_parity == Some(1);
                    let need_even =
                        sp.break_type == SectionBreakType::EvenPage || restart_parity == Some(0);
                    if need_odd || need_even {
                        let explicit_parity_break = matches!(
                            sp.break_type,
                            SectionBreakType::OddPage | SectionBreakType::EvenPage
                        );
                        let filler = if explicit_parity_break {
                            // Word's filler page follows the number the section would get by
                            // continuing. A restarted section skips it (and bumps its own start
                            // instead, see page_numbers) unless evenAndOddHeaders/mirrorMargins
                            // ask for print-ready sheets (measured with Word probes, 2026-10-03).
                            (sp.page_num_start.is_none()
                                || doc.even_and_odd_headers
                                || doc.mirror_margins)
                                && {
                                    let continuing =
                                        page_numbers(&doc.sections, &state.pb.page_section_indices)
                                            .last()
                                            .map_or(1, |n| n + 1);
                                    (continuing % 2 == 1) != need_odd
                                }
                        } else {
                            // The evenAndOddHeaders heuristic (NextPage break with a restart)
                            // lands the section on the right PHYSICAL sheet for header selection.
                            let physical = state.pb.page_count() + 1;
                            (need_odd && physical % 2 == 0) || (need_even && physical % 2 == 1)
                        };
                        if filler {
                            state.pb.push_blank_page(sect_idx - 1);
                        }
                    }

                    open_section_first_page(&mut state, sect_idx, sp, &ctx);
                }
                SectionBreakType::Continuous => {
                    if state.pb.is_at_page_top(cur_sp)
                        && !doc.sections[sect_idx - 1].blocks.iter().any(
                            |block| matches!(block, Block::Paragraph(p) if p.frame_props.is_some()),
                        )
                    {
                        // No content on this page belongs to the preceding
                        // section. The continuous section therefore owns this
                        // sheet, including its first-page header/footer variant.
                        open_section_first_page(&mut state, sect_idx, sp, &ctx);
                    }
                    // Mid-page, the sheet keeps the section that started it;
                    // geometry updates on the next page without a forced break.
                }
            }
        }

        // §17.6.8 restart: newSection starts over here; newPage does too when
        // the section owns a fresh sheet (a continuous section joining a page
        // mid-way keeps that page's count).
        if let Some(ln) = &sp.line_numbering
            && match ln.restart {
                LineNumberRestart::NewSection => true,
                LineNumberRestart::NewPage => state.pb.page_hf_section == sect_idx,
                LineNumberRestart::Continuous => false,
            }
        {
            state.pb.line_number_counter = 0;
        }

        cur_sp = sp;
        // If the first block itself moves to the next sheet, that sheet is the
        // continuous section's first page, whatever the block is (Word
        // examples: a 100pt table row, floating table, inline picture or kept
        // 6-line paragraph, also one with an anchored picture or text box, or
        // in two columns, below a page of section 1 opens page 2 under the
        // first-page footer, numbered 1). Once it flows here, or when the
        // section already owns this page, a later break gets the default
        // variant as before.
        state.pb.pending_section_first_page = (state.pb.page_hf_section != sect_idx
            && matches!(sp.break_type, SectionBreakType::Continuous)
            && state.pb.float_zone.is_none())
        .then_some((sect_idx, state.pb.page_count(), state.pb.slot_top));
        let text_width = sp.text_width();

        // Column geometry: vec of (x_offset, width) for each column
        let col_config = sp.columns.as_ref();
        let col_count = col_config.map(|c| c.columns.len()).unwrap_or(1);
        let col_geometry = column_geometry(sp);
        state.current_col = 0;
        // Record the starting y for this section's columns on the current
        // page. For a mid-page continuous section, both columns begin at the
        // same y rather than at the top of the page.
        state.pb.column_top_y = state.pb.slot_top;
        // Mid-page, the columns start below the pending space after; the
        // first paragraph still opens with max(space after, its space before),
        // so a heading's extra space stays in column 1 (covid's column 2 starts
        // above its heading; case80's columns line up).
        if col_count > 1 && !state.pb.is_at_page_top(sp) {
            state.pb.column_top_y -= state.prev_space_after;
        }
        state.pb.region_sep_xs = match col_config {
            Some(cfg) if cfg.sep => col_geometry
                .iter()
                .zip(&cfg.columns)
                .take(col_count - 1)
                .map(|(&(x, w), col)| x + w + col.space / 2.0)
                .collect(),
            _ => Vec::new(),
        };

        // The page pre-scan lays lines out twice; most sections have no band
        // to look for, and past the last paragraph or table with one it finds nothing.
        let last_band_block = (col_count == 1)
            .then(|| {
                section.blocks.iter().rposition(|b| match b {
                    Block::Paragraph(p) => has_bands(p, sp, col_geometry[0], text_width),
                    Block::Table(t) => floating_table_band(t, sp, col_geometry[0], &ctx).is_some(),
                })
            })
            .flatten();
        let empty_frame_anchor_page = state.pb.page_count();
        let empty_frame_bands = if col_count == 1 {
            empty_frame_table_bands(
                &section.blocks,
                sp,
                &ctx,
                sp.page_height - state.pb.slot_top,
                state.pb.slot_top - sp.margin_bottom,
            )
        } else {
            Vec::new()
        };

        let layout_blocks = |state: &mut LayoutState| {
            let mut frame: Option<OpenFrame> = None;
            let mut open_flow_frame: Option<FlowFrame> = None;
            let mut last_block_consumption = None;
            for (block_idx, block) in section.blocks.iter().enumerate() {
                let previous_block_consumption = last_block_consumption.take();
                let block_start = (state.pb.page_count(), state.pb.slot_top);
                let block_frame = lifted_frame(block);
                if let Some(f) = frame.take_if(|f| block_frame.is_none_or(|fp| fp != f.props)) {
                    f.close(state, sp);
                }
                let block_flow_frame = leading_para(block).and_then(flow_frame_of);
                if let Some(f) = open_flow_frame.take_if(|f| block_flow_frame != Some(f.props)) {
                    f.close(state);
                }
                if open_flow_frame.is_none()
                    && let Some(fp) = block_flow_frame
                {
                    let blocks = &section.blocks[block_idx..];
                    open_flow_frame = Some(FlowFrame::open(
                        fp,
                        blocks,
                        state,
                        &ctx,
                        sp,
                        (sect_idx, &col_geometry),
                    ));
                }
                // The first block stayed on this page, so a later break opens an
                // ordinary one, even from where the first block started.
                if block_idx > 0 {
                    state.pb.pending_section_first_page = None;
                }
                state.band_prescan = last_band_block.is_some_and(|l| block_idx <= l)
                    && frame.is_none()
                    && block_frame.is_none();
                if frame.is_none() {
                    match block_frame {
                        Some(fp) => {
                            let blocks = &section.blocks[block_idx..];
                            frame = Some(OpenFrame::open(
                                fp,
                                blocks,
                                state,
                                &ctx,
                                sp,
                                (sect_idx, &col_geometry),
                            ));
                        }
                        None => {
                            // The break comes first: the bands are those of the
                            // page the paragraph starts (Word examples: a page-top
                            // paragraph with page break before, or one breaking
                            // mid-page, goes below the new page's table). An empty
                            // break paragraph lays nothing out and keeps out of it.
                            if last_band_block.is_some()
                                && let Block::Paragraph(p) = block
                                && p.page_break_before
                                && !lays_out_nothing(p)
                            {
                                take_page_break_before(state, p, sect_idx, sp, &ctx);
                                state.page_break_taken = true;
                            }
                            let col = col_geometry[state.current_col];
                            let last = last_band_block.map_or(block_idx, |l| l.max(block_idx));
                            let rest = &section.blocks[block_idx..=last];
                            prescan_new_page(state, rest, sect_idx, sp, col, text_width, &ctx);
                            step_below_bands(
                                state, block_idx, block, sp, col, text_width, None, &ctx,
                            );
                        }
                    }
                }
                let (geometry, cols, width) = match &frame {
                    Some(f) => (std::slice::from_ref(&f.geometry), 1, f.geometry.1),
                    None => (&col_geometry[..], col_count, text_width),
                };

                // A table may move the text below its wrap zone while the
                // paragraph-relative connector retains the original anchor.
                let table_connector_anchor = state
                    .pb
                    .float_zone
                    .as_ref()
                    .filter(|zone| zone.from_table)
                    .map(|_| (state.pb.page_count(), state.pb.slot_top));

                // With no side strip, the whole empty clearing paragraph sits
                // below the floating table, including the line ended by br.
                if matches!(block, Block::Paragraph(p) if p.clears_floats && p.runs.iter().all(|r| r.is_line_break || (r.text.is_empty() && !r.is_tab && r.inline_image.is_none())))
                    && follows_floating_table(&section.blocks, block_idx)
                    && let Some(ref zone) = state.pb.float_zone
                {
                    if zone.side_strip(col_geometry[state.current_col]) < MIN_EMPTY_STRIP
                        && state.pb.slot_top > zone.bottom_y
                    {
                        state.pb.slot_top = zone.bottom_y;
                        state.pb.float_zone = None;
                    }
                }

                if frame.is_none()
                    && !empty_frame_bands.is_empty()
                    && state.pb.page_count() == empty_frame_anchor_page
                    && matches!(block, Block::Table(t) if t.position.is_none()
                        && t.col_widths.iter().sum::<f32>() >= text_width - 0.05)
                {
                    let top = sp.page_height - state.pb.slot_top;
                    let bottom = empty_frame_bands
                        .iter()
                        .filter(|(start, end)| top >= *start && top < *end)
                        .map(|(_, end)| *end)
                        .fold(top, f32::max);
                    if bottom > top {
                        state.pb.slot_top = sp.page_height - bottom;
                    }
                }

                let before_float_clearance = state.pb.slot_top;
                let mut table_cleared_float = false;
                // If a float zone is active, decide whether to wrap text beside
                // the object or push it below.
                if let Some(ref fz) = state.pb.float_zone {
                    let rect_clear = rectangular_polygon_clearance(
                        block,
                        fz,
                        state.prev_space_after,
                        &ctx,
                        col_geometry[state.current_col],
                    );
                    let clear_bottom = rect_clear.as_ref().map_or(fz.bottom_y, |rect| rect.bottom);
                    if state.pb.slot_top <= clear_bottom {
                        // Already past the zone — clear it
                        state.pb.next_float_zone();
                    } else if match &rect_clear {
                        Some(rect) => state.pb.slot_top - rect.line_h < rect.top,
                        None => {
                            state.pb.slot_top <= fz.top_y
                            || (fz.para_relative && state.pb.slot_top <= fz.top_y + 30.0)
                            // A floating table meets a paragraph whose first
                            // line reaches its top: physical_education's empty
                            // anchor 0.05pt above a full-width table goes below
                            // it with its line and space after. A table anchored
                            // to the page or margin does the same (Word examples,
                            // fixture page-table-line-reach: the anchor's and
                            // later lines 1 to 10pt into the table, or ending at
                            // its top, go below it; a 0.5pt gap keeps the line).
                            || (fz.from_table
                                && matches!(block, Block::Paragraph(p) if {
                                    let col_w = col_geometry[state.current_col].1;
                                    let reach = first_line_reach(
                                        p,
                                        (&section.blocks, block_idx),
                                        col_w,
                                        state,
                                        sp,
                                        &ctx,
                                    );
                                    state.pb.slot_top - reach < fz.top_y + 0.1
                                }))
                        }
                    } {
                        // Cursor is within, entering, or (for paragraph-relative
                        // zones) slightly above the zone.  Paragraph-relative
                        // images with a positive vertical offset create zones that
                        // start below the anchor paragraph; the next paragraph's
                        // cursor may still be above the zone top.
                        let (col_x, col_w) = col_geometry[state.current_col];
                        let (ex_left, ex_right) = fz.exclusion_at_y(state.pb.slot_top);
                        let space_right = (col_x + col_w) - (ex_right + fz.right_from_text);
                        let space_left = (ex_left - fz.left_from_text) - col_x;
                        // Word 2013+ requires even an empty line's indented box
                        // to fit beside a float; legacy mode ignores its indents.
                        // A negative indent's strip in the margin holds nothing
                        // (examples: an empty line indented 42.5pt into the margin
                        // beside a 37pt gap there goes below the table).
                        let (space_left, space_right) = match block {
                            Block::Paragraph(p)
                                if ctx.compat_mode >= 15 && is_text_empty(&p.runs) =>
                            {
                                let left = col_x + p.indent_left.max(0.0);
                                let right = col_x + col_w - p.indent_right.max(0.0);
                                (
                                    (ex_left - fz.left_from_text).min(right) - left,
                                    right - (ex_right + fz.right_from_text).max(left),
                                )
                            }
                            _ => (space_left, space_right),
                        };
                        let first_word = match block {
                            Block::Paragraph(p) => first_word_width(p, &ctx),
                            Block::Table(_) => None,
                        };
                        let enough_space = match first_word {
                            // A line goes beside a float when its first word
                            // fits a strip there (Word examples: "C1P0", 25pt,
                            // holds a 30pt strip, not a 24pt one).
                            Some(w) => {
                                let room = side_room(fz.wrap_text, space_left, space_right);
                                room >= w.max(MIN_EMPTY_STRIP)
                            }
                            None => {
                                wrap_room(fz.wrap_text, space_left, space_right)
                                    >= MIN_TABLE_SIDE_ROOM
                            }
                        };
                        if !enough_space {
                            // Empty paragraphs can be absorbed within a wide
                            // image's vertical extent without needing wrap space —
                            // but only when a usable side strip exists for their
                            // line boxes. When the float spans the full column
                            // (sample500kB: image width == text width) Word stacks
                            // even empty paragraphs below it; with a real strip
                            // (brazilian: ~42pt) they sit beside. 18pt threshold
                            // splits the two observed cases.
                            // Include paragraphs with only line breaks (w:br)
                            // as "empty" — they have no visible text content.
                            let has_side_strip = space_right.max(space_left) >= MIN_EMPTY_STRIP;
                            let is_empty_para = matches!(block,
                                Block::Paragraph(p) if p.runs.iter().all(|r|
                                    r.vanish || r.is_line_break
                                    || (r.text.is_empty() && !r.is_tab && r.inline_image.is_none())
                                )
                                    && p.image.is_none()
                                    && p.inline_chart.is_none()
                                    && p.smartart.is_empty()
                            );
                            if !is_empty_para || !has_side_strip {
                                table_cleared_float = matches!(block, Block::Table(_));
                                // Its floating paragraph-relative drawings stay where
                                // it stood (japanese_hr's form box anchored in the
                                // empty paragraph Word moves below the table); inline
                                // ones go with its line (page-table-line-reach).
                                let mut cleared = rect_clear
                                    .as_ref()
                                    .map_or(fz.bottom_y, |rect| rect.cleared_y(state.pb.slot_top));
                                if fz.from_table && matches!(block, Block::Paragraph(_)) {
                                    state.pb.pin_float_anchor(
                                        state.pb.slot_top - state.prev_space_after,
                                    );
                                    // Below the table the previous paragraph's space
                                    // after is gone, as below a band; the paragraph
                                    // right after a floating table keeps all of its
                                    // space before (Word examples, 6pt after and 12pt
                                    // before: 6pt below the table, 12pt with the
                                    // table between the two paragraphs).
                                    if follows_floating_table(&section.blocks, block_idx) {
                                        state.prev_space_after = 0.0;
                                    } else {
                                        cleared += state.prev_space_after;
                                    }
                                }
                                state.pb.slot_top = cleared;
                                state.pb.next_float_zone();
                            }
                        }
                        // Otherwise leave zone active — paragraph layout adjusts width
                    }
                }

                // Clearing a floating table may move the first body line into
                // a later frame band. Recheck the existing bands at that cursor.
                if frame.is_none() && state.pb.slot_top != before_float_clearance {
                    let col = col_geometry[state.current_col];
                    step_below_existing_bands(state, block, sp, col, None, &ctx);
                }

                match block {
                    Block::Paragraph(para) => {
                        let skip = render_paragraph_block(
                            para,
                            state,
                            &ctx,
                            cur_sp,
                            geometry,
                            cols,
                            width,
                            sect_idx,
                            block_idx,
                            &section.blocks,
                            &images,
                            doc,
                            smartart_font_key,
                            debug_wrap,
                            table_connector_anchor,
                        );
                        state.page_break_taken = false;
                        state.pb.tags.attach_hoisted();
                        if skip {
                            continue;
                        }
                    }

                    Block::Table(table) => {
                        state.pb.lists.close();
                        state.pb.toc = None;
                        let override_pos = table.position.as_ref().map(|pos| {
                            let (col_x, col_w) = col_geometry[state.current_col];
                            let mut resolved = FloatingTablePos::resolve(
                                table,
                                pos,
                                sp,
                                col_x,
                                col_w,
                                state.pb.slot_top,
                                &ctx,
                            );
                            if table_cleared_float {
                                resolved.constrain_nonoverlap_left(pos, sp, col_x, ctx.compat_mode);
                            }
                            if table_cleared_float
                                && !pos.allow_overlap
                                && resolved.v_anchor_text
                                && resolved.v_offset_pt >= 0.0
                            {
                                resolved.y = state.pb.slot_top - pos.top_from_text;
                            }
                            resolved
                        });
                        if table.position.is_none() {
                            let (col_x, col_w) = col_geometry[state.current_col];
                            let w: f32 = table.col_widths.iter().sum();
                            let left = col_x + table.table_indent;
                            let right = if w > 0.0 { left + w } else { col_x + col_w };
                            state.pb.slot_top = state.pb.below_header_floats(sp, &ctx, left, right);
                            state.pb.table_extent = Some((left, right));
                        }
                        let col_bounds =
                            (cols > 1 || frame.is_some()).then(|| geometry[state.current_col]);
                        let table_tags =
                            tagging::TableTags::for_table(&mut state.pb.tags, tagging::ROOT, table);
                        state.pb.table_tags = Some(table_tags);
                        render_table(
                            table,
                            sp,
                            &ctx,
                            &mut state.pb,
                            sect_idx,
                            state.prev_space_after,
                            override_pos,
                            &doc.footnotes,
                            &mut state.effective_margin_bottom,
                            col_bounds,
                        );
                        state.pb.table_extent = None;
                        // A full-width, negatively positioned text float displaces
                        // its preceding empty anchor line below the table in Word.
                        // Keep the table's measured position, but reserve that line
                        // in its clearing extent instead of losing it above the float.
                        if table
                            .position
                            .as_ref()
                            .is_some_and(|p| p.v_anchor == "text" && p.v_offset_pt < 0.0)
                            && block_idx
                                .checked_sub(1)
                                .and_then(|i| section.blocks.get(i))
                                .is_some_and(|b| {
                                    matches!(b, Block::Paragraph(p) if is_text_empty(&p.runs)
                                    && p.content_height == 0.0 && p.image.is_none()
                                    && !p.is_section_break && p.frame_props.is_none())
                                })
                            && let Some((page, consumed)) = previous_block_consumption
                            && page == state.pb.page_count()
                            && consumed > 0.0
                            && let Some(zone) = state.pb.float_zone.as_mut()
                        {
                            if zone.side_strip(geometry[state.current_col]) < MIN_EMPTY_STRIP {
                                zone.bottom_y -= consumed;
                            }
                        }
                        if let Some(mut tags) = state.pb.table_tags.take() {
                            tags.finish(&mut state.pb.tags);
                        }
                        state.prev_space_after = 0.0;

                        // Update styleref tracking (footnotes are already tracked
                        // inside render_table incrementally per row).
                        for p in table.all_paragraphs() {
                            update_styleref_from_para(
                                &mut state.pb.styleref_running,
                                &mut state.pb.styleref_page_first,
                                p,
                                &doc.style_id_to_name,
                            );
                        }
                    }
                }
                // §17.3.3.1 br clear="all": content after this paragraph restarts
                // below any floating objects.
                if let Block::Paragraph(p) = block
                    && p.clears_floats
                    && let Some(ref fz) = state.pb.float_zone
                {
                    let empty_clear = p
                        .runs
                        .iter()
                        .all(|r| r.text.is_empty() && r.inline_image.is_none())
                        && p.image.is_none()
                        && p.inline_chart.is_none()
                        && p.smartart.is_empty()
                        && p.floating_images.is_empty()
                        && p.textboxes.is_empty();
                    let extra_breaks = if empty_clear {
                        p.runs
                            .iter()
                            .filter(|r| r.is_line_break)
                            .count()
                            .saturating_sub(1)
                    } else {
                        0
                    };
                    let (fs, lhr, _) = tallest_run_metrics(&p.runs, ctx.fonts);
                    let ls = p.line_spacing.unwrap_or(ctx.doc_line_spacing);
                    let (_, break_lhr) = unsized_line_metrics(p, fs, ctx.fonts);
                    let mark_pitch = mark_line_h(p, ls, break_lhr.or(lhr), ctx.fonts)
                        .unwrap_or_else(|| resolve_line_h(ls, fs, lhr.or(break_lhr)));
                    let break_lines: f32 = p
                        .runs
                        .iter()
                        .filter(|r| r.is_line_break)
                        .skip(1)
                        .take(extra_breaks)
                        .map(|br| {
                            let (size, (br_lhr, _)) = break_font(br, ctx.fonts);
                            resolve_line_h(ls, size, br_lhr.or(lhr))
                        })
                        .sum();
                    // It clears every float the cursor has not passed yet
                    // (Word example: two boxes 250pt and 80pt tall, text below
                    // the taller).
                    let float_bottom = state
                        .pb
                        .more_float_zones
                        .iter()
                        .map(|z| z.bottom_y)
                        .filter(|&b| b < state.pb.slot_top)
                        .fold(fz.bottom_y, f32::min);
                    // Clear the line after a trailing break even when its end
                    // already extends below the table: its start can still be
                    // inside the table's vertical band.
                    let trailing_clear = fz.from_table
                        && !empty_clear
                        && p.runs.last().is_some_and(|r| r.is_line_break)
                        && p.runs.iter().filter(|r| r.is_line_break).count() == 1
                        && p.runs.iter().all(|r| r.inline_image.is_none())
                        && p.image.is_none();
                    if state.pb.slot_top > float_bottom
                        || (fz.from_table && extra_breaks > 0)
                        || (trailing_clear && state.pb.slot_top + mark_pitch > fz.bottom_y)
                    {
                        // The line following the break resumes below the
                        // float and still occupies its full line height
                        // there (the break paragraph's mark line).
                        // A continuous section mark terminates this paragraph;
                        // unlike an ordinary clearing paragraph, its mark has
                        // no additional line below the float in Word.
                        let section_mark = ctx.compat_mode < 15
                            && fz.from_table
                            && empty_clear
                            && p.is_section_break
                            && block_idx > 0
                            && doc.sections.get(sect_idx + 1).is_some_and(|next| {
                                next.properties.break_type == SectionBreakType::Continuous
                            });
                        let mark_lines = if section_mark { 0.0 } else { mark_pitch };
                        state.pb.slot_top = float_bottom - mark_lines - break_lines;
                    }
                    state.pb.float_zone = None;
                    state.pb.more_float_zones.clear();
                }
                // Clear float zone once cursor passes below it
                if let Some(ref fz) = state.pb.float_zone
                    && state.pb.slot_top <= fz.bottom_y
                {
                    state.pb.next_float_zone();
                }

                last_block_consumption = (block_start.0 == state.pb.page_count())
                    .then_some((block_start.0, block_start.1 - state.pb.slot_top));
                state.global_block_idx += 1;
            }
            if let Some(f) = frame {
                f.close(state, sp);
            }
            if let Some(f) = open_flow_frame {
                f.close(state);
            }
        };

        // Word balances the columns of a region that ends at a continuous
        // section break, unless it holds a column break (probe: column 2 then
        // runs to the page foot and the next section moves to page 2).
        let balance = col_count > 1
            && doc
                .sections
                .get(sect_idx + 1)
                .is_some_and(|s| s.properties.break_type == SectionBreakType::Continuous)
            && !section.blocks.iter().any(|b| {
                matches!(b, Block::Paragraph(p) if p.column_break_before || p.column_break_after)
            });
        // A region spanning pages is balanced on its last page (Word probes:
        // a two-page region splits its last page 10/10 lines). `flushes` is
        // how many pages it fills before that one at full height.
        let base = state.pb.page_count().min(1);
        let trial_at = |state: &LayoutState, flushes: usize, bottom: f32| {
            let mut trial = state.trial(state.effective_margin_bottom);
            if flushes == 0 {
                trial.effective_margin_bottom = bottom;
            } else {
                trial.pb.balance_floor = Some((base + flushes, bottom));
            }
            layout_blocks(&mut trial);
            // The region's content height on its last page, as one stream,
            // with the space before a page-top heading dropped (case81 p6).
            let total = trial.pb.col_heights + trial.pb.column_top_y
                - (trial.pb.slot_top - trial.prev_space_after)
                + trial.pb.top_suppressed;
            (trial.pb.page_count() - base, total, trial.frame_overflow)
        };
        let natural = balance.then(|| trial_at(&state, 0, state.effective_margin_bottom));
        // ponytail: every trial lays the whole region out again, so a long
        // region costs a few full layouts; start trials at its last page if slow.
        let balanced_h = if let Some((flushes, total, _)) = natural {
            let (floor_page, page_bottom, top) = if flushes == 0 {
                (0, state.effective_margin_bottom, state.pb.column_top_y)
            } else {
                let page = state.pb.page_num(ctx.sections) + flushes;
                (
                    state.pb.page_count() + flushes,
                    compute_effective_margin_bottom(sp, false, page, &ctx),
                    effective_slot_top(sp, false, page, &ctx),
                )
            };
            // Word starts at the content height over the column count and adds
            // a line until it fits, which is not always the shortest fit (a
            // 3-column page split 14/14/12 lines where 14/13/13 fits).
            // A line must fit with all its leading here, unlike at the page
            // foot: case80's two-column region keeps its third paragraph whole
            // in column 2 though its first two lines' text fits column 1.
            let (step, lead) = section
                .blocks
                .iter()
                .find_map(|b| match b {
                    Block::Paragraph(p) if !is_text_empty(&p.runs) => {
                        let (fs, lhr, _) = tallest_run_metrics(&p.runs, ctx.fonts);
                        let ls = p.line_spacing.unwrap_or(ctx.doc_line_spacing);
                        let line_h = resolve_line_h(ls, fs, lhr);
                        Some((line_h, spacing_leading(line_h, fs, lhr)))
                    }
                    _ => None,
                })
                .unwrap_or((12.0, 0.0));
            let step = step.max(1.0);
            let mut h = total / col_count as f32;
            while h < top - page_bottom {
                let (pages, _, overflow) = trial_at(&state, flushes, top - h + lead + 0.01);
                if pages == flushes && !overflow {
                    break;
                }
                h += step;
            }
            let lo = if h < top - page_bottom {
                top - h + lead + 0.01
            } else {
                page_bottom
            };
            if flushes == 0 {
                state.effective_margin_bottom = lo;
            } else {
                state.pb.balance_floor = Some((floor_page, lo));
            }
            layout_blocks(&mut state);
            state.pb.balance_floor = None;
            // Keep footnote space booked while the region was laid out.
            state.effective_margin_bottom += page_bottom - lo;
            Some(h.min(top - page_bottom))
        } else {
            layout_blocks(&mut state);
            None
        };

        if col_count > 1 {
            // What follows a column region starts below its deepest column,
            // the last one's final space after included; in a balanced region
            // that space reaches no lower than the balancing height (Word
            // probes: one paragraph over three columns ends 3pt above its
            // space after; a lone two-line tail keeps all of it).
            let mut last = state.pb.slot_top - state.prev_space_after;
            if let Some(h) = balanced_h {
                last = last.max(state.pb.column_top_y - h).min(state.pb.slot_top);
            }
            let bottom = state.pb.col_bottom.min(last);
            state.pb.push_col_seps(bottom);
            state.pb.slot_top = bottom;
            state.prev_space_after = 0.0;
        }
        state.pb.region_sep_xs.clear();
    }
    let last_sect = doc.sections.len() - 1;
    state.pb.flush_page(last_sect);
    if !state.pb.footnote_ids.is_empty() {
        // Notes carried past the document's end fill a page of their own
        // under the first-page header; endnotes start it at its top under the
        // default one (Word examples: a 600pt row with a footnote
        // ends the document, with and without an endnote).
        let first = state.pb.endnote_ids.is_empty();
        let sp = &doc.sections[last_sect].properties;
        let top = effective_slot_top(sp, first, state.pb.page_num(ctx.sections), &ctx);
        state.pb.slot_top = top;
        state.prev_space_after = 0.0;
        state.pb.is_first_page_of_section = first;
        state.pb.flush_page(last_sect);
    }
    // For §17.6.23 vAlign centering, Word's content box includes the trailing
    // space_after of the last paragraph, which `slot_top` (and thus the recorded
    // content bottom) excludes. Extend the last page's content bottom by it so
    // the centered block matches Word's vertical position.
    if let Some(last) = state.pb.all_content_bottom.last_mut() {
        *last -= state.prev_space_after;
    }

    let t_layout = t0.elapsed();

    // Phase 2b: column separator lines
    for (content, seps) in state.pb.all_contents.iter_mut().zip(&state.pb.all_col_seps) {
        for (xs, top, bottom) in seps {
            for &x in xs {
                // Word's separator is a 0.75pt bar.
                stroke_segment(content, (x, *bottom), (x, *top), 0.75, None);
            }
        }
    }

    // Phase 2b': change bars (markup view). Word draws them in the left
    // margin (the outside margin with mirrored margins), half an inch from
    // the text but no deeper than halfway into the margin, 0.72pt wide, on
    // its 0.48pt device grid (Word for Mac probes 2026-11-05: margins of
    // 0.5, 1, 1.25 and 2 inches; two columns still bar at the page's left).
    for (page_idx, (content, bars)) in state
        .pb
        .all_contents
        .iter_mut()
        .zip(&state.pb.all_change_bars)
        .enumerate()
    {
        let (.., si) = state.pb.page_section_indices[page_idx];
        draw_change_bars(
            content,
            bars,
            &doc.sections[si].properties,
            doc.mirror_margins,
            page_idx,
        );
    }

    // Phase 2c: render footnotes at page bottom (above footer). Endnotes
    // (default pos=docEnd) flow inline after the last body block on the final
    // page (see render_endnotes_inline), NOT pinned to the bottom: their
    // separator paragraph starts at the cursor below the last paragraph and
    // its space_after.
    let last_page_idx = state.pb.all_contents.len().saturating_sub(1);
    let endnote_top_y = state.pb.slot_top - state.prev_space_after;
    let page_numbers = page_number_formats(doc, &state.pb.page_section_indices);
    for (page_idx, content) in state.pb.all_contents.iter_mut().enumerate() {
        let (hf_si, is_first, si) = state.pb.page_section_indices[page_idx];
        let sp = &doc.sections[hf_si].properties;
        let eff_bottom =
            compute_effective_margin_bottom(sp, is_first, page_numbers[page_idx].0, &ctx);
        let content_sp = &doc.sections[si].properties;
        let text_width = content_sp.text_width();
        let bottom = eff_bottom;
        // One block per column, at its foot and in its width.
        let ids = &state.pb.all_footnote_ids[page_idx];
        let cols = &state.pb.all_footnote_cols[page_idx];
        let mut tops = Vec::new();
        let mut done: Vec<(f32, f32)> = Vec::new();
        for &col in cols {
            if done.contains(&col) {
                continue;
            }
            done.push(col);
            let col_ids: Vec<u32> = ids
                .iter()
                .zip(cols)
                .filter(|&(_, &c)| c == col)
                .map(|(&id, _)| id)
                .collect();
            tops.extend(render_page_footnotes(
                content,
                &col_ids,
                &doc.footnotes,
                &footnote_display_order,
                &ctx,
                col.0,
                bottom,
                col.1,
                &mut state.pb.all_gradient_specs[page_idx],
                tagging::NoteTagger {
                    tags: &mut state.pb.tags,
                    page: page_idx,
                    endnote: false,
                    links: &mut state.pb.all_links[page_idx],
                },
            ));
        }
        for (id, y) in tops {
            state
                .bookmark_positions
                .insert(footnotes::note_anchor(false, id), (page_idx, y));
        }
        // Footnotes render after the body's bars were drawn; theirs would
        // otherwise all fall to the first page's header pass.
        let note_bars = layout::take_change_bars();
        draw_change_bars(
            content,
            &note_bars,
            content_sp,
            doc.mirror_margins,
            page_idx,
        );
        if page_idx == last_page_idx && !state.pb.endnote_ids.is_empty() {
            let tops = render_endnotes_inline(
                content,
                endnote_top_y,
                &state.pb.endnote_ids,
                &doc.endnotes,
                &endnote_display_order,
                &ctx,
                content_sp.margin_left,
                text_width,
                &mut state.pb.all_gradient_specs[page_idx],
                tagging::NoteTagger {
                    tags: &mut state.pb.tags,
                    page: page_idx,
                    endnote: true,
                    links: &mut state.pb.all_links[page_idx],
                },
            );
            for (id, y) in tops {
                state
                    .bookmark_positions
                    .insert(footnotes::note_anchor(true, id), (page_idx, y));
            }
        }
    }

    let t_headers = t0.elapsed();

    // Phase 2d: render headers/footers into separate content streams (behind body)
    let total_pages = state.pb.all_contents.len();

    // Pre-index header/footer image maps by (section_index, hf_type)
    // Fields: (para_images, inline_images, floating_images, effect_para)
    type HfMaps = (
        HashMap<usize, String>,
        HashMap<(usize, usize), String>,
        HashMap<(usize, usize), String>,
        HashMap<usize, EffectXObjs>,
    );
    let mut hf_maps_index: HashMap<(usize, u8), HfMaps> = HashMap::new();
    for ((s, t, pi), name) in &images.hf_image_names {
        hf_maps_index
            .entry((*s, *t))
            .or_default()
            .0
            .insert(*pi, name.clone());
    }
    for ((s, t, pi, ri), name) in &images.hf_inline_image_names {
        hf_maps_index
            .entry((*s, *t))
            .or_default()
            .1
            .insert((*pi, *ri), name.clone());
    }
    for ((s, t, pi, fi), name) in &images.hf_floating_image_names {
        hf_maps_index
            .entry((*s, *t))
            .or_default()
            .2
            .insert((*pi, *fi), name.clone());
    }
    for ((s, t, pi), fx) in &images.effect_hf_names {
        hf_maps_index
            .entry((*s, *t))
            .or_default()
            .3
            .insert(*pi, fx.clone());
    }
    let empty_hf_maps: HfMaps = Default::default();

    let section_page_counts =
        section_page_counts(doc.sections.len(), &state.pb.page_section_indices);

    let empty_styleref: HashMap<String, String> = HashMap::new();
    // The first occurrence of each style in the document. The running values
    // hold every style seen up to a page, so a style missing from them and
    // from the page itself first appears here.
    let mut styleref_doc_first: HashMap<String, String> = HashMap::new();
    for first in &state.pb.all_first_styleref {
        for (k, v) in first {
            if !styleref_doc_first.contains_key(k) {
                styleref_doc_first.insert(k.clone(), v.clone());
            }
        }
    }
    let mut page_styleref_merged: HashMap<String, String> = HashMap::new();
    let mut all_hf_contents: Vec<Option<Content>> = (0..total_pages).map(|_| None).collect();
    for (page_idx, hf_content) in all_hf_contents.iter_mut().enumerate() {
        let (si, is_first, _content_si) = state.pb.page_section_indices[page_idx];
        if state.pb.filler_pages.contains(&page_idx) {
            continue;
        }
        let sp = &doc.sections[si].properties;
        ctx.cell_grid_pitch.set(cell_grid_pitch(doc, sp));
        ctx.page_height.set(sp.page_height);

        let (page_num, effective_page_num_format) = page_numbers[page_idx];
        // Word counts the pages of the section the page starts in.
        let section_pages = section_page_counts[si];

        // Per spec §17.16.5.59: in headers/footers of a printed document, STYLEREF
        // searches the current page top-to-bottom first, then backward to doc start.
        let page_first = state
            .pb
            .all_first_styleref
            .get(page_idx)
            .unwrap_or(&empty_styleref);
        let prev_running = if page_idx > 0 {
            state
                .pb
                .all_styleref
                .get(page_idx - 1)
                .unwrap_or(&empty_styleref)
        } else {
            &empty_styleref
        };
        // Failing both, Word searches forward to the end: a contents page's
        // running head names the act whose title paragraph comes later.
        page_styleref_merged.clone_from(&styleref_doc_first);
        page_styleref_merged.extend(prev_running.iter().map(|(k, v)| (k.clone(), v.clone())));
        // Current-page first occurrences take priority (top-to-bottom search)
        for (k, v) in page_first {
            page_styleref_merged.insert(k.clone(), v.clone());
        }
        let page_styleref = &page_styleref_merged;

        let mut hf = Content::new();
        let mut has_hf = false;

        // Header first: the content stream order is part of the output.
        for (resolved, is_header) in [
            (resolve_header_for_page(doc, si, is_first, page_num), true),
            (resolve_footer_for_page(doc, si, is_first, page_num), false),
        ] {
            let (Some(data), hf_type, hf_si) = resolved else {
                continue;
            };
            let (pi_map, ii_map, fi_map, sh_para) = hf_maps_index
                .get(&(hf_si, hf_type))
                .unwrap_or(&empty_hf_maps);
            let pc = HfPageContext {
                page_num,
                total_pages,
                section_pages,
                para_image_names: pi_map,
                inline_image_names: ii_map,
                floating_image_names: fi_map,
                effect_para_names: sh_para,

                styleref_values: page_styleref,
                page_num_format: effective_page_num_format,
            };
            render_header_footer(
                &mut hf,
                data,
                &ctx,
                sp,
                is_header,
                &pc,
                &mut state.pb.all_gradient_specs[page_idx],
                &mut state.pb.all_links[page_idx],
            );
            has_hf = true;
        }

        // The header's and footer's own revised lines bar at the same x
        // (Word for Mac probe p09: bars beside the header's and footer's
        // inserted text at 54pt, like the body's).
        let hf_bars = layout::take_change_bars();
        if !hf_bars.is_empty() {
            draw_change_bars(&mut hf, &hf_bars, sp, doc.mirror_margins, page_idx);
            has_hf = true;
        }
        if has_hf {
            *hf_content = Some(hf);
        }
    }

    // Links drawn in artifacts (headers and footers, repeated table header
    // rows) have no Link element, but PDF/UA puts every link annotation in
    // one (Word leaves them untagged): one per link, holding only its
    // annotations, after the body.
    for links in &mut state.pb.all_links {
        let untagged = |a: &LinkAnnotation, b: &LinkAnnotation| {
            a.node.is_none() && b.node.is_none() && a.url == b.url
        };
        for link in links.chunk_by_mut(untagged) {
            if link[0].node.is_none() {
                let node = state.pb.tags.add(tagging::ROOT, "Link");
                link.iter_mut().for_each(|l| l.node = Some(node));
            }
        }
    }

    // §17.6.23 w:vAlign — shift each page's body block down so it is centered
    // (or bottom-aligned) in the text region. `slack` is the empty space between
    // the content bottom and the bottom margin; center splits it, bottom takes
    // it all. `both` (justify) and `top` leave content where it flowed.
    let valign_offsets: Vec<f32> = (0..total_pages)
        .map(|page_idx| {
            let (_, is_first, content_si) = state.pb.page_section_indices[page_idx];
            let sp = &doc.sections[content_si].properties;
            let frac = match sp.vertical_align {
                PageVerticalAlign::Center => 0.5,
                PageVerticalAlign::Bottom => 1.0,
                _ => return 0.0,
            };
            let region_bottom =
                compute_effective_margin_bottom(sp, is_first, page_numbers[page_idx].0, &ctx);
            let slack = state.pb.all_content_bottom[page_idx] - region_bottom;
            if slack * frac > 0.01 {
                slack * frac
            } else {
                0.0
            }
        })
        .collect();

    assemble_pdf_pages(
        &mut pdf,
        &mut alloc,
        catalog_id,
        pages_id,
        &valign_offsets,
        state.pb.all_contents,
        state.pb.all_deferred_shapes,
        state.pb.all_behind_shapes,
        &mut all_hf_contents,
        &state.pb.all_links,
        &state.pb.all_comment_anchors,
        &state.pb.all_alpha_states,
        &state.pb.all_gradient_specs,
        &state.pb.page_section_indices,
        ctx.fonts,
        &font_order,
        &images.image_xobjects,
        doc,
        &state.bookmark_positions,
        &state.heading_entries,
        &state.pb.tags,
    );

    let t_assembly = t0.elapsed();

    log::info!(
        "Render phases: fonts={:.1}ms, images={:.1}ms, layout={:.1}ms, headers={:.1}ms, assembly={:.1}ms",
        t_fonts.as_secs_f64() * 1000.0,
        (t_images - t_fonts).as_secs_f64() * 1000.0,
        (t_layout - t_images).as_secs_f64() * 1000.0,
        (t_headers - t_layout).as_secs_f64() * 1000.0,
        (t_assembly - t_headers).as_secs_f64() * 1000.0,
    );

    let page_fields = BodyPageFields {
        pages: page_numbers
            .into_iter()
            .map(|(n, fmt)| (n, fmt.map(str::to_string)))
            .collect(),
        section_pages: section_page_counts,
        total: total_pages,
    };
    Ok((objstm::pack(pdf.finish()), page_fields, valign_offsets))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn exact_lines_fit_by_their_whole_box() {
        // Auto spacing: the 13.8pt text of a 25pt pitch may hang its leading
        // past the margin; exact spacing (whole box) and lines above
        // footnotes need the full pitch.
        assert_eq!(line_fit_height(false, Some(25.0), 25.0, 13.8), 13.8);
        assert_eq!(line_fit_height(false, None, 25.0, 13.8), 13.8);
        assert_eq!(line_fit_height(true, Some(25.0), 25.0, 13.8), 25.0);
        assert_eq!(line_fit_height(true, None, 25.0, 13.8), 25.0);
        // A pitch shorter than the text (an own-run line) is the line's size.
        assert_eq!(line_fit_height(false, Some(9.0), 25.0, 13.8), 9.0);
    }

    #[test]
    fn a_break_line_that_does_not_fit_leaves_a_blank_page() {
        use std::io::{Cursor, Write};
        let pages = |break_twips: u32| {
            let exact = |t: u32| {
                format!(
                    r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="{t}" w:lineRule="exact"/></w:pPr>"#
                )
            };
            // 25 exact 25pt lines leave 23pt of a 648pt body.
            let lines = (0..25)
                .map(|i| format!("<w:p>{}<w:r><w:t>Line {i}</w:t></w:r></w:p>", exact(500)))
                .collect::<String>();
            let xml = format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{lines}<w:p>{}<w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#,
                exact(break_twips)
            );
            let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
            zip.start_file(
                "word/document.xml",
                zip::write::SimpleFileOptions::default(),
            )
            .unwrap();
            zip.write_all(xml.as_bytes()).unwrap();
            let doc = crate::docx::parse_bytes(
                &zip.finish().unwrap().into_inner(),
                crate::RevisionView::Final,
            )
            .unwrap();
            render_pass(&doc, None, &[], &ConvertOptions::default())
                .unwrap()
                .1
                .total
        };
        assert_eq!(pages(500), 3); // p29a
        assert_eq!(pages(300), 2); // p29c
    }

    #[test]
    fn footnote_space_is_charged_to_the_line_holding_the_reference() {
        let line_refs = vec![vec![], vec![4, 5], vec![], vec![6]];
        let tracked = HashSet::from([4]);
        let (per_line, total) =
            per_line_footnote_extra(&line_refs, &[4, 5, 6, 7], &tracked, 12.0, |id| {
                id as f32 * 10.0
            });
        // 4 is already on the page; 5 opens the footnote area so it carries the
        // separator; 7 produced no chunk and is charged to the last line.
        assert_eq!(per_line, vec![0.0, 62.0, 0.0, 130.0]);
        assert_eq!(total, 192.0);
    }

    #[test]
    fn whole_line_fit_ignores_the_float_gap_below_it() {
        let line = TextLine {
            pitch: Some(25.0),
            gap_after: 30.0,
            ..Default::default()
        };
        let lines = [line];
        assert_eq!(
            lines_that_fit(&lines, &[0.0], 45.0, true, 25.0, 13.8, (10.0, 3.0)),
            1
        );
        assert_eq!(
            lines_that_fit(&lines, &[0.0], 24.0, true, 25.0, 13.8, (10.0, 3.0)),
            0
        );
    }

    #[test]
    fn keep_with_next_needs_an_unsplittable_paragraph_whole() {
        assert_eq!(lines_kept_together(false, || 5), 1);
        assert_eq!(lines_kept_together(true, || 1), 1);
        assert_eq!(lines_kept_together(true, || 3), 3);
        assert_eq!(lines_kept_together(true, || 4), 2);
    }

    #[test]
    fn footnote_free_paragraph_costs_nothing() {
        let (per_line, total) =
            per_line_footnote_extra(&[vec![], vec![]], &[], &HashSet::new(), 12.0, |_| 99.0);
        assert_eq!(per_line, vec![0.0, 0.0]);
        assert_eq!(total, 0.0);
    }
}
