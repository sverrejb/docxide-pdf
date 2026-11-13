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
mod images;
mod layout;
mod line_metrics;
mod list_label;
mod objstm;
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
    Block, Document, FieldCode, FrameProperties, LineNumberRestart, LineSpacing,
    PageVerticalAlign, Paragraph, ParagraphBorder, Run, SectionBreakType, SectionProperties,
    ShapeFill, ShapeGeometry, VRelativeFrom, VerticalPosition, WrapText, WrapType,
};

use crate::fonts::font_key;
use assembly::{HeadingEntry, assemble_pdf_pages};
use bands::{Band, FlowFrame, OpenFrame, add_bands, auto_frame_width, block_para, column_geometry, empty_frame_table_bands, first_line_reach, first_word_width, floating_table_band, flow_frame, follows_floating_table, frame_beside_text, frame_content_height, has_bands, hf_page_bands, inter_gap, is_empty_wrapping_frame, lays_out_nothing, leading_para, lifted_frame, open_section_first_page, page_anchored, paragraph_line_count, place_first_line, place_moved_first_line, prescan_new_page, prescan_page_bands, step_below_bands, step_below_existing_bands, step_lines_below_bands, take_page_break_before};
use color::{fill_rgb, stroke_border, stroke_segment};
use fonts::collect_and_register_fonts;
use footnotes::{footnote_height, render_endnotes_inline, render_page_footnotes};
use header_footer::{
    HfPageContext, compute_effective_margin_bottom, effective_slot_top, render_header_footer,
    resolve_footer_for_page, resolve_header_for_page,
};
pub(super) use helpers::resolve_line_h;
use helpers::{
    horizontal_border_ext, side_border_outsets,
    align_offset, auto_ascent_scale, between_space, border_band, border_box_indents, bottom_border,
    collect_paras, draw_horizontal_rule, joins_border_group, spacing_leading,
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
use line_metrics::{break_font, break_run_lhr, label_boosted_baseline_offset, label_boosted_line_h, mark_font, mark_line_h, mark_position_stretch, size_break_lines, size_mark_line, size_tab_mark_line, size_unsized_lines, unsized_line_font, unsized_line_metrics};
use list_label::{label_font_key, render_list_label};
use paragraph::render_paragraph_block;
use positioning::{
    FloatTableExtent, FloatZone, FloatingTablePos, MIN_EMPTY_STRIP, MIN_TABLE_SIDE_ROOM, blocking_textbox_band, draw_debug_wrap_overlay, first_line_region_x, hangs_below_anchor, image_band, image_blocks_band, keep_open, left_text_region, no_strip_beside, para_line_h, plain_line_h, poly_scanline, polygon_edges, rectangular_polygon_clearance, render_connector, render_floating_images, resolve_fi_x, right_text_region, side_room, side_strip, skip_blocked_lines, step_lines_below_floats, textbox_blocks_band, textbox_room_beside, textbox_wraps_beside, textbox_zones, wrap_room, wrap_side, wraps_in_column,
};
pub(super) use positioning::{resolve_fi_y_top, resolve_h_position};
use prepass::{body_page_field_text, body_runs, cell_grid_pitch, compute_bookmark_positions, document_lang, keeps_section_mark_line};
use smartart::draw_shape_path;
use table::render_table;
use textbox_render::render_single_textbox;

/// A frame's width: w:w, or its widest block within `avail_w`.
fn frame_width(props: &FrameProperties, blocks: &[Block], ctx: &RenderContext, avail_w: f32) -> f32 {
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
        props.height.max(frame_content_height(props, blocks, ctx, width))
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

/// A one-shot override of where the next paragraph's drawings hang from.
#[derive(Clone, Copy)]
pub(super) struct PendingAnchor {
    pub(super) top: f32,
    pub(super) kind: AnchorKind,
}

#[derive(Clone, Copy)]
pub(super) enum AnchorKind {
    /// Drawings on the line (inline canvases, inline text boxes) follow the
    /// line wherever it goes.
    Floats,
    /// Drawings on the line take it too.
    TableMoved,
    /// Drawings on and positioned relative to the line take it too (see
    /// `ParaAnchor::drawing_top`).
    LookAhead,
}

/// Where a paragraph's drawings hang from. Paragraph-relative offsets count
/// from below the previous paragraph's space after but above this one's own
/// space before: Word keeps indonesian_school's "Format 12" box in place for
/// before = 0, 10 and 200 twips, and case124's stamp below the section-break
/// paragraph's 8pt. Drawings on the line (inline canvases, inline text boxes)
/// count from the line top, below the space before.
#[derive(Clone, Copy)]
pub(super) struct ParaAnchor {
    float: f32,
    on_line: f32,
    /// The paragraph's top where it flowed, space before included, or the
    /// look-ahead float's anchor.
    flowed: f32,
    /// The highest a wrapping paragraph-relative picture may sit: the top
    /// margin under compat 15 (Word probes: -30pt from a first paragraph
    /// lands on the margin; compat 14 and older keep the -30pt).
    ceiling: f32,
}

impl ParaAnchor {
    /// `pending`: a floating table's one-shot override for its anchor
    /// paragraph; `line_top`: the first line's top, below the space before.
    fn new(
        pending: Option<PendingAnchor>,
        line_top: f32,
        own_before: f32,
        sp: &SectionProperties,
        ctx: &RenderContext,
    ) -> Self {
        let below_before = line_top + own_before;
        let (float, on_line, flowed) = match pending {
            None => (below_before, line_top, below_before),
            Some(PendingAnchor {
                top,
                kind: AnchorKind::Floats,
            }) => (top, line_top, below_before),
            Some(PendingAnchor {
                top,
                kind: AnchorKind::TableMoved,
            }) => (top, top, below_before),
            Some(PendingAnchor {
                top,
                kind: AnchorKind::LookAhead,
            }) => (top, top, top),
        };
        Self {
            float,
            on_line,
            flowed,
            // Compat 14 keeps a negative offset only down to the page edge
            // (Word probes p20c/p20d: -82.7pt from a page-top paragraph puts
            // the picture at y 0).
            ceiling: if ctx.compat_mode >= 15 {
                sp.page_height - sp.margin_top
            } else {
                sp.page_height
            },
        }
    }

    /// Where a drawing's vertical offset counts from. One placed relative to
    /// the line follows its paragraph below a band or floating table that
    /// pushed it, where a paragraph-relative one stays where the paragraph
    /// stood (Word examples, fixture line-anchor-follows-push: offsets -30, 0
    /// and 10pt; with 24pt before, the line one counts from above it). It
    /// follows the line below a floating table moved whole to the next page
    /// too, but keeps a look-ahead float's anchor like a paragraph-relative
    /// one (fixture line-anchor-overrides).
    pub(super) fn drawing_top(self, on_line: bool, v_from_line: bool) -> f32 {
        if on_line {
            self.on_line
        } else if v_from_line {
            self.flowed
        } else {
            self.float
        }
    }

    /// A body float's top (`resolve_fi_y_top`) under the anchor's ceiling.
    pub(super) fn fi_y_top(self, fi: &crate::model::FloatingImage, sp: &SectionProperties) -> f32 {
        let y = resolve_fi_y_top(fi, sp, self.drawing_top(fi.on_line, fi.v_from_line));
        let wrapping_para = fi.wrap_type != WrapType::None
            && matches!(fi.v_relative_from, VRelativeFrom::Paragraph);
        // A wrapping picture taller than the page body fits no page: Word
        // keeps it on its anchor's sheet, moved up to the page top (an
        // external corpus cover: a 797pt picture anchored -82.7pt above the
        // third paragraph of a 595pt landscape page prints at y 0, the title
        // wrapped beside it; its A4 sibling, where the picture fits, keeps
        // 18.85). One that fits a page is not moved: its anchor paragraph
        // goes to the next page instead (leading-anchor fixtures).
        let h = fi.image.display_height;
        let y = if wrapping_para
            && h > sp.page_height - sp.margin_top - sp.margin_bottom
            && y - h < 0.0
        {
            y.max(h.min(sp.page_height))
        } else {
            y
        };
        if wrapping_para {
            y.min(self.ceiling)
        } else {
            y
        }
    }
}

impl PageBuilder {
    /// The pending float anchor, when it is section block `block_idx`'s (or
    /// the next paragraph's, when no block was named).
    fn float_anchor_for(&self, block_idx: usize) -> Option<PendingAnchor> {
        self.pending_float_anchor
            .filter(|&(_, b)| b.is_none_or(|b| b == block_idx))
            .map(|(a, _)| a)
    }

    fn take_float_anchor_for(&mut self, block_idx: usize) -> Option<PendingAnchor> {
        let anchor = self.float_anchor_for(block_idx)?;
        self.pending_float_anchor = None;
        Some(anchor)
    }
}

pub(super) struct PageBuilder {
    /// The body table being rendered, as `(left, right)`: a page it opens
    /// starts below the header floats it would overlap.
    pub(super) table_extent: Option<(f32, f32)>,
    // Current page state
    pub(super) content: Content,
    pub(super) links: Vec<LinkAnnotation>,
    /// (comment_id, anchor_x, anchor_y) — anchor is the end of the last chunk
    /// covered by that comment on this page. Used by the comment-pane renderer
    /// to draw the connector line from highlighted phrase to callout.
    pub(super) comment_anchors: Vec<(u32, f32, f32, f32)>,
    pub(super) footnote_ids: Vec<u32>,
    footnote_ids_set: HashSet<u32>,
    /// Endnote IDs encountered across the whole document, in encounter order.
    /// Endnotes default to `pos=docEnd`, so they all render on the final page
    /// rather than on the page where the reference appears.
    pub(super) endnote_ids: Vec<u32>,
    endnote_ids_set: HashSet<u32>,
    pub(super) alpha_states: HashSet<u8>,
    pub(super) gradient_specs: Vec<GradientSpec>,

    // Cross-page running state
    styleref_running: HashMap<String, String>,
    styleref_page_first: HashMap<String, String>,

    // Layout position state
    pub(super) slot_top: f32,
    /// Y-position at which the current column on this page begins. For a
    /// continuous 2-column section that starts mid-page, both the left and
    /// right columns share this top-y so advancing from left to right returns
    /// to where the section started rather than the top of the page.
    pub(super) column_top_y: f32,
    /// Where this page's body starts: below a header taller than the top
    /// margin, not at the margin (`is_at_page_top`).
    pub(super) page_top_y: f32,
    pub(super) is_first_page_of_section: bool,
    /// A continuous section whose first block has not flowed yet.
    pending_section_first_page: Option<(usize, usize, f32)>,
    /// Section that owns the current page for header/footer purposes.
    /// For continuous section breaks, this stays as the section that started
    /// the page, not the section that continues mid-page.
    page_hf_section: usize,
    /// Floating table exclusion zone on this page; paragraph layout
    /// uses horizontal bounds to decide wrap-beside vs push-below.
    pub(super) float_zone: Option<FloatZone>,
    /// The other wrapping floats of the paragraph that set `float_zone`,
    /// next in line once the cursor passes it.
    pub(super) more_float_zones: Vec<FloatZone>,
    /// The last floating table's horizontal extent, kept after its zone is
    /// consumed or lost to a page break.
    pub(super) float_table_extent: Option<FloatTableExtent>,
    /// The bands of this page's lifted frames and blocking floats: body
    /// lines may not overlap them (see `OpenFrame`, `step_below_bands`).
    frame_bands: Vec<Band>,
    /// The (section, page) whose upcoming float bands `prescan_page_bands`
    /// collected.
    bands_prescanned: Option<(usize, usize)>,
    /// One-shot anchor override for the paragraph that follows a floating table
    /// which was pushed whole onto a fresh page. That paragraph (the table's
    /// vertAnchor="text" anchor) must position its paragraph-relative shapes
    /// from the top of the page body — where the anchor naturally flows — even
    /// though the flow cursor stays below the table so following text doesn't
    /// overlap it. Without this the shapes drop by the table's height.
    /// With it, the section block the anchor is for, when a look-ahead found
    /// its float more than one paragraph ahead; `None` means the next one.
    pub(super) pending_float_anchor: Option<(PendingAnchor, Option<usize>)>,
    next_text_below_picture: Option<f32>,
    /// Anchored shapes for this page, painted above the text layer sorted by
    /// relativeHeight: Word stacks floating shapes in z order regardless of
    /// which paragraph anchors them.
    pub(super) deferred_shapes: Vec<(u32, Content)>,
    /// Behind-text drawings for this page, painted under all of its body text
    /// (shading included) sorted by relativeHeight, whatever paragraph or
    /// table cell anchors them.
    pub(super) behind_shapes: Vec<(u32, Content)>,

    // Accumulated pages
    all_contents: Vec<Content>,
    /// An exact frame in the flow clips this page's content (`FlowFrame`).
    pub(super) frame_clip: bool,
    /// Per-page y of the cursor at flush time (bottom of the last body block).
    /// Used to compute the §17.6.23 `w:vAlign` center/bottom offset.
    all_content_bottom: Vec<f32>,
    pub(super) all_deferred_shapes: Vec<Vec<(u32, Content)>>,
    pub(super) all_behind_shapes: Vec<Vec<(u32, Content)>>,
    all_links: Vec<Vec<LinkAnnotation>>,
    pub(super) all_comment_anchors: Vec<Vec<(u32, f32, f32, f32)>>,
    all_footnote_ids: Vec<Vec<u32>>,
    /// Each booked footnote's column (x, width), parallel to `footnote_ids`:
    /// Word sets a multi-column section's footnotes at the foot of the column
    /// that cites them, in its width.
    footnote_cols: Vec<(f32, f32)>,
    all_footnote_cols: Vec<Vec<(f32, f32)>>,
    /// Footnote space booked in the current column.
    pub(super) col_fn_reserved: f32,
    /// Footnotes (id, column, height, separator) cited on this page with no
    /// room left for them: the next page books them.
    carried_footnotes: Vec<(u32, (f32, f32), Option<f32>, f32)>,
    /// §17.6.8 body line-number counter (0-based count of lines seen).
    pub(super) line_number_counter: u32,
    /// The current section's `lnNumType restart` is newPage (the default):
    /// the counter goes back to zero with every page flushed.
    pub(super) line_numbers_restart_each_page: bool,
    all_alpha_states: Vec<HashSet<u8>>,
    all_gradient_specs: Vec<Vec<GradientSpec>>,
    /// Pages inserted by odd/even section breaks; Word prints them without header or footer.
    filler_pages: Vec<usize>,
    /// Per-page tuples: (hf_section, is_first_page, content_section).
    /// hf_section: which section provides headers/footers.
    /// content_section: which section is being rendered (for page numbering, geometry).
    page_section_indices: Vec<(usize, bool, usize)>,
    /// Highest column index the current column region reached on this page;
    /// Word draws a column separator only up to the last column holding text.
    last_col: usize,
    /// Lowest point a finished column of the current region reached on this page.
    col_bottom: f32,
    /// Separator x positions of the current column region (empty without w:sep).
    pub(super) region_sep_xs: Vec<f32>,
    /// Separator lines on the current page: (x positions, top y, bottom y).
    col_seps: Vec<(Vec<f32>, f32, f32)>,
    all_col_seps: Vec<Vec<(Vec<f32>, f32, f32)>>,
    /// Per page, (bottom, top) of each revised line (markup view).
    all_change_bars: Vec<Vec<(f32, f32)>>,
    /// Summed heights of the current region's finished columns on this page.
    col_heights: f32,
    /// Space before dropped at the top of this page (balancing still counts it).
    top_suppressed: f32,
    /// (page index, y): columns on that page end no lower than y, to balance
    /// the last page of a region spanning pages.
    balance_floor: Option<(usize, f32)>,
    all_styleref: Vec<HashMap<String, String>>,
    all_first_styleref: Vec<HashMap<String, String>>,
    pub(super) tags: tagging::Tags,
    pub(super) lists: tagging::Lists,
    /// The open TOC element while consecutive "toc N" paragraphs are tagged.
    pub(super) toc: Option<usize>,
    /// A TOC field has begun and its entries haven't ended yet.
    toc_field: bool,
    /// Structure of the body table being rendered (see `table::render_table`).
    pub(super) table_tags: Option<tagging::TableTags>,
}

impl PageBuilder {
    /// The cursor left `float_zone` behind: the next of its paragraph's
    /// floats still below the cursor takes its place.
    fn next_float_zone(&mut self) {
        let cursor = self.slot_top;
        self.more_float_zones.retain(|z| z.bottom_y < cursor);
        self.float_zone =
            (!self.more_float_zones.is_empty()).then(|| self.more_float_zones.remove(0));
    }

    /// The wrapping floats a paragraph just placed replace the page's zones,
    /// its last one first.
    fn set_float_zones(&mut self, mut zones: Vec<FloatZone>) {
        if let Some(last) = zones.pop() {
            self.float_zone = Some(last);
            self.more_float_zones = zones;
        }
    }

    /// A paragraph moved below a band or a floating table: its floating
    /// paragraph-relative drawings still hang from where it stood, while inline
    /// ones go with its line.
    fn pin_float_anchor(&mut self, top: f32) {
        self.pending_float_anchor.get_or_insert((
            PendingAnchor {
                top,
                kind: AnchorKind::Floats,
            },
            None,
        ));
    }

    fn new(slot_top: f32) -> Self {
        PageBuilder {
            content: tagging::artifact_content(),
            links: Vec::new(),
            comment_anchors: Vec::new(),
            footnote_ids: Vec::new(),
            footnote_ids_set: HashSet::new(),
            endnote_ids: Vec::new(),
            endnote_ids_set: HashSet::new(),
            alpha_states: HashSet::new(),
            gradient_specs: Vec::new(),
            styleref_running: HashMap::new(),
            styleref_page_first: HashMap::new(),
            slot_top,
            column_top_y: slot_top,
            page_top_y: slot_top,
            is_first_page_of_section: true,
            pending_section_first_page: None,
            page_hf_section: 0,
            float_zone: None,
            more_float_zones: Vec::new(),
            float_table_extent: None,
            frame_bands: Vec::new(),
            bands_prescanned: None,
            pending_float_anchor: None,
            table_extent: None,
            next_text_below_picture: None,
            deferred_shapes: Vec::new(),
            behind_shapes: Vec::new(),
            all_contents: Vec::new(),
            frame_clip: false,
            all_content_bottom: Vec::new(),
            all_deferred_shapes: Vec::new(),
            all_behind_shapes: Vec::new(),
            all_links: Vec::new(),
            all_comment_anchors: Vec::new(),
            all_footnote_ids: Vec::new(),
            footnote_cols: Vec::new(),
            all_footnote_cols: Vec::new(),
            col_fn_reserved: 0.0,
            carried_footnotes: Vec::new(),
            line_number_counter: 0,
            line_numbers_restart_each_page: false,
            all_alpha_states: Vec::new(),
            all_gradient_specs: Vec::new(),
            filler_pages: Vec::new(),
            page_section_indices: Vec::new(),
            last_col: 0,
            col_bottom: f32::INFINITY,
            region_sep_xs: Vec::new(),
            col_seps: Vec::new(),
            all_col_seps: Vec::new(),
            all_change_bars: Vec::new(),
            col_heights: 0.0,
            top_suppressed: 0.0,
            balance_floor: None,
            all_styleref: Vec::new(),
            all_first_styleref: Vec::new(),
            tags: tagging::Tags::new(),
            lists: tagging::Lists::default(),
            toc: None,
            toc_field: false,
            table_tags: None,
        }
    }

    /// Start `node`'s content on the current page (see `tagging`).
    fn begin_tag(&mut self, node: usize) {
        let page = self.all_contents.len();
        self.tags.begin(&mut self.content, page, node);
    }

    fn end_tag(&mut self) {
        tagging::Tags::end(&mut self.content);
    }

    /// Structure nodes for a body paragraph: (Lbl, element for its text). List
    /// items become LI > Lbl + LBody; numbered headings stay headings.
    fn para_tags(&mut self, para: &Paragraph, doc: &Document) -> (Option<usize>, usize) {
        // Word tags a table of contents as one flat TOC holding a TOCI per entry.
        let style_name = para
            .style_id
            .as_ref()
            .and_then(|id| doc.style_id_to_name.get(id));
        let is_toc_entry = style_name.is_some_and(|n| {
            n.get(..4).is_some_and(|p| p.eq_ignore_ascii_case("toc "))
                && n[4..].parse::<u8>().is_ok()
        });
        // Only inside a TOC field: its begin can sit in the entry or a heading above.
        self.toc_field |= para.starts_toc_field;
        if is_toc_entry && self.toc_field {
            self.lists.close();
            let toc = *self
                .toc
                .get_or_insert_with(|| self.tags.add(tagging::ROOT, "TOC"));
            return (None, self.tags.add(toc, "TOCI"));
        }
        self.toc = None;
        self.toc_field = para.starts_toc_field;
        self.tags.para_nodes(
            &mut self.lists,
            tagging::ROOT,
            para.list_item.filter(|_| para.outline_level.is_none()),
            !para.list_label.is_empty(),
            para_tag_kind(para, style_name),
        )
    }

    /// Tag a picture, chart or diagram paragraph the way Word does: an empty
    /// element for the paragraph mark, then the Figure hoisted beside it.
    fn begin_figure(&mut self, para: &Paragraph, doc: &Document, alt: Option<&str>) {
        self.tag_empty_para(para, doc);
        let figure = self.tags.add_figure(tagging::ROOT, alt);
        self.begin_tag(figure);
    }

    /// A chart or SmartArt paragraph as Word tags it: the paragraph mark, then
    /// a Figure whose drawing (labels included) stays an artifact, so only the
    /// alt text speaks for it.
    // ponytail: content-less Figure; tag the drawing itself if a validator asks for content
    fn figure_without_content(&mut self, para: &Paragraph, doc: &Document, alt: Option<&str>) {
        self.tag_empty_para(para, doc);
        self.tags.add_figure(tagging::ROOT, alt);
    }

    /// The paragraph's elements with nothing drawn inside.
    fn tag_empty_para(&mut self, para: &Paragraph, doc: &Document) {
        let mark = self.para_tags(para, doc);
        self.begin_para_tags(mark, |_| {});
        self.end_tag();
    }

    /// Draw the list label as its own Lbl (or inside the paragraph's element
    /// when it has none) and leave the paragraph text's tag open.
    fn begin_para_tags(
        &mut self,
        nodes: (Option<usize>, usize),
        draw_label: impl FnOnce(&mut Content),
    ) {
        let page = self.all_contents.len();
        self.tags
            .begin_para(&mut self.content, page, nodes, draw_label);
    }

    fn end_frame_clip(&mut self) {
        if std::mem::take(&mut self.frame_clip) {
            self.content.restore_state();
        }
    }

    pub(super) fn flush_page(&mut self, sect_idx: usize) {
        // A frame's clip ends with its page.
        self.end_frame_clip();
        self.all_contents.push(std::mem::replace(
            &mut self.content,
            tagging::artifact_content(),
        ));
        self.all_content_bottom.push(self.slot_top);
        // Stable sort: equal relativeHeight keeps document order
        self.deferred_shapes.sort_by_key(|(z, _)| *z);
        self.all_deferred_shapes
            .push(std::mem::take(&mut self.deferred_shapes));
        self.behind_shapes.sort_by_key(|(z, _)| *z);
        self.all_behind_shapes
            .push(std::mem::take(&mut self.behind_shapes));
        self.all_links.push(std::mem::take(&mut self.links));
        self.all_comment_anchors
            .push(std::mem::take(&mut self.comment_anchors));
        self.all_footnote_ids
            .push(std::mem::take(&mut self.footnote_ids));
        self.all_footnote_cols
            .push(std::mem::take(&mut self.footnote_cols));
        self.col_fn_reserved = 0.0;
        if self.line_numbers_restart_each_page {
            self.line_number_counter = 0;
        }
        self.footnote_ids_set.clear();
        let mut reserved = 0.0;
        for (id, col, height, separator_h) in std::mem::take(&mut self.carried_footnotes) {
            self.book_footnote(id, col, height, separator_h, &mut reserved);
        }
        self.all_alpha_states
            .push(std::mem::take(&mut self.alpha_states));
        self.all_gradient_specs
            .push(std::mem::take(&mut self.gradient_specs));
        self.page_section_indices.push((
            self.page_hf_section,
            self.is_first_page_of_section,
            sect_idx,
        ));
        self.push_col_seps(self.slot_top);
        self.all_col_seps.push(std::mem::take(&mut self.col_seps));
        self.all_change_bars.push(layout::take_change_bars());
        self.top_suppressed = 0.0;
        self.all_styleref.push(self.styleref_running.clone());
        self.all_first_styleref
            .push(std::mem::take(&mut self.styleref_page_first));
        self.float_zone = None;
        self.more_float_zones.clear();
        self.frame_bands.clear();
        // An anchor measured on the old page must not place a float on this one
        self.pending_float_anchor = None;
        self.next_text_below_picture = None;
        // After flush, the new page starts with the current section
        self.page_hf_section = sect_idx;
    }

    fn push_blank_page(&mut self, sect_idx: usize) {
        self.filler_pages.push(self.all_contents.len());
        self.all_contents.push(tagging::artifact_content());
        // Blank page has no body content; record top so vAlign yields no shift.
        self.all_content_bottom.push(self.slot_top);
        self.all_deferred_shapes.push(Vec::new());
        self.all_behind_shapes.push(Vec::new());
        self.all_links.push(Vec::new());
        self.all_comment_anchors.push(Vec::new());
        self.all_footnote_ids.push(Vec::new());
        self.all_footnote_cols.push(Vec::new());
        self.all_alpha_states.push(HashSet::new());
        self.all_gradient_specs.push(Vec::new());
        self.page_section_indices
            .push((self.page_hf_section, false, sect_idx));
        self.all_col_seps.push(Vec::new());
        self.all_change_bars.push(Vec::new());
        self.all_styleref.push(self.styleref_running.clone());
        self.all_first_styleref
            .push(std::mem::take(&mut self.styleref_page_first));
        self.page_hf_section = sect_idx;
    }

    fn page_count(&self) -> usize {
        self.all_contents.len()
    }

    /// The number the open page shows (`page_numbers`), which picks its even
    /// or odd header.
    pub(super) fn page_num(&self, sections: &[crate::model::Section]) -> usize {
        let open = (self.page_hf_section, false, self.page_hf_section);
        last_page_number(sections, self.page_section_indices.iter().chain([&open]))
    }

    /// Book footnote `id` (height None when it doesn't exist) into the foot
    /// of column `col` once per page, shrinking the column by its height and,
    /// for the column's first note, the separator.
    pub(super) fn book_footnote(
        &mut self,
        id: u32,
        col: (f32, f32),
        height: Option<f32>,
        separator_h: f32,
        effective_margin_bottom: &mut f32,
    ) {
        if !self.footnote_ids_set.insert(id) {
            return;
        }
        self.footnote_ids.push(id);
        self.footnote_cols.push(col);
        if let Some(h) = height {
            let sep = if self.col_fn_reserved == 0.0 {
                separator_h
            } else {
                0.0
            };
            *effective_margin_bottom += sep + h;
            self.col_fn_reserved += sep + h;
        }
    }

    /// Leave footnote `id` for the next page, as `book_footnote` takes it.
    pub(super) fn carry_footnote(
        &mut self,
        id: u32,
        col: (f32, f32),
        height: Option<f32>,
        separator_h: f32,
    ) {
        if !self.footnote_ids_set.contains(&id) {
            self.carried_footnotes.push((id, col, height, separator_h));
        }
    }

    /// Close the current column region's separators on this page: Word draws
    /// them from the region's top down to its deepest column (`bottom` is the
    /// column in progress), and none when the text never left column 1.
    fn push_col_seps(&mut self, bottom: f32) {
        let n = self.last_col.min(self.region_sep_xs.len());
        if n > 0 {
            let bottom = bottom.min(self.col_bottom);
            self.col_seps
                .push((self.region_sep_xs[..n].to_vec(), self.column_top_y, bottom));
        }
        self.last_col = 0;
        self.col_bottom = f32::INFINITY;
        self.col_heights = 0.0;
    }

    /// Endnotes render at the end of the document: collect the id once, in
    /// encounter order, for the last page (Phase 2c).
    fn track_endnote(&mut self, id: u32) {
        if self.endnote_ids_set.insert(id) {
            self.endnote_ids.push(id);
        }
    }

    /// A page opened (by a break) that nothing has been placed on yet.
    fn is_empty_page(&self) -> bool {
        self.page_count() > 0
            && (self.slot_top - self.page_top_y).abs() < 0.01
            && self.deferred_shapes.is_empty()
            && self.footnote_ids.is_empty()
    }

    /// Where a paragraph's floats anchor: its slot top, except that a page's
    /// first paragraph pushed down by header floats anchors where the header's
    /// text alone ends (Word probes under a 133pt floating letterhead: offsets
    /// 0 and +30pt land at the 56.7pt margin and 30pt below it, while the
    /// text sits under the letterhead; a header of flowing text moves both).
    fn footer_top(&self, sp: &SectionProperties, ctx: &RenderContext) -> f32 {
        header_footer::footer_top(
            sp,
            self.is_first_page_of_section,
            self.page_num(ctx.sections),
            ctx,
        )
    }

    fn anchor_slot_top(&self, sp: &SectionProperties, ctx: &RenderContext) -> f32 {
        match self.page_top_header(sp, ctx) {
            Some(is_first) => self.slot_top.max(header_footer::flow_slot_top(
                sp,
                is_first,
                self.page_num(ctx.sections),
                ctx,
            )),
            None => self.slot_top,
        }
    }

    /// Which header variant (first page or not) opened this page, when the
    /// cursor still sits at its top: the one whose body top matches.
    fn page_top_header(&self, sp: &SectionProperties, ctx: &RenderContext) -> Option<bool> {
        if !self.is_at_page_top(sp) {
            return None;
        }
        let page = self.page_num(ctx.sections);
        [true, false].into_iter().find(|&is_first| {
            (effective_slot_top(sp, is_first, page, ctx) - self.page_top_y).abs() < 0.01
        })
    }

    /// Where a table spanning `left..right` starts: a page-top table drops
    /// below the header's wrapping floats it would overlap, see
    /// `header_float_floor`.
    fn below_header_floats(
        &self,
        sp: &SectionProperties,
        ctx: &RenderContext,
        left: f32,
        right: f32,
    ) -> f32 {
        self.page_top_header(sp, ctx)
            .and_then(|is_first| {
                header_footer::header_float_floor(
                    sp,
                    is_first,
                    self.page_num(ctx.sections),
                    ctx,
                    left,
                    right,
                )
            })
            .map_or(self.slot_top, |floor| self.slot_top.min(floor))
    }

    fn is_at_page_top(&self, sp: &SectionProperties) -> bool {
        // nabl's 47.49pt header ends below its 36pt top margin: Word still
        // drops the space before the first paragraph under it.
        (self.slot_top - (sp.page_height - sp.margin_top)).abs() < 1.0
            || (self.slot_top - self.page_top_y).abs() < 0.01
    }

    /// A fresh page always takes a line, except a balanced one whose short
    /// columns must make the trial fail instead.
    fn takes_a_line(&self, sp: &SectionProperties) -> bool {
        self.is_at_page_top(sp)
            && self
                .balance_floor
                .is_none_or(|(page, _)| page != self.page_count())
    }

    /// The gap above a block at the top of a new page (None elsewhere): none,
    /// except on a section's first page, where the block's space before
    /// collapses with the previous section's trailing space after.
    fn page_top_gap(
        &self,
        sp: &SectionProperties,
        space_before: f32,
        prev_space_after: f32,
    ) -> Option<f32> {
        (!self.all_contents.is_empty() && self.is_at_page_top(sp)).then(|| {
            if self.is_first_page_of_section {
                own_space_before(space_before, prev_space_after)
            } else {
                0.0
            }
        })
    }

    /// Text that didn't fit: on to the next column, but when this column is
    /// still empty the others are no taller, so to the next page (case80's
    /// five-column region starts on page 2). The column's separator reaches
    /// the `trailing_space` after its last paragraph (case80 page 5).
    #[allow(clippy::too_many_arguments)]
    fn overflow_column_or_page(
        &mut self,
        current_col: &mut usize,
        col_count: usize,
        sect_idx: usize,
        sp: &SectionProperties,
        effective_margin_bottom: &mut f32,
        ctx: &RenderContext,
        trailing_space: f32,
    ) {
        let col_count = if self.slot_top >= self.column_top_y - 0.01 {
            1
        } else {
            self.slot_top -= trailing_space;
            col_count
        };
        self.advance_column_or_page(
            current_col,
            col_count,
            sect_idx,
            sp,
            effective_margin_bottom,
            ctx,
        );
    }

    /// Advance to the next column if available, otherwise flush the current page.
    fn advance_column_or_page(
        &mut self,
        current_col: &mut usize,
        col_count: usize,
        sect_idx: usize,
        sp: &SectionProperties,
        effective_margin_bottom: &mut f32,
        ctx: &RenderContext,
    ) {
        if *current_col + 1 < col_count {
            *current_col += 1;
            self.last_col = self.last_col.max(*current_col);
            self.col_bottom = self.col_bottom.min(self.slot_top);
            self.col_heights += self.column_top_y - self.slot_top;
            *effective_margin_bottom -= self.col_fn_reserved;
            self.col_fn_reserved = 0.0;
            self.slot_top = self.column_top_y;
            self.pending_float_anchor = None;
        } else {
            *current_col = 0;
            self.begin_next_page(sect_idx, sp, effective_margin_bottom, ctx);
            if let Some((page, y)) = self.balance_floor
                && page == self.page_count()
            {
                *effective_margin_bottom = effective_margin_bottom.max(y);
            }
        }
    }

    /// The text after a full-width picture resumes below the lowest one.
    fn lower_next_text_to(&mut self, bottom: f32) {
        self.next_text_below_picture =
            Some(self.next_text_below_picture.map_or(bottom, |old| old.min(bottom)));
    }

    fn take_pending_first_page(&mut self, sect_idx: usize) -> bool {
        self.pending_section_first_page
            .take()
            .is_some_and(|(section, page, y)| {
                section == sect_idx && page == self.page_count() && (self.slot_top - y).abs() < 0.01
            })
    }

    /// A table opening the section's pending first page moves the cursor
    /// from `from` to its position, space before and top border; nothing has
    /// flowed yet.
    pub(super) fn keep_pending_first_page(&mut self, from: f32) {
        if let Some((_, _, y)) = &mut self.pending_section_first_page
            && (*y - from).abs() < 0.01
        {
            *y = self.slot_top;
        }
    }

    /// Footnotes carried past the end of a section get a page of their own
    /// under the first-page header before the next section starts (Word
    /// examples: a 630pt row whose note does not fit ends section 1;
    /// the note fills page 2 alone under section 1's first header for a
    /// next-page break, section 2's for a continuous one, whose text starts
    /// page 3).
    fn flush_note_page(&mut self, sect_idx: usize) {
        if !self.footnote_ids.is_empty() {
            self.is_first_page_of_section = true;
            self.flush_page(sect_idx);
        }
    }

    /// Flush this page and open the next one of section `sp`.
    pub(super) fn begin_next_page(
        &mut self,
        sect_idx: usize,
        sp: &SectionProperties,
        effective_margin_bottom: &mut f32,
        ctx: &RenderContext,
    ) {
        let first_page = self.take_pending_first_page(sect_idx);
        // A section that moved off the page in full leaves it to the one before:
        // the page neither counts for it nor starts its numbering.
        let ends_in = if first_page { sect_idx - 1 } else { sect_idx };
        self.flush_page(ends_in);
        self.page_hf_section = sect_idx;
        if first_page {
            self.flush_note_page(sect_idx);
        }
        let page = self.page_num(ctx.sections);
        self.slot_top = effective_slot_top(sp, first_page, page, ctx);
        self.page_top_y = self.slot_top;
        if let Some((left, right)) = self.table_extent {
            self.slot_top = self.below_header_floats(sp, ctx, left, right);
        }
        self.column_top_y = self.slot_top;
        self.page_top_y = self.slot_top;
        *effective_margin_bottom =
            compute_effective_margin_bottom(sp, first_page, page, ctx) + self.col_fn_reserved;
        self.is_first_page_of_section = first_page;
        let bands = hf_page_bands(sp, first_page, page, *effective_margin_bottom, ctx);
        add_bands(&mut self.frame_bands, bands);
    }
}

/// Bundles the mutable render loop state so it can be passed to extracted functions.
pub(super) struct LayoutState {
    pub(super) pb: PageBuilder,
    pub(super) prev_space_after: f32,
    pub(super) prev_between_below: f32,
    pub(super) effective_margin_bottom: f32,
    pub(super) current_col: usize,
    pub(super) global_block_idx: usize,
    pub(super) heading_entries: Vec<HeadingEntry>,
    pub(super) bookmark_positions: HashMap<String, (usize, f32)>,
    /// The keep-with-next chain being laid out is too long for a page: its
    /// paragraphs keep only their own link to the next.
    pub(super) long_keep_chain: bool,
    /// The gap above a paragraph laid out again in the column it overflowed
    /// into: the overflow's, not its spacing recomputed.
    pub(super) reflowed: bool,
    /// The page's upcoming float bands are collected ahead of its lines
    /// (`prescan_page_bands`): a single-column section with bands, outside
    /// frames.
    pub(super) band_prescan: bool,
    /// The next paragraph's page break before was taken ahead of the page's
    /// float bands, which may have moved it off the page top since.
    pub(super) page_break_taken: bool,
    /// A frame in the flow runs past the bottom, or moved on to a column
    /// with more lines than it left in the one before: a balancing trial at
    /// this height fails (Word examples: a frame of four lines after one
    /// stays in column 1, fixture frame-balanced-columns; six lines after
    /// seven open column 2, frame-balanced-opens-column).
    pub(super) frame_overflow: bool,
}

impl LayoutState {
    fn advance_column_or_page(
        &mut self,
        col_count: usize,
        sect_idx: usize,
        sp: &SectionProperties,
        ctx: &RenderContext,
    ) {
        self.pb.advance_column_or_page(
            &mut self.current_col,
            col_count,
            sect_idx,
            sp,
            &mut self.effective_margin_bottom,
            ctx,
        );
    }

    fn begin_next_page(&mut self, sect_idx: usize, sp: &SectionProperties, ctx: &RenderContext) {
        self.pb
            .begin_next_page(sect_idx, sp, &mut self.effective_margin_bottom, ctx);
    }

    fn take_page_break_after(
        &mut self,
        para: &Paragraph,
        sect_idx: usize,
        sp: &SectionProperties,
        ctx: &RenderContext,
    ) {
        if para.page_break_after {
            self.begin_next_page(sect_idx, sp, ctx);
            self.prev_space_after = 0.0;
            self.current_col = 0;
        }
    }

    /// A throwaway copy for laying out a column region again with its
    /// columns ending at `bottom`; its output is discarded.
    fn trial(&self, bottom: f32) -> LayoutState {
        let mut pb = PageBuilder::new(self.pb.slot_top);
        pb.column_top_y = self.pb.column_top_y;
        pb.page_top_y = self.pb.page_top_y;
        pb.is_first_page_of_section = self.pb.is_first_page_of_section;
        pb.page_hf_section = self.pb.page_hf_section;
        pb.page_section_indices = self.pb.page_section_indices.clone();
        pb.float_zone = self.pb.float_zone.clone();
        pb.more_float_zones = self.pb.more_float_zones.clone();
        pb.pending_float_anchor = self.pb.pending_float_anchor;
        pb.next_text_below_picture = self.pb.next_text_below_picture;
        pb.footnote_ids_set = self.pb.footnote_ids_set.clone();
        pb.col_fn_reserved = self.pb.col_fn_reserved;
        pb.line_number_counter = self.pb.line_number_counter;
        pb.line_numbers_restart_each_page = self.pb.line_numbers_restart_each_page;
        // Page-top rules look at whether a page was already flushed.
        if !self.pb.all_contents.is_empty() {
            pb.all_contents.push(Content::new());
        }
        LayoutState {
            pb,
            prev_space_after: self.prev_space_after,
            prev_between_below: self.prev_between_below,
            effective_margin_bottom: bottom,
            current_col: self.current_col,
            global_block_idx: self.global_block_idx,
            heading_entries: Vec::new(),
            bookmark_positions: self.bookmark_positions.clone(),
            long_keep_chain: self.long_keep_chain,
            reflowed: false,
            band_prescan: self.band_prescan,
            page_break_taken: false,
            frame_overflow: false,
        }
    }
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

/// Each page's number and number format (None = decimal), each section's
/// page count, and the page count.
pub(super) struct BodyPageFields {
    pages: Vec<(usize, Option<String>)>,
    section_pages: Vec<usize>,
    total: usize,
}

/// How many pages each section is on (SECTIONPAGES). A page where a
/// continuous section starts counts for both sections.
fn section_page_counts(sections: usize, pages: &[(usize, bool, usize)]) -> Vec<usize> {
    let mut counts = vec![0; sections];
    for &(first, _, last) in pages {
        for count in &mut counts[first..=last] {
            *count += 1;
        }
    }
    counts
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
                    open_flow_frame =
                        Some(FlowFrame::open(fp, blocks, state, &ctx, sp, (sect_idx, &col_geometry)));
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
                            frame = Some(OpenFrame::open(fp, blocks, state, &ctx, sp, (sect_idx, &col_geometry)));
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
    fn continuous_first_page_is_pending_only_before_the_first_block_flows() {
        let mut page = PageBuilder::new(97.9);
        page.pending_section_first_page = Some((2, 0, 97.9));
        assert!(page.take_pending_first_page(2));
        assert!(!page.take_pending_first_page(2));
        page.pending_section_first_page = Some((2, 0, 97.9));
        page.slot_top -= 16.0; // A first paragraph has flowed on this sheet.
        assert!(!page.take_pending_first_page(2));
        page.slot_top = 97.9;
        page.pending_section_first_page = Some((2, 0, 97.9));
        assert!(!page.take_pending_first_page(3));
        page.pending_section_first_page = Some((2, 0, 97.9));
        page.all_contents.push(Content::new());
        assert!(!page.take_pending_first_page(2));
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

/// The number the last of `page_section_indices` shows (`page_numbers`), 1
/// for none.
fn last_page_number<'a>(
    sections: &[crate::model::Section],
    page_section_indices: impl IntoIterator<Item = &'a (usize, bool, usize)>,
) -> usize {
    page_numbers(sections, page_section_indices)
        .last()
        .copied()
        .unwrap_or(1)
}

/// Logical page number of each page: a section with `w:pgNumType @start` restarts,
/// others continue. A restart after an odd/even section break with the wrong
/// parity takes the next number instead (Word skips a number, not a sheet).
/// A page shows the number of the section at its top; a restart in a section
/// that begins further down counts that page as its start, so the next page
/// shows start + 1 (Word probes, 2026-10-03).
fn page_numbers<'a>(
    sections: &[crate::model::Section],
    page_section_indices: impl IntoIterator<Item = &'a (usize, bool, usize)>,
) -> Vec<usize> {
    let mut numbers = Vec::new();
    let mut running = 0;
    let mut prev_si = None;
    for &(si, _, last_si) in page_section_indices {
        let sp = &sections[si].properties;
        running = match sp.page_num_start {
            Some(start) if prev_si != Some(si) => {
                let start = start as usize;
                let wrong_parity = si > 0
                    && match sp.break_type {
                        SectionBreakType::OddPage => start % 2 == 0,
                        SectionBreakType::EvenPage => start % 2 == 1,
                        _ => false,
                    };
                start + wrong_parity as usize
            }
            _ => running + 1,
        };
        numbers.push(running);
        if let Some(start) = (si + 1..=last_si)
            .filter_map(|s| sections[s].properties.page_num_start)
            .last()
        {
            running = start as usize;
        }
        prev_si = Some(last_si);
    }
    numbers
}

/// Each page's logical number and number format (None = decimal): the format
/// of the section a page ends in, which never inherits from the one before
/// (OOXML §17.6.12; Word renders arabic after roman front matter).
fn page_number_formats<'a>(
    doc: &'a Document,
    page_section_indices: &[(usize, bool, usize)],
) -> Vec<(usize, Option<&'a str>)> {
    page_numbers(&doc.sections, page_section_indices)
        .into_iter()
        .zip(page_section_indices)
        .map(|(n, &(_, _, si))| (n, doc.sections[si].properties.page_num_format.as_deref()))
        .collect()
}

/// The docGrid pitch table-cell lines snap to in a section: only under
/// `w:compat/w:adjustLineHeightInTable`, else 0.
/// How an empty, zero-height section-break paragraph passes spacing on: the
/// drop below the cursor and the space after left pending. Before a
/// continuous section the paragraph before it keeps its whole space after,
/// and the break's own space after only absorbs that much of the next
/// paragraph's space before (Word probes: 12pt after, 8pt on the break, 10pt
/// before give 14pt; covid_insomnia's two columns start 12pt below its
/// keywords). Before a new page the break's own space after counts.
fn section_break_spacing(prev_after: f32, break_after: f32, next_continuous: bool) -> (f32, f32) {
    if next_continuous {
        (prev_after - break_after, break_after)
    } else {
        (0.0, break_after)
    }
}
