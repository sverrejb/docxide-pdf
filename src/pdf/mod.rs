mod assembly;
mod chart_legend;
mod charts;
mod charts_radial;
mod comments;
mod emf;
pub(crate) mod color;
mod fonts;
mod footnotes;
mod header_footer;
mod helpers;
mod images;
mod list_label;
mod layout;
mod positioning;
mod smartart;
mod table;
mod table_layout;
mod textbox_render;
mod wordart;

use std::collections::{HashMap, HashSet};

use pdf_writer::{Content, Name, Pdf, Ref};

use crate::error::Error;
use crate::fonts::FontEntry;
use crate::model::{
    Alignment, Block, DocGridType, Document, FieldCode,
    HorizontalPosition, LineSpacing, Paragraph, ParagraphBorder,
    Run, SectionBreakType, SectionProperties, ShapeFill, ShapeGeometry,
    VRelativeFrom, VerticalPosition, WrapText, WrapType,
};

use assembly::{HeadingEntry, assemble_pdf_pages};
use fonts::collect_and_register_fonts;
use footnotes::{compute_footnote_height, render_page_footnotes};
use header_footer::{
    HfPageContext, compute_effective_margin_bottom, effective_slot_top, render_header_footer,
    resolve_footer_for_page, resolve_header_for_page,
};
pub(super) use helpers::resolve_line_h;
use helpers::borders_match;
use positioning::{render_connector, render_floating_images, resolve_fi_x};
pub(super) use positioning::{resolve_h_position, resolve_fi_y_top};
use images::{EffectXObjs, EmbeddedImages, embed_all_images};
use layout::{
    DualRegion, LinkAnnotation, build_paragraph_lines, build_tabbed_line, is_text_empty,
    render_paragraph_lines, tallest_run_metrics,
};
use crate::fonts::font_key;
use color::{fill_rgb, stroke_rgb};
use list_label::{collect_paras, label_font_key, para_runs_with_textboxes, render_list_label};
use smartart::draw_shape_path;
use table::render_table;
use textbox_render::render_single_textbox;

/// Word stacks overlapping anchored shapes by wp:anchor relativeHeight, not
/// document order. Stable sort keeps document order for equal values.
fn sorted_by_z<'a>(iter: impl Iterator<Item = &'a crate::model::Textbox>) -> Vec<&'a crate::model::Textbox> {
    let mut v: Vec<&crate::model::Textbox> = iter.collect();
    v.sort_by_key(|t| t.z_index);
    v
}

pub(super) struct RenderContext<'a> {
    pub(super) fonts: &'a HashMap<String, FontEntry>,
    pub(super) doc_line_spacing: LineSpacing,
    pub(super) default_tab_stop: f32,
    /// Image names for inline images in table cells, keyed by Arc data pointer address.
    pub(super) table_cell_image_names: &'a HashMap<usize, String>,
    pub(super) effect_table_names: &'a HashMap<usize, EffectXObjs>,
    /// Image names for images inside textbox paragraphs, keyed by Arc data pointer address.
    pub(super) textbox_image_names: &'a HashMap<usize, String>,
    pub(super) effect_textbox_names: &'a HashMap<usize, EffectXObjs>,
    pub(super) chart_font_name: &'a str,
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

/// Compute line height from the list label font if it exceeds the text-run line height.
/// Word includes the numbering label character's font metrics in the tallest-font
/// calculation, so a bullet from Symbol font can make the line taller than text-only Calibri.
/// Only applied when label and text use the same font size (the common bullet case);
/// when the label has a dramatically different size it's a special numbering style
/// handled separately via the existing label_fs > font_size path.
fn label_boosted_line_h(
    para: &Paragraph,
    fonts: &HashMap<String, FontEntry>,
    text_line_h: f32,
    effective_ls: LineSpacing,
    text_font_size: f32,
) -> f32 {
    if para.list_label.is_empty() {
        return text_line_h;
    }
    let label_fs = para.list_label_font_size.unwrap_or(text_font_size);
    // Only boost when the label font size matches the text font size (common bullet case).
    // Large label sizes (e.g. 20pt label on 10pt text) are handled separately.
    if (label_fs - text_font_size).abs() > 1.0 {
        return text_line_h;
    }
    let Some(key) = label_font_key(para) else {
        return text_line_h;
    };
    let Some(entry) = fonts.get(&key) else {
        return text_line_h;
    };
    let label_lh = resolve_line_h(effective_ls, label_fs, entry.line_h_ratio);
    text_line_h.max(label_lh)
}

/// Look up the line_h_ratio for a break run's font, matching by font_size.
fn break_run_lhr(
    runs: &[Run],
    break_fs: f32,
    fonts: &HashMap<String, FontEntry>,
) -> Option<f32> {
    // Find the break run with the matching font size
    let br_run = runs.iter()
        .filter(|r| r.is_line_break && (r.font_size - break_fs).abs() < 0.01)
        .last();
    if let Some(run) = br_run {
        let key = font_key(run);
        fonts.get(&key).and_then(|e| e.line_h_ratio)
    } else {
        None
    }
}

fn styleref_insert(
    map: &mut HashMap<String, String>,
    id: &str,
    text: &str,
    style_id_to_name: &HashMap<String, String>,
) {
    map.insert(id.to_string(), text.to_string());
    if let Some(name) = style_id_to_name.get(id) {
        map.insert(name.clone(), text.to_string());
    }
}

fn styleref_insert_first(
    map: &mut HashMap<String, String>,
    id: &str,
    text: &str,
    style_id_to_name: &HashMap<String, String>,
) {
    map.entry(id.to_string())
        .or_insert_with(|| text.to_string());
    if let Some(name) = style_id_to_name.get(id) {
        map.entry(name.clone()).or_insert_with(|| text.to_string());
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
            styleref_insert(running, sid, &text, style_id_to_name);
            styleref_insert_first(page_first, sid, &text, style_id_to_name);
        }
    }
    for run in &para.runs {
        if let Some(ref csid) = run.char_style_id {
            if !run.text.is_empty() {
                styleref_insert(running, csid, &run.text, style_id_to_name);
                styleref_insert_first(page_first, csid, &run.text, style_id_to_name);
            }
        }
    }
}

pub(super) struct FloatZone {
    pub top_y: f32,
    pub bottom_y: f32,
    pub obj_left: f32,
    pub obj_right: f32,
    pub left_from_text: f32,
    pub right_from_text: f32,
    /// Polygon vertices in absolute page coords (PDF: x from left, y from bottom)
    pub polygon_pts: Option<Vec<(f32, f32)>>,
    pub wrap_text: WrapText,
    /// True when the zone was created by a paragraph-relative floating image
    /// (positionV relativeFrom="paragraph").  Paragraphs whose cursor is
    /// slightly above the zone should still be pushed below wide images.
    pub para_relative: bool,
}

impl FloatZone {
    /// Returns (left_edge, right_edge) of the exclusion zone at the given Y.
    /// Falls back to rectangular bounds if no polygon or scanline misses.
    fn exclusion_at_y(&self, y: f32) -> (f32, f32) {
        if let Some(ref pts) = self.polygon_pts {
            if let Some((left, right)) = poly_scanline(pts, y) {
                return (left, right);
            }
        }
        (self.obj_left, self.obj_right)
    }

    /// Narrow paragraph geometry to fit beside this floating object.
    /// Returns `Some((para_text_x, para_text_width, label_x))` if the zone
    /// is active at `slot_top`, or `None` if no narrowing is needed.
    #[allow(dead_code)]
    fn narrow_paragraph_geometry(
        &self,
        slot_top: f32,
        col_x: f32,
        col_w: f32,
        indent_left: f32,
        indent_right: f32,
        indent_hanging: f32,
    ) -> Option<(f32, f32, f32)> {
        if !(slot_top <= self.top_y && slot_top > self.bottom_y) {
            return None;
        }
        let col_right = col_x + col_w;
        let (ex_left, ex_right) = self.exclusion_at_y(slot_top);
        let space_right = col_right - (ex_right + self.right_from_text);
        let space_left = (ex_left - self.left_from_text) - col_x;

        let mut para_text_x = col_x + indent_left;
        let mut para_text_width = (col_w - indent_left - indent_right).max(1.0);
        let mut label_x = col_x + indent_left - indent_hanging;

        if self.wrap_text == WrapText::BothSides {
            let lw = (space_left - indent_left).max(0.0);
            let rw = (space_right - indent_right).max(0.0);
            if rw > lw {
                let new_left = ex_right + self.right_from_text;
                para_text_width = rw.max(1.0);
                para_text_x = new_left + indent_left;
                label_x = new_left + indent_left - indent_hanging;
            } else if lw > 0.0 {
                para_text_width = lw.max(1.0);
            }
        } else {
            let use_right = match self.wrap_text {
                WrapText::Right => space_right >= 1.0,
                WrapText::Left => !(space_left >= 1.0),
                _ => space_right >= space_left && space_right >= 72.0,
            };
            let use_left = match self.wrap_text {
                WrapText::Left => space_left >= 1.0,
                WrapText::Right => false,
                _ => space_left >= 72.0,
            };
            if use_right {
                let new_left = ex_right + self.right_from_text;
                para_text_width = (col_right - new_left - indent_right).max(1.0);
                para_text_x = new_left + indent_left;
                label_x = new_left + indent_left - indent_hanging;
            } else if use_left {
                let avail_right = ex_left - self.left_from_text;
                para_text_width =
                    (avail_right - col_x - indent_left - indent_right).max(1.0);
            }
        }
        Some((para_text_x, para_text_width, label_x))
    }
}

/// Scanline intersection: find the leftmost and rightmost x where polygon edges cross y.
fn poly_scanline(pts: &[(f32, f32)], y: f32) -> Option<(f32, f32)> {
    let n = pts.len();
    if n < 3 {
        return None;
    }
    let mut min_x = f32::MAX;
    let mut max_x = f32::MIN;
    for i in 0..n {
        let (x0, y0) = pts[i];
        let (x1, y1) = pts[(i + 1) % n];
        if (y0 <= y && y1 >= y) || (y1 <= y && y0 >= y) {
            if (y1 - y0).abs() < 0.001 {
                min_x = min_x.min(x0).min(x1);
                max_x = max_x.max(x0).max(x1);
            } else {
                let t = (y - y0) / (y1 - y0);
                let x = x0 + t * (x1 - x0);
                min_x = min_x.min(x);
                max_x = max_x.max(x);
            }
        }
    }
    if min_x <= max_x {
        Some((min_x, max_x))
    } else {
        None
    }
}

/// Convert polygon vertices from 1/21600-of-extent coords to absolute page coords.
fn convert_polygon_to_page_coords(
    vertices: &[(i32, i32)],
    img_x: f32,
    img_y_top: f32,
    display_w: f32,
    display_h: f32,
) -> Vec<(f32, f32)> {
    vertices
        .iter()
        .map(|&(px, py)| {
            let x_pt = img_x + (px as f32 / 21600.0) * display_w;
            let y_pt = img_y_top - (py as f32 / 21600.0) * display_h;
            (x_pt, y_pt)
        })
        .collect()
}

/// Draw debug overlays for the wrap polygon and effective exclusion zone.
/// Green = raw polygon, blue = wrap zone (polygon ± dist margins), red = top/bottom bounds.
fn draw_debug_wrap_overlay(content: &mut Content, fz: &FloatZone) {
    content.save_state();

    // Green: raw polygon outline
    if let Some(ref pts) = fz.polygon_pts {
        if pts.len() >= 3 {
            content.set_stroke_rgb(0.0, 0.7, 0.0);
            content.set_line_width(0.5);
            content.move_to(pts[0].0, pts[0].1);
            for &(x, y) in &pts[1..] {
                content.line_to(x, y);
            }
            content.close_path();
            content.stroke();
        }
    }

    // Blue: effective wrap zone boundaries (polygon shifted by dist margins)
    if let Some(ref pts) = fz.polygon_pts {
        content.set_stroke_rgb(0.0, 0.0, 0.8);
        content.set_line_width(0.3);
        let steps = ((fz.top_y - fz.bottom_y) / 1.0).ceil() as usize;
        if steps > 0 {
            let mut left_pts = Vec::with_capacity(steps + 1);
            let mut right_pts = Vec::with_capacity(steps + 1);
            for i in 0..=steps {
                let y = fz.top_y - i as f32 * (fz.top_y - fz.bottom_y) / steps as f32;
                if let Some((l, r)) = poly_scanline(pts, y) {
                    left_pts.push((l - fz.left_from_text, y));
                    right_pts.push((r + fz.right_from_text, y));
                }
            }
            if left_pts.len() >= 2 {
                content.move_to(left_pts[0].0, left_pts[0].1);
                for &(x, y) in &left_pts[1..] {
                    content.line_to(x, y);
                }
                content.stroke();
            }
            if right_pts.len() >= 2 {
                content.move_to(right_pts[0].0, right_pts[0].1);
                for &(x, y) in &right_pts[1..] {
                    content.line_to(x, y);
                }
                content.stroke();
            }
        }
    }

    // Red: top and bottom zone boundaries
    content.set_stroke_rgb(0.8, 0.0, 0.0);
    content.set_line_width(0.3);
    let page_left = fz.obj_left - 20.0;
    let page_right = fz.obj_right + 20.0;
    content.move_to(page_left, fz.top_y);
    content.line_to(page_right, fz.top_y);
    content.stroke();
    content.move_to(page_left, fz.bottom_y);
    content.line_to(page_right, fz.bottom_y);
    content.stroke();

    content.restore_state();
}

pub(super) struct FloatingTablePos {
    pub x: f32,
    pub y: f32,
    pub top_from_text: f32,
    pub bottom_from_text: f32,
    pub left_from_text: f32,
    pub right_from_text: f32,
}

pub(super) struct PageBuilder {
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
    pub(super) is_first_page_of_section: bool,
    /// Section that owns the current page for header/footer purposes.
    /// For continuous section breaks, this stays as the section that started
    /// the page, not the section that continues mid-page.
    page_hf_section: usize,
    /// Floating table exclusion zone on this page; paragraph layout
    /// uses horizontal bounds to decide wrap-beside vs push-below.
    pub(super) float_zone: Option<FloatZone>,
    /// Anchored shapes for this page, painted above the text layer sorted by
    /// relativeHeight: Word stacks floating shapes in z order regardless of
    /// which paragraph anchors them.
    pub(super) deferred_shapes: Vec<(u32, Content)>,

    // Accumulated pages
    all_contents: Vec<Content>,
    pub(super) all_deferred_shapes: Vec<Vec<(u32, Content)>>,
    all_links: Vec<Vec<LinkAnnotation>>,
    pub(super) all_comment_anchors: Vec<Vec<(u32, f32, f32, f32)>>,
    all_footnote_ids: Vec<Vec<u32>>,
    all_alpha_states: Vec<HashSet<u8>>,
    all_gradient_specs: Vec<Vec<GradientSpec>>,
    /// Per-page tuples: (hf_section, is_first_page, content_section).
    /// hf_section: which section provides headers/footers.
    /// content_section: which section is being rendered (for page numbering, geometry).
    page_section_indices: Vec<(usize, bool, usize)>,
    all_styleref: Vec<HashMap<String, String>>,
    all_first_styleref: Vec<HashMap<String, String>>,
}

impl PageBuilder {
    fn new(slot_top: f32) -> Self {
        PageBuilder {
            content: Content::new(),
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
            is_first_page_of_section: true,
            page_hf_section: 0,
            float_zone: None,
            deferred_shapes: Vec::new(),
            all_contents: Vec::new(),
            all_deferred_shapes: Vec::new(),
            all_links: Vec::new(),
            all_comment_anchors: Vec::new(),
            all_footnote_ids: Vec::new(),
            all_alpha_states: Vec::new(),
            all_gradient_specs: Vec::new(),
            page_section_indices: Vec::new(),
            all_styleref: Vec::new(),
            all_first_styleref: Vec::new(),
        }
    }

    pub(super) fn flush_page(&mut self, sect_idx: usize) {
        self.all_contents
            .push(std::mem::replace(&mut self.content, Content::new()));
        // Stable sort: equal relativeHeight keeps document order
        self.deferred_shapes.sort_by_key(|(z, _)| *z);
        self.all_deferred_shapes
            .push(std::mem::take(&mut self.deferred_shapes));
        self.all_links.push(std::mem::take(&mut self.links));
        self.all_comment_anchors
            .push(std::mem::take(&mut self.comment_anchors));
        self.all_footnote_ids
            .push(std::mem::take(&mut self.footnote_ids));
        self.footnote_ids_set.clear();
        self.all_alpha_states
            .push(std::mem::take(&mut self.alpha_states));
        self.all_gradient_specs
            .push(std::mem::take(&mut self.gradient_specs));
        self.page_section_indices.push((
            self.page_hf_section,
            self.is_first_page_of_section,
            sect_idx,
        ));
        self.all_styleref.push(self.styleref_running.clone());
        self.all_first_styleref
            .push(std::mem::take(&mut self.styleref_page_first));
        self.float_zone = None;
        // After flush, the new page starts with the current section
        self.page_hf_section = sect_idx;
    }

    fn push_blank_page(&mut self, sect_idx: usize) {
        self.all_contents.push(Content::new());
        self.all_deferred_shapes.push(Vec::new());
        self.all_links.push(Vec::new());
        self.all_comment_anchors.push(Vec::new());
        self.all_footnote_ids.push(Vec::new());
        self.all_alpha_states.push(HashSet::new());
        self.all_gradient_specs.push(Vec::new());
        self.page_section_indices
            .push((self.page_hf_section, false, sect_idx));
        self.all_styleref.push(self.styleref_running.clone());
        self.all_first_styleref
            .push(std::mem::take(&mut self.styleref_page_first));
        self.page_hf_section = sect_idx;
    }

    fn page_count(&self) -> usize {
        self.all_contents.len()
    }

    fn is_at_page_top(&self, sp: &SectionProperties) -> bool {
        (self.slot_top - (sp.page_height - sp.margin_top)).abs() < 1.0
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
            self.slot_top = self.column_top_y;
        } else {
            *current_col = 0;
            self.flush_page(sect_idx);
            self.slot_top = effective_slot_top(sp, false, ctx);
            self.column_top_y = self.slot_top;
            *effective_margin_bottom = compute_effective_margin_bottom(sp, false, ctx);
            self.is_first_page_of_section = false;
        }
    }
}

/// Bundles the mutable render loop state so it can be passed to extracted functions.
pub(super) struct LayoutState {
    pub(super) pb: PageBuilder,
    pub(super) prev_space_after: f32,
    pub(super) effective_margin_bottom: f32,
    pub(super) current_col: usize,
    pub(super) global_block_idx: usize,
    pub(super) heading_entries: Vec<HeadingEntry>,
    pub(super) bookmark_positions: HashMap<String, (usize, f32)>,
}

/// Compute effective first-line hanging indent for a paragraph.
fn compute_text_hanging(para: &Paragraph, default_tab_stop: f32) -> f32 {
    if !para.list_label.is_empty() {
        if let Some(nts) = para.num_level_tab_stop {
            if nts < para.indent_left
                && (para.indent_left - para.indent_hanging).abs() < 0.5
            {
                (para.indent_left - nts).max(0.0)
            } else if nts > para.indent_left
                && para.indent_hanging == 0.0
                && para.indent_first_line == 0.0
            {
                -(nts - para.indent_left)
            } else if para.indent_first_line > 0.0 && para.indent_hanging == 0.0 {
                -para.indent_first_line
            } else {
                0.0
            }
        } else if para.indent_hanging == 0.0
            && para.indent_first_line == 0.0
            && default_tab_stop > 0.0
        {
            // No num tab stop defined: Word renders label at indent_left, then a
            // tab advances text to the next default tab stop *position* (not by
            // that amount). Target the next multiple of default_tab_stop that
            // is strictly greater than indent_left.
            let next_tab =
                ((para.indent_left / default_tab_stop).floor() + 1.0) * default_tab_stop;
            -(next_tab - para.indent_left)
        } else if para.indent_first_line > 0.0 && para.indent_hanging == 0.0 {
            -para.indent_first_line
        } else {
            0.0
        }
    } else if para.indent_hanging > 0.0 {
        para.indent_hanging
    } else {
        -para.indent_first_line
    }
}

/// Pre-compute bookmark page positions so PAGEREF fields (e.g. TOC) can
/// show correct page numbers. Simulates page layout without rendering.
fn compute_bookmark_positions(
    doc: &Document,
    ctx: &RenderContext,
) -> HashMap<String, (usize, f32)> {
    let has_pagerefs = doc.sections.iter().any(|s| {
        s.blocks.iter().any(|b| match b {
            Block::Paragraph(p) => p
                .runs
                .iter()
                .any(|r| matches!(&r.field_code, Some(FieldCode::PageRef(_)))),
            _ => false,
        })
    });
    if !has_pagerefs {
        return HashMap::new();
    }

    let mut bookmark_positions: HashMap<String, (usize, f32)> = HashMap::new();
    let mut page_idx = 0usize;
    let first_sp = &doc.sections[0].properties;
    let mut sp = first_sp;
    let mut slot_top = effective_slot_top(sp, true, ctx);
    let mut margin_bottom = compute_effective_margin_bottom(sp, true, ctx);
    let mut prev_space_after: f32 = 0.0;
    let mut prev_contextual: bool = false;
    let empty_imgs: HashMap<usize, String> = HashMap::new();
    let empty_fx: HashMap<usize, images::EffectXObjs> = HashMap::new();

    for (si, section) in doc.sections.iter().enumerate() {
        sp = &section.properties;
        if si > 0 {
            match sp.break_type {
                SectionBreakType::NextPage
                | SectionBreakType::OddPage
                | SectionBreakType::EvenPage => {
                    page_idx += 1;
                    slot_top = effective_slot_top(sp, true, ctx);
                    margin_bottom = compute_effective_margin_bottom(sp, true, ctx);
                    prev_space_after = 0.0;
                }
                SectionBreakType::Continuous => {}
            }
        }
        let text_width = sp.page_width - sp.margin_left - sp.margin_right;
        let blocks = &section.blocks;
        for (bi, block) in blocks.iter().enumerate() {
            match block {
                Block::Paragraph(para) => {
                    if para.page_break_before
                        && slot_top < effective_slot_top(sp, false, ctx)
                    {
                        page_idx += 1;
                        slot_top = effective_slot_top(sp, false, ctx);
                        margin_bottom = compute_effective_margin_bottom(sp, false, ctx);
                        prev_space_after = 0.0;
                    }
                    for bm in &para.bookmarks {
                        bookmark_positions.insert(bm.clone(), (page_idx, slot_top));
                    }
                    if para.is_section_break && is_text_empty(&para.runs) {
                        continue;
                    }
                    let (mut font_size, mut tallest_lhr, _) =
                        tallest_run_metrics(&para.runs, ctx.fonts);
                    if tallest_lhr.is_none() {
                        if let Some(mark_fs) = para.paragraph_mark_font_size {
                            font_size = mark_fs;
                        }
                        if let Some(ref mark_fn) = para.paragraph_mark_font_name {
                            if let Some(entry) = ctx.fonts.get(mark_fn.as_str()) {
                                tallest_lhr = entry.line_h_ratio;
                            }
                        }
                    }
                    let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                    let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
                    let line_h = if para.snap_to_grid
                        && matches!(
                            sp.grid_type,
                            DocGridType::Lines
                                | DocGridType::LinesAndChars
                                | DocGridType::SnapToChars
                        )
                        && !matches!(effective_ls, LineSpacing::Exact(_))
                        && sp.line_pitch > 0.0
                    {
                        (line_h / sp.line_pitch).ceil() * sp.line_pitch
                    } else {
                        line_h
                    };
                    let para_w =
                        (text_width - para.indent_left - para.indent_right).max(1.0);
                    let hanging = compute_text_hanging(para, ctx.default_tab_stop);
                    let has_tabs = para.runs.iter().any(|r| r.is_tab);
                    let lines = if is_text_empty(&para.runs) {
                        vec![]
                    } else if has_tabs {
                        build_tabbed_line(
                            &para.runs,
                            ctx.fonts,
                            &para.tab_stops,
                            para.indent_left,
                            para_w,
                            para.indent_right,
                            hanging,
                            &empty_imgs,
                            &empty_fx,
                            doc.default_tab_stop,
                        )
                    } else {
                        build_paragraph_lines(
                            &para.runs, ctx.fonts, para_w, hanging, &empty_imgs,
                            &empty_fx, None, None, None,
                            para.auto_space_de || para.auto_space_dn,
                        )
                    };
                    let num_lines = lines.len().max(1);
                    let para_has_inline_img =
                        para.runs.iter().any(|r| r.inline_image.is_some());
                    let content_h = if para.image.is_some() || para.inline_chart.is_some()
                    {
                        para.content_height
                    } else if para_has_inline_img && para.content_height > 0.0 {
                        para.content_height.max(num_lines as f32 * line_h)
                    } else {
                        num_lines as f32 * line_h
                    };
                    let effective_sb = if para.contextual_spacing && prev_contextual {
                        0.0
                    } else {
                        para.space_before
                    };
                    let next_contextual = blocks.get(bi + 1).is_some_and(|b| {
                        matches!(b, Block::Paragraph(p) if p.contextual_spacing)
                    });
                    let effective_sa = if para.contextual_spacing && next_contextual {
                        0.0
                    } else {
                        para.space_after
                    };
                    let inter_gap = f32::max(prev_space_after, effective_sb);
                    let needed = inter_gap + content_h;
                    if slot_top - needed < margin_bottom
                        && slot_top < effective_slot_top(sp, false, ctx)
                    {
                        page_idx += 1;
                        slot_top = effective_slot_top(sp, false, ctx);
                        margin_bottom =
                            compute_effective_margin_bottom(sp, false, ctx);
                        slot_top -= content_h;
                    } else {
                        slot_top -= inter_gap + content_h;
                    }
                    prev_space_after = effective_sa;
                    prev_contextual = para.contextual_spacing;
                }
                Block::Table(table) => {
                    let para_count: usize = table
                        .rows
                        .iter()
                        .flat_map(|r| r.cells.iter())
                        .map(|c| c.all_paragraphs().len().max(1))
                        .max()
                        .unwrap_or(1)
                        * table.rows.len();
                    let est_h = para_count as f32 * 14.0;
                    if slot_top - est_h < margin_bottom {
                        page_idx += 1;
                        slot_top = effective_slot_top(sp, false, ctx);
                        margin_bottom =
                            compute_effective_margin_bottom(sp, false, ctx);
                    }
                    slot_top -= est_h;
                    prev_space_after = 0.0;
                    prev_contextual = false;
                }
            }
        }
    }
    bookmark_positions
}

/// Render a single paragraph block. Returns `true` if the block was skipped
/// (the caller should `continue` the block loop).
#[allow(clippy::too_many_arguments)]
fn render_paragraph_block(
    para: &Paragraph,
    state: &mut LayoutState,
    ctx: &RenderContext,
    sp: &SectionProperties,
    col_geometry: &[(f32, f32)],
    col_count: usize,
    text_width: f32,
    sect_idx: usize,
    block_idx: usize,
    section_blocks: &[Block],
    floating_image_pdf_names: &HashMap<(usize, usize), String>,
    inline_image_pdf_names: &HashMap<(usize, usize), String>,
    image_pdf_names: &HashMap<usize, String>,
    effect_names: &HashMap<usize, EffectXObjs>,
    effect_floating_names: &HashMap<(usize, usize), EffectXObjs>,
    effect_inline_names: &HashMap<(usize, usize), EffectXObjs>,
    footnote_display_order: &HashMap<u32, u32>,
    endnote_display_order: &HashMap<u32, u32>,
    doc: &Document,
    smartart_font_key: &str,
    smartart_image_names: &HashMap<usize, String>,
    debug_wrap: bool,
) -> bool {
    let adjacent_para = |idx: usize| -> Option<&Paragraph> {
        match section_blocks.get(idx)? {
            Block::Paragraph(p) => Some(p),
            Block::Table(_) => None,
        }
    };

    // Skip empty section-break paragraphs — Word gives these zero height
    if para.is_section_break
        && is_text_empty(&para.runs)
        && para.image.is_none()
        && para.inline_chart.is_none()
        && para.smartart.is_empty()
        && para.floating_images.is_empty()
        && para.textboxes.is_empty()
    {
        state.global_block_idx += 1;
        return true;
    }

    // Handle explicit page breaks.
    // `<w:br w:type="page"/>` (page_break_before_explicit) advances
    // unconditionally — even at the top of a page, Word emits a blank
    // page when the explicit break follows a section break. The
    // `<w:pageBreakBefore/>` style property is idempotent and skipped
    // when already at the top.
    if para.page_break_before {
        let at_top = state.pb.is_at_page_top(sp);
        if !at_top || para.page_break_before_explicit {
            state.pb.flush_page(sect_idx);
            state.pb.slot_top = effective_slot_top(sp, false, &ctx);
            state.pb.column_top_y = state.pb.slot_top;
            state.effective_margin_bottom =
                compute_effective_margin_bottom(sp, false, &ctx);
            state.pb.is_first_page_of_section = false;
            state.current_col = 0;
        }
        state.prev_space_after = 0.0;
        if is_text_empty(&para.runs) {
            state.global_block_idx += 1;
            return true;
        }
    }

    // Handle explicit column breaks
    if para.column_break_before && col_count > 1 {
        state.pb.advance_column_or_page(
            &mut state.current_col,
            col_count,
            sect_idx,
            sp,
            &mut state.effective_margin_bottom,
            &ctx,
        );
        state.prev_space_after = 0.0;
    }

    let next_para = adjacent_para(block_idx + 1);
    let prev_para = if block_idx > 0 {
        adjacent_para(block_idx - 1)
    } else {
        None
    };

    let effective_space_before = if para.contextual_spacing
        && prev_para.is_some_and(|p| p.contextual_spacing)
    {
        0.0
    } else {
        para.space_before
    };
    let effective_space_after = if para.contextual_spacing
        && next_para.is_some_and(|p| p.contextual_spacing)
    {
        0.0
    } else {
        para.space_after
    };

    let mut inter_gap = f32::max(state.prev_space_after, effective_space_before);

    let (mut font_size, mut tallest_lhr, tallest_ar) =
        tallest_run_metrics(&para.runs, ctx.fonts);
    // When runs are empty, use the paragraph mark's font metrics instead
    // of the 12pt / 1.2x fallback that tallest_run_metrics returns.
    if tallest_lhr.is_none() {
        if let Some(mark_fs) = para.paragraph_mark_font_size {
            font_size = mark_fs;
        }
        let mark_font_name = para
            .paragraph_mark_font_name
            .as_deref()
            .unwrap_or(&para.runs.first().map(|r| r.font_name.as_str()).unwrap_or(""));
        if !mark_font_name.is_empty() {
            if let Some(entry) = ctx.fonts.get(mark_font_name) {
                tallest_lhr = entry.line_h_ratio;
            }
        }
    }
    let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
    let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
    let grid_snapped = para.snap_to_grid
        && matches!(
            sp.grid_type,
            DocGridType::Lines
                | DocGridType::LinesAndChars
                | DocGridType::SnapToChars
        )
        && !matches!(effective_ls, LineSpacing::Exact(_))
        && sp.line_pitch > 0.0;
    let line_h = if grid_snapped {
        (line_h / sp.line_pitch).ceil() * sp.line_pitch
    } else {
        line_h
    };

    let (col_x, col_w) = col_geometry[state.current_col];
    let mut para_text_x = col_x + para.indent_left;
    let mut para_text_width =
        (col_w - para.indent_left - para.indent_right).max(1.0);
    let mut label_x = col_x + para.indent_left - para.indent_hanging;

    // When inside a floating object zone, narrow the paragraph to
    // fit beside the object rather than overlapping it.
    if let Some(ref fz) = state.pb.float_zone {
        if state.pb.slot_top <= fz.top_y && state.pb.slot_top > fz.bottom_y {
            let col_right = col_x + col_w;
            let (ex_left, ex_right) = fz.exclusion_at_y(state.pb.slot_top);
            let space_right =
                col_right - (ex_right + fz.right_from_text);
            let space_left = (ex_left - fz.left_from_text) - col_x;

            if fz.wrap_text == WrapText::BothSides {
                // For bothSides, use the wider region as
                // primary text width (dual geometry handles
                // both regions per-line).
                let lw = (space_left - para.indent_left).max(0.0);
                let rw = (space_right - para.indent_right).max(0.0);
                if rw > lw {
                    let new_left = ex_right + fz.right_from_text;
                    para_text_width = rw.max(1.0);
                    para_text_x = new_left + para.indent_left;
                    label_x =
                        new_left + para.indent_left - para.indent_hanging;
                } else if lw > 0.0 {
                    para_text_width = lw.max(1.0);
                }
            } else {
                let use_right = match fz.wrap_text {
                    WrapText::Right => space_right >= 1.0,
                    WrapText::Left => !(space_left >= 1.0),
                    _ => space_right >= space_left && space_right >= 72.0,
                };
                let use_left = match fz.wrap_text {
                    WrapText::Left => space_left >= 1.0,
                    WrapText::Right => false,
                    _ => space_left >= 72.0,
                };
                if use_right {
                    let new_left = ex_right + fz.right_from_text;
                    para_text_width =
                        (col_right - new_left - para.indent_right).max(1.0);
                    para_text_x = new_left + para.indent_left;
                    label_x =
                        new_left + para.indent_left - para.indent_hanging;
                } else if use_left {
                    let avail_right = ex_left - fz.left_from_text;
                    para_text_width = (avail_right - col_x
                        - para.indent_left
                        - para.indent_right)
                        .max(1.0);
                }
            }
        }
    }

    let text_hanging = compute_text_hanging(para, ctx.default_tab_stop);

    // Substitute footnote/endnote refs and resolve PAGEREF fields
    let has_footnote_refs = para.runs.iter().any(|r| r.footnote_id.is_some());
    let has_endnote_refs = para.runs.iter().any(|r| r.endnote_id.is_some());
    let has_pageref = para.runs.iter().any(|r| matches!(&r.field_code, Some(FieldCode::PageRef(_))));
    let effective_runs: std::borrow::Cow<'_, Vec<Run>> = if has_footnote_refs || has_endnote_refs || has_pageref {
        let substituted: Vec<Run> = para
            .runs
            .iter()
            .map(|run| {
                if let Some(id) = run.footnote_id {
                    let num = footnote_display_order.get(&id).copied().unwrap_or(0);
                    let mut r = run.clone();
                    r.text = num.to_string();
                    r
                } else if let Some(id) = run.endnote_id {
                    let num = endnote_display_order.get(&id).copied().unwrap_or(0);
                    let mut r = run.clone();
                    r.text = num.to_string();
                    r
                } else if let Some(FieldCode::PageRef(ref bookmark)) = run.field_code {
                    let mut r = run.clone();
                    if let Some(&(page_idx, _)) = state.bookmark_positions.get(bookmark) {
                        r.text = (page_idx + 1).to_string();
                    }
                    r
                } else {
                    run.clone()
                }
            })
            .collect();
        std::borrow::Cow::Owned(substituted)
    } else {
        std::borrow::Cow::Borrowed(&para.runs)
    };

    let text_empty = is_text_empty(&effective_runs);
    let has_tabs = effective_runs.iter().any(|r| r.is_tab);
    let block_inline_images: HashMap<usize, String> = inline_image_pdf_names
        .iter()
        .filter(|((bi, _), _)| *bi == state.global_block_idx)
        .map(|((_, ri), name)| (*ri, name.clone()))
        .collect();
    let block_effect_inlines: HashMap<usize, images::EffectXObjs> = effect_inline_names
        .iter()
        .filter(|((bi, _), _)| *bi == state.global_block_idx)
        .map(|((_, ri), fx)| (*ri, fx.clone()))
        .collect();
    // Self-wrapping: if this paragraph anchors a wrapping float
    // and has text, set up the float zone NOW so width-narrowing
    // applies to this paragraph's own lines.  Always replace any
    // previous float zone — the paragraph's own image takes priority.
    if !para.floating_images.is_empty()
        && !text_empty
    {
        if let Some(fi) = para.floating_images.iter().find(|fi| {
            matches!(
                fi.wrap_type,
                WrapType::Square | WrapType::Tight | WrapType::Through
            )
        }) {
            let fi_x =
                resolve_fi_x(fi, sp, col_x, col_w, text_width);
            let fi_y_top =
                resolve_fi_y_top(fi, sp, state.pb.slot_top);
            let fi_y_bottom =
                fi_y_top - fi.image.display_height;
            let polygon_pts =
                fi.wrap_polygon.as_ref().map(|verts| {
                    convert_polygon_to_page_coords(
                        verts,
                        fi_x,
                        fi_y_top,
                        fi.image.display_width,
                        fi.image.display_height,
                    )
                });
            state.pb.float_zone = Some(FloatZone {
                top_y: fi_y_top + fi.dist_top,
                bottom_y: fi_y_bottom - fi.dist_bottom,
                obj_left: fi_x,
                obj_right: fi_x + fi.image.display_width,
                left_from_text: fi.dist_left,
                right_from_text: fi.dist_right,
                polygon_pts,
                wrap_text: fi.wrap_text,
                para_relative: fi.v_relative_from == VRelativeFrom::Paragraph,
            });
            // Re-narrow para_text_x / para_text_width using the
            // new float zone (same logic as the block above).
            let fz = state.pb.float_zone.as_ref().unwrap();
            if state.pb.slot_top <= fz.top_y && state.pb.slot_top > fz.bottom_y
            {
                let col_right = col_x + col_w;
                let (ex_left, ex_right) =
                    fz.exclusion_at_y(state.pb.slot_top);
                let space_right = col_right
                    - (ex_right + fz.right_from_text);
                let space_left =
                    (ex_left - fz.left_from_text) - col_x;

                if fz.wrap_text == WrapText::BothSides {
                    let lw = (space_left - para.indent_left).max(0.0);
                    let rw = (space_right - para.indent_right).max(0.0);
                    if rw > lw {
                        let new_left = ex_right + fz.right_from_text;
                        para_text_width = rw.max(1.0);
                        para_text_x = new_left + para.indent_left;
                        label_x = new_left + para.indent_left
                            - para.indent_hanging;
                    } else if lw > 0.0 {
                        para_text_width = lw.max(1.0);
                    }
                } else {
                    let use_right = match fz.wrap_text {
                        WrapText::Right => space_right >= 1.0,
                        WrapText::Left => !(space_left >= 1.0),
                        _ => space_right >= space_left && space_right >= 72.0,
                    };
                    let use_left = match fz.wrap_text {
                        WrapText::Left => space_left >= 1.0,
                        WrapText::Right => false,
                        _ => space_left >= 72.0,
                    };
                    if use_right {
                        let new_left =
                            ex_right + fz.right_from_text;
                        para_text_width = (col_right
                            - new_left
                            - para.indent_right)
                            .max(1.0);
                        para_text_x =
                            new_left + para.indent_left;
                        label_x = new_left + para.indent_left
                            - para.indent_hanging;
                    } else if use_left {
                        let avail_right =
                            ex_left - fz.left_from_text;
                        para_text_width = (avail_right - col_x
                            - para.indent_left
                            - para.indent_right)
                            .max(1.0);
                    }
                }
            }
        }
    }

    // Additional wrapping floats anchored to this same paragraph beyond the
    // first (which became `float_zone` above). The single-zone geometry below
    // can't model e.g. a logo on each side of a centered title, so when these
    // exist we compute per-line free intervals across all zones instead.
    let extra_float_zones: Vec<FloatZone> = if !text_empty {
        para.floating_images
            .iter()
            .filter(|fi| {
                matches!(
                    fi.wrap_type,
                    WrapType::Square | WrapType::Tight | WrapType::Through
                )
            })
            .skip(1)
            .map(|fi| {
                let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
                let fi_y_top = resolve_fi_y_top(fi, sp, state.pb.slot_top);
                let fi_y_bottom = fi_y_top - fi.image.display_height;
                let polygon_pts = fi.wrap_polygon.as_ref().map(|verts| {
                    convert_polygon_to_page_coords(
                        verts,
                        fi_x,
                        fi_y_top,
                        fi.image.display_width,
                        fi.image.display_height,
                    )
                });
                FloatZone {
                    top_y: fi_y_top + fi.dist_top,
                    bottom_y: fi_y_bottom - fi.dist_bottom,
                    obj_left: fi_x,
                    obj_right: fi_x + fi.image.display_width,
                    left_from_text: fi.dist_left,
                    right_from_text: fi.dist_right,
                    polygon_pts,
                    wrap_text: fi.wrap_text,
                    para_relative: fi.v_relative_from == VRelativeFrom::Paragraph,
                }
            })
            .collect()
    } else {
        Vec::new()
    };

    // Build per-line geometry and dual-region geometry
    let (poly_line_geom, poly_dual_geom): (Option<Vec<(f32, f32)>>, Option<Vec<DualRegion>>) =
        if let Some(fz) = state.pb.float_zone.as_ref() {
            let eff_top = state.pb.slot_top - inter_gap;
            if !extra_float_zones.is_empty() {
                // Multiple wrapping floats on one paragraph: subtract every
                // float's exclusion span per line and lay the text in the
                // widest remaining gap (Word places the line between the
                // floats). Dual regions only model a single float, so they
                // are skipped here.
                let zones: Vec<&FloatZone> =
                    std::iter::once(fz).chain(extra_float_zones.iter()).collect();
                if zones.iter().all(|z| eff_top <= z.bottom_y) {
                    (None, None)
                } else {
                    let full_w =
                        (col_w - para.indent_left - para.indent_right).max(1.0);
                    let col_right = col_x + col_w;
                    let lowest_bottom = zones
                        .iter()
                        .map(|z| z.bottom_y)
                        .fold(f32::INFINITY, f32::min);
                    let max_lines =
                        ((((eff_top - lowest_bottom) / line_h).ceil() as usize) + 5)
                            .max(50);
                    let mut geom = Vec::with_capacity(max_lines);
                    for i in 0..max_lines {
                        let line_top = eff_top - i as f32 * line_h;
                        let line_bottom = line_top - line_h;
                        let mut intervals: Vec<(f32, f32)> =
                            vec![(col_x, col_right)];
                        for z in &zones {
                            let partial_overlap =
                                z.para_relative || z.polygon_pts.is_some();
                            let in_zone = if partial_overlap {
                                line_bottom < z.top_y
                            } else {
                                line_top <= z.top_y
                            };
                            if !(in_zone && line_top > z.bottom_y + line_h * 0.2) {
                                continue;
                            }
                            let query_y = line_top.min(z.top_y);
                            let (ex_left, ex_right) = z.exclusion_at_y(query_y);
                            let sl = ex_left - z.left_from_text;
                            let sr = ex_right + z.right_from_text;
                            let mut next = Vec::with_capacity(intervals.len() + 1);
                            for &(a, b) in &intervals {
                                if sr <= a || sl >= b {
                                    next.push((a, b));
                                    continue;
                                }
                                if sl > a {
                                    next.push((a, sl));
                                }
                                if sr < b {
                                    next.push((sr, b));
                                }
                            }
                            intervals = next;
                        }
                        let best = intervals
                            .into_iter()
                            .max_by(|x, y| (x.1 - x.0).total_cmp(&(y.1 - y.0)));
                        match best {
                            Some((a, b))
                                if b - a
                                    > para.indent_left + para.indent_right + 1.0 =>
                            {
                                geom.push((
                                    a + para.indent_left,
                                    (b - a - para.indent_left - para.indent_right)
                                        .max(1.0),
                                ));
                            }
                            _ => geom.push((col_x + para.indent_left, full_w)),
                        }
                    }
                    (Some(geom), None)
                }
            } else if eff_top <= fz.bottom_y {
                (None, None)
            } else {
                let is_both_sides =
                    fz.wrap_text == WrapText::BothSides;
                let full_w = (col_w
                    - para.indent_left
                    - para.indent_right)
                    .max(1.0);
                let col_right = col_x + col_w;
                let max_lines = ((eff_top - fz.bottom_y)
                    / line_h)
                    .ceil() as usize
                    + 5;
                let max_lines = max_lines.max(50);
                let mut geom = Vec::with_capacity(max_lines);
                let mut dual = if is_both_sides {
                    Some(Vec::with_capacity(max_lines))
                } else {
                    None
                };
                // Bottom threshold of 0.2 * line_h excludes lines
                // barely overlapping the zone, matching Word's behavior.
                let bottom_threshold = fz.bottom_y + line_h * 0.2;
                // Paragraph-relative float images and polygon
                // zones: check if any part of the line overlaps
                // (catches lines starting just above the zone
                // whose bottom extends into it). Floating table
                // zones use the simpler line-top check because
                // their boundaries already include topFromText /
                // bottomFromText clearance.
                let partial_overlap =
                    fz.para_relative || fz.polygon_pts.is_some();
                for i in 0..max_lines {
                    let line_top = eff_top - i as f32 * line_h;
                    let line_bottom = line_top - line_h;
                    let in_zone = if partial_overlap {
                        line_bottom < fz.top_y
                    } else {
                        line_top <= fz.top_y
                    };
                    if in_zone && line_top > bottom_threshold
                    {
                        let query_y = line_top.min(fz.top_y);
                        let (ex_left, ex_right) =
                            fz.exclusion_at_y(query_y);
                        let sr = col_right
                            - (ex_right + fz.right_from_text);
                        let sl =
                            (ex_left - fz.left_from_text) - col_x;

                        if is_both_sides {
                            // BothSides: provide both regions
                            let lx = col_x + para.indent_left;
                            let lw = (sl - para.indent_left).max(0.0);
                            let rx = ex_right + fz.right_from_text;
                            let rw = (sr - para.indent_right).max(0.0);
                            if let Some(ref mut d) = dual {
                                d.push((lx, lw, rx, rw));
                            }
                            // Single-region geometry always stores the
                            // LEFT region — render_paragraph_lines uses
                            // this for left-chunk x positioning.
                            geom.push((lx, lw));
                        } else {
                            // Left/Right/Largest: pick one side
                            let use_right = match fz.wrap_text {
                                WrapText::Right => sr >= 1.0,
                                WrapText::Left => !(sl >= 1.0),
                                _ => sr >= sl && sr >= 72.0,
                            };
                            let use_left = match fz.wrap_text {
                                WrapText::Left => sl >= 1.0,
                                WrapText::Right => false,
                                _ => sl >= 72.0,
                            };
                            if use_right {
                                let nl =
                                    ex_right + fz.right_from_text;
                                let w = (col_right
                                    - nl
                                    - para.indent_right)
                                    .max(1.0);
                                geom.push((
                                    nl + para.indent_left,
                                    w,
                                ));
                            } else if use_left {
                                let ar =
                                    ex_left - fz.left_from_text;
                                let w = (ar - col_x
                                    - para.indent_left
                                    - para.indent_right)
                                    .max(1.0);
                                geom.push((
                                    col_x + para.indent_left,
                                    w,
                                ));
                            } else {
                                geom.push((
                                    col_x + para.indent_left,
                                    full_w,
                                ));
                            }
                        }
                    } else {
                        geom.push((
                            col_x + para.indent_left,
                            full_w,
                        ));
                        if let Some(ref mut d) = dual {
                            d.push((col_x + para.indent_left, full_w, 0.0, 0.0));
                        }
                    }
                }
                (Some(geom), dual)
            }
        } else {
            (None, None)
        };

    let poly_line_widths: Option<Vec<f32>> =
        poly_line_geom.as_ref().map(|g| {
            g.iter().map(|&(_, w)| w).collect()
        });

    let mut float_width_change: Option<(usize, f32)> = None;
    // For look-ahead: (narrow_x, narrow_w) for lines after the split
    let mut lookahead_narrow: Option<(f32, f32)> = None;
    let has_inline_image_runs = effective_runs.iter().any(|r| r.inline_image.is_some());
    let lines = if para.image.is_some() || (text_empty && !has_inline_image_runs) {
        vec![]
    } else if has_tabs {
        build_tabbed_line(
            &effective_runs,
            ctx.fonts,
            &para.tab_stops,
            para.indent_left,
            para_text_width,
            para.indent_right,
            text_hanging,
            &block_inline_images,
            &block_effect_inlines,
            doc.default_tab_stop,
        )
    } else {
        // Look-ahead: if next block is an image-only paragraph
        // with wrapping, build lines at full width first, then
        // check if the bottom lines need narrowing.
        let lookahead_fi = if state.pb.float_zone.is_none() {
            section_blocks.get(block_idx + 1).and_then(|b| {
                if let Block::Paragraph(np) = b {
                    if !np.floating_images.is_empty()
                        && is_text_empty(&np.runs)
                        && np.image.is_none()
                        && np.inline_chart.is_none()
                    {
                        np.floating_images.iter().find(|fi| matches!(
                            fi.wrap_type,
                            WrapType::Square | WrapType::Tight | WrapType::Through
                        ) && fi.image.display_width < text_width * 0.5)
                    } else { None }
                } else { None }
            })
        } else { None };

        let auto_space = para.auto_space_de || para.auto_space_dn;
        let (lines, final_width_change) = if let Some(fi) = lookahead_fi {
            // Two-pass: build at full width, then narrow bottom lines
            let full_lines = build_paragraph_lines(
                &effective_runs, ctx.fonts, para_text_width,
                text_hanging, &block_inline_images, &block_effect_inlines, None, None, None,
                auto_space,
            );
            let num_lines = full_lines.len();
            let _content_h_est = num_lines as f32 * line_h;
            let fi_x = resolve_fi_x(fi, sp, col_x, col_w, col_w);
            let space_right = (col_x + col_w)
                - (fi_x + fi.image.display_width + fi.dist_right);
            let space_left = (fi_x - fi.dist_left) - col_x;
            let best = space_right.max(space_left);
            // Zone overlap: image starts at ~(slot_top - content_h - space_after)
            // zone extends dist_top above that into the current paragraph
            // Image-only paragraphs effectively have zero height
            // in Word, so the image anchors right at the preceding
            // paragraph's bottom. The zone extends dist_top above.
            let overlap = fi.dist_top;
            let lines_to_narrow = if best >= 72.0 && overlap > 0.0 {
                ((overlap / line_h).ceil() as usize).min(num_lines)
            } else { 0 };
            if lines_to_narrow > 0 {
                let lines_above = num_lines.saturating_sub(lines_to_narrow);
                let narrow_w = if space_right >= space_left {
                    ((col_x + col_w) - (fi_x + fi.image.display_width + fi.dist_right)
                        - para.indent_right).max(1.0)
                } else {
                    (fi_x - fi.dist_left - col_x
                        - para.indent_left - para.indent_right).max(1.0)
                };
                let narrow_x = if space_right >= space_left {
                    fi_x + fi.image.display_width + fi.dist_right
                        + para.indent_left
                } else {
                    col_x + para.indent_left
                };
                lookahead_narrow = Some((narrow_x, narrow_w));
                let rebuilt = build_paragraph_lines(
                    &effective_runs, ctx.fonts, para_text_width,
                    text_hanging, &block_inline_images, &block_effect_inlines,
                    Some((lines_above, narrow_w)), None, None,
                    auto_space,
                );
                (rebuilt, Some((lines_above, narrow_w)))
            } else {
                (full_lines, None)
            }
        } else {
            // Per-line geometry handles narrow→wide transitions;
            // dual geometry takes priority over single-region widths.
            let plw: Option<&[f32]> = if poly_dual_geom.is_some() { None } else { poly_line_widths.as_deref() };
            let built = build_paragraph_lines(
                &effective_runs, ctx.fonts, para_text_width,
                text_hanging, &block_inline_images, &block_effect_inlines, None,
                plw, poly_dual_geom.as_deref(),
                auto_space,
            );
            (built, None)
        };
        float_width_change = final_width_change;
        lines
    };

    // For lines containing inline images, use the tallest element as line height
    let max_inline_img_h = lines
        .iter()
        .flat_map(|l| l.chunks.iter())
        .map(|c| c.inline_image_height)
        .fold(0.0f32, f32::max);

    let mut content_h = if para.inline_chart.is_some() {
        para.content_height
    } else if para.image.is_some() {
        para.content_height
    } else if max_inline_img_h > 0.0 {
        let mut h = 0.0f32;
        for line in &lines {
            let img_h = line
                .chunks
                .iter()
                .map(|c| c.inline_image_height)
                .fold(0.0f32, f32::max);
            h += if img_h > line_h { img_h } else { line_h };
        }
        h
    } else if text_empty {
        if para.paragraph_mark_vanish {
            0.0
        } else if para.content_height > 0.0 {
            para.content_height
        } else {
            line_h
        }
    } else {
        let num_lines = lines.len();
        // List label font size doesn't boost line height — labels
        // sit in the margin and Word uses the text font for line
        // height, not the numbering run's font.
        let first_line_h = label_boosted_line_h(para, ctx.fonts, line_h, effective_ls, font_size);
        if num_lines <= 1 {
            // If the single line was created by a break, use its font size
            if let Some(bfs) = lines.first().and_then(|l| l.break_font_size) {
                let blhr = break_run_lhr(&effective_runs, bfs, ctx.fonts);
                resolve_line_h(effective_ls, bfs, blhr)
            } else {
                first_line_h
            }
        } else {
            // Per-line height: break-created lines use the break
            // run's font metrics instead of the paragraph's text metrics.
            let mut h = first_line_h;
            for line in lines.iter().skip(1) {
                if let Some(bfs) = line.break_font_size {
                    let blhr = break_run_lhr(&effective_runs, bfs, ctx.fonts);
                    h += resolve_line_h(effective_ls, bfs, blhr);
                } else {
                    h += line_h;
                }
            }
            h
        }
    };

    // Extra height from floating images that extends beyond
    // the text content — used only for page-break decisions,
    // not for cursor advancement (text wraps beside the image).
    let mut float_overflow_h = 0.0f32;

    for fi in &para.floating_images {
        let reserve = match fi.wrap_type {
            WrapType::TopAndBottom => true,
            // Square/Tight/Through use float zones for vertical
            // extent — don't block content_h, even for wide images.
            WrapType::Square | WrapType::Tight | WrapType::Through => false,
            WrapType::None => false,
        };
        let fi_h = match fi.v_position {
            VerticalPosition::Offset(o) => {
                o + fi.dist_top + fi.image.display_height + fi.dist_bottom
            }
            _ => fi.dist_top + fi.image.display_height + fi.dist_bottom,
        };
        if reserve {
            // Wide images block all text — add to content_h
            content_h = content_h.max(fi_h);
        } else if fi.v_relative_from == VRelativeFrom::Paragraph
            && matches!(
                fi.wrap_type,
                WrapType::Square | WrapType::Tight | WrapType::Through
            )
        {
            // Paragraph-relative wrapping images: track overflow
            // for page-break check only (text wraps beside them).
            float_overflow_h = float_overflow_h.max(fi_h);
        }
    }

    for tb in &para.textboxes {
        let reserve = match tb.wrap_type {
            WrapType::TopAndBottom => true,
            WrapType::Square => tb.width_pt >= text_width * 0.5,
            _ => false,
        };
        if reserve {
            let tb_bottom = tb.v_offset_pt + tb.height_pt + tb.dist_bottom;
            match tb.v_relative_from {
                VRelativeFrom::Paragraph => {
                    content_h = content_h.max(tb_bottom);
                }
                _ => {
                    content_h += tb_bottom;
                }
            }
        }
    }

    // Vanished paragraph mark: zero out height and spacing
    if text_empty && para.paragraph_mark_vanish {
        content_h = 0.0;
        inter_gap = 0.0;
    }

    // When consecutive paragraphs share matching borders, Word
    // suppresses the border padding on the continuous sides — only
    // the first paragraph in a group gets top padding and only the
    // last gets bottom padding.
    let prev_borders_match = prev_para
        .is_some_and(|pp| borders_match(&pp.borders, &para.borders));
    let next_borders_match = next_para
        .is_some_and(|np| borders_match(&para.borders, &np.borders));

    // Per ISO/IEC 29500 §17.3.1.5/.7, identical adjoining paragraphs share a
    // single `between` rule (or none, if unspecified) instead of individual
    // bottom/top borders. Word follows this — collapsing the divider — when a
    // `between` border is defined, when a competing top border meets this
    // bottom border, or when either paragraph is *empty* (e.g. blank spacer
    // paragraphs in a consent form, where the whole run collapses to one rule
    // at the group's outer edge). The exception is two *non-empty* identical
    // paragraphs separated only by a bottom rule — e.g. survey rows — where
    // Word keeps each rule, so the divider between them must still be drawn.
    let next_has_top = next_para.is_some_and(|np| np.borders.top.is_some());
    let next_is_empty = next_para.is_some_and(|np| is_text_empty(&np.runs));
    let bottom_collapses = next_borders_match
        && (next_has_top
            || para.borders.between.is_some()
            || text_empty
            || next_is_empty);

    let bdr_top_pad = if prev_borders_match {
        0.0
    } else {
        para.borders
            .top
            .as_ref()
            .map(|b| b.space_pt + b.width_pt / 2.0)
            .unwrap_or(0.0)
    };
    let bdr_bottom_pad = if bottom_collapses {
        0.0
    } else {
        para.borders
            .bottom
            .as_ref()
            .map(|b| b.space_pt + b.width_pt / 2.0)
            .unwrap_or(0.0)
    };
    // Full extent of bottom border below content (to border bottom edge)
    let bdr_bottom_extent = if bottom_collapses {
        0.0
    } else {
        para.borders
            .bottom
            .as_ref()
            .map(|b| b.space_pt + b.width_pt)
            .unwrap_or(0.0)
    };

    // Word measures the bottom border `space` attribute from
    // the full line-height content bottom, not from the text
    // descent.  No trailing-lead adjustment is needed.

    let needed = inter_gap + bdr_top_pad + content_h + bdr_bottom_extent;
    // For page-break decisions, also account for floating
    // images that extend below the text content.
    let needed_with_floats = needed.max(inter_gap + float_overflow_h);
    let at_page_top = state.pb.is_at_page_top(sp);

    // Word allows the last line's trailing inter-line
    // spacing to extend past the bottom margin — only the
    // text (ascent + descent) must fit inside the content
    // area.  Compute the excess leading so the page-break
    // check can tolerate it.
    let last_line_lead = if !lines.is_empty()
        && para.image.is_none()
        && para.inline_chart.is_none()
        && para.smartart.is_empty()
        && !matches!(effective_ls, LineSpacing::Exact(_))
    {
        let single_h = tallest_lhr
            .map(|r| font_size * r)
            .unwrap_or(font_size * 1.2);
        (line_h - single_h).max(0.0)
    } else {
        0.0
    };

    let keep_next_extra = if para.keep_next {
        let mut extra = 0.0;
        let mut prev_sa = effective_space_after;
        let mut i = block_idx + 1;
        while let Some(next) = adjacent_para(i) {
            if next.page_break_before {
                extra = f32::MAX;
                break;
            }
            let (nfs, nlhr, _) = tallest_run_metrics(&next.runs, ctx.fonts);
            let next_inter = f32::max(prev_sa, next.space_before);
            let next_first_line_h =
                nlhr.map(|ratio| nfs * ratio).unwrap_or(nfs * 1.2);
            if !next.keep_next {
                let next_ls = next.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let next_line_h = resolve_line_h(next_ls, nfs, nlhr);
                extra += next_inter + next_first_line_h + next_line_h;
                break;
            }
            if next.page_break_after {
                extra = f32::MAX;
                break;
            }
            extra += next_inter + next_first_line_h;
            prev_sa = next.space_after;
            i += 1;
        }
        extra
    } else {
        0.0
    };

    // Pre-compute footnote space for this paragraph so the
    // page-break check accounts for footnotes the paragraph
    // introduces (otherwise they're only tracked after
    // rendering, which can cause body/footnote overlap).
    let para_fn_extra = {
        let mut extra = 0.0f32;
        let mut seen_ids = HashSet::new();
        for run in para.runs.iter() {
            if let Some(id) = run.footnote_id {
                if !state.pb.footnote_ids_set.contains(&id)
                    && seen_ids.insert(id)
                {
                    if let Some(footnote) = doc.footnotes.get(&id) {
                        extra += compute_footnote_height(
                            footnote, &ctx, text_width,
                        );
                    }
                }
            }
        }
        if extra > 0.0 && state.pb.footnote_ids.is_empty() {
            extra += 12.0;
        }
        extra
    };

    if !at_page_top
        && state.pb.slot_top - needed_with_floats - keep_next_extra + last_line_lead
            < state.effective_margin_bottom + para_fn_extra
    {
        let available =
            state.pb.slot_top - inter_gap - state.effective_margin_bottom - para_fn_extra;
        let first_line_h = tallest_lhr
            .map(|ratio| font_size * ratio)
            .unwrap_or(font_size);
        let mut lines_that_fit = if line_h > 0.0 && available >= first_line_h {
            1 + ((available - first_line_h) / line_h).floor() as usize
        } else {
            0
        };

        if para.widow_control {
            // Ensure at least 2 lines remain on next page (orphan prevention)
            if lines_that_fit > 0
                && lines.len().saturating_sub(lines_that_fit) < 2
            {
                lines_that_fit = lines.len().saturating_sub(2);
            }
        }

        // keepLines: don't split — move entire paragraph to next column/page
        if para.keep_lines {
            lines_that_fit = 0;
        }

        let min_split_lines = if para.widow_control { 2 } else { 1 };
        if lines_that_fit >= min_split_lines && lines_that_fit < lines.len() {
            let first_part = &lines[..lines_that_fit];
            state.pb.slot_top -= inter_gap;
            let ascender_ratio = tallest_ar.unwrap_or(0.75);
            let baseline_offset = if grid_snapped {
                sp.line_pitch
            } else {
                font_size * ascender_ratio
            };
            let baseline_y = state.pb.slot_top - baseline_offset;

            render_list_label(
                &mut state.pb.content,
                para,
                ctx.fonts,
                label_x,
                baseline_y,
                font_size,
            );

            render_paragraph_lines(
                &mut state.pb.content,
                first_part,
                &para.alignment,
                para_text_x,
                para_text_width,
                baseline_y,
                line_h,
                lines.len(),
                0,
                &mut state.pb.links,
                text_hanging,
                ctx.fonts,
                poly_line_geom.as_deref(),
                &mut state.pb.gradient_specs,
                Some(&mut state.pb.comment_anchors),
            );

            state.pb.advance_column_or_page(
                &mut state.current_col,
                col_count,
                sect_idx,
                sp,
                &mut state.effective_margin_bottom,
                &ctx,
            );

            let rest = &lines[lines_that_fit..];
            let rest_content_h = rest.len() as f32 * line_h;
            let baseline_offset2 = if grid_snapped {
                sp.line_pitch
            } else {
                font_size * ascender_ratio
            };
            let baseline_y2 = state.pb.slot_top - baseline_offset2;

            let (rest_col_x, rest_col_w) = col_geometry[state.current_col];
            let rest_text_x = rest_col_x + para.indent_left;
            let rest_text_width =
                (rest_col_w - para.indent_left - para.indent_right).max(1.0);

            render_paragraph_lines(
                &mut state.pb.content,
                rest,
                &para.alignment,
                rest_text_x,
                rest_text_width,
                baseline_y2,
                line_h,
                lines.len(),
                lines_that_fit,
                &mut state.pb.links,
                text_hanging,
                ctx.fonts,
                None,
                &mut state.pb.gradient_specs,
                Some(&mut state.pb.comment_anchors),
            );

            state.pb.slot_top -= rest_content_h;
            state.prev_space_after = effective_space_after;

            // Track footnotes for the split paragraph on the new page
            for run in para.runs.iter() {
                if let Some(id) = run.footnote_id {
                    if state.pb.footnote_ids_set.insert(id) {
                        state.pb.footnote_ids.push(id);
                        if let Some(footnote) = doc.footnotes.get(&id) {
                            let fn_height =
                                compute_footnote_height(footnote, &ctx, text_width);
                            let separator_h = if state.pb.footnote_ids.len() == 1 {
                                12.0
                            } else {
                                0.0
                            };
                            state.effective_margin_bottom += separator_h + fn_height;
                        }
                    }
                }
                if let Some(id) = run.endnote_id {
                    if state.pb.endnote_ids_set.insert(id) {
                        state.pb.endnote_ids.push(id);
                    }
                }
            }

            state.global_block_idx += 1;
            return true;
        }

        state.pb.advance_column_or_page(
            &mut state.current_col,
            col_count,
            sect_idx,
            sp,
            &mut state.effective_margin_bottom,
            &ctx,
        );
        inter_gap = 0.0;
    }

    // Suppress space_before at the top of a page
    let at_new_page_top = !state.pb.all_contents.is_empty() && state.pb.is_at_page_top(sp);
    if at_new_page_top {
        if state.pb.is_first_page_of_section {
            // Section break: collapse with the previous section's trailing space_after
            inter_gap = (effective_space_before - state.prev_space_after).max(0.0);
        } else {
            inter_gap = 0.0;
        }
    }

    let applied_inter_gap = inter_gap;
    state.pb.slot_top -= inter_gap;

    for bookmark in &para.bookmarks {
        state.bookmark_positions
            .insert(bookmark.clone(), (state.pb.all_contents.len(), state.pb.slot_top));
    }

    if let Some(level) = para.outline_level {
        let title: String = para.runs.iter().map(|r| r.text.as_str()).collect();
        if !title.trim().is_empty() {
            state.heading_entries.push(HeadingEntry {
                title: title.trim().to_string(),
                level,
                page_idx: state.pb.all_contents.len(),
                y_position: state.pb.slot_top,
            });
        }
    }

    // Re-fetch column geometry (may have changed after overflow)
    let (col_x, col_w) = col_geometry[state.current_col];
    para_text_x = col_x + para.indent_left;
    para_text_width = (col_w - para.indent_left - para.indent_right).max(1.0);
    label_x = col_x + para.indent_left - para.indent_hanging;

    // Re-apply float zone adjustment after potential column change
    if let Some(ref fz) = state.pb.float_zone {
        if state.pb.slot_top <= fz.top_y && state.pb.slot_top > fz.bottom_y {
            let col_right = col_x + col_w;
            let (ex_left, ex_right) = fz.exclusion_at_y(state.pb.slot_top);
            let space_right =
                col_right - (ex_right + fz.right_from_text);
            let space_left = (ex_left - fz.left_from_text) - col_x;

            if fz.wrap_text == WrapText::BothSides {
                let lw = (space_left - para.indent_left).max(0.0);
                let rw = (space_right - para.indent_right).max(0.0);
                if rw > lw {
                    let new_left = ex_right + fz.right_from_text;
                    para_text_width = rw.max(1.0);
                    para_text_x = new_left + para.indent_left;
                    label_x =
                        new_left + para.indent_left - para.indent_hanging;
                } else if lw > 0.0 {
                    para_text_width = lw.max(1.0);
                }
            } else {
                let use_right = match fz.wrap_text {
                    WrapText::Right => space_right >= 1.0,
                    WrapText::Left => !(space_left >= 1.0),
                    _ => space_right >= space_left && space_right >= 72.0,
                };
                let use_left = match fz.wrap_text {
                    WrapText::Left => space_left >= 1.0,
                    WrapText::Right => false,
                    _ => space_left >= 72.0,
                };
                if use_right {
                    let new_left = ex_right + fz.right_from_text;
                    para_text_width =
                        (col_right - new_left - para.indent_right).max(1.0);
                    para_text_x = new_left + para.indent_left;
                    label_x =
                        new_left + para.indent_left - para.indent_hanging;
                } else if use_left {
                    let avail_right = ex_left - fz.left_from_text;
                    para_text_width = (avail_right - col_x
                        - para.indent_left
                        - para.indent_right)
                        .max(1.0);
                }
            }
        }
    }

    // Render behind-doc layer: floating images + textboxes
    render_floating_images(
        &para.floating_images,
        true,
        state.global_block_idx,
        &floating_image_pdf_names,
        &effect_floating_names,
        sp,
        col_x,
        col_w,
        text_width,
        state.pb.slot_top,
        &mut state.pb.content,
    );
    for tb in sorted_by_z(para.textboxes.iter().filter(|t| t.behind_doc)) {
        render_single_textbox(
            tb,
            sp,
            col_x,
            col_w,
            text_width,
            state.pb.slot_top,
            &mut state.pb.content,
            &mut state.pb.gradient_specs,
            &ctx,
            &mut state.pb.links,
        );
    }

    // Draw paragraph shading (background), extending outward to match borders
    if let Some(shd_color) = para.shading {
        let shd_left_outset = para
            .borders
            .left
            .as_ref()
            .map(|b| b.space_pt)
            .unwrap_or(0.0);
        let shd_right_outset = para
            .borders
            .right
            .as_ref()
            .map(|b| b.space_pt)
            .unwrap_or(0.0);
        let shd_left = col_x - shd_left_outset;
        let shd_right = col_x + col_w + shd_right_outset;
        let shd_top = state.pb.slot_top
            + if prev_borders_match { applied_inter_gap } else { 0.0 };
        let shd_bottom =
            state.pb.slot_top - bdr_top_pad - content_h - bdr_bottom_pad;
        state.pb.content.save_state();
        fill_rgb(&mut state.pb.content, shd_color);
        state.pb.content.rect(
            shd_left,
            shd_bottom,
            shd_right - shd_left,
            shd_top - shd_bottom,
        );
        state.pb.content.fill_nonzero();
        state.pb.content.restore_state();
    }

    // Render foreground layer: floating images + textboxes
    render_floating_images(
        &para.floating_images,
        false,
        state.global_block_idx,
        &floating_image_pdf_names,
        &effect_floating_names,
        sp,
        col_x,
        col_w,
        text_width,
        state.pb.slot_top,
        &mut state.pb.content,
    );

    // Set FloatZone for wrapping floating images
    // (may already be set by self-wrapping above; overwrite
    // to ensure polygon data is included).
    for fi in &para.floating_images {
        match fi.wrap_type {
            WrapType::Square | WrapType::Tight | WrapType::Through => {
                let fi_x =
                    resolve_fi_x(fi, sp, col_x, col_w, text_width);
                let fi_y_top =
                    resolve_fi_y_top(fi, sp, state.pb.slot_top);
                let fi_y_bottom =
                    fi_y_top - fi.image.display_height;
                let polygon_pts =
                    fi.wrap_polygon.as_ref().map(|verts| {
                        convert_polygon_to_page_coords(
                            verts,
                            fi_x,
                            fi_y_top,
                            fi.image.display_width,
                            fi.image.display_height,
                        )
                    });
                state.pb.float_zone = Some(FloatZone {
                    top_y: fi_y_top + fi.dist_top,
                    bottom_y: fi_y_bottom - fi.dist_bottom,
                    obj_left: fi_x,
                    obj_right: fi_x + fi.image.display_width,
                    left_from_text: fi.dist_left,
                    right_from_text: fi.dist_right,
                    polygon_pts,
                    wrap_text: fi.wrap_text,
                    para_relative: fi.v_relative_from == VRelativeFrom::Paragraph,
                });
            }
            _ => {}
        }
    }

    if debug_wrap {
        if let Some(ref fz) = state.pb.float_zone {
            draw_debug_wrap_overlay(&mut state.pb.content, fz);
        }
    }

    for tb in para.textboxes.iter().filter(|t| !t.behind_doc) {
        // Render into a per-shape buffer; flush_page paints these above the
        // page's text layer sorted by relativeHeight (Word's z-order for
        // floating shapes spans paragraphs).
        let mut shape_content = Content::new();
        render_single_textbox(
            tb,
            sp,
            col_x,
            col_w,
            text_width,
            state.pb.slot_top,
            &mut shape_content,
            &mut state.pb.gradient_specs,
            &ctx,
            &mut state.pb.links,
        );
        state.pb.deferred_shapes.push((tb.z_index, shape_content));
    }

    for conn in &para.connectors {
        // Same page-level z-stack as textboxes — anchored connectors must
        // interleave with shapes by relativeHeight (e.g. letter strokes
        // drawn over gradient circles)
        let mut shape_content = Content::new();
        render_connector(conn, &mut shape_content, col_x, state.pb.slot_top);
        state.pb.deferred_shapes.push((conn.z_index, shape_content));
    }

    if let Some(ref ic) = para.inline_chart {
        let chart_x = col_x
            + match para.alignment {
                Alignment::Center => (col_w - ic.display_width).max(0.0) / 2.0,
                Alignment::Right => (col_w - ic.display_width).max(0.0),
                _ => 0.0,
            };
        charts::render_chart(
            ic,
            &mut state.pb.content,
            chart_x,
            state.pb.slot_top,
            ctx.fonts,
            ctx.chart_font_name,
            &mut state.pb.alpha_states,
        );
    } else if !para.smartart.is_empty() {
        for (i, diagram) in para.smartart.iter().enumerate() {
            if i > 0 {
                state.pb.slot_top -= diagram.display_height;
            }
            smartart::render_smartart(
                &mut state.pb.content,
                diagram,
                col_x,
                state.pb.slot_top,
                ctx.fonts,
                smartart_font_key,
                &smartart_image_names,
            );
        }
    } else if let Some(ref hr) = para.horizontal_rule {
        let rule_w = col_w * hr.width_pct / 100.0;
        let rule_x = col_x
            + match para.alignment {
                Alignment::Center => (col_w - rule_w) / 2.0,
                Alignment::Right => col_w - rule_w,
                _ => 0.0,
            };
        // Standard HRs (o:hrstd) render as a thin 0.5pt line
        // centered in the specified height space
        let draw_h = if hr.is_standard { 0.5 } else { hr.height_pt };
        let rule_y =
            state.pb.slot_top - (content_h - draw_h) / 2.0 - draw_h;
        state.pb.content.save_state();
        fill_rgb(&mut state.pb.content, hr.fill_color);
        state.pb.content
            .rect(rule_x, rule_y, rule_w, draw_h);
        state.pb.content.fill_nonzero();
        state.pb.content.restore_state();
    } else if para.image.is_some() && para.content_height > 0.0 {
        if let Some(pdf_name) = image_pdf_names.get(&state.global_block_idx) {
            let img = para.image.as_ref().unwrap();
            let y_bottom = state.pb.slot_top - img.layout_extra_top - img.display_height;
            let x = col_x
                + match para.alignment {
                    Alignment::Center => (col_w - img.display_width).max(0.0) / 2.0,
                    Alignment::Right => (col_w - img.display_width).max(0.0),
                    _ => 0.0,
                };
            let img_fx = effect_names.get(&state.global_block_idx);
            if let Some(ref shadow) = img.shadow {
                color::draw_image_shadow(
                    &mut state.pb.content, shadow, x, y_bottom,
                    img.display_width, img.display_height,
                    img_fx.and_then(|fx| fx.shadow.as_deref()),
                );
            }
            if let Some(ref glow) = img.glow {
                color::draw_image_glow(
                    &mut state.pb.content, glow, x, y_bottom,
                    img.display_width, img.display_height,
                    img_fx.and_then(|fx| fx.glow.as_deref()),
                );
            }
            smartart::render_image_with_clip(
                &mut state.pb.content, pdf_name, x, y_bottom,
                img.display_width, img.display_height,
                img.clip_geometry.as_ref(),
            );
            if let Some(sc) = img.stroke_color {
                smartart::stroke_image_border(
                    &mut state.pb.content, x, y_bottom,
                    img.display_width, img.display_height,
                    sc, img.stroke_width,
                    img.clip_geometry.as_ref(),
                );
            }
            // Post-image effects: inner shadow, reflection (drawn on top / below)
            if let Some(ref inner) = img.inner_shadow {
                color::draw_inner_shadow(
                    &mut state.pb.content, inner, x, y_bottom,
                    img.display_width, img.display_height,
                    img_fx.and_then(|fx| fx.inner_shadow.as_deref()),
                );
            }
            if let Some(ref refl) = img.reflection {
                color::draw_reflection(
                    &mut state.pb.content, refl, x, y_bottom,
                    img.display_width, img.display_height,
                    img_fx.and_then(|fx| fx.reflection.as_deref()),
                );
            }
        } else if para.image.is_some() {
            state.pb.content
                .set_fill_gray(0.5)
                .rect(col_x, state.pb.slot_top - content_h, col_w, content_h)
                .fill_nonzero()
                .set_fill_gray(0.0);
        }
    } else if !lines.is_empty() {
        let ascender_ratio = tallest_ar.unwrap_or(0.75);
        // When the document grid snaps line heights, align the first
        // baseline one linePitch below the slot top so text sits on
        // the grid rather than at a font-metric-dependent offset.
        let baseline_offset = if grid_snapped {
            sp.line_pitch
        } else {
            font_size * ascender_ratio
        };
        let baseline_y = state.pb.slot_top - bdr_top_pad - baseline_offset;

        render_list_label(
            &mut state.pb.content,
            para,
            ctx.fonts,
            label_x,
            baseline_y,
            font_size,
        );

        if let Some((split_at, _after_w)) = float_width_change {
            if split_at < lines.len() {
                // Render first part
                render_paragraph_lines(
                    &mut state.pb.content,
                    &lines[..split_at],
                    &para.alignment,
                    para_text_x,
                    para_text_width,
                    baseline_y,
                    line_h,
                    lines.len(),
                    0,
                    &mut state.pb.links,
                    text_hanging,
                    ctx.fonts,
                    poly_line_geom.as_deref(),
                    &mut state.pb.gradient_specs,
                    Some(&mut state.pb.comment_anchors),
                );
                // float_width_change only comes from the lookahead path,
                // which always sets lookahead_narrow in the same branch.
                let (after_x, after_w) = lookahead_narrow
                    .expect("lookahead_narrow set when float_width_change is Some");
                let below_baseline = baseline_y - split_at as f32 * line_h;
                render_paragraph_lines(
                    &mut state.pb.content,
                    &lines[split_at..],
                    &para.alignment,
                    after_x,
                    after_w,
                    below_baseline,
                    line_h,
                    lines.len(),
                    split_at,
                    &mut state.pb.links,
                    text_hanging,
                    ctx.fonts,
                    poly_line_geom.as_deref(),
                    &mut state.pb.gradient_specs,
                    Some(&mut state.pb.comment_anchors),
                );
            } else {
                // Split point beyond paragraph — render all at once
                render_paragraph_lines(
                    &mut state.pb.content,
                    &lines,
                    &para.alignment,
                    para_text_x,
                    para_text_width,
                    baseline_y,
                    line_h,
                    lines.len(),
                    0,
                    &mut state.pb.links,
                    text_hanging,
                    ctx.fonts,
                    poly_line_geom.as_deref(),
                    &mut state.pb.gradient_specs,
                    Some(&mut state.pb.comment_anchors),
                );
            }
        } else {
            render_paragraph_lines(
                &mut state.pb.content,
                &lines,
                &para.alignment,
                para_text_x,
                para_text_width,
                baseline_y,
                line_h,
                lines.len(),
                0,
                &mut state.pb.links,
                text_hanging,
                ctx.fonts,
                poly_line_geom.as_deref(),
                &mut state.pb.gradient_specs,
                Some(&mut state.pb.comment_anchors),
            );
        }
    }

    // Draw paragraph borders — left/right borders extend outward
    // from the text area so text inside stays aligned with text outside
    {
        let bdr = &para.borders;
        let box_top = state.pb.slot_top;
        let box_bottom =
            state.pb.slot_top - bdr_top_pad - content_h - bdr_bottom_pad;
        let bdr_left_outset = bdr
            .left
            .as_ref()
            .map(|b| b.space_pt + b.width_pt / 2.0)
            .unwrap_or(0.0);
        let bdr_right_outset = bdr
            .right
            .as_ref()
            .map(|b| b.space_pt + b.width_pt / 2.0)
            .unwrap_or(0.0);
        let box_left = col_x - bdr_left_outset;
        let box_right = col_x + col_w + bdr_right_outset;

        // Extend horizontal borders past the corners so they cover
        // the corner gap left by butt-capped vertical border strokes.
        let h_left_ext = bdr.left.as_ref().map(|b| b.width_pt / 2.0).unwrap_or(0.0);
        let h_right_ext = bdr.right.as_ref().map(|b| b.width_pt / 2.0).unwrap_or(0.0);
        let draw_h_border = |content: &mut Content, b: &ParagraphBorder, y: f32| {
            content.save_state();
            content.set_line_width(b.width_pt);
            stroke_rgb(content, b.color);
            content.move_to(box_left - h_left_ext, y);
            content.line_to(box_right + h_right_ext, y);
            content.stroke();
            content.restore_state();
        };
        // When this paragraph continues a border group, extend vertical
        // borders upward through the inter-paragraph gap so there is no
        // visible break between consecutive paragraphs' left/right borders.
        let v_border_top = box_top
            + if prev_borders_match { applied_inter_gap } else { 0.0 };
        let draw_v_border = |content: &mut Content, b: &ParagraphBorder, x: f32| {
            content.save_state();
            content.set_line_width(b.width_pt);
            stroke_rgb(content, b.color);
            content.move_to(x, v_border_top);
            content.line_to(x, box_bottom);
            content.stroke();
            content.restore_state();
        };

        if !prev_borders_match {
            if let Some(b) = &bdr.top {
                draw_h_border(&mut state.pb.content, b, box_top);
            }
        }
        if bottom_collapses {
            if let Some(b) = &bdr.between {
                draw_h_border(&mut state.pb.content, b, box_bottom);
            }
        } else if let Some(b) = &bdr.bottom {
            draw_h_border(&mut state.pb.content, b, box_bottom);
        }
        if let Some(b) = &bdr.left {
            draw_v_border(&mut state.pb.content, b, box_left);
        }
        if let Some(b) = &bdr.right {
            draw_v_border(&mut state.pb.content, b, box_right);
        }
    }

    state.pb.slot_top -= content_h + bdr_top_pad + bdr_bottom_extent;
    if state.pb.slot_top < state.effective_margin_bottom - 1.0 {
        log::warn!(
            "Body overflow: slot_top={:.2} < eff_margin_bottom={:.2} after paragraph on page {}",
            state.pb.slot_top, state.effective_margin_bottom, state.pb.all_contents.len(),
        );
    }
    if !(text_empty && para.paragraph_mark_vanish) {
        state.prev_space_after = effective_space_after;
    }

    // Track footnotes referenced on this page
    for run in para.runs.iter() {
        if let Some(id) = run.footnote_id {
            if state.pb.footnote_ids_set.insert(id) {
                state.pb.footnote_ids.push(id);
                if let Some(footnote) = doc.footnotes.get(&id) {
                    let fn_height =
                        compute_footnote_height(footnote, &ctx, text_width);
                    let separator_h = if state.pb.footnote_ids.len() == 1 {
                        12.0
                    } else {
                        0.0
                    };
                    state.effective_margin_bottom += separator_h + fn_height;
                }
            }
        }
        // Endnotes render at end of document; just collect IDs in encounter
        // order — they're flushed to the last page in Phase 2c.
        if let Some(id) = run.endnote_id {
            if state.pb.endnote_ids_set.insert(id) {
                state.pb.endnote_ids.push(id);
            }
        }
    }

    update_styleref_from_para(
        &mut state.pb.styleref_running,
        &mut state.pb.styleref_page_first,
        para,
        &doc.style_id_to_name,
    );

    if para.page_break_after {
        state.pb.flush_page(sect_idx);
        state.pb.slot_top = effective_slot_top(sp, false, &ctx);
        state.pb.column_top_y = state.pb.slot_top;
        state.effective_margin_bottom =
            compute_effective_margin_bottom(sp, false, &ctx);
        state.pb.is_first_page_of_section = false;
        state.prev_space_after = 0.0;
        state.current_col = 0;
    }

    false
}

pub fn render(doc: &Document) -> Result<Vec<u8>, Error> {
    let debug_wrap = std::env::var("DOCXSIDE_DEBUG_WRAP").is_ok();
    let t0 = std::time::Instant::now();
    let mut pdf = Pdf::new();
    let mut next_id = 1i32;
    let mut alloc = || {
        let r = Ref::new(next_id);
        next_id += 1;
        r
    };

    let catalog_id = alloc();
    let pages_id = alloc();

    let (seen_fonts, font_order) = collect_and_register_fonts(doc, &mut pdf, &mut alloc);
    let smartart_font_key = font_order.first().map(|s| s.as_str()).unwrap_or("");
    let t_fonts = t0.elapsed();

    let EmbeddedImages {
        image_pdf_names,
        inline_image_pdf_names,
        floating_image_pdf_names,
        image_xobjects,
        hf_image_names,
        hf_inline_image_names,
        hf_floating_image_names,
        table_cell_image_names,
        textbox_image_names,
        smartart_image_names,
        effect_names,
        effect_floating_names,
        effect_inline_names,
        effect_hf_names,
        effect_hf_inline_names: _,
        effect_hf_floating_names,
        effect_table_names,
        effect_textbox_names,
    } = embed_all_images(doc, &mut pdf, &mut alloc);

    let ctx = RenderContext {
        fonts: &seen_fonts,
        doc_line_spacing: doc.line_spacing,
        default_tab_stop: doc.default_tab_stop,
        table_cell_image_names: &table_cell_image_names,
        effect_table_names: &effect_table_names,
        textbox_image_names: &textbox_image_names,
        effect_textbox_names: &effect_textbox_names,
        chart_font_name: &doc.chart_font_name,
    };

    let t_images = t0.elapsed();

    // Pre-compute footnote and endnote display order: scan body runs for
    // footnote_id / endnote_id, assign sequential numbers in encounter order.
    let mut footnote_display_order: HashMap<u32, u32> = HashMap::new();
    let mut endnote_display_order: HashMap<u32, u32> = HashMap::new();
    {
        let mut next_fn_num = 1u32;
        let mut next_en_num = 1u32;
        for section in &doc.sections {
            for block in &section.blocks {
                let runs: Box<dyn Iterator<Item = &Run>> = match block {
                    Block::Paragraph(p) => Box::new(p.runs.iter()),
                    Block::Table(t) => Box::new(
                        t.rows
                            .iter()
                            .flat_map(|row| row.cells.iter())
                            .flat_map(|cell| cell.all_paragraphs())
                            .flat_map(|p| p.runs.iter()),
                    ),
                };
                for run in runs {
                    if let Some(id) = run.footnote_id {
                        if !footnote_display_order.contains_key(&id) {
                            footnote_display_order.insert(id, next_fn_num);
                            next_fn_num += 1;
                        }
                    }
                    if let Some(id) = run.endnote_id {
                        if !endnote_display_order.contains_key(&id) {
                            endnote_display_order.insert(id, next_en_num);
                            next_en_num += 1;
                        }
                    }
                }
            }
        }
    }

    let bookmark_positions = compute_bookmark_positions(doc, &ctx);

    // Phase 2: build multi-page content streams (section-aware)
    let first_sp = &doc.sections[0].properties;
    let mut cur_sp = first_sp;
    let initial_slot_top = effective_slot_top(cur_sp, true, &ctx);
    let mut state = LayoutState {
        pb: PageBuilder::new(initial_slot_top),
        prev_space_after: 0.0,
        effective_margin_bottom: compute_effective_margin_bottom(cur_sp, true, &ctx),
        current_col: 0,
        global_block_idx: 0,
        heading_entries: Vec::new(),
        bookmark_positions,
    };

    for (sect_idx, section) in doc.sections.iter().enumerate() {
        let sp = &section.properties;

        // Section break handling (not for the first section)
        if sect_idx > 0 {
            match sp.break_type {
                SectionBreakType::NextPage
                | SectionBreakType::OddPage
                | SectionBreakType::EvenPage => {
                    state.pb.flush_page(sect_idx - 1);

                    // Insert blank page for odd/even page alignment
                    let need_odd = match sp.break_type {
                        SectionBreakType::OddPage => true,
                        _ if doc.even_and_odd_headers && sp.page_num_start.is_some() => {
                            sp.page_num_start.unwrap() % 2 == 1
                        }
                        _ => false,
                    };
                    let need_even = match sp.break_type {
                        SectionBreakType::EvenPage => true,
                        _ if doc.even_and_odd_headers && sp.page_num_start.is_some() => {
                            sp.page_num_start.unwrap() % 2 == 0
                        }
                        _ => false,
                    };
                    if need_odd || need_even {
                        // For an explicit OddPage/EvenPage section break, parity refers to the
                        // new section's LOGICAL page number (pgNumType w:start): a restarted
                        // section already begins at that number, so a filler page is only
                        // needed when its parity is wrong (without a restart, numbering
                        // continues and the physical index is the right proxy).
                        //
                        // The evenAndOddHeaders alignment heuristic also sets need_odd/need_even
                        // (on a NextPage break, from page_num_start parity) but there the goal
                        // is to land the section on the correct PHYSICAL sheet for even/odd
                        // header selection — so it must keep using the physical index.
                        let explicit_parity_break = matches!(
                            sp.break_type,
                            SectionBreakType::OddPage | SectionBreakType::EvenPage
                        );
                        let parity_ref = if explicit_parity_break {
                            sp.page_num_start
                                .map(|s| s as usize)
                                .unwrap_or_else(|| state.pb.page_count() + 1)
                        } else {
                            state.pb.page_count() + 1
                        };
                        if (need_odd && parity_ref % 2 == 0) || (need_even && parity_ref % 2 == 1) {
                            state.pb.push_blank_page(sect_idx - 1);
                        }
                    }

                    state.pb.slot_top = effective_slot_top(sp, true, &ctx);
                    state.pb.column_top_y = state.pb.slot_top;
                    state.effective_margin_bottom = compute_effective_margin_bottom(sp, true, &ctx);
                    state.pb.page_hf_section = sect_idx;
                    state.pb.is_first_page_of_section = true;
                }
                SectionBreakType::Continuous => {
                    // No forced break; geometry updates on next page.
                    // Don't update page_hf_section — the current page keeps
                    // the section that started it for header/footer purposes.
                }
            }
        }

        cur_sp = sp;
        let text_width = sp.page_width - sp.margin_left - sp.margin_right;

        // Column geometry: vec of (x_offset, width) for each column
        let col_config = sp.columns.as_ref();
        let col_count = col_config.map(|c| c.columns.len()).unwrap_or(1);
        let col_geometry: Vec<(f32, f32)> = if let Some(cfg) = col_config {
            let mut x = sp.margin_left;
            cfg.columns
                .iter()
                .map(|col| {
                    let result = (x, col.width);
                    x += col.width + col.space;
                    result
                })
                .collect()
        } else {
            vec![(sp.margin_left, text_width)]
        };
        state.current_col = 0;
        // Record the starting y for this section's columns on the current
        // page. For a mid-page continuous section, both columns begin at the
        // same y rather than at the top of the page.
        state.pb.column_top_y = state.pb.slot_top;

        for (block_idx, block) in section.blocks.iter().enumerate() {
            // If a float zone is active, decide whether to wrap text beside
            // the object or push it below.
            if let Some(ref fz) = state.pb.float_zone {
                if state.pb.slot_top <= fz.bottom_y {
                    // Already past the zone — clear it
                    state.pb.float_zone = None;
                } else if state.pb.slot_top <= fz.top_y
                    || (fz.para_relative && state.pb.slot_top <= fz.top_y + 30.0)
                {
                    // Cursor is within, entering, or (for paragraph-relative
                    // zones) slightly above the zone.  Paragraph-relative
                    // images with a positive vertical offset create zones that
                    // start below the anchor paragraph; the next paragraph's
                    // cursor may still be above the zone top.
                    let (col_x, col_w) = col_geometry[state.current_col];
                    let (ex_left, ex_right) = fz.exclusion_at_y(state.pb.slot_top);
                    let space_right = (col_x + col_w) - (ex_right + fz.right_from_text);
                    let space_left = (ex_left - fz.left_from_text) - col_x;
                    let min_wrap_w: f32 = 72.0;
                    let enough_space = if fz.wrap_text == WrapText::BothSides {
                        // For bothSides, check combined width of both regions
                        (space_left + space_right) >= min_wrap_w
                    } else {
                        let best_side = match fz.wrap_text {
                            WrapText::Left => space_left,
                            WrapText::Right => space_right,
                            _ => space_right.max(space_left),
                        };
                        best_side >= min_wrap_w
                    };
                    if !enough_space {
                        // Empty paragraphs can be absorbed within a wide
                        // image's vertical extent without needing wrap space.
                        // Include paragraphs with only line breaks (w:br)
                        // as "empty" — they have no visible text content.
                        let is_empty_para = matches!(block,
                            Block::Paragraph(p) if p.runs.iter().all(|r|
                                r.vanish || r.is_line_break
                                || (r.text.is_empty() && !r.is_tab && r.inline_image.is_none())
                            )
                                && p.image.is_none()
                                && p.inline_chart.is_none()
                                && p.smartart.is_empty()
                        );
                        if !is_empty_para {
                            state.pb.slot_top = fz.bottom_y;
                            state.pb.float_zone = None;
                        }
                    }
                    // Otherwise leave zone active — paragraph layout adjusts width
                }
            }

            match block {
                Block::Paragraph(para) => {
                    if render_paragraph_block(
                        para,
                        &mut state,
                        &ctx,
                        cur_sp,
                        &col_geometry,
                        col_count,
                        text_width,
                        sect_idx,
                        block_idx,
                        &section.blocks,
                        &floating_image_pdf_names,
                        &inline_image_pdf_names,
                        &image_pdf_names,
                        &effect_names,
                        &effect_floating_names,
                        &effect_inline_names,
                        &footnote_display_order,
                        &endnote_display_order,
                        doc,
                        smartart_font_key,
                        &smartart_image_names,
                        debug_wrap,
                    ) {
                        continue;
                    }
                }

                Block::Table(table) => {
                    let override_pos = table.position.as_ref().map(|pos| {
                        let table_total_w: f32 = table.col_widths.iter().sum();
                        let x = match pos.h_anchor {
                            "page" => match pos.h_position {
                                HorizontalPosition::AlignCenter => {
                                    (sp.page_width - table_total_w) / 2.0
                                }
                                HorizontalPosition::AlignRight => sp.page_width - table_total_w,
                                HorizontalPosition::AlignLeft => 0.0,
                                HorizontalPosition::Offset(o) => o,
                            },
                            "margin" => match pos.h_position {
                                HorizontalPosition::AlignCenter => {
                                    sp.margin_left + (text_width - table_total_w) / 2.0
                                }
                                HorizontalPosition::AlignRight => {
                                    sp.margin_left + text_width - table_total_w
                                }
                                HorizontalPosition::AlignLeft => sp.margin_left,
                                HorizontalPosition::Offset(o) => sp.margin_left + o,
                            },
                            _ => {
                                let (col_x, col_w) = col_geometry[state.current_col];
                                match pos.h_position {
                                    HorizontalPosition::AlignCenter => {
                                        col_x + (col_w - table_total_w) / 2.0
                                    }
                                    HorizontalPosition::AlignRight => col_x + col_w - table_total_w,
                                    HorizontalPosition::AlignLeft => col_x,
                                    HorizontalPosition::Offset(o) => col_x + o,
                                }
                            }
                        };
                        let y = match pos.v_anchor {
                            "page" => sp.page_height - pos.v_offset_pt,
                            "margin" => sp.page_height - sp.margin_top - pos.v_offset_pt,
                            _ => state.pb.slot_top - pos.v_offset_pt,
                        };
                        FloatingTablePos {
                            x,
                            y,
                            top_from_text: pos.top_from_text,
                            bottom_from_text: pos.bottom_from_text,
                            left_from_text: pos.left_from_text,
                            right_from_text: pos.right_from_text,
                        }
                    });
                    let col_bounds = if col_count > 1 {
                        Some(col_geometry[state.current_col])
                    } else {
                        None
                    };
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
                    state.prev_space_after = 0.0;

                    // Update styleref tracking (footnotes are already tracked
                    // inside render_table incrementally per row).
                    for row in &table.rows {
                        for cell in &row.cells {
                            for p in cell.all_paragraphs() {
                                update_styleref_from_para(
                                    &mut state.pb.styleref_running,
                                    &mut state.pb.styleref_page_first,
                                    p,
                                    &doc.style_id_to_name,
                                );
                            }
                        }
                    }
                }
            }
            // Clear float zone once cursor passes below it
            if let Some(ref fz) = state.pb.float_zone {
                if state.pb.slot_top <= fz.bottom_y {
                    state.pb.float_zone = None;
                }
            }

            state.global_block_idx += 1;
        }
    }
    state.pb.flush_page(doc.sections.len() - 1);

    let t_layout = t0.elapsed();

    // Phase 2b: column separator lines
    for (page_idx, content) in state.pb.all_contents.iter_mut().enumerate() {
        let (.., si) = state.pb.page_section_indices[page_idx];
        let sp = &doc.sections[si].properties;

        if let Some(cfg) = &sp.columns {
            if cfg.sep {
                let mut x = sp.margin_left;
                for (i, col) in cfg.columns.iter().enumerate() {
                    x += col.width;
                    if i < cfg.columns.len() - 1 {
                        let mid_x = x + col.space / 2.0;
                        content.save_state();
                        content.set_line_width(0.5);
                        content.move_to(mid_x, sp.margin_bottom);
                        content.line_to(mid_x, sp.page_height - sp.margin_top);
                        content.stroke();
                        content.restore_state();
                        x += col.space;
                    }
                }
            }
        }
    }

    // Phase 2c: render footnotes at page bottom (above footer). Endnotes
    // (default pos=docEnd) render once on the final page, stacked above any
    // footnotes on that page.
    let last_page_idx = state.pb.all_contents.len().saturating_sub(1);
    for (page_idx, content) in state.pb.all_contents.iter_mut().enumerate() {
        let (hf_si, is_first, si) = state.pb.page_section_indices[page_idx];
        let sp = &doc.sections[hf_si].properties;
        let eff_bottom = compute_effective_margin_bottom(sp, is_first, &ctx);
        let content_sp = &doc.sections[si].properties;
        let text_width = content_sp.page_width - content_sp.margin_left - content_sp.margin_right;
        let mut bottom = eff_bottom;
        render_page_footnotes(
            content,
            &state.pb.all_footnote_ids[page_idx],
            &doc.footnotes,
            &footnote_display_order,
            &ctx,
            content_sp.margin_left,
            bottom,
            text_width,
            &mut state.pb.all_gradient_specs[page_idx],
        );
        if page_idx == last_page_idx && !state.pb.endnote_ids.is_empty() {
            // Stack endnotes above any footnotes on the final page.
            let fn_height: f32 = state.pb.all_footnote_ids[page_idx]
                .iter()
                .filter_map(|id| doc.footnotes.get(id))
                .map(|fn_note| compute_footnote_height(fn_note, &ctx, text_width))
                .sum();
            if fn_height > 0.0 {
                bottom += fn_height + 12.0;
            }
            render_page_footnotes(
                content,
                &state.pb.endnote_ids,
                &doc.endnotes,
                &endnote_display_order,
                &ctx,
                content_sp.margin_left,
                bottom,
                text_width,
                &mut state.pb.all_gradient_specs[page_idx],
            );
        }
    }

    let t_headers = t0.elapsed();

    // Phase 2d: render headers/footers into separate content streams (behind body)
    let total_pages = state.pb.all_contents.len();

    // Pre-index header/footer image maps by (section_index, hf_type)
    // Fields: (para_images, inline_images, floating_images, effect_para, effect_floating)
    type HfMaps = (
        HashMap<usize, String>,
        HashMap<(usize, usize), String>,
        HashMap<(usize, usize), String>,
        HashMap<usize, EffectXObjs>,
        HashMap<(usize, usize), EffectXObjs>,
    );
    let mut hf_maps_index: HashMap<(usize, u8), HfMaps> = HashMap::new();
    for ((s, t, pi), name) in &hf_image_names {
        hf_maps_index.entry((*s, *t)).or_default().0.insert(*pi, name.clone());
    }
    for ((s, t, pi, ri), name) in &hf_inline_image_names {
        hf_maps_index.entry((*s, *t)).or_default().1.insert((*pi, *ri), name.clone());
    }
    for ((s, t, pi, fi), name) in &hf_floating_image_names {
        hf_maps_index.entry((*s, *t)).or_default().2.insert((*pi, *fi), name.clone());
    }
    for ((s, t, pi), fx) in &effect_hf_names {
        hf_maps_index.entry((*s, *t)).or_default().3.insert(*pi, fx.clone());
    }
    for ((s, t, pi, fi), fx) in &effect_hf_floating_names {
        hf_maps_index.entry((*s, *t)).or_default().4.insert((*pi, *fi), fx.clone());
    }
    let empty_hf_maps: HfMaps = Default::default();

    // Pre-compute page numbers and formats: sections without w:pgNumType @start
    // continue numbering from the previous section, but the format never
    // inherits — fmt applies only to its own section and an omitted fmt means
    // decimal (OOXML §17.6.12; Word renders arabic after roman front matter)
    let mut page_numbers: Vec<usize> = Vec::with_capacity(total_pages);
    // Track which section's format applies to each page (None = decimal default).
    // Uses a section index to avoid cloning the format string for every page.
    let mut page_format_sources: Vec<Option<usize>> = Vec::with_capacity(total_pages);
    {
        let mut running_num: usize = 0;
        let mut running_format_si: Option<usize> = None;
        let mut prev_content_si: Option<usize> = None;
        for page_idx in 0..total_pages {
            let (_, _, content_si) = state.pb.page_section_indices[page_idx];
            let csp = &doc.sections[content_si].properties;
            if prev_content_si != Some(content_si) {
                // New section boundary
                running_format_si = if csp.page_num_format.is_some() {
                    Some(content_si)
                } else {
                    None
                };
                if let Some(start) = csp.page_num_start {
                    running_num = start as usize;
                } else {
                    running_num += 1;
                }
            } else {
                running_num += 1;
            }
            page_numbers.push(running_num);
            page_format_sources.push(running_format_si);
            prev_content_si = Some(content_si);
        }
    }

    let empty_styleref: HashMap<String, String> = HashMap::new();
    let mut page_styleref_merged: HashMap<String, String> = HashMap::new();
    let mut all_hf_contents: Vec<Option<Content>> = (0..total_pages).map(|_| None).collect();
    for (page_idx, hf_content) in all_hf_contents.iter_mut().enumerate() {
        let (si, is_first, _content_si) = state.pb.page_section_indices[page_idx];
        let sp = &doc.sections[si].properties;

        let page_num = page_numbers[page_idx];
        let effective_page_num_format = page_format_sources[page_idx]
            .and_then(|si| doc.sections[si].properties.page_num_format.as_deref());

        // Per spec §17.16.5.59: in headers/footers of a printed document, STYLEREF
        // searches the current page top-to-bottom first, then backward to doc start.
        let page_first = state.pb
            .all_first_styleref
            .get(page_idx)
            .unwrap_or(&empty_styleref);
        let prev_running = if page_idx > 0 {
            state.pb.all_styleref.get(page_idx - 1).unwrap_or(&empty_styleref)
        } else {
            &empty_styleref
        };
        page_styleref_merged.clone_from(prev_running);
        // Current-page first occurrences take priority (top-to-bottom search)
        for (k, v) in page_first {
            page_styleref_merged.insert(k.clone(), v.clone());
        }
        let page_styleref = &page_styleref_merged;

        let mut hf = Content::new();
        let mut has_hf = false;

        let (header, hdr_type, hdr_si) =
            resolve_header_for_page(doc, si, is_first, page_num);
        if let Some(header_data) = header {
            let (pi_map, ii_map, fi_map, sh_para, sh_float) = hf_maps_index.get(&(hdr_si, hdr_type)).unwrap_or(&empty_hf_maps);
            let pc = HfPageContext {
                page_num, total_pages,
                para_image_names: pi_map, inline_image_names: ii_map,
                floating_image_names: fi_map,
                effect_para_names: sh_para, effect_floating_names: sh_float,
                styleref_values: page_styleref,
                page_num_format: effective_page_num_format,
            };
            render_header_footer(
                &mut hf, header_data, &ctx, sp, true, &pc,
                &mut state.pb.all_gradient_specs[page_idx],
            );
            has_hf = true;
        }

        let (footer, ftr_type, ftr_si) =
            resolve_footer_for_page(doc, si, is_first, page_num);
        if let Some(footer_data) = footer {
            let (pi_map, ii_map, fi_map, sh_para, sh_float) = hf_maps_index.get(&(ftr_si, ftr_type)).unwrap_or(&empty_hf_maps);
            let pc = HfPageContext {
                page_num, total_pages,
                para_image_names: pi_map, inline_image_names: ii_map,
                floating_image_names: fi_map,
                effect_para_names: sh_para, effect_floating_names: sh_float,
                styleref_values: page_styleref,
                page_num_format: effective_page_num_format,
            };
            render_header_footer(
                &mut hf, footer_data, &ctx, sp, false, &pc,
                &mut state.pb.all_gradient_specs[page_idx],
            );
            has_hf = true;
        }

        if has_hf {
            *hf_content = Some(hf);
        }
    }

    assemble_pdf_pages(
        &mut pdf,
        &mut alloc,
        catalog_id,
        pages_id,
        state.pb.all_contents,
        state.pb.all_deferred_shapes,
        &mut all_hf_contents,
        &state.pb.all_links,
        &state.pb.all_comment_anchors,
        &state.pb.all_alpha_states,
        &state.pb.all_gradient_specs,
        &state.pb.page_section_indices,
        ctx.fonts,
        &font_order,
        &image_xobjects,
        doc,
        &state.bookmark_positions,
        &state.heading_entries,
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

    Ok(pdf.finish())
}
