mod assembly;
mod chart_legend;
mod charts;
mod charts_radial;
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
    Run, SectionBreakType, SectionProperties, ShapeFill, ShapeGeometry, TextAnchor, Textbox,
    VRelativeFrom, VerticalPosition, WrapText, WrapType,
};

use assembly::{HeadingEntry, assemble_pdf_pages};
use fonts::collect_and_register_fonts;
use footnotes::{compute_footnote_height, render_page_footnotes};
use header_footer::{
    compute_effective_margin_bottom, effective_slot_top, render_header_footer,
};
pub(super) use helpers::resolve_line_h;
use helpers::borders_match;
use positioning::{render_connector, render_floating_images, resolve_fi_x};
pub(super) use positioning::{resolve_h_position, resolve_fi_y_top};
use images::{EmbeddedImages, embed_all_images};
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

pub(super) struct RenderContext<'a> {
    pub(super) fonts: &'a HashMap<String, FontEntry>,
    pub(super) doc_line_spacing: LineSpacing,
    pub(super) default_tab_stop: f32,
    /// Image names for inline images in table cells, keyed by Arc data pointer address.
    pub(super) table_cell_image_names: &'a HashMap<usize, String>,
    pub(super) chart_font_name: &'a str,
}

pub(super) struct GradientSpec {
    pattern_name: String,
    stops: Vec<([u8; 3], f32)>,
    angle_deg: f32,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
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
    pub(super) footnote_ids: Vec<u32>,
    pub(super) alpha_states: HashSet<u8>,
    pub(super) gradient_specs: Vec<GradientSpec>,

    // Cross-page running state
    styleref_running: HashMap<String, String>,
    styleref_page_first: HashMap<String, String>,

    // Layout position state
    pub(super) slot_top: f32,
    pub(super) is_first_page_of_section: bool,
    /// Section that owns the current page for header/footer purposes.
    /// For continuous section breaks, this stays as the section that started
    /// the page, not the section that continues mid-page.
    page_hf_section: usize,
    /// Floating table exclusion zone on this page; paragraph layout
    /// uses horizontal bounds to decide wrap-beside vs push-below.
    pub(super) float_zone: Option<FloatZone>,

    // Accumulated pages
    all_contents: Vec<Content>,
    all_links: Vec<Vec<LinkAnnotation>>,
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
            footnote_ids: Vec::new(),
            alpha_states: HashSet::new(),
            gradient_specs: Vec::new(),
            styleref_running: HashMap::new(),
            styleref_page_first: HashMap::new(),
            slot_top,
            is_first_page_of_section: true,
            page_hf_section: 0,
            float_zone: None,
            all_contents: Vec::new(),
            all_links: Vec::new(),
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
        self.all_links.push(std::mem::take(&mut self.links));
        self.all_footnote_ids
            .push(std::mem::take(&mut self.footnote_ids));
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
        self.all_links.push(Vec::new());
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
            self.slot_top = effective_slot_top(sp, false, ctx);
        } else {
            *current_col = 0;
            self.flush_page(sect_idx);
            self.slot_top = effective_slot_top(sp, false, ctx);
            *effective_margin_bottom = compute_effective_margin_bottom(sp, false, ctx);
            self.is_first_page_of_section = false;
        }
    }
}


pub fn render(doc: &Document) -> Result<Vec<u8>, Error> {
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
    } = embed_all_images(doc, &mut pdf, &mut alloc);

    let ctx = RenderContext {
        fonts: &seen_fonts,
        doc_line_spacing: doc.line_spacing,
        default_tab_stop: doc.default_tab_stop,
        table_cell_image_names: &table_cell_image_names,
        chart_font_name: &doc.chart_font_name,
    };

    let t_images = t0.elapsed();

    // Pre-compute footnote display order: scan body runs for footnote_id, assign sequential numbers
    let mut footnote_display_order: HashMap<u32, u32> = HashMap::new();
    {
        let mut next_fn_num = 1u32;
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
                }
            }
        }
    }

    // Map bookmarks to page indices so PAGEREF fields (e.g. TOC) can show
    // correct page numbers for bookmarks that appear later in the document.
    // Uses real line-building for accurate paragraph heights.
    let has_pagerefs = doc.sections.iter().any(|s| s.blocks.iter().any(|b| match b {
        Block::Paragraph(p) => p.runs.iter().any(|r| matches!(&r.field_code, Some(FieldCode::PageRef(_)))),
        _ => false,
    }));
    let mut bookmark_positions: HashMap<String, (usize, f32)> = HashMap::new();
    if has_pagerefs {
        let mut page_idx = 0usize;
        let first_sp = &doc.sections[0].properties;
        let mut sp = first_sp;
        let mut slot_top = effective_slot_top(sp, true, &ctx);
        let mut margin_bottom = compute_effective_margin_bottom(sp, true, &ctx);
        let mut prev_space_after: f32 = 0.0;
        let mut prev_contextual: bool = false;
        let empty_imgs: HashMap<usize, String> = HashMap::new();

        for (si, section) in doc.sections.iter().enumerate() {
            sp = &section.properties;
            if si > 0 {
                match sp.break_type {
                    SectionBreakType::NextPage
                    | SectionBreakType::OddPage
                    | SectionBreakType::EvenPage => {
                        page_idx += 1;
                        slot_top = effective_slot_top(sp, true, &ctx);
                        margin_bottom = compute_effective_margin_bottom(sp, true, &ctx);
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
                        if para.page_break_before && slot_top < effective_slot_top(sp, false, &ctx) {
                            page_idx += 1;
                            slot_top = effective_slot_top(sp, false, &ctx);
                            margin_bottom = compute_effective_margin_bottom(sp, false, &ctx);
                            prev_space_after = 0.0;
                        }
                        for bm in &para.bookmarks {
                            bookmark_positions.insert(bm.clone(), (page_idx, slot_top));
                        }
                        if para.is_section_break && is_text_empty(&para.runs) {
                            continue;
                        }
                        let (font_size, tallest_lhr, _) =
                            tallest_run_metrics(&para.runs, ctx.fonts);
                        let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                        let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
                        let line_h = if para.snap_to_grid
                            && matches!(sp.grid_type, DocGridType::Lines | DocGridType::LinesAndChars | DocGridType::SnapToChars)
                            && !matches!(effective_ls, LineSpacing::Exact(_))
                            && sp.line_pitch > 0.0
                        {
                            (line_h / sp.line_pitch).ceil() * sp.line_pitch
                        } else {
                            line_h
                        };
                        let para_w = (text_width - para.indent_left - para.indent_right).max(1.0);
                        let hanging = if !para.list_label.is_empty() {
                            if let Some(nts) = para.num_level_tab_stop {
                                if nts < para.indent_left && (para.indent_left - para.indent_hanging).abs() < 0.5 {
                                    (para.indent_left - nts).max(0.0)
                                } else { 0.0 }
                            } else { 0.0 }
                        } else if para.indent_hanging > 0.0 {
                            para.indent_hanging
                        } else {
                            -para.indent_first_line
                        };
                        let has_tabs = para.runs.iter().any(|r| r.is_tab);
                        let lines = if is_text_empty(&para.runs) {
                            vec![]
                        } else if has_tabs {
                            build_tabbed_line(
                                &para.runs, ctx.fonts, &para.tab_stops,
                                para.indent_left, para_w, hanging,
                                &empty_imgs, doc.default_tab_stop,
                            )
                        } else {
                            build_paragraph_lines(
                                &para.runs, ctx.fonts, para_w,
                                hanging, &empty_imgs, None, None, None,
                            )
                        };
                        let num_lines = lines.len().max(1);
                        let content_h = if para.image.is_some() || para.inline_chart.is_some() {
                            para.content_height
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
                            && slot_top < effective_slot_top(sp, false, &ctx)
                        {
                            page_idx += 1;
                            slot_top = effective_slot_top(sp, false, &ctx);
                            margin_bottom = compute_effective_margin_bottom(sp, false, &ctx);
                            slot_top -= content_h;
                        } else {
                            slot_top -= inter_gap + content_h;
                        }
                        prev_space_after = effective_sa;
                        prev_contextual = para.contextual_spacing;
                    }
                    Block::Table(table) => {
                        // ~12pt default font at ~1.15x line spacing ≈ 14pt per line
                        let para_count: usize = table.rows.iter()
                            .flat_map(|r| r.cells.iter())
                            .map(|c| c.all_paragraphs().len().max(1))
                            .max()
                            .unwrap_or(1)
                            * table.rows.len();
                        let est_h = para_count as f32 * 14.0;
                        if slot_top - est_h < margin_bottom {
                            page_idx += 1;
                            slot_top = effective_slot_top(sp, false, &ctx);
                            margin_bottom = compute_effective_margin_bottom(sp, false, &ctx);
                        }
                        slot_top -= est_h;
                        prev_space_after = 0.0;
                        prev_contextual = false;
                    }
                }
            }
        }
    }

    // Phase 2: build multi-page content streams (section-aware)
    let first_sp = &doc.sections[0].properties;
    let mut cur_sp = first_sp;
    let initial_slot_top = effective_slot_top(cur_sp, true, &ctx);
    let mut pb = PageBuilder::new(initial_slot_top);
    let mut heading_entries: Vec<HeadingEntry> = Vec::new();
    let mut prev_space_after: f32 = 0.0;
    let mut effective_margin_bottom: f32 = compute_effective_margin_bottom(cur_sp, true, &ctx);
    let mut global_block_idx: usize = 0;

    for (sect_idx, section) in doc.sections.iter().enumerate() {
        let sp = &section.properties;

        // Section break handling (not for the first section)
        if sect_idx > 0 {
            match sp.break_type {
                SectionBreakType::NextPage
                | SectionBreakType::OddPage
                | SectionBreakType::EvenPage => {
                    pb.flush_page(sect_idx - 1);

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
                        let next_phys = pb.page_count() + 1;
                        let next_is_odd = next_phys % 2 == 1;
                        if (need_odd && !next_is_odd) || (need_even && next_is_odd) {
                            pb.push_blank_page(sect_idx - 1);
                        }
                    }

                    pb.slot_top = effective_slot_top(sp, true, &ctx);
                    effective_margin_bottom = compute_effective_margin_bottom(sp, true, &ctx);
                    pb.page_hf_section = sect_idx;
                    pb.is_first_page_of_section = true;
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
        let mut current_col: usize = 0;

        let adjacent_para = |idx: usize| -> Option<&Paragraph> {
            match section.blocks.get(idx)? {
                Block::Paragraph(p) => Some(p),
                Block::Table(_) => None,
            }
        };

        for (block_idx, block) in section.blocks.iter().enumerate() {
            // If a float zone is active, decide whether to wrap text beside
            // the object or push it below.
            if let Some(ref fz) = pb.float_zone {
                if pb.slot_top <= fz.bottom_y {
                    // Already past the zone — clear it
                    pb.float_zone = None;
                } else if pb.slot_top <= fz.top_y {
                    // Cursor is within or entering the zone — check horizontal space
                    let (col_x, col_w) = col_geometry[current_col];
                    let (ex_left, ex_right) = fz.exclusion_at_y(pb.slot_top);
                    let space_right = (col_x + col_w) - (ex_right + fz.right_from_text);
                    let space_left = (ex_left - fz.left_from_text) - col_x;
                    const MIN_WRAP_WIDTH: f32 = 72.0; // ~1 inch minimum
                    let enough_space = if fz.wrap_text == WrapText::BothSides {
                        // For bothSides, check combined width of both regions
                        (space_left + space_right) >= MIN_WRAP_WIDTH
                    } else {
                        let best_side = match fz.wrap_text {
                            WrapText::Left => space_left,
                            WrapText::Right => space_right,
                            _ => space_right.max(space_left),
                        };
                        best_side >= MIN_WRAP_WIDTH
                    };
                    if !enough_space {
                        // Not enough room — push below
                        pb.slot_top = fz.bottom_y;
                        pb.float_zone = None;
                    }
                    // Otherwise leave zone active — paragraph layout adjusts width
                }
            }

            match block {
                Block::Paragraph(para) => {
                    // Skip empty section-break paragraphs — Word gives these zero height
                    if para.is_section_break
                        && is_text_empty(&para.runs)
                        && para.image.is_none()
                        && para.inline_chart.is_none()
                        && para.smartart.is_none()
                        && para.floating_images.is_empty()
                        && para.textboxes.is_empty()
                    {
                        global_block_idx += 1;
                        continue;
                    }

                    // Handle explicit page breaks
                    if para.page_break_before {
                        let at_top = pb.is_at_page_top(cur_sp);
                        if !at_top {
                            pb.flush_page(sect_idx);
                            pb.slot_top = effective_slot_top(cur_sp, false, &ctx);
                            effective_margin_bottom =
                                compute_effective_margin_bottom(cur_sp, false, &ctx);
                            pb.is_first_page_of_section = false;
                            current_col = 0;
                        }
                        prev_space_after = 0.0;
                        if is_text_empty(&para.runs) {
                            global_block_idx += 1;
                            continue;
                        }
                    }

                    // Handle explicit column breaks
                    if para.column_break_before && col_count > 1 {
                        pb.advance_column_or_page(
                            &mut current_col,
                            col_count,
                            sect_idx,
                            cur_sp,
                            &mut effective_margin_bottom,
                            &ctx,
                        );
                        prev_space_after = 0.0;
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

                    let mut inter_gap = f32::max(prev_space_after, effective_space_before);

                    let (font_size, tallest_lhr, tallest_ar) =
                        tallest_run_metrics(&para.runs, ctx.fonts);
                    let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                    let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
                    let line_h = if para.snap_to_grid
                        && matches!(
                            cur_sp.grid_type,
                            DocGridType::Lines
                                | DocGridType::LinesAndChars
                                | DocGridType::SnapToChars
                        )
                        && !matches!(effective_ls, LineSpacing::Exact(_))
                        && cur_sp.line_pitch > 0.0
                    {
                        (line_h / cur_sp.line_pitch).ceil() * cur_sp.line_pitch
                    } else {
                        line_h
                    };

                    let (col_x, col_w) = col_geometry[current_col];
                    let mut para_text_x = col_x + para.indent_left;
                    let mut para_text_width =
                        (col_w - para.indent_left - para.indent_right).max(1.0);
                    let mut label_x = col_x + para.indent_left - para.indent_hanging;

                    // When inside a floating object zone, narrow the paragraph to
                    // fit beside the object rather than overlapping it.
                    if let Some(ref fz) = pb.float_zone {
                        if pb.slot_top <= fz.top_y && pb.slot_top > fz.bottom_y {
                            let col_right = col_x + col_w;
                            let (ex_left, ex_right) = fz.exclusion_at_y(pb.slot_top);
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

                    let text_hanging = if !para.list_label.is_empty() {
                        if let Some(nts) = para.num_level_tab_stop {
                            if nts < para.indent_left && (para.indent_left - para.indent_hanging).abs() < 0.5 {
                                (para.indent_left - nts).max(0.0)
                            } else {
                                0.0
                            }
                        } else {
                            0.0
                        }
                    } else if para.indent_hanging > 0.0 {
                        para.indent_hanging
                    } else {
                        -para.indent_first_line
                    };

                    // Substitute footnote refs and resolve PAGEREF fields
                    let has_footnote_refs = para.runs.iter().any(|r| r.footnote_id.is_some());
                    let has_pageref = para.runs.iter().any(|r| matches!(&r.field_code, Some(FieldCode::PageRef(_))));
                    let effective_runs: std::borrow::Cow<'_, Vec<Run>> = if has_footnote_refs || has_pageref {
                        let substituted: Vec<Run> = para
                            .runs
                            .iter()
                            .map(|run| {
                                if let Some(id) = run.footnote_id {
                                    let num = footnote_display_order.get(&id).copied().unwrap_or(0);
                                    let mut r = run.clone();
                                    r.text = num.to_string();
                                    r
                                } else if let Some(FieldCode::PageRef(ref bookmark)) = run.field_code {
                                    let mut r = run.clone();
                                    if let Some(&(page_idx, _)) = bookmark_positions.get(bookmark) {
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
                        .filter(|((bi, _), _)| *bi == global_block_idx)
                        .map(|((_, ri), name)| (*ri, name.clone()))
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
                                resolve_fi_y_top(fi, sp, pb.slot_top);
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
                            pb.float_zone = Some(FloatZone {
                                top_y: fi_y_top + fi.dist_top,
                                bottom_y: fi_y_bottom - fi.dist_bottom,
                                obj_left: fi_x,
                                obj_right: fi_x + fi.image.display_width,
                                left_from_text: fi.dist_left,
                                right_from_text: fi.dist_right,
                                polygon_pts,
                                wrap_text: fi.wrap_text,
                            });
                            // Re-narrow para_text_x / para_text_width using the
                            // new float zone (same logic as the block above).
                            let fz = pb.float_zone.as_ref().unwrap();
                            if pb.slot_top <= fz.top_y && pb.slot_top > fz.bottom_y
                            {
                                let col_right = col_x + col_w;
                                let (ex_left, ex_right) =
                                    fz.exclusion_at_y(pb.slot_top);
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

                    // Build per-line geometry and dual-region geometry
                    let (poly_line_geom, poly_dual_geom): (Option<Vec<(f32, f32)>>, Option<Vec<DualRegion>>) =
                        if let Some(fz) = pb.float_zone.as_ref() {
                            let eff_top = pb.slot_top - inter_gap;
                            if eff_top <= fz.bottom_y {
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
                                for i in 0..max_lines {
                                    // Use line slot top for zone check (not baseline).
                                    let line_top = eff_top - i as f32 * line_h;
                                    if line_top <= fz.top_y
                                        && line_top > bottom_threshold
                                    {
                                        let (ex_left, ex_right) =
                                            fz.exclusion_at_y(line_top);
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
                    let lines = if para.image.is_some() || text_empty {
                        vec![]
                    } else if has_tabs {
                        build_tabbed_line(
                            &effective_runs,
                            ctx.fonts,
                            &para.tab_stops,
                            para.indent_left,
                            para_text_width,
                            text_hanging,
                            &block_inline_images,
                            doc.default_tab_stop,
                        )
                    } else {
                        // Look-ahead: if next block is an image-only paragraph
                        // with wrapping, build lines at full width first, then
                        // check if the bottom lines need narrowing.
                        let lookahead_fi = if pb.float_zone.is_none() {
                            section.blocks.get(block_idx + 1).and_then(|b| {
                                if let Block::Paragraph(np) = b {
                                    if !np.floating_images.is_empty()
                                        && is_text_empty(&np.runs)
                                        && np.image.is_none()
                                        && np.inline_chart.is_none()
                                    {
                                        np.floating_images.iter().find(|fi| matches!(
                                            fi.wrap_type,
                                            WrapType::Square | WrapType::Tight | WrapType::Through
                                        ))
                                    } else { None }
                                } else { None }
                            })
                        } else { None };

                        let (lines, final_width_change) = if let Some(fi) = lookahead_fi {
                            // Two-pass: build at full width, then narrow bottom lines
                            let full_lines = build_paragraph_lines(
                                &effective_runs, ctx.fonts, para_text_width,
                                text_hanging, &block_inline_images, None, None, None,
                            );
                            let num_lines = full_lines.len();
                            let content_h_est = num_lines as f32 * line_h;
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
                                    text_hanging, &block_inline_images,
                                    Some((lines_above, narrow_w)), None, None,
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
                                text_hanging, &block_inline_images, None,
                                plw, poly_dual_geom.as_deref(),
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
                    } else if text_empty {
                        if para.paragraph_mark_vanish {
                            0.0
                        } else if para.content_height > 0.0 {
                            para.content_height
                        } else {
                            line_h
                        }
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
                    } else {
                        let num_lines = lines.len();
                        let first_line_h = if let Some(label_fs) = para.list_label_font_size {
                            if label_fs > font_size {
                                resolve_line_h(effective_ls, label_fs, tallest_lhr)
                            } else {
                                line_h
                            }
                        } else {
                            line_h
                        };
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
                            WrapType::Square | WrapType::Tight | WrapType::Through => {
                                fi.image.display_width >= text_width * 0.9
                            }
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
                            // Narrower paragraph-relative images: track overflow
                            // for page-break check only (text wraps beside them)
                            float_overflow_h = float_overflow_h.max(fi_h);
                        }
                    }

                    for tb in &para.textboxes {
                        let reserve = match tb.wrap_type {
                            WrapType::TopAndBottom => true,
                            WrapType::Square => tb.width_pt >= text_width * 0.9,
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

                    let bdr_top_pad = para
                        .borders
                        .top
                        .as_ref()
                        .map(|b| b.space_pt + b.width_pt / 2.0)
                        .unwrap_or(0.0);
                    let bdr_bottom_pad = para
                        .borders
                        .bottom
                        .as_ref()
                        .map(|b| b.space_pt + b.width_pt / 2.0)
                        .unwrap_or(0.0);
                    // Full extent of bottom border below content (to border bottom edge)
                    let bdr_bottom_extent = para
                        .borders
                        .bottom
                        .as_ref()
                        .map(|b| b.space_pt + b.width_pt)
                        .unwrap_or(0.0);

                    // Word measures the bottom border `space` attribute from
                    // the full line-height content bottom, not from the text
                    // descent.  No trailing-lead adjustment is needed.

                    let needed = inter_gap + bdr_top_pad + content_h + bdr_bottom_extent;
                    // For page-break decisions, also account for floating
                    // images that extend below the text content.
                    let needed_with_floats = needed.max(inter_gap + float_overflow_h);
                    let at_page_top = pb.is_at_page_top(cur_sp);

                    // Word allows the last line's trailing inter-line
                    // spacing to extend past the bottom margin — only the
                    // text (ascent + descent) must fit inside the content
                    // area.  Compute the excess leading so the page-break
                    // check can tolerate it.
                    let last_line_lead = if !lines.is_empty()
                        && para.image.is_none()
                        && para.inline_chart.is_none()
                        && para.smartart.is_none()
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

                    if !at_page_top
                        && pb.slot_top - needed_with_floats - keep_next_extra + last_line_lead
                            < effective_margin_bottom
                    {
                        let available = pb.slot_top - inter_gap - effective_margin_bottom;
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
                            pb.slot_top -= inter_gap;
                            let ascender_ratio = tallest_ar.unwrap_or(0.75);
                            let baseline_y = pb.slot_top - font_size * ascender_ratio;

                            render_list_label(
                                &mut pb.content,
                                para,
                                ctx.fonts,
                                label_x,
                                baseline_y,
                                font_size,
                            );

                            render_paragraph_lines(
                                &mut pb.content,
                                first_part,
                                &para.alignment,
                                para_text_x,
                                para_text_width,
                                baseline_y,
                                line_h,
                                lines.len(),
                                0,
                                &mut pb.links,
                                text_hanging,
                                ctx.fonts,
                                poly_line_geom.as_deref(),
                            );

                            pb.advance_column_or_page(
                                &mut current_col,
                                col_count,
                                sect_idx,
                                cur_sp,
                                &mut effective_margin_bottom,
                                &ctx,
                            );

                            let rest = &lines[lines_that_fit..];
                            let rest_content_h = rest.len() as f32 * line_h;
                            let baseline_y2 = pb.slot_top - font_size * ascender_ratio;

                            let (rest_col_x, rest_col_w) = col_geometry[current_col];
                            let rest_text_x = rest_col_x + para.indent_left;
                            let rest_text_width =
                                (rest_col_w - para.indent_left - para.indent_right).max(1.0);

                            render_paragraph_lines(
                                &mut pb.content,
                                rest,
                                &para.alignment,
                                rest_text_x,
                                rest_text_width,
                                baseline_y2,
                                line_h,
                                lines.len(),
                                lines_that_fit,
                                &mut pb.links,
                                text_hanging,
                                ctx.fonts,
                                None,
                            );

                            pb.slot_top -= rest_content_h;
                            prev_space_after = effective_space_after;
                            global_block_idx += 1;
                            continue;
                        }

                        pb.advance_column_or_page(
                            &mut current_col,
                            col_count,
                            sect_idx,
                            cur_sp,
                            &mut effective_margin_bottom,
                            &ctx,
                        );
                        inter_gap = 0.0;
                    }

                    // Suppress space_before at the top of a page
                    let at_new_page_top = !pb.all_contents.is_empty() && pb.is_at_page_top(cur_sp);
                    if at_new_page_top {
                        if pb.is_first_page_of_section {
                            // Section break: collapse with the previous section's trailing space_after
                            inter_gap = (effective_space_before - prev_space_after).max(0.0);
                        } else {
                            inter_gap = 0.0;
                        }
                    }

                    pb.slot_top -= inter_gap;

                    for bookmark in &para.bookmarks {
                        bookmark_positions
                            .insert(bookmark.clone(), (pb.all_contents.len(), pb.slot_top));
                    }

                    if let Some(level) = para.outline_level {
                        let title: String = para.runs.iter().map(|r| r.text.as_str()).collect();
                        if !title.trim().is_empty() {
                            heading_entries.push(HeadingEntry {
                                title: title.trim().to_string(),
                                level,
                                page_idx: pb.all_contents.len(),
                                y_position: pb.slot_top,
                            });
                        }
                    }

                    // Re-fetch column geometry (may have changed after overflow)
                    let (col_x, col_w) = col_geometry[current_col];
                    para_text_x = col_x + para.indent_left;
                    para_text_width = (col_w - para.indent_left - para.indent_right).max(1.0);
                    label_x = col_x + para.indent_left - para.indent_hanging;

                    // Re-apply float zone adjustment after potential column change
                    if let Some(ref fz) = pb.float_zone {
                        if pb.slot_top <= fz.top_y && pb.slot_top > fz.bottom_y {
                            let col_right = col_x + col_w;
                            let (ex_left, ex_right) = fz.exclusion_at_y(pb.slot_top);
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
                        global_block_idx,
                        &floating_image_pdf_names,
                        sp,
                        col_x,
                        col_w,
                        text_width,
                        pb.slot_top,
                        &mut pb.content,
                    );
                    for tb in para.textboxes.iter().filter(|t| t.behind_doc) {
                        render_single_textbox(
                            tb,
                            sp,
                            col_x,
                            col_w,
                            text_width,
                            pb.slot_top,
                            &mut pb.content,
                            &mut pb.gradient_specs,
                            &ctx,
                            &mut pb.links,
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
                        let shd_top = pb.slot_top;
                        let shd_bottom =
                            pb.slot_top - bdr_top_pad - content_h - bdr_bottom_pad;
                        pb.content.save_state();
                        fill_rgb(&mut pb.content, shd_color);
                        pb.content.rect(
                            shd_left,
                            shd_bottom,
                            shd_right - shd_left,
                            shd_top - shd_bottom,
                        );
                        pb.content.fill_nonzero();
                        pb.content.restore_state();
                    }

                    // Render foreground layer: floating images + textboxes
                    render_floating_images(
                        &para.floating_images,
                        false,
                        global_block_idx,
                        &floating_image_pdf_names,
                        sp,
                        col_x,
                        col_w,
                        text_width,
                        pb.slot_top,
                        &mut pb.content,
                    );

                    // Set FloatZone for wrapping floating images
                    // (may already be set by self-wrapping above; overwrite
                    // to ensure polygon data is included)
                    for fi in &para.floating_images {
                        match fi.wrap_type {
                            WrapType::Square | WrapType::Tight | WrapType::Through => {
                                let fi_x =
                                    resolve_fi_x(fi, sp, col_x, col_w, text_width);
                                let fi_y_top =
                                    resolve_fi_y_top(fi, sp, pb.slot_top);
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
                                pb.float_zone = Some(FloatZone {
                                    top_y: fi_y_top + fi.dist_top,
                                    bottom_y: fi_y_bottom - fi.dist_bottom,
                                    obj_left: fi_x,
                                    obj_right: fi_x + fi.image.display_width,
                                    left_from_text: fi.dist_left,
                                    right_from_text: fi.dist_right,
                                    polygon_pts,
                                    wrap_text: fi.wrap_text,
                                });
                            }
                            _ => {}
                        }
                    }

                    for tb in para.textboxes.iter().filter(|t| !t.behind_doc) {
                        render_single_textbox(
                            tb,
                            sp,
                            col_x,
                            col_w,
                            text_width,
                            pb.slot_top,
                            &mut pb.content,
                            &mut pb.gradient_specs,
                            &ctx,
                            &mut pb.links,
                        );
                    }

                    for conn in &para.connectors {
                        render_connector(conn, &mut pb.content, col_x, pb.slot_top);
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
                            &mut pb.content,
                            chart_x,
                            pb.slot_top,
                            ctx.fonts,
                            ctx.chart_font_name,
                            &mut pb.alpha_states,
                        );
                    } else if let Some(ref diagram) = para.smartart {
                        smartart::render_smartart(
                            &mut pb.content,
                            diagram,
                            col_x,
                            pb.slot_top,
                            ctx.fonts,
                            smartart_font_key,
                        );
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
                            pb.slot_top - (content_h - draw_h) / 2.0 - draw_h;
                        pb.content.save_state();
                        fill_rgb(&mut pb.content, hr.fill_color);
                        pb.content
                            .rect(rule_x, rule_y, rule_w, draw_h);
                        pb.content.fill_nonzero();
                        pb.content.restore_state();
                    } else if (para.image.is_some() || text_empty) && para.content_height > 0.0 {
                        if let Some(pdf_name) = image_pdf_names.get(&global_block_idx) {
                            let img = para.image.as_ref().unwrap();
                            let y_bottom = pb.slot_top - img.layout_extra_top - img.display_height;
                            let x = col_x
                                + match para.alignment {
                                    Alignment::Center => (col_w - img.display_width).max(0.0) / 2.0,
                                    Alignment::Right => (col_w - img.display_width).max(0.0),
                                    _ => 0.0,
                                };
                            pb.content.save_state();
                            pb.content.transform([
                                img.display_width,
                                0.0,
                                0.0,
                                img.display_height,
                                x,
                                y_bottom,
                            ]);
                            pb.content.x_object(Name(pdf_name.as_bytes()));
                            pb.content.restore_state();
                        } else if para.image.is_some() {
                            pb.content
                                .set_fill_gray(0.5)
                                .rect(col_x, pb.slot_top - content_h, col_w, content_h)
                                .fill_nonzero()
                                .set_fill_gray(0.0);
                        }
                    } else if !lines.is_empty() {
                        let ascender_ratio = tallest_ar.unwrap_or(0.75);
                        let baseline_y = pb.slot_top - bdr_top_pad - font_size * ascender_ratio;

                        render_list_label(
                            &mut pb.content,
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
                                    &mut pb.content,
                                    &lines[..split_at],
                                    &para.alignment,
                                    para_text_x,
                                    para_text_width,
                                    baseline_y,
                                    line_h,
                                    lines.len(),
                                    0,
                                    &mut pb.links,
                                    text_hanging,
                                    ctx.fonts,
                                    poly_line_geom.as_deref(),
                                );
                                // float_width_change only comes from the lookahead path,
                                // which always sets lookahead_narrow in the same branch.
                                let (after_x, after_w) = lookahead_narrow
                                    .expect("lookahead_narrow set when float_width_change is Some");
                                let below_baseline = baseline_y - split_at as f32 * line_h;
                                render_paragraph_lines(
                                    &mut pb.content,
                                    &lines[split_at..],
                                    &para.alignment,
                                    after_x,
                                    after_w,
                                    below_baseline,
                                    line_h,
                                    lines.len(),
                                    split_at,
                                    &mut pb.links,
                                    text_hanging,
                                    ctx.fonts,
                                    poly_line_geom.as_deref(),
                                );
                            } else {
                                // Split point beyond paragraph — render all at once
                                render_paragraph_lines(
                                    &mut pb.content,
                                    &lines,
                                    &para.alignment,
                                    para_text_x,
                                    para_text_width,
                                    baseline_y,
                                    line_h,
                                    lines.len(),
                                    0,
                                    &mut pb.links,
                                    text_hanging,
                                    ctx.fonts,
                                    poly_line_geom.as_deref(),
                                );
                            }
                        } else {
                            render_paragraph_lines(
                                &mut pb.content,
                                &lines,
                                &para.alignment,
                                para_text_x,
                                para_text_width,
                                baseline_y,
                                line_h,
                                lines.len(),
                                0,
                                &mut pb.links,
                                text_hanging,
                                ctx.fonts,
                                poly_line_geom.as_deref(),
                            );
                        }
                    }

                    // Draw paragraph borders — left/right borders extend outward
                    // from the text area so text inside stays aligned with text outside
                    {
                        let bdr = &para.borders;
                        let box_top = pb.slot_top;
                        let box_bottom =
                            pb.slot_top - bdr_top_pad - content_h - bdr_bottom_pad;
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

                        let draw_h_border = |content: &mut Content, b: &ParagraphBorder, y: f32| {
                            content.save_state();
                            content.set_line_width(b.width_pt);
                            stroke_rgb(content, b.color);
                            content.move_to(box_left, y);
                            content.line_to(box_right, y);
                            content.stroke();
                            content.restore_state();
                        };
                        let draw_v_border = |content: &mut Content, b: &ParagraphBorder, x: f32| {
                            content.save_state();
                            content.set_line_width(b.width_pt);
                            stroke_rgb(content, b.color);
                            content.move_to(x, box_top);
                            content.line_to(x, box_bottom);
                            content.stroke();
                            content.restore_state();
                        };

                        let prev_borders_match = prev_para
                            .is_some_and(|pp| borders_match(&pp.borders, &para.borders));
                        let next_borders_match = next_para
                            .is_some_and(|np| borders_match(&para.borders, &np.borders));

                        if !prev_borders_match {
                            if let Some(b) = &bdr.top {
                                draw_h_border(&mut pb.content, b, box_top);
                            }
                        }
                        if next_borders_match {
                            if let Some(b) = &bdr.between {
                                draw_h_border(&mut pb.content, b, box_bottom);
                            }
                        } else if let Some(b) = &bdr.bottom {
                            draw_h_border(&mut pb.content, b, box_bottom);
                        }
                        if let Some(b) = &bdr.left {
                            draw_v_border(&mut pb.content, b, box_left);
                        }
                        if let Some(b) = &bdr.right {
                            draw_v_border(&mut pb.content, b, box_right);
                        }
                    }

                    pb.slot_top -= content_h + bdr_top_pad + bdr_bottom_extent;
                    if !(text_empty && para.paragraph_mark_vanish) {
                        prev_space_after = effective_space_after;
                    }

                    // Track footnotes referenced on this page
                    for run in para.runs.iter() {
                        if let Some(id) = run.footnote_id {
                            if !pb.footnote_ids.contains(&id) {
                                pb.footnote_ids.push(id);
                                if let Some(footnote) = doc.footnotes.get(&id) {
                                    let fn_height =
                                        compute_footnote_height(footnote, &ctx, text_width);
                                    let separator_h = if pb.footnote_ids.len() == 1 {
                                        12.0
                                    } else {
                                        0.0
                                    };
                                    effective_margin_bottom += separator_h + fn_height;
                                }
                            }
                        }
                    }

                    update_styleref_from_para(
                        &mut pb.styleref_running,
                        &mut pb.styleref_page_first,
                        para,
                        &doc.style_id_to_name,
                    );

                    if para.page_break_after {
                        pb.flush_page(sect_idx);
                        pb.slot_top = effective_slot_top(cur_sp, false, &ctx);
                        effective_margin_bottom =
                            compute_effective_margin_bottom(cur_sp, false, &ctx);
                        pb.is_first_page_of_section = false;
                        prev_space_after = 0.0;
                        current_col = 0;
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
                                let (col_x, col_w) = col_geometry[current_col];
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
                            _ => pb.slot_top - pos.v_offset_pt,
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
                    render_table(
                        table,
                        sp,
                        &ctx,
                        &mut pb,
                        sect_idx,
                        prev_space_after,
                        override_pos,
                    );
                    prev_space_after = 0.0;

                    for row in &table.rows {
                        for cell in &row.cells {
                            for p in cell.all_paragraphs() {
                                update_styleref_from_para(
                                    &mut pb.styleref_running,
                                    &mut pb.styleref_page_first,
                                    p,
                                    &doc.style_id_to_name,
                                );
                            }
                        }
                    }
                }
            }
            // Clear float zone once cursor passes below it
            if let Some(ref fz) = pb.float_zone {
                if pb.slot_top <= fz.bottom_y {
                    pb.float_zone = None;
                }
            }

            global_block_idx += 1;
        }
    }
    pb.flush_page(doc.sections.len() - 1);

    let t_layout = t0.elapsed();

    // Phase 2b: column separator lines
    for (page_idx, content) in pb.all_contents.iter_mut().enumerate() {
        let (.., si) = pb.page_section_indices[page_idx];
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

    // Phase 2c: render footnotes at page bottom
    for (page_idx, content) in pb.all_contents.iter_mut().enumerate() {
        let (.., si) = pb.page_section_indices[page_idx];
        let sp = &doc.sections[si].properties;
        let text_width = sp.page_width - sp.margin_left - sp.margin_right;
        render_page_footnotes(
            content,
            &pb.all_footnote_ids[page_idx],
            &doc.footnotes,
            &footnote_display_order,
            &ctx,
            sp.margin_left,
            sp.margin_bottom,
            text_width,
        );
    }

    let t_headers = t0.elapsed();

    // Phase 2d: render headers/footers into separate content streams (behind body)
    let total_pages = pb.all_contents.len();
    let build_hf_maps = |si: usize,
                         hf_type: u8|
     -> (
        HashMap<usize, String>,
        HashMap<(usize, usize), String>,
        HashMap<(usize, usize), String>,
    ) {
        let pi_map: HashMap<usize, String> = hf_image_names
            .iter()
            .filter(|((s, t, _), _)| *s == si && *t == hf_type)
            .map(|((_, _, pi), name)| (*pi, name.clone()))
            .collect();
        let ii_map: HashMap<(usize, usize), String> = hf_inline_image_names
            .iter()
            .filter(|((s, t, _, _), _)| *s == si && *t == hf_type)
            .map(|((_, _, pi, ri), name)| ((*pi, *ri), name.clone()))
            .collect();
        let fi_map: HashMap<(usize, usize), String> = hf_floating_image_names
            .iter()
            .filter(|((s, t, _, _), _)| *s == si && *t == hf_type)
            .map(|((_, _, pi, fi), name)| ((*pi, *fi), name.clone()))
            .collect();
        (pi_map, ii_map, fi_map)
    };

    let empty_styleref: HashMap<String, String> = HashMap::new();
    let mut all_hf_contents: Vec<Option<Content>> = (0..total_pages).map(|_| None).collect();
    for (page_idx, hf_content) in all_hf_contents.iter_mut().enumerate() {
        let (si, is_first, content_si) = pb.page_section_indices[page_idx];
        let sp = &doc.sections[si].properties;

        // Page numbering uses content_section (the section being rendered),
        // not hf_section (which may differ for continuous section breaks)
        let num_sp = &doc.sections[content_si].properties;
        let page_num = if let Some(start) = num_sp.page_num_start {
            // Section specifies explicit start: count pages within this section
            let pages_before_in_section = pb.page_section_indices[..page_idx]
                .iter()
                .filter(|&&(_, _, cs)| cs == content_si)
                .count();
            start as usize + pages_before_in_section
        } else {
            // No explicit start: continue absolute numbering
            page_idx + 1
        };

        // Per spec §17.16.5.59: in headers/footers of a printed document, STYLEREF
        // searches the current page top-to-bottom first, then backward to doc start.
        let page_first = pb
            .all_first_styleref
            .get(page_idx)
            .unwrap_or(&empty_styleref);
        let prev_running = if page_idx > 0 {
            pb.all_styleref.get(page_idx - 1).unwrap_or(&empty_styleref)
        } else {
            &empty_styleref
        };
        let mut page_styleref_merged = prev_running.clone();
        // Current-page first occurrences take priority (top-to-bottom search)
        for (k, v) in page_first {
            page_styleref_merged.insert(k.clone(), v.clone());
        }
        let page_styleref = &page_styleref_merged;

        let mut hf = Content::new();
        let mut has_hf = false;

        // Resolve header with inheritance: fall back to previous sections
        let (header, hdr_type, hdr_si) = {
            let mut resolved = (None, 0u8, si);
            for idx in (0..=si).rev() {
                let s = &doc.sections[idx].properties;
                let (h, t) = if idx == si {
                    // Current section: respect is_first and even/odd
                    if is_first && s.different_first_page {
                        (s.header_first.as_ref(), 1u8)
                    } else if doc.even_and_odd_headers && page_num % 2 == 0 && s.header_even.is_some() {
                        (s.header_even.as_ref(), 4u8)
                    } else {
                        (s.header_default.as_ref(), 0u8)
                    }
                } else {
                    // Inherited section: always use default
                    (s.header_default.as_ref(), 0u8)
                };
                if h.is_some() {
                    resolved = (h, t, idx);
                    break;
                }
            }
            resolved
        };
        if let Some(header_data) = header {
            let (pi_map, ii_map, fi_map) = build_hf_maps(hdr_si, hdr_type);
            render_header_footer(
                &mut hf,
                header_data,
                &ctx,
                sp,
                true,
                page_num,
                total_pages,
                &pi_map,
                &ii_map,
                &fi_map,
                page_styleref,
                &mut pb.all_gradient_specs[page_idx],
            );
            has_hf = true;
        }

        // Resolve footer with inheritance: fall back to previous sections
        let (footer, ftr_type, ftr_si) = {
            let mut resolved = (None, 0u8, si);
            for idx in (0..=si).rev() {
                let s = &doc.sections[idx].properties;
                let (f, t) = if idx == si {
                    // Current section: respect is_first and even/odd
                    if is_first && s.different_first_page {
                        (s.footer_first.as_ref(), 3u8)
                    } else if doc.even_and_odd_headers && page_num % 2 == 0 && s.footer_even.is_some() {
                        (s.footer_even.as_ref(), 5u8)
                    } else {
                        (s.footer_default.as_ref(), 2u8)
                    }
                } else {
                    // Inherited section: always use default
                    (s.footer_default.as_ref(), 2u8)
                };
                if f.is_some() {
                    resolved = (f, t, idx);
                    break;
                }
            }
            resolved
        };
        if let Some(footer_data) = footer {
            let (pi_map, ii_map, fi_map) = build_hf_maps(ftr_si, ftr_type);
            render_header_footer(
                &mut hf,
                footer_data,
                &ctx,
                sp,
                false,
                page_num,
                total_pages,
                &pi_map,
                &ii_map,
                &fi_map,
                page_styleref,
                &mut pb.all_gradient_specs[page_idx],
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
        pb.all_contents,
        &mut all_hf_contents,
        &pb.all_links,
        &pb.all_alpha_states,
        &pb.all_gradient_specs,
        &pb.page_section_indices,
        ctx.fonts,
        &font_order,
        &image_xobjects,
        doc,
        &bookmark_positions,
        &heading_entries,
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
