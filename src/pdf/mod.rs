mod chart_legend;
mod charts;
mod charts_radial;
mod footnotes;
mod header_footer;
mod layout;
mod smartart;
mod table;

use std::collections::{HashMap, HashSet};

use pdf_writer::{Content, Filter, Name, Pdf, Rect, Ref, Str};

use crate::error::Error;
use crate::fonts::{
    FontEntry, encode_as_gids, font_key, font_key_buf, register_font, to_winansi_bytes,
};
use crate::model::{
    Alignment, Block, ConnectorShape, ConnectorType, Document, EmbeddedImage, FieldCode,
    FloatingImage, HRelativeFrom, HeaderFooter, HorizontalPosition, ImageFormat, LineSpacing,
    Paragraph, ParagraphBorder, ParagraphBorders, Run, SectionBreakType, SectionProperties,
    ShapeFill, ShapeGeometry, Table, TextAnchor, Textbox, VRelativeFrom, VerticalPosition,
    WrapType,
};

use footnotes::{compute_footnote_height, render_page_footnotes};
use header_footer::{
    compute_effective_margin_bottom, effective_slot_top, hf_paragraphs, render_header_footer,
};
use layout::{
    LinkAnnotation, build_paragraph_lines, build_tabbed_line, is_text_empty,
    render_paragraph_lines, tallest_run_metrics,
};
use smartart::draw_shape_path;
use table::render_table;

pub(super) struct RenderContext<'a> {
    pub(super) fonts: &'a HashMap<String, FontEntry>,
    pub(super) doc_line_spacing: LineSpacing,
    /// Image names for inline images in table cells, keyed by Arc data pointer address.
    pub(super) table_cell_image_names: &'a HashMap<usize, String>,
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
        ShapeFill::Solid([r, g, b]) => {
            content.save_state();
            content.set_fill_rgb(*r as f32 / 255.0, *g as f32 / 255.0, *b as f32 / 255.0);
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

fn resolve_h_position(
    h_relative_from: HRelativeFrom,
    h_position: &HorizontalPosition,
    obj_width: f32,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> f32 {
    let (origin, area_width) = match h_relative_from {
        HRelativeFrom::Page => (0.0, sp.page_width),
        HRelativeFrom::Column => (col_x, col_w),
        HRelativeFrom::Margin => (sp.margin_left, text_width),
    };
    match h_position {
        HorizontalPosition::AlignCenter => origin + (area_width - obj_width) / 2.0,
        HorizontalPosition::AlignRight => origin + area_width - obj_width,
        HorizontalPosition::AlignLeft => origin,
        HorizontalPosition::Offset(o) => origin + o,
    }
}

fn resolve_fi_x(
    fi: &FloatingImage,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> f32 {
    resolve_h_position(
        fi.h_relative_from,
        &fi.h_position,
        fi.image.display_width,
        sp,
        col_x,
        col_w,
        text_width,
    )
}

fn resolve_fi_y_top(fi: &FloatingImage, sp: &SectionProperties, slot_top: f32) -> f32 {
    let img = &fi.image;
    match fi.v_position {
        VerticalPosition::Offset(v_offset) => match fi.v_relative_from {
            VRelativeFrom::Page => sp.page_height - v_offset,
            VRelativeFrom::Margin | VRelativeFrom::TopMargin => {
                sp.page_height - sp.margin_top - v_offset
            }
            VRelativeFrom::Paragraph => slot_top - v_offset,
        },
        VerticalPosition::AlignTop => match fi.v_relative_from {
            VRelativeFrom::Page => sp.page_height,
            _ => sp.page_height - sp.margin_top,
        },
        VerticalPosition::AlignCenter => match fi.v_relative_from {
            VRelativeFrom::Page => (sp.page_height + img.display_height) / 2.0,
            _ => {
                let area = sp.page_height - sp.margin_top - sp.margin_bottom;
                sp.page_height - sp.margin_top - (area - img.display_height) / 2.0
            }
        },
        VerticalPosition::AlignBottom => match fi.v_relative_from {
            VRelativeFrom::Page => img.display_height,
            _ => sp.margin_bottom + img.display_height,
        },
    }
}

fn render_single_textbox(
    tb: &Textbox,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
    slot_top: f32,
    content: &mut Content,
    gradient_specs: &mut Vec<GradientSpec>,
    ctx: &RenderContext,
    page_links: &mut Vec<LinkAnnotation>,
) {
    let tb_x = resolve_h_position(
        tb.h_relative_from,
        &tb.h_position,
        tb.width_pt,
        sp,
        col_x,
        col_w,
        text_width,
    );
    let tb_y_top = match tb.v_relative_from {
        VRelativeFrom::Page => sp.page_height - tb.v_offset_pt,
        VRelativeFrom::Margin | VRelativeFrom::TopMargin => {
            sp.page_height - sp.margin_top - tb.v_offset_pt
        }
        VRelativeFrom::Paragraph => slot_top - tb.v_offset_pt,
    };

    if let Some(ref fill) = tb.fill {
        render_shape_fill(
            content,
            fill,
            tb_x,
            tb_y_top - tb.height_pt,
            tb.width_pt,
            tb.height_pt,
            &tb.shape_type,
            gradient_specs,
        );
    }

    if let Some(stroke) = tb.stroke_color {
        if tb.stroke_width > 0.0 {
            content.save_state();
            content.set_line_width(tb.stroke_width);
            content.set_stroke_rgb(
                stroke[0] as f32 / 255.0,
                stroke[1] as f32 / 255.0,
                stroke[2] as f32 / 255.0,
            );
            draw_shape_path(
                content,
                tb_x,
                tb_y_top - tb.height_pt,
                tb.width_pt,
                tb.height_pt,
                &tb.shape_type,
            );
            content.stroke();
            content.restore_state();
        }
    }

    let content_x = tb_x + tb.margin_left;
    let content_w = if tb.no_text_wrap {
        10000.0
    } else {
        (tb.width_pt - tb.margin_left - tb.margin_right).max(0.0)
    };

    let anchor_offset = match tb.text_anchor {
        TextAnchor::Top => 0.0,
        TextAnchor::Middle | TextAnchor::Bottom => {
            let empty_inline_imgs_pre: HashMap<usize, String> = HashMap::new();
            let mut total_h = 0.0f32;
            for tp in &tb.paragraphs {
                let tp_ls = tp.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let tp_text_w = (content_w - tp.indent_left - tp.indent_right).max(1.0);
                let text_hanging = if !tp.list_label.is_empty() {
                    0.0
                } else if tp.indent_hanging > 0.0 {
                    tp.indent_hanging
                } else {
                    -tp.indent_first_line
                };
                let has_tabs = tp.runs.iter().any(|r| r.is_tab);
                let lines = if has_tabs {
                    build_tabbed_line(
                        &tp.runs, ctx.fonts, &tp.tab_stops, tp.indent_left,
                        tp_text_w, text_hanging, &empty_inline_imgs_pre,
                    )
                } else {
                    build_paragraph_lines(
                        &tp.runs, ctx.fonts, tp_text_w, text_hanging, &empty_inline_imgs_pre,
                    )
                };
                let (fs, lhr, _) = tallest_run_metrics(&tp.runs, ctx.fonts);
                let lh = resolve_line_h(tp_ls, fs, lhr);
                let n = lines.len().max(1) as f32;
                total_h += tp.space_before + n * lh + tp.space_after;
            }
            let available = tb.height_pt - tb.margin_top - tb.margin_bottom;
            let gap = (available - total_h).max(0.0);
            match tb.text_anchor {
                TextAnchor::Middle => gap / 2.0,
                TextAnchor::Bottom => gap,
                TextAnchor::Top => 0.0,
            }
        }
    };

    let mut cursor_y = tb_y_top - tb.margin_top - anchor_offset;
    let empty_inline_imgs: HashMap<usize, String> = HashMap::new();
    for tp in &tb.paragraphs {
        let tp_ls = tp.line_spacing.unwrap_or(ctx.doc_line_spacing);
        let tp_text_x = content_x + tp.indent_left;
        let tp_text_w = (content_w - tp.indent_left - tp.indent_right).max(1.0);
        let text_hanging = if !tp.list_label.is_empty() {
            0.0
        } else if tp.indent_hanging > 0.0 {
            tp.indent_hanging
        } else {
            -tp.indent_first_line
        };
        let has_tabs = tp.runs.iter().any(|r| r.is_tab);
        let tb_lines = if has_tabs {
            build_tabbed_line(
                &tp.runs,
                ctx.fonts,
                &tp.tab_stops,
                tp.indent_left,
                tp_text_w,
                text_hanging,
                &empty_inline_imgs,
            )
        } else {
            build_paragraph_lines(
                &tp.runs,
                ctx.fonts,
                tp_text_w,
                text_hanging,
                &empty_inline_imgs,
            )
        };
        if tb_lines.is_empty() {
            let (fs, lhr, _) = tallest_run_metrics(&tp.runs, ctx.fonts);
            let lh = resolve_line_h(tp_ls, fs, lhr);
            cursor_y -= tp.space_before + lh + tp.space_after;
            continue;
        }
        let (tb_fs, tb_lhr, tb_ar) = tallest_run_metrics(&tp.runs, ctx.fonts);
        let tb_ascender = tb_ar.unwrap_or(0.75);
        let tb_line_h = resolve_line_h(tp_ls, tb_fs, tb_lhr);
        let tb_baseline = cursor_y - tp.space_before - tb_fs * tb_ascender;
        render_list_label(
            content,
            tp,
            ctx.fonts,
            content_x + tp.indent_left - tp.indent_hanging,
            tb_baseline,
            tb_fs,
        );
        render_paragraph_lines(
            content,
            &tb_lines,
            &tp.alignment,
            tp_text_x,
            tp_text_w,
            tb_baseline,
            tb_line_h,
            tb_lines.len(),
            0,
            page_links,
            0.0,
            ctx.fonts,
        );
        cursor_y -= tp.space_before + (tb_lines.len() as f32) * tb_line_h + tp.space_after;
    }
}

fn render_floating_images(
    floating_images: &[FloatingImage],
    behind_doc: bool,
    global_block_idx: usize,
    pdf_names: &HashMap<(usize, usize), String>,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
    slot_top: f32,
    content: &mut Content,
) {
    for (fi_idx, fi) in floating_images.iter().enumerate() {
        if fi.behind_doc != behind_doc {
            continue;
        }
        if let Some(pdf_name) = pdf_names.get(&(global_block_idx, fi_idx)) {
            let img = &fi.image;
            let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
            let fi_y_top = resolve_fi_y_top(fi, sp, slot_top);
            let fi_y_bottom = fi_y_top - img.display_height;
            content.save_state();
            content.transform([
                img.display_width,
                0.0,
                0.0,
                img.display_height,
                fi_x,
                fi_y_bottom,
            ]);
            content.x_object(Name(pdf_name.as_bytes()));
            content.restore_state();
        }
    }
}

fn render_connector(conn: &ConnectorShape, content: &mut Content, col_x: f32, slot_top: f32) {
    let cx = col_x + conn.x;
    let cy = slot_top - conn.y;

    content.save_state();
    content.set_stroke_rgb(
        conn.stroke_color[0] as f32 / 255.0,
        conn.stroke_color[1] as f32 / 255.0,
        conn.stroke_color[2] as f32 / 255.0,
    );
    content.set_line_width(conn.stroke_width);

    match &conn.connector_type {
        ConnectorType::Line { flip_h, flip_v } => {
            let (x0, y0, x1, y1) = match (*flip_h, *flip_v) {
                (false, false) => (cx, cy, cx + conn.width, cy - conn.height),
                (true, false) => (cx + conn.width, cy, cx, cy - conn.height),
                (false, true) => (cx, cy - conn.height, cx + conn.width, cy),
                (true, true) => (cx + conn.width, cy - conn.height, cx, cy),
            };
            content.move_to(x0, y0);
            content.line_to(x1, y1);
            content.stroke();
        }
        ConnectorType::Arc {
            start_angle,
            end_angle,
            rotation,
        } => {
            render_arc(
                content,
                cx,
                cy,
                conn.width,
                conn.height,
                *start_angle,
                *end_angle,
                *rotation,
            );
        }
    }

    content.restore_state();
}

fn render_arc(
    content: &mut Content,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    start_deg: f32,
    end_deg: f32,
    rotation_deg: f32,
) {
    let rx = w / 2.0;
    let ry = h / 2.0;
    let cx = x + rx;
    let cy = y - ry;

    // OOXML sweep: from start_deg (adj1) to end_deg (adj2), always positive
    let mut sweep_deg = end_deg - start_deg;
    if sweep_deg <= 0.0 {
        sweep_deg += 360.0;
    }

    // OOXML uses standard trig angles (0°=right, CCW positive) displayed
    // in y-down coords (so visually clockwise). For PDF y-up, negate angles.
    let math_start = (-(start_deg + rotation_deg)).to_radians();
    let total = -(sweep_deg).to_radians();

    if total.abs() < 0.001 {
        return;
    }

    // Approximate arc with cubic bezier segments (max 90° each)
    let n_segs = ((total.abs() / std::f32::consts::FRAC_PI_2).ceil() as usize).max(1);
    let step = total / n_segs as f32;

    let pt = |a: f32| -> (f32, f32) { (cx + rx * a.cos(), cy + ry * a.sin()) };

    let mut angle = math_start;
    let (sx, sy) = pt(angle);
    content.move_to(sx, sy);

    for _ in 0..n_segs {
        let a0 = angle;
        let a1 = angle + step;
        let alpha = 4.0 * (1.0 - (step / 2.0).cos()) / (step / 2.0).sin() / 3.0;
        let (x0, y0) = pt(a0);
        let (x3, y3) = pt(a1);
        let cp1x = x0 - alpha * rx * a0.sin();
        let cp1y = y0 + alpha * ry * a0.cos();
        let cp2x = x3 + alpha * rx * a1.sin();
        let cp2y = y3 - alpha * ry * a1.cos();
        content.cubic_to(cp1x, cp1y, cp2x, cp2y, x3, y3);
        angle = a1;
    }
    content.stroke();
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

fn border_eq(a: &Option<ParagraphBorder>, b: &Option<ParagraphBorder>) -> bool {
    match (a, b) {
        (None, None) => true,
        (Some(a), Some(b)) => a.width_pt == b.width_pt && a.color == b.color,
        _ => false,
    }
}

fn borders_match(a: &ParagraphBorders, b: &ParagraphBorders) -> bool {
    border_eq(&a.top, &b.top)
        && border_eq(&a.bottom, &b.bottom)
        && border_eq(&a.left, &b.left)
        && border_eq(&a.right, &b.right)
        && border_eq(&a.between, &b.between)
}

fn resolve_line_h(ls: LineSpacing, font_size: f32, tallest_lhr: Option<f32>) -> f32 {
    match ls {
        LineSpacing::Auto(mult) => tallest_lhr
            .map(|ratio| font_size * ratio * mult)
            .unwrap_or(font_size * 1.2 * mult),
        LineSpacing::Exact(pts) => pts,
        LineSpacing::AtLeast(min_pts) => {
            let natural = tallest_lhr
                .map(|ratio| font_size * ratio)
                .unwrap_or(font_size * 1.2);
            natural.max(min_pts)
        }
    }
}

fn para_runs_with_textboxes(para: &Paragraph) -> Vec<&Run> {
    let mut out: Vec<&Run> = para.runs.iter().collect();
    for tb in &para.textboxes {
        for tp in &tb.paragraphs {
            out.extend(para_runs_with_textboxes(tp));
        }
    }
    out
}

fn collect_paras(para: &Paragraph) -> Vec<&Paragraph> {
    let mut out = vec![para];
    for tb in &para.textboxes {
        for tp in &tb.paragraphs {
            out.extend(collect_paras(tp));
        }
    }
    out
}

struct EmbeddedImages {
    image_pdf_names: HashMap<usize, String>,
    inline_image_pdf_names: HashMap<(usize, usize), String>,
    floating_image_pdf_names: HashMap<(usize, usize), String>,
    image_xobjects: Vec<(String, Ref)>,
    hf_image_names: HashMap<(usize, u8, usize), String>,
    hf_inline_image_names: HashMap<(usize, u8, usize, usize), String>,
    hf_floating_image_names: HashMap<(usize, u8, usize, usize), String>,
    /// Images in table cell paragraphs, keyed by Arc data pointer address.
    table_cell_image_names: HashMap<usize, String>,
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

    // Accumulated pages
    all_contents: Vec<Content>,
    all_links: Vec<Vec<LinkAnnotation>>,
    all_footnote_ids: Vec<Vec<u32>>,
    all_alpha_states: Vec<HashSet<u8>>,
    all_gradient_specs: Vec<Vec<GradientSpec>>,
    page_section_indices: Vec<(usize, bool)>,
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
        self.page_section_indices
            .push((sect_idx, self.is_first_page_of_section));
        self.all_styleref.push(self.styleref_running.clone());
        self.all_first_styleref
            .push(std::mem::take(&mut self.styleref_page_first));
    }

    fn push_blank_page(&mut self, sect_idx: usize) {
        self.all_contents.push(Content::new());
        self.all_links.push(Vec::new());
        self.all_footnote_ids.push(Vec::new());
        self.all_alpha_states.push(HashSet::new());
        self.all_gradient_specs.push(Vec::new());
        self.page_section_indices.push((sect_idx, false));
        self.all_styleref.push(self.styleref_running.clone());
        self.all_first_styleref
            .push(std::mem::take(&mut self.styleref_page_first));
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

fn embed_single_image(
    img: &EmbeddedImage,
    image_xobjects: &mut Vec<(String, Ref)>,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> String {
    let xobj_ref = alloc();
    let pdf_name = format!("Im{}", image_xobjects.len() + 1);

    match img.format {
        ImageFormat::Jpeg => {
            let mut xobj = pdf.image_xobject(xobj_ref, &*img.data);
            xobj.filter(Filter::DctDecode);
            xobj.width(img.pixel_width as i32);
            xobj.height(img.pixel_height as i32);
            match img.jpeg_components {
                1 => xobj.color_space().device_gray(),
                4 => xobj.color_space().device_cmyk(),
                _ => xobj.color_space().device_rgb(),
            };
            xobj.bits_per_component(8);
            xobj.interpolate(true);
        }
        ImageFormat::Png => {
            let cursor = std::io::Cursor::new(img.data.as_slice());
            let reader = image::ImageReader::with_format(
                std::io::BufReader::new(cursor),
                image::ImageFormat::Png,
            );
            let decoded = match reader.decode() {
                Ok(d) => d,
                Err(e) => {
                    log::warn!("PNG decode failed: {e} — writing 1x1 placeholder");
                    let mut xobj = pdf.image_xobject(xobj_ref, &[255, 255, 255]);
                    xobj.width(1);
                    xobj.height(1);
                    xobj.color_space().device_rgb();
                    xobj.bits_per_component(8);
                    image_xobjects.push((pdf_name.clone(), xobj_ref));
                    return pdf_name;
                }
            };
            let rgba: image::RgbaImage = decoded.to_rgba8();
            let (w, h) = (rgba.width(), rgba.height());
            let has_alpha = rgba.pixels().any(|p| p.0[3] < 255);

            let rgb_data: Vec<u8> = rgba
                .pixels()
                .flat_map(|p| [p.0[0], p.0[1], p.0[2]])
                .collect();
            let compressed_rgb = miniz_oxide::deflate::compress_to_vec_zlib(&rgb_data, 6);

            let smask_ref = if has_alpha {
                let alpha_data: Vec<u8> = rgba.pixels().map(|p| p.0[3]).collect();
                let compressed_alpha = miniz_oxide::deflate::compress_to_vec_zlib(&alpha_data, 6);
                let mask_ref = alloc();
                let mut mask = pdf.image_xobject(mask_ref, &compressed_alpha);
                mask.filter(Filter::FlateDecode);
                mask.width(w as i32);
                mask.height(h as i32);
                mask.color_space().device_gray();
                mask.bits_per_component(8);
                Some(mask_ref)
            } else {
                None
            };

            let mut xobj = pdf.image_xobject(xobj_ref, &compressed_rgb);
            xobj.filter(Filter::FlateDecode);
            xobj.width(w as i32);
            xobj.height(h as i32);
            xobj.color_space().device_rgb();
            xobj.bits_per_component(8);
            xobj.interpolate(true);
            if let Some(mask_ref) = smask_ref {
                xobj.s_mask(mask_ref);
            }
        }
    }

    image_xobjects.push((pdf_name.clone(), xobj_ref));
    pdf_name
}

fn collect_all_runs(doc: &Document) -> Vec<&Run> {
    let hf_runs = doc.sections.iter().flat_map(|s| {
        [
            &s.properties.header_default,
            &s.properties.header_first,
            &s.properties.header_even,
            &s.properties.footer_default,
            &s.properties.footer_first,
            &s.properties.footer_even,
        ]
        .into_iter()
        .filter_map(|hf| hf.as_ref())
        .flat_map(|hf| hf_paragraphs(hf))
        .flat_map(|p| p.runs.iter())
    });

    let footnote_runs = doc
        .footnotes
        .values()
        .flat_map(|fn_| fn_.paragraphs.iter())
        .flat_map(|p| p.runs.iter());

    doc.sections
        .iter()
        .flat_map(|s| s.blocks.iter())
        .flat_map(|block| -> Vec<&Run> {
            match block {
                Block::Paragraph(para) => para_runs_with_textboxes(para),
                Block::Table(table) => table
                    .rows
                    .iter()
                    .flat_map(|row| row.cells.iter())
                    .flat_map(|cell| cell.paragraphs.iter())
                    .flat_map(|para| para_runs_with_textboxes(para))
                    .collect(),
            }
        })
        .chain(hf_runs)
        .chain(footnote_runs)
        .collect()
}

fn collect_used_chars(doc: &Document, all_runs: &[&Run]) -> HashMap<String, HashSet<char>> {
    let mut used: HashMap<String, HashSet<char>> = HashMap::new();
    let mut key_buf = String::new();

    for run in all_runs {
        let key = font_key_buf(run, &mut key_buf);
        let chars = used.entry(key.to_string()).or_default();
        if run.caps || run.small_caps {
            chars.extend(run.text.to_uppercase().chars());
        } else {
            chars.extend(run.text.chars());
        }
        if let Some(ref fc) = run.field_code {
            match fc {
                FieldCode::Page | FieldCode::NumPages => {
                    chars.extend('0'..='9');
                }
                FieldCode::StyleRef(_) => {}
            }
        }
        if run.footnote_id.is_some() || run.is_footnote_ref_mark {
            chars.extend('0'..='9');
        }
    }

    let all_paras: Vec<&Paragraph> = doc
        .sections
        .iter()
        .flat_map(|s| s.blocks.iter())
        .flat_map(|block| -> Vec<&Paragraph> {
            match block {
                Block::Paragraph(p) => collect_paras(p),
                Block::Table(t) => t
                    .rows
                    .iter()
                    .flat_map(|row| row.cells.iter())
                    .flat_map(|cell| cell.paragraphs.iter())
                    .flat_map(|p| collect_paras(p))
                    .collect(),
            }
        })
        .collect();

    for para in &all_paras {
        if !para.list_label.is_empty() {
            if let Some(key) = label_font_key(para) {
                used.entry(key)
                    .or_default()
                    .extend(para.list_label.chars());
            }
        }
        for stop in &para.tab_stops {
            if let Some(leader_char) = stop.leader
                && let Some(run) = para.runs.first()
            {
                let key = font_key_buf(run, &mut key_buf).to_string();
                used.entry(key).or_default().insert(leader_char);
            }
        }
    }

    if let Some(first_run) = all_runs.first() {
        let sa_key = font_key_buf(first_run, &mut key_buf).to_string();
        let chars = used.entry(sa_key).or_default();
        for para in &all_paras {
            if let Some(ref diagram) = para.smartart {
                for shape in &diagram.shapes {
                    chars.extend(shape.text.chars());
                }
            }
        }
    }

    for section in &doc.sections {
        for hf in [
            &section.properties.header_default,
            &section.properties.header_first,
            &section.properties.header_even,
            &section.properties.footer_default,
            &section.properties.footer_first,
            &section.properties.footer_even,
        ]
        .into_iter()
        .flatten()
        {
            for para in hf_paragraphs(hf) {
                for run in &para.runs {
                    let key = font_key_buf(run, &mut key_buf);
                    let chars = used.entry(key.to_string()).or_default();
                    if run.caps || run.small_caps {
                        chars.extend(run.text.to_uppercase().chars());
                    } else {
                        chars.extend(run.text.chars());
                    }
                    if let Some(ref fc) = run.field_code {
                        match fc {
                            FieldCode::Page | FieldCode::NumPages => {
                                chars.extend('0'..='9');
                            }
                            FieldCode::StyleRef(_) => {
                                chars.extend('0'..='9');
                                chars.extend('A'..='Z');
                                chars.extend('a'..='z');
                                chars.extend([' ', '.', ',', '/', '-', '(', ')']);
                            }
                        }
                    }
                }
            }
        }
    }

    for chars in used.values_mut() {
        chars.insert(' ');
    }

    used
}

fn collect_and_register_fonts(
    doc: &Document,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> (HashMap<String, FontEntry>, Vec<String>) {
    let mut seen_fonts: HashMap<String, FontEntry> = HashMap::new();
    let mut font_order: Vec<String> = Vec::new();
    let all_runs = collect_all_runs(doc);
    let used_chars_per_font = collect_used_chars(doc, &all_runs);
    let mut key_buf = String::new();

    for run in &all_runs {
        let key = font_key_buf(run, &mut key_buf);
        if !seen_fonts.contains_key(key) {
            let key_owned = key.to_string();
            let pdf_name = format!("F{}", font_order.len() + 1);
            let used = used_chars_per_font
                .get(&key_owned)
                .cloned()
                .unwrap_or_default();
            let entry = register_font(
                pdf,
                &run.font_name,
                run.bold,
                run.italic,
                pdf_name,
                alloc,
                &doc.embedded_fonts,
                &used,
                &doc.font_table,
            );
            font_order.push(key_owned.clone());
            seen_fonts.insert(key_owned, entry);
        }
    }

    for (key, used) in &used_chars_per_font {
        if !seen_fonts.contains_key(key) {
            let pdf_name = format!("F{}", font_order.len() + 1);
            let entry = register_font(
                pdf,
                key,
                false,
                false,
                pdf_name,
                alloc,
                &doc.embedded_fonts,
                used,
                &doc.font_table,
            );
            seen_fonts.insert(key.clone(), entry);
            font_order.push(key.clone());
        }
    }

    if seen_fonts.is_empty() {
        let pdf_name = "F1".to_string();
        let entry = register_font(
            pdf,
            "Helvetica",
            false,
            false,
            pdf_name,
            alloc,
            &doc.embedded_fonts,
            &HashSet::new(),
            &doc.font_table,
        );
        seen_fonts.insert("Helvetica".to_string(), entry);
        font_order.push("Helvetica".to_string());
    }

    (seen_fonts, font_order)
}

fn embed_all_images(
    doc: &Document,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> EmbeddedImages {
    let mut image_pdf_names: HashMap<usize, String> = HashMap::new();
    let mut inline_image_pdf_names: HashMap<(usize, usize), String> = HashMap::new();
    let mut image_xobjects: Vec<(String, Ref)> = Vec::new();
    let mut floating_image_pdf_names: HashMap<(usize, usize), String> = HashMap::new();

    {
        let mut global_block_idx = 0usize;
        for section in &doc.sections {
            for block in &section.blocks {
                if let Block::Paragraph(para) = block {
                    if let Some(img) = &para.image {
                        let name = embed_single_image(img, &mut image_xobjects, pdf, alloc);
                        image_pdf_names.insert(global_block_idx, name);
                    }
                    for (run_idx, run) in para.runs.iter().enumerate() {
                        if let Some(img) = &run.inline_image {
                            let name = embed_single_image(img, &mut image_xobjects, pdf, alloc);
                            inline_image_pdf_names.insert((global_block_idx, run_idx), name);
                        }
                    }
                    for (fi_idx, fi) in para.floating_images.iter().enumerate() {
                        let name = embed_single_image(&fi.image, &mut image_xobjects, pdf, alloc);
                        floating_image_pdf_names.insert((global_block_idx, fi_idx), name);
                    }
                }
                global_block_idx += 1;
            }
        }
    }

    let mut hf_image_names: HashMap<(usize, u8, usize), String> = HashMap::new();
    let mut hf_inline_image_names: HashMap<(usize, u8, usize, usize), String> = HashMap::new();
    let mut hf_floating_image_names: HashMap<(usize, u8, usize, usize), String> = HashMap::new();
    {
        let hf_variants: [(u8, fn(&SectionProperties) -> Option<&HeaderFooter>); 6] = [
            (0, |sp| sp.header_default.as_ref()),
            (1, |sp| sp.header_first.as_ref()),
            (2, |sp| sp.footer_default.as_ref()),
            (3, |sp| sp.footer_first.as_ref()),
            (4, |sp| sp.header_even.as_ref()),
            (5, |sp| sp.footer_even.as_ref()),
        ];
        for (si, section) in doc.sections.iter().enumerate() {
            for &(hf_type, accessor) in &hf_variants {
                if let Some(hf) = accessor(&section.properties) {
                    let mut pi = 0usize;
                    for block in &hf.blocks {
                        if let Block::Paragraph(para) = block {
                            if let Some(img) = &para.image {
                                let name = embed_single_image(img, &mut image_xobjects, pdf, alloc);
                                hf_image_names.insert((si, hf_type, pi), name);
                            }
                            for (ri, run) in para.runs.iter().enumerate() {
                                if let Some(img) = &run.inline_image {
                                    let name =
                                        embed_single_image(img, &mut image_xobjects, pdf, alloc);
                                    hf_inline_image_names.insert((si, hf_type, pi, ri), name);
                                }
                            }
                            for (fi, floating) in para.floating_images.iter().enumerate() {
                                let name = embed_single_image(
                                    &floating.image,
                                    &mut image_xobjects,
                                    pdf,
                                    alloc,
                                );
                                hf_floating_image_names.insert((si, hf_type, pi, fi), name);
                            }
                            pi += 1;
                        }
                    }
                }
            }
        }
    }

    let mut table_cell_image_names: HashMap<usize, String> = HashMap::new();
    {
        let mut tables: Vec<&Table> = Vec::new();
        for section in &doc.sections {
            for block in &section.blocks {
                if let Block::Table(table) = block {
                    tables.push(table);
                }
            }
            let hf_list: [Option<&HeaderFooter>; 6] = [
                section.properties.header_default.as_ref(),
                section.properties.header_first.as_ref(),
                section.properties.footer_default.as_ref(),
                section.properties.footer_first.as_ref(),
                section.properties.header_even.as_ref(),
                section.properties.footer_even.as_ref(),
            ];
            for hf_opt in hf_list {
                if let Some(hf) = hf_opt {
                    for block in &hf.blocks {
                        if let Block::Table(table) = block {
                            tables.push(table);
                        }
                    }
                }
            }
        }
        for table in tables {
            for row in &table.rows {
                for cell in &row.cells {
                    for para in &cell.paragraphs {
                        if let Some(img) = &para.image {
                            let key = std::sync::Arc::as_ptr(&img.data) as usize;
                            if !table_cell_image_names.contains_key(&key) {
                                let name =
                                    embed_single_image(img, &mut image_xobjects, pdf, alloc);
                                table_cell_image_names.insert(key, name);
                            }
                        }
                    }
                }
            }
        }
    }

    EmbeddedImages {
        image_pdf_names,
        inline_image_pdf_names,
        floating_image_pdf_names,
        image_xobjects,
        hf_image_names,
        hf_inline_image_names,
        hf_floating_image_names,
        table_cell_image_names,
    }
}

fn srgb_to_linear(s: f32) -> f32 {
    if s <= 0.04045 {
        s / 12.92
    } else {
        ((s + 0.055) / 1.055).powf(2.4)
    }
}

fn srgb_byte_to_linear(c: u8) -> f32 {
    srgb_to_linear(c as f32 / 255.0)
}

fn linear_to_oklab(r: f32, g: f32, b: f32) -> [f32; 3] {
    let l = 0.4122214708 * r + 0.5363325363 * g + 0.0514459929 * b;
    let m = 0.2119034982 * r + 0.6806995451 * g + 0.1073969566 * b;
    let s = 0.0883024619 * r + 0.2220049874 * g + 0.6896925507 * b;
    let l_ = l.cbrt();
    let m_ = m.cbrt();
    let s_ = s.cbrt();
    [
        0.2104542553 * l_ + 0.7936177850 * m_ - 0.0040720468 * s_,
        1.9779984951 * l_ - 2.4285922050 * m_ + 0.4505937099 * s_,
        0.0259040371 * l_ + 0.7827717662 * m_ - 0.8086757660 * s_,
    ]
}

fn oklab_to_linear(lab: [f32; 3]) -> [f32; 3] {
    let l_ = lab[0] + 0.3963377774 * lab[1] + 0.2158037573 * lab[2];
    let m_ = lab[0] - 0.1055613458 * lab[1] - 0.0638541728 * lab[2];
    let s_ = lab[0] - 0.0894841775 * lab[1] - 1.2914855480 * lab[2];
    let l = l_ * l_ * l_;
    let m = m_ * m_ * m_;
    let s = s_ * s_ * s_;
    [
        (4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s).clamp(0.0, 1.0),
        (-1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s).clamp(0.0, 1.0),
        (-0.0041960863 * l - 0.7034186147 * m + 1.7076147010 * s).clamp(0.0, 1.0),
    ]
}

fn srgb_to_oklab(c: [u8; 3]) -> [f32; 3] {
    linear_to_oklab(
        srgb_byte_to_linear(c[0]),
        srgb_byte_to_linear(c[1]),
        srgb_byte_to_linear(c[2]),
    )
}

const GRAD_SUBDIVISIONS: usize = 16;

fn build_oklab_gradient_func(
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
    c0: [u8; 3],
    c1: [u8; 3],
) -> Ref {
    let lab0 = srgb_to_oklab(c0);
    let lab1 = srgb_to_oklab(c1);
    let n = GRAD_SUBDIVISIONS;
    let mut stops: Vec<[f32; 3]> = Vec::with_capacity(n + 1);
    for i in 0..=n {
        let t = i as f32 / n as f32;
        let lab = [
            lab0[0] + t * (lab1[0] - lab0[0]),
            lab0[1] + t * (lab1[1] - lab0[1]),
            lab0[2] + t * (lab1[2] - lab0[2]),
        ];
        stops.push(oklab_to_linear(lab));
    }
    let sub_refs: Vec<Ref> = stops
        .windows(2)
        .map(|pair| {
            let fref = alloc();
            pdf.exponential_function(fref)
                .domain([0.0, 1.0])
                .c0(pair[0])
                .c1(pair[1])
                .n(1.0);
            fref
        })
        .collect();
    let bounds: Vec<f32> = (1..n).map(|i| i as f32 / n as f32).collect();
    let encode: Vec<f32> = sub_refs.iter().flat_map(|_| [0.0, 1.0]).collect();
    let stitch_ref = alloc();
    pdf.stitching_function(stitch_ref)
        .domain([0.0, 1.0])
        .functions(sub_refs)
        .bounds(bounds)
        .encode(encode);
    stitch_ref
}

fn srgb_to_linear_rgb(c: [u8; 3]) -> [f32; 3] {
    [
        srgb_byte_to_linear(c[0]),
        srgb_byte_to_linear(c[1]),
        srgb_byte_to_linear(c[2]),
    ]
}

#[allow(clippy::too_many_arguments)]
fn assemble_pdf_pages(
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
    catalog_id: Ref,
    pages_id: Ref,
    all_contents: Vec<Content>,
    all_hf_contents: &mut Vec<Option<Content>>,
    all_page_links: &[Vec<LinkAnnotation>],
    all_page_alpha_states: &[HashSet<u8>],
    all_page_gradient_specs: &[Vec<GradientSpec>],
    page_section_indices: &[(usize, bool)],
    seen_fonts: &HashMap<String, FontEntry>,
    font_order: &[String],
    image_xobjects: &[(String, Ref)],
    doc: &Document,
) {
    let n = all_contents.len();
    let page_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();
    let content_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();

    let page_annot_refs: Vec<Vec<Ref>> = all_page_links
        .iter()
        .map(|links| {
            links
                .iter()
                .map(|link| {
                    let annot_ref = alloc();
                    let mut annot = pdf.annotation(annot_ref);
                    annot
                        .subtype(pdf_writer::types::AnnotationType::Link)
                        .rect(link.rect)
                        .border(0.0, 0.0, 0.0, None);
                    annot
                        .action()
                        .action_type(pdf_writer::types::ActionType::Uri)
                        .uri(Str(link.url.as_bytes()));
                    annot_ref
                })
                .collect()
        })
        .collect();

    let all_alpha_values: HashSet<u8> = all_page_alpha_states
        .iter()
        .flat_map(|s| s.iter().copied())
        .collect();
    let alpha_gs_refs: HashMap<u8, Ref> = all_alpha_values
        .iter()
        .map(|&pct| {
            let gs_ref = alloc();
            pdf.ext_graphics(gs_ref)
                .non_stroking_alpha(pct as f32 / 100.0);
            (pct, gs_ref)
        })
        .collect();

    let all_page_pattern_refs: Vec<Vec<(String, Ref)>> = all_page_gradient_specs
        .iter()
        .map(|specs| {
            specs
                .iter()
                .map(|spec| {
                    let func_ref = if spec.stops.len() <= 2 {
                        let (c0, c1) = if spec.stops.len() >= 2 {
                            (spec.stops[0].0, spec.stops[spec.stops.len() - 1].0)
                        } else {
                            (spec.stops[0].0, spec.stops[0].0)
                        };
                        build_oklab_gradient_func(pdf, alloc, c0, c1)
                    } else {
                        let segment_refs: Vec<Ref> = spec
                            .stops
                            .windows(2)
                            .map(|pair| {
                                build_oklab_gradient_func(pdf, alloc, pair[0].0, pair[1].0)
                            })
                            .collect();

                        let bounds: Vec<f32> = spec.stops[1..spec.stops.len() - 1]
                            .iter()
                            .map(|s| s.1)
                            .collect();
                        let encode: Vec<f32> =
                            segment_refs.iter().flat_map(|_| [0.0, 1.0]).collect();

                        let stitch_ref = alloc();
                        pdf.stitching_function(stitch_ref)
                            .domain([0.0, 1.0])
                            .functions(segment_refs)
                            .bounds(bounds)
                            .encode(encode);
                        stitch_ref
                    };

                    let ang_rad = spec.angle_deg.to_radians();
                    let (sin_a, cos_a) = ang_rad.sin_cos();
                    let cx = spec.x + spec.w / 2.0;
                    let cy = spec.y + spec.h / 2.0;
                    let half_len = ((spec.w / 2.0 * cos_a).powi(2)
                        + (spec.h / 2.0 * sin_a).powi(2))
                    .sqrt();
                    let x0 = cx - half_len * cos_a;
                    let y0 = cy + half_len * sin_a;
                    let x1 = cx + half_len * cos_a;
                    let y1 = cy - half_len * sin_a;

                    let pat_ref = alloc();
                    let mut pattern = pdf.shading_pattern(pat_ref);
                    let mut shading = pattern.function_shading();
                    shading
                        .shading_type(pdf_writer::types::FunctionShadingType::Axial)
                        .color_space()
                        .cal_rgb(
                            [0.9505, 1.0, 1.0890],
                            None,
                            None,
                            Some([
                                0.4124, 0.2126, 0.0193, 0.3576, 0.7152, 0.1192, 0.1805, 0.0722,
                                0.9505,
                            ]),
                        );
                    shading
                        .function(func_ref)
                        .coords([x0, y0, x1, y1])
                        .extend([true, true]);

                    (spec.pattern_name.clone(), pat_ref)
                })
                .collect()
        })
        .collect();

    for (i, c) in all_contents.into_iter().enumerate() {
        let body_raw = c.finish();
        if let Some(hf) = all_hf_contents[i].take() {
            let hf_raw = hf.finish();
            let mut combined = Vec::with_capacity(hf_raw.len() + 1 + body_raw.len());
            combined.extend_from_slice(hf_raw.as_slice());
            combined.push(b'\n');
            combined.extend_from_slice(body_raw.as_slice());
            let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&combined, 6);
            pdf.stream(content_ids[i], &compressed)
                .filter(Filter::FlateDecode);
        } else {
            let compressed = miniz_oxide::deflate::compress_to_vec_zlib(body_raw.as_slice(), 6);
            pdf.stream(content_ids[i], &compressed)
                .filter(Filter::FlateDecode);
        }
    }

    pdf.catalog(catalog_id).pages(pages_id);
    pdf.pages(pages_id)
        .kids(page_ids.iter().copied())
        .count(n as i32);

    let font_pairs: Vec<(String, Ref)> = font_order
        .iter()
        .map(|name| (seen_fonts[name].pdf_name.clone(), seen_fonts[name].font_ref))
        .collect();

    for i in 0..n {
        let (si, _) = page_section_indices[i];
        let sp = &doc.sections[si].properties;
        let mut page = pdf.page(page_ids[i]);
        page.media_box(Rect::new(0.0, 0.0, sp.page_width, sp.page_height))
            .parent(pages_id)
            .contents(content_ids[i]);
        if !page_annot_refs[i].is_empty() {
            page.annotations(page_annot_refs[i].iter().copied());
        }
        {
            let mut resources = page.resources();
            {
                let mut fonts = resources.fonts();
                for (name, font_ref) in &font_pairs {
                    fonts.pair(Name(name.as_bytes()), *font_ref);
                }
            }
            if !image_xobjects.is_empty() {
                let mut xobjects = resources.x_objects();
                for (name, xobj_ref) in image_xobjects {
                    xobjects.pair(Name(name.as_bytes()), *xobj_ref);
                }
            }
            if let Some(alpha_set) = all_page_alpha_states.get(i).filter(|s| !s.is_empty()) {
                let mut gs_dict = resources.ext_g_states();
                for &pct in alpha_set {
                    let gs_name = format!("GSa{pct}");
                    let gs_ref = alpha_gs_refs[&pct];
                    gs_dict.pair(Name(gs_name.as_bytes()), gs_ref);
                }
            }
            if let Some(pat_refs) = all_page_pattern_refs.get(i).filter(|p| !p.is_empty()) {
                let mut patterns = resources.patterns();
                for (name, pat_ref) in pat_refs {
                    patterns.pair(Name(name.as_bytes()), *pat_ref);
                }
            }
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
        table_cell_image_names: &table_cell_image_names,
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
                            .flat_map(|cell| cell.paragraphs.iter())
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

    // Phase 2: build multi-page content streams (section-aware)
    let first_sp = &doc.sections[0].properties;
    let mut cur_sp = first_sp;
    let initial_slot_top = effective_slot_top(cur_sp, true, &ctx);
    let mut pb = PageBuilder::new(initial_slot_top);
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
                }
                SectionBreakType::Continuous => {
                    // No forced break; geometry updates on next page
                }
            }
            pb.is_first_page_of_section = true;
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

                    let (col_x, col_w) = col_geometry[current_col];
                    let para_text_x = col_x + para.indent_left;
                    let para_text_width = (col_w - para.indent_left - para.indent_right).max(1.0);
                    let label_x = col_x + para.indent_left - para.indent_hanging;
                    let text_hanging = if !para.list_label.is_empty() {
                        0.0
                    } else if para.indent_hanging > 0.0 {
                        para.indent_hanging
                    } else {
                        -para.indent_first_line
                    };

                    // Substitute footnote reference runs with display numbers
                    let has_footnote_refs = para.runs.iter().any(|r| r.footnote_id.is_some());
                    let effective_runs: std::borrow::Cow<'_, Vec<Run>> = if has_footnote_refs {
                        let substituted: Vec<Run> = para
                            .runs
                            .iter()
                            .map(|run| {
                                if let Some(id) = run.footnote_id {
                                    let num = footnote_display_order.get(&id).copied().unwrap_or(0);
                                    let mut r = run.clone();
                                    r.text = num.to_string();
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
                        )
                    } else {
                        build_paragraph_lines(
                            &effective_runs,
                            ctx.fonts,
                            para_text_width,
                            text_hanging,
                            &block_inline_images,
                        )
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
                        para.content_height.max(sp.line_pitch)
                    } else if text_empty {
                        if para.content_height > 0.0 {
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
                            first_line_h
                        } else {
                            first_line_h + (num_lines - 1) as f32 * line_h
                        }
                    };

                    for fi in &para.floating_images {
                        let reserve = match fi.wrap_type {
                            WrapType::TopAndBottom => true,
                            WrapType::Square | WrapType::Tight | WrapType::Through => {
                                fi.image.display_width >= text_width * 0.9
                            }
                            WrapType::None => false,
                        };
                        if reserve {
                            let fi_h = match fi.v_position {
                                VerticalPosition::Offset(o) => o + fi.image.display_height,
                                _ => fi.image.display_height,
                            };
                            content_h = content_h.max(fi_h);
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

                    let needed = inter_gap + bdr_top_pad + content_h;
                    let at_page_top = pb.is_at_page_top(cur_sp);

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
                        && pb.slot_top - needed - keep_next_extra < effective_margin_bottom
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

                        // Reduce to ensure at least 2 lines remain on next page (orphan control)
                        if lines_that_fit > 0 && lines.len().saturating_sub(lines_that_fit) < 2 {
                            lines_that_fit = lines.len().saturating_sub(2);
                        }

                        // keepLines: don't split — move entire paragraph to next column/page
                        if para.keep_lines {
                            lines_that_fit = 0;
                        }

                        if lines_that_fit >= 2 && lines_that_fit < lines.len() {
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

                    // Re-fetch column geometry (may have changed after overflow)
                    let (col_x, col_w) = col_geometry[current_col];
                    let para_text_x = col_x + para.indent_left;
                    let para_text_width = (col_w - para.indent_left - para.indent_right).max(1.0);
                    let label_x = col_x + para.indent_left - para.indent_hanging;

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
                    if let Some([r, g, b]) = para.shading {
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
                        let shd_bottom = pb.slot_top - bdr_top_pad - content_h - bdr_bottom_pad;
                        pb.content.save_state();
                        pb.content.set_fill_rgb(
                            r as f32 / 255.0,
                            g as f32 / 255.0,
                            b as f32 / 255.0,
                        );
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
                        let default_font = ctx
                            .fonts
                            .keys()
                            .next()
                            .map(|s| s.as_str())
                            .unwrap_or("Helvetica");
                        charts::render_chart(
                            ic,
                            &mut pb.content,
                            chart_x,
                            pb.slot_top,
                            ctx.fonts,
                            default_font,
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
                    } else if (para.image.is_some() || text_empty) && para.content_height > 0.0 {
                        if let Some(pdf_name) = image_pdf_names.get(&global_block_idx) {
                            let img = para.image.as_ref().unwrap();
                            let y_bottom = pb.slot_top - img.display_height;
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
                        );
                    }

                    // Draw paragraph borders — left/right borders extend outward
                    // from the text area so text inside stays aligned with text outside
                    {
                        let bdr = &para.borders;
                        let box_top = pb.slot_top;
                        let box_bottom = pb.slot_top - bdr_top_pad - content_h - bdr_bottom_pad;
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
                            let [r, g, b_c] = b.color;
                            content.save_state();
                            content.set_line_width(b.width_pt);
                            content.set_stroke_rgb(
                                r as f32 / 255.0,
                                g as f32 / 255.0,
                                b_c as f32 / 255.0,
                            );
                            content.move_to(box_left, y);
                            content.line_to(box_right, y);
                            content.stroke();
                            content.restore_state();
                        };
                        let draw_v_border = |content: &mut Content, b: &ParagraphBorder, x: f32| {
                            let [r, g, b_c] = b.color;
                            content.save_state();
                            content.set_line_width(b.width_pt);
                            content.set_stroke_rgb(
                                r as f32 / 255.0,
                                g as f32 / 255.0,
                                b_c as f32 / 255.0,
                            );
                            content.move_to(x, box_top);
                            content.line_to(x, box_bottom);
                            content.stroke();
                            content.restore_state();
                        };

                        let prev_has_between = prev_para.is_some_and(|pp| {
                            pp.borders.between.is_some()
                                && borders_match(&pp.borders, &para.borders)
                        });
                        let next_has_between = next_para.is_some_and(|np| {
                            bdr.between.is_some() && borders_match(&para.borders, &np.borders)
                        });

                        if !prev_has_between {
                            if let Some(b) = &bdr.top {
                                draw_h_border(&mut pb.content, b, box_top);
                            }
                        }
                        if next_has_between {
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

                    pb.slot_top -= content_h + bdr_top_pad;
                    prev_space_after = effective_space_after;

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
                        let restore = pos.v_anchor != "text";
                        let y = match pos.v_anchor {
                            "page" => sp.page_height - pos.v_offset_pt,
                            "margin" => sp.page_height - sp.margin_top - pos.v_offset_pt,
                            _ => pb.slot_top - pos.v_offset_pt,
                        };
                        (x, y, restore)
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
                            for p in &cell.paragraphs {
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
            global_block_idx += 1;
        }
    }
    pb.flush_page(doc.sections.len() - 1);

    let t_layout = t0.elapsed();

    // Phase 2b: column separator lines
    for (page_idx, content) in pb.all_contents.iter_mut().enumerate() {
        let (si, _) = pb.page_section_indices[page_idx];
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
        let (si, _) = pb.page_section_indices[page_idx];
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
        let (si, is_first) = pb.page_section_indices[page_idx];
        let sp = &doc.sections[si].properties;

        let page_num = if let Some(start) = sp.page_num_start {
            // Section specifies explicit start: count pages within this section
            let pages_before_in_section = pb.page_section_indices[..page_idx]
                .iter()
                .filter(|&&(s, _)| s == si)
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

        let (header, hdr_type) = if is_first && sp.different_first_page {
            (sp.header_first.as_ref(), 1u8)
        } else if doc.even_and_odd_headers && page_num % 2 == 0 && sp.header_even.is_some() {
            (sp.header_even.as_ref(), 4u8)
        } else {
            (sp.header_default.as_ref(), 0u8)
        };
        if let Some(header_data) = header {
            let (pi_map, ii_map, fi_map) = build_hf_maps(si, hdr_type);
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

        let (footer, ftr_type) = if is_first && sp.different_first_page {
            (sp.footer_first.as_ref(), 3u8)
        } else if doc.even_and_odd_headers && page_num % 2 == 0 && sp.footer_even.is_some() {
            (sp.footer_even.as_ref(), 5u8)
        } else {
            (sp.footer_default.as_ref(), 2u8)
        };
        if let Some(footer_data) = footer {
            let (pi_map, ii_map, fi_map) = build_hf_maps(si, ftr_type);
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

fn label_font_key(para: &Paragraph) -> Option<String> {
    if let Some(ref bf) = para.list_label_font {
        let mut k = bf.clone();
        if para.list_label_bold {
            k.push_str("/B");
        }
        Some(k)
    } else {
        let run = para.runs.first()?;
        let key_run = Run {
            bold: para.list_label_bold || run.bold,
            ..run.clone()
        };
        Some(font_key(&key_run))
    }
}

fn label_for_paragraph<'a>(
    para: &Paragraph,
    seen_fonts: &'a HashMap<String, FontEntry>,
) -> (&'a str, Vec<u8>) {
    let Some(key) = label_font_key(para) else {
        return ("", vec![]);
    };
    let Some(entry) = seen_fonts.get(&key) else {
        return ("", vec![]);
    };
    let bytes = match &entry.char_to_gid {
        Some(map) => encode_as_gids(&para.list_label, map),
        None => to_winansi_bytes(&para.list_label),
    };
    (entry.pdf_name.as_str(), bytes)
}

fn render_list_label(
    content: &mut Content,
    para: &Paragraph,
    fonts: &HashMap<String, FontEntry>,
    label_x: f32,
    baseline_y: f32,
    fallback_font_size: f32,
) {
    if para.list_label.is_empty() {
        return;
    }
    let (label_font_name, label_bytes) = label_for_paragraph(para, fonts);
    let label_color = para
        .list_label_color
        .or_else(|| para.runs.first().and_then(|r| r.color));
    if let Some([r, g, b]) = label_color {
        content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
    }
    let label_fs = para.list_label_font_size.unwrap_or(fallback_font_size);
    content
        .begin_text()
        .set_font(Name(label_font_name.as_bytes()), label_fs)
        .next_line(label_x, baseline_y)
        .show(Str(&label_bytes))
        .end_text();
    if label_color.is_some() {
        content.set_fill_gray(0.0);
    }
}
