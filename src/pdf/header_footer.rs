use std::collections::HashMap;

use pdf_writer::Content;

use crate::model::{
    Alignment, Block, Document, FieldCode, FrameProperties, HRelativeFrom, HeaderFooter, IfPart,
    LineSpacing, Paragraph, Run, SectionProperties, TextAnchor, VRelativeFrom, VerticalPosition,
    WrapType,
};

use super::color::stroke_segment;
use super::helpers::{align_offset, draw_horizontal_rule};
use super::layout::{
    LineOpts, LinkAnnotation, build_lines, is_text_empty, lines_height, picture_line_bottom,
    position_stretch, render_paragraph_lines, runs_max_image_h, size_lines_by_own_runs,
    tallest_run_metrics,
};
use super::positioning::resolve_h_position;
use super::table;
use super::{RenderContext, resolve_line_h};

pub(super) fn substitute_hf_runs(
    runs: &[Run],
    page_num: usize,
    total_pages: usize,
    styleref_values: &HashMap<String, String>,
    page_num_format: Option<&str>,
) -> Vec<Run> {
    let values = FieldValues {
        page_num,
        total_pages,
        styleref_values,
        page_num_format,
    };
    runs.iter()
        .map(|run| {
            let mut r = run.clone();
            if let Some(ref fc) = run.field_code {
                r.field_code = None;
                r.text = values.eval(fc).unwrap_or_else(|| run.text.clone());
            }
            r
        })
        .collect()
}

/// The key a STYLEREF value is stored under: Word matches style names
/// case-insensitively ("CharSchno" for CharSchNo), and the `\n` switch reads
/// the paragraph's list number.
pub(super) fn styleref_key(name: &str, number: bool) -> String {
    let key = name.to_lowercase();
    if number { key + "\\n" } else { key }
}

/// What a header or footer field shows on one page.
struct FieldValues<'a> {
    page_num: usize,
    total_pages: usize,
    styleref_values: &'a HashMap<String, String>,
    page_num_format: Option<&'a str>,
}

impl FieldValues<'_> {
    /// None keeps the cached result.
    fn eval(&self, fc: &FieldCode) -> Option<String> {
        Some(match fc {
            FieldCode::Page => match self.page_num_format {
                Some(fmt) => crate::docx::numbering::format_number(self.page_num as u32, fmt),
                None => self.page_num.to_string(),
            },
            FieldCode::NumPages => self.total_pages.to_string(),
            FieldCode::StyleRef { name, number } => self
                .styleref_values
                .get(&styleref_key(name, *number))
                .cloned()
                .unwrap_or_default(),
            FieldCode::PageRef(_) => return None,
            FieldCode::If(parts) => return self.eval_if(parts),
        })
    }

    /// `IF expr1 op expr2 "true" "false"` after substituting the nested
    /// fields.
    fn eval_if(&self, parts: &[IfPart]) -> Option<String> {
        let mut instr = String::new();
        for part in parts {
            match part {
                IfPart::Text(t) => instr.push_str(t),
                IfPart::Field(fc) => instr.push_str(&self.eval(fc)?),
            }
        }
        let args = field_args(instr.trim_start().get(2..)?);
        let [a, op, b, t, f] = args.as_slice() else {
            return None;
        };
        let ord = match (a.trim().parse::<f64>(), b.trim().parse::<f64>()) {
            (Ok(x), Ok(y)) => x.partial_cmp(&y)?,
            _ => a.cmp(b),
        };
        let holds = match op.as_str() {
            "=" => ord.is_eq(),
            "<>" => ord.is_ne(),
            "<" => ord.is_lt(),
            "<=" => ord.is_le(),
            ">" => ord.is_gt(),
            ">=" => ord.is_ge(),
            _ => return None,
        };
        Some(if holds { t } else { f }.clone())
    }
}

/// Field arguments: whitespace-separated, double quotes group (and are
/// dropped), comparison operators stand alone even without spaces.
fn field_args(s: &str) -> Vec<String> {
    let mut args = Vec::new();
    let mut chars = s.chars().peekable();
    while let Some(&c) = chars.peek() {
        if c.is_whitespace() {
            chars.next();
        } else if c == '"' {
            chars.next();
            args.push(chars.by_ref().take_while(|&c| c != '"').collect());
        } else if "=<>".contains(c) {
            let mut op = String::new();
            while let Some(&c) = chars.peek().filter(|c| "=<>".contains(**c)) {
                op.push(c);
                chars.next();
            }
            args.push(op);
        } else {
            let mut word = String::new();
            while let Some(&c) = chars
                .peek()
                .filter(|c| !c.is_whitespace() && !"\"=<>".contains(**c))
            {
                word.push(c);
                chars.next();
            }
            args.push(word);
        }
    }
    args
}

/// Vertical bands (top, bottom from the page top) of the header's page- or
/// margin-anchored frames that forbid text beside them (`w:wrap`
/// none/notBeside with an explicit `w:h`). An in-flow header line that would
/// overlap such a band is laid out below it: bosch's first-page header has a
/// 106pt "Persbericht" frame at 33pt, its first two 14.75pt lines fit above
/// it and the third lands at 139pt, which is why Word starts the body at
/// 153pt while the top margin says 86pt (annotation #195). ponytail: `w:h`
/// is taken as the frame height whatever `w:hRule` says; grow it by the
/// frame's content height if a frame ever overflows it.
fn blocking_frame_bands(hf: &HeaderFooter, sp: &SectionProperties) -> Vec<(f32, f32)> {
    hf.blocks
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => p.frame_props.as_ref(),
            _ => None,
        })
        .filter(|fp| fp.text_below && fp.height > 0.0)
        .filter_map(|fp| anchored_frame_top(fp, fp.height, sp).map(|top| (top, top + fp.height)))
        .collect()
}

/// A page- or margin-anchored frame's top, down from the page top; None for a
/// paragraph-anchored frame, which stays in the flow. `height` places a
/// bottom- or center-aligned frame.
pub(super) fn anchored_frame_top(
    fp: &FrameProperties,
    height: f32,
    sp: &SectionProperties,
) -> Option<f32> {
    if fp.v_relative_from == VRelativeFrom::Paragraph {
        return None;
    }
    let top = resolve_tb_y_top(fp.v_relative_from, &fp.v_position, height, sp, 0.0);
    Some(sp.page_height - top)
}

/// The band a bottom border adds below a paragraph: its `space` and its stroke
/// (the body path's `bdr_bottom_extent`). The next paragraph starts below it.
fn bottom_border_band(para: &Paragraph) -> f32 {
    super::helpers::border_band(para.borders.bottom.as_ref())
}

/// The band a top border adds above a paragraph's lines, as in the body
/// (`bdr_top_pad`): carbon_farming's footer rule sits 1.75pt above its text.
fn top_border_band(para: &Paragraph) -> f32 {
    super::helpers::border_band(para.borders.top.as_ref())
}

/// How far a header (footer) paragraph's text drops below its own floats:
/// one it may not sit beside, starting at or above the paragraph top, pushes
/// the text under it. Word probes on renewable_dispatch's 568pt header
/// shape: topAndBottom pushes at any width, a square wrap only when no room
/// is left beside it. A float starting lower leaves the first line above it,
/// and one behind the text never pushes (corpus header logos).
/// ponytail: "no room" = as wide as the text; measure the side gaps if a
/// partly covering float turns up.
fn float_text_push(para: &Paragraph, text_width: f32) -> f32 {
    para.floating_images
        .iter()
        .filter(|fi| matches!(fi.v_relative_from, VRelativeFrom::Paragraph) && !fi.behind_doc)
        .filter(|fi| match fi.wrap_type {
            WrapType::TopAndBottom => true,
            w => w.wraps_beside() && fi.image.display_width >= text_width,
        })
        .filter_map(|fi| match fi.v_position {
            VerticalPosition::Offset(o) if o <= 0.0 => {
                Some(o + fi.image.display_height + fi.dist_bottom)
            }
            _ => None,
        })
        .fold(0.0, f32::max)
}

/// Where a line of height `line_h` whose top sits `top` below the page top
/// really starts: below every blocking frame band it would overlap, including
/// one it meets only below another (massachusetts' logo band ends where its
/// governor table's begins).
pub(super) fn below_blocking_frames(mut top: f32, line_h: f32, bands: &[(f32, f32)]) -> f32 {
    while let Some(&(_, b_bot)) = bands
        .iter()
        .find(|&&(b_top, b_bot)| top < b_bot && top + line_h > b_top)
    {
        top = b_bot;
    }
    top
}

/// The mark's descent separates automatically wrapped picture lines, but is
/// not retained again below the final line. A differently sized paragraph mark
/// requires its own line metrics; keep that existing path until it is resolved.
fn wrapped_picture_trailing_descent(
    para: &Paragraph,
    font_size: f32,
    descent: f32,
    spacing: LineSpacing,
) -> f32 {
    if matches!(spacing, LineSpacing::Exact(_))
        || para.runs.iter().any(|r| r.is_line_break)
        || para
            .paragraph_mark_font_size
            .is_some_and(|mark| (mark - font_size).abs() > 0.01)
    {
        0.0
    } else {
        descent
    }
}

fn wrapped_header_picture_height(para: &Paragraph, ctx: &RenderContext, width: f32) -> Option<f32> {
    if para
        .runs
        .iter()
        .filter(|r| r.inline_image.is_some())
        .count()
        < 2
        || para
            .runs
            .iter()
            .any(|r| !r.text.trim().is_empty() || r.field_code.is_some())
    {
        return None;
    }
    let images: HashMap<usize, String> = para
        .runs
        .iter()
        .enumerate()
        .filter(|(_, r)| r.inline_image.is_some())
        .map(|(i, _)| (i, format!("Im{i}")))
        .collect();
    let (font_size, lhr, ar) = tallest_run_metrics(&para.runs, ctx.fonts);
    let spacing = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
    let line_h = resolve_line_h(spacing, font_size, lhr);
    let mut lines = build_lines(
        &para.runs,
        ctx,
        (width - para.indent_left - para.indent_right).max(1.0),
        ctx.cjk(true, para.alignment),
        &LineOpts {
            inline_images: Some(&images),
            tab_stops: &para.tab_stops,
            indent_left: para.indent_left,
            indent_right: para.indent_right,
            hanging: para.indent_hanging - para.indent_first_line,
            ..Default::default()
        },
    );
    if lines.len() < 2 {
        return None;
    }
    for (chunk, image) in lines
        .iter_mut()
        .flat_map(|l| &mut l.chunks)
        .filter(|c| c.inline_image_name.is_some())
        .zip(para.runs.iter().filter_map(|r| r.inline_image.as_ref()))
    {
        chunk.inline_image_height += image.layout_extra_height;
        chunk.y_offset = image.layout_extra_height - image.layout_extra_top;
    }
    let metrics = (
        font_size * ar.unwrap_or(0.75),
        font_size * super::layout::descender_ratio(lhr, ar),
    );
    if !matches!(spacing, LineSpacing::Exact(_)) {
        size_lines_by_own_runs(&mut lines, ctx.fonts, spacing, line_h, metrics.0);
    }
    let trailing_descent = wrapped_picture_trailing_descent(para, font_size, metrics.1, spacing);
    Some(lines_height(&lines, line_h, metrics) - trailing_descent)
}

fn compute_header_height(
    hf: &HeaderFooter,
    ctx: &RenderContext,
    sp: &SectionProperties,
    is_header: bool,
) -> f32 {
    let text_width = sp.text_width();
    let mut height = 0.0f32;
    let mut prev_space_after = 0.0f32;
    let mut prev_para: Option<&Paragraph> = None;
    // How far page- or margin-placed floats reach from the header (footer)
    // margin, whatever the flow above them.
    let mut placed_extent = 0.0f32;
    let bands = if is_header {
        blocking_frame_bands(hf, sp)
    } else {
        Vec::new()
    };
    for block in &hf.blocks {
        match block {
            Block::Paragraph(para) if para.frame_props.is_some() => {
                // Frame paragraphs are out-of-flow; skip height contribution
            }
            Block::Paragraph(para) => {
                height += hf_paragraph_gap(prev_para, prev_space_after, para);
                prev_para = Some(para);
                let (font_size, tallest_lhr, _) = tallest_run_metrics(&para.runs, ctx.fonts);
                let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
                height = below_blocking_frames(sp.header_margin + height, line_h, &bands)
                    - sp.header_margin;
                // Positioned runs and marks stretch their lines as in the render.
                let line_h = line_h
                    + if is_text_empty(&para.runs) {
                        super::mark_position_stretch(para, effective_ls, false)
                    } else {
                        position_stretch(&para.runs, effective_ls)
                    };
                // Mirrors the render loop's advance: a picture line is the
                // picture plus the text descent (`inline_line_advance`).
                let picture_h = runs_max_image_h(&para.runs);
                let wrapped_h = wrapped_header_picture_height(para, ctx, text_width);
                let mut content_h = if let Some(h) = wrapped_h {
                    h
                } else if picture_h > 0.0 {
                    line_h.max(
                        picture_h + picture_line_bottom(&para.runs, para, ctx.fonts, effective_ls),
                    )
                } else {
                    line_h
                } + float_text_push(para, text_width);

                for fi in &para.floating_images {
                    // Text flows beside a narrow wrapping image, and Word does
                    // not extend the header to its bottom: czech_municipal's
                    // logo hangs 5.5pt below the header text, where the body
                    // starts.
                    if matches!(fi.wrap_type, WrapType::None)
                        || (fi.wrap_type.wraps_beside()
                            && fi.image.display_width < text_width * 0.5)
                    {
                        continue;
                    }
                    let fi_h = match fi.v_position {
                        // A paragraph-relative float raised above its paragraph
                        // ends that much higher (americas_counter_terrorism: a
                        // -9pt logo ends 0.35pt above the body top, which Word
                        // leaves in place).
                        VerticalPosition::Offset(o)
                            if matches!(fi.v_relative_from, VRelativeFrom::Paragraph) =>
                        {
                            (o + fi.image.display_height).max(0.0)
                        }
                        // A page- or margin-placed float covers its own band of
                        // the page: cyprus_ucits' footer logo 675.8pt below the
                        // top margin made an 85-page document of 3.
                        VerticalPosition::Offset(o) => {
                            let top = match fi.v_relative_from {
                                VRelativeFrom::Page => o,
                                _ => sp.margin_top + o,
                            };
                            placed_extent = placed_extent.max(if is_header {
                                top + fi.image.display_height - sp.header_margin
                            } else {
                                sp.page_height - sp.footer_margin - top
                            });
                            continue;
                        }
                        _ => fi.image.display_height,
                    };
                    // TopAndBottom or wide wrapping image
                    content_h = content_h.max(fi_h);
                }

                for tb in &para.textboxes {
                    if matches!(tb.wrap_type, WrapType::TopAndBottom) {
                        let tb_bottom = tb.v_offset_pt + tb.height_pt + tb.dist_bottom;
                        match tb.v_relative_from {
                            VRelativeFrom::Paragraph => {
                                content_h = content_h.max(tb_bottom);
                            }
                            VRelativeFrom::Page if is_header => {
                                // tb_bottom is absolute from page top; convert to
                                // distance relative to the current position in the header
                                let current_pos = sp.header_margin + height;
                                let contribution = (tb_bottom - current_pos).max(0.0);
                                content_h = content_h.max(contribution);
                            }
                            _ => {
                                content_h += tb_bottom;
                            }
                        }
                    }
                }

                // Each w:br (line break) in the paragraph creates an additional line.
                let br_count = para.runs.iter().filter(|r| r.is_line_break).count();
                if wrapped_h.is_none() {
                    content_h += br_count as f32 * line_h;
                }

                height += top_border_band(para) + content_h + bottom_border_band(para);
                prev_space_after = para.space_after;
            }
            // A floating table takes no room in the header's flow
            Block::Table(table) if table.position.is_some() => {}
            Block::Table(table) => {
                let content_w = sp.text_width();
                height += table::compute_hf_table_height(table, ctx, content_w);
                prev_space_after = 0.0;
                prev_para = None;
            }
        }
    }
    (height + prev_space_after).max(placed_extent)
}

/// The body top of a section's page; `page_idx` (0-based) picks the even
/// header where the document has them.
pub(super) fn effective_slot_top(
    sp: &SectionProperties,
    is_first: bool,
    page_idx: usize,
    ctx: &RenderContext,
) -> f32 {
    let header = layout_hf(sp, is_first, page_idx, true, ctx);
    let base = sp.page_height - sp.margin_top;
    match header {
        Some(_) if sp.margin_top_fixed => base,
        Some(hf) => {
            base.min(sp.page_height - sp.header_margin - compute_header_height(hf, ctx, sp, true))
        }
        None => base,
    }
}

pub(super) fn compute_effective_margin_bottom(
    sp: &SectionProperties,
    is_first: bool,
    page_idx: usize,
    ctx: &RenderContext,
) -> f32 {
    let footer = layout_hf(sp, is_first, page_idx, false, ctx);
    let base = sp.margin_bottom;
    match footer {
        Some(_) if sp.margin_bottom_fixed => base,
        Some(hf) => base.max(sp.footer_margin + compute_header_height(hf, ctx, sp, false)),
        None => base,
    }
}

/// The header (or footer) whose extent a section's page lays out around: its
/// own, else the one it inherits, as drawn (radiographer's later sections
/// inherit a two-line empty header that starts their body 27pt below the
/// header), and on an even page (0-based `page_idx` odd) the even header.
/// ponytail: physical parity; the drawn header follows the displayed number.
fn layout_hf<'a>(
    sp: &'a SectionProperties,
    is_first: bool,
    page_idx: usize,
    is_header: bool,
    ctx: &RenderContext<'a>,
) -> Option<&'a HeaderFooter> {
    let variant = hf_variant(ctx.even_and_odd_headers, sp, is_first, page_idx + 1);
    // `sp` is always one of the document's sections
    let idx = ctx
        .sections
        .iter()
        .position(|s| std::ptr::eq(&s.properties, sp))?;
    inherited_hf(ctx.sections, idx, variant, is_header).map(|(hf, _)| hf)
}

pub(super) fn hf_paragraphs(hf: &HeaderFooter) -> Vec<&Paragraph> {
    hf.blocks
        .iter()
        .flat_map(|block| match block {
            Block::Paragraph(p) => vec![p],
            Block::Table(t) => t
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .flat_map(|cell| cell.all_paragraphs())
                .collect(),
        })
        .collect()
}

pub(super) fn resolve_tb_y_top(
    v_relative_from: VRelativeFrom,
    v_position: &crate::model::VerticalPosition,
    tb_height: f32,
    sp: &SectionProperties,
    slot_top: f32,
) -> f32 {
    use crate::model::VerticalPosition;
    let (region_top, region_bottom) = match v_relative_from {
        VRelativeFrom::Page => (sp.page_height, 0.0),
        VRelativeFrom::Margin | VRelativeFrom::TopMargin => {
            (sp.page_height - sp.margin_top, sp.margin_bottom)
        }
        VRelativeFrom::Paragraph => (slot_top, slot_top - tb_height),
    };
    match v_position {
        VerticalPosition::Offset(o) => region_top - o,
        VerticalPosition::AlignTop => region_top,
        VerticalPosition::AlignBottom => region_bottom + tb_height,
        VerticalPosition::AlignCenter => {
            region_top - ((region_top - region_bottom) - tb_height) / 2.0
        }
    }
}

/// A wrapping float in a header or footer: its box (PDF coordinates) and the
/// distances text keeps from its sides.
#[derive(Clone, Copy)]
struct HfFloatZone {
    left: f32,
    top: f32,
    bottom: f32,
    right: f32,
    dist_left: f32,
    dist_right: f32,
}

impl HfFloatZone {
    fn for_float(fi: &crate::model::FloatingImage, fi_x: f32, fi_y_top: f32) -> Self {
        HfFloatZone {
            left: fi_x,
            top: fi_y_top,
            bottom: fi_y_top - fi.image.display_height,
            right: fi_x + fi.image.display_width,
            dist_left: fi.dist_left,
            dist_right: fi.dist_right,
        }
    }
}

/// Per-page context for header/footer rendering: field substitution values
/// and image name mappings resolved by the caller.
pub(super) struct HfPageContext<'a> {
    pub(super) page_num: usize,
    pub(super) total_pages: usize,
    pub(super) para_image_names: &'a HashMap<usize, String>,
    pub(super) inline_image_names: &'a HashMap<(usize, usize), String>,
    pub(super) floating_image_names: &'a HashMap<(usize, usize), String>,
    pub(super) effect_para_names: &'a HashMap<usize, super::images::EffectXObjs>,
    pub(super) styleref_values: &'a HashMap<String, String>,
    pub(super) page_num_format: Option<&'a str>,
}

pub(super) fn render_header_footer(
    content: &mut Content,
    hf: &HeaderFooter,
    ctx: &RenderContext,
    sp: &SectionProperties,
    is_header: bool,
    pc: &HfPageContext,
    gradient_specs: &mut Vec<super::GradientSpec>,
    // The page's header/footer links; the text stays an artifact.
    links: &mut Vec<LinkAnnotation>,
) {
    let page_num = pc.page_num;
    let total_pages = pc.total_pages;
    let para_image_names = pc.para_image_names;
    let inline_image_names = pc.inline_image_names;
    let floating_image_names = pc.floating_image_names;
    let styleref_values = pc.styleref_values;
    let page_num_format = pc.page_num_format;
    let text_width = sp.text_width();
    let mut cursor_y = if is_header {
        sp.page_height - sp.header_margin
    } else {
        sp.footer_margin + compute_header_height(hf, ctx, sp, false)
    };

    let mut pi = 0usize;
    let mut prev_space_after = 0.0f32;
    let mut prev_para: Option<&Paragraph> = None;
    // A header can hold several wrapping floats (e.g. a logo on each side of a
    // centered letterhead) — all of them constrain the text bounds together.
    let mut hdr_fz: Vec<HfFloatZone> = Vec::new();
    let bands = if is_header {
        blocking_frame_bands(hf, sp)
    } else {
        Vec::new()
    };
    for block in &hf.blocks {
        match block {
            Block::Table(table) => {
                table::render_header_footer_table(
                    table,
                    sp,
                    ctx,
                    content,
                    &mut cursor_y,
                    page_num,
                    total_pages,
                    styleref_values,
                    page_num_format,
                    gradient_specs,
                    links,
                );
                prev_space_after = 0.0;
                prev_para = None;
            }
            Block::Paragraph(para) if para.frame_props.is_some() => {
                let fp = para.frame_props.as_ref().unwrap();
                let substituted_runs = substitute_hf_runs(
                    &para.runs,
                    page_num,
                    total_pages,
                    styleref_values,
                    page_num_format,
                );
                let (font_size, tallest_lhr, tallest_ar) =
                    tallest_run_metrics(&substituted_runs, ctx.fonts);
                let ascender_ratio = tallest_ar.unwrap_or(0.75);
                let frame_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let frame_ascent = super::layout::boxed_line_ascent(
                    frame_ls,
                    resolve_line_h(frame_ls, font_size, tallest_lhr),
                    font_size,
                    tallest_lhr,
                    tallest_ar,
                    &substituted_runs,
                    ctx.fonts,
                )
                .unwrap_or(font_size * ascender_ratio);

                let lines = build_lines(
                    &substituted_runs,
                    ctx,
                    text_width,
                    ctx.cjk(true, para.alignment),
                    &LineOpts {
                        tab_stops: &para.tab_stops,
                        ..Default::default()
                    },
                );
                let content_width = lines.iter().map(|l| l.total_width).fold(0.0f32, f32::max);

                // A framePr with an explicit width (w:w) positions a fixed-width
                // frame box within the anchor area; text then flows from the
                // frame's left edge. Without a width we fall back to aligning the
                // text itself (legacy behaviour).
                let frame_x = if fp.width > 0.0 {
                    let (origin, area_width) = match fp.h_relative_from {
                        HRelativeFrom::Page => (0.0, sp.page_width),
                        HRelativeFrom::Margin | HRelativeFrom::Column => {
                            (sp.margin_left, text_width)
                        }
                    };
                    fp.h_position.place(origin, area_width, fp.width) + para.indent_left
                } else {
                    resolve_h_position(
                        fp.h_relative_from,
                        &fp.h_position,
                        content_width,
                        sp,
                        sp.margin_left,
                        text_width,
                        text_width,
                    )
                };
                // vAnchor + w:y pin the frame top to the page/margin, independent
                // of the flowing header/footer cursor. Paragraph-anchored frames
                // keep the in-flow position.
                let frame_h = if fp.height > 0.0 {
                    fp.height
                } else {
                    lines.len() as f32 * resolve_line_h(frame_ls, font_size, tallest_lhr)
                };
                let frame_top = anchored_frame_top(fp, frame_h, sp)
                    .map_or(cursor_y, |top| sp.page_height - top);
                let frame_baseline = frame_top - frame_ascent;

                // Frame text carries no inline pictures, so no descent is needed.
                render_paragraph_lines(
                    content,
                    &lines,
                    &Alignment::Left,
                    frame_x,
                    content_width,
                    frame_baseline,
                    font_size,
                    (font_size * ascender_ratio, 0.0),
                    lines.len(),
                    0,
                    links,
                    0.0,
                    ctx.fonts,
                    None,
                    gradient_specs,
                    None,
                    None,
                    None,
                );
                // Frame paragraphs are out-of-flow: do not advance cursor_y
                pi += 1;
            }
            Block::Paragraph(para) => {
                let has_para_image = para.image.is_some();
                let has_field_code = para.runs.iter().any(|r| r.field_code.is_some());
                let text_empty = !has_field_code && is_text_empty(&para.runs);

                cursor_y -= hf_paragraph_gap(prev_para, prev_space_after, para);
                prev_para = Some(para);

                let substituted_runs = substitute_hf_runs(
                    &para.runs,
                    page_num,
                    total_pages,
                    styleref_values,
                    page_num_format,
                );

                let (font_size, tallest_lhr, tallest_ar) =
                    tallest_run_metrics(&substituted_runs, ctx.fonts);
                let ascender_ratio = tallest_ar.unwrap_or(0.75);
                let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
                // Mirrors compute_header_height: a line overlapping a frame the
                // text may not sit beside moves below it.
                cursor_y = sp.page_height
                    - below_blocking_frames(sp.page_height - cursor_y, line_h, &bands);
                let slot_top = cursor_y;
                cursor_y -= top_border_band(para) + float_text_push(para, text_width);

                // DrawingML lines live in connectors rather than pictures or
                // textboxes. Empty header/footer paragraphs still carry their
                // paragraph-relative anchors.
                for connector in &para.connectors {
                    let (x, y) = super::positioning::connector_top_left(
                        connector,
                        sp,
                        sp.margin_left,
                        text_width,
                        text_width,
                        slot_top,
                    );
                    super::positioning::render_connector(connector, content, x, y);
                }

                // Paragraph borders span the laid-out height, so each exit below
                // draws them once it knows it (ut_koer: a header staff image
                // with a bottom border); empty bordered paragraphs render too.
                let bdr = &para.borders;
                let (box_left, box_right, box_top) =
                    (sp.margin_left, sp.margin_left + text_width, slot_top);
                let draw_para_borders = |content: &mut Content, box_bottom: f32| {
                    let draw_h_border =
                        |content: &mut Content, b: &crate::model::ParagraphBorder, y: f32| {
                            stroke_segment(
                                content,
                                (box_left, y),
                                (box_right, y),
                                b.width_pt,
                                Some(b.color),
                            );
                        };
                    if let Some(b) = &bdr.top {
                        draw_h_border(content, b, box_top - b.width_pt / 2.0);
                    }
                    if let Some(b) = &bdr.bottom {
                        // The stroke sits `space` below the box, outside it,
                        // as the body path's bottom pad places it.
                        draw_h_border(content, b, box_bottom - b.space_pt - b.width_pt / 2.0);
                    }
                };

                let mut baseline_y = cursor_y
                    - super::layout::boxed_line_ascent(
                        effective_ls,
                        line_h,
                        font_size,
                        tallest_lhr,
                        tallest_ar,
                        &substituted_runs,
                        ctx.fonts,
                    )
                    .unwrap_or(font_size * ascender_ratio);

                // Render textboxes
                for tb in &para.textboxes {
                    let tb_height = super::textbox_render::textbox_height(tb, ctx);
                    let tb_x = super::resolve_h_position(
                        tb.h_relative_from,
                        &tb.h_position,
                        tb.width_pt,
                        sp,
                        sp.margin_left,
                        text_width,
                        text_width,
                    );
                    let tb_y_top = resolve_tb_y_top(
                        tb.v_relative_from,
                        &tb.v_position,
                        tb_height,
                        sp,
                        slot_top,
                    );

                    if let Some(ref fill) = tb.fill {
                        super::render_shape_fill(
                            content,
                            fill,
                            tb_x,
                            tb_y_top - tb_height,
                            tb.width_pt,
                            tb_height,
                            &tb.shape_type,
                            gradient_specs,
                        );
                    }

                    let content_x = tb_x + tb.margin_left;
                    let content_w = (tb.width_pt - tb.margin_left - tb.margin_right).max(0.0);

                    // Word vertically anchors textbox content per bodyPr anchor (ctr/b).
                    // The body textbox path (textbox_render::render_single_textbox) applies this,
                    // but the header/footer loop is a separate renderer that historically
                    // top-anchored everything. Mirror the body math here. The per-paragraph
                    // height calc must match the rendering advances below (block image / empty /
                    // text), so keep the two in sync.
                    let anchor_offset = match tb.text_anchor {
                        TextAnchor::Top => 0.0,
                        TextAnchor::Middle | TextAnchor::Bottom => {
                            let tp_height = |tp: &Paragraph| -> f32 {
                                let tp_ls = tp.line_spacing.unwrap_or(ctx.doc_line_spacing);
                                let tp_text_w =
                                    (content_w - tp.indent_left - tp.indent_right).max(1.0);
                                let tp_hanging = if !tp.list_label.is_empty() {
                                    if tp.indent_first_line > 0.0 && tp.indent_hanging == 0.0 {
                                        -tp.indent_first_line
                                    } else {
                                        0.0
                                    }
                                } else if tp.indent_hanging > 0.0 {
                                    tp.indent_hanging
                                } else {
                                    -tp.indent_first_line
                                };
                                if let Some(img) =
                                    super::textbox_render::textbox_para_block_image(tp)
                                {
                                    return tp.space_before
                                        + img.display_height
                                        + img.layout_extra_height
                                        + tp.space_after;
                                }
                                let inline_imgs: HashMap<usize, String> =
                                    if tp.runs.iter().any(|r| r.inline_image.is_some()) {
                                        tp.runs
                                            .iter()
                                            .enumerate()
                                            .filter_map(|(ri, run)| {
                                                let img = run.inline_image.as_ref()?;
                                                ctx.textbox_image_names
                                                    .get(&img.key())
                                                    .map(|name| (ri, name.clone()))
                                            })
                                            .collect()
                                    } else {
                                        HashMap::new()
                                    };
                                let tb_lines = build_lines(
                                    &tp.runs,
                                    ctx,
                                    tp_text_w,
                                    ctx.cjk(true, tp.alignment),
                                    &LineOpts {
                                        inline_images: Some(&inline_imgs),
                                        tab_stops: &tp.tab_stops,
                                        indent_left: tp.indent_left,
                                        indent_right: tp.indent_right,
                                        hanging: tp_hanging,
                                        ..Default::default()
                                    },
                                );
                                if tb_lines.is_empty() {
                                    let (fs, _, _) = tallest_run_metrics(&tp.runs, ctx.fonts);
                                    let lh = resolve_line_h(tp_ls, fs, None);
                                    return tp.space_before + lh + tp.space_after;
                                }
                                let (tb_fs, _, tb_ar) = tallest_run_metrics(&tp.runs, ctx.fonts);
                                let tb_line_h = resolve_line_h(tp_ls, tb_fs, tb_ar);
                                tp.space_before
                                    + (tb_lines.len() as f32) * tb_line_h
                                    + tp.space_after
                            };
                            let total_h: f32 = tb.paragraphs.iter().map(tp_height).sum();
                            let available = (tb_height - tb.margin_top - tb.margin_bottom).max(0.0);
                            let gap = (available - total_h).max(0.0);
                            match tb.text_anchor {
                                TextAnchor::Middle => gap / 2.0,
                                TextAnchor::Bottom => gap,
                                TextAnchor::Top => 0.0,
                            }
                        }
                    };
                    let mut tb_cursor = tb_y_top - tb.margin_top - anchor_offset;
                    for tp in &tb.paragraphs {
                        let tp_ls = tp.line_spacing.unwrap_or(ctx.doc_line_spacing);
                        let tp_text_w = (content_w - tp.indent_left - tp.indent_right).max(1.0);
                        let tp_hanging = if !tp.list_label.is_empty() {
                            if tp.indent_first_line > 0.0 && tp.indent_hanging == 0.0 {
                                -tp.indent_first_line
                            } else {
                                0.0
                            }
                        } else if tp.indent_hanging > 0.0 {
                            tp.indent_hanging
                        } else {
                            -tp.indent_first_line
                        };

                        if let Some(img) = super::textbox_render::textbox_para_block_image(tp) {
                            if let Some(pdf_name) = ctx.textbox_image_names.get(&img.key()) {
                                let img_x = content_x
                                    + tp.indent_left
                                    + align_offset(
                                        tp.alignment,
                                        (tp_text_w - img.display_width).max(0.0),
                                    );
                                let img_y = tb_cursor - tp.space_before - img.display_height;
                                super::smartart::render_image_with_clip(
                                    content,
                                    pdf_name,
                                    img_x,
                                    img_y,
                                    img.display_width,
                                    img.display_height,
                                    img.clip_geometry.as_ref(),
                                );
                            }
                            tb_cursor -= tp.space_before
                                + img.display_height
                                + img.layout_extra_height
                                + tp.space_after;
                            continue;
                        }

                        let inline_imgs: HashMap<usize, String> =
                            if tp.runs.iter().any(|r| r.inline_image.is_some()) {
                                tp.runs
                                    .iter()
                                    .enumerate()
                                    .filter_map(|(ri, run)| {
                                        let img = run.inline_image.as_ref()?;
                                        ctx.textbox_image_names
                                            .get(&img.key())
                                            .map(|name| (ri, name.clone()))
                                    })
                                    .collect()
                            } else {
                                HashMap::new()
                            };
                        let tb_lines = build_lines(
                            &tp.runs,
                            ctx,
                            tp_text_w,
                            ctx.cjk(true, tp.alignment),
                            &LineOpts {
                                inline_images: Some(&inline_imgs),
                                tab_stops: &tp.tab_stops,
                                indent_left: tp.indent_left,
                                indent_right: tp.indent_right,
                                hanging: tp_hanging,
                                ..Default::default()
                            },
                        );
                        if tb_lines.is_empty() {
                            let (fs, _, _) = tallest_run_metrics(&tp.runs, ctx.fonts);
                            let lh = resolve_line_h(tp_ls, fs, None);
                            tb_cursor -= tp.space_before + lh + tp.space_after;
                            continue;
                        }
                        let (tb_fs, _, tb_ar) = tallest_run_metrics(&tp.runs, ctx.fonts);
                        let tb_ascender = tb_ar.unwrap_or(0.75);
                        let tb_line_h = resolve_line_h(tp_ls, tb_fs, tb_ar);
                        let tb_baseline = tb_cursor - tp.space_before - tb_fs * tb_ascender;
                        let tb_metrics = (
                            tb_fs * tb_ascender,
                            if inline_imgs.is_empty() {
                                0.0
                            } else {
                                picture_line_bottom(&tp.runs, tp, ctx.fonts, tp_ls)
                            },
                        );
                        super::render_list_label(
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
                            content_x + tp.indent_left,
                            tp_text_w,
                            tb_baseline,
                            tb_line_h,
                            tb_metrics,
                            tb_lines.len(),
                            0,
                            links,
                            0.0,
                            ctx.fonts,
                            None,
                            gradient_specs,
                            None,
                            None,
                            None,
                        );
                        tb_cursor -=
                            tp.space_before + (tb_lines.len() as f32) * tb_line_h + tp.space_after;
                    }
                }

                // Render floating images and register float zones
                for (fi_idx, fi) in para.floating_images.iter().enumerate() {
                    if let Some(pdf_name) = floating_image_names.get(&(pi, fi_idx)) {
                        let img = &fi.image;
                        let fi_x = super::resolve_h_position(
                            fi.h_relative_from,
                            &fi.h_position,
                            img.display_width,
                            sp,
                            sp.margin_left,
                            text_width,
                            text_width,
                        );
                        let fi_y_top = super::resolve_fi_y_top(fi, sp, slot_top);
                        super::smartart::render_image_with_clip(
                            content,
                            pdf_name,
                            fi_x,
                            fi_y_top - img.display_height,
                            img.display_width,
                            img.display_height,
                            img.clip_geometry.as_ref(),
                        );
                        // Register float zone for wrapping images
                        if fi.wrap_type.wraps_beside() {
                            hdr_fz.push(HfFloatZone::for_float(fi, fi_x, fi_y_top));
                        }
                    }
                }

                if (has_para_image || text_empty) && para.content_height > 0.0 {
                    if let Some(pdf_name) = para_image_names.get(&pi) {
                        let img = para.image.as_ref().unwrap();
                        let natural_line_h = font_size * tallest_lhr.unwrap_or(1.2);
                        let bottom_depth = if para.content_height > line_h {
                            img.layout_extra_top + img.display_height
                        } else {
                            para.content_height.max(natural_line_h)
                                - (img.layout_extra_height - img.layout_extra_top)
                        };
                        let y_bottom = cursor_y - bottom_depth;
                        let x = sp.margin_left
                            + align_offset(
                                para.alignment,
                                (text_width - img.display_width).max(0.0),
                            );
                        let hf_fx = pc.effect_para_names.get(&pi);
                        if let Some(ref shadow) = img.shadow {
                            super::color::draw_image_shadow(
                                content,
                                shadow,
                                x,
                                y_bottom,
                                img.display_width,
                                img.display_height,
                                hf_fx.and_then(|fx| fx.shadow.as_deref()),
                            );
                        }
                        if let Some(ref glow) = img.glow {
                            super::color::draw_image_glow(
                                content,
                                glow,
                                x,
                                y_bottom,
                                img.display_width,
                                img.display_height,
                                hf_fx.and_then(|fx| fx.glow.as_deref()),
                            );
                        }
                        super::smartart::render_image_with_clip(
                            content,
                            pdf_name,
                            x,
                            y_bottom,
                            img.display_width,
                            img.display_height,
                            img.clip_geometry.as_ref(),
                        );
                        if let Some(sc) = img.stroke_color {
                            super::smartart::stroke_image_border(
                                content,
                                x,
                                y_bottom,
                                img.display_width,
                                img.display_height,
                                sc,
                                img.stroke_width,
                                img.clip_geometry.as_ref(),
                            );
                        }
                    }
                    draw_para_borders(content, cursor_y - line_h);
                    cursor_y -= line_h + bottom_border_band(para);
                    prev_space_after = para.space_after;
                    pi += 1;
                    continue;
                }

                // Inline pictures sit on the baseline and grow their line to the
                // picture plus the text descent (`inline_line_advance`).
                let picture_bottom = if runs_max_image_h(&para.runs) > 0.0 {
                    picture_line_bottom(&substituted_runs, para, ctx.fonts, effective_ls)
                } else {
                    0.0
                };

                // VML horizontal rules (o:hr) are carried on otherwise-empty
                // paragraphs, so draw them before the text_empty skip below —
                // mirrors the body render path in pdf::mod.
                if let Some(ref hr) = para.horizontal_rule {
                    draw_horizontal_rule(
                        content,
                        para,
                        hr,
                        sp.margin_left,
                        text_width,
                        cursor_y - line_h,
                    );
                }

                if text_empty {
                    let mut advance = line_h
                        + super::mark_position_stretch(para, effective_ls, false)
                        + bottom_border_band(para);
                    // TopAndBottom textboxes push content below them
                    for tb in &para.textboxes {
                        if matches!(tb.wrap_type, WrapType::TopAndBottom) {
                            let tb_bottom_y = match tb.v_relative_from {
                                VRelativeFrom::Page => {
                                    sp.page_height
                                        - (tb.v_offset_pt + tb.height_pt + tb.dist_bottom)
                                }
                                _ => cursor_y - tb.v_offset_pt - tb.height_pt - tb.dist_bottom,
                            };
                            let needed = (cursor_y - tb_bottom_y).max(0.0);
                            advance = advance.max(needed);
                        }
                    }
                    draw_para_borders(content, cursor_y - line_h);
                    cursor_y -= advance;
                    prev_space_after = para.space_after;
                    pi += 1;
                    continue;
                }

                let block_inline_images: HashMap<usize, String> = inline_image_names
                    .iter()
                    .filter(|((pi2, _), _)| *pi2 == pi)
                    .map(|((_, ri), name)| (*ri, name.clone()))
                    .collect();

                let mut para_text_x = sp.margin_left + para.indent_left;
                let mut para_text_width =
                    (text_width - para.indent_left - para.indent_right).max(1.0);
                let text_hanging = if !para.list_label.is_empty() {
                    if let Some(nts) = para.num_level_tab_stop {
                        if nts < para.indent_left
                            && (para.indent_left - para.indent_hanging).abs() < 0.5
                        {
                            (para.indent_left - nts).max(0.0)
                        } else if para.indent_first_line > 0.0 && para.indent_hanging == 0.0 {
                            -para.indent_first_line
                        } else {
                            0.0
                        }
                    } else if para.indent_first_line > 0.0 && para.indent_hanging == 0.0 {
                        -para.indent_first_line
                    } else {
                        0.0
                    }
                } else if para.indent_hanging > 0.0 {
                    para.indent_hanging
                } else {
                    -para.indent_first_line
                };

                // Narrow text width when wrapping floating images overlap this paragraph.
                // Combine same-paragraph floats and cross-paragraph float zones so a
                // logo on each side of a centered letterhead constrains both edges.
                let mut hdr_line_geom: Option<Vec<(f32, f32)>> = None;
                let mut zones: Vec<HfFloatZone> = para
                    .floating_images
                    .iter()
                    .filter(|fi| fi.wrap_type.wraps_beside())
                    .map(|fi| {
                        let fi_x = super::resolve_h_position(
                            fi.h_relative_from,
                            &fi.h_position,
                            fi.image.display_width,
                            sp,
                            sp.margin_left,
                            text_width,
                            text_width,
                        );
                        let fi_y_top = super::resolve_fi_y_top(fi, sp, slot_top);
                        HfFloatZone::for_float(fi, fi_x, fi_y_top)
                    })
                    .collect();
                zones.extend(hdr_fz.iter().copied());

                {
                    let col_x = sp.margin_left;
                    let col_right = sp.margin_left + text_width;

                    // Combined text bounds over [y_lo, y_hi]: a float with more room
                    // on its right raises the left bound, otherwise it lowers the
                    // right bound. Returns None when no zone overlaps.
                    let bounds_at = |y_hi: f32, y_lo: f32| -> Option<(f32, f32)> {
                        let mut left = col_x;
                        let mut right = col_right;
                        let mut hit = false;
                        for z in &zones {
                            if y_hi > z.bottom && y_lo < z.top {
                                hit = true;
                                let space_right = col_right - (z.right + z.dist_right);
                                let space_left = (z.left - z.dist_left) - col_x;
                                if space_right >= space_left && space_right >= 36.0 {
                                    left = left.max(z.right + z.dist_right);
                                } else if space_left >= 36.0 {
                                    right = right.min(z.left - z.dist_left);
                                }
                            }
                        }
                        hit.then_some((left, right))
                    };

                    // Word measures paragraph indents from the column edge and lets
                    // float bounds clip the region — indents are NOT added on top of
                    // the float edges.
                    let indented = |lb: f32, rb: f32| -> (f32, f32) {
                        let tx = (col_x + para.indent_left).max(lb);
                        let tr = (col_right - para.indent_right).min(rb);
                        (tx, (tr - tx).max(1.0))
                    };

                    if let Some((lb, rb)) = bounds_at(cursor_y, baseline_y) {
                        (para_text_x, para_text_width) = indented(lb, rb);

                        // Build per-line geometry for multi-line paragraphs that may
                        // span above and through the float zones
                        let ascender_ratio_e = tallest_ar.unwrap_or(0.75);
                        let full_w = (text_width - para.indent_left - para.indent_right).max(1.0);
                        let deepest_bottom =
                            zones.iter().map(|z| z.bottom).fold(f32::INFINITY, f32::min);
                        let max_lines = ((cursor_y - deepest_bottom) / line_h).ceil() as usize + 10;
                        let max_lines = max_lines.max(20);
                        let mut geom = Vec::with_capacity(max_lines);
                        for i in 0..max_lines {
                            let y = cursor_y - font_size * ascender_ratio_e - i as f32 * line_h;
                            match bounds_at(y, y) {
                                Some((lb, rb)) => geom.push(indented(lb, rb)),
                                None => geom.push((col_x + para.indent_left, full_w)),
                            }
                        }
                        hdr_line_geom = Some(geom);
                    }
                }

                let per_line_widths: Option<Vec<f32>> = hdr_line_geom
                    .as_ref()
                    .map(|g| g.iter().map(|&(_, w)| w).collect());

                let mut lines = build_lines(
                    &substituted_runs,
                    ctx,
                    para_text_width,
                    ctx.cjk(true, para.alignment),
                    &LineOpts {
                        inline_images: Some(&block_inline_images),
                        tab_stops: &para.tab_stops,
                        indent_left: para.indent_left,
                        indent_right: para.indent_right,
                        hanging: text_hanging,
                        per_line_widths: per_line_widths.as_deref(),
                        ..Default::default()
                    },
                );

                // A single short picture sits on the paragraph mark's natural
                // line baseline, like the body path, with its bottom effect
                // extent below that baseline. An empty tab can keep the picture
                // in runs instead of Paragraph.image, so handle that slot too.
                if lines.len() == 1 && substituted_runs.iter().all(|r| r.text.trim().is_empty()) {
                    let mut images = substituted_runs
                        .iter()
                        .filter_map(|r| r.inline_image.as_ref());
                    if let Some(image) = images.next()
                        && images.next().is_none()
                        && image.display_height + image.layout_extra_height <= line_h
                    {
                        baseline_y = cursor_y - font_size * tallest_lhr.unwrap_or(1.2)
                            + image.layout_extra_height
                            - image.layout_extra_top;
                    }
                }
                let wrapped_pictures = lines.len() > 1
                    && substituted_runs
                        .iter()
                        .filter(|r| r.inline_image.is_some())
                        .count()
                        > 1
                    && substituted_runs.iter().all(|r| r.text.trim().is_empty());
                // A wrapped image-only running head retains the paragraph
                // mark's descent between its picture lines. effectExtent's
                // bottom also sits below the baseline, rather than moving the
                // visible picture down. Keep other picture-line paths intact.
                if wrapped_pictures {
                    let images = substituted_runs.iter().enumerate().filter_map(|(ri, r)| {
                        block_inline_images
                            .contains_key(&ri)
                            .then_some(r.inline_image.as_ref())
                            .flatten()
                    });
                    for (chunk, image) in lines
                        .iter_mut()
                        .flat_map(|l| &mut l.chunks)
                        .filter(|c| c.inline_image_name.is_some())
                        .zip(images)
                    {
                        chunk.inline_image_extra_height = image.layout_extra_height;
                        chunk.y_offset = image.layout_extra_height - image.layout_extra_top;
                    }
                }
                let wrapped_picture_bottom = if wrapped_pictures {
                    font_size * super::layout::descender_ratio(tallest_lhr, tallest_ar)
                } else {
                    picture_bottom
                };
                let metrics = (font_size * ascender_ratio, wrapped_picture_bottom);
                // Each line as tall as its own runs, as in the body.
                if !matches!(effective_ls, LineSpacing::Exact(_)) {
                    size_lines_by_own_runs(&mut lines, ctx.fonts, effective_ls, line_h, metrics.0);
                }
                render_paragraph_lines(
                    content,
                    &lines,
                    &para.alignment,
                    para_text_x,
                    para_text_width,
                    baseline_y,
                    line_h,
                    metrics,
                    lines.len(),
                    0,
                    links,
                    text_hanging,
                    ctx.fonts,
                    hdr_line_geom.as_deref(),
                    gradient_specs,
                    None,
                    None,
                    None,
                );

                let trailing_descent = if wrapped_pictures {
                    wrapped_picture_trailing_descent(
                        para,
                        font_size,
                        wrapped_picture_bottom,
                        effective_ls,
                    )
                } else {
                    0.0
                };
                let para_h = lines_height(&lines, line_h, metrics) - trailing_descent;
                draw_para_borders(content, cursor_y - para_h);
                cursor_y -= para_h + bottom_border_band(para);
                prev_space_after = para.space_after;
                pi += 1;
            }
        }
    }
}

/// Which header/footer variant a page shows. A section that lacks that variant
/// inherits it from earlier sections, and shows none if no section defines it
/// (§17.10.5) — it never falls back to the default variant.
#[derive(Clone, Copy)]
enum HfVariant {
    Default,
    First,
    Even,
}

fn hf_variant(
    even_and_odd_headers: bool,
    sp: &SectionProperties,
    is_first_page: bool,
    page_num: usize,
) -> HfVariant {
    if is_first_page && sp.different_first_page {
        HfVariant::First
    } else if even_and_odd_headers && page_num.is_multiple_of(2) {
        HfVariant::Even
    } else {
        HfVariant::Default
    }
}

/// A section's header (or footer) of `variant`, else the nearest earlier
/// section's, with the index of the section that owns it.
fn inherited_hf(
    sections: &[crate::model::Section],
    idx: usize,
    variant: HfVariant,
    is_header: bool,
) -> Option<(&HeaderFooter, usize)> {
    sections[..=idx]
        .iter()
        .enumerate()
        .rev()
        .find_map(|(i, s)| {
            let p = &s.properties;
            let slot = match (variant, is_header) {
                (HfVariant::Default, true) => &p.header_default,
                (HfVariant::First, true) => &p.header_first,
                (HfVariant::Even, true) => &p.header_even,
                (HfVariant::Default, false) => &p.footer_default,
                (HfVariant::First, false) => &p.footer_first,
                (HfVariant::Even, false) => &p.footer_even,
            };
            slot.as_ref().map(|hf| (hf, i))
        })
}

/// Resolve which header to use for a given page, walking sections backward
/// for inheritance. Returns `(header_data, hf_type_id, section_index)`.
pub(super) fn resolve_header_for_page(
    doc: &Document,
    section_idx: usize,
    is_first_page: bool,
    page_num: usize,
) -> (Option<&HeaderFooter>, u8, usize) {
    let sp = &doc.sections[section_idx].properties;
    let variant = hf_variant(doc.even_and_odd_headers, sp, is_first_page, page_num);
    let t = match variant {
        HfVariant::Default => 0,
        HfVariant::First => 1,
        HfVariant::Even => 4,
    };
    match inherited_hf(&doc.sections, section_idx, variant, true) {
        Some((hf, idx)) => (Some(hf), t, idx),
        None => (None, t, section_idx),
    }
}

/// Resolve which footer to use for a given page, walking sections backward
/// for inheritance. Returns `(footer_data, hf_type_id, section_index)`.
pub(super) fn resolve_footer_for_page(
    doc: &Document,
    section_idx: usize,
    is_first_page: bool,
    page_num: usize,
) -> (Option<&HeaderFooter>, u8, usize) {
    let sp = &doc.sections[section_idx].properties;
    let variant = hf_variant(doc.even_and_odd_headers, sp, is_first_page, page_num);
    let t = match variant {
        HfVariant::Default => 2,
        HfVariant::First => 3,
        HfVariant::Even => 5,
    };
    match inherited_hf(&doc.sections, section_idx, variant, false) {
        Some((hf, idx)) => (Some(hf), t, idx),
        None => (None, t, section_idx),
    }
}

/// The gap above a header/footer paragraph: contextual spacing drops both
/// sides between same-style paragraphs, as in the body (two 18pt-before
/// motion header lines step 15.5 in Word, not 33.6).
pub(super) fn hf_paragraph_gap(
    prev: Option<&Paragraph>,
    prev_space_after: f32,
    para: &Paragraph,
) -> f32 {
    let after = if prev.is_some_and(|p| super::helpers::drops_contextual_spacing(p, Some(para))) {
        0.0
    } else {
        prev_space_after
    };
    after.max(super::helpers::effective_space_before(para, prev))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn if_over_stylerefs_takes_the_page_value() {
        let sr = |number| {
            IfPart::Field(FieldCode::StyleRef {
                name: "CharPartNo".into(),
                number,
            })
        };
        let text = |t: &str| IfPart::Text(t.into());
        // IF {STYLEREF CharPartNo \n} = 0 "{STYLEREF CharPartNo}" "Part {STYLEREF CharPartNo \n}"
        let parts = [
            text(" IF "),
            sr(true),
            text(" = 0 \""),
            sr(false),
            text("\" \"Part "),
            sr(true),
            text("\""),
        ];
        let mut styleref_values = HashMap::new();
        let eval = |values: &HashMap<String, String>| {
            FieldValues {
                page_num: 1,
                total_pages: 1,
                styleref_values: values,
                page_num_format: None,
            }
            .eval(&FieldCode::If(parts.to_vec()))
        };
        assert_eq!(eval(&styleref_values), None);
        styleref_values.insert(styleref_key("charpartno", false), "Part 3".to_string());
        styleref_values.insert(styleref_key("charpartno", true), "0".to_string());
        assert_eq!(eval(&styleref_values).as_deref(), Some("Part 3"));
        styleref_values.insert(styleref_key("charpartno", true), "4".to_string());
        assert_eq!(eval(&styleref_values).as_deref(), Some("Part 4"));
        assert_eq!(field_args("a<>\"b c\" 2"), ["a", "<>", "b c", "2"]);
    }
}
