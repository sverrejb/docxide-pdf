use std::collections::HashMap;

use pdf_writer::Content;

use crate::model::{
    Alignment, Block, Document, FieldCode, FrameProperties, HRelativeFrom, HeaderFooter, IfPart,
    LineSpacing, Paragraph, Run, SectionProperties, TextAnchor, VRelativeFrom, VerticalPosition,
    WrapType,
};

use super::color::stroke_border;
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

pub(super) use crate::model::styleref_key;

/// A PAGE or NUMPAGES value in its numFmt (None = decimal).
pub(super) fn page_field_text(n: usize, fmt: Option<&str>) -> String {
    fmt.map_or_else(
        || n.to_string(),
        |fmt| crate::docx::numbering::format_number(n as u32, fmt),
    )
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
            FieldCode::Page(fmt) => page_field_text(self.page_num, fmt.or(self.page_num_format)),
            FieldCode::NumPages(fmt) => page_field_text(self.total_pages, *fmt),
            FieldCode::StyleRef { name, form } => self
                .styleref_values
                .get(&styleref_key(name, *form))
                .cloned()
                .unwrap_or_default(),
            FieldCode::PageRef(_) | FieldCode::Seq { .. } => return None,
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

/// Header lines size as body lines do: runs of nothing but spaces or tabs
/// size nothing, and a line left with nothing takes the paragraph mark's
/// metrics. A header of a tab and spaces in 11pt Tahoma under a 20pt mark is
/// a 20pt Arial line in Word (body probes: spaces, tabs or both under a 20pt
/// mark give the empty mark's line, a 20pt tab under an 8pt mark an 8pt
/// line), and "GOVERNOR", 113 spaces and "SECRETARY" in 10pt with the spaces
/// at 11pt is a 10pt line.
fn hf_run_metrics(
    runs: &[crate::model::Run],
    para: &Paragraph,
    fonts: &HashMap<String, crate::fonts::FontEntry>,
) -> (f32, Option<f32>, Option<f32>) {
    let glyphs = super::layout::tallest_glyph_run_metrics(runs, fonts);
    let nothing_sizes = glyphs.1.is_none()
        && runs.iter().any(|r| !r.vanish)
        && !runs
            .iter()
            .any(|r| r.is_line_break || r.inline_image.is_some());
    let mark = para.paragraph_mark_font_size.filter(|_| nothing_sizes);
    let Some(fs) = mark else {
        return if glyphs.1.is_some() {
            glyphs
        } else {
            tallest_run_metrics(runs, fonts)
        };
    };
    let (lhr, ar) = para
        .paragraph_mark_font_name
        .as_deref()
        .and_then(|n| fonts.get(n))
        .map_or((None, None), |e| super::layout::run_line_metrics(e, ""));
    (fs, lhr, ar)
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

fn blocking_table_bands(
    hf: &HeaderFooter,
    sp: &SectionProperties,
    ctx: &RenderContext,
) -> Vec<(f32, f32)> {
    hf.blocks
        .iter()
        .filter_map(|b| match b {
            Block::Table(t) => blocking_table_band(t, sp, ctx),
            _ => None,
        })
        .collect()
}

/// Where a header paragraph starts (down from the page top, before its
/// space before) once blocking floating tables have pushed it: the whole
/// paragraph, its space before and its first line or picture, moves below
/// (Word probes: 20pt before lands 20pt under the table, a 59pt logo below a
/// table that starts 13pt down still moves).
fn below_blocking_tables(
    top: f32,
    gap: f32,
    para: &Paragraph,
    line_h: f32,
    bands: &[(f32, f32)],
) -> f32 {
    below_blocking_frames(top, gap + line_h.max(runs_max_image_h(&para.runs)), bands)
}

/// A page- or margin-anchored floating table that leaves no room beside it
/// across the text column: the header lines after it start below its rows'
/// content. Word probes on a letterhead table over a logo paragraph: the
/// logo drops by the row's text lines (one line 12.65, three lines 37.95),
/// moves with tblpY, ignores the cells' space after, and stays at the
/// header top when the table is narrow or lower than the logo.
fn blocking_table_band(
    table: &crate::model::Table,
    sp: &SectionProperties,
    ctx: &RenderContext,
) -> Option<(f32, f32)> {
    let pos = table.position.as_ref()?;
    if pos.v_anchor == "text" {
        return None;
    }
    let text_width = sp.text_width();
    let fp = super::FloatingTablePos::resolve(table, pos, sp, sp.margin_left, text_width, 0.0, ctx);
    let width: f32 = table.col_widths.iter().sum();
    // ponytail: "no room" = covers the column; measure the side gaps if a
    // partly covering header table turns up.
    if fp.x - fp.left_from_text > sp.margin_left + 1.0
        || fp.x + width + fp.right_from_text < sp.margin_left + text_width - 1.0
    {
        return None;
    }
    let trailing_after = table
        .rows
        .last()
        .into_iter()
        .flat_map(|r| &r.cells)
        .filter_map(|c| match c.content.last() {
            Some(Block::Paragraph(p)) => Some(p.space_after),
            _ => None,
        })
        .fold(0.0f32, f32::max);
    let height = table::compute_hf_table_height(table, ctx, text_width) - trailing_after;
    let top = sp.page_height - fp.y;
    Some((top, top + height.max(0.0)))
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
        ctx.cjk(true, para),
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
        size_lines_by_own_runs(&mut lines, ctx.fonts, spacing, line_h, metrics.0, para.text_alignment == crate::model::TextAlignment::Top);
    }
    let trailing_descent = wrapped_picture_trailing_descent(para, font_size, metrics.1, spacing);
    Some(lines_height(&lines, line_h, metrics) - trailing_descent)
}

// Inline groups are flattened into an invisible flow box and positioned leaves.
// Keep the terminal tab on the following line when the group fills the line,
// and advance past the box before laying out subsequent header paragraphs.
fn inline_canvas_tab_height(para: &Paragraph, width: f32, line_h: f32) -> Option<f32> {
    let terminal_tab = para
        .runs
        .iter()
        .rev()
        .find(|run| {
            run.is_tab || run.is_line_break || !run.text.is_empty() || run.inline_image.is_some()
        })
        .is_some_and(|run| run.is_tab);
    if !terminal_tab {
        return None;
    }
    let available = (width - para.indent_left - para.indent_right - para.indent_first_line
        + para.indent_hanging)
        .max(0.0);
    para.textboxes
        .iter()
        .filter(|tb| {
            tb.on_line
                && tb.no_text_wrap
                && tb.paragraphs.is_empty()
                && matches!(tb.wrap_type, WrapType::TopAndBottom)
        })
        .map(|tb| {
            tb.v_offset_pt
                + tb.height_pt
                + tb.dist_bottom
                + if tb.width_pt > available + 0.05 {
                    line_h
                } else {
                    0.0
                }
        })
        .reduce(f32::max)
}

fn compute_header_height(
    hf: &HeaderFooter,
    ctx: &RenderContext,
    sp: &SectionProperties,
    is_header: bool,
) -> f32 {
    header_height(hf, ctx, sp, is_header, true)
}

/// A header/footer (by address, it lives as long as the document) and the
/// section geometry its measured height depends on.
pub(super) type HfHeightKey = (usize, [u32; 8], bool, bool);

/// `compute_header_height`, optionally without the floats: the extent of the
/// header's own text flow.
fn header_height(
    hf: &HeaderFooter,
    ctx: &RenderContext,
    sp: &SectionProperties,
    is_header: bool,
    with_floats: bool,
) -> f32 {
    let geometry = [
        sp.page_width,
        sp.page_height,
        sp.margin_left,
        sp.margin_right,
        sp.margin_top,
        sp.margin_bottom,
        sp.header_margin,
        sp.footer_margin,
    ]
    .map(f32::to_bits);
    let key = (
        hf as *const HeaderFooter as usize,
        geometry,
        is_header,
        with_floats,
    );
    if let Some(&h) = ctx.hf_heights.borrow().get(&key) {
        return h;
    }
    let h = measure_header_height(hf, ctx, sp, is_header, with_floats);
    ctx.hf_heights.borrow_mut().insert(key, h);
    h
}

fn measure_header_height(
    hf: &HeaderFooter,
    ctx: &RenderContext,
    sp: &SectionProperties,
    is_header: bool,
    with_floats: bool,
) -> f32 {
    let text_width = sp.text_width();
    let mut height = 0.0f32;
    let mut prev_space_after = 0.0f32;
    let mut prev_para: Option<&Paragraph> = None;
    // How far page- or margin-placed floats reach from the header (footer)
    // margin, whatever the flow above them.
    let mut placed_extent = 0.0f32;
    let (bands, table_bands) = if is_header {
        (
            blocking_frame_bands(hf, sp),
            blocking_table_bands(hf, sp, ctx),
        )
    } else {
        Default::default()
    };
    for block in &hf.blocks {
        match block {
            Block::Paragraph(para) if para.frame_props.is_some() => {
                // Frame paragraphs are out-of-flow; skip height contribution
            }
            Block::Paragraph(para) => {
                let gap = hf_paragraph_gap(prev_para, prev_space_after, para);
                prev_para = Some(para);
                let (font_size, tallest_lhr, _) = hf_run_metrics(&para.runs, para, ctx.fonts);
                let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
                height = below_blocking_tables(
                    sp.header_margin + height,
                    gap,
                    para,
                    line_h,
                    &table_bands,
                ) - sp.header_margin
                    + gap;
                height = below_blocking_frames(sp.header_margin + height, line_h, &bands)
                    - sp.header_margin;
                // Positioned runs and marks stretch their lines as in the render.
                let line_h = line_h
                    + if is_text_empty(&para.runs) {
                        super::mark_position_stretch(para, effective_ls, false, ctx.fonts)
                    } else {
                        position_stretch(&para.runs, effective_ls)
                    };
                // Mirrors the render loop's advance: a picture line is the
                // picture plus the text descent (`inline_line_advance`).
                let picture_h = runs_max_image_h(&para.runs);
                let wrapped_h = wrapped_header_picture_height(para, ctx, text_width);
                // A text paragraph is as many lines as it wraps to (Word bottom-
                // aligns a nine-line footer at the footer margin; one line per
                // paragraph put it 100pt low, off the page); its w:br breaks
                // are among those lines.
                let text_lines = (wrapped_h.is_none() && picture_h <= 0.0).then(|| {
                    build_lines(
                        &para.runs,
                        ctx,
                        (text_width - para.indent_left - para.indent_right).max(1.0),
                        ctx.cjk(true, para),
                        &LineOpts {
                            tab_stops: &para.tab_stops,
                            indent_left: para.indent_left,
                            indent_right: para.indent_right,
                            hanging: para.indent_hanging - para.indent_first_line,
                            ..Default::default()
                        },
                    )
                    .len()
                    .max(1)
                });
                let mut content_h = if let Some(h) = wrapped_h {
                    h
                } else if picture_h > 0.0 {
                    line_h.max(
                        picture_h + picture_line_bottom(&para.runs, para, ctx.fonts, effective_ls),
                    )
                } else {
                    text_lines.unwrap_or(1) as f32 * line_h
                } + if with_floats {
                    float_text_push(para, text_width)
                } else {
                    0.0
                };

                for fi in para.floating_images.iter().filter(|_| with_floats) {
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

                for tb in para.textboxes.iter().filter(|_| with_floats) {
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
                if is_header && with_floats {
                    if let Some(canvas_h) = inline_canvas_tab_height(para, text_width, line_h) {
                        content_h = content_h.max(canvas_h);
                    }
                }
                let br_count = para.runs.iter().filter(|r| r.is_line_break).count();
                if wrapped_h.is_none() && text_lines.is_none() {
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

/// The body top a page's header text alone would leave: where Word anchors
/// a page's first paragraph's floats when header floats push its lines lower
/// (a 133pt floating letterhead pushes the text, not the anchor).
pub(super) fn flow_slot_top(
    sp: &SectionProperties,
    is_first: bool,
    page_idx: usize,
    ctx: &RenderContext,
) -> f32 {
    let base = sp.page_height - sp.margin_top;
    match layout_hf(sp, is_first, page_idx, true, ctx) {
        Some(hf) if !sp.margin_top_fixed => {
            base.min(sp.page_height - sp.header_margin - header_height(hf, ctx, sp, true, false))
        }
        _ => base,
    }
}

/// The lowest bottom (page coordinates, text distance included) of the
/// page header's wrapping floats that overlap `left..right`. Body text flows
/// beside them; a table cannot, so one at the page top starts below them
/// (cote_dazur: two logos end 81.2pt from the top, where Word starts the
/// body table 10.4pt below the top margin).
// ponytail: floats anchor at the header's top line; a letterhead rarely
// anchors its logo in a later header paragraph.
pub(super) fn header_float_floor(
    sp: &SectionProperties,
    is_first: bool,
    page_idx: usize,
    ctx: &RenderContext,
    left: f32,
    right: f32,
) -> Option<f32> {
    let hf = layout_hf(sp, is_first, page_idx, true, ctx)?;
    let text_width = sp.text_width();
    let slot_top = sp.page_height - sp.header_margin;
    hf.blocks
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(para) => Some(para),
            Block::Table(_) => None,
        })
        .flat_map(|para| &para.floating_images)
        .filter(|fi| fi.wrap_type.wraps_beside() || matches!(fi.wrap_type, WrapType::TopAndBottom))
        .filter_map(|fi| {
            let fi_x = super::resolve_h_position(
                fi.h_relative_from,
                &fi.h_position,
                fi.image.display_width,
                sp,
                sp.margin_left,
                text_width,
                text_width,
            );
            let overlaps = fi_x - fi.dist_left < right && fi_x + fi.image.display_width + fi.dist_right > left;
            overlaps.then(|| {
                super::resolve_fi_y_top(fi, sp, slot_top) - fi.image.display_height - fi.dist_bottom
            })
        })
        .min_by(f32::total_cmp)
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
    let (bands, table_bands) = if is_header {
        (
            blocking_frame_bands(hf, sp),
            blocking_table_bands(hf, sp, ctx),
        )
    } else {
        Default::default()
    };
    for (bi, block) in hf.blocks.iter().enumerate() {
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
                    ctx.cjk(true, para),
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
                    if fp.width > 0.0 {
                        para.indent_left
                    } else {
                        0.0
                    },
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

                let gap = hf_paragraph_gap(prev_para, prev_space_after, para);
                prev_para = Some(para);

                let substituted_runs = substitute_hf_runs(
                    &para.runs,
                    page_num,
                    total_pages,
                    styleref_values,
                    page_num_format,
                );

                let (font_size, tallest_lhr, tallest_ar) =
                    hf_run_metrics(&substituted_runs, para, ctx.fonts);
                let ascender_ratio = tallest_ar.unwrap_or(0.75);
                let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
                cursor_y = sp.page_height
                    - below_blocking_tables(
                        sp.page_height - cursor_y,
                        gap,
                        para,
                        line_h,
                        &table_bands,
                    )
                    - gap;
                // Mirrors compute_header_height: a line overlapping a frame the
                // text may not sit beside moves below it.
                cursor_y = sp.page_height
                    - below_blocking_frames(sp.page_height - cursor_y, line_h, &bands);
                let slot_top = cursor_y;
                cursor_y -= top_border_band(para) + float_text_push(para, text_width);

                // Foreground anchored lines must cover an inline picture's white
                // background, so they wait for the late pass; behind-document
                // lines (and all lines when there is no picture) paint early.
                let defer_connectors = has_para_image
                    || para.runs.iter().any(|run| run.inline_image.is_some())
                    || para.textboxes.iter().any(|tb| tb.fill.is_some() && tb.behind_doc);
                let draw_connectors = |content: &mut Content, late_pass: bool| {
                    for connector in &para.connectors {
                        if (defer_connectors && !connector.behind_doc) != late_pass {
                            continue;
                        }
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
                };
                draw_connectors(content, false);

                // Paragraph borders span the laid-out height, so each exit below
                // draws them once it knows it (ut_koer: a header staff image
                // with a bottom border); empty bordered paragraphs render too.
                let bdr = &para.borders;
                let (box_left, box_right, box_top) =
                    (sp.margin_left, sp.margin_left + text_width, slot_top);
                let draw_para_borders = |content: &mut Content, box_bottom: f32| {
                    if let Some(b) = &bdr.top {
                        let y = box_top - b.extent() / 2.0;
                        stroke_border(content, b, (box_left, y), (box_right, y), -1.0);
                    }
                    if let Some(b) = &bdr.bottom {
                        // The stroke sits `space` below the box, outside it,
                        // as the body path's bottom pad places it.
                        let y = box_bottom - b.space_pt - b.extent() / 2.0;
                        stroke_border(content, b, (box_left, y), (box_right, y), 1.0);
                    }
                    // Side borders as in the body: their inner edge `space`
                    // plus Word's 1.47 / 1.73pt outside the text.
                    let v_bottom = box_bottom - bdr.bottom.as_ref().map_or(0.0, |b| b.space_pt);
                    let (left_x, right_x) = super::side_border_outsets(bdr, true);
                    for (b, x, inward) in [
                        (&bdr.left, box_left - left_x, 1.0),
                        (&bdr.right, box_right + right_x, -1.0),
                    ] {
                        if let Some(b) = b {
                            stroke_border(content, b, (x, box_top), (x, v_bottom), inward);
                        }
                    }
                };
                // Paragraph shading under the text, out to the side borders'
                // inner edges and down to the bottom border (a footer's blue
                // address bar with white text drew as white on white).
                let draw_para_shading = |content: &mut Content, box_bottom: f32| {
                    let Some(color) = para.shading else { return };
                    let (left_out, right_out) = super::side_border_outsets(bdr, false);
                    let (left, right) = (box_left - left_out, box_right + right_out);
                    let bottom = box_bottom - bdr.bottom.as_ref().map_or(0.0, |b| b.space_pt);
                    content.save_state();
                    super::color::fill_rgb(content, color);
                    content.rect(left, bottom, right - left, box_top - bottom);
                    content.fill_nonzero();
                    content.restore_state();
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

                // Paint in z-order: the shapes above the pictures after them
                // (a white "Agenda" title box over a banner picture).
                let pictures_first = para
                    .floating_images
                    .iter()
                    .map(|f| f.z_index)
                    .max()
                    .zip(para.textboxes.iter().map(|t| t.z_index).min())
                    .is_some_and(|(top_picture, low_box)| top_picture < low_box);
                for pictures in [pictures_first, !pictures_first] {
                    if pictures {
                        // Render floating images and register float zones,
                        // overlapping ones in z-order like the body does.
                        let mut by_z: Vec<_> = para.floating_images.iter().enumerate().collect();
                        by_z.sort_by_key(|(_, fi)| fi.z_index);
                        for (fi_idx, fi) in by_z {
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
                    } else {
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
                            let rotated = super::textbox_render::rotate_about_center(
                                content,
                                tb.rotation_deg,
                                tb_x,
                                tb_y_top,
                                tb.width_pt,
                                tb_height,
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
                            let content_w =
                                (tb.width_pt - tb.margin_left - tb.margin_right).max(0.0);

                            // WordArt (a header watermark) draws its glyphs as warped
                            // outlines, as the body path does.
                            if tb.text_warp.as_ref().is_some_and(|w| w.preset != "textNoShape")
                                && super::wordart::render_warped_textbox(
                                    tb, content, ctx.fonts, tb_x, tb_y_top, content_w,
                                )
                            {
                                if rotated {
                                    content.restore_state();
                                }
                                continue;
                            }

                            // Match Word and body textboxes: fixed boxes clip overflow.
                            let needs_clip = !matches!(tb.auto_fit, crate::model::AutoFit::Shape);
                            if needs_clip {
                                content.save_state();
                                content.rect(tb_x, tb_y_top - tb_height, tb.width_pt, tb_height);
                                content.clip_nonzero();
                                content.end_path();
                            }

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
                                            if tp.indent_first_line > 0.0
                                                && tp.indent_hanging == 0.0
                                            {
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
                                        let inline_imgs = super::layout::run_image_map(
                                            &tp.runs,
                                            ctx.textbox_image_names,
                                        );
                                        // Footer page numbers live in text boxes as often
                                        // as in plain paragraphs; without this the box
                                        // kept the template's cached PAGE/NUMPAGES text.
                                        let hf_runs = substitute_hf_runs(
                                            &tp.runs,
                                            page_num,
                                            total_pages,
                                            styleref_values,
                                            page_num_format,
                                        );
                                        let tb_lines = build_lines(
                                            &hf_runs,
                                            ctx,
                                            tp_text_w,
                                            ctx.cjk(true, tp),
                                            &LineOpts {
                                                inline_images: Some(&inline_imgs),
                                                tab_stops: &tp.tab_stops,
                                                indent_left: tp.indent_left,
                                                indent_right: tp.indent_right,
                                                hanging: tp_hanging,
                                                ..Default::default()
                                            },
                                        );
                                        let (fs, lhr, _) = tallest_run_metrics(&hf_runs, ctx.fonts);
                                        let line_h = resolve_line_h(tp_ls, fs, lhr);
                                        tp.space_before
                                            + tb_lines.len().max(1) as f32 * line_h
                                            + tp.space_after
                                    };
                                    let total_h: f32 = tb.paragraphs.iter().map(tp_height).sum();
                                    let available =
                                        (tb_height - tb.margin_top - tb.margin_bottom).max(0.0);
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
                                    if let Some(pdf_name) = ctx.textbox_image_names.get(&img.key())
                                    {
                                        let img_x = content_x
                                            + tp.indent_left
                                            + align_offset(
                                                tp.alignment,
                                                (tp_text_w - img.display_width).max(0.0),
                                            );
                                        let img_y =
                                            tb_cursor - tp.space_before - img.display_height;
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

                                let inline_imgs =
                                    super::layout::run_image_map(&tp.runs, ctx.textbox_image_names);
                                // Same substitution as the height pass above: the
                                // footer text box's PAGE/NUMPAGES must be evaluated
                                // for this page, not left at the cached value.
                                let hf_runs = substitute_hf_runs(
                                    &tp.runs,
                                    page_num,
                                    total_pages,
                                    styleref_values,
                                    page_num_format,
                                );
                                let tb_lines = build_lines(
                                    &hf_runs,
                                    ctx,
                                    tp_text_w,
                                    ctx.cjk(true, tp),
                                    &LineOpts {
                                        inline_images: Some(&inline_imgs),
                                        tab_stops: &tp.tab_stops,
                                        indent_left: tp.indent_left,
                                        indent_right: tp.indent_right,
                                        hanging: tp_hanging,
                                        ..Default::default()
                                    },
                                );
                                let (tb_fs, tb_lhr, tb_ar) =
                                    tallest_run_metrics(&hf_runs, ctx.fonts);
                                let tb_line_h = resolve_line_h(tp_ls, tb_fs, tb_lhr);
                                if tb_lines.is_empty() {
                                    tb_cursor -= tp.space_before + tb_line_h + tp.space_after;
                                    continue;
                                }
                                let tb_ascender = tb_ar.unwrap_or(0.75);
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
                                    tp.indent_left,
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
                                tb_cursor -= tp.space_before
                                    + (tb_lines.len() as f32) * tb_line_h
                                    + tp.space_after;
                            }
                            if needs_clip {
                                content.restore_state();
                            }
                            if rotated {
                                content.restore_state();
                            }
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
                    draw_connectors(content, true);
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
                        + super::mark_position_stretch(para, effective_ls, false, ctx.fonts)
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
                    draw_connectors(content, true);
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
                // A wrapping float the next paragraph raises above its own top
                // wraps this paragraph too, as the body's look-ahead does:
                // bulgarian_farmland's emblem, anchored in the letterhead's
                // second line, sits beside the first. ponytail: the next
                // paragraph's top assumes this one keeps its full-width lines.
                if let Some(Block::Paragraph(next)) = hf.blocks.get(bi + 1) {
                    let raised: Vec<_> = next
                        .floating_images
                        .iter()
                        .filter(|fi| {
                            fi.wrap_type.wraps_beside()
                                && fi.v_relative_from == VRelativeFrom::Paragraph
                                && matches!(fi.v_position, VerticalPosition::Offset(o) if o < 0.0)
                        })
                        .collect();
                    if !raised.is_empty() {
                        let full_lines = build_lines(
                            &substituted_runs,
                            ctx,
                            (text_width - para.indent_left - para.indent_right).max(1.0),
                            ctx.cjk(true, para),
                            &LineOpts {
                                tab_stops: &para.tab_stops,
                                ..Default::default()
                            },
                        );
                        let next_top = cursor_y
                            - full_lines.len() as f32 * line_h
                            - hf_paragraph_gap(Some(para), para.space_after, next);
                        for fi in raised {
                            let fi_x = super::resolve_h_position(
                                fi.h_relative_from,
                                &fi.h_position,
                                fi.image.display_width,
                                sp,
                                sp.margin_left,
                                text_width,
                                text_width,
                            );
                            let fi_y_top = super::resolve_fi_y_top(fi, sp, next_top);
                            zones.push(HfFloatZone::for_float(fi, fi_x, fi_y_top));
                        }
                    }
                }

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

                // Tab stops are measured from the column, so a paragraph that
                // starts beside a float (its lines shifted right of a logo)
                // sees them that much nearer: bulgarian_farmland's letterhead
                // tabs to 63.8pt from the margin, not from the emblem's edge.
                // ponytail: default-interval stops are not shifted.
                let tab_shift = para_text_x - (sp.margin_left + para.indent_left);
                let shifted_stops: Vec<crate::model::TabStop> = para
                    .tab_stops
                    .iter()
                    .map(|t| crate::model::TabStop {
                        position: t.position - tab_shift,
                        ..t.clone()
                    })
                    .collect();

                let mut lines = build_lines(
                    &substituted_runs,
                    ctx,
                    para_text_width,
                    ctx.cjk(true, para),
                    &LineOpts {
                        inline_images: Some(&block_inline_images),
                        tab_stops: &shifted_stops,
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
                    size_lines_by_own_runs(&mut lines, ctx.fonts, effective_ls, line_h, metrics.0, para.text_alignment == crate::model::TextAlignment::Top);
                }
                draw_para_shading(content, cursor_y - lines_height(&lines, line_h, metrics));
                render_paragraph_lines(
                    content,
                    &lines,
                    &para.alignment,
                    para_text_x,
                    para.indent_left,
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
                let mut para_h = lines_height(&lines, line_h, metrics) - trailing_descent;
                if is_header {
                    if let Some(canvas_h) = inline_canvas_tab_height(para, text_width, line_h) {
                        para_h = para_h.max(canvas_h);
                    }
                }
                draw_connectors(content, true);
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
    let after = if prev.is_some_and(|p| super::helpers::drops_space_after(p, Some(para))) {
        0.0
    } else {
        prev_space_after
    };
    after.max(super::helpers::effective_space_before(para, prev))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::StyleRefForm;

    #[test]
    fn if_over_stylerefs_takes_the_page_value() {
        let sr = |form| {
            IfPart::Field(FieldCode::StyleRef {
                name: "CharPartNo".into(),
                form,
            })
        };
        let text = |t: &str| IfPart::Text(t.into());
        // IF {STYLEREF CharPartNo \n} = 0 "{STYLEREF CharPartNo}" "Part {STYLEREF CharPartNo \n}"
        let parts = [
            text(" IF "),
            sr(StyleRefForm::Label),
            text(" = 0 \""),
            sr(StyleRefForm::Text),
            text("\" \"Part "),
            sr(StyleRefForm::Label),
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
        let key = |form| styleref_key("charpartno", form);
        styleref_values.insert(key(StyleRefForm::Text), "Part 3".to_string());
        styleref_values.insert(key(StyleRefForm::Label), "0".to_string());
        assert_eq!(eval(&styleref_values).as_deref(), Some("Part 3"));
        styleref_values.insert(key(StyleRefForm::Label), "4".to_string());
        assert_eq!(eval(&styleref_values).as_deref(), Some("Part 4"));
        assert_eq!(field_args("a<>\"b c\" 2"), ["a", "<>", "b c", "2"]);
    }
}
