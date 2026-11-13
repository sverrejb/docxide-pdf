use std::collections::HashMap;

use pdf_writer::Content;

use crate::model::{
    Alignment, Block, Document, FrameProperties, HRelativeFrom, HeaderFooter, LineSpacing,
    Paragraph, Run, SectionProperties, TextAnchor, VRelativeFrom, VerticalPosition, WrapType,
};

use super::color::stroke_border;
use super::helpers::{align_offset, draw_horizontal_rule};
pub(super) use super::hf_fields::{page_field_text, substitute_hf_runs, substitute_hf_runs_with};
use super::layout::{
    LineOpts, LinkAnnotation, TextLine, build_lines, is_text_empty, line_h_with_picture,
    lines_height, picture_line_bottom, position_stretch, render_paragraph_lines, runs_max_image_h,
    size_lines_by_own_runs, tallest_run_metrics,
};
use super::positioning::resolve_h_position;
use super::table;
use super::table_layout::HfSubstitution;
use super::{RenderContext, resolve_line_h};

pub(super) use crate::model::styleref_key;

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
    if glyphs.1.is_some() {
        return glyphs;
    }
    let mark_sizes = !runs.is_empty()
        && (runs.iter().all(|r| r.vanish)
            || !runs
                .iter()
                .any(|r| r.is_line_break || r.inline_image.is_some()));
    if mark_sizes && let Some(mark) = super::layout::mark_metrics(para, fonts) {
        return mark;
    }
    super::layout::tallest_run_metrics(runs, fonts)
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
    let (top_band, bottom_band) = super::table_layout::outer_half_bands(table);
    let height = top_band + table::compute_hf_table_height(table, ctx, text_width) + bottom_band
        - trailing_after;
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

/// Whether a header or footer paragraph shares its border group with the
/// paragraph before and after it, as in the body (`joins_border_group`).
#[derive(Clone, Copy, Default)]
struct BorderGroup {
    prev: bool,
    next: bool,
    /// The band a top border adds above the lines, as in the body
    /// (`bdr_top_pad`): carbon_farming's footer rule sits 1.75pt above its
    /// text. Inside a border group only the first paragraph has one.
    top: f32,
    /// `helpers::bottom_border`'s band below the lines (the body path's
    /// `bdr_bottom_extent`); the next paragraph starts below it.
    bottom: (f32, f32),
    /// The part of the previous paragraph's band below its `between` border,
    /// when `prev`; it takes this paragraph's shading.
    below_prev: f32,
    /// Whether the gap above shares the previous paragraph's shading too.
    shade_gap: bool,
}

impl BorderGroup {
    fn of(blocks: &[Block], i: usize) -> Self {
        let flow_para = |j: usize| match blocks.get(j) {
            Some(Block::Paragraph(p)) if p.frame_props.is_none() => Some(p),
            _ => None,
        };
        let Some(para) = flow_para(i) else {
            return Self::default();
        };
        let next = flow_para(i + 1).filter(|n| super::helpers::joins_border_group(para, n));
        let prev = i
            .checked_sub(1)
            .and_then(flow_para)
            .filter(|p| super::helpers::joins_border_group(p, para));
        Self {
            prev: prev.is_some(),
            next: next.is_some(),
            top: if prev.is_some() {
                0.0
            } else {
                super::helpers::border_band(para.borders.top.as_ref())
            },
            bottom: super::helpers::bottom_border(para, next).1,
            below_prev: prev.map_or(0.0, super::helpers::between_space),
            shade_gap: prev.is_some_and(|p| p.shading == para.shading),
        }
    }
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
/// not retained again below the final line unless a picture is underlined.
/// A differently sized paragraph mark requires its own line metrics; keep that
/// existing path until it is resolved.
fn wrapped_picture_trailing_descent(
    para: &Paragraph,
    font_size: f32,
    descent: f32,
    spacing: LineSpacing,
) -> f32 {
    if matches!(spacing, LineSpacing::Exact(_))
        || para.runs.iter().any(|r| r.is_line_break)
        || para
            .runs
            .iter()
            .any(|r| !r.vanish && r.underlined_picture())
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
        size_lines_by_own_runs(
            &mut lines,
            ctx.fonts,
            spacing,
            line_h,
            metrics.0,
            1.0,
            para.text_alignment == crate::model::TextAlignment::Top,
        );
        super::size_tab_mark_line(&mut lines, para, spacing, lhr, ctx.fonts);
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

/// A header/footer (by address, it lives as long as the document) and the
/// section geometry its measured height depends on.
pub(super) type HfHeightKey = (usize, [u32; 8], bool, bool);

/// The header's height, cached per section geometry; without the floats, the
/// extent of the header's own text flow.
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
    // The header's wrapping floats so far, placed as the render places them.
    let mut float_zones: Vec<HfFloatZone> = Vec::new();
    for (bi, block) in hf.blocks.iter().enumerate() {
        match block {
            Block::Paragraph(para) if para.frame_props.is_some() => {
                // Frame paragraphs are out-of-flow; skip height contribution
            }
            Block::Paragraph(para) => {
                let gap = hf_paragraph_gap(prev_para, prev_space_after, para);
                let own_before = hf_own_space_before(prev_para, prev_space_after, para);
                let group = BorderGroup::of(&hf.blocks, bi);
                prev_para = Some(para);
                let (font_size, tallest_lhr, tallest_ar) =
                    hf_run_metrics(&para.runs, para, ctx.fonts);
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
                let count_lines = |geometry: Option<&((f32, f32), Vec<(f32, f32)>)>| {
                    let per_line_widths: Option<Vec<f32>> =
                        geometry.map(|(_, g)| g.iter().map(|&(_, w)| w).collect());
                    let (shift, width) = geometry.map_or(
                        (
                            0.0,
                            (text_width - para.indent_left - para.indent_right).max(1.0),
                        ),
                        |&((x, w), _)| (x - (sp.margin_left + para.indent_left), w),
                    );
                    build_lines(
                        &para.runs,
                        ctx,
                        width,
                        ctx.cjk(true, para),
                        &LineOpts {
                            tab_stops: &shifted_tab_stops(para, shift),
                            indent_left: para.indent_left,
                            indent_right: para.indent_right,
                            hanging: para.indent_hanging - para.indent_first_line,
                            per_line_widths: per_line_widths.as_deref(),
                            ..Default::default()
                        },
                    )
                    .len()
                    .max(1)
                };
                // Positioned runs and marks stretch their lines as in the render.
                let stretch = if is_text_empty(&para.runs) {
                    super::mark_position_stretch(para, effective_ls, false, ctx.fonts)
                } else {
                    position_stretch(&para.runs, effective_ls)
                };
                // Header lines flow beside the wrapping floats their boxes
                // overlap, as rendered, and Word reserves all of them.
                // ponytail: a footer's floats are placed from where it ends up,
                // which its height decides; its lines keep the full width here.
                let geometry = if is_header {
                    let slot_top = sp.page_height - sp.header_margin - height;
                    let top = slot_top - group.top - float_text_push(para, text_width);
                    let baseline = top
                        - super::layout::boxed_line_ascent(
                            effective_ls,
                            line_h,
                            font_size,
                            tallest_lhr,
                            tallest_ar,
                            &para.runs,
                            ctx.fonts,
                        )
                        .unwrap_or(font_size * tallest_ar.unwrap_or(0.75));
                    float_zones.extend(
                        para.floating_images
                            .iter()
                            .filter(|fi| fi.wrap_type.wraps_beside())
                            .map(|fi| hf_float_zone(fi, sp, slot_top, own_before)),
                    );
                    let zones = hf_float_zones(
                        hf,
                        bi,
                        para,
                        &para.runs,
                        &float_zones,
                        top,
                        line_h,
                        sp,
                        ctx,
                    );
                    // ponytail: Word fills both sides of a float with room
                    // on each (ut_koer's centred logo between two
                    // columns), the render only the wider side; such
                    // lines are measured at full width.
                    hf_line_geometry(&zones, para, sp, (top, baseline), line_h).filter(|g| {
                        let bottom = top - count_lines(Some(g)) as f32 * (line_h + stretch);
                        zones
                            .iter()
                            .filter(|z| top > z.bottom && bottom < z.top)
                            .all(|z| {
                                let (left, right) = z.side_room(sp);
                                left < MIN_WRAP_ROOM || right < MIN_WRAP_ROOM
                            })
                    })
                } else {
                    None
                };
                let line_h = line_h + stretch;
                // Mirrors the render loop's advance: a picture line is the
                // picture plus the text descent (`inline_line_advance`).
                let picture_h = runs_max_image_h(&para.runs);
                let wrapped_h = wrapped_header_picture_height(para, ctx, text_width);
                // A text paragraph is as many lines as it wraps to (Word bottom-
                // aligns a nine-line footer at the footer margin; one line per
                // paragraph put it 100pt low, off the page); its w:br breaks
                // are among those lines.
                let text_lines = (wrapped_h.is_none() && picture_h <= 0.0)
                    .then(|| count_lines(geometry.as_ref()));
                let mut content_h = if let Some(h) = wrapped_h {
                    h
                } else if picture_h > 0.0 {
                    line_h_with_picture(
                        picture_h,
                        &para.runs,
                        para,
                        ctx.fonts,
                        effective_ls,
                        line_h,
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
                                VRelativeFrom::Page | VRelativeFrom::TopMargin => o,
                                VRelativeFrom::BottomMargin => {
                                    sp.page_height - sp.margin_bottom + o
                                }
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
                            VRelativeFrom::Page | VRelativeFrom::TopMargin if is_header => {
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

                height += group.top + content_h + group.bottom.1;
                prev_space_after = para.space_after;
            }
            // A floating table takes no room in the header's flow
            Block::Table(table) if table.position.is_some() => {}
            Block::Table(table) => {
                let content_w = sp.text_width();
                let (top_band, bottom_band) = if is_header {
                    super::table_layout::outer_half_bands(table)
                } else {
                    (0.0, 0.0)
                };
                height +=
                    top_band + table::compute_hf_table_height(table, ctx, content_w) + bottom_band;
                prev_space_after = 0.0;
                prev_para = None;
            }
        }
    }
    (height + prev_space_after).max(placed_extent)
}

/// The body top of a section's page; `page_num`, the number the page shows,
/// picks the even header where the document has them.
pub(super) fn effective_slot_top(
    sp: &SectionProperties,
    is_first: bool,
    page_num: usize,
    ctx: &RenderContext,
) -> f32 {
    let header = layout_hf(sp, is_first, page_num, true, ctx);
    let base = sp.page_height - sp.margin_top;
    match header {
        Some(_) if sp.margin_top_fixed => base,
        Some(hf) => {
            base.min(sp.page_height - sp.header_margin - header_height(hf, ctx, sp, true, true))
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
    page_num: usize,
    ctx: &RenderContext,
) -> f32 {
    let base = sp.page_height - sp.margin_top;
    match layout_hf(sp, is_first, page_num, true, ctx) {
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
    page_num: usize,
    ctx: &RenderContext,
    left: f32,
    right: f32,
) -> Option<f32> {
    let hf = layout_hf(sp, is_first, page_num, true, ctx)?;
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
            let overlaps =
                fi_x - fi.dist_left < right && fi_x + fi.image.display_width + fi.dist_right > left;
            overlaps.then(|| {
                super::resolve_fi_y_top(fi, sp, slot_top) - fi.image.display_height - fi.dist_bottom
            })
        })
        .min_by(f32::total_cmp)
}

pub(super) fn compute_effective_margin_bottom(
    sp: &SectionProperties,
    is_first: bool,
    page_num: usize,
    ctx: &RenderContext,
) -> f32 {
    let footer = layout_hf(sp, is_first, page_num, false, ctx);
    let base = sp.margin_bottom;
    let mut bottom = match footer {
        Some(_) if sp.margin_bottom_fixed => base,
        Some(hf) => base.max(sp.footer_margin + header_height(hf, ctx, sp, false, true)),
        None => base,
    };
    let cols = super::column_geometry(sp);
    let bands: Vec<_> = hf_table_bands(sp, is_first, page_num, ctx)
        .into_iter()
        .filter(|b| cols.iter().all(|&col| b.in_column(col)))
        .collect();
    while let Some(top) = bands
        .iter()
        .filter(|b| b.top < sp.page_height - bottom && b.bottom > sp.page_height - bottom)
        .map(|b| b.top)
        .reduce(f32::min)
    {
        bottom = sp.page_height - top;
    }
    bottom
}

/// The bands, down from the page top, of the page header's and footer's
/// floating tables placed against the page or margin, text distances
/// included, holding in the columns they leave no strip beside them in
/// (`Band::in_column`). Body lines keep out of them as out of a body table's:
/// a line reaching into one goes below it, or on when the band reaches the
/// bottom margin (Word examples, fixture hf-floating-table-band: a footer
/// table 704.25pt down with its 28.35pt topFromText ends the page after a
/// line ending at 692.7pt; without topFromText the lines run to 728.7pt).
/// Each of a page's lines keeps out, judged by the widths the table is drawn
/// with; a table over one of two columns moves only that column's lines
/// (fixture hf-floating-table-band-flow). A band reaching the bottom margin
/// in every column ends the page above it, and so does one crossing that
/// raised foot (`compute_effective_margin_bottom`; Word example: footer bands
/// 550-650 and 600-750pt end the page at 550pt).
// ponytail: tables anchored to the header's text, and tables with a strip
// beside the text column, which Word wraps text around, take no part.
pub(super) fn hf_table_bands(
    sp: &SectionProperties,
    is_first: bool,
    page_num: usize,
    ctx: &RenderContext,
) -> Vec<super::Band> {
    let col = (sp.margin_left, sp.text_width());
    [true, false]
        .into_iter()
        .filter_map(|is_header| layout_hf(sp, is_first, page_num, is_header, ctx))
        .flat_map(|hf| &hf.blocks)
        .filter_map(|block| match block {
            Block::Table(t) => Some(t),
            Block::Paragraph(_) => None,
        })
        .filter_map(|t| {
            let pos = super::page_anchored(t)?;
            let fp = super::FloatingTablePos::resolve(t, pos, sp, col.0, col.1, 0.0, ctx);
            let widths = table::hf_col_widths(t, ctx, col.1);
            let width: f32 = widths.iter().sum();
            let top = sp.page_height - fp.y;
            let (top_band, bottom_band) = super::table_layout::outer_half_bands(t);
            let height = top_band + table::hf_rows_height(t, &widths, ctx) + bottom_band;
            Some(super::Band {
                x: (fp.x - fp.left_from_text, fp.x + width + fp.right_from_text),
                beside: true,
                ..super::Band::new(top - fp.top_from_text, top + height + fp.bottom_from_text)
            })
        })
        .collect()
}

/// The top of the footer's line boxes, from the page bottom (the footer
/// margin when the page has no footer): where Word stops the lines a
/// page-wide float pushes past the bottom margin.
pub(super) fn footer_top(
    sp: &SectionProperties,
    is_first: bool,
    page_num: usize,
    ctx: &RenderContext,
) -> f32 {
    match layout_hf(sp, is_first, page_num, false, ctx) {
        Some(hf) => sp.footer_margin + header_height(hf, ctx, sp, false, true),
        None => sp.footer_margin,
    }
}

/// The header (or footer) whose extent a section's page lays out around: its
/// own, else the one it inherits, as drawn (radiographer's later sections
/// inherit a two-line empty header that starts their body 27pt below the
/// header), and on a page showing an even number the even header, as drawn
/// (Word example: numbering from 2, the first sheet keeps body lines out of a
/// table only its even header holds).
fn layout_hf<'a>(
    sp: &'a SectionProperties,
    is_first: bool,
    page_num: usize,
    is_header: bool,
    ctx: &RenderContext<'a>,
) -> Option<&'a HeaderFooter> {
    let variant = hf_variant(ctx.even_and_odd_headers, sp, is_first, page_num);
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
        VRelativeFrom::Margin => (sp.page_height - sp.margin_top, sp.margin_bottom),
        VRelativeFrom::TopMargin => (sp.page_height, sp.page_height - sp.margin_top),
        VRelativeFrom::BottomMargin => (sp.margin_bottom, 0.0),
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

const MIN_WRAP_ROOM: f32 = 36.0;

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

    /// The room text has on its left and right in the margins' column.
    fn side_room(&self, sp: &SectionProperties) -> (f32, f32) {
        (
            (self.left - self.dist_left) - sp.margin_left,
            sp.margin_left + sp.text_width() - (self.right + self.dist_right),
        )
    }
}

fn shifted_tab_stops(para: &Paragraph, shift: f32) -> Vec<crate::model::TabStop> {
    para.tab_stops
        .iter()
        .map(|t| crate::model::TabStop {
            position: t.position - shift,
            ..t.clone()
        })
        .collect()
}

/// The wrapping floats a header (footer) paragraph's lines flow beside: its
/// own and those of the paragraphs before it (`earlier`), and those the next
/// paragraph raises above its own top. `cursor_y` is the top of the
/// paragraph's first line.
fn hf_float_zones(
    hf: &HeaderFooter,
    bi: usize,
    para: &Paragraph,
    runs: &[Run],
    earlier: &[HfFloatZone],
    cursor_y: f32,
    line_h: f32,
    sp: &SectionProperties,
    ctx: &RenderContext,
) -> Vec<HfFloatZone> {
    // Word probes on ut_koer's letterhead logo: in compat mode 14
    // and older, header and footer text runs straight through a
    // float (square or tight, in front or behind, at any offset);
    // compat 15 wraps it.
    if ctx.compat_mode < 15 {
        return Vec::new();
    }
    let text_width = sp.text_width();
    let mut zones = earlier.to_vec();
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
                runs,
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
            let next_own_before = hf_own_space_before(Some(para), para.space_after, next);
            zones.extend(
                raised
                    .into_iter()
                    .map(|fi| hf_float_zone(fi, sp, next_top, next_own_before)),
            );
        }
    }
    zones
}

/// Where a header (footer) paragraph's drawings hang from: its `slot_top`
/// above its `own_before`, or the line top when on the line.
fn hf_anchor_top(slot_top: f32, own_before: f32, on_line: bool) -> f32 {
    slot_top + if on_line { 0.0 } else { own_before }
}

fn hf_float_zone(
    fi: &crate::model::FloatingImage,
    sp: &SectionProperties,
    slot_top: f32,
    own_before: f32,
) -> HfFloatZone {
    let text_width = sp.text_width();
    let fi_x = super::resolve_h_position(
        fi.h_relative_from,
        &fi.h_position,
        fi.image.display_width,
        sp,
        sp.margin_left,
        text_width,
        text_width,
    );
    let top = hf_anchor_top(slot_top, own_before, fi.on_line);
    HfFloatZone::for_float(fi, fi_x, super::resolve_fi_y_top(fi, sp, top))
}

/// Where a header (footer) paragraph's lines go when a float in `zones`
/// overlaps its first line (`top` down to its `baseline`): the first line's
/// (x, width) and each line's, a line flowing beside every float its box
/// overlaps (Word example: a picture ending 4pt into a header's first line
/// narrows it). Paragraph indents count from the column edge and the float
/// bounds clip them, they do not add up.
fn hf_line_geometry(
    zones: &[HfFloatZone],
    para: &Paragraph,
    sp: &SectionProperties,
    (top, baseline): (f32, f32),
    line_h: f32,
) -> Option<((f32, f32), Vec<(f32, f32)>)> {
    let col_x = sp.margin_left;
    let text_width = sp.text_width();
    let col_right = sp.margin_left + text_width;
    // Combined text bounds over [y_lo, y_hi]: a float with more room on its
    // right raises the left bound, otherwise it lowers the right bound.
    let bounds_at = |y_hi: f32, y_lo: f32| -> Option<(f32, f32)> {
        let mut left = col_x;
        let mut right = col_right;
        let mut hit = false;
        for z in zones {
            if y_hi > z.bottom && y_lo < z.top {
                hit = true;
                let (space_left, space_right) = z.side_room(sp);
                if space_right >= space_left && space_right >= MIN_WRAP_ROOM {
                    left = left.max(z.right + z.dist_right);
                } else if space_left >= MIN_WRAP_ROOM {
                    right = right.min(z.left - z.dist_left);
                }
            }
        }
        hit.then_some((left, right))
    };
    let indented = |(lb, rb): (f32, f32)| -> (f32, f32) {
        let tx = (col_x + para.indent_left).max(lb);
        let tr = (col_right - para.indent_right).min(rb);
        (tx, (tr - tx).max(1.0))
    };
    let first = indented(bounds_at(top, baseline)?);
    let full_w = (text_width - para.indent_left - para.indent_right).max(1.0);
    let deepest_bottom = zones.iter().map(|z| z.bottom).fold(f32::INFINITY, f32::min);
    let max_lines = (((top - deepest_bottom) / line_h).ceil() as usize + 10).max(20);
    let geom = (0..max_lines)
        .map(|i| {
            let line_top = top - i as f32 * line_h;
            bounds_at(line_top, line_top - line_h)
                .map_or((col_x + para.indent_left, full_w), indented)
        })
        .collect();
    Some((first, geom))
}

/// Per-page context for header/footer rendering: field substitution values
/// and image name mappings resolved by the caller.
pub(super) struct HfPageContext<'a> {
    pub(super) page_num: usize,
    pub(super) total_pages: usize,
    pub(super) section_pages: usize,
    pub(super) para_image_names: &'a HashMap<usize, String>,
    pub(super) inline_image_names: &'a HashMap<(usize, usize), String>,
    pub(super) floating_image_names: &'a HashMap<(usize, usize), String>,
    pub(super) effect_para_names: &'a HashMap<usize, super::images::EffectXObjs>,
    pub(super) styleref_values: &'a HashMap<String, String>,
    pub(super) page_num_format: Option<&'a str>,
}

/// A header text box's paragraph laid out at the box's content width: the
/// anchor-offset pass and the draw pass advance by the same height.
struct HfBoxPara {
    lines: Vec<TextLine>,
    inline_imgs: HashMap<usize, String>,
    ls: LineSpacing,
    text_w: f32,
    fs: f32,
    ar: Option<f32>,
    line_h: f32,
}

impl HfBoxPara {
    fn height(&self, tp: &Paragraph) -> f32 {
        tp.space_before + self.lines.len().max(1) as f32 * self.line_h + tp.space_after
    }
}

/// The text width inside a header text box: its content width less the
/// paragraph's indents.
fn hf_box_text_w(tp: &Paragraph, content_w: f32) -> f32 {
    (content_w - tp.indent_left - tp.indent_right).max(1.0)
}

fn hf_box_paragraph(
    tp: &Paragraph,
    content_w: f32,
    sub: &HfSubstitution,
    ctx: &RenderContext,
) -> HfBoxPara {
    let ls = tp.line_spacing.unwrap_or(ctx.doc_line_spacing);
    let text_w = hf_box_text_w(tp, content_w);
    let hanging = if !tp.list_label.is_empty() {
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
    let inline_imgs = super::layout::run_image_map(&tp.runs, ctx.textbox_image_names);
    // Footer page numbers live in text boxes as often as in plain paragraphs;
    // without this the box kept the template's cached PAGE/NUMPAGES text.
    let runs = substitute_hf_runs_with(&tp.runs, sub);
    let lines = build_lines(
        &runs,
        ctx,
        text_w,
        ctx.cjk(true, tp),
        &LineOpts {
            inline_images: Some(&inline_imgs),
            tab_stops: &tp.tab_stops,
            indent_left: tp.indent_left,
            indent_right: tp.indent_right,
            hanging,
            ..Default::default()
        },
    );
    let (fs, lhr, ar) = tallest_run_metrics(&runs, ctx.fonts);
    let line_h = resolve_line_h(ls, fs, lhr);
    HfBoxPara {
        lines,
        inline_imgs,
        ls,
        text_w,
        fs,
        ar,
        line_h,
    }
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
    let sub = HfSubstitution {
        page_num: pc.page_num,
        total_pages: pc.total_pages,
        section_pages: pc.section_pages,
        styleref_values: pc.styleref_values,
        page_num_format: pc.page_num_format,
    };
    let text_width = sp.text_width();
    let mut cursor_y = if is_header {
        sp.page_height - sp.header_margin
    } else {
        sp.footer_margin + header_height(hf, ctx, sp, false, true)
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
                    sub.page_num,
                    sub.total_pages,
                    sub.section_pages,
                    sub.styleref_values,
                    sub.page_num_format,
                    gradient_specs,
                    links,
                    is_header,
                );
                prev_space_after = 0.0;
                prev_para = None;
            }
            Block::Paragraph(para) if para.frame_props.is_some() => {
                render_hf_frame_paragraph(
                    content,
                    para,
                    ctx,
                    sp,
                    &sub,
                    cursor_y,
                    gradient_specs,
                    links,
                );
                // Frame paragraphs are out-of-flow: do not advance cursor_y
                pi += 1;
            }
            Block::Paragraph(para) => {
                let has_para_image = para.image.is_some();
                let has_field_code = para.runs.iter().any(|r| r.field_code.is_some());
                let text_empty = !has_field_code && is_text_empty(&para.runs);

                let gap = hf_paragraph_gap(prev_para, prev_space_after, para);
                let own_before = hf_own_space_before(prev_para, prev_space_after, para);
                prev_para = Some(para);

                let substituted_runs = substitute_hf_runs_with(&para.runs, &sub);

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
                // Mirrors header_height: a line overlapping a frame the
                // text may not sit beside moves below it.
                cursor_y = sp.page_height
                    - below_blocking_frames(sp.page_height - cursor_y, line_h, &bands);
                let slot_top = cursor_y;
                let group = BorderGroup::of(&hf.blocks, bi);
                // Drawings hang from above the paragraph's own space before,
                // those on the line from the line top.
                let anchor_top = |on_line: bool| hf_anchor_top(slot_top, own_before, on_line);
                cursor_y -= group.top + float_text_push(para, text_width);

                // Foreground anchored lines must cover an inline picture's white
                // background, so they wait for the late pass; behind-document
                // lines (and all lines when there is no picture) paint early.
                let defer_connectors = has_para_image
                    || para.runs.iter().any(|run| run.inline_image.is_some())
                    || para.textboxes.iter().any(|tb| {
                        (tb.fill.is_some() || tb.picture_fill.is_some()) && tb.behind_doc
                    });
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
                            anchor_top(false),
                        );
                        super::positioning::render_connector(connector, content, x, y);
                    }
                };
                draw_connectors(content, false);

                // Paragraph borders span the laid-out height, so each exit below
                // draws them once it knows it (ut_koer: a header staff image
                // with a bottom border); empty bordered paragraphs render too.
                let bdr = &para.borders;
                let (box_left_in, box_right_in) = super::helpers::border_box_indents(para);
                let (box_left, box_right, box_top) = (
                    sp.margin_left + box_left_in,
                    sp.margin_left + text_width - box_right_in,
                    slot_top,
                );
                let on_grid = |y, band| super::layout::rule_on_grid(y, sp.page_height, band);
                let draw_para_borders = |content: &mut Content, box_bottom: f32| {
                    // Side borders as in the body: their inner edge `space`
                    // plus Word's 1.47 / 1.73pt outside the text, and the
                    // horizontal borders run out to their outer edges.
                    let (left_x, right_x) = super::side_border_outsets(bdr, true);
                    let (l_ext, r_ext) = super::horizontal_border_ext(bdr);
                    let h_left = box_left - left_x - l_ext;
                    let h_right = box_right + right_x + r_ext;
                    if !group.prev
                        && let Some(b) = &bdr.top
                    {
                        let y = on_grid(box_top - b.extent() / 2.0, b.extent());
                        stroke_border(content, b, (h_left, y), (h_right, y), -1.0);
                    }
                    // The stroke sits `space` below the box, outside it,
                    // as the body path's bottom pad places it.
                    if let Some(b) = if group.next {
                        &bdr.between
                    } else {
                        &bdr.bottom
                    } {
                        let y = on_grid(box_bottom - group.bottom.0 - b.extent() / 2.0, b.extent());
                        stroke_border(content, b, (h_left, y), (h_right, y), 1.0);
                    }
                    let v_top = on_grid(box_top + if group.prev { gap } else { 0.0 }, 0.0);
                    let v_bottom = on_grid(
                        box_bottom
                            - if group.next {
                                group.bottom.1
                            } else {
                                group.bottom.0
                            },
                        0.0,
                    );
                    for (b, x, inward) in [
                        (&bdr.left, box_left - left_x, 1.0),
                        (&bdr.right, box_right + right_x, -1.0),
                    ] {
                        if let Some(b) = b {
                            stroke_border(content, b, (x, v_top), (x, v_bottom), inward);
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
                    let bottom = box_bottom - group.bottom.0;
                    let top = box_top + group.below_prev + if group.shade_gap { gap } else { 0.0 };
                    content.save_state();
                    super::color::fill_rgb(content, color);
                    content.rect(left, bottom, right - left, top - bottom);
                    content.fill_nonzero();
                    content.restore_state();
                };

                let line_ascent = super::layout::boxed_line_ascent(
                    effective_ls,
                    line_h,
                    font_size,
                    tallest_lhr,
                    tallest_ar,
                    &substituted_runs,
                    ctx.fonts,
                )
                .unwrap_or(font_size * ascender_ratio);
                let mut baseline_y = cursor_y - line_ascent;

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
                            if let Some(pdf_name) = pc.floating_image_names.get(&(pi, fi_idx)) {
                                let img = &fi.image;
                                let zone = hf_float_zone(fi, sp, slot_top, own_before);
                                super::smartart::render_image_with_clip(
                                    content,
                                    pdf_name,
                                    zone.left,
                                    zone.bottom,
                                    img.display_width,
                                    img.display_height,
                                    img.clip_geometry.as_ref(),
                                );
                                if let Some(sc) = img.stroke_color {
                                    super::smartart::stroke_image_border(
                                        content,
                                        zone.left,
                                        zone.bottom,
                                        img.display_width,
                                        img.display_height,
                                        sc,
                                        img.stroke_width,
                                        img.clip_geometry.as_ref(),
                                    );
                                }
                                // Register float zone for wrapping images
                                if fi.wrap_type.wraps_beside() {
                                    hdr_fz.push(zone);
                                }
                            }
                        }
                    } else {
                        // Render textboxes
                        for tb in &para.textboxes {
                            render_hf_textbox(
                                content,
                                tb,
                                ctx,
                                sp,
                                anchor_top(tb.on_line),
                                &sub,
                                gradient_specs,
                                links,
                            );
                        }
                    }
                }

                if (has_para_image || text_empty) && para.content_height > 0.0 {
                    if let Some(pdf_name) = pc.para_image_names.get(&pi) {
                        let img = para.image.as_ref().unwrap();
                        let natural_line_h = font_size * tallest_lhr.unwrap_or(1.2);
                        let bottom_depth = if para.content_height > line_h {
                            img.layout_extra_top + img.display_height
                        } else {
                            let foot = match effective_ls {
                                LineSpacing::AtLeast(_) => line_h,
                                _ => para.content_height.max(natural_line_h),
                            };
                            foot - (img.layout_extra_height - img.layout_extra_top)
                        };
                        let y_bottom = cursor_y - bottom_depth;
                        let x = sp.margin_left
                            + align_offset(para.alignment, (text_width - img.box_width()).max(0.0));
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
                    cursor_y -= line_h + group.bottom.1;
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
                        + group.bottom.1;
                    // TopAndBottom textboxes push content below them
                    for tb in &para.textboxes {
                        if matches!(tb.wrap_type, WrapType::TopAndBottom) {
                            let tb_bottom_y = match tb.v_relative_from {
                                VRelativeFrom::Page | VRelativeFrom::TopMargin => {
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

                let block_inline_images: HashMap<usize, String> = pc
                    .inline_image_names
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

                let zones = hf_float_zones(
                    hf,
                    bi,
                    para,
                    &substituted_runs,
                    &hdr_fz,
                    cursor_y,
                    line_h,
                    sp,
                    ctx,
                );
                let hdr_line_geom =
                    hf_line_geometry(&zones, para, sp, (cursor_y, baseline_y), line_h).map(
                        |(first, geom)| {
                            (para_text_x, para_text_width) = first;
                            geom
                        },
                    );

                let per_line_widths: Option<Vec<f32>> = hdr_line_geom
                    .as_ref()
                    .map(|g| g.iter().map(|&(_, w)| w).collect());

                // Tab stops are measured from the column, so a paragraph that
                // starts beside a float (its lines shifted right of a logo)
                // sees them that much nearer: bulgarian_farmland's letterhead
                // tabs to 63.8pt from the margin, not from the emblem's edge.
                // ponytail: default-interval stops are not shifted.
                let shifted_stops =
                    shifted_tab_stops(para, para_text_x - (sp.margin_left + para.indent_left));

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
                // The picture then hangs from that baseline as it is: it lowers
                // the line no further (Word examples: 10.5pt at single spacing
                // and 40pt under a 40pt minimum end at the line foot).
                let mut picture_ascent = None;
                if lines.len() == 1 && substituted_runs.iter().all(|r| r.text.trim().is_empty()) {
                    let mut images = substituted_runs
                        .iter()
                        .filter_map(|r| r.inline_image.as_ref());
                    if let Some(image) = images.next()
                        && images.next().is_none()
                        && image.display_height + image.layout_extra_height <= line_h
                    {
                        let foot = match effective_ls {
                            LineSpacing::AtLeast(_) => line_h,
                            _ => font_size * tallest_lhr.unwrap_or(1.2),
                        };
                        baseline_y =
                            cursor_y - foot + image.layout_extra_height - image.layout_extra_top;
                        picture_ascent = Some(cursor_y - baseline_y);
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
                let metrics = (
                    picture_ascent.unwrap_or(line_ascent),
                    wrapped_picture_bottom,
                );
                // Each line as tall as its own runs, as in the body.
                if !matches!(effective_ls, LineSpacing::Exact(_)) {
                    size_lines_by_own_runs(
                        &mut lines,
                        ctx.fonts,
                        effective_ls,
                        line_h,
                        line_ascent,
                        1.0,
                        para.text_alignment == crate::model::TextAlignment::Top,
                    );
                    super::size_tab_mark_line(
                        &mut lines,
                        para,
                        effective_ls,
                        tallest_lhr,
                        ctx.fonts,
                    );
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
                    None,
                    Some(super::layout::LineGrid::new(
                        sp.page_height,
                        cursor_y - baseline_y,
                        effective_ls,
                        line_h,
                        font_size * tallest_lhr.unwrap_or(1.2),
                        metrics.0,
                        para.space_after,
                    )),
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
                cursor_y -= para_h + group.bottom.1;
                prev_space_after = para.space_after;
                pi += 1;
            }
        }
    }
}

/// A framed (`w:framePr`) header or footer paragraph: drawn where its frame
/// sits, out of the flow.
fn render_hf_frame_paragraph(
    content: &mut Content,
    para: &Paragraph,
    ctx: &RenderContext,
    sp: &SectionProperties,
    sub: &HfSubstitution,
    cursor_y: f32,
    gradient_specs: &mut Vec<super::GradientSpec>,
    links: &mut Vec<LinkAnnotation>,
) {
    let text_width = sp.text_width();
    let fp = para.frame_props.as_ref().unwrap();
    let substituted_runs = substitute_hf_runs_with(&para.runs, sub);
    let (font_size, tallest_lhr, tallest_ar) = tallest_run_metrics(&substituted_runs, ctx.fonts);
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
            HRelativeFrom::Margin | HRelativeFrom::Column => (sp.margin_left, text_width),
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
    let frame_top =
        anchored_frame_top(fp, frame_h, sp).map_or(cursor_y, |top| sp.page_height - top);
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
        None,
        None,
    );
}

/// One of a header or footer paragraph's text boxes: its fill, WordArt or
/// paragraphs and tables, clipped and anchored as the body draws them.
fn render_hf_textbox(
    content: &mut Content,
    tb: &crate::model::Textbox,
    ctx: &RenderContext,
    sp: &SectionProperties,
    anchor_top: f32,
    sub: &HfSubstitution,
    gradient_specs: &mut Vec<super::GradientSpec>,
    links: &mut Vec<LinkAnnotation>,
) {
    let text_width = sp.text_width();
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
        anchor_top,
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
    super::textbox_render::render_picture_fill(
        content,
        tb,
        (tb_x, tb_y_top - tb_height, tb.width_pt, tb_height),
        ctx,
    );

    let content_x = tb_x + tb.margin_left;
    let content_w = (tb.width_pt - tb.margin_left - tb.margin_right).max(0.0);
    let hf_table = |t| super::table::TextboxTable::layout(t, content_w, ctx, Some(sub));

    // WordArt (a header watermark) draws its glyphs as warped
    // outlines, as the body path does.
    if tb
        .text_warp
        .as_ref()
        .is_some_and(|w| w.preset != "textNoShape")
        && super::wordart::render_warped_textbox(tb, content, ctx.fonts, tb_x, tb_y_top, content_w)
    {
        if rotated {
            content.restore_state();
        }
        return;
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
    // top-anchored everything. Mirror the body math here; `hf_box_paragraph`
    // gives this pass and the draw loop below the same advance.
    let anchor_offset = match tb.text_anchor {
        TextAnchor::Top => 0.0,
        TextAnchor::Middle | TextAnchor::Bottom => {
            let tp_height = |tp: &Paragraph| -> f32 {
                if let Some(img) = super::textbox_render::textbox_para_block_image(tp) {
                    return tp.space_before
                        + img.display_height
                        + img.layout_extra_height
                        + tp.space_after;
                }
                hf_box_paragraph(tp, content_w, sub, ctx).height(tp)
            };
            let tables: f32 = tb.tables.iter().map(|(_, t)| hf_table(t).height()).sum();
            let total_h: f32 = tb.paragraphs.iter().map(tp_height).sum::<f32>() + tables;
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
    for (tp_idx, tp) in tb.paragraphs.iter().enumerate() {
        for (_, table) in tb.tables.iter().filter(|(at, _)| *at == tp_idx) {
            hf_table(table).render(
                (content_x, content_w),
                &mut tb_cursor,
                content,
                ctx,
                gradient_specs,
                links,
                None,
                Some(sub),
            );
        }
        if let Some(img) = super::textbox_render::textbox_para_block_image(tp) {
            if let Some(pdf_name) = ctx.textbox_image_names.get(&img.id) {
                let img_x = content_x
                    + tp.indent_left
                    + align_offset(
                        tp.alignment,
                        (hf_box_text_w(tp, content_w) - img.display_width).max(0.0),
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
            tb_cursor -=
                tp.space_before + img.display_height + img.layout_extra_height + tp.space_after;
            continue;
        }

        let bp = hf_box_paragraph(tp, content_w, sub, ctx);
        if bp.lines.is_empty() {
            tb_cursor -= bp.height(tp);
            continue;
        }
        let tb_ascender = bp.ar.unwrap_or(0.75);
        let tb_baseline = tb_cursor - tp.space_before - bp.fs * tb_ascender;
        let tb_metrics = (
            bp.fs * tb_ascender,
            if bp.inline_imgs.is_empty() {
                0.0
            } else {
                picture_line_bottom(&tp.runs, tp, ctx.fonts, bp.ls)
            },
        );
        super::render_list_label(
            content,
            tp,
            ctx.fonts,
            content_x + tp.indent_left - tp.indent_hanging,
            tb_baseline,
            bp.fs,
        );
        render_paragraph_lines(
            content,
            &bp.lines,
            &tp.alignment,
            content_x + tp.indent_left,
            tp.indent_left,
            bp.text_w,
            tb_baseline,
            bp.line_h,
            tb_metrics,
            bp.lines.len(),
            0,
            links,
            0.0,
            ctx.fonts,
            None,
            gradient_specs,
            None,
            None,
            None,
            None,
            None,
        );
        tb_cursor -= bp.height(tp);
    }
    if needs_clip {
        content.restore_state();
    }
    if rotated {
        content.restore_state();
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

/// Resolve which header (or footer) to use for a given page, walking sections
/// backward for inheritance. Returns `(header_data, hf_type_id, section_index)`.
fn resolve_hf_for_page(
    doc: &Document,
    section_idx: usize,
    is_first_page: bool,
    page_num: usize,
    is_header: bool,
) -> (Option<&HeaderFooter>, u8, usize) {
    let sp = &doc.sections[section_idx].properties;
    let variant = hf_variant(doc.even_and_odd_headers, sp, is_first_page, page_num);
    let t = match (variant, is_header) {
        (HfVariant::Default, true) => 0,
        (HfVariant::First, true) => 1,
        (HfVariant::Even, true) => 4,
        (HfVariant::Default, false) => 2,
        (HfVariant::First, false) => 3,
        (HfVariant::Even, false) => 5,
    };
    match inherited_hf(&doc.sections, section_idx, variant, is_header) {
        Some((hf, idx)) => (Some(hf), t, idx),
        None => (None, t, section_idx),
    }
}

pub(super) fn resolve_header_for_page(
    doc: &Document,
    section_idx: usize,
    is_first_page: bool,
    page_num: usize,
) -> (Option<&HeaderFooter>, u8, usize) {
    resolve_hf_for_page(doc, section_idx, is_first_page, page_num, true)
}

pub(super) fn resolve_footer_for_page(
    doc: &Document,
    section_idx: usize,
    is_first_page: bool,
    page_num: usize,
) -> (Option<&HeaderFooter>, u8, usize) {
    resolve_hf_for_page(doc, section_idx, is_first_page, page_num, false)
}

/// The gap above a header/footer paragraph: contextual spacing drops both
/// sides between same-style paragraphs, as in the body (two 18pt-before
/// motion header lines step 15.5 in Word, not 33.6).
pub(super) fn hf_paragraph_gap(
    prev: Option<&Paragraph>,
    prev_space_after: f32,
    para: &Paragraph,
) -> f32 {
    let (after, before) = hf_spacing(prev, prev_space_after, para);
    after.max(before)
}

/// How far above its first line a header/footer paragraph's drawings hang
/// from: its own space before, the part of the gap above the previous
/// paragraph's space after (`ParaAnchor` in the body). Word examples: a
/// letterhead logo 4pt below a first paragraph with 36pt before sits at
/// header margin + 4; after a paragraph with 12pt after, 24pt before puts it
/// 12pt below the previous line.
fn hf_own_space_before(prev: Option<&Paragraph>, prev_space_after: f32, para: &Paragraph) -> f32 {
    let (after, before) = hf_spacing(prev, prev_space_after, para);
    super::own_space_before(before, after)
}

fn hf_spacing(prev: Option<&Paragraph>, prev_space_after: f32, para: &Paragraph) -> (f32, f32) {
    let after = if prev.is_some_and(|p| super::helpers::drops_space_after(p, Some(para))) {
        0.0
    } else {
        prev_space_after
    };
    (after, super::helpers::effective_space_before(para, prev))
}
