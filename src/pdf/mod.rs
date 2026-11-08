mod assembly;
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
mod list_label;
mod objstm;
mod positioning;
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
    Block, Document, FieldCode, FrameProperties, HRelativeFrom, LineNumberRestart, LineSpacing,
    PageVerticalAlign, Paragraph, ParagraphBorder, Run, SectionBreakType, SectionProperties,
    ShapeFill, ShapeGeometry, VRelativeFrom, VerticalPosition, WrapText, WrapType,
};

use crate::fonts::font_key;
use assembly::{HeadingEntry, assemble_pdf_pages};
use color::{fill_rgb, stroke_border, stroke_segment};
use fonts::collect_and_register_fonts;
use footnotes::{footnote_height, render_endnotes_inline, render_page_footnotes};
use header_footer::{
    HfPageContext, compute_effective_margin_bottom, effective_slot_top, render_header_footer,
    resolve_footer_for_page, resolve_header_for_page,
};
pub(super) use helpers::resolve_line_h;
use helpers::{
    align_offset, auto_ascent_scale, between_space, border_band, bottom_border,
    border_box_indents, collect_paras, draw_horizontal_rule, joins_border_group,
    para_runs_with_textboxes, spacing_leading,
};
use images::{EffectXObjs, EmbeddedImages, embed_all_images};
use layout::{
    CjkLayout, DualRegion, EMPTY_EFFECTS, EMPTY_INLINE_IMAGES, LineNumberArg, LineOpts,
    LinkAnnotation, LinkTagger, PUNCT_PULL_FREE, PUNCT_SQUEEZE, TextLine, build_lines,
    build_paragraph_lines, descender_ratio, grid_baseline_offset, grid_snapped_line_h,
    inline_image_line_extra, inserted_picture_descent, is_text_empty, line_h_with_picture,
    line_max_image_h, lines_height, picture_line_bottom,
    render_paragraph_lines, run_line_metrics, size_lines_by_own_runs, tallest_glyph_run_metrics,
    tallest_run_metrics,
};
use list_label::{label_font_key, render_list_label};
use positioning::{
    image_band, image_blocks_band, render_connector, render_floating_images, resolve_fi_x,
    wraps_in_column,
};
pub(super) use positioning::{resolve_fi_y_top, resolve_h_position};
use smartart::draw_shape_path;
use table::render_table;
use textbox_render::render_single_textbox;

/// The paragraph a table opens with, which contextual spacing compares a
/// paragraph before the table against: bulgarian_road_safety's empty Normal
/// paragraph keeps no space after above a table whose first cell is Normal.
fn first_cell_paragraph(t: &crate::model::Table) -> Option<&Paragraph> {
    t.first_cell()?.content.iter().find_map(|b| match b {
        Block::Paragraph(p) => Some(p),
        Block::Table(_) => None,
    })
}

/// The paragraph contextual spacing compares a neighbour of `blocks[i]` with:
/// the block itself, or a table's first cell paragraph.
fn block_para(blocks: &[Block], i: usize) -> Option<&Paragraph> {
    blocks.get(i).and_then(leading_para)
}

fn leading_para(block: &Block) -> Option<&Paragraph> {
    match block {
        Block::Paragraph(p) => Some(p),
        Block::Table(t) => first_cell_paragraph(t),
    }
}

/// The frame Word lifts `block` out of the flow into: any page- or
/// margin-anchored framePr (a table's is its first cell paragraph's), which
/// keeps text off its sides or, wrapping, lets text flow on both of them
/// (empty-frame-control), or a paragraph's wrapping frame at an explicit
/// offset (`paragraph_offset_frame`), which text then wraps around like a float.
/// ponytail: other text-anchored frames stay in the flow; lift them when a
/// fixture needs it.
fn lifted_frame(block: &Block) -> Option<&FrameProperties> {
    let fp = leading_para(block)?.frame_props.as_ref()?;
    if fp.v_relative_from == VRelativeFrom::Paragraph
        && (matches!(block, Block::Table(_)) || !paragraph_offset_frame(fp))
    {
        return None;
    }
    Some(fp)
}

/// A frame wholly beside the text area, which no text can wrap around
/// (3ec631ca50's address block in the right margin).
fn frame_beside_text(fp: &FrameProperties, sp: &SectionProperties) -> bool {
    let (text_x, text_w) = (sp.margin_left, sp.text_width());
    let x = resolve_h_position(fp.h_relative_from, &fp.h_position, fp.width, sp, text_x, text_w, text_w);
    fp.width > 0.0 && (x >= text_x + text_w || x + fp.width <= text_x)
}

/// A wrapping paragraph-anchored frame with an explicit height and a nonzero
/// offset, or a drop cap: it keeps its own coordinates, and text wraps around it.
fn paragraph_offset_frame(fp: &FrameProperties) -> bool {
    fp.drop_cap.is_some()
        || (!fp.text_below
            && fp.v_relative_from == VRelativeFrom::Paragraph
            && fp.height > 0.0
            && matches!(fp.v_position, VerticalPosition::Offset(y) if y != 0.0))
}

/// The blocks from the start of `blocks` that share the frame `props`.
fn frame_blocks<'b>(
    props: &FrameProperties,
    blocks: &'b [Block],
) -> impl Iterator<Item = &'b Block> {
    blocks
        .iter()
        .take_while(move |b| leading_para(b).and_then(|p| p.frame_props.as_ref()) == Some(props))
}

/// The room beside a float on the side(s) its wrapText lets text use.
fn side_room(wrap_text: WrapText, left: f32, right: f32) -> f32 {
    match wrap_text {
        WrapText::Left => left,
        WrapText::Right => right,
        WrapText::BothSides | WrapText::Largest => left.max(right),
    }
}

/// Whether a page- or margin-anchored textbox keeps body lines out of its
/// whole band, like a lifted frame: a topAndBottom box (issue #148: text
/// flows above a page-placed one and resumes below it), or one leaving no
/// room on its wrap side (massachusetts' secretary box at the right edge wraps
/// text on its right only, so its anchor's line and the date start below it).
fn textbox_blocks_band(
    tb: &crate::model::Textbox,
    para: &Paragraph,
    sp: &SectionProperties,
    (col_x, col_w): (f32, f32),
    text_width: f32,
) -> bool {
    if tb.v_relative_from == VRelativeFrom::Paragraph {
        return false;
    }
    match tb.wrap_type {
        WrapType::TopAndBottom => true,
        WrapType::Square | WrapType::Tight => {
            textbox_room_beside(tb, para, sp, (col_x, col_w), text_width)
                .is_some_and(|room| room < MIN_EMPTY_STRIP)
        }
        _ => false,
    }
}

fn textbox_x(
    tb: &crate::model::Textbox,
    sp: &SectionProperties,
    (col_x, col_w): (f32, f32),
    text_width: f32,
) -> f32 {
    let (rel, pos) = (tb.h_relative_from, &tb.h_position);
    resolve_h_position(rel, pos, tb.width_pt, sp, col_x, col_w, text_width)
}

/// The rectangle (left, top, width, height; y from the page bottom) a text
/// box or shape at `x`, `y_top` keeps text off: its extent turned to the
/// nearest quarter turn (Word examples: a 128×60pt box rotated 90° wraps as
/// 60×128 about its centre, one rotated 30° as its unrotated extent).
fn textbox_wrap_rect(
    tb: &crate::model::Textbox,
    x: f32,
    y_top: f32,
    h: f32,
) -> (f32, f32, f32, f32) {
    let quarter = (tb.rotation_deg.rem_euclid(180.0) / 90.0).round() as i32 % 2 == 1;
    if !quarter {
        return (x, y_top, tb.width_pt, h);
    }
    let (cx, cy) = (x + tb.width_pt / 2.0, y_top - h / 2.0);
    (cx - h / 2.0, cy + tb.width_pt / 2.0, h, tb.width_pt)
}

/// A wrapping float's horizontal span with its wrap distances and its
/// vertical extent (top, bottom; y from the page bottom, its anchor at the
/// top margin).
type WrapSpan = ((f32, f32), (f32, f32));

fn textbox_span(
    tb: &crate::model::Textbox,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
) -> WrapSpan {
    let anchor = sp.page_height - sp.margin_top;
    let x = textbox_x(tb, sp, col, text_width);
    let (rel, pos) = (tb.v_relative_from, &tb.v_position);
    let y = header_footer::resolve_tb_y_top(rel, pos, tb.height_pt, sp, anchor);
    let (left, top, w, h) = textbox_wrap_rect(tb, x, y, tb.height_pt);
    (
        (left - tb.dist_left, left + w + tb.dist_right),
        (top + tb.dist_top, top - h - tb.dist_bottom),
    )
}

fn picture_span(
    fi: &crate::model::FloatingImage,
    sp: &SectionProperties,
    (col_x, col_w): (f32, f32),
    text_width: f32,
) -> WrapSpan {
    let x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
    let y = resolve_fi_y_top(fi, sp, sp.page_height - sp.margin_top);
    (
        (x - fi.dist_left, x + fi.image.display_width + fi.dist_right),
        (y + fi.dist_top, y - fi.image.display_height - fi.dist_bottom),
    )
}

/// Whether text wraps beside a floating text box or shape (square, tight or
/// through) like beside a picture: Word needs a strip of `MIN_EMPTY_STRIP`
/// on a side its wrapText allows (examples: 20pt beside a box keeps its
/// anchor's empty line there, 17pt moves it below the box).
fn textbox_wraps_beside(
    tb: &crate::model::Textbox,
    para: &Paragraph,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
) -> bool {
    !tb.on_line
        && tb.wrap_type.wraps_beside()
        && textbox_room_beside(tb, para, sp, col, text_width)
            .is_some_and(|room| room >= MIN_EMPTY_STRIP)
}

/// The room beside a text box on the side(s) its wrapText lets text use,
/// None for a box that stays out of the column.
fn textbox_room_beside(
    tb: &crate::model::Textbox,
    para: &Paragraph,
    sp: &SectionProperties,
    (col_x, col_w): (f32, f32),
    text_width: f32,
) -> Option<f32> {
    let col = (col_x, col_w);
    let ((own_left, own_right), (top, bottom)) = textbox_span(tb, sp, col, text_width);
    if own_left >= col_x + col_w || own_right <= col_x {
        return None;
    }
    // The widest strips (left, right) the column leaves beside the box, the
    // paragraph's other wrapping floats beside it taken out too: two boxes
    // side by side across the column leave no room, though each alone would.
    let others = para
        .textboxes
        .iter()
        .filter(|o| !std::ptr::eq(*o, tb) && !o.on_line && o.wrap_type.wraps_beside())
        .map(|o| textbox_span(o, sp, col, text_width))
        .chain(
            para.floating_images
                .iter()
                .filter(|fi| fi.wrap_type.wraps_beside())
                .map(|fi| picture_span(fi, sp, col, text_width)),
        )
        .filter(|&(_, (t, b))| t > bottom && b < top);
    let mut free = vec![(col_x, col_x + col_w)];
    for (l, r) in std::iter::once((own_left, own_right)).chain(others.map(|(x, _)| x)) {
        free = free
            .into_iter()
            .flat_map(|(a, z)| [(a, z.min(l)), (a.max(r), z)])
            .filter(|(a, z)| z > a)
            .collect();
    }
    let widest = |keep: &dyn Fn(f32, f32) -> bool| {
        free.iter().filter(|&&(a, z)| keep(a, z)).map(|(a, z)| z - a).fold(0.0, f32::max)
    };
    Some(side_room(
        tb.wrap_text,
        widest(&|_, z| z <= own_left + 0.01),
        widest(&|a, _| a >= own_right - 0.01),
    ))
}

/// The wrap zones of a paragraph's text boxes that `textbox_wraps_beside`,
/// its drawings' anchor top at `anchor_top`.
fn textbox_zones(
    para: &Paragraph,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    anchor_top: f32,
    ctx: &RenderContext,
) -> Vec<FloatZone> {
    para.textboxes
        .iter()
        .filter(|tb| textbox_wraps_beside(tb, para, sp, col, text_width))
        .map(|tb| textbox_zone(tb, sp, col, text_width, anchor_top, ctx))
        .collect()
}

/// The wrap zone of a text box that `textbox_wraps_beside`, its anchor's top
/// at `anchor_top`.
fn textbox_zone(
    tb: &crate::model::Textbox,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    anchor_top: f32,
    ctx: &RenderContext,
) -> FloatZone {
    let h = textbox_render::textbox_height(tb, ctx);
    let (rel, pos) = (tb.v_relative_from, &tb.v_position);
    let y_top = header_footer::resolve_tb_y_top(rel, pos, h, sp, anchor_top);
    FloatZone::for_textbox(tb, textbox_x(tb, sp, col, text_width), y_top, h)
}

/// The band (down from the page top) of a textbox that `textbox_blocks_band`.
fn blocking_textbox_band(
    tb: &crate::model::Textbox,
    para: &Paragraph,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    ctx: &RenderContext,
) -> Option<(f32, f32)> {
    if !textbox_blocks_band(tb, para, sp, col, text_width) {
        return None;
    }
    let h = textbox_render::textbox_height(tb, ctx);
    let top = sp.page_height
        - header_footer::resolve_tb_y_top(tb.v_relative_from, &tb.v_position, h, sp, 0.0);
    Some((top, top + h + tb.dist_bottom))
}

/// A paragraph-relative topAndBottom box hanging below its anchor's text
/// keeps the following paragraphs out of its band; they fill the gap above
/// it first (issue #148: a group 27mm below its anchor's top leaves the next
/// three paragraphs above it). One reaching into the anchor's own text keeps
/// the anchor's reserve.
fn hangs_below_anchor(
    tb: &crate::model::Textbox,
    p: &Paragraph,
    col_w: f32,
    ctx: &RenderContext,
) -> bool {
    tb.v_relative_from == VRelativeFrom::Paragraph
        && !tb.on_line
        && tb.wrap_type == WrapType::TopAndBottom
        && matches!(tb.v_position, VerticalPosition::Offset(o) if o > 0.0 && {
            let lines = paragraph_line_count(p, col_w, ctx).max(1) as f32;
            o >= p.space_before + lines * para_line_h(p, ctx)
        })
}

/// A paragraph's line height from its runs, or from its mark when no run
/// sizes the line.
fn plain_line_h(p: &Paragraph, ctx: &RenderContext) -> f32 {
    let (fs, lhr, _) = tallest_run_metrics(&p.runs, ctx.fonts);
    let (fs, lhr) = if lhr.is_none() {
        unsized_line_metrics(p, fs, ctx.fonts)
    } else {
        (fs, lhr)
    };
    resolve_line_h(p.line_spacing.unwrap_or(ctx.doc_line_spacing), fs, lhr)
}

fn para_line_h(p: &Paragraph, ctx: &RenderContext) -> f32 {
    let (fs, lhr, _) = tallest_run_metrics(&p.runs, ctx.fonts);
    resolve_line_h(p.line_spacing.unwrap_or(ctx.doc_line_spacing), fs, lhr)
}

/// Whether `paragraph_bands` has any, without measuring a textbox.
fn has_bands(p: &Paragraph, sp: &SectionProperties, col: (f32, f32), text_width: f32) -> bool {
    p.textboxes
        .iter()
        .any(|tb| textbox_blocks_band(tb, p, sp, col, text_width))
        || p.floating_images
            .iter()
            .any(|fi| image_blocks_band(fi, sp, col.0, col.1, text_width))
}

/// The bands of a paragraph's blocking page- or margin-anchored floats.
fn paragraph_bands(
    p: &Paragraph,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    ctx: &RenderContext,
) -> Vec<(f32, f32)> {
    let textboxes = p
        .textboxes
        .iter()
        .filter_map(|tb| blocking_textbox_band(tb, p, sp, col, text_width, ctx));
    let pictures = p
        .floating_images
        .iter()
        .filter(|fi| image_blocks_band(fi, sp, col.0, col.1, text_width))
        .map(|fi| image_band(fi, sp));
    textboxes.chain(pictures).collect()
}

fn add_bands(bands: &mut Vec<(f32, f32)>, new: Vec<(f32, f32)>) {
    for b in new {
        if !bands.contains(&b) {
            bands.push(b);
        }
    }
}

/// Before a body block: its paragraph's blocking floats join the page's
/// bands, and its first line steps below any band it would overlap.
/// ponytail: checked at the block start only, like the float zones; a
/// paragraph that runs into a band part way down keeps its later lines there.
fn step_below_bands(
    state: &mut LayoutState,
    block: &Block,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    ctx: &RenderContext,
) {
    let Block::Paragraph(p) = block else {
        step_below_existing_bands(state, block, sp, col.1, ctx);
        return;
    };
    add_bands(
        &mut state.pb.frame_bands,
        paragraph_bands(p, sp, col, text_width, ctx),
    );
    let anchor_y = state
        .pb
        .pending_float_anchor
        .map_or(state.pb.slot_top - state.prev_space_after, |(a, _)| a);
    step_below_existing_bands(state, block, sp, col.1, ctx);
    // After the anchor's own step: these bands only hold off what follows.
    let hanging = p
        .textboxes
        .iter()
        .filter(|tb| hangs_below_anchor(tb, p, col.1, ctx))
        .map(|tb| {
            let top = sp.page_height - anchor_y + tb.v_offset_pt;
            (top, top + textbox_render::textbox_height(tb, ctx) + tb.dist_bottom)
        })
        .collect();
    add_bands(&mut state.pb.frame_bands, hanging);
}

/// Word places all of a page's page- or margin-anchored floats before it lays
/// out the page's lines, so text above an anchor keeps out of its band too
/// (issue #148: a letterhead band at the page top anchored in the third
/// paragraph moves the first two below it). On a page's first block, the
/// bands of the paragraphs estimated to start on the page join its bands.
/// ponytail: plain line counts estimate where the page ends, stopping at the
/// first table or page break; an anchor that still lands on the next page
/// keeps its band on this one too.
fn prescan_page_bands(
    state: &mut LayoutState,
    blocks: &[Block],
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    ctx: &RenderContext,
) {
    let bottom = sp.page_height - state.effective_margin_bottom;
    let mut y = sp.page_height - state.pb.slot_top;
    let mut bands = std::mem::take(&mut state.pb.frame_bands);
    for (i, block) in blocks.iter().enumerate() {
        let Block::Paragraph(p) = block else { break };
        if i > 0 && p.page_break_before {
            break;
        }
        add_bands(&mut bands, paragraph_bands(p, sp, col, text_width, ctx));
        let line_h = para_line_h(p, ctx);
        let lines = paragraph_line_count(p, col.1, ctx).max(1) as f32;
        y = header_footer::below_blocking_frames(y + p.space_before, line_h, &bands)
            + lines * line_h
            + p.space_after;
        if y > bottom || p.page_break_after || p.page_break_at.is_some() {
            break;
        }
    }
    state.pb.frame_bands = bands;
}

/// The width of a paragraph's first word, None for one without text.
fn first_word_width(p: &Paragraph, ctx: &RenderContext) -> Option<f32> {
    if is_text_empty(&p.runs) {
        return None;
    }
    layout::opening_word_width(&p.runs, ctx.fonts)
}

/// How many lines a paragraph takes in a column `col_w` wide.
fn paragraph_line_count(p: &Paragraph, col_w: f32, ctx: &RenderContext) -> usize {
    build_lines(
        &p.runs,
        ctx,
        (col_w - p.indent_left - p.indent_right).max(1.0),
        ctx.cjk(true, p),
        &LineOpts {
            tab_stops: &p.tab_stops,
            indent_left: p.indent_left,
            indent_right: p.indent_right,
            hanging: p.indent_hanging - p.indent_first_line,
            ..Default::default()
        },
    )
    .len()
}

fn step_below_existing_bands(
    state: &mut LayoutState,
    block: &Block,
    sp: &SectionProperties,
    col_w: f32,
    ctx: &RenderContext,
) {
    let top = sp.page_height - state.pb.slot_top;
    let Some(p) = leading_para(block) else { return };
    if state
        .pb
        .frame_bands
        .iter()
        .all(|&(_, bottom)| top >= bottom)
    {
        return;
    }
    let line_h = para_line_h(p, ctx);
    // The first line fits above a band with the space before it, and a
    // one-line paragraph only with its space after too (Word probes, 12.65pt
    // line + 6pt after: an 18.0pt gap between two topAndBottom boxes is
    // skipped, 19.0pt holds it; 14.2pt holds it with no space after).
    let one_line = paragraph_line_count(p, col_w, ctx) <= 1;
    let need = line_h + if one_line { p.space_after } else { 0.0 };
    let line_top = top + state.prev_space_after.max(p.space_before);
    let below = header_footer::below_blocking_frames(line_top, need, &state.pb.frame_bands);
    if below != line_top {
        // Its paragraph-relative drawings still hang from where the paragraph
        // stood (probe: a wrapNone box at offset 0 in a paragraph pushed below
        // a band stays above the band; vaccines' circles under their banner).
        let anchor = state.pb.slot_top - state.prev_space_after;
        state.pb.pending_float_anchor.get_or_insert((anchor, None));
        // Below a band the previous paragraph's space after is gone and this
        // one's own space before stays (probe, 6pt after and 12pt before: the
        // line starts 6pt below the box).
        state.pb.slot_top = sp.page_height - (below - state.prev_space_after);
    }
}

/// A lifted frame being laid out: its blocks render in the frame's own column
/// from its top, and the flow then resumes where it left off, clear of the
/// frame's band (Word probes: body lines step below it with no gap).
struct OpenFrame<'a> {
    props: &'a FrameProperties,
    top: f32,
    /// The frame's column: its left edge and width.
    geometry: (f32, f32),
    /// The height text wraps around (`w:h`, or a drop cap's lines).
    wrap_height: f32,
    flow: SavedFlow,
}

struct SavedFlow {
    slot_top: f32,
    effective_margin_bottom: f32,
    prev_space_after: f32,
    current_col: usize,
    float_zone: Option<FloatZone>,
    more_float_zones: Vec<FloatZone>,
}

impl<'a> OpenFrame<'a> {
    fn open(
        props: &'a FrameProperties,
        blocks: &[Block],
        state: &mut LayoutState,
        ctx: &RenderContext,
        sp: &SectionProperties,
    ) -> Self {
        let width = if props.width > 0.0 {
            props.width
        } else {
            auto_frame_width(props, blocks, ctx, sp.text_width())
        };
        let (text_x, text_w) = (sp.margin_left, sp.text_width());
        let x = match props.drop_cap {
            Some(dc) if dc.in_margin => text_x - props.h_space - width,
            Some(_) => text_x,
            None => resolve_h_position(
                props.h_relative_from,
                &props.h_position,
                width,
                sp,
                text_x,
                text_w,
                text_w,
            ),
        };
        // A drop cap is as tall as the next paragraph's first `lines` lines.
        // ponytail: that paragraph's space before and indents are not added;
        // the probed drop cap has neither.
        let wrap_height = props.drop_cap.map_or(props.height, |dc| {
            let next = block_para(blocks, frame_blocks(props, blocks).count());
            dc.lines as f32 * next.map_or(0.0, |p| table_layout::spacer_line_h(p, ctx))
        });
        // Empty frames draw nothing but keep their marks' lines as their box:
        // text wraps on both sides of it (empty-frame-control, Word online).
        let empty_h: f32 = frame_blocks(props, blocks)
            .filter_map(|b| match b {
                Block::Paragraph(p) if is_empty_wrapping_frame(p) => Some(para_line_h(p, ctx)),
                _ => None,
            })
            .sum();
        let wrap_height = wrap_height.max(empty_h);
        let height = match props.v_position {
            VerticalPosition::AlignBottom | VerticalPosition::AlignCenter
                if props.height == 0.0 =>
            {
                frame_content_height(props, blocks, ctx, width)
            }
            _ => props.height,
        };
        let top = sp.page_height
            - header_footer::resolve_tb_y_top(
                props.v_relative_from,
                &props.v_position,
                height,
                sp,
                state.pb.slot_top,
            );
        let flow = SavedFlow {
            slot_top: std::mem::replace(&mut state.pb.slot_top, sp.page_height - top),
            // A frame never breaks across pages: 3ec631ca50's bottom-aligned
            // address block ends on the bottom margin, where float rounding
            // would push its last line onto the next page.
            effective_margin_bottom: std::mem::take(&mut state.effective_margin_bottom),
            prev_space_after: std::mem::take(&mut state.prev_space_after),
            current_col: std::mem::take(&mut state.current_col),
            float_zone: state.pb.float_zone.take(),
            more_float_zones: std::mem::take(&mut state.pb.more_float_zones),
        };
        OpenFrame {
            props,
            top,
            geometry: (x, width),
            wrap_height,
            flow,
        }
    }

    fn close(self, state: &mut LayoutState, sp: &SectionProperties) {
        let bottom = sp.page_height - state.pb.slot_top;
        // A full-width wrapping frame blocks body lines even if a previous
        // table already owns float_zone. Its overflowing content must not be
        // lost when restoring that table's exclusion zone below.
        let full_width = side_room(
            WrapText::BothSides,
            self.geometry.0 - sp.margin_left,
            sp.page_width - sp.margin_right - self.geometry.0 - self.geometry.1,
        ) < MIN_EMPTY_STRIP;
        if self.props.text_below || full_width
        {
            let wrap_bottom = if full_width && self.props.height_exact && self.wrap_height > 0.0 {
                self.top + self.wrap_height
            } else if full_width { self.top + self.wrap_height.max(bottom - self.top) }
            else { bottom };
            state.pb.frame_bands.push((self.top, wrap_bottom));
        }
        let flow = self.flow;
        state.pb.slot_top = flow.slot_top;
        state.effective_margin_bottom = flow.effective_margin_bottom;
        state.prev_space_after = flow.prev_space_after;
        state.current_col = flow.current_col;
        state.pb.float_zone = flow.float_zone;
        state.pb.more_float_zones = flow.more_float_zones;
        let wrap_height = if self.props.height_exact {
            self.wrap_height
        } else {
            self.wrap_height.max(bottom - self.top)
        };
        let top = sp.page_height - self.top;
        let bottom_y = top - wrap_height - self.props.v_space;
        // A frame the flow has already passed (one anchored at the margin top
        // by a paragraph lower down) leaves nothing to wrap around.
        if !self.props.text_below && !full_width && state.pb.slot_top > bottom_y {
            state.pb.float_zone.get_or_insert(FloatZone {
                top_y: top + self.props.v_space,
                bottom_y,
                obj_left: self.geometry.0,
                obj_right: self.geometry.0 + self.geometry.1,
                left_from_text: self.props.h_space,
                right_from_text: self.props.h_space,
                polygon_pts: None,
                wrap_text: WrapText::BothSides,
                para_relative: false,
                from_table: false,
                table_on_page: false,
            });
        }
    }
}

/// A frame without w:w is as wide as its widest block: a paragraph's picture
/// or longest line, a table's grid.
fn auto_frame_width(
    props: &FrameProperties,
    blocks: &[Block],
    ctx: &RenderContext,
    col_w: f32,
) -> f32 {
    frame_blocks(props, blocks)
        .map(|b| match b {
            Block::Paragraph(p) => p.image.as_ref().map_or_else(
                || {
                    let opts = LineOpts {
                        tab_stops: &p.tab_stops,
                        ..Default::default()
                    };
                    build_lines(&p.runs, ctx, col_w, ctx.cjk(true, p), &opts)
                        .iter()
                        .map(|l| l.total_width)
                        .fold(0.0, f32::max)
                        + p.indent_left
                        + p.indent_right
                },
                |img| img.display_width,
            ),
            Block::Table(t) => t.col_widths.iter().sum(),
        })
        .fold(0.0, f32::max)
}

/// A frame without w:h is as tall as its blocks laid out at its width.
fn frame_content_height(
    props: &FrameProperties,
    blocks: &[Block],
    ctx: &RenderContext,
    width: f32,
) -> f32 {
    let mut height = 0.0;
    let mut prev: Option<&Paragraph> = None;
    let mut prev_space_after = 0.0;
    for block in frame_blocks(props, blocks) {
        match block {
            Block::Paragraph(p) => {
                height += header_footer::hf_paragraph_gap(prev, prev_space_after, p);
                let opts = LineOpts {
                    tab_stops: &p.tab_stops,
                    ..Default::default()
                };
                let line_w = (width - p.indent_left - p.indent_right).max(1.0);
                let lines = build_lines(&p.runs, ctx, line_w, ctx.cjk(true, p), &opts);
                let (fs, lhr, _) = tallest_run_metrics(&p.runs, ctx.fonts);
                let ls = p.line_spacing.unwrap_or(ctx.doc_line_spacing);
                height += lines.len().max(1) as f32 * resolve_line_h(ls, fs, lhr);
                prev = Some(p);
                prev_space_after = p.space_after;
            }
            Block::Table(t) => {
                height += prev_space_after + table::compute_hf_table_height(t, ctx, width);
                prev = None;
                prev_space_after = 0.0;
            }
        }
    }
    height + prev_space_after
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
    /// Word's `wpJustification` (see `docx::settings`).
    pub(super) wp_justification: bool,
    /// `Document::additive_paragraph_spacing`.
    pub(super) additive_paragraph_spacing: bool,
    /// The current section's docGrid line pitch when its grid snaps lines
    /// and `adjustLineHeightInTable` is set (0 otherwise): table cells then
    /// snap to it like body text; physical_therapy (no flag) keeps natural
    /// cell lines on its 18pt grid.
    pub(super) cell_grid_pitch: std::cell::Cell<f32>,
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
            expand_shift_return: !self.do_not_expand_shift_return,
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

/// Compute line height from the list label font if it exceeds the text-run line height.
/// Word includes the numbering label character's font metrics in the tallest-font
/// calculation, so a bullet from Symbol font can make the line taller than text-only
/// Calibri, and an oversized label (e.g. 20pt number on 10pt text) makes the first
/// line taller outright.
pub(super) fn label_boosted_line_h(
    para: &Paragraph,
    fonts: &HashMap<String, FontEntry>,
    text_line_h: f32,
    effective_ls: LineSpacing,
    text_font_size: f32,
    text_lhr: Option<f32>,
    text_ar: Option<f32>,
) -> f32 {
    if para.list_label.is_empty() {
        return text_line_h;
    }
    let label_fs = para.list_label_font_size.unwrap_or(text_font_size);
    let Some(key) = label_font_key(para) else {
        return text_line_h;
    };
    let Some(entry) = fonts.get(&key) else {
        return text_line_h;
    };
    let Some(label_ar) = entry.ascender_ratio else {
        return text_line_h;
    };
    if para.text_alignment == crate::model::TextAlignment::Center {
        return if matches!(effective_ls, LineSpacing::Exact(_)) {
            text_line_h
        } else {
            text_line_h.max(label_fs * entry.line_h_ratio.unwrap_or(1.0))
        };
    }
    // Word raises the line by the marker's ascent but keeps the text's descent,
    // never one font's whole line height: a Symbol bullet reaches 0.6pt above
    // 11pt Calibri and Word's line grows by exactly that (annotation #66, 16.0pt
    // not 15.5), while a Courier New "o" or a Symbol bullet on Arial, whose
    // descents are deeper than the text's, leave the line at the text height
    // (streamnet p5, dialysis). Measured against Word, not from the spec.
    let text_ascent = text_font_size * text_ar.unwrap_or(0.75);
    let label_ascent = label_fs * label_ar;
    let natural = match effective_ls {
        // A multiple scales the text's own line and the marker's extra ascent
        // is added once, unscaled: SymbolMT on 12pt Aptos at 278/240 gives
        // 16.97 + 0.80 = 17.76 (Word 17.75, case3), on 11pt Calibri at 1.15
        // 15.44 + 0.59 = 16.03 (Word 16.00, case33); scaling the marker's
        // ascent with the text gave 17.89 and 16.12.
        LineSpacing::Auto(_) => {
            resolve_line_h(effective_ls, text_font_size, text_lhr)
                + (label_ascent - text_ascent).max(0.0)
        }
        _ => {
            let descent = text_font_size * descender_ratio(text_lhr, text_ar);
            resolve_line_h(
                effective_ls,
                1.0,
                Some(text_ascent.max(label_ascent) + descent),
            )
        }
    };
    natural.max(text_line_h)
}

/// First-baseline offset including the list label's ascent. The label is a run
/// on the first line, so a numbering label that reaches higher than the text
/// (oversized, or a taller face) pushes the first baseline down to its ascent.
fn label_boosted_baseline_offset(
    para: &Paragraph,
    fonts: &HashMap<String, FontEntry>,
    text_offset: f32,
    text_font_size: f32,
) -> f32 {
    if para.list_label.is_empty() || para.text_alignment == crate::model::TextAlignment::Center {
        return text_offset;
    }
    let label_fs = para.list_label_font_size.unwrap_or(text_font_size);
    // Compared by ascent, not size: an 11pt Symbol bullet reaches 0.6pt above
    // 11pt Calibri, and Word puts that above the first baseline (annotation #66).
    let label_ar = label_font_key(para)
        .and_then(|k| fonts.get(&k))
        .and_then(|e| e.ascender_ratio)
        .unwrap_or(0.75);
    text_offset.max(label_fs * label_ar)
}

/// (font_size, line_h_ratio) for a paragraph whose runs size nothing (empty,
/// breaks, whitespace): a break sizes the line it ends (samtale's 10pt breaks
/// under an 11pt mark), otherwise only the paragraph mark is left (eco_int's
/// lone 9.5pt space takes the mark's Calibri 11 line).
fn unsized_line_metrics(
    para: &Paragraph,
    font_size: f32,
    fonts: &HashMap<String, FontEntry>,
) -> (f32, Option<f32>) {
    let (size, metrics) = unsized_line_font(para, font_size, fonts);
    (size, metrics.0)
}

/// The font that sizes a line no glyph run sizes (only tabs and spaces): the
/// first break's, else the paragraph mark's, as (size, run_line_metrics).
fn unsized_line_font(
    para: &Paragraph,
    font_size: f32,
    fonts: &HashMap<String, FontEntry>,
) -> (f32, (Option<f32>, Option<f32>)) {
    match para.runs.iter().find(|r| r.is_line_break) {
        Some(br) => break_font(br, fonts),
        None => mark_font(para, font_size, fonts),
    }
}

fn break_font(br: &Run, fonts: &HashMap<String, FontEntry>) -> (f32, (Option<f32>, Option<f32>)) {
    let metrics = fonts
        .get(&font_key(br))
        .map_or((None, None), |e| run_line_metrics(e, ""));
    (br.font_size, metrics)
}

fn mark_font(
    para: &Paragraph,
    font_size: f32,
    fonts: &HashMap<String, FontEntry>,
) -> (f32, (Option<f32>, Option<f32>)) {
    let metrics = para
        .paragraph_mark_font_name
        .as_deref()
        .and_then(|n| fonts.get(n))
        .map_or((None, None), |e| run_line_metrics(e, ""));
    (para.paragraph_mark_font_size.unwrap_or(font_size), metrics)
}

/// Each line no glyph sizes (only tabs and spaces) takes the font of what
/// ends it, its ascent included: its break's, or the paragraph mark's on the
/// last line, whatever sizes the paragraph's other lines (Word example: tab
/// lines ended by a 20pt break and an 11pt mark sit 18.1pt and 9.96pt
/// below their tops, beside text lines too). Under at-least spacing that is
/// the font's baseline in the taller box (`boxed_line_ascent`), where a text
/// line in that font would sit. `ascent` is where the paragraph puts its
/// first baseline.
fn size_unsized_lines(
    lines: &mut [TextLine],
    para: &Paragraph,
    runs: &[Run],
    ls: LineSpacing,
    (line_h, font_size, ascent): (f32, f32, f32),
    fonts: &HashMap<String, FontEntry>,
) {
    let mut breaks = runs.iter().filter(|r| r.is_line_break && !r.vanish);
    let last = lines.len().saturating_sub(1);
    for (i, line) in lines.iter_mut().enumerate() {
        let br = if line.ends_with_break { breaks.next() } else { None };
        if line.pitch.is_some()
            || line.break_font_size.is_some()
            || line.float_skip
            || line.chunks.iter().any(|c| {
                c.math.is_some() || c.inline_image_name.is_some() || !c.text.trim().is_empty()
            })
        {
            continue;
        }
        let (size, (lhr, ar)) = match br {
            Some(br) => break_font(br, fonts),
            None if i == last => mark_font(para, font_size, fonts),
            None => continue,
        };
        let pitch = resolve_line_h(ls, size, lhr);
        let own_ascent = layout::boxed_line_ascent(ls, pitch, size, lhr, ar, &[], fonts)
            .unwrap_or(size * ar.unwrap_or(0.75));
        if (pitch - line_h).abs() > 0.01 || (own_ascent - ascent).abs() > 0.01 {
            line.pitch = Some(pitch);
            line.ascent_shift = own_ascent - ascent;
        }
    }
}

/// Size the shifted mark's box before applying the paragraph's line rule.
/// Moving the entire mark also moves its descent, so a small lowering can
/// leave the occupied height unchanged rather than stretching both sides.
fn mark_position_stretch(
    para: &Paragraph,
    ls: LineSpacing,
    grid_snapped: bool,
    fonts: &HashMap<String, FontEntry>,
) -> f32 {
    if grid_snapped || matches!(ls, LineSpacing::Exact(_)) || para.paragraph_mark_position == 0.0 {
        return 0.0;
    }
    let font = para
        .paragraph_mark_font_name
        .as_deref()
        .and_then(|name| fonts.get(name));
    let metrics = font.map(|font| run_line_metrics(font, ""));
    let (Some(font), (Some(lhr), Some(ar))) = (font, metrics.unwrap_or((None, None))) else {
        return para.paragraph_mark_position.abs();
    };
    let size = para.paragraph_mark_font_size.unwrap_or(11.0);
    let pos = para.paragraph_mark_position;
    let natural = size * lhr;
    let shifted = if para.text_alignment == crate::model::TextAlignment::Top {
        natural + layout::top_aligned_drop(font, size, pos)
    } else {
        (size * ar + pos).max(0.0) + (size * (lhr - ar).max(0.0) - pos).max(0.0)
    };
    match ls {
        LineSpacing::Auto(multiplier) => (shifted - natural).max(0.0) * multiplier,
        LineSpacing::AtLeast(minimum) => shifted.max(minimum) - natural.max(minimum),
        LineSpacing::Exact(_) => 0.0,
    }
}

/// A line holding only a w:br is as tall as the break run, wherever it
/// falls: pasto's title opens with an unformatted <w:br/> (11pt) above
/// its 12pt bold text, and Word steps 12.65 for that line, not 13.80.
fn size_break_lines(
    lines: &mut [TextLine],
    runs: &[Run],
    ls: LineSpacing,
    line_h: f32,
    fonts: &HashMap<String, FontEntry>,
) {
    for line in lines
        .iter_mut()
        .filter(|l| l.ends_with_break && l.pitch.is_none())
    {
        if let Some(bfs) = line.break_font_size {
            let lhr = line.break_lhr.or_else(|| break_run_lhr(runs, bfs, fonts));
            let pitch = resolve_line_h(ls, bfs, lhr);
            if (pitch - line_h).abs() > 0.01 {
                line.pitch = Some(pitch);
            }
        }
    }
}

/// The pitch of a line holding only the paragraph mark.
fn mark_line_h(
    para: &Paragraph,
    ls: LineSpacing,
    fallback_lhr: Option<f32>,
    fonts: &HashMap<String, FontEntry>,
) -> Option<f32> {
    let mark_size = para.paragraph_mark_font_size?;
    let mark_lhr = mark_font(para, mark_size, fonts).1.0.or(fallback_lhr);
    Some(resolve_line_h(ls, mark_size, mark_lhr))
}

/// A final empty line after a break holds the paragraph mark, which sizes
/// it rather than the preceding text.
fn size_mark_line(
    lines: &mut [TextLine],
    para: &Paragraph,
    runs: &[Run],
    ls: LineSpacing,
    fallback_lhr: Option<f32>,
    fonts: &HashMap<String, FontEntry>,
) {
    if runs.last().is_some_and(|r| r.is_line_break)
        && let Some(last) = lines.last_mut().filter(|l| l.chunks.is_empty())
        && let Some(mark_h) = mark_line_h(para, ls, fallback_lhr, fonts)
    {
        last.pitch = Some(mark_h);
    }
}

/// The blank line a wrapped tab leaves at the paragraph end also holds the
/// mark, so it is as tall as the larger of the two (geneva: a 10pt tab
/// before an 11pt mark steps 12.65, the mark's pitch).
fn size_tab_mark_line(
    lines: &mut [TextLine],
    para: &Paragraph,
    ls: LineSpacing,
    fallback_lhr: Option<f32>,
    fonts: &HashMap<String, FontEntry>,
) {
    if let Some(last) = lines
        .last_mut()
        .filter(|l| l.leading_tab_wrap && l.chunks.is_empty())
        && let Some(pitch) = last.pitch
        && let Some(mark_h) = mark_line_h(para, ls, fallback_lhr, fonts)
    {
        last.pitch = Some(pitch.max(mark_h));
    }
}

/// Look up the line_h_ratio for a break run's font, matching by font_size.
fn break_run_lhr(runs: &[Run], break_fs: f32, fonts: &HashMap<String, FontEntry>) -> Option<f32> {
    // Find the break run with the matching font size
    let br_run = runs
        .iter()
        .rfind(|r| r.is_line_break && (r.font_size - break_fs).abs() < 0.01);
    if let Some(run) = br_run {
        let key = font_key(run);
        fonts.get(&key).and_then(|e| e.line_h_ratio)
    } else {
        None
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

/// Minimum side-strip width for a line box to sit beside a wrapping float.
/// Empirical: brazilian_logistics_study absorbs empty paragraphs beside a
/// float with ~42pt strips; sample500kB (image width == column width, 0pt
/// strips) stacks them below.
const MIN_EMPTY_STRIP: f32 = 18.0;

/// The part of a collapsed paragraph gap that is the paragraph's own space
/// before, beyond the previous paragraph's space after.
pub(super) fn own_space_before(gap: f32, prev_space_after: f32) -> f32 {
    (gap - prev_space_after).max(0.0)
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
    /// The highest a wrapping paragraph-relative picture may sit: the top
    /// margin under compat 15 (Word probes: -30pt from a first paragraph
    /// lands on the margin; compat 14 and older keep the -30pt).
    ceiling: f32,
}

impl ParaAnchor {
    /// `pending`: a floating table's one-shot override for its anchor
    /// paragraph; `line_top`: the first line's top, below the space before.
    fn new(
        pending: Option<f32>,
        line_top: f32,
        own_before: f32,
        sp: &SectionProperties,
        ctx: &RenderContext,
    ) -> Self {
        Self {
            float: pending.unwrap_or(line_top + own_before),
            on_line: pending.unwrap_or(line_top),
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

    pub(super) fn top(self, on_line: bool) -> f32 {
        if on_line { self.on_line } else { self.float }
    }

    /// A body float's top (`resolve_fi_y_top`) under the anchor's ceiling.
    pub(super) fn fi_y_top(self, fi: &crate::model::FloatingImage, sp: &SectionProperties) -> f32 {
        let y = resolve_fi_y_top(fi, sp, self.top(fi.on_line));
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
        let y = if wrapping_para && h > sp.page_height - sp.margin_top - sp.margin_bottom && y - h < 0.0 {
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
    fn float_anchor_for(&self, block_idx: usize) -> Option<f32> {
        self.pending_float_anchor
            .filter(|&(_, b)| b.is_none_or(|b| b == block_idx))
            .map(|(a, _)| a)
    }

    fn take_float_anchor_for(&mut self, block_idx: usize) -> Option<f32> {
        let anchor = self.float_anchor_for(block_idx)?;
        self.pending_float_anchor = None;
        Some(anchor)
    }
}

pub(super) struct FloatTableExtent {
    pub(super) left: f32,
    pub(super) right: f32,
    pub(super) left_from_text: f32,
    pub(super) right_from_text: f32,
}

/// The wider of the two strips a column leaves beside a float.
fn side_strip(
    (obj_left, obj_right): (f32, f32),
    (left_from_text, right_from_text): (f32, f32),
    (col_x, col_w): (f32, f32),
) -> f32 {
    (obj_left - left_from_text - col_x).max(col_x + col_w - obj_right - right_from_text)
}

#[derive(Clone)]
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
    /// Set for a floating table: like a paragraph-relative float it sits
    /// tblpY below its anchor paragraph, which comes next in the flow.
    pub from_table: bool,
    /// A floating table anchored to the page or margin, not to its text.
    pub table_on_page: bool,
}

/// The text region right of a float's wrap edge `float_right`: Word starts
/// the line at the wrap edge or at the paragraph's left indent, whichever is
/// further right, for tables, pictures and text boxes alike (examples: a box
/// whose wrap edge is 139pt in starts lines with a 0pt and a 129pt indent at
/// its edge, a 200pt one at the indent).
fn right_text_region(
    col_x: f32,
    col_right: f32,
    float_right: f32,
    para: &Paragraph,
) -> (f32, f32) {
    let left = (col_x + para.indent_left).max(float_right);
    (left, (col_right - para.indent_right - left).max(0.0))
}

/// The text region left of a float's wrap edge `float_left`: the paragraph's
/// box cut off at the edge (example: a right indent inside the float's area
/// takes nothing more off the line).
fn left_text_region(
    col_x: f32,
    col_right: f32,
    float_left: f32,
    para: &Paragraph,
) -> (f32, f32) {
    let left = col_x + para.indent_left;
    (left, (float_left.min(col_right - para.indent_right) - left).max(0.0))
}

impl FloatZone {
    /// The exclusion zone of a wrapping float whose frame's left edge is `fi_x`
    /// and top edge `fi_y_top` (page coordinates, y from the bottom).
    fn for_float(fi: &crate::model::FloatingImage, fi_x: f32, fi_y_top: f32) -> Self {
        let (w, h) = (fi.image.display_width, fi.image.display_height);
        FloatZone {
            top_y: fi_y_top + fi.dist_top,
            bottom_y: fi_y_top - h - fi.dist_bottom,
            obj_left: fi_x,
            obj_right: fi_x + w,
            left_from_text: fi.dist_left,
            right_from_text: fi.dist_right,
            polygon_pts: fi
                .wrap_polygon
                .as_ref()
                .map(|verts| convert_polygon_to_page_coords(verts, fi_x, fi_y_top, w, h)),
            wrap_text: fi.wrap_text,
            para_relative: fi.v_relative_from == VRelativeFrom::Paragraph,
            from_table: false,
            table_on_page: false,
        }
    }

    /// The exclusion zone of a text box or shape `h` tall whose frame's left
    /// edge is `x` and top edge `y_top`: the same model as a picture's.
    fn for_textbox(tb: &crate::model::Textbox, x: f32, y_top: f32, h: f32) -> Self {
        let (left, top, w, h) = textbox_wrap_rect(tb, x, y_top, h);
        let unturned = left == x && top == y_top;
        FloatZone {
            top_y: top + tb.dist_top,
            bottom_y: top - h - tb.dist_bottom,
            obj_left: left,
            obj_right: left + w,
            left_from_text: tb.dist_left,
            right_from_text: tb.dist_right,
            polygon_pts: tb
                .wrap_polygon
                .as_ref()
                .filter(|_| unturned)
                .map(|verts| convert_polygon_to_page_coords(verts, left, top, w, h)),
            wrap_text: tb.wrap_text,
            para_relative: tb.v_relative_from == VRelativeFrom::Paragraph,
            from_table: false,
            table_on_page: false,
        }
    }

    /// Returns (left_edge, right_edge) of the exclusion zone at the given Y.
    /// Falls back to rectangular bounds if no polygon or scanline misses.
    fn exclusion_at_y(&self, y: f32) -> (f32, f32) {
        if let Some(ref pts) = self.polygon_pts
            && let Some((left, right)) = poly_scanline(pts, y)
        {
            return (left, right);
        }
        (self.obj_left, self.obj_right)
    }

    /// Exclusion over a whole line box, `top` to `bottom`: Word keeps a line
    /// clear of the polygon anywhere in its height (case42's wrap edges match
    /// it to 0.3pt), where one scanline at the line top lags a line behind
    /// a shape that widens downwards.
    fn exclusion_in_band(&self, top: f32, bottom: f32) -> (f32, f32) {
        let Some(pts) = self.polygon_pts.as_ref() else {
            return (self.obj_left, self.obj_right);
        };
        // A polygon's extremes over a band lie on the band's edges or its vertices.
        let inner = pts.iter().map(|p| p.1).filter(|&y| y < top && y > bottom);
        [top, bottom]
            .into_iter()
            .chain(inner)
            .filter_map(|y| poly_scanline(pts, y))
            .reduce(|(l0, r0), (l1, r1)| (l0.min(l1), r0.max(r1)))
            .unwrap_or((self.obj_left, self.obj_right))
    }

    /// Narrow a paragraph's text box (`text_x`, `text_w`, `label_x`) to fit
    /// beside this floating object when `y`, the paragraph's first line top,
    /// is inside the zone. Leaves the box as it is otherwise, or when no side
    /// is wide enough.
    #[allow(clippy::too_many_arguments)]
    fn narrow_paragraph(
        &self,
        y: f32,
        col_x: f32,
        col_w: f32,
        para: &Paragraph,
        text_x: &mut f32,
        text_w: &mut f32,
        label_x: &mut f32,
        legacy_indents: bool,
    ) {
        if !(y <= self.top_y && y > self.bottom_y) {
            return;
        }
        let col_right = col_x + col_w;
        let (ex_left, ex_right) = self.exclusion_at_y(y);
        let space_right = col_right - (ex_right + self.right_from_text);
        let space_left = (ex_left - self.left_from_text) - col_x;

        let float_left = ex_left - self.left_from_text;
        let float_right = ex_right + self.right_from_text;
        let (_, lw) = left_text_region(col_x, col_right, float_left, para);
        let (rx, rw) = right_text_region(col_x, col_right, float_right, para);
        let mut use_right_region = || {
            *text_w = rw.max(1.0);
            *text_x = rx;
            let first =
                first_line_region_x(rx, float_right, para, para.indent_hanging, legacy_indents);
            *label_x = first - para.indent_hanging + para.indent_first_line;
        };
        if self.wrap_text == WrapText::BothSides {
            // For bothSides, use the wider region as primary text width (dual
            // geometry handles both regions per-line).
            if rw > lw {
                use_right_region();
            } else if lw > 0.0 {
                *text_w = lw.max(1.0);
            }
        } else {
            match wrap_side(self.wrap_text, space_left, space_right) {
                Some(true) => use_right_region(),
                Some(false) => *text_w = lw.max(1.0),
                None => {}
            }
        }
    }
}

/// The side of a float a line beside it takes (true: the right), None when
/// the side its wrapText allows (for largest, the wider) leaves less than
/// `MIN_EMPTY_STRIP` (Word examples: a 30pt strip beside a wrapText=largest
/// picture holds lines, a 17pt one sends them below it).
fn wrap_side(wrap_text: WrapText, left: f32, right: f32) -> Option<bool> {
    let right_side = match wrap_text {
        WrapText::Left => false,
        WrapText::Right => true,
        WrapText::BothSides | WrapText::Largest => right >= left,
    };
    let room = if right_side { right } else { left };
    (room >= MIN_EMPTY_STRIP).then_some(right_side)
}

/// Where a paragraph's first line region beside a float starts, `x` being
/// the region's start for the other lines and `edge` the float's wrap edge:
/// the first-line indent moves the line right of it, a hanging indent never
/// left of the edge (Word examples: left 36pt + first line 18pt starts 18pt
/// past the edge; left 129pt or 150pt with a 40pt hanging indent at the
/// edge). Before compat 15 a hanging indent's first line counts its indent
/// from the edge instead (examples: left 150pt, hanging 40pt starts 110pt past
/// it; french_sexual_health_youth_strategy's arrows beside its logo). The
/// renderer shifts the first line by the hanging indent, so the region
/// starts that much further right.
fn first_line_region_x(
    x: f32,
    edge: f32,
    para: &Paragraph,
    hanging: f32,
    legacy: bool,
) -> f32 {
    if legacy && para.indent_hanging > 0.0 {
        edge + (para.indent_left - para.indent_hanging).max(0.0) + para.indent_hanging
    } else {
        x.max(edge + hanging.max(0.0))
    }
}

/// The per-line entries of the lines `blocked` does not mark.
fn keep_open<T: Copy>(per_line: &[T], blocked: &[bool]) -> Vec<T> {
    let blocked = blocked.iter().copied().chain(std::iter::repeat(false));
    per_line.iter().zip(blocked).filter(|(_, b)| !b).map(|(&v, _)| v).collect()
}

/// `built`, laid out over the lines `blocked` does not mark, with an empty
/// line (`float_skip`) at each marked one: a line beside a float with no room
/// on a side text may take stays empty and the text goes on below it.
fn skip_blocked_lines(blocked: &[bool], built: Vec<TextLine>) -> Vec<TextLine> {
    let mut built = built.into_iter().peekable();
    let mut lines = Vec::new();
    for &b in blocked {
        if built.peek().is_none() {
            break;
        }
        lines.push(match b {
            false => built.next().unwrap(),
            true => TextLine {
                float_skip: true,
                ..TextLine::default()
            },
        });
    }
    lines.extend(built);
    lines
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
    if let Some(ref pts) = fz.polygon_pts
        && pts.len() >= 3
    {
        content.set_stroke_rgb(0.0, 0.7, 0.0);
        content.set_line_width(0.5);
        content.move_to(pts[0].0, pts[0].1);
        for &(x, y) in &pts[1..] {
            content.line_to(x, y);
        }
        content.close_path();
        content.stroke();
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
    /// Raw `tblpY` offset in points (positive = down from the anchor). Needed by
    /// the flushed-to-next-page path, which re-anchors a `vertAnchor="text"`
    /// table to the new page body top and must re-apply this offset so the table
    /// sits where Word puts it (often above the top margin for negative tblpY).
    pub v_offset_pt: f32,
    /// True when `vertAnchor="text"` — the offset is relative to the anchor
    /// paragraph (re-applicable on a fresh page), not the page/margin.
    pub v_anchor_text: bool,
}

impl FloatingTablePos {
    /// Preserve the text distance when a non-overlapping table moves below a float.
    pub(super) fn constrain_nonoverlap_left(
        &mut self,
        pos: &crate::model::TablePosition,
        sp: &SectionProperties,
        col_x: f32,
        compat_mode: u32,
    ) {
        if compat_mode >= 15
            && !pos.allow_overlap
            && matches!(pos.h_position, crate::model::HorizontalPosition::Offset(v) if v >= 0.0)
            && matches!(pos.h_anchor, "margin" | "column")
        {
            let text_left = if pos.h_anchor == "margin" {
                sp.margin_left
            } else {
                col_x
            };
            self.x = self.x.max(text_left + pos.left_from_text);
        }
    }

    /// Where a floating table goes: `tblpX`/`tblpXSpec` against its anchor
    /// column (`col_x`, `col_w`), `tblpY` below the page, the margin or the
    /// text (`text_y`: where its anchor paragraph's flow is).
    pub(super) fn resolve(
        table: &crate::model::Table,
        pos: &crate::model::TablePosition,
        sp: &SectionProperties,
        col_x: f32,
        col_w: f32,
        text_y: f32,
        ctx: &RenderContext,
    ) -> Self {
        let h_relative_from = match pos.h_anchor {
            "page" => HRelativeFrom::Page,
            "margin" => HRelativeFrom::Margin,
            _ => HRelativeFrom::Column,
        };
        let x = resolve_h_position(
            h_relative_from,
            &pos.h_position,
            table.col_widths.iter().sum(),
            sp,
            col_x,
            col_w,
            sp.text_width(),
        );
        // Before compat 15 an offset places the first cell's text, not
        // its border, as tblInd does inline: Word draws case46's R1 at
        // the margin and the border a cell margin left of it.
        // From compat 15 the offset places the left border's outer edge, as
        // tblInd does inline (Word probe: 0.5/1.5/6pt borders put floating and
        // inline tables' first-column text at the same x).
        use crate::model::HorizontalPosition::Offset;
        let x = match pos.h_position {
            Offset(_) if ctx.compat_mode < 15 => x - table.first_cell_left_margin(),
            Offset(_) => x + table::word2013_border_shift(table),
            _ => x,
        };
        let y = match pos.v_anchor {
            "page" => sp.page_height - pos.v_offset_pt,
            "margin" => sp.page_height - sp.margin_top - pos.v_offset_pt,
            _ => text_y - pos.v_offset_pt,
        };
        FloatingTablePos {
            x,
            y,
            top_from_text: pos.top_from_text,
            bottom_from_text: pos.bottom_from_text,
            left_from_text: pos.left_from_text,
            right_from_text: pos.right_from_text,
            v_offset_pt: pos.v_offset_pt,
            v_anchor_text: pos.v_anchor == "text",
        }
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
    /// A continuous section whose first ordinary paragraph has not flowed yet.
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
    /// The (top, bottom) bands of this page's lifted frames, down from the
    /// page top: body lines may not overlap them (see `OpenFrame`).
    frame_bands: Vec<(f32, f32)>,
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
    pub(super) pending_float_anchor: Option<(f32, Option<usize>)>,
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
            all_content_bottom: Vec::new(),
            all_deferred_shapes: Vec::new(),
            all_behind_shapes: Vec::new(),
            all_links: Vec::new(),
            all_comment_anchors: Vec::new(),
            all_footnote_ids: Vec::new(),
            footnote_cols: Vec::new(),
            all_footnote_cols: Vec::new(),
            col_fn_reserved: 0.0,
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

    pub(super) fn flush_page(&mut self, sect_idx: usize) {
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
        header_footer::footer_top(sp, self.is_first_page_of_section, self.page_count(), ctx)
    }

    fn anchor_slot_top(&self, sp: &SectionProperties, ctx: &RenderContext) -> f32 {
        match self.page_top_header(sp, ctx) {
            Some(is_first) => self
                .slot_top
                .max(header_footer::flow_slot_top(sp, is_first, self.page_count(), ctx)),
            None => self.slot_top,
        }
    }

    /// Which header variant (first page or not) opened this page, when the
    /// cursor still sits at its top: the one whose body top matches.
    fn page_top_header(&self, sp: &SectionProperties, ctx: &RenderContext) -> Option<bool> {
        if !self.is_at_page_top(sp) {
            return None;
        }
        let page = self.page_count();
        [true, false].into_iter().find(|&is_first| {
            (effective_slot_top(sp, is_first, page, ctx) - self.page_top_y).abs() < 0.01
        })
    }

    /// Where a table spanning `left..right` starts: a page-top table drops
    /// below the header's wrapping floats it would overlap, see
    /// `header_float_floor`.
    fn below_header_floats(&self, sp: &SectionProperties, ctx: &RenderContext, left: f32, right: f32) -> f32 {
        self.page_top_header(sp, ctx)
            .and_then(|is_first| {
                header_footer::header_float_floor(sp, is_first, self.page_count(), ctx, left, right)
            })
            .map_or(self.slot_top, |floor| self.slot_top.min(floor))
    }

    fn is_at_page_top(&self, sp: &SectionProperties) -> bool {
        // nabl's 47.49pt header ends below its 36pt top margin: Word still
        // drops the space before the first paragraph under it.
        (self.slot_top - (sp.page_height - sp.margin_top)).abs() < 1.0
            || (self.slot_top - self.page_top_y).abs() < 0.01
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

    fn take_pending_first_page(&mut self, sect_idx: usize) -> bool {
        self.pending_section_first_page
            .take()
            .is_some_and(|(section, page, y)| {
                section == sect_idx && page == self.page_count() && (self.slot_top - y).abs() < 0.01
            })
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
        self.flush_page(sect_idx);
        let page = self.page_count();
        self.slot_top = effective_slot_top(sp, first_page, page, ctx);
        self.page_top_y = self.slot_top;
        if let Some((left, right)) = self.table_extent {
            self.slot_top = self.below_header_floats(sp, ctx, left, right);
        }
        self.column_top_y = self.slot_top;
        self.page_top_y = self.slot_top;
        *effective_margin_bottom = compute_effective_margin_bottom(sp, first_page, page, ctx);
        self.is_first_page_of_section = first_page;
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
}

impl LayoutState {
    /// A throwaway copy for laying out a column region again with its
    /// columns ending at `bottom`; its output is discarded.
    fn trial(&self, bottom: f32) -> LayoutState {
        let mut pb = PageBuilder::new(self.pb.slot_top);
        pb.column_top_y = self.pb.column_top_y;
        pb.page_top_y = self.pb.page_top_y;
        pb.is_first_page_of_section = self.pb.is_first_page_of_section;
        pb.page_hf_section = self.pb.page_hf_section;
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
/// else its whole pitch — above a footnote area, and for an exact-spaced
/// line, whose box is fixed (Word for Mac probe p27: 25 lines of 25pt on a
/// 648pt body, the 26th moves on although its 12pt glyphs would end 0.5pt
/// above the margin).
fn line_fit_height(whole_box: bool, own_pitch: Option<f32>, line_h: f32, text_h: f32) -> f32 {
    if whole_box {
        own_pitch.unwrap_or(line_h)
    } else {
        own_pitch.map_or(text_h, |p| p.min(text_h))
    }
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

/// Modern Word keeps a margin-anchored empty frame's box even though its
/// paragraphs consume no body flow. A full-width inline table cannot wrap
/// beside that box, nor can body text above the anchor (Word probes:
/// empty-margin-frame-visible-body-before-frame). Only boxes whose anchor
/// fits this page are collected; frames after forced breaks are left alone.
fn empty_frame_table_bands(
    blocks: &[Block],
    sp: &SectionProperties,
    ctx: &RenderContext,
    body_top: f32,
    available_h: f32,
) -> Vec<(f32, f32)> {
    if ctx.compat_mode < 15 {
        return Vec::new();
    }
    let mut bands = Vec::new();
    let mut i = 0;
    while i < blocks.len() {
        let Block::Paragraph(p) = &blocks[i] else {
            i += 1;
            continue;
        };
        if p.is_section_break {
            break;
        }
        let Some(fp) = p
            .frame_props
            .as_ref()
            .filter(|_| is_empty_wrapping_frame(p))
        else {
            if p.runs.iter().any(|r| {
                r.inline_image.is_some()
                    || r.field_code.is_some()
                    || r.checkbox.is_some()
                    || r.footnote_id.is_some()
                    || r.endnote_id.is_some()
            }) || p.image.is_some()
                || p.inline_chart.is_some()
                || !p.smartart.is_empty()
                || !p.floating_images.is_empty()
                || !p.textboxes.is_empty()
                || !p.connectors.is_empty()
                || p.page_break_before
                || p.page_break_after
                || p.page_break_at.is_some()
                || p.column_break_before
                || p.column_break_after
            {
                break;
            }
            i += 1;
            continue;
        };
        if fp.v_relative_from != VRelativeFrom::Margin
            || !matches!(&fp.v_position, VerticalPosition::Offset(y) if *y == 0.0)
            || fp.width <= 0.0
            || p.paragraph_mark_vanish
        {
            i += 1;
            continue;
        }
        let start = i;
        let mut height = 0.0;
        let mut prev_after: f32 = 0.0;
        while let Some(Block::Paragraph(q)) = blocks.get(i) {
            if q.is_section_break
                || q.paragraph_mark_vanish
                || !is_empty_wrapping_frame(q)
                || q.frame_props.as_ref() != Some(fp)
            {
                break;
            }
            height += prev_after.max(q.space_before)
                + plain_line_h(q, ctx) * line_count(q, ctx, fp.width).max(1) as f32;
            prev_after = q.space_after;
            i += 1;
        }
        let top = sp.margin_top - fp.v_space;
        let bottom = sp.margin_top + height + prev_after + fp.v_space;
        let x = resolve_h_position(
            fp.h_relative_from,
            &fp.h_position,
            fp.width,
            sp,
            sp.margin_left,
            sp.text_width(),
            sp.text_width(),
        );
        if bottom <= body_top
            || x + fp.width + fp.h_space <= sp.margin_left
            || x - fp.h_space >= sp.page_width - sp.margin_right
        {
            continue;
        }
        // Conservatively count floating extents too. Never reserve a box on
        // this page if its introductory anchor could already be on the next.
        let mut prefix_h = height + prev_after;
        let fits = blocks[..start].iter().all(|b| {
            match b {
                Block::Table(t) => {
                    if t.rows
                        .iter()
                        .flat_map(|r| r.cells.iter())
                        .flat_map(|c| c.all_paragraphs())
                        .any(|q| {
                            q.page_break_before
                                || q.page_break_after
                                || q.page_break_at.is_some()
                                || q.runs
                                    .iter()
                                    .any(|r| r.footnote_id.is_some() || r.endnote_id.is_some())
                        })
                    {
                        return false;
                    }
                    let widths =
                        table_layout::auto_fit_columns(
                            t,
                            ctx.fonts,
                            None,
                            Some(sp.text_width()),
                            t.autofit_overhang(ctx.compat_mode),
                        );
                    prefix_h += table_layout::compute_row_layouts(t, &widths, ctx, None)
                        .iter()
                        .map(|r| r.height)
                        .sum::<f32>();
                    if let Some(pos) = &t.position {
                        prefix_h += pos.v_offset_pt.max(0.0)
                            + pos.top_from_text.max(0.0)
                            + pos.bottom_from_text.max(0.0);
                    }
                }
                Block::Paragraph(q) => {
                    if !is_empty_wrapping_frame(q) {
                        prefix_h += q.space_before.max(0.0)
                            + q.space_after.max(0.0)
                            + plain_line_h(q, ctx)
                                * line_count(q, ctx, sp.text_width()).max(1) as f32;
                    }
                }
            }
            prefix_h <= available_h
        }) && prefix_h <= available_h;
        if fits {
            bands.push((top, bottom));
        }
    }
    bands
}

/// Empty auto-height wrapping frames have no body-flow content. Ordinary
/// empty paragraphs and explicit breaks retain their paragraph-mark lines.
/// Inside a lifted frame the mark keeps its line too: Word spaces
/// 3ec631ca50's address groups with empty frame paragraphs.
fn is_empty_wrapping_frame(para: &Paragraph) -> bool {
    para.frame_props
        .as_ref()
        .is_some_and(|fp| !fp.text_below && fp.height == 0.0)
        && para
            .runs
            .iter()
            .all(|r| r.vanish || r.text.bytes().all(|b| b == b' '))
        && !para.runs.iter().any(|r| {
            r.is_line_break
                || r.is_tab
                || r.field_code.is_some()
                || r.inline_image.is_some()
                || r.checkbox.is_some()
        })
        && para.image.is_none()
        && para.inline_chart.is_none()
        && para.smartart.is_empty()
        && para.floating_images.is_empty()
        && para.textboxes.is_empty()
        && para.connectors.is_empty()
        && para.list_label.is_empty()
        && para.shading.is_none()
        && para.borders.top.is_none()
        && para.borders.bottom.is_none()
        && para.borders.left.is_none()
        && para.borders.right.is_none()
        && para.borders.between.is_none()
        && !para.page_break_before
        && !para.page_break_after
        && para.page_break_at.is_none()
        && !para.column_break_before
        && !para.column_break_after
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

/// Where a split paragraph's final fragment goes around its trailing float.
enum FragmentPlan {
    /// Lines `from..to` sit beside the float in a narrowed box.
    Beside { from: usize, to: usize, text_x: f32, text_w: f32 },
    /// Lines from `from` on start below the float, at `top`.
    Below { from: usize, top: f32 },
    /// Below is off the page: the lines above the band keep their place
    /// (the page's first line stays even inside the band) and the rest
    /// move on, the float with them.
    NextPage { keep: usize },
    /// The anchor line alone on a page, below the float's off-page bottom:
    /// Word sets it at the page bottom above the footer.
    PageBottom,
}

/// The float's anchor follows every run with visible text, a break, a tab or
/// an inline picture: its line is the paragraph's last.
/// No run from `index` on shows anything. ponytail: the index counts the
/// runs as parsed; a run dropped since (a deleted page break) only moves it
/// past the end.
fn blank_after(runs: &[crate::model::Run], index: usize) -> bool {
    runs[index.min(runs.len())..]
        .iter()
        .all(|r| r.text.trim().is_empty() && !r.is_line_break && !r.is_tab && r.inline_image.is_none())
}

/// The floats `keep` selects, with their XObject names re-keyed to their
/// positions in the subset (`render_floating_images` keys by index).
fn float_subset(
    fis: &[crate::model::FloatingImage],
    pdf_names: &HashMap<(usize, usize), String>,
    effect_names: &HashMap<(usize, usize), images::EffectXObjs>,
    block_idx: usize,
    keep: impl Fn(&crate::model::FloatingImage) -> bool,
) -> (
    Vec<crate::model::FloatingImage>,
    HashMap<(usize, usize), String>,
    HashMap<(usize, usize), images::EffectXObjs>,
) {
    let mut sub = Vec::new();
    let mut names = HashMap::new();
    let mut effects = HashMap::new();
    for (i, fi) in fis.iter().enumerate() {
        if !keep(fi) {
            continue;
        }
        let j = sub.len();
        if let Some(n) = pdf_names.get(&(block_idx, i)) {
            names.insert((block_idx, j), n.clone());
        }
        if let Some(e) = effect_names.get(&(block_idx, i)) {
            effects.insert((block_idx, j), e.clone());
        }
        sub.push(fi.clone());
    }
    (sub, names, effects)
}

/// Whether a leading picture leaves no usable text strip on either side.
fn leading_full_width_picture(
    fi: &crate::model::FloatingImage,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> bool {
    if !fi.anchor_at_start
        || fi.v_relative_from != VRelativeFrom::Paragraph
        || !matches!(fi.v_position, VerticalPosition::Offset(y) if y > 0.0)
        || !matches!(fi.wrap_type, WrapType::Square | WrapType::Tight)
    {
        return false;
    }
    let x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
    (x - fi.dist_left - col_x).max(col_x + col_w - x - fi.image.display_width - fi.dist_right)
        < MIN_EMPTY_STRIP
}

// Full-width rectangular contour wrapping advances through candidate lines,
// rather than applying the square-wrap jump to image/effect bounds.
struct RectangularClearance {
    top: f32,
    bottom: f32,
    pitch: f32,
    line_h: f32,
}

impl RectangularClearance {
    fn cleared_y(&self, start: f32) -> f32 {
        let steps = ((start - self.bottom) / self.pitch).ceil().max(0.0);
        start - steps * self.pitch
    }
}

fn rectangular_polygon_bounds(pts: &[(f32, f32)]) -> Option<(f32, f32, f32, f32)> {
    if !(4..=5).contains(&pts.len()) {
        return None;
    }
    let left = pts.iter().map(|p| p.0).fold(f32::INFINITY, f32::min);
    let right = pts.iter().map(|p| p.0).fold(f32::NEG_INFINITY, f32::max);
    let bottom = pts.iter().map(|p| p.1).fold(f32::INFINITY, f32::min);
    let top = pts.iter().map(|p| p.1).fold(f32::NEG_INFINITY, f32::max);
    if left >= right || bottom >= top {
        return None;
    }
    for (i, &(x, y)) in pts.iter().enumerate() {
        if !(x == left || x == right) || !(y == bottom || y == top) {
            return None;
        }
        let next = pts[(i + 1) % pts.len()];
        if x != next.0 && y != next.1 {
            return None;
        }
    }
    [(left, bottom), (right, bottom), (left, top), (right, top)]
        .iter()
        .all(|c| pts.contains(c))
        .then_some((left, right, top, bottom))
}

fn rectangular_polygon_clearance(
    block: &Block,
    zone: &FloatZone,
    prev_after: f32,
    ctx: &RenderContext,
    (col_x, col_w): (f32, f32),
) -> Option<RectangularClearance> {
    if zone.from_table || !zone.para_relative || prev_after != 0.0 {
        return None;
    }
    let (left, right, top, bottom) = rectangular_polygon_bounds(zone.polygon_pts.as_ref()?)?;
    if side_strip((left, right), (zone.left_from_text, zone.right_from_text), (col_x, col_w))
        >= MIN_EMPTY_STRIP
    {
        return None;
    }
    let Block::Paragraph(p) = block else {
        return None;
    };
    if p.space_before != 0.0
        || p.space_before_auto
        || p.clears_floats
        || p.image.is_some()
        || p.inline_chart.is_some()
        || !p.smartart.is_empty()
        || !p.floating_images.is_empty()
        || !p.textboxes.is_empty()
        || p.frame_props.is_some()
        || p.runs
            .iter()
            .any(|r| r.inline_image.is_some() || r.field_code.is_some() || r.math.is_some())
    {
        return None;
    }
    let ls = p.line_spacing.unwrap_or(ctx.doc_line_spacing);
    if !matches!(ls, LineSpacing::Auto(_)) {
        return None;
    }
    let (fs, lhr, _) = tallest_run_metrics(&p.runs, ctx.fonts);
    let (fs, lhr) = if is_text_empty(&p.runs) {
        unsized_line_metrics(p, fs, ctx.fonts)
    } else {
        (fs, lhr)
    };
    let pitch = resolve_line_h(LineSpacing::Auto(1.0), fs, lhr);
    let line_h = resolve_line_h(ls, fs, lhr);
    (pitch > 0.0 && pitch.is_finite()).then_some(RectangularClearance {
        top,
        bottom,
        pitch,
        line_h,
    })
}

/// Compute effective first-line hanging indent for a paragraph.
fn compute_text_hanging(
    para: &Paragraph,
    default_tab_stop: f32,
    fonts: &HashMap<String, FontEntry>,
) -> f32 {
    list_label::text_hanging(para, default_tab_stop, fonts)
}

/// Every run of the body: its paragraphs and table cells.
fn body_runs<'a>(doc: &'a Document) -> impl Iterator<Item = &'a Run> + 'a {
    doc.sections.iter().flat_map(|s| s.blocks.iter()).flat_map(
        |block| -> Box<dyn Iterator<Item = &'a Run> + 'a> {
            match block {
                Block::Paragraph(p) => Box::new(p.runs.iter()),
                Block::Table(t) => Box::new(
                    t.rows
                        .iter()
                        .flat_map(|row| row.cells.iter())
                        .flat_map(|cell| cell.all_paragraphs())
                        .flat_map(|p| p.runs.iter()),
                ),
            }
        },
    )
}

/// The catalog `/Lang`: the language most of the body text is in (by letters,
/// East Asian ones by their own language), so only passages in another one
/// need a `/Lang` Span. The most common primary subtag wins (en-US and en-GB
/// together outvote fr-FR), then its most common tag; else the declared
/// default, else Word's en-US.
fn document_lang(doc: &Document) -> String {
    let mut letters: HashMap<&str, usize> = HashMap::new();
    for run in body_runs(doc) {
        for ch in run.text.chars().filter(|c| c.is_alphabetic()) {
            if let Some(lang) = layout::run_lang(run, layout::Script::of(ch)) {
                *letters.entry(lang).or_default() += 1;
            }
        }
    }
    let primary = |lang: &str| tagging::primary_subtag(lang).to_ascii_lowercase();
    let mut by_primary: HashMap<String, usize> = HashMap::new();
    for (lang, n) in &letters {
        *by_primary.entry(primary(lang)).or_default() += n;
    }
    // Ties go to the alphabetically first, so the output doesn't follow hash order.
    letters
        .iter()
        .max_by_key(|&(lang, n)| (by_primary[&primary(lang)], *n, std::cmp::Reverse(*lang)))
        .map(|(lang, _)| lang.to_string())
        .or_else(|| doc.default_lang.clone())
        .unwrap_or_else(|| "en-US".to_string())
}

/// A body PAGE, NUMPAGES or SECTIONPAGES field's value on page `page_idx`,
/// from the first pass (None before it). ponytail: a paragraph that runs onto
/// the next page shows its first page's number throughout.
fn body_page_field_text(
    run: &Run,
    ctx: &RenderContext,
    page_idx: usize,
    sect_idx: usize,
) -> Option<String> {
    let fields = ctx.body_page_fields?;
    match run.field_code {
        Some(FieldCode::Page(fmt)) => {
            let (num, section_fmt) = fields.pages.get(page_idx)?;
            Some(header_footer::page_field_text(
                *num,
                fmt.or(section_fmt.as_deref()),
            ))
        }
        Some(FieldCode::NumPages(fmt)) => Some(header_footer::page_field_text(fields.total, fmt)),
        Some(FieldCode::SectionPages(fmt)) => Some(header_footer::page_field_text(
            fields.section_pages[sect_idx],
            fmt,
        )),
        _ => None,
    }
}

/// Pre-compute bookmark page positions so PAGEREF fields (e.g. TOC) can
/// show correct page numbers, and each simulated page's label. Simulates
/// page layout without rendering.
fn compute_bookmark_positions(
    doc: &Document,
    ctx: &RenderContext,
) -> (HashMap<String, (usize, f32)>, Vec<String>) {
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
        return (HashMap::new(), Vec::new());
    }

    let mut bookmark_positions: HashMap<String, (usize, f32)> = HashMap::new();
    let mut page_idx = 0usize;
    // The section each simulated page opens in, shaped like
    // `PageBuilder::page_section_indices` for `page_number_formats`.
    let mut pages = vec![(0, false, 0)];
    let first_sp = &doc.sections[0].properties;
    let mut sp = first_sp;
    let mut slot_top = effective_slot_top(sp, true, page_idx, ctx);
    let mut margin_bottom = compute_effective_margin_bottom(sp, true, page_idx, ctx);
    let mut prev_space_after: f32 = 0.0;
    let mut prev_para: Option<&Paragraph> = None;

    for (si, section) in doc.sections.iter().enumerate() {
        sp = &section.properties;
        if si > 0 {
            match sp.break_type {
                SectionBreakType::NextPage
                | SectionBreakType::OddPage
                | SectionBreakType::EvenPage => {
                    page_idx += 1;
                    pages.push((si, false, si));
                    slot_top = effective_slot_top(sp, true, page_idx, ctx);
                    margin_bottom = compute_effective_margin_bottom(sp, true, page_idx, ctx);
                    prev_space_after = 0.0;
                }
                SectionBreakType::Continuous => {}
            }
        }
        let text_width = sp.text_width();
        let blocks = &section.blocks;
        for (bi, block) in blocks.iter().enumerate() {
            match block {
                Block::Paragraph(para) => {
                    if para.page_break_before
                        && slot_top < effective_slot_top(sp, false, page_idx, ctx)
                    {
                        page_idx += 1;
                        pages.push((si, false, si));
                        slot_top = effective_slot_top(sp, false, page_idx, ctx);
                        margin_bottom = compute_effective_margin_bottom(sp, false, page_idx, ctx);
                        prev_space_after = 0.0;
                    }
                    for bm in &para.bookmarks {
                        bookmark_positions.insert(bm.clone(), (page_idx, slot_top));
                    }
                    if is_empty_wrapping_frame(para) {
                        continue;
                    }
                    let next_continuous = doc.sections.get(si + 1).is_some_and(|next| {
                        next.properties.break_type == SectionBreakType::Continuous
                    });
                    let keeps_line = keeps_section_mark_line(
                        ctx.compat_mode,
                        next_continuous,
                        bi.checked_sub(1).and_then(|i| blocks.get(i)),
                    );
                    if para.is_section_break && bi != 0 && !keeps_line && is_text_empty(&para.runs)
                    {
                        let drop;
                        (drop, prev_space_after) = section_break_spacing(
                            prev_space_after,
                            para.space_after,
                            next_continuous,
                        );
                        slot_top -= drop;
                        continue;
                    }
                    let (mut font_size, mut tallest_lhr, _) =
                        tallest_glyph_run_metrics(&para.runs, ctx.fonts);
                    if tallest_lhr.is_none() {
                        (font_size, tallest_lhr) = unsized_line_metrics(para, font_size, ctx.fonts);
                    }
                    let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                    let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
                    let line_h = if para.snap_to_grid
                        && sp.line_grid_pitch().is_some()
                        && !matches!(effective_ls, LineSpacing::Exact(_))
                    {
                        grid_snapped_line_h(
                            &para.runs,
                            ctx.fonts,
                            effective_ls,
                            line_h,
                            sp.line_pitch,
                        )
                    } else {
                        line_h
                    };
                    let para_w = (text_width - para.indent_left - para.indent_right).max(1.0);
                    let hanging = compute_text_hanging(para, ctx.default_tab_stop, ctx.fonts);
                    let lines = if is_text_empty(&para.runs) {
                        vec![]
                    } else {
                        build_lines(
                            &para.runs,
                            ctx,
                            para_w,
                            ctx.cjk(para.auto_space_de || para.auto_space_dn, para),
                            &LineOpts {
                                tab_stops: &para.tab_stops,
                                indent_left: para.indent_left,
                                indent_right: para.indent_right,
                                hanging,
                                ..Default::default()
                            },
                        )
                    };
                    let num_lines = lines.len().max(1);
                    let para_has_inline_img = para.runs.iter().any(|r| r.inline_image.is_some());
                    let content_h = if para.image.is_some() || para.inline_chart.is_some() {
                        para.content_height
                    } else if para_has_inline_img && para.content_height > 0.0 {
                        para.content_height.max(num_lines as f32 * line_h)
                    } else {
                        num_lines as f32 * line_h
                    };
                    let effective_sb = helpers::effective_space_before(para, prev_para);
                    let next_para = match blocks.get(bi + 1) {
                        Some(Block::Paragraph(p)) => Some(p),
                        Some(Block::Table(t)) => first_cell_paragraph(t),
                        None => None,
                    };
                    let effective_sa = helpers::effective_space_after(para, next_para);
                    let inter_gap = f32::max(prev_space_after, effective_sb);
                    let needed = inter_gap + content_h;
                    if slot_top - needed < margin_bottom
                        && slot_top < effective_slot_top(sp, false, page_idx, ctx)
                    {
                        page_idx += 1;
                        pages.push((si, false, si));
                        slot_top = effective_slot_top(sp, false, page_idx, ctx);
                        margin_bottom = compute_effective_margin_bottom(sp, false, page_idx, ctx);
                        slot_top -= content_h;
                    } else {
                        slot_top -= inter_gap + content_h;
                    }
                    prev_space_after = effective_sa;
                    prev_para = Some(para);
                }
                Block::Table(table) => {
                    let cells: Vec<Vec<&Paragraph>> = table
                        .rows
                        .iter()
                        .flat_map(|r| r.cells.iter())
                        .map(|c| c.all_paragraphs())
                        .collect();
                    let para_count =
                        cells.iter().map(|c| c.len().max(1)).max().unwrap_or(1) * table.rows.len();
                    let est_h = para_count as f32 * 14.0;
                    if slot_top - est_h < margin_bottom {
                        page_idx += 1;
                        pages.push((si, false, si));
                        slot_top = effective_slot_top(sp, false, page_idx, ctx);
                        margin_bottom = compute_effective_margin_bottom(sp, false, page_idx, ctx);
                    }
                    // Captions in cells count as on the table's page (issue
                    // #146: a panel table's TOC entries had no page numbers).
                    // ponytail: a table that splits keeps its first page.
                    for bm in cells.iter().flatten().flat_map(|p| &p.bookmarks) {
                        bookmark_positions.insert(bm.clone(), (page_idx, slot_top));
                    }
                    slot_top -= est_h;
                    prev_space_after = 0.0;
                    prev_para = None;
                }
            }
        }
    }
    // A page's label as its PAGE field shows it (issue #146: a TOC after
    // "i" front matter shows 1, not the sheet count).
    let labels = page_number_formats(doc, &pages)
        .into_iter()
        .map(|(n, fmt)| header_footer::page_field_text(n, fmt))
        .collect();
    (bookmark_positions, labels)
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

/// How far outside its `space` a paragraph's left or right border sits.
const LEFT_BORDER_GAP: f32 = 1.47;
const RIGHT_BORDER_GAP: f32 = 1.73;

/// How far a top or bottom border runs past its box at each side: to a side
/// border's outer edge, or without one by the side gaps Word keeps anyway
/// (paragraph-border-styles: 70.5–541.7 on a 72–540 measure).
pub(super) fn horizontal_border_ext(borders: &crate::model::ParagraphBorders) -> (f32, f32) {
    let ext = |b: &Option<ParagraphBorder>, gap: f32| b.as_ref().map_or(gap, |b| b.extent() / 2.0);
    (
        ext(&borders.left, LEFT_BORDER_GAP),
        ext(&borders.right, RIGHT_BORDER_GAP),
    )
}

/// How far outside the text the left and right borders' inner edges sit
/// (`space` plus the gap), or their stroke centres with `to_centre`.
pub(super) fn side_border_outsets(
    borders: &crate::model::ParagraphBorders,
    to_centre: bool,
) -> (f32, f32) {
    let outset = |b: Option<&crate::model::ParagraphBorder>, gap: f32| {
        b.map_or(0.0, |b| {
            b.space_pt + gap + if to_centre { b.extent() / 2.0 } else { 0.0 }
        })
    };
    (
        outset(borders.left.as_ref(), LEFT_BORDER_GAP),
        outset(borders.right.as_ref(), RIGHT_BORDER_GAP),
    )
}

/// An empty section mark keeps its line after an inline table (the rule is
/// explained in `render_paragraph_block`); bookmark positions follow it too.
fn keeps_section_mark_line(compat_mode: u32, next_continuous: bool, prev: Option<&Block>) -> bool {
    compat_mode >= 15
        && next_continuous
        && matches!(prev, Some(Block::Table(t)) if t.position.is_none())
}

#[allow(clippy::too_many_arguments)]
fn render_para_textbox(
    tb: &crate::model::Textbox,
    para: &Paragraph,
    anchor: ParaAnchor,
    sp: &SectionProperties,
    (col_x, col_w): (f32, f32),
    text_width: f32,
    ctx: &RenderContext,
    pb: &mut PageBuilder,
) {
    let tb_x = col_x + if tb.on_line { para.indent_left } else { 0.0 };
    let page = pb.all_contents.len();
    // Each box renders into its own buffer; flush_page paints these behind or
    // above the page's text layer sorted by relativeHeight (Word's z-order for
    // floating shapes spans paragraphs).
    let mut shape_content = tagging::artifact_content();
    render_single_textbox(
        tb,
        sp,
        tb_x,
        col_w,
        text_width,
        anchor.top(tb.on_line),
        &mut shape_content,
        &mut pb.gradient_specs,
        ctx,
        &mut pb.links,
        &mut pb.tags,
        page,
        &mut pb.alpha_states,
    );
    let shapes = if tb.behind_doc { &mut pb.behind_shapes } else { &mut pb.deferred_shapes };
    shapes.push((tb.z_index, shape_content));
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
    doc: &Document,
    smartart_font_key: &str,
    smartart_image_names: &HashMap<usize, String>,
    debug_wrap: bool,
    table_connector_anchor: Option<(usize, f32)>,
) -> bool {
    let prev_between_below = std::mem::take(&mut state.prev_between_below);
    let reflowed = std::mem::take(&mut state.reflowed);
    // One wholly beside the text keeps its (empty) paragraph, which Word
    // tags (empty-margin-frame-off-page).
    if is_empty_wrapping_frame(para)
        && !para.frame_props.as_ref().is_some_and(|fp| frame_beside_text(fp, sp))
    {
        state.global_block_idx += 1;
        return true;
    }

    // §17.6.8: per-section line-number config (None if disabled). Holds no borrow
    // of `state`, so each render call can freshly borrow the shared counter.
    // Word numbers neither frame text (a drop cap's letter) nor suppressed lines.
    let numbered = para.frame_props.is_none() && !para.suppress_line_numbers;
    let ln_cfg: Option<(i32, u32, f32)> =
        sp.line_numbering.as_ref().filter(|_| numbered).map(|ln| {
            (
                ln.start,
                ln.count_by,
                sp.margin_left - ln.distance.unwrap_or(18.0),
            )
        });
    let ln_font: Option<(String, f32)> =
        sp.line_numbering
            .as_ref()
            .filter(|_| numbered)
            .and_then(|ln| {
                ctx.fonts
                    .get(&crate::fonts::font_key_for(&ln.font_name, false, false))
                    .map(|e| (e.pdf_name.clone(), ln.font_size))
            });
    if para.frame_props.is_none() && para.runs.iter().any(|run| !run.text.trim().is_empty()) {
        if let Some(bottom) = state.pb.next_text_below_picture.take() {
            state.pb.slot_top = state.pb.slot_top.min(bottom);
        }
    }
    let adjacent_para = |idx: usize| block_para(section_blocks, idx);

    // Skip empty section-break paragraphs — Word gives these zero height, also
    // before a continuous section that changes the columns (Word probes in
    // compat 14 and 15; covid_insomnia's two columns start 12pt higher) —
    // unless the paragraph is the first block of its section, including a
    // continuous empty section:
    // transition_to_work's contents start a line and 8pt below the top of the
    // page that empty section opens.
    let next_continuous = doc
        .sections
        .get(sect_idx + 1)
        .is_some_and(|next| next.properties.break_type == SectionBreakType::Continuous);
    // In modern compatibility mode, Word preserves the section mark's own
    // paragraph line after an inline table. Floating tables do not reserve
    // that extra line. Legacy mode still collapses it for both kinds.
    // Only before a continuous section: before a new page the line could only
    // push itself onto a blank page, which Word doesn't do (indigenous_innovation
    // ends a page with its signature table and starts Schedule A on the next).
    let keeps_line = block_idx == 0
        || keeps_section_mark_line(
            ctx.compat_mode,
            next_continuous,
            block_idx.checked_sub(1).and_then(|i| section_blocks.get(i)),
        )
        // A modern continuous section mark can occupy a side strip beside
        // a floating table. Full-width floats still absorb that mark line.
        || (ctx.compat_mode >= 15 && next_continuous
            && block_idx.checked_sub(1).and_then(|i| section_blocks.get(i))
                .is_some_and(|b| matches!(b, Block::Table(t) if t.position.is_some()))
            && state.pb.float_zone.as_ref().is_some_and(|zone| {
                let (x, w) = col_geometry[state.current_col];
                (zone.obj_left - zone.left_from_text - x)
                    .max(x + w - zone.obj_right - zone.right_from_text) >= MIN_EMPTY_STRIP
            }));

    if para.is_section_break
        && !keeps_line
        && is_text_empty(&para.runs)
        && para.image.is_none()
        && para.inline_chart.is_none()
        && para.smartart.is_empty()
        && para.floating_images.is_empty()
        && para.textboxes.is_empty()
    {
        // Its space after still meets the next section's first space before:
        // case25's sections (break paragraph after=10pt) start their 24pt
        // heading 14pt down, victorian's (after=0) its 26pt heading 26pt down.
        // At a page top nothing moves (both this branch's and main's probes,
        // 2026-10-05: gap = prev after + max(0, next before − break after)).
        let drop;
        (drop, state.prev_space_after) =
            section_break_spacing(state.prev_space_after, para.space_after, next_continuous);
        if !state.pb.is_at_page_top(sp) {
            state.pb.slot_top -= drop;
        }
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
        // The break still sits on its paragraph's first line, and that line
        // must fit first: one that does not moves to a page of its own before
        // breaking, leaving that page blank (Word probes p29a–c: a 25pt exact
        // break line under 23pt of room makes a blank page, a 15pt one does
        // not). The space above it may run past the page bottom (erasmus_plus:
        // a 17pt line with 22pt left under 12pt of space after stays put).
        if !at_top && para.page_break_before_explicit && {
            let (font_size, mut lhr, _) = tallest_glyph_run_metrics(&para.runs, ctx.fonts);
            let mut font_size = font_size;
            if lhr.is_none() {
                (font_size, lhr) = unsized_line_metrics(para, font_size, ctx.fonts);
            }
            let ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
            state.pb.slot_top - resolve_line_h(ls, font_size, lhr) < state.effective_margin_bottom
        } {
            state
                .pb
                .begin_next_page(sect_idx, sp, &mut state.effective_margin_bottom, ctx);
        }
        if !at_top || para.page_break_before_explicit {
            state
                .pb
                .begin_next_page(sect_idx, sp, &mut state.effective_margin_bottom, ctx);
            state.current_col = 0;
        }
        state.prev_space_after = 0.0;
        let has_floats = !para.floating_images.is_empty()
            || !para.textboxes.is_empty()
            || !para.connectors.is_empty()
            || !para.smartart.is_empty();
        if is_text_empty(&para.runs) && !has_floats {
            state.global_block_idx += 1;
            return true;
        }
    }

    // Handle explicit column breaks
    // In a one-column section Word breaks the page (bosch's page 2 ends there).
    if para.column_break_before {
        // The column ends below the space after its last paragraph (its
        // separator reaches there, Word probe).
        if col_count > 1 {
            state.pb.slot_top -= state.prev_space_after;
        }
        state.pb.advance_column_or_page(
            &mut state.current_col,
            col_count,
            sect_idx,
            sp,
            &mut state.effective_margin_bottom,
            ctx,
        );
        state.prev_space_after = 0.0;
    }

    let next_para = adjacent_para(block_idx + 1);
    let prev_para = if block_idx > 0 {
        adjacent_para(block_idx - 1)
    } else {
        None
    };

    // A continuous section boundary does not interrupt contextual spacing.
    // Border grouping still uses the section-local neighbours above.
    let context_prev = prev_para.or_else(|| {
        (block_idx == 0 && sp.break_type == SectionBreakType::Continuous)
            .then(|| sect_idx.checked_sub(1).and_then(|i| doc.sections.get(i)))
            .flatten()
            .and_then(|section| section.blocks.last())
            .and_then(leading_para)
    });
    let context_next = next_para.or_else(|| {
        (block_idx + 1 == section_blocks.len())
            .then(|| doc.sections.get(sect_idx + 1))
            .flatten()
            .filter(|section| section.properties.break_type == SectionBreakType::Continuous)
            .and_then(|section| section.blocks.first())
            .and_then(leading_para)
    });

    // A leading floating table is already the first body block. Word drops
    // the following empty anchor's before-spacing at the page top, while a
    // document's ordinary first paragraph retains its own before-spacing.
    let empty_anchor_after_float = is_text_empty(&para.runs)
        && state.pb.is_at_page_top(sp)
        && block_idx.checked_sub(1).and_then(|i| section_blocks.get(i))
            .is_some_and(|b| matches!(b, Block::Table(t) if t.position.is_some()));
    let effective_space_before = if empty_anchor_after_float {
        0.0
    } else {
        helpers::effective_space_before(para, context_prev)
    };
    let effective_space_after = helpers::effective_space_after(para, context_next);

    // Word overlaps the space after and the next space before (the larger
    // wins) under its HTML spacing rule; doNotUseHTMLParagraphAutoSpacing
    // adds them (5pt + 5pt list items step 23.76 on a 13.8pt line).
    let mut inter_gap = if reflowed {
        0.0
    } else if ctx.additive_paragraph_spacing {
        state.prev_space_after + effective_space_before
    } else {
        f32::max(state.prev_space_after, effective_space_before)
    };
    // The anchor the wrap and reservation passes below see, before a page
    // break or a page top's suppressed space before is known (peeked here,
    // taken when the drawings render).
    let pre_anchor = ParaAnchor::new(
        state.pb.float_anchor_for(block_idx),
        state.pb.anchor_slot_top(sp, ctx) - inter_gap,
        own_space_before(inter_gap, state.prev_space_after),
        sp,
        ctx,
    );

    let (mut font_size, mut tallest_lhr, mut tallest_ar) =
        tallest_glyph_run_metrics(&para.runs, ctx.fonts);
    // A line of only tabs and spaces sits on the mark's (or break's) font,
    // its ascent included: an underlined tab hangs from that baseline.
    if tallest_lhr.is_none() {
        (font_size, (tallest_lhr, tallest_ar)) = unsized_line_font(para, font_size, ctx.fonts);
    }
    let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
    let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
    let grid_snapped = para.snap_to_grid
        && sp.line_grid_pitch().is_some()
        && !matches!(effective_ls, LineSpacing::Exact(_));
    let line_h = if grid_snapped {
        grid_snapped_line_h(&para.runs, ctx.fonts, effective_ls, line_h, sp.line_pitch)
    } else {
        line_h
    };
    let grid_baseline = grid_snapped
        .then(|| grid_baseline_offset(&para.runs, ctx.fonts, line_h))
        .flatten()
        .unwrap_or(sp.line_pitch);

    // Word bottom-aligns text within an exact-height line box: the baseline
    // sits winDescent above the box bottom (identity: line_h_ratio −
    // ascender_ratio = winDescent/upm), however large the font's ascent is.
    // Placing the baseline at font_size * ascender_ratio instead pushes a
    // large-lineGap CJK substitute's descenders out of the fixed box and
    // into whatever follows (annotation #219: heading into table border).
    let exact_baseline_base = layout::boxed_line_ascent(
        effective_ls,
        line_h,
        font_size,
        tallest_lhr,
        tallest_ar,
        &para.runs,
        ctx.fonts,
    );

    let follows_floating_table = block_idx
        .checked_sub(1)
        .and_then(|i| section_blocks.get(i))
        .is_some_and(|b| matches!(b, Block::Table(t) if t.position.is_some()));
    // A single leading clearing break keeps its empty anchor line beside the
    // immediately preceding floating table, but its visible text resumes below.
    // Clearing only after painting the paragraph misplaces that first text line.
    let leading_table_clear = if para.clears_floats
        && follows_floating_table
        && matches!(para.runs.as_slice(), [first, rest @ ..]
            if first.is_line_break && !rest.iter().any(|r| r.is_line_break))
        && para
            .runs
            .iter()
            .any(|r| !r.text.trim().is_empty() || r.inline_image.is_some())
        && let Some(zone) = state
            .pb
            .float_zone
            .take_if(|zone| zone.from_table && state.pb.slot_top > zone.bottom_y)
    {
        state.pb.slot_top = zone.bottom_y + inter_gap;
        true
    } else {
        false
    };

    let (col_x, col_w) = col_geometry[state.current_col];
    let mut para_text_x = col_x + para.indent_left;
    let mut para_text_width = (col_w - para.indent_left - para.indent_right).max(1.0);
    let mut label_x = list_label::label_x(para, col_x, col_w, ctx.fonts);

    // When inside a floating object zone, narrow the paragraph to
    // fit beside the object rather than overlapping it.
    // The paragraph's first line starts below the gap: a centred line under
    // a logo sits clear of it once the 8pt after-space is counted.
    let first_line_top = state.pb.slot_top - inter_gap;
    if let Some(ref fz) = state.pb.float_zone {
        fz.narrow_paragraph(
            first_line_top,
            col_x,
            col_w,
            para,
            &mut para_text_x,
            &mut para_text_width,
            &mut label_x,
            ctx.compat_mode < 15,
        );
    }

    let text_hanging = compute_text_hanging(para, ctx.default_tab_stop, ctx.fonts);

    // Substitute footnote/endnote refs and resolve PAGEREF fields
    let has_footnote_refs = para.runs.iter().any(|r| r.footnote_id.is_some());
    let has_endnote_refs = para.runs.iter().any(|r| r.endnote_id.is_some());
    let has_pageref = para.runs.iter().any(|r| {
        r.field_code
            .as_ref()
            .is_some_and(|fc| fc.is_page_count() || matches!(fc, FieldCode::PageRef(_)))
    });
    let effective_runs: std::borrow::Cow<'_, Vec<Run>> =
        if has_footnote_refs || has_endnote_refs || has_pageref {
            let substituted: Vec<Run> = para
                .runs
                .iter()
                .map(|run| {
                    if let Some(text) = ctx.note_mark_text(run) {
                        Run {
                            text,
                            ..run.clone()
                        }
                    } else if let Some(FieldCode::PageRef(ref bookmark)) = run.field_code {
                        let mut r = run.clone();
                        // Word prints PAGEREF's cached result unless fields are updated
                        // before printing; only an empty result needs our estimate.
                        if r.text.trim().is_empty()
                            && let Some(&(page_idx, _)) = state.bookmark_positions.get(bookmark)
                        {
                            r.text = ctx
                                .page_labels
                                .get(page_idx)
                                .cloned()
                                .unwrap_or_else(|| (page_idx + 1).to_string());
                        }
                        r
                    } else if let Some(text) =
                        body_page_field_text(run, ctx, state.pb.all_contents.len(), sect_idx)
                    {
                        Run {
                            text,
                            field_code: None,
                            ..run.clone()
                        }
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
    // A float anchored after the paragraph's text belongs to the page its
    // last line lands on (the split path draws it there); when the forced
    // breaks alone outrun this page the paragraph will split, so that float
    // wraps none of the lines laid out here (an external tracked corpus
    // cover: five pages of break-separated lines, a page-wide picture after
    // them, set one letter per line beside a picture Word prints on page 5).
    let forced_lines = para.runs.iter().filter(|r| r.is_line_break).count() as f32 + 1.0;
    let splits_for_sure = forced_lines * line_h > first_line_top - state.effective_margin_bottom;
    let wraps_here = |fi: &crate::model::FloatingImage| {
        !(splits_for_sure && blank_after(&para.runs, fi.anchor_run))
    };
    // Self-wrapping: if this paragraph anchors a wrapping float
    // and has text, set up the float zone NOW so width-narrowing
    // applies to this paragraph's own lines.  Always replace any
    // previous float zone — the paragraph's own image takes priority.
    let mut own_textbox_zones: Vec<FloatZone> = if text_empty {
        Vec::new()
    } else {
        textbox_zones(para, sp, (col_x, col_w), text_width, pre_anchor.top(false), ctx)
    };
    let own_picture = para
        .floating_images
        .iter()
        .filter(|fi| !text_empty && wraps_here(fi))
        .find(|fi| wraps_in_column(fi, sp, col_x, col_w, text_width));
    // The previous paragraph may have re-wrapped around this float and
    // grown; the float stays where that look-ahead anchored it
    // (peeked here, taken below).
    let own_zone = match own_picture {
        Some(fi) => Some(FloatZone::for_float(
            fi,
            resolve_fi_x(fi, sp, col_x, col_w, text_width),
            pre_anchor.fi_y_top(fi, sp),
        )),
        None if !own_textbox_zones.is_empty() => Some(own_textbox_zones.remove(0)),
        None => None,
    };
    let own_zone_set = own_zone.is_some();
    let entry_zones = (state.pb.float_zone.clone(), state.pb.more_float_zones.clone());
    if let Some(zone) = own_zone {
        state.pb.float_zone = Some(zone);
        // Re-narrow para_text_x / para_text_width using the
        // new float zone (same logic as the block above).
        let fz = state.pb.float_zone.as_ref().unwrap();
        fz.narrow_paragraph(
            first_line_top,
            col_x,
            col_w,
            para,
            &mut para_text_x,
            &mut para_text_width,
            &mut label_x,
            ctx.compat_mode < 15,
        );
    }

    let cjk = ctx.cjk(para.auto_space_de || para.auto_space_dn, para);

    // Look-ahead: a wrapping float anchored in the *next* block (an image-only
    // paragraph) sits at that block's top, which Word computes from this
    // paragraph laid out at full width — and then re-wraps this paragraph's
    // lines around the float without moving it (case41 p3: the paragraph before
    // a centred 4.5in picture wraps beside it from its second line, annotation
    // #152). Install that zone now so the per-line geometry below narrows the
    // lines it reaches, and hand the anchor position to the next paragraph so
    // it draws the picture there rather than where it now flows.
    // (anchor paragraph top, the zone's real top edge) once installed.
    let mut lookahead: Option<(f32, f32, usize)> = None;
    if state.pb.float_zone.is_none() && !text_empty && !has_tabs {
        // The anchor paragraph may carry text of its own (case41 p6, annotation
        // #240): the anchor is its top either way. It may also sit up to three
        // paragraphs ahead: Word places the float from the paragraphs between
        // laid out at full width, then re-wraps them all beside it (Word for
        // Mac probes p26a and p26b: two short titles before the anchor both
        // wrap beside a page-tall picture at y 0 and beside a 300pt one
        // hoisted 60pt above the third paragraph). ponytail: plain paragraphs
        // between only, three at most.
        let qualifies = |fi: &crate::model::FloatingImage| {
            // Only floats that hang off the anchor paragraph itself;
            // a page- or margin-relative float does not move with it.
            fi.v_relative_from == VRelativeFrom::Paragraph
                && matches!(
                    fi.v_position,
                    VerticalPosition::Offset(_) | VerticalPosition::AlignTop
                )
                && wraps_in_column(fi, sp, col_x, col_w, text_width)
                // Only a float text can sit beside: Word wraps the
                // preceding paragraph next to case41 p3's centred
                // picture (64.8pt free on each side) but leaves
                // brazilian p9's caption alone above a figure with
                // 37.5pt beside it. ponytail: 48pt threshold, two
                // calibration points.
                && positioning::side_gaps(fi, sp, col_x, col_w, text_width)
                    .is_some_and(|(left, right)| left.max(right) >= 48.0)
        };
        // Find the anchor first: the paragraphs between are only measured when
        // there is one, so a plain paragraph costs no line breaking ahead.
        let mut found = None;
        let mut gaps = [0.0f32; 3];
        let mut prev_after = para.space_after;
        for k in 1..=3 {
            let Some(Block::Paragraph(np)) = section_blocks.get(block_idx + k) else {
                break;
            };
            if np.image.is_some() || np.inline_chart.is_some() || np.page_break_before {
                break;
            }
            gaps[k - 1] = if ctx.additive_paragraph_spacing {
                prev_after + np.space_before
            } else {
                prev_after.max(np.space_before)
            };
            if let Some(fi) = np.floating_images.iter().find(|fi| qualifies(fi)) {
                found = Some((fi, k));
                break;
            }
            if !np.floating_images.is_empty() || !np.textboxes.is_empty() || np.page_break_after {
                break;
            }
            prev_after = np.space_after;
        }
        let first_gap = gaps[0];
        let found = found.map(|(fi, k)| {
            let mut between = 0.0f32;
            for j in 1..k {
                let Some(Block::Paragraph(np)) = section_blocks.get(block_idx + j) else {
                    unreachable!("the scan above only passes paragraphs");
                };
                between +=
                    gaps[j - 1] + line_count(np, ctx, col_w).max(1) as f32 * plain_line_h(np, ctx);
            }
            (fi, between + gaps[k - 1], block_idx + k)
        });
        if let Some((fi, gap, anchor_block)) = found {
            let full_lines = build_lines(
                &effective_runs,
                ctx,
                para_text_width,
                cjk,
                &LineOpts {
                    inline_images: Some(&block_inline_images),
                    effects: Some(&block_effect_inlines),
                    hanging: text_hanging,
                    ..Default::default()
                },
            );
            let anchor_top = state.pb.slot_top - inter_gap - full_lines.len() as f32 * line_h - gap;
            let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
            let fi_y_top = match fi.v_position {
                VerticalPosition::Offset(o) => anchor_top - o,
                _ => anchor_top,
            };
            let mut zone = FloatZone::for_float(fi, fi_x, fi_y_top);
            // Only when the float's top and its anchor land on this page: a
            // paragraph that breaks before its anchor would hand the next
            // page a stale anchor. (A page-tall picture's foot lies below the
            // margin and still wraps this page: p26a.)
            if zone.top_y > state.effective_margin_bottom
                && anchor_top > state.effective_margin_bottom
            {
                let true_top = zone.top_y;
                // Word treats the last line's space-after as part of that line
                // when testing overlap, so the zone reaches up through the gap
                // for this paragraph's geometry (restored after the lines are
                // built).
                zone.top_y += first_gap;
                state.pb.float_zone = Some(zone);
                lookahead = Some((anchor_top, true_top, anchor_block));
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
            .filter(|fi| fi.wrap_type.wraps_beside() && wraps_here(fi))
            .filter(|fi| !image_blocks_band(fi, sp, col_x, col_w, text_width))
            .skip(usize::from(own_picture.is_some()))
            .map(|fi| {
                let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
                let fi_y_top = pre_anchor.fi_y_top(fi, sp);
                FloatZone::for_float(fi, fi_x, fi_y_top)
            })
            .chain(own_textbox_zones)
            .chain(if own_zone_set { Vec::new() } else { state.pb.more_float_zones.clone() })
            .collect()
    } else {
        Vec::new()
    };

    // Lines beside floats with no room on a side text may take (floating
    // tables keep their own rule).
    let mut blocked_lines = vec![false; 0];
    // Build per-line geometry and dual-region geometry
    let (mut poly_line_geom, mut poly_dual_geom): (Option<Vec<(f32, f32)>>, Option<Vec<DualRegion>>) =
        if let Some(fz) = state.pb.float_zone.as_ref() {
            let eff_top = state.pb.slot_top - inter_gap;
            if !extra_float_zones.is_empty() {
                // Multiple wrapping floats on one paragraph: subtract every
                // float's exclusion span per line, keep to the sides each
                // float's wrapText allows, and lay the text in the gaps left,
                // the first two from the left as a dual region (Word example:
                // between two boxes and right of the second).
                let zones: Vec<&FloatZone> = std::iter::once(fz)
                    .chain(extra_float_zones.iter())
                    .collect();
                if zones.iter().all(|z| eff_top <= z.bottom_y) {
                    (None, None)
                } else {
                    let full_w = (col_w - para.indent_left - para.indent_right).max(1.0);
                    let col_right = col_x + col_w;
                    let lowest_bottom = zones
                        .iter()
                        .map(|z| z.bottom_y)
                        .fold(f32::INFINITY, f32::min);
                    let max_lines =
                        ((((eff_top - lowest_bottom) / line_h).ceil() as usize) + 5).max(50);
                    let mut geom = Vec::with_capacity(max_lines);
                    let mut dual: Vec<DualRegion> = Vec::with_capacity(max_lines);
                    let mut two_regions = false;
                    let legacy = ctx.compat_mode < 15;
                    for i in 0..max_lines {
                        let line_top = eff_top - i as f32 * line_h;
                        let line_bottom = line_top - line_h;
                        let para_left = col_x + para.indent_left;
                        let mut intervals: Vec<(f32, f32)> =
                            vec![(para_left, col_right - para.indent_right)];
                        let mut float_edges: Vec<f32> = Vec::new();
                        let mut overlapped = false;
                        for z in &zones {
                            let partial_overlap = z.para_relative || z.polygon_pts.is_some();
                            // Require the line to dip >20% of its height below
                            // the zone top before counting as in-zone, matching
                            // the symmetric bottom_threshold. Without this, a
                            // caption line directly above a float offset down by
                            // ~a line height got flagged in-zone and wrapped to
                            // one word per line (annotation #120).
                            let in_zone = if partial_overlap {
                                line_bottom < z.top_y - line_h * 0.2
                            } else {
                                line_top <= z.top_y
                            };
                            if !(in_zone && line_top > z.bottom_y + line_h * 0.2) {
                                continue;
                            }
                            let (ex_left, ex_right) =
                                z.exclusion_in_band(line_top.min(z.top_y), line_bottom);
                            let sl = ex_left - z.left_from_text;
                            let sr = ex_right + z.right_from_text;
                            float_edges.push(sr);
                            overlapped |= !z.from_table;
                            // Text keeps to the sides the float's wrapText allows.
                            let side = (z.wrap_text != WrapText::BothSides)
                                .then(|| wrap_side(z.wrap_text, sl - col_x, col_right - sr));
                            let (lo, hi) = match side {
                                None => (f32::MIN, f32::MAX),
                                Some(Some(true)) => (sr, f32::MAX),
                                Some(Some(false)) => (f32::MIN, sl),
                                Some(None) => (0.0, 0.0),
                            };
                            intervals = intervals
                                .into_iter()
                                .flat_map(|(a, b)| [(a, b.min(sl)), (a.max(sr), b)])
                                .map(|(a, b)| (a.max(lo), b.min(hi)))
                                .filter(|(a, b)| b > a)
                                .collect();
                        }
                        let min_w = if overlapped { MIN_EMPTY_STRIP } else { 1.0 };
                        let open: Vec<(f32, f32)> =
                            intervals.into_iter().filter(|(a, b)| b - a >= min_w).collect();
                        let region = |(a, b): (f32, f32)| {
                            let edge = float_edges
                                .iter()
                                .copied()
                                .filter(|&e| e <= a + 0.01)
                                .reduce(f32::max);
                            let a = match edge {
                                Some(e) if i == 0 => {
                                    first_line_region_x(a, e, para, text_hanging, legacy)
                                }
                                _ => a,
                            };
                            (a, (b - a).max(1.0))
                        };
                        let pick = match open.len() {
                            n if para.bidi && n > 2 => &open[n - 2..],
                            _ => &open[..],
                        };
                        let (x, w, rx, rw) = match pick {
                            [] => {
                                blocked_lines.resize(i, false);
                                blocked_lines.push(overlapped);
                                (para_left, full_w, 0.0, 0.0)
                            }
                            [one] => {
                                let (x, w) = region(*one);
                                (x, w, 0.0, 0.0)
                            }
                            [first, second, ..] => {
                                two_regions = true;
                                let ((x, w), (rx, rw)) = (region(*first), region(*second));
                                (x, w, rx, rw)
                            }
                        };
                        geom.push((x, w));
                        dual.push((x, w, rx, rw));
                    }
                    (Some(geom), two_regions.then_some(dual))
                }
            } else if eff_top <= fz.bottom_y {
                (None, None)
            } else {
                let is_both_sides = fz.wrap_text == WrapText::BothSides;
                let full_w = (col_w - para.indent_left - para.indent_right).max(1.0);
                let col_right = col_x + col_w;
                let max_lines = ((eff_top - fz.bottom_y) / line_h).ceil() as usize + 5;
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
                // A line starting just above a float can still intersect it.
                // Apply the same overlap threshold to rectangular table zones
                // as to paragraph-relative images and polygon zones.
                for i in 0..max_lines {
                    let line_top = eff_top - i as f32 * line_h;
                    let line_bottom = line_top - line_h;
                    // See annotation #120: require >20% line-height overlap
                    // (symmetric with bottom_threshold) so a caption line
                    // directly above a downward-offset float is not wrongly
                    // squeezed into the float's side margin.
                    let in_zone = line_bottom < fz.top_y - line_h * 0.2;
                    if in_zone && line_top > bottom_threshold {
                        let (ex_left, ex_right) =
                            fz.exclusion_in_band(line_top.min(fz.top_y), line_bottom);
                        let float_left = ex_left - fz.left_from_text;
                        let float_right = ex_right + fz.right_from_text;
                        let sr = col_right - float_right;
                        let sl = float_left - col_x;
                        let (lx, mut lw) = left_text_region(col_x, col_right, float_left, para);
                        let mut right_of_float =
                            right_text_region(col_x, col_right, float_right, para);
                        if i == 0 {
                            let (x, w) = right_of_float;
                            let legacy = ctx.compat_mode < 15;
                            let first = first_line_region_x(x, float_right, para, text_hanging, legacy);
                            right_of_float = (first, (w - (first - x)).max(0.0));
                            // The renderer widens the first line leftwards by a
                            // hanging indent: not past a float's edge.
                            if text_hanging > 0.0 && lw == 0.0 {
                                lw = (float_left - lx).max(-text_hanging);
                            }
                        }

                        if is_both_sides {
                            // A strip under the minimum stays empty on its own
                            // (Word example: 17pt left of the box, text right only).
                            let narrow = |room: f32| !fz.from_table && room < MIN_EMPTY_STRIP;
                            if narrow(sl) && narrow(sr) {
                                blocked_lines.resize(i, false);
                                blocked_lines.push(true);
                            }
                            if narrow(sl) {
                                lw = lw.min(0.0);
                            }
                            // BothSides: provide both regions
                            let (rx, rw) = right_of_float;
                            let rw = if narrow(sr) { 0.0 } else { rw };
                            if let Some(ref mut d) = dual {
                                d.push((lx, lw, rx, rw));
                            }
                            // Single-region geometry always stores the
                            // LEFT region — render_paragraph_lines uses
                            // this for left-chunk x positioning.
                            geom.push((lx, lw));
                        } else {
                            match wrap_side(fz.wrap_text, sl, sr) {
                                Some(true) => {
                                    let (x, w) = right_of_float;
                                    geom.push((x, w.max(1.0)));
                                }
                                Some(false) => geom.push((lx, lw.max(1.0))),
                                None => {
                                    blocked_lines.resize(i, false);
                                    blocked_lines.push(!fz.from_table);
                                    geom.push((col_x + para.indent_left, full_w));
                                }
                            }
                        }
                    } else {
                        geom.push((col_x + para.indent_left, full_w));
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
    // A paragraph whose first word fits no strip beside a float starts below
    // it, the word whole (Word examples: "Internationalization" and "Hi" +
    // no-break space + "Internationalization" beside a 40pt strip, the anchor's
    // own box or an earlier one).
    let from_table = state.pb.float_zone.iter().chain(&extra_float_zones).any(|z| z.from_table);
    if let Some(word) = layout::opening_word_width(&effective_runs, ctx.fonts)
        .filter(|_| !text_empty && !from_table)
    {
        let full_w = (col_w - para.indent_left - para.indent_right).max(1.0);
        let room = |i: usize| match (&poly_dual_geom, &poly_line_geom) {
            (Some(d), _) => d.get(i).map(|&(_, w, _, rw)| w.max(rw)),
            (None, Some(g)) => g.get(i).map(|&(_, w)| w),
            _ => None,
        };
        for i in 0.. {
            let Some(room) = room(i) else { break };
            if blocked_lines.get(i).copied().unwrap_or(false) {
                continue;
            }
            if room >= word || room >= full_w - 0.5 {
                break;
            }
            blocked_lines.resize(blocked_lines.len().max(i + 1), false);
            blocked_lines[i] = true;
        }
    }
    // A right-to-left paragraph fills the strips beside floats from the right
    // (Word example: Hebrew between a wrapText=right box and a bothSides one
    // starts in the strip right of the second).
    if para.bidi
        && let Some(dual) = poly_dual_geom.as_mut()
    {
        for (i, d) in dual.iter_mut().enumerate() {
            if d.3 > 0.0 {
                *d = (d.2, d.3, d.0, d.1);
                if let Some(g) = poly_line_geom.as_mut().and_then(|g| g.get_mut(i)) {
                    *g = (d.0, d.1);
                }
            }
        }
    }
    // A first line with no room beside a float moves to the float's foot, not
    // a whole line lower (Word example: a 17pt strip beside a wrapText=largest
    // picture puts the first line right under it).
    let blocked_top = blocked_lines.iter().take_while(|&&b| b).count();
    let float_drop = if blocked_top > 0 {
        let eff_top = state.pb.slot_top - inter_gap;
        let reach = (blocked_top as f32 + 0.2) * line_h;
        state.pb.float_zone.iter().chain(extra_float_zones.iter())
            .map(|z| eff_top - z.bottom_y)
            .filter(|&d| d <= reach)
            .fold((blocked_top - 1) as f32 * line_h, f32::max)
    } else {
        0.0
    };
    blocked_lines.drain(..blocked_top);
    if let Some(g) = poly_line_geom.as_mut() {
        g.drain(..blocked_top.min(g.len()));
    }
    if let Some(d) = poly_dual_geom.as_mut() {
        d.drain(..blocked_top.min(d.len()));
    }
    if blocked_top > 0 {
        // The first line sits below the float: the lines past the per-line
        // geometry take the whole measure, not the strip beside it.
        para_text_x = col_x + para.indent_left;
        para_text_width = (col_w - para.indent_left - para.indent_right).max(1.0);
        label_x = list_label::label_x(para, col_x, col_w, ctx.fonts);
    }

    let poly_line_widths: Option<Vec<f32>> = poly_line_geom
        .as_ref()
        .map(|g| g.iter().map(|&(_, w)| w).collect());

    let has_inline_image_runs = effective_runs.iter().any(|r| r.inline_image.is_some());
    // Word advances a left tab past any floating image whose body sits on the
    // line, snapping to the first stop clear of the image's right edge. Collect
    // the horizontal spans (in from-text-margin coords) of wrapping images that
    // overlap this paragraph's first line so build_tabbed_line can skip them.
    let tab_exclusions: Vec<(f32, f32)> = if has_tabs && !text_empty {
        let slot_top = state.pb.slot_top;
        para.floating_images
            .iter()
            .filter(|fi| fi.wrap_type.wraps_beside() && wraps_here(fi))
            .filter_map(|fi| {
                let fi_y_top = pre_anchor.fi_y_top(fi, sp);
                let fi_y_bottom = fi_y_top - fi.image.display_height;
                // Only images vertically overlapping the first line band.
                if fi_y_bottom > slot_top + 2.0 || fi_y_top < slot_top - 40.0 {
                    return None;
                }
                let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
                Some((fi_x - col_x, fi_x + fi.image.display_width - col_x))
            })
            .collect()
    } else {
        Vec::new()
    };
    let mut lines = if para.image.is_some() || (text_empty && !has_inline_image_runs) {
        vec![]
    } else {
        // Per-line geometry handles narrow→wide transitions;
        // dual geometry takes priority over single-region widths.
        let plw: Option<&[f32]> = if poly_dual_geom.is_some() {
            None
        } else {
            poly_line_widths.as_deref()
        };
        let build = |per_line_widths: Option<&[f32]>, dual: Option<&[DualRegion]>| {
            build_lines(
                &effective_runs,
                ctx,
                para_text_width,
                cjk,
                &LineOpts {
                    inline_images: Some(&block_inline_images),
                    effects: Some(&block_effect_inlines),
                    tab_stops: &para.tab_stops,
                    indent_left: para.indent_left,
                    indent_right: para.indent_right,
                    hanging: text_hanging,
                    tab_exclusions: &tab_exclusions,
                    per_line_widths,
                    dual,
                },
            )
        };
        if blocked_lines.contains(&true) {
            let open = |v: &[_]| keep_open(v, &blocked_lines);
            let widths = plw.map(open);
            let dual = poly_dual_geom.as_deref().map(|d| keep_open(d, &blocked_lines));
            skip_blocked_lines(&blocked_lines, build(widths.as_deref(), dual.as_deref()))
        } else {
            build(plw, poly_dual_geom.as_deref())
        }
    };
    // A pure one-break anchor can be taller than a short floating table.
    // In that case its first line extends below the float and occupies flow
    // height; dropping it also drops a full line before the paragraph mark.
    let short_empty_clear = || {
        let mut breaks = para.runs.iter().filter(|r| r.is_line_break);
        let (Some(br), None) = (breaks.next(), breaks.next()) else {
            return false;
        };
        para.runs
            .iter()
            .all(|r| r.text.is_empty() && !r.is_tab && r.inline_image.is_none())
            && para.image.is_none()
            && para.inline_chart.is_none()
            && para.smartart.is_empty()
            && para.floating_images.is_empty()
            && para.textboxes.is_empty()
            && state.pb.float_zone.as_ref().is_some_and(|zone| {
                let lhr = break_run_lhr(&effective_runs, br.font_size, ctx.fonts);
                let first_pitch = resolve_line_h(effective_ls, br.font_size, lhr);
                zone.from_table && state.pb.slot_top - inter_gap - first_pitch < zone.bottom_y
            })
    };
    // A clearing break that opens the anchor paragraph of the floating table
    // above it sits beside that table's top; only what follows the break comes
    // below the table (indigenous_innovation's defined terms: one line, not two).
    // The zone is gone once the table broke across pages or left no strip
    // (case129's full-width table: both lines below it), so ask the table's
    // own extent.
    if (leading_table_clear
        || para.clears_floats
            && follows_floating_table
            && !short_empty_clear()
            && state.pb.float_table_extent.as_ref().is_some_and(|t| {
                side_strip(
                    (t.left, t.right),
                    (t.left_from_text, t.right_from_text),
                    (col_x, col_w),
                ) >= MIN_EMPTY_STRIP
            }))
        && lines.len() > 1
        && lines[0].ends_with_break
        && lines[0].chunks.is_empty()
    {
        lines.remove(0);
    }
    // The look-ahead zone reached up through this paragraph's space-after only
    // for its own geometry; following paragraphs see the float's real edge.
    if let (Some((_, top, _)), Some(fz)) = (lookahead, state.pb.float_zone.as_mut()) {
        fz.top_y = top;
    }

    let max_inline_img_h = lines.iter().map(line_max_image_h).fold(0.0f32, f32::max);

    // Paragraph ascent/descent in points: the first baseline sits `para_ascent`
    // below the paragraph top, and picture lines are sized from both (see
    // `inline_line_advance` and `picture_line_bottom`). The bottom part is only
    // read for picture lines, so most paragraphs skip its run scan.
    let para_ascent = exact_baseline_base.unwrap_or(font_size * tallest_ar.unwrap_or(0.75));
    let first_baseline_offset =
        label_boosted_baseline_offset(para, ctx.fonts, para_ascent, font_size)
            * auto_ascent_scale(effective_ls);
    // A grid or an exact rule gives every line the same box.
    if !grid_snapped && !matches!(effective_ls, LineSpacing::Exact(_)) {
        size_lines_by_own_runs(&mut lines, ctx.fonts, effective_ls, line_h, para_ascent, para.text_alignment == crate::model::TextAlignment::Top);
        size_break_lines(&mut lines, &effective_runs, effective_ls, line_h, ctx.fonts);
        size_unsized_lines(
            &mut lines,
            para,
            &effective_runs,
            effective_ls,
            (line_h, font_size, para_ascent),
            ctx.fonts,
        );
        size_tab_mark_line(&mut lines, para, effective_ls, tallest_lhr, ctx.fonts);
    }
    // Only glyphs on a picture's own line hang a descent below it: when the
    // text after leading spaces and a picture wraps to the next line, Word
    // closes the picture line at the picture (leading-picture-image-fits).
    let picture_lines_glyphless = lines
        .iter()
        .filter(|l| line_max_image_h(l) > 0.0)
        .all(|l| {
            l.chunks
                .iter()
                .all(|c| c.inline_image_name.is_some() || c.text.trim().is_empty())
        });
    let picture_runs: Vec<Run> = if picture_lines_glyphless {
        effective_runs
            .iter()
            .filter(|r| r.inline_image.is_some() || r.text.trim().is_empty())
            .cloned()
            .collect()
    } else {
        Vec::new()
    };
    let mut para_metrics = (
        para_ascent,
        if max_inline_img_h > 0.0 {
            let runs = if picture_lines_glyphless { &picture_runs[..] } else { &effective_runs[..] };
            picture_line_bottom(runs, para, ctx.fonts, effective_ls)
        } else {
            0.0
        },
    );

    // Legacy Word puts the mark's extra automatic leading only after the
    // last picture line. Modern Word adds it after every picture line.
    let legacy_picture_leading = if ctx.compat_mode < 15
        && max_inline_img_h > 0.0
        && matches!(effective_ls, LineSpacing::Auto(factor) if factor > 1.0)
        && !effective_runs.is_empty()
        && effective_runs
            .iter()
            .all(|r| r.vanish || r.inline_image.is_some())
    {
        std::mem::replace(&mut para_metrics.1, 0.0)
    } else {
        0.0
    };

    // One line of the mark's font at single spacing: a picture-only line's
    // baseline below its top (and the base its leading is measured from).
    let natural_line_h = font_size * tallest_lhr.unwrap_or(1.2);
    // A picture taller than the text line takes its paragraph's own
    // line-spacing leading below it, sized by the paragraph mark
    // (dental_amalgam: a 68.25pt logo under Normal's 1.15 lines is
    // 70.1pt tall in Word although the next paragraph is single-spaced).
    // The mark's, not the picture run's: Word probe at 3× spacing, 12pt
    // Times New Roman, 12pt Calibri, 24pt Times New Roman and 11pt Cambria
    // marks add 2 × their line whatever the picture run's font and size.
    // Lazy: most paragraphs hold no picture.
    let mark_line = || {
        let (mark_fs, mark_lhr) = unsized_line_metrics(para, font_size, ctx.fonts);
        let mark_line_h = resolve_line_h(effective_ls, mark_fs, mark_lhr);
        (mark_line_h, spacing_leading(mark_line_h, mark_fs, mark_lhr))
    };
    // A shorter one sits on the baseline of a full line of the mark's
    // font (indigenous_innovation's 2.25pt rule under the title). Taller
    // than a natural line is enough for the leading: a 30pt picture under
    // 3 × 13.8pt lines is 57.6pt in Word, not 41.4. Charts and an inline
    // shape alone on its line are pictures too (torture_test: Word's next
    // line is 2.09pt lower than the chart's bottom).
    // Markup view: an inserted picture keeps its run's descent below it like
    // the underlined text it is; a short one then sits on the text baseline
    // instead of the line's foot (Word for Mac probe p22d).
    let inserted_descent = inserted_picture_descent(&effective_runs, ctx.fonts);
    let picture_line_h = |bottom: f32| {
        if bottom + inserted_descent > natural_line_h {
            bottom + inserted_descent + mark_line().1
        } else {
            line_h
        }
    };
    let mut content_h = if para.inline_chart.is_some() {
        picture_line_h(para.content_height)
    } else if let Some(img) = &para.image {
        // Modern Word swaps the inline layout axes in the nearest quarter-turn
        // quadrant; it does not reserve the exact oblique bounding rectangle.
        let swap_axes = ctx.compat_mode >= 15
            && (img.rotation_deg.rem_euclid(360.0) + 45.0).rem_euclid(180.0) >= 90.0;
        let image_height = if swap_axes {
            img.display_width + img.layout_extra_height
        } else {
            para.content_height
        };
        let picture_h = picture_line_h(image_height);
        // An OLE preview wider than the column leaves its line no room for the
        // paragraph mark, which wraps onto a line of its own (alfies_arc: an
        // 11pt line follows a 1014pt-wide OLE logo strip, annotation #186).
        // The paragraph's own indents do not count: learning_cultures keeps
        // the mark beside a column-wide picture in a right-indented paragraph.
        // Compared in whole twips: a picture sized to the column (8616.99
        // twips from its EMU extent against an 8617 twip column) fits.
        let twips = |pt: f32| (pt * 20.0).round();
        if img.is_ole_preview && twips(img.layout_size().0) > twips(col_w) {
            picture_h + mark_line().0
        } else {
            picture_h
        }
    } else if max_inline_img_h > 0.0 {
        lines_height(&lines, line_h, para_metrics)
    } else if text_empty {
        if para.paragraph_mark_vanish {
            0.0
        } else if para.content_height > 0.0 {
            para.content_height
        } else {
            line_h + mark_position_stretch(para, effective_ls, grid_snapped, ctx.fonts)
        }
    } else {
        let num_lines = lines.len();
        // The numbering label is a run on the first line, so its font
        // metrics participate in that line's height.
        let first_line_h = label_boosted_line_h(
            para,
            ctx.fonts,
            line_h,
            effective_ls,
            font_size,
            tallest_lhr,
            tallest_ar,
        ) - lines
            .first()
            .and_then(|l| l.pitch)
            .map_or(0.0, |p| line_h - p);
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
            let last = lines.len() - 1;
            for (i, line) in lines.iter().enumerate().skip(1) {
                if line.leading_tab_wrap
                    && let Some(pitch) = line.pitch
                {
                    h += pitch;
                } else if let Some(bfs) = line.break_font_size {
                    // The empty line a trailing break leaves holds only the
                    // paragraph mark, whose rPr sizes it — the break char
                    // sizes the line it terminates, not this one (samtale:
                    // 26pt br before a 12pt mark, annotation #121).
                    let mark_h = (i == last && line.chunks.is_empty())
                        .then(|| mark_line_h(para, effective_ls, tallest_lhr, ctx.fonts))
                        .flatten();
                    if let Some(mark_h) = mark_h {
                        h += mark_h;
                    } else {
                        let blhr = break_run_lhr(&effective_runs, bfs, ctx.fonts);
                        h += resolve_line_h(effective_ls, bfs, blhr);
                    }
                } else {
                    h += line.pitch.unwrap_or(line_h);
                }
            }
            h
        }
    };

    if legacy_picture_leading > 0.0
        && lines
            .last()
            .is_some_and(|line| line_max_image_h(line) > para_ascent)
    {
        content_h += legacy_picture_leading;
    }

    // A tall inline picture on the first line lowers that line's baseline; the
    // list label sits on the lowered one (render_paragraph_lines drops the text
    // lines itself).
    let first_line_drop = lines
        .first()
        .map_or(0.0, |l| inline_image_line_extra(l, para_ascent));

    // The shading covers the lines, not the room reserved below them for a
    // float: indonesian_school's "Format 12" box shows between its anchor's
    // white shading and the next paragraph's.
    let lines_h = content_h;

    // Extra height from floating images that extends beyond
    // the text content — used only for page-break decisions,
    // not for cursor advancement (text wraps beside the image).
    let mut float_overflow_h = 0.0f32;
    // How far a topAndBottom float over the first line pushes the text down.
    let mut text_drop = float_drop;
    content_h += float_drop;

    // A band keeps the lines out by itself (`step_below_bands`).
    for fi in para
        .floating_images
        .iter()
        .filter(|fi| wraps_here(fi) && !image_blocks_band(fi, sp, col_x, col_w, text_width))
    {
        let reserve = match fi.wrap_type {
            WrapType::TopAndBottom => true,
            // wrapSquare/Tight with a usable side strip: content (even empty
            // spacer paragraphs) flows beside via the float zone — reserving
            // here would double-count the image height (brazilian, ~42pt
            // strips). With no usable strip (sample500kB: image width ==
            // column width) Word stacks everything below the float, which
            // reserving the image height in the anchor reproduces.
            WrapType::Square | WrapType::Tight => {
                let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
                let left_gap = (fi_x - fi.dist_left) - col_x;
                let right_gap = (col_x + col_w) - (fi_x + fi.image.display_width + fi.dist_right);
                left_gap.max(right_gap) < MIN_EMPTY_STRIP
            }
            WrapType::Through => false,
            WrapType::None => false,
        };
        let fi_h = match fi.v_position {
            VerticalPosition::Offset(o) => {
                o + fi.dist_top + fi.image.display_height + fi.dist_bottom
            }
            _ => fi.dist_top + fi.image.display_height + fi.dist_bottom,
        };
        if reserve && !text_empty && leading_full_width_picture(fi, sp, col_x, col_w, text_width) {
            float_overflow_h = float_overflow_h.max(fi_h);
        } else if reserve {
            // Wide images block all text — add to content_h. An empty anchor
            // paragraph's own line moves below such a float when the float
            // covers it: cyprus' page-wide letterhead (margin-relative, top
            // 83.6pt above the margin) keeps the mark's 14.65pt line under its
            // bottom, and Word probes with paragraph- and margin-relative
            // offsets, square and topAndBottom wrapping all do the same.
            let content_top = state.pb.slot_top - inter_gap;
            let zone = FloatZone::for_float(
                fi,
                resolve_fi_x(fi, sp, col_x, col_w, text_width),
                pre_anchor.fi_y_top(fi, sp),
            );
            // ponytail: only a float that starts at or below the paragraph top,
            // or a paragraph opening the page (the probed cases); a float
            // reaching over earlier lines on the page (sample500kB) keeps the
            // plain reserve.
            let covers_line = zone.top_y > content_top - content_h
                && zone.bottom_y < content_top
                && (zone.top_y <= state.pb.slot_top + 0.5 || state.pb.is_at_page_top(sp));
            // A topAndBottom picture over a text paragraph's first line moves
            // its lines below it (Word probe: offset 0 puts the first baseline
            // under the picture; +30pt leaves the first line above it).
            let covers_first_line = matches!(fi.wrap_type, WrapType::TopAndBottom)
                && zone.top_y > content_top - line_h
                && zone.bottom_y < content_top
                && (zone.top_y <= state.pb.slot_top + 0.5 || state.pb.is_at_page_top(sp));
            if text_empty && !para.paragraph_mark_vanish && covers_line {
                content_h = content_h.max(content_top - zone.bottom_y + content_h);
            } else if !text_empty && covers_first_line {
                text_drop = text_drop.max(content_top - zone.bottom_y);
                content_h = content_h.max(fi_h).max(lines_h + text_drop);
            } else if fi.v_relative_from == VRelativeFrom::Paragraph {
                content_h = content_h.max(fi_h);
            } else {
                // A page- or margin-placed float's offset is not from this
                // paragraph: it reaches down to the float's own bottom (a
                // 45pt rubric banner at 484.8pt from the page top made its
                // one-line anchor 530pt tall and sent it to the next page).
                // The space after overlaps the float: Word starts the next
                // paragraph at the banner's foot. Only a float the next line
                // would run into counts; one further down the page leaves the
                // text flowing (a page-foot logo 680pt below its anchor).
                // ponytail: lines that reach a lower float later are not moved.
                let next_line_bottom = content_top - content_h - para.space_after - line_h;
                if zone.top_y > next_line_bottom {
                    content_h = content_h.max(content_top - zone.bottom_y - para.space_after);
                }
            }
        } else if fi.v_relative_from == VRelativeFrom::Paragraph
            && fi.wrap_type.wraps_beside()
            && fi_h <= sp.page_height - sp.margin_top - sp.margin_bottom
        {
            // Paragraph-relative wrapping images: track overflow
            // for page-break check only (text wraps beside them). One taller
            // than the page body fits no page: it stays, moved up to the
            // page top (`ParaAnchor::fi_y_top`).
            float_overflow_h = float_overflow_h.max(fi_h);
        }
    }

    for tb in &para.textboxes {
        // A band keeps the lines out by itself (`step_below_bands`).
        let reserve = tb.v_relative_from == VRelativeFrom::Paragraph
            && !hangs_below_anchor(tb, para, col_w, ctx)
            && match tb.wrap_type {
                WrapType::TopAndBottom => true,
                WrapType::Square => {
                    textbox_room_beside(tb, para, sp, (col_x, col_w), text_width)
                        .is_some_and(|room| room < MIN_EMPTY_STRIP)
                }
                _ => false,
            };
        if reserve {
            let tb_bottom =
                tb.v_offset_pt + textbox_render::textbox_height(tb, ctx) + tb.dist_bottom;
            match tb.v_relative_from {
                // The offset counts from above the paragraph's own space
                // before, which already sits above the content. An empty
                // anchor's own line that would overlap the box moves below it,
                // keeping that space (indonesian_school's "Format 12" with
                // before=200 twips: Word +22.5pt = 10 + box bottom 33.18 +
                // line - 33.18).
                // Drawings on the line count from it (isla_language's lesson
                // canvas).
                // An inline shape alone on its line is a picture there.
                _ if tb.on_line => {
                    let h = if text_empty {
                        picture_line_h(tb_bottom)
                    } else {
                        tb_bottom
                    };
                    content_h = content_h.max(h);
                }
                _ => {
                    let before = own_space_before(inter_gap, state.prev_space_after);
                    let line_overlaps = tb.v_offset_pt < before + lines_h && tb_bottom > before;
                    // Only a mark line moves below the box: a picture-only
                    // paragraph keeps its picture line beside it (a 65pt logo
                    // beside a square-wrapped title box ends the paragraph at
                    // the logo's foot in Word, not the two stacked).
                    // Side-by-side boxes move the line below the lower one, not
                    // below their stacked heights (torture_test's linked pair).
                    // A box over a text paragraph's first line moves its lines
                    // below it, as a picture does (Word examples: topAndBottom,
                    // and square with a 17pt strip beside it).
                    let first_line_overlaps =
                        tb.v_offset_pt < before + line_h && tb_bottom > before;
                    content_h = if text_empty && para.image.is_none() && line_overlaps {
                        content_h.max(tb_bottom + lines_h)
                    } else if !text_empty && first_line_overlaps {
                        text_drop = text_drop.max(tb_bottom - before);
                        content_h.max(lines_h + text_drop)
                    } else {
                        content_h.max(tb_bottom - before)
                    };
                }
            }
        } else if tb.v_relative_from == VRelativeFrom::Paragraph
            && textbox_wraps_beside(tb, para, sp, (col_x, col_w), text_width)
        {
            // Like a picture text wraps beside: only the page break counts
            // it (Word example: an empty anchor 40pt above the bottom margin
            // goes to the next page with its 180pt box).
            let offset = match tb.v_position {
                VerticalPosition::Offset(o) => o,
                _ => 0.0,
            };
            let h = textbox_render::textbox_height(tb, ctx);
            let tb_h = offset + tb.dist_top + h + tb.dist_bottom;
            if tb_h <= sp.page_height - sp.margin_top - sp.margin_bottom {
                float_overflow_h = float_overflow_h.max(tb_h);
            }
        }
    }

    // Vanished paragraph mark: zero out height and spacing
    if text_empty && para.paragraph_mark_vanish {
        content_h = 0.0;
        inter_gap = 0.0;
    }

    // Word treats consecutive paragraphs with identical borders and matching
    // box edges as one border group: top padding/rule only on the first, bottom
    // padding/rule only on the last, and a `between` rule (if any) at the joins.
    let prev_borders_match = prev_para.is_some_and(|pp| joins_border_group(pp, para));
    let next_borders_match = next_para.is_some_and(|np| joins_border_group(para, np));
    let bottom_collapses = next_borders_match;

    // The top border's band sits inside the paragraph, like the bottom one
    // (`bdr_bottom_extent`): case17's 0.5pt and 1pt boxes start their band at
    // the paragraph top and their text `space` below its lower edge.
    let bdr_top_pad = if prev_borders_match {
        0.0
    } else {
        border_band(para.borders.top.as_ref())
    };
    let bdr_top_half_band = if prev_borders_match {
        0.0
    } else {
        para.borders.top.as_ref().map_or(0.0, |b| b.extent() / 2.0)
    };
    let (bottom_bdr, bottom_band) =
        bottom_border(para, next_para.filter(|_| bottom_collapses));
    let bdr_bottom_pad = bottom_band.0 + bottom_bdr.map_or(0.0, |b| b.extent() / 2.0);
    // Full extent of bottom border below content (to border bottom edge)
    let bdr_bottom_extent = bottom_band.1;

    // Word measures the bottom border `space` attribute from
    // the full line-height content bottom, not from the text
    // descent.  No trailing-lead adjustment is needed.

    let needed = inter_gap + bdr_top_pad + content_h + bdr_bottom_extent;
    // For page-break decisions, also account for floating
    // images that extend below the text content.
    let needed_with_floats = needed.max(inter_gap + float_overflow_h);
    let at_page_top = state.pb.is_at_page_top(sp);

    // Pre-compute footnote space for this paragraph so the
    // page-break check accounts for footnotes the paragraph
    // introduces (otherwise they're only tracked after
    // rendering, which can cause body/footnote overlap).
    let line_refs: Vec<Vec<u32>> = lines
        .iter()
        .map(|l| l.chunks.iter().filter_map(|c| c.footnote_id).collect())
        .collect();
    let run_refs: Vec<u32> = para.runs.iter().filter_map(|r| r.footnote_id).collect();
    let (line_fn_extra, para_fn_extra) = per_line_footnote_extra(
        &line_refs,
        &run_refs,
        &state.pb.footnote_ids_set,
        if state.pb.col_fn_reserved == 0.0 {
            ctx.note_separator.height
        } else {
            0.0
        },
        |id| footnote_height(id, &doc.footnotes, ctx, col_geometry[state.current_col].1),
    );

    // Word allows the last line's trailing inter-line
    // spacing to extend past the bottom margin — only the
    // text (ascent + descent) must fit inside the content
    // area.  Compute the excess leading so the page-break
    // check can tolerate it.
    // A lone paragraph mark is a last line too: Word keeps an empty 1.15-spaced
    // Calibri paragraph whose leading overhangs the margin by 0.8pt
    // (sao_paulo_procurement_contract p2) and an empty 1.5-spaced Arial one
    // (czech_wastewater_discharge_permit p1).
    let mark_only = text_empty && !para.paragraph_mark_vanish && para.content_height <= 0.0;
    // Not into a footnote area, though, as in the per-line check below
    // (zimbabwe_gold p3: Word moves a double-spaced paragraph whose last line
    // would end 0.6pt into it; probes: no tolerance).
    let last_line_lead = if (!lines.is_empty() || mark_only)
        && state.pb.col_fn_reserved == 0.0
        && para_fn_extra == 0.0
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

    let keep_next_extra = if para.keep_next && !at_page_top {
        let mut extra = 0.0;
        // Footnotes the kept paragraphs bring come along too
        // (uk_commercial_lease's "Break Date" moves with its definition's note 8).
        let mut chain_notes: Vec<u32> = Vec::new();
        let mut prev_sa = effective_space_after;
        let mut first_link: Option<f32> = None;
        let mut i = block_idx + 1;
        loop {
            let next = match section_blocks.get(i) {
                Some(Block::Paragraph(p)) => p,
                // The chain reaches into a following in-flow table: the paragraph
                // stays with its first row (czech_wastewater's "8. Seznam…"
                // heading moves to page 4 with its table).
                Some(Block::Table(t)) if t.position.is_none() => {
                    extra += prev_sa + first_row_height(t, &ctx, col_geometry[state.current_col].1);
                    break;
                }
                _ => break,
            };
            if next.page_break_before {
                // A page break opening the next paragraph, a `<w:br>` or
                // pageBreakBefore, ends the chain where it is: Word keeps the
                // kept paragraphs on this page (online export probes,
                // 2026-10-03; run-borders' items before a pageBreakBefore
                // Heading 3).
                break;
            }
            chain_notes.extend(next.runs.iter().filter_map(|r| r.footnote_id));
            let (nfs, nlhr, _) = tallest_run_metrics(&next.runs, ctx.fonts);
            let next_inter = f32::max(prev_sa, next.space_before);
            let next_ls = next.line_spacing.unwrap_or(ctx.doc_line_spacing);
            // An exact or at-least box is the whole first line (issue #145's
            // heading moves with its 40pt-exact successor); an auto multiple's
            // extra leading may hang past the margin like `last_line_lead`.
            let first_ls = match next_ls {
                LineSpacing::Auto(_) => LineSpacing::Auto(1.0),
                boxed => boxed,
            };
            // A picture makes its line as tall as itself (issue #146: a
            // heading kept with an 80pt picture under a 14pt line moves with
            // it); a paragraph of one picture holds it in `image`, not a run.
            let picture_h = if next.image.is_some() {
                next.content_height
            } else {
                layout::runs_max_image_h(&next.runs)
            };
            let next_first_line_h = line_h_with_picture(
                picture_h,
                &next.runs,
                next,
                ctx.fonts,
                next_ls,
                resolve_line_h(first_ls, nfs, nlhr),
            );
            // The chain needs as many of the next paragraph's lines as must
            // stay together on this page: all of a keepLines paragraph; else
            // one without widow control, with it two, or all of a paragraph
            // of three or fewer (it can't split without leaving a lone line).
            // lithuanian's headings end on an empty paragraph and fit at the
            // foot; western_australia's end on a three-line item and move.
            // A kept paragraph that stays whole passes the chain on to its
            // successor; one that may split ends it (australian_higher's
            // keepLines items move with heading 6.5.1).
            let mut n = None;
            let mut count = || {
                *n.get_or_insert_with(|| line_count(next, ctx, col_geometry[state.current_col].1))
            };
            let needed = if next.keep_lines {
                count()
            } else {
                lines_kept_together(next.widow_control, &mut count)
            };
            extra += next_inter
                + next_first_line_h
                + needed.saturating_sub(1) as f32 * resolve_line_h(next_ls, nfs, nlhr);
            first_link.get_or_insert(extra);
            if !next.keep_next || needed < count() {
                break;
            }
            if next.page_break_after {
                extra = f32::MAX;
                break;
            }
            prev_sa = next.space_after;
            i += 1;
        }
        let opens_note_area = state.pb.col_fn_reserved == 0.0 && para_fn_extra == 0.0;
        let (_, notes_h) = per_line_footnote_extra(
            &[],
            &chain_notes,
            &state.pb.footnote_ids_set,
            if opens_note_area {
                ctx.note_separator.height
            } else {
                0.0
            },
            |id| footnote_height(id, &doc.footnotes, ctx, col_geometry[state.current_col].1),
        );
        // A chain no page can hold starts on a new page and flows from
        // there: its later paragraphs keep only their own link to the next,
        // or each would strand on a page of its own (a bibliography of
        // Heading 3 entries opens page 3 in Word, the rest follows, each
        // entry still with its successor's first lines).
        let page_h = effective_slot_top(sp, false, state.pb.page_count(), ctx)
            - state.effective_margin_bottom;
        let inside_chain = block_idx
            .checked_sub(1)
            .and_then(|i| section_blocks.get(i))
            .is_some_and(|b| matches!(b, Block::Paragraph(p) if p.keep_next));
        // Judged once, where the chain starts: from a later member the rest
        // may fit a page and would move again.
        if !inside_chain {
            state.long_keep_chain = needed_with_floats + extra + notes_h > page_h;
        }
        if inside_chain && state.long_keep_chain {
            first_link.unwrap_or(extra)
        } else {
            extra + notes_h
        }
    } else {
        0.0
    };

    // A multi-line paragraph that opens the page and outruns it splits like
    // any other: one of 108 manual line breaks, double spaced, took one page
    // instead of five (corpus 2026-11-05; probe p18b, 70 break lines). A
    // single oversized line stays put (case150's 700pt exact spacer).
    if (!at_page_top || lines.len() > 1)
        && state.pb.slot_top - needed_with_floats - keep_next_extra + last_line_lead
            < state.effective_margin_bottom + para_fn_extra
    {
        let available = state.pb.slot_top - inter_gap - text_drop - state.effective_margin_bottom;
        let first_line_h = tallest_lhr
            .map(|ratio| font_size * ratio)
            .unwrap_or(font_size);
        // Each line must fit together with the footnotes it introduces.
        let mut lines_that_fit = 0usize;
        if line_h > 0.0 {
            let mut fn_acc = 0.0f32;
            // Line i fits when the advances of the lines above it plus its own
            // text height fit (its trailing leading may hang past the margin),
            // but not past a footnote area: there the whole line must fit
            // (environmental_law_clinic's double-spaced lines stop a line
            // earlier above the footnotes on every page in Word).
            let mut above = 0.0f32;
            for (i, fn_extra) in line_fn_extra.iter().enumerate() {
                fn_acc += fn_extra;
                let room = available - fn_acc;
                let own_pitch = lines.get(i).and_then(|l| l.pitch);
                let above_footnotes = state.pb.col_fn_reserved > 0.0 || fn_acc > 0.0;
                let own_h = line_fit_height(
                    above_footnotes || matches!(effective_ls, LineSpacing::Exact(_)),
                    own_pitch,
                    line_h,
                    first_line_h,
                );
                if above + own_h > room {
                    break;
                }
                lines_that_fit = i + 1;
                above += own_pitch.unwrap_or(line_h);
            }
        }

        let text_would_split = lines_h > available;
        if para.widow_control {
            // Ensure at least 2 lines remain on next page (orphan prevention)
            if lines_that_fit > 0 && lines.len().saturating_sub(lines_that_fit) < 2 {
                lines_that_fit = lines.len().saturating_sub(2);
            }
        }

        // keepLines: don't split — move entire paragraph to next column/page
        let leading_picture_keeps_text = content_h
            <= sp.page_height - sp.margin_top - sp.margin_bottom
            && text_would_split
            && para
                .floating_images
                .iter()
                .any(|fi| leading_full_width_picture(fi, sp, col_x, col_w, text_width));
        if para.keep_lines || leading_picture_keeps_text {
            lines_that_fit = 0;
        }

        let min_split_lines = if para.widow_control { 2 } else { 1 };
        if lines_that_fit >= min_split_lines && lines_that_fit < lines.len() {
            let first_part = &lines[..lines_that_fit];
            state.pb.slot_top -= inter_gap;
            layout::set_change_bar_extension(
                own_space_before(inter_gap, state.prev_space_after),
                effective_space_after,
            );
            let split_anchor = ParaAnchor::new(
                state.pb.take_float_anchor_for(block_idx),
                state.pb.anchor_slot_top(sp, ctx),
                own_space_before(inter_gap, state.prev_space_after),
                sp,
                ctx,
            );
            let baseline_offset = if grid_snapped {
                grid_baseline
            } else {
                first_baseline_offset
            };
            let baseline_y = state.pb.slot_top - baseline_offset - text_drop;

            // Its text boxes hang from its top on this page, as an unsplit
            // paragraph's do.
            for tb in &para.textboxes {
                render_para_textbox(tb, para, split_anchor, sp, (col_x, col_w), text_width, ctx, &mut state.pb);
            }

            // One element for both halves: its content continues on the next page.
            let tags = state.pb.para_tags(para, doc);
            let tag = tags.1;
            state.pb.begin_para_tags(tags, |content| {
                render_list_label(
                    content,
                    para,
                    ctx.fonts,
                    label_x,
                    baseline_y - first_line_drop,
                    font_size,
                )
            });

            render_paragraph_lines(
                &mut state.pb.content,
                first_part,
                &para.alignment,
                para_text_x,
                para.indent_left,
                para_text_width,
                baseline_y,
                line_h,
                para_metrics,
                lines.len(),
                0,
                &mut state.pb.links,
                text_hanging,
                ctx.fonts,
                poly_line_geom.as_deref(),
                &mut state.pb.gradient_specs,
                Some(&mut state.pb.comment_anchors),
                ln_cfg.map(|(start, count_by, right_x)| LineNumberArg {
                    counter: &mut state.pb.line_number_counter,
                    start,
                    count_by,
                    right_x,
                    font: ln_font.clone(),
                }),
                Some(LinkTagger::new(
                    &mut state.pb.tags,
                    state.pb.all_contents.len(),
                    tag,
                )),
                None,
            );
            state.pb.end_tag();

            // Floats anchored before the paragraph's last text hang from its
            // top on this page, like an unsplit paragraph's (the split path
            // returns before the ordinary float rendering below).
            let (leading, leading_names, leading_effects) = float_subset(
                &para.floating_images,
                floating_image_pdf_names,
                effect_floating_names,
                state.global_block_idx,
                |fi| !blank_after(&para.runs, fi.anchor_run),
            );
            if !leading.is_empty() {
                let page = state.pb.all_contents.len();
                render_floating_images(&leading, state.global_block_idx,
                    &leading_names, &leading_effects, sp, col_x, col_w, text_width,
                    split_anchor, &mut state.pb.behind_shapes, &mut state.pb.deferred_shapes,
                    &mut state.pb.tags, page);
            }

            // Footnotes referenced on the lines that stay here belong to this
            // page's footnote area; the flush below would otherwise carry them
            // to the continuation page while the space stays reserved here.
            let first_part_fn_ids = line_footnote_ids(first_part);
            for &id in &first_part_fn_ids {
                track_page_footnote(state, doc, ctx, col_geometry[state.current_col], id);
            }

            // The column ends below the lines that stay (its separator reaches them).
            state.pb.slot_top -= text_drop + lines_height(first_part, line_h, para_metrics);
            state.pb.advance_column_or_page(
                &mut state.current_col,
                col_count,
                sect_idx,
                sp,
                &mut state.effective_margin_bottom,
                ctx,
            );

            let baseline_offset2 = if grid_snapped {
                grid_baseline
            } else {
                para_ascent * auto_ascent_scale(effective_ls)
            };
            // The rest may itself outrun the column: a long paragraph spans as
            // many columns or pages as it needs (case80 probes: one paragraph
            // balanced over three columns).
            // The trailing pictures of a split paragraph belong to its
            // final fragment: a float goes on the page its anchor run's
            // line lands on, hung from the paragraph's top on that page
            // (an external tracked corpus cover: a deleted page-wide
            // picture after five pages of inserted text prints on the
            // fifth page). Full-width pictures also keep following visible
            // text below them, while intervening empty marks may occupy
            // the space above.
            let (trailing, trailing_names, trailing_effects) = float_subset(
                &para.floating_images,
                floating_image_pdf_names,
                effect_floating_names,
                state.global_block_idx,
                |fi| blank_after(&para.runs, fi.anchor_run),
            );
            let mut start = lines_that_fit;
            loop {
                let remaining = &lines[start..];
                let room = state.pb.slot_top - state.effective_margin_bottom;
                let mut fit = 0usize;
                let mut above = 0.0f32;
                for l in remaining {
                    if above + l.pitch.map_or(first_line_h, |p| p.min(first_line_h)) > room {
                        break;
                    }
                    fit += 1;
                    above += l.pitch.unwrap_or(line_h);
                }
                if para.widow_control && fit < remaining.len() {
                    if remaining.len() - fit < 2 {
                        fit = remaining.len().saturating_sub(2);
                    }
                    // Widow control keeps two lines together here too; with
                    // room for fewer, the column is skipped.
                    if fit < 2 {
                        fit = 0;
                    }
                }
                // A fresh page always takes a line, except a balanced one whose
                // short columns must make the trial fail instead.
                let fresh_page = state.pb.is_at_page_top(sp)
                    && state
                        .pb
                        .balance_floor
                        .is_none_or(|(page, _)| page != state.pb.page_count());
                if fit == 0 && !fresh_page {
                    state.pb.advance_column_or_page(
                        &mut state.current_col,
                        col_count,
                        sect_idx,
                        sp,
                        &mut state.effective_margin_bottom,
                        ctx,
                    );
                    continue;
                }
                let chunk = &remaining[..fit.clamp(1, remaining.len())];
                let baseline_y2 = state.pb.slot_top - baseline_offset2;
                let (rest_col_x, rest_col_w) = col_geometry[state.current_col];
                let rest_text_x = rest_col_x + para.indent_left;
                let rest_text_width = (rest_col_w - para.indent_left - para.indent_right).max(1.0);

                let last_fragment = fit >= remaining.len() && !trailing.is_empty();
                let anchor = ParaAnchor::new(None, state.pb.slot_top, 0.0, sp, ctx);
                // Where the fragment's lines go around the first wrapping
                // trailing float (Word probes p20c/p20e: lines whose box meets
                // the picture's band sit beside it, x 249 by a 150pt picture
                // at 90; the corpus cover's page-wide picture leaves no side,
                // its lines go below, and when below is off the page the
                // lines move to the next page, one line staying per page, and
                // the anchor line alone sits at the page bottom above the
                // footer: the corpus cover's "Montering", an empty line and
                // "Med forbehold" take one page each).
                let mut plan: Option<FragmentPlan> = None;
                if last_fragment
                    && let Some(fi) = trailing.iter().find(|fi| fi.wrap_type.wraps_beside())
                {
                    let fi_x = resolve_fi_x(fi, sp, rest_col_x, rest_col_w, text_width);
                    let zone = FloatZone::for_float(fi, fi_x, anchor.fi_y_top(fi, sp));
                    let mut y = state.pb.slot_top;
                    let mut first_in = None;
                    let mut after = chunk.len();
                    for (i, l) in chunk.iter().enumerate() {
                        let adv = lines_height(std::slice::from_ref(l), line_h, para_metrics);
                        let in_band = y > zone.bottom_y && y - adv < zone.top_y;
                        if in_band && first_in.is_none() {
                            first_in = Some((i, y));
                        }
                        if !in_band && first_in.is_some() && after == chunk.len() {
                            after = i;
                        }
                        y -= adv;
                    }
                    if let Some((from, y_b)) = first_in {
                        let (mut tx, mut tw, mut lx) = (rest_text_x, rest_text_width, 0.0);
                        // The line's box meets the band though its top may sit above it.
                        let legacy = ctx.compat_mode < 15;
                        zone.narrow_paragraph(y_b.min(zone.top_y), rest_col_x, rest_col_w, para, &mut tx, &mut tw, &mut lx, legacy);
                        plan = Some(if tw < rest_text_width - 0.5 {
                            FragmentPlan::Beside { from, to: after, text_x: tx, text_w: tw }
                        } else if zone.bottom_y - lines_height(&chunk[from..], line_h, para_metrics)
                            >= state.effective_margin_bottom
                        {
                            FragmentPlan::Below { from, top: zone.bottom_y }
                        } else if chunk.len() > from.max(1) {
                            FragmentPlan::NextPage { keep: from.max(1) }
                        } else {
                            FragmentPlan::PageBottom
                        });
                    }
                }
                if let Some(FragmentPlan::NextPage { keep }) = plan {
                    fit = keep;
                }
                let chunk = &remaining[..fit.clamp(1, remaining.len())];
                let last_fragment = last_fragment && fit >= remaining.len();
                if last_fragment {
                    let page = state.pb.all_contents.len();
                    render_floating_images(&trailing, state.global_block_idx,
                        &trailing_names, &trailing_effects, sp, rest_col_x, rest_col_w,
                        text_width, anchor, &mut state.pb.behind_shapes,
                        &mut state.pb.deferred_shapes, &mut state.pb.tags, page);
                }
                // Full-width pictures keep following visible text below them,
                // while intervening empty marks may occupy the space above.
                let trailing_images_fit = last_fragment
                    && trailing.iter().all(|fi| {
                        fi.v_relative_from == VRelativeFrom::Paragraph
                            && matches!(fi.v_position, VerticalPosition::Offset(y) if y > 0.0)
                            && fi.wrap_type.wraps_beside()
                            && {
                                let x = resolve_fi_x(fi, sp, rest_col_x, rest_col_w, text_width);
                                (x - fi.dist_left - rest_col_x).max(
                                    rest_col_x + rest_col_w - x - fi.image.display_width - fi.dist_right
                                ) < MIN_EMPTY_STRIP
                            }
                            && resolve_fi_y_top(fi, sp, state.pb.slot_top)
                                - fi.image.display_height - fi.dist_bottom >= state.effective_margin_bottom
                    });
                if trailing_images_fit {
                    for fi in &trailing {
                        let x = resolve_fi_x(fi, sp, rest_col_x, rest_col_w, text_width);
                        let top = resolve_fi_y_top(fi, sp, anchor.top(fi.on_line));
                        state.pb.float_zone = Some(FloatZone::for_float(fi, x, top));
                        let bottom = top - fi.image.display_height - fi.dist_bottom;
                        state.pb.next_text_below_picture = Some(
                            state.pb.next_text_below_picture.map_or(bottom, |old| old.min(bottom)));
                    }
                }
                // (from, to, text_x, text_w, baseline of the segment's first line)
                let mut segments: Vec<(usize, usize, f32, f32, f32)> = Vec::new();
                let mut extra_drop = 0.0f32;
                // (`lines_height` of no lines is one pitch.)
                let seg_h = |a: usize, b: usize| {
                    if a < b { lines_height(&chunk[a..b], line_h, para_metrics) } else { 0.0 }
                };
                match plan {
                    Some(FragmentPlan::Beside { from, to, text_x, text_w }) if last_fragment => {
                        let h_a = seg_h(0, from);
                        let h_b = seg_h(from, to);
                        segments.push((0, from, rest_text_x, rest_text_width, baseline_y2));
                        segments.push((from, to, text_x, text_w, baseline_y2 - h_a));
                        segments.push((to, chunk.len(), rest_text_x, rest_text_width, baseline_y2 - h_a - h_b));
                    }
                    Some(FragmentPlan::Below { from, top }) if last_fragment => {
                        let h_a = seg_h(0, from);
                        segments.push((0, from, rest_text_x, rest_text_width, baseline_y2));
                        segments.push((from, chunk.len(), rest_text_x, rest_text_width, top - baseline_offset2));
                        extra_drop = (state.pb.slot_top - h_a) - top;
                    }
                    Some(FragmentPlan::PageBottom) if last_fragment => {
                        let top = state.pb.footer_top(sp, ctx) + lines_height(chunk, line_h, para_metrics);
                        segments.push((0, chunk.len(), rest_text_x, rest_text_width, top - baseline_offset2));
                        extra_drop = state.pb.slot_top - top;
                    }
                    _ => segments.push((0, chunk.len(), rest_text_x, rest_text_width, baseline_y2)),
                }
                state.pb.begin_tag(tag);
                for &(from, to, seg_x, seg_w, seg_baseline) in segments.iter().filter(|s| s.1 > s.0) {
                render_paragraph_lines(
                    &mut state.pb.content,
                    &chunk[from..to],
                    &para.alignment,
                    seg_x,
                    para.indent_left,
                    seg_w,
                    seg_baseline,
                    line_h,
                    para_metrics,
                    lines.len(),
                    start + from,
                    &mut state.pb.links,
                    text_hanging,
                    ctx.fonts,
                    None,
                    &mut state.pb.gradient_specs,
                    Some(&mut state.pb.comment_anchors),
                    ln_cfg.map(|(start, count_by, right_x)| LineNumberArg {
                        counter: &mut state.pb.line_number_counter,
                        start,
                        count_by,
                        right_x,
                        font: ln_font.clone(),
                    }),
                    Some(LinkTagger::new(
                        &mut state.pb.tags,
                        state.pb.all_contents.len(),
                        tag,
                    )),
                    None,
                );
                }
                state.pb.end_tag();

                state.pb.slot_top -= lines_height(chunk, line_h, para_metrics) + extra_drop;
                start += chunk.len();
                if start >= lines.len() {
                    break;
                }
                state.pb.advance_column_or_page(
                    &mut state.current_col,
                    col_count,
                    sect_idx,
                    sp,
                    &mut state.effective_margin_bottom,
                    ctx,
                );
            }
            state.prev_space_after = effective_space_after;
            layout::set_change_bar_extension(0.0, 0.0);

            // Track the remaining footnotes for the split paragraph on the new page
            for run in para.runs.iter() {
                if let Some(id) = run.footnote_id
                    && !first_part_fn_ids.contains(&id)
                {
                    track_page_footnote(state, doc, ctx, col_geometry[state.current_col], id);
                }
                if let Some(id) = run.endnote_id {
                    state.pb.track_endnote(id);
                }
            }

            state.global_block_idx += 1;
            if para.page_break_after {
                state
                    .pb
                    .begin_next_page(sect_idx, sp, &mut state.effective_margin_bottom, ctx);
                state.prev_space_after = 0.0;
                state.current_col = 0;
            }
            return true;
        }

        let page = state.pb.page_count();
        state.pb.overflow_column_or_page(
            &mut state.current_col,
            col_count,
            sect_idx,
            sp,
            &mut state.effective_margin_bottom,
            ctx,
            state.prev_space_after,
        );
        inter_gap = 0.0;
        // Its own floats move to the new column with it, an earlier anchor's
        // stay behind: lay it out again there, its gap the overflow's (Word
        // examples: a 100pt box and its anchor's text start 280pt right in the
        // second of two 200pt columns, at its top after a 24pt space after; a
        // paragraph beside an earlier box moved on by widow control takes the
        // next column's or page's full measure).
        let new_page = state.pb.page_count() != page;
        let earlier_floats = poly_line_geom.is_some() && !own_zone_set;
        if (own_zone_set || earlier_floats)
            && ((col_geometry[state.current_col].0 - col_x).abs() > 0.01
                || (new_page && earlier_floats))
            && !para.page_break_before
            && !para.column_break_before
            && !leading_table_clear
        {
            if !new_page {
                (state.pb.float_zone, state.pb.more_float_zones) = entry_zones;
            }
            state.reflowed = true;
            return render_paragraph_block(
                para,
                state,
                ctx,
                sp,
                col_geometry,
                col_count,
                text_width,
                sect_idx,
                block_idx,
                section_blocks,
                floating_image_pdf_names,
                inline_image_pdf_names,
                image_pdf_names,
                effect_names,
                effect_floating_names,
                effect_inline_names,
                doc,
                smartart_font_key,
                smartart_image_names,
                debug_wrap,
                table_connector_anchor,
            );
        }
    }

    // Suppress space_before at the top of a page
    if let Some(gap) = state
        .pb
        .page_top_gap(sp, effective_space_before, state.prev_space_after)
    {
        state.pb.top_suppressed = (effective_space_before - gap).max(0.0);
        inter_gap = gap;
    }

    let applied_inter_gap = inter_gap;
    // The part of the band above that lies below the previous paragraph's
    // `between` border takes this paragraph's shading.
    let prev_band_below = if prev_borders_match && !state.pb.is_at_page_top(sp) {
        prev_between_below
    } else {
        0.0
    };
    state.pb.slot_top -= inter_gap;

    for bookmark in &para.bookmarks {
        state.bookmark_positions.insert(
            bookmark.clone(),
            (state.pb.all_contents.len(), state.pb.slot_top),
        );
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
    label_x = list_label::label_x(para, col_x, col_w, ctx.fonts);

    // Re-apply float zone adjustment after potential column change
    let first_line_top = state.pb.slot_top - inter_gap;
    if let Some(ref fz) = state.pb.float_zone {
        fz.narrow_paragraph(
            first_line_top,
            col_x,
            col_w,
            para,
            &mut para_text_x,
            &mut para_text_width,
            &mut label_x,
            ctx.compat_mode < 15,
        );
    }
    // A centred or right-aligned paragraph carries its label with the text:
    // Word lays out label, tab gap and first line as one left-aligned unit
    // and shifts the unit by the line's slack (issue #145 probe: bullet at
    // 244.7, text at 266.6 on a 453.85pt measure, the 21.85pt gap the same
    // as the left-aligned control's); `cell_label_alignment_offset` is the
    // same rule for cells. ponytail: left-to-right paragraphs only.
    if !para.bidi {
        let first_width = lines.first().map_or(0.0, |l| l.left_content_width());
        let slack = (para_text_width + text_hanging - first_width).max(0.0);
        label_x += align_offset(para.alignment, slack);
    }

    // A floating table pushed onto a fresh page hands its anchor paragraph the
    // page-body top so paragraph-relative shapes anchor there (the flow cursor
    // stays below the table). One-shot: consumed by this, the next, paragraph.
    let anchor = ParaAnchor::new(
        state.pb.take_float_anchor_for(block_idx),
        state.pb.anchor_slot_top(sp, ctx),
        own_space_before(applied_inter_gap, state.prev_space_after),
        sp,
        ctx,
    );
    // Hand the look-ahead float's anchor to the paragraph that actually
    // carries it only after this paragraph has taken its own.
    if let Some((anchor_top, _, anchor_block)) = lookahead {
        state.pb.pending_float_anchor = Some((anchor_top, Some(anchor_block)));
    } else if state
        .pb
        .pending_float_anchor
        .is_none_or(|(_, b)| b.is_none_or(|b| b <= block_idx))
    {
        // One-shot: an anchor meant for this paragraph (or any next one) is
        // spent; one for a later block rides through the paragraphs between.
        state.pb.pending_float_anchor = None;
    }

    // Margin anchors belong to the page that already started. A continuous
    // section's new margins apply on its next sheet, not to this page's float.
    let page_sp = &doc.sections[state.pb.page_hf_section].properties;
    let render_images = if state.pb.page_hf_section != sect_idx
        && (page_sp.margin_left != sp.margin_left || page_sp.margin_right != sp.margin_right)
    {
        let mut images = para.floating_images.clone();
        for fi in &mut images {
            if fi.h_relative_from == crate::model::HRelativeFrom::Margin
                && fi.wrap_type == WrapType::None
            {
                let x = fi.h_position.place(
                    page_sp.margin_left,
                    page_sp.text_width(),
                    fi.image.display_width,
                );
                fi.h_relative_from = crate::model::HRelativeFrom::Page;
                fi.h_position = crate::model::HorizontalPosition::Offset(x);
            }
        }
        std::borrow::Cow::Owned(images)
    } else {
        std::borrow::Cow::Borrowed(&para.floating_images)
    };

    // Floating images + textboxes, each into a per-shape buffer on its layer:
    // behind-doc ones paint under the page's text, the rest above it. Both
    // layers are sorted by relativeHeight in flush_page, so foreground images
    // interleave with foreground shapes by z-order rather than always
    // painting beneath them (annotation #191), and Word's z-order for
    // floating shapes spans paragraphs.
    let page = state.pb.all_contents.len();
    render_floating_images(
        &render_images,
        state.global_block_idx,
        floating_image_pdf_names,
        effect_floating_names,
        sp,
        col_x,
        col_w,
        text_width,
        anchor,
        &mut state.pb.behind_shapes,
        &mut state.pb.deferred_shapes,
        &mut state.pb.tags,
        page,
    );
    for tb in &para.textboxes {
        render_para_textbox(tb, para, anchor, sp, (col_x, col_w), text_width, ctx, &mut state.pb);
    }

    // Draw paragraph shading (background), extending outward to match borders
    if let Some(shd_color) = para.shading {
        // Up to the side borders' inner edges: stopping at their space left a
        // white seam inside english_town_council's blue section bars.
        let (shd_left_outset, shd_right_outset) = side_border_outsets(&para.borders, false);
        let (box_left_in, box_right_in) = border_box_indents(para);
        let shd_left = col_x + box_left_in - shd_left_outset;
        let shd_right = col_x + col_w - box_right_in + shd_right_outset;
        // Only a group member with the same fill closes the gap above:
        // renewable_dispatch's blue headings join the unshaded list item
        // above them in a (borderless) group, and Word leaves that gap white.
        let shd_top = state.pb.slot_top
            + if prev_borders_match && prev_para.is_some_and(|pp| pp.shading == para.shading) {
                applied_inter_gap
            } else {
                0.0
            }
            + prev_band_below;
        // min: a vanished paragraph mark zeroes content_h after `lines_h`.
        let shd_bottom_pad = if bottom_collapses {
            bottom_band.0
        } else {
            bdr_bottom_pad
        };
        let shd_bottom = state.pb.slot_top - bdr_top_pad - content_h.min(lines_h) - shd_bottom_pad;
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

    // Set FloatZone for wrapping floating images
    // (may already be set by self-wrapping above; overwrite
    // to ensure polygon data is included).
    // A float entirely outside the text column (e.g. a QR code in the left
    // margin) never narrows text — installing its zone would only clobber a
    // still-active in-column zone from an earlier paragraph's float.
    let mut zones = Vec::new();
    for fi in para
        .floating_images
        .iter()
        .filter(|fi| wraps_in_column(fi, sp, col_x, col_w, text_width))
    {
        let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
        let fi_y_top = anchor.fi_y_top(fi, sp);
        zones.push(FloatZone::for_float(fi, fi_x, fi_y_top));
        if !text_empty && leading_full_width_picture(fi, sp, col_x, col_w, text_width) {
            let bottom = fi_y_top - fi.image.display_height - fi.dist_bottom;
            state.pb.next_text_below_picture = Some(
                state
                    .pb
                    .next_text_below_picture
                    .map_or(bottom, |old| old.min(bottom)),
            );
        }
    }
    zones.extend(textbox_zones(para, sp, (col_x, col_w), text_width, anchor.top(false), ctx));
    state.pb.set_float_zones(zones);

    if debug_wrap && let Some(ref fz) = state.pb.float_zone {
        draw_debug_wrap_overlay(&mut state.pb.content, fz);
    }

    for conn in &para.connectors {
        // Same page-level z-stack as textboxes — anchored connectors must
        // interleave with shapes by relativeHeight (e.g. letter strokes
        // drawn over gradient circles)
        let mut shape_content = tagging::artifact_content();
        let (x, y) =
            positioning::connector_top_left(conn, sp, col_x, col_w, text_width,
                    table_connector_anchor.filter(|(page, _)| *page == state.pb.page_count())
                        .map_or(anchor.top(false), |(_, top)| top));
        render_connector(conn, &mut shape_content, x, y);
        state.pb.deferred_shapes.push((conn.z_index, shape_content));
    }

    for diagram in &para.floating_smartart {
        let Some(anchor) = &diagram.anchor else {
            continue;
        };
        let x = positioning::resolve_h_position(
            anchor.h_relative_from,
            &anchor.h_position,
            anchor.width,
            sp,
            col_x,
            col_w,
            text_width,
        );
        let y = header_footer::resolve_tb_y_top(
            anchor.v_relative_from,
            &anchor.v_position,
            diagram.display_height,
            sp,
            state.pb.slot_top,
        );
        // A content-less Figure after the paragraph, as for an inline diagram.
        let alt = smartart_alt(std::slice::from_ref(diagram));
        state.pb.tags.hoist_figure(alt.as_deref(), 0);
        let mut shape_content = tagging::artifact_content();
        smartart::render_smartart(
            &mut shape_content,
            diagram,
            x,
            y,
            ctx.fonts,
            smartart_font_key,
            smartart_image_names,
        );
        state
            .pb
            .deferred_shapes
            .push((anchor.z_index, shape_content));
    }

    if let Some(ref ic) = para.inline_chart {
        let chart_x = col_x + align_offset(para.alignment, (col_w - ic.display_width).max(0.0));
        state.pb.figure_without_content(para, doc, None);
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
        let alt = smartart_alt(&para.smartart);
        state.pb.figure_without_content(para, doc, alt.as_deref());
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
                smartart_image_names,
            );
        }
    } else if let Some(ref hr) = para.horizontal_rule {
        let line_bottom = state.pb.slot_top - content_h;
        draw_horizontal_rule(&mut state.pb.content, para, hr, col_x, col_w, line_bottom);
    } else if para.image.is_some() && para.content_height > 0.0 {
        if let Some(pdf_name) = image_pdf_names.get(&state.global_block_idx) {
            let img = para.image.as_ref().unwrap();
            // Decorative pictures stay artifacts, as in Word's export.
            if img.decorative {
                // The paragraph mark still has its element (a heading's H3).
                state.pb.tag_empty_para(para, doc);
            } else {
                state.pb.begin_figure(para, doc, img.alt.as_deref());
            }
            // A picture shorter than the line sits, effect extents included, on
            // a baseline one natural line of the mark's font down, whatever the
            // spacing (Word probe: 2.25-12pt pictures under Arial 12 at single
            // and 1.5 lines).
            let turned = img.rotation_deg.abs() > 0.01;
            let swap_axes = ctx.compat_mode >= 15
                && (img.rotation_deg.rem_euclid(360.0) + 45.0).rem_euclid(180.0) >= 90.0;
            let (box_w, box_h) = if swap_axes {
                (img.display_height, img.display_width)
            } else {
                (img.display_width, img.display_height)
            };
            let image_height = box_h + img.layout_extra_height;
            let bottom_depth = if image_height + inserted_descent > line_h {
                img.layout_extra_top + box_h
            } else {
                let baseline =
                    (image_height + inserted_descent).max(natural_line_h) - inserted_descent;
                baseline - (img.layout_extra_height - img.layout_extra_top)
            };
            let box_bottom = state.pb.slot_top - bottom_depth;
            // Keep the picture paragraph's existing first-line and side indents.
            let first = if para.list_label.is_empty() {
                para.indent_first_line - para.indent_hanging
            } else {
                0.0
            };
            let start = para.indent_left + first;
            let room = col_w - start - para.indent_right - box_w;
            let box_x = col_x + start + align_offset(para.alignment, room.max(0.0));
            let x = box_x + (box_w - img.display_width) / 2.0;
            let y_bottom = box_bottom + (box_h - img.display_height) / 2.0;
            if turned {
                state.pb.content.save_state();
                // Legacy inline layout clips a turned picture to its original
                // frame instead of growing the occupied line box.
                if ctx.compat_mode < 15 {
                    state.pb.content.rect(box_x, box_bottom, box_w, box_h);
                    state.pb.content.clip_nonzero();
                    state.pb.content.end_path();
                }
                positioning::push_center_rotation(
                    &mut state.pb.content,
                    box_x,
                    box_bottom,
                    box_w,
                    box_h,
                    img.rotation_deg,
                );
            }
            let img_fx = effect_names.get(&state.global_block_idx);
            if let Some(ref shadow) = img.shadow {
                color::draw_image_shadow(
                    &mut state.pb.content,
                    shadow,
                    x,
                    y_bottom,
                    img.display_width,
                    img.display_height,
                    img_fx.and_then(|fx| fx.shadow.as_deref()),
                );
            }
            if let Some(ref glow) = img.glow {
                color::draw_image_glow(
                    &mut state.pb.content,
                    glow,
                    x,
                    y_bottom,
                    img.display_width,
                    img.display_height,
                    img_fx.and_then(|fx| fx.glow.as_deref()),
                );
            }
            smartart::render_image_with_clip(
                &mut state.pb.content,
                pdf_name,
                x,
                y_bottom,
                img.display_width,
                img.display_height,
                img.clip_geometry.as_ref(),
            );
            if let Some(sc) = img.stroke_color {
                smartart::stroke_image_border(
                    &mut state.pb.content,
                    x,
                    y_bottom,
                    img.display_width,
                    img.display_height,
                    sc,
                    img.stroke_width,
                    img.clip_geometry.as_ref(),
                );
            }
            // Post-image effects: inner shadow, reflection (drawn on top / below)
            if let Some(ref inner) = img.inner_shadow {
                color::draw_inner_shadow(
                    &mut state.pb.content,
                    inner,
                    x,
                    y_bottom,
                    img.display_width,
                    img.display_height,
                    img_fx.and_then(|fx| fx.inner_shadow.as_deref()),
                );
            }
            if let Some(ref refl) = img.reflection {
                color::draw_reflection(
                    &mut state.pb.content,
                    refl,
                    x,
                    y_bottom,
                    img.display_width,
                    img.display_height,
                    img_fx.and_then(|fx| fx.reflection.as_deref()),
                );
            }
            if turned {
                state.pb.content.restore_state();
            }
            // Markup view: the inserted picture's revision underline, the run's
            // post-table geometry on a 0.25pt grid like the text path's.
            if inserted_descent > 0.0
                && let Some(run) = para.runs.iter().find(|r| r.revised && r.underline)
                && let Some((offset, thick)) = ctx
                    .fonts
                    .get(&font_key(run))
                    .and_then(|e| layout::post_underline(e, run.font_size))
            {
                color::fill_color_or_black(&mut state.pb.content, run.color);
                state
                    .pb
                    .content
                    .rect(box_x, box_bottom - offset - thick, box_w, thick)
                    .fill_nonzero();
            }
            if !img.decorative {
                state.pb.end_tag();
            }
        } else if para.image.is_some() {
            state
                .pb
                .content
                .set_fill_gray(0.5)
                .rect(col_x, state.pb.slot_top - content_h, col_w, content_h)
                .fill_nonzero()
                .set_fill_gray(0.0);
        }
    } else if lines.is_empty()
        && para.list_label.is_empty()
        && para.mark_revised
        && !para.paragraph_mark_vanish
    {
        layout::set_change_bar_extension(
            own_space_before(applied_inter_gap, state.prev_space_after),
            effective_space_after,
        );
        // Markup view: an empty paragraph with a revised mark gets its bar
        // (Word's bar runs unbroken over a decree's empty inserted paragraphs).
        let top = state.pb.slot_top - bdr_top_pad;
        layout::record_change_bar(top - content_h, top);
    } else if !lines.is_empty() || !para.list_label.is_empty() {
        // An empty numbered paragraph still shows its label: uk_commercial's
        // "SCHEDULE 1" headings are empty paragraphs whose style numbers them.
        // When the document grid snaps line heights, align the first
        // baseline one linePitch below the slot top so text sits on
        // the grid rather than at a font-metric-dependent offset.
        let baseline_offset = if grid_snapped {
            grid_baseline
        } else {
            first_baseline_offset
        };
        let baseline_y = state.pb.slot_top - bdr_top_pad - baseline_offset - text_drop;
        layout::set_change_bar_extension(
            own_space_before(applied_inter_gap, state.prev_space_after),
            effective_space_after,
        );

        let tags = state.pb.para_tags(para, doc);
        state.pb.begin_para_tags(tags, |content| {
            render_list_label(
                content,
                para,
                ctx.fonts,
                label_x,
                baseline_y - first_line_drop,
                font_size,
            )
        });

        if !lines.is_empty() {
            render_paragraph_lines(
                &mut state.pb.content,
                &lines,
                &para.alignment,
                para_text_x,
                para.indent_left,
                para_text_width,
                baseline_y,
                line_h,
                para_metrics,
                lines.len(),
                0,
                &mut state.pb.links,
                text_hanging,
                ctx.fonts,
                poly_line_geom.as_deref(),
                &mut state.pb.gradient_specs,
                Some(&mut state.pb.comment_anchors),
                ln_cfg.map(|(start, count_by, right_x)| LineNumberArg {
                    counter: &mut state.pb.line_number_counter,
                    start,
                    count_by,
                    right_x,
                    font: ln_font.clone(),
                }),
                Some(LinkTagger::new(
                    &mut state.pb.tags,
                    state.pb.all_contents.len(),
                    tags.1,
                )),
                None,
            );
        }
        state.pb.end_tag();
    } else {
        // Word tags empty paragraphs too; keeping them keeps the P sequence aligned.
        state.pb.tag_empty_para(para, doc);
    }

    // Draw paragraph borders — left/right borders extend outward
    // from the text area so text inside stays aligned with text outside
    {
        let bdr = &para.borders;
        let box_top = state.pb.slot_top - bdr_top_half_band;
        let box_bottom = state.pb.slot_top - bdr_top_pad - content_h - bdr_bottom_pad;
        // Word puts a side border's inner edge its space plus 1.47pt (left) or
        // 1.73pt (right) outside the text, whatever its width (online export
        // probes: sz 4/12/24 × space 0/4/12, with and without indents).
        let (bdr_left_outset, bdr_right_outset) = side_border_outsets(bdr, true);
        let (box_left_in, box_right_in) = border_box_indents(para);
        let box_left = col_x + box_left_in - bdr_left_outset;
        let box_right = col_x + col_w - box_right_in + bdr_right_outset;

        // Extend horizontal borders past the corners so they cover the
        // corner gap left by butt-capped vertical border strokes.
        let (h_left_ext, h_right_ext) = horizontal_border_ext(bdr);
        // `inward`: +1 for a border below the text, -1 for one above.
        let draw_h_border = |content: &mut Content, b: &ParagraphBorder, y: f32, inward: f32| {
            let (x0, x1) = (box_left - h_left_ext, box_right + h_right_ext);
            stroke_border(content, b, (x0, y), (x1, y), inward);
        };
        // When this paragraph continues a border group, extend vertical
        // borders upward through the inter-paragraph gap so there is no
        // visible break between consecutive paragraphs' left/right borders.
        let v_border_top = box_top
            + if prev_borders_match {
                applied_inter_gap
            } else {
                0.0
            };
        // `inward`: +1 for the left border, -1 for the right.
        // Down through the band of a `between` border the group goes on below.
        let v_border_bottom = box_bottom
            - if bottom_collapses {
                bdr_bottom_extent - bdr_bottom_pad
            } else {
                0.0
            };
        let draw_v_border = |content: &mut Content, b: &ParagraphBorder, x: f32, inward: f32| {
            stroke_border(content, b, (x, v_border_top), (x, v_border_bottom), inward);
        };

        if !prev_borders_match && let Some(b) = &bdr.top {
            draw_h_border(&mut state.pb.content, b, box_top, -1.0);
        }
        if let Some(b) = bottom_bdr {
            draw_h_border(&mut state.pb.content, b, box_bottom, 1.0);
        }
        if let Some(b) = &bdr.left {
            draw_v_border(&mut state.pb.content, b, box_left, 1.0);
        }
        if let Some(b) = &bdr.right {
            draw_v_border(&mut state.pb.content, b, box_right, -1.0);
        }
    }

    // Word numbers a paragraph without text lines too (an empty anchor under
    // its boxes, a chart or a picture), on its mark line at the foot.
    if lines.is_empty()
        && content_h > 0.0
        && let Some((start, count_by, right_x)) = ln_cfg
    {
        // ponytail: no grid snap; a grid-snapped empty anchor's number may sit off its line.
        // A picture's, chart's or inline shape's bottom sits on its line's
        // baseline.
        let picture_bottom = if para.image.is_some() || para.inline_chart.is_some() {
            Some(para.content_height)
        } else {
            para.textboxes
                .iter()
                .filter(|tb| tb.on_line)
                .map(crate::model::Textbox::flow_bottom)
                .reduce(f32::max)
        };
        let mark_baseline = if let Some(bottom) = picture_bottom {
            state.pb.slot_top - bdr_top_pad - bottom
        } else {
            state.pb.slot_top - bdr_top_pad - content_h + (line_h - first_baseline_offset).max(0.0)
        };
        let fonts = fonts::by_pdf_name(ctx.fonts);
        LineNumberArg {
            counter: &mut state.pb.line_number_counter,
            start,
            count_by,
            right_x,
            font: ln_font.clone(),
        }
        .number(
            &mut state.pb.content,
            0.0,
            mark_baseline,
            || None,
            &fonts,
            None,
        );
    }

    state.pb.slot_top -= content_h + bdr_top_pad + bdr_bottom_extent;
    state.prev_between_below = if bottom_collapses {
        between_space(para)
    } else {
        0.0
    };
    if state.pb.slot_top < state.effective_margin_bottom - 1.0 {
        log::warn!(
            "Body overflow: slot_top={:.2} < eff_margin_bottom={:.2} after paragraph on page {}",
            state.pb.slot_top,
            state.effective_margin_bottom,
            state.pb.all_contents.len(),
        );
    }
    if !(text_empty && para.paragraph_mark_vanish) {
        state.prev_space_after = effective_space_after;
        layout::set_change_bar_extension(0.0, 0.0);
    }

    // Track footnotes referenced on this page
    for run in para.runs.iter() {
        if let Some(id) = run.footnote_id {
            track_page_footnote(state, doc, ctx, col_geometry[state.current_col], id);
        }
        if let Some(id) = run.endnote_id {
            state.pb.track_endnote(id);
        }
    }

    update_styleref_from_para(
        &mut state.pb.styleref_running,
        &mut state.pb.styleref_page_first,
        para,
        &doc.style_id_to_name,
    );

    if para.page_break_after {
        state
            .pb
            .begin_next_page(sect_idx, sp, &mut state.effective_margin_bottom, ctx);
        state.prev_space_after = 0.0;
        state.current_col = 0;
    }
    if para.column_break_after {
        state.pb.advance_column_or_page(
            &mut state.current_col,
            col_count,
            sect_idx,
            sp,
            &mut state.effective_margin_bottom,
            ctx,
        );
        state.prev_space_after = 0.0;
    }

    false
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
    let (pdf, page_fields) = render_pass(doc, None, options)?;
    // Word shows body PAGE, NUMPAGES and SECTIONPAGES fields live, not their cached result,
    // and their values are known only once the document is paginated.
    let has_page_fields =
        body_runs(doc).any(|r| r.field_code.as_ref().is_some_and(FieldCode::is_page_count));
    if has_page_fields {
        return render_pass(doc, Some(&page_fields), options).map(|(pdf, _)| pdf);
    }
    Ok(pdf)
}

fn render_pass(
    doc: &Document,
    page_fields: Option<&BodyPageFields>,
    options: &ConvertOptions,
) -> Result<(Vec<u8>, BodyPageFields), Error> {
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
        effect_table_names,
    } = embed_all_images(
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
        table_cell_image_names: &table_cell_image_names,
        effect_table_names: &effect_table_names,
        textbox_image_names: &textbox_image_names,
        chart_font_name: &doc.theme_minor_font,
        compress_punctuation: doc.compress_punctuation,
        footnote_marks: &footnote_display_order,
        endnote_marks: &endnote_display_order,
        body_page_fields: page_fields,
        page_labels: Vec::new(),
        compat_mode: doc.compat_mode,
        open_type_features: doc.open_type_features,
        east_asian_lang: doc.east_asian_lang,
        do_not_expand_shift_return: doc.do_not_expand_shift_return,
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
        cell_comment_anchors: Default::default(),
        hf_heights: Default::default(),
    };

    let (bookmark_positions, page_labels) = compute_bookmark_positions(doc, &ctx);
    ctx.page_labels = page_labels;

    // Phase 2: build multi-page content streams (section-aware)
    let first_sp = &doc.sections[0].properties;
    let mut cur_sp = first_sp;
    let initial_slot_top = effective_slot_top(cur_sp, true, 0, &ctx);
    let mut state = LayoutState {
        pb: PageBuilder::new(initial_slot_top),
        prev_space_after: 0.0,
        prev_between_below: 0.0,
        effective_margin_bottom: compute_effective_margin_bottom(cur_sp, true, 0, &ctx),
        current_col: 0,
        global_block_idx: 0,
        heading_entries: Vec::new(),
        bookmark_positions,
        long_keep_chain: false,
        reflowed: false,
    };
    state.pb.tags.lang = document_lang(doc);

    for (sect_idx, section) in doc.sections.iter().enumerate() {
        let sp = &section.properties;
        ctx.cell_grid_pitch.set(cell_grid_pitch(doc, sp));

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
                                        page_numbers(doc, &state.pb.page_section_indices)
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

                    state.pb.slot_top = effective_slot_top(sp, true, state.pb.page_count(), &ctx);
                    state.pb.column_top_y = state.pb.slot_top;
                    state.pb.page_top_y = state.pb.slot_top;
                    state.effective_margin_bottom =
                        compute_effective_margin_bottom(sp, true, state.pb.page_count(), &ctx);
                    state.pb.page_hf_section = sect_idx;
                    state.pb.is_first_page_of_section = true;
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
                        state.pb.page_hf_section = sect_idx;
                        state.pb.is_first_page_of_section = true;
                        state.pb.slot_top =
                            effective_slot_top(sp, true, state.pb.page_count(), &ctx);
                        state.pb.column_top_y = state.pb.slot_top;
                        state.pb.page_top_y = state.pb.slot_top;
                        state.effective_margin_bottom =
                            compute_effective_margin_bottom(sp, true, state.pb.page_count(), &ctx);
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
        // If the first ordinary paragraph itself moves to the next sheet,
        // that sheet is the continuous section's first page. Once it flows
        // here, a later break gets the default variant as before. Tables,
        // lifted frames and drawing paragraphs keep their existing handling.
        state.pb.pending_section_first_page =
            (matches!(sp.break_type, SectionBreakType::Continuous)
                && sp.columns.as_ref().is_none_or(|c| c.columns.len() <= 1)
                && state.pb.float_zone.is_none()
                && section.blocks.first().is_some_and(|b| {
                    matches!(b,
                        Block::Paragraph(p) if p.frame_props.is_none()
                            && p.image.is_none() && p.inline_chart.is_none()
                            && p.floating_images.is_empty() && p.textboxes.is_empty()
                            && p.smartart.is_empty() && p.connectors.is_empty()
                            && p.runs.iter().all(|r| r.inline_image.is_none())
                    )
                }))
            .then_some((sect_idx, state.pb.page_count(), state.pb.slot_top));
        let text_width = sp.text_width();

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
        // to look for, and past the last paragraph with one it finds nothing.
        let last_band_block = (col_count == 1)
            .then(|| {
                section.blocks.iter().rposition(|b| {
                    matches!(b, Block::Paragraph(p)
                        if has_bands(p, sp, col_geometry[0], text_width))
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
            let mut last_block_consumption = None;
            for (block_idx, block) in section.blocks.iter().enumerate() {
                let previous_block_consumption = last_block_consumption.take();
                let block_start = (state.pb.page_count(), state.pb.slot_top);
                let block_frame = lifted_frame(block);
                if let Some(f) = frame.take_if(|f| block_frame.is_none_or(|fp| fp != f.props)) {
                    f.close(state, sp);
                }
                if frame.is_none() {
                    match block_frame {
                        Some(fp) => {
                            let blocks = &section.blocks[block_idx..];
                            frame = Some(OpenFrame::open(fp, blocks, state, &ctx, sp));
                        }
                        None => {
                            let col = col_geometry[state.current_col];
                            let page = (sect_idx, state.pb.page_count());
                            if let Some(last) = last_band_block.filter(|&l| block_idx <= l)
                                && state.pb.bands_prescanned != Some(page)
                            {
                                state.pb.bands_prescanned = Some(page);
                                let rest = &section.blocks[block_idx..=last];
                                prescan_page_bands(state, rest, sp, col, text_width, &ctx);
                            }
                            step_below_bands(state, block, sp, col, text_width, &ctx);
                        }
                    }
                }
                let (geometry, cols, width) = match &frame {
                    Some(f) => (std::slice::from_ref(&f.geometry), 1, f.geometry.1),
                    None => (&col_geometry[..], col_count, text_width),
                };

                // A table may move the text below its wrap zone while the
                // paragraph-relative connector retains the original anchor.
                let table_connector_anchor = state.pb.float_zone.as_ref()
                    .filter(|zone| zone.from_table)
                    .map(|_| (state.pb.page_count(), state.pb.slot_top));

                // With no side strip, the whole empty clearing paragraph sits
                // below the floating table, including the line ended by br.
                if matches!(block, Block::Paragraph(p) if p.clears_floats && p.runs.iter().all(|r| r.is_line_break || (r.text.is_empty() && !r.is_tab && r.inline_image.is_none())))
                    && block_idx
                        .checked_sub(1)
                        .and_then(|i| section.blocks.get(i))
                        .is_some_and(|b| matches!(b, Block::Table(t) if t.position.is_some()))
                    && let Some(ref zone) = state.pb.float_zone
                {
                    let (x, w) = col_geometry[state.current_col];
                    let gap = side_strip(
                        (zone.obj_left, zone.obj_right),
                        (zone.left_from_text, zone.right_from_text),
                        (x, w),
                    );
                    if gap < MIN_EMPTY_STRIP && state.pb.slot_top > zone.bottom_y {
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
                        None => state.pb.slot_top <= fz.top_y
                            || (fz.para_relative && state.pb.slot_top <= fz.top_y + 30.0)
                            // A floating table meets a paragraph whose first
                            // line reaches its top: physical_education's empty
                            // anchor 0.05pt above a full-width table goes below
                            // it with its line and space after. Before a table
                            // anchored to the page, a line that starts above the
                            // top keeps its place even where it reaches into the
                            // table (japanese_hr_consultation's 10pt line 7pt above
                            // the table overlaps it; Word probes 2026-10-23 with
                            // the table top 2pt and 8pt into the line kept it, 2pt
                            // above its top pushed it). Text-anchored tables keep
                            // the reach rule (case32, case174, east_asia_conference).
                            || (fz.from_table
                                && matches!(block, Block::Paragraph(p) if {
                                    let (fs, lhr, _) = tallest_run_metrics(&p.runs, ctx.fonts);
                                    let ls = p.line_spacing.unwrap_or(ctx.doc_line_spacing);
                                    // An indented empty anchor can enter the float
                                    // through its before-spacing, before its line starts.
                                    let before = if ctx.compat_mode >= 15
                                        && (p.indent_left != 0.0 || p.indent_right != 0.0)
                                        && is_text_empty(&p.runs)
                                    {
                                        state.prev_space_after.max(p.space_before)
                                    } else { 0.0 };
                                    let reach = if fz.table_on_page {
                                        0.0
                                    } else {
                                        resolve_line_h(ls, fs, lhr)
                                    };
                                    state.pb.slot_top - before - reach < fz.top_y + 0.1
                                })),
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
                        let (space_left, space_right) = match block {
                            Block::Paragraph(p)
                                if ctx.compat_mode >= 15 && is_text_empty(&p.runs) =>
                            {
                                let left = col_x + p.indent_left;
                                let right = col_x + col_w - p.indent_right;
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
                                let room = if fz.wrap_text == WrapText::BothSides {
                                    space_left.max(space_right)
                                } else {
                                    side_room(fz.wrap_text, space_left, space_right)
                                };
                                room >= w.max(MIN_EMPTY_STRIP)
                            }
                            None if fz.wrap_text == WrapText::BothSides => {
                                // For bothSides, check combined width of both regions.
                                // A float overhanging a margin has no strip on that
                                // side, not a negative one (an external tracked corpus
                                // cover 71.5pt into the left margin left 147pt on the
                                // right, and the title went below the page-tall picture).
                                (space_left.max(0.0) + space_right.max(0.0)) >= 72.0
                            }
                            None => side_room(fz.wrap_text, space_left, space_right) >= 72.0,
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
                                state.pb.slot_top = rect_clear
                                    .as_ref()
                                    .map_or(fz.bottom_y, |rect| rect.cleared_y(state.pb.slot_top));
                                state.pb.next_float_zone();
                            }
                        }
                        // Otherwise leave zone active — paragraph layout adjusts width
                    }
                }

                // Clearing a floating table may move the first body line into
                // a later frame band. Recheck the existing bands at that cursor.
                if frame.is_none() && state.pb.slot_top != before_float_clearance {
                    let col_w = col_geometry[state.current_col].1;
                    step_below_existing_bands(state, block, sp, col_w, &ctx);
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
                            &floating_image_pdf_names,
                            &inline_image_pdf_names,
                            &image_pdf_names,
                            &effect_names,
                            &effect_floating_names,
                            &effect_inline_names,
                            doc,
                            smartart_font_key,
                            &smartart_image_names,
                            debug_wrap,
                            table_connector_anchor,
                        );
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
                        if table.position.as_ref().is_some_and(|p| p.v_anchor == "text" && p.v_offset_pt < 0.0)
                            && block_idx.checked_sub(1).and_then(|i| section.blocks.get(i)).is_some_and(|b|
                                matches!(b, Block::Paragraph(p) if is_text_empty(&p.runs)
                                    && p.content_height == 0.0 && p.image.is_none()
                                    && !p.is_section_break && p.frame_props.is_none()))
                            && let Some((page, consumed)) = previous_block_consumption
                            && page == state.pb.page_count() && consumed > 0.0
                            && let Some(zone) = state.pb.float_zone.as_mut()
                        {
                            let (x, w) = geometry[state.current_col];
                            let gap = (zone.obj_left - zone.left_from_text - x)
                                .max(x + w - zone.obj_right - zone.right_from_text);
                            if gap < MIN_EMPTY_STRIP { zone.bottom_y -= consumed; }
                        }
                        if let Some(mut tags) = state.pb.table_tags.take() {
                            tags.finish(&mut state.pb.tags);
                        }
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
                    let mark_pitch = resolve_line_h(ls, fs, lhr);
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
                        let section_mark = ctx.compat_mode < 15 && fz.from_table
                            && empty_clear && p.is_section_break && block_idx > 0
                            && doc.sections.get(sect_idx + 1).is_some_and(|next|
                                next.properties.break_type == SectionBreakType::Continuous);
                        let mark_lines = usize::from(!section_mark) + extra_breaks;
                        state.pb.slot_top = float_bottom - mark_pitch * mark_lines as f32;
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
            (trial.pb.page_count() - base, total)
        };
        let natural = balance.then(|| trial_at(&state, 0, state.effective_margin_bottom));
        // ponytail: every trial lays the whole region out again, so a long
        // region costs a few full layouts; start trials at its last page if slow.
        let balanced_h = if let Some((flushes, total)) = natural {
            let (floor_page, page_bottom, top) = if flushes == 0 {
                (0, state.effective_margin_bottom, state.pb.column_top_y)
            } else {
                (
                    state.pb.page_count() + flushes,
                    compute_effective_margin_bottom(
                        sp,
                        false,
                        state.pb.page_count() + flushes,
                        &ctx,
                    ),
                    effective_slot_top(sp, false, state.pb.page_count() + flushes, &ctx),
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
                if trial_at(&state, flushes, top - h + lead + 0.01).0 == flushes {
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
    state.pb.flush_page(doc.sections.len() - 1);
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
        draw_change_bars(content, bars, &doc.sections[si].properties, doc.mirror_margins, page_idx);
    }

    // Phase 2c: render footnotes at page bottom (above footer). Endnotes
    // (default pos=docEnd) flow inline after the last body block on the final
    // page (see render_endnotes_inline), NOT pinned to the bottom: their
    // separator paragraph starts at the cursor below the last paragraph and
    // its space_after.
    let last_page_idx = state.pb.all_contents.len().saturating_sub(1);
    let endnote_top_y = state.pb.slot_top - state.prev_space_after;
    for (page_idx, content) in state.pb.all_contents.iter_mut().enumerate() {
        let (hf_si, is_first, si) = state.pb.page_section_indices[page_idx];
        let sp = &doc.sections[hf_si].properties;
        let eff_bottom = compute_effective_margin_bottom(sp, is_first, page_idx, &ctx);
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
        draw_change_bars(content, &note_bars, content_sp, doc.mirror_margins, page_idx);
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
    for ((s, t, pi), name) in &hf_image_names {
        hf_maps_index
            .entry((*s, *t))
            .or_default()
            .0
            .insert(*pi, name.clone());
    }
    for ((s, t, pi, ri), name) in &hf_inline_image_names {
        hf_maps_index
            .entry((*s, *t))
            .or_default()
            .1
            .insert((*pi, *ri), name.clone());
    }
    for ((s, t, pi, fi), name) in &hf_floating_image_names {
        hf_maps_index
            .entry((*s, *t))
            .or_default()
            .2
            .insert((*pi, *fi), name.clone());
    }
    for ((s, t, pi), fx) in &effect_hf_names {
        hf_maps_index
            .entry((*s, *t))
            .or_default()
            .3
            .insert(*pi, fx.clone());
    }
    let empty_hf_maps: HfMaps = Default::default();

    let page_numbers = page_number_formats(doc, &state.pb.page_section_indices);
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

        let (header, hdr_type, hdr_si) = resolve_header_for_page(doc, si, is_first, page_num);
        if let Some(header_data) = header {
            let (pi_map, ii_map, fi_map, sh_para) = hf_maps_index
                .get(&(hdr_si, hdr_type))
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
                header_data,
                &ctx,
                sp,
                true,
                &pc,
                &mut state.pb.all_gradient_specs[page_idx],
                &mut state.pb.all_links[page_idx],
            );
            has_hf = true;
        }

        let (footer, ftr_type, ftr_si) = resolve_footer_for_page(doc, si, is_first, page_num);
        if let Some(footer_data) = footer {
            let (pi_map, ii_map, fi_map, sh_para) = hf_maps_index
                .get(&(ftr_si, ftr_type))
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
                footer_data,
                &ctx,
                sp,
                false,
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
            let region_bottom = compute_effective_margin_bottom(sp, is_first, page_idx, &ctx);
            let slack = state.pb.all_content_bottom[page_idx] - region_bottom;
            if slack > 0.0 { slack * frac } else { 0.0 }
        })
        .collect();

    assemble_pdf_pages(
        &mut pdf,
        &mut alloc,
        catalog_id,
        pages_id,
        valign_offsets,
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
        &image_xobjects,
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
    Ok((objstm::pack(pdf.finish()), page_fields))
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
    fn a_trailing_anchor_has_only_blank_runs_after_it() {
        let text = |t: &str| crate::model::Run { text: t.to_string(), ..Default::default() };
        let brk = crate::model::Run { is_line_break: true, ..Default::default() };
        let runs = [text("Med forbehold"), text(" "), brk, text("")];
        assert!(blank_after(&runs, 4));
        assert!(blank_after(&runs, 3));
        assert!(blank_after(&runs, 9), "an index past the end is a trailing anchor");
        assert!(!blank_after(&runs, 2), "a break after the anchor is a line");
        assert!(!blank_after(&runs, 0));
    }

    #[test]
    fn contour_clearance_keeps_the_natural_line_phase() {
        let rect = RectangularClearance {
            top: 80.0,
            bottom: 15.0,
            pitch: 10.0,
            line_h: 15.0,
        };
        assert_eq!(rect.cleared_y(100.0), 10.0);
        assert_eq!(rect.cleared_y(98.0), 8.0);
        assert_eq!(rect.cleared_y(15.0), 15.0);
        assert_eq!(rect.cleared_y(10.0), 10.0);
    }

    #[test]
    fn rectangle_bounds_reject_diagonal_and_nonrectangular_contours() {
        let rect = [
            (0.0, 0.0),
            (0.0, 20.0),
            (30.0, 20.0),
            (30.0, 0.0),
            (0.0, 0.0),
        ];
        assert_eq!(
            rectangular_polygon_bounds(&rect),
            Some((0.0, 30.0, 20.0, 0.0))
        );
        assert_eq!(
            rectangular_polygon_bounds(&rect[..4]),
            Some((0.0, 30.0, 20.0, 0.0))
        );
        assert!(
            rectangular_polygon_bounds(&[(0.0, 0.0), (30.0, 20.0), (0.0, 20.0), (30.0, 0.0)])
                .is_none()
        );
        assert!(
            rectangular_polygon_bounds(&[(0.0, 0.0), (0.0, 20.0), (15.0, 10.0), (30.0, 0.0)])
                .is_none()
        );
    }

    #[test]
    fn a_break_line_that_does_not_fit_leaves_a_blank_page() {
        use std::io::{Cursor, Write};
        let pages = |break_twips: u32| {
            let exact = |t: u32| format!(r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="{t}" w:lineRule="exact"/></w:pPr>"#);
            // 25 exact 25pt lines leave 23pt of a 648pt body.
            let lines = (0..25)
                .map(|i| format!("<w:p>{}<w:r><w:t>Line {i}</w:t></w:r></w:p>", exact(500)))
                .collect::<String>();
            let xml = format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{lines}<w:p>{}<w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#,
                exact(break_twips)
            );
            let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
            zip.start_file("word/document.xml", zip::write::SimpleFileOptions::default())
                .unwrap();
            zip.write_all(xml.as_bytes()).unwrap();
            let doc = crate::docx::parse_bytes(
                &zip.finish().unwrap().into_inner(),
                crate::RevisionView::Final,
            )
            .unwrap();
            render_pass(&doc, None, &ConvertOptions::default()).unwrap().1.total
        };
        assert_eq!(pages(500), 3); // p29a
        assert_eq!(pages(300), 2); // p29c
    }

    #[test]
    fn space_only_wrapping_frame_is_empty_but_content_controls_keep_flow() {
        let mut doc = crate::docx::parse_bytes(
            include_bytes!("../../tests/fixtures/cases/space-only-frame/input.docx"),
            crate::RevisionView::Final,
        )
        .unwrap();
        let section = &mut doc.sections[0];
        let Block::Paragraph(para) = &mut section.blocks[0] else {
            panic!("frame paragraph expected");
        };
        assert!(is_empty_wrapping_frame(para));
        para.runs[0].text = "   ".into();
        assert!(is_empty_wrapping_frame(para));
        assert!(!is_text_empty(&para.runs)); // Ordinary paragraphs still flow.
        for text in ["\u{a0}", "\t", "x"] {
            para.runs[0].text = text.into();
            assert!(!is_empty_wrapping_frame(para));
        }
        para.runs[0].text = " ".into();
        para.runs[0].is_tab = true;
        assert!(!is_empty_wrapping_frame(para));
        para.runs[0].is_tab = false;
        para.runs[0].is_line_break = true;
        assert!(!is_empty_wrapping_frame(para));
        para.runs[0].is_line_break = false;
        para.frame_props.as_mut().unwrap().height = 20.0;
        assert!(!is_empty_wrapping_frame(para));
        para.frame_props.as_mut().unwrap().height = 0.0;
        para.frame_props.as_mut().unwrap().text_below = true;
        assert!(!is_empty_wrapping_frame(para));
    }

    #[test]
    fn continuous_first_page_is_pending_only_before_any_paragraph_flow() {
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
    fn keep_with_next_needs_an_unsplittable_paragraph_whole() {
        assert_eq!(lines_kept_together(false, || 5), 1);
        assert_eq!(lines_kept_together(true, || 1), 1);
        assert_eq!(lines_kept_together(true, || 3), 3);
        assert_eq!(lines_kept_together(true, || 4), 2);
    }

    #[test]
    fn unsized_line_takes_its_break_else_the_mark() {
        let mut para = Paragraph {
            paragraph_mark_font_size: Some(11.0),
            ..Default::default()
        };
        assert_eq!(unsized_line_metrics(&para, 9.5, &HashMap::new()).0, 11.0);
        para.runs.push(Run {
            font_size: 10.0,
            is_line_break: true,
            ..Default::default()
        });
        assert_eq!(unsized_line_metrics(&para, 9.5, &HashMap::new()).0, 10.0);
    }

    #[test]
    fn footnote_free_paragraph_costs_nothing() {
        let (per_line, total) =
            per_line_footnote_extra(&[vec![], vec![]], &[], &HashSet::new(), 12.0, |_| 99.0);
        assert_eq!(per_line, vec![0.0, 0.0]);
        assert_eq!(total, 0.0);
    }

    fn font(lhr: f32, ar: f32) -> FontEntry {
        FontEntry {
            pdf_name: "F".to_string(),
            line_h_ratio: Some(lhr),
            ascender_ratio: Some(ar),
            plain_line_h_ratio: Some(lhr),
            plain_ascender_ratio: Some(ar),
            ..FontEntry::stub()
        }
    }

    /// case33: an 11pt Symbol bullet on 11pt Calibri gives Word a 16.0pt line
    /// (marker ascent + text descent, ×1.15), not the 15.5pt of either font alone.
    #[test]
    fn symbol_bullet_adds_its_extra_ascent_once() {
        let (cal_lhr, cal_ar) = (1.220703, 0.952148);
        let fonts = HashMap::from([
            ("Symbol".to_string(), font(1.225098, 1.005371)),
            ("Calibri".to_string(), font(cal_lhr, cal_ar)),
            ("Courier New".to_string(), font(1.132813, 0.832520)),
        ]);
        let text_line_h = 11.0 * cal_lhr * 1.15;
        let mut para = Paragraph {
            list_label: "\u{2022}".to_string(),
            ..Default::default()
        };
        let boosted = |para: &Paragraph| {
            label_boosted_line_h(
                para,
                &fonts,
                text_line_h,
                LineSpacing::Auto(1.15),
                11.0,
                Some(cal_lhr),
                Some(cal_ar),
            )
        };

        para.list_label_font = Some("Symbol".to_string());
        // 15.44 (Calibri at 1.15) + 0.59 unscaled extra ascent; Word 16.00.
        assert!(
            (boosted(&para) - 16.027).abs() < 0.01,
            "got {}",
            boosted(&para)
        );
        // Symbol reaches higher than Calibri, so the first baseline drops too.
        let off = label_boosted_baseline_offset(&para, &fonts, 11.0 * cal_ar, 11.0);
        assert!((off - 11.0 * 1.005371).abs() < 0.001);

        para.list_label_font = Some("Calibri".to_string());
        assert_eq!(boosted(&para), text_line_h);
        // Courier New's deeper descent does not count (streamnet p5 sub-bullets).
        para.list_label_font = Some("Courier New".to_string());
        assert_eq!(boosted(&para), text_line_h);
        para.list_label_font = Some("Symbol".to_string());
        para.text_alignment = crate::model::TextAlignment::Center;
        assert_eq!(boosted(&para), text_line_h);
        assert_eq!(
            label_boosted_baseline_offset(&para, &fonts, 11.0 * cal_ar, 11.0),
            11.0 * cal_ar
        );
        assert_eq!(
            label_boosted_line_h(
                &para,
                &fonts,
                8.0,
                LineSpacing::Exact(8.0),
                11.0,
                Some(cal_lhr),
                Some(cal_ar)
            ),
            8.0
        );
    }
}

/// Logical page number of each page: a section with `w:pgNumType @start` restarts,
/// others continue. A restart after an odd/even section break with the wrong
/// parity takes the next number instead (Word skips a number, not a sheet).
/// A page shows the number of the section at its top; a restart in a section
/// that begins further down counts that page as its start, so the next page
/// shows start + 1 (Word probes, 2026-10-03).
fn page_numbers(doc: &Document, page_section_indices: &[(usize, bool, usize)]) -> Vec<usize> {
    let mut numbers = Vec::with_capacity(page_section_indices.len());
    let mut running = 0;
    let mut prev_si = None;
    for &(si, _, last_si) in page_section_indices {
        let sp = &doc.sections[si].properties;
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
            .filter_map(|s| doc.sections[s].properties.page_num_start)
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
    page_numbers(doc, page_section_indices)
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

fn cell_grid_pitch(doc: &Document, sp: &crate::model::SectionProperties) -> f32 {
    sp.line_grid_pitch()
        .filter(|_| doc.adjust_line_height_in_table)
        .unwrap_or(0.0)
}

#[cfg(test)]
mod table_wrap_indent_tests {
    use super::*;

    #[test]
    fn right_table_strip_intersects_margin_relative_paragraph_box() {
        let mut para = Paragraph {
            indent_left: 255.0,
            indent_right: 12.0,
            ..Default::default()
        };
        assert_eq!(
            right_text_region(78.0, 552.0, 299.0, &para),
            (333.0, 207.0)
        );
        para.indent_left = 100.0;
        assert_eq!(
            right_text_region(78.0, 552.0, 299.0, &para),
            (299.0, 241.0)
        );
        para.indent_left = -10.0;
        assert_eq!(
            right_text_region(78.0, 552.0, 299.0, &para),
            (299.0, 241.0)
        );
        para.indent_left = 500.0;
        assert_eq!(
            right_text_region(78.0, 552.0, 299.0, &para),
            (578.0, 0.0)
        );
    }
}
