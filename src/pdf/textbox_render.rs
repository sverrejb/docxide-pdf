use std::collections::HashSet;

use pdf_writer::Content;

use crate::model::{
    EmbeddedImage, ImageGlow, ImageReflection, LineSpacing, Paragraph, Run, SectionProperties,
    ShapeFill, ShapeGeometry, Table, TextAlignment, TextAnchor, Textbox,
};

use super::color::{fill_rgb, set_fill_alpha, stroke_rgb};
use super::header_footer::resolve_tb_y_top;
use super::helpers::align_offset;
use super::layout::{
    LineOpts, LinkAnnotation, LinkTagger, TextLine, build_lines, line_max_image_h, lines_height,
    picture_line_bottom, render_paragraph_lines, tallest_run_metrics,
};
use super::list_label::render_list_label;
use super::positioning::resolve_h_position;
use super::smartart::draw_shape_path;
use super::table::TextboxTable;
use super::tagging::{Lists, Tags};
use super::wordart;
use super::{GradientSpec, RenderContext, render_shape_fill, resolve_line_h};

/// Returns the block image for an image-only textbox paragraph (no text).
/// Textbox paragraphs are parsed with `resolve_drawings: false`, so the
/// `inline_image_count == 1 && !has_text` promotion in `docx::paragraph`
/// doesn't fire — `tp.image` stays `None` and the image lives on a run.
/// This helper re-derives the block image at render time so callers can
/// size the paragraph as image-height (not text-line-height) and emit the
/// image directly instead of going through line layout.
pub(super) fn textbox_para_block_image(tp: &Paragraph) -> Option<&EmbeddedImage> {
    if let Some(img) = &tp.image {
        return Some(img);
    }
    let has_text = tp.runs.iter().any(|r| !r.text.is_empty() || r.is_tab);
    if has_text || tp.runs.iter().any(Run::underlined_picture) {
        return None;
    }
    tp.runs.iter().find_map(Run::shown_image)
}

/// `textbox_para_block_image` unless the paragraph's lines are exact: an exact
/// line keeps its height and its picture stands on the baseline above it, cut
/// at the box (Word example: a 100pt picture in an exactly 12pt line of a
/// shape-fitted box leaves the box 12pt tall, the picture shown inside it).
fn box_block_image<'p>(tp: &'p Paragraph, ctx: &RenderContext) -> Option<&'p EmbeddedImage> {
    textbox_para_block_image(tp).filter(|_| !exact_lines(tp, ctx))
}

fn exact_lines(tp: &Paragraph, ctx: &RenderContext) -> bool {
    matches!(
        tp.line_spacing.unwrap_or(ctx.doc_line_spacing),
        LineSpacing::Exact(_)
    )
}

fn exact_picture_line(tp: &Paragraph, ctx: &RenderContext) -> bool {
    exact_lines(tp, ctx) && tp.runs.iter().any(|r| r.shown_image().is_some())
}

/// A picture-only paragraph's height: the picture plus the leading its
/// mark's font gets from multiple line spacing (Word: a 50pt picture in 10pt
/// Calibri steps 50, 51.6 and 56.1 at 1, 1.13 and 1.5 lines).
fn image_block_height(img: &EmbeddedImage, tp: &Paragraph, ctx: &RenderContext) -> f32 {
    let h = img.display_height + img.layout_extra_height;
    match tp.line_spacing.unwrap_or(ctx.doc_line_spacing) {
        ls @ LineSpacing::Auto(_) => h + picture_line_bottom(&tp.runs, tp, ctx.fonts, ls),
        LineSpacing::AtLeast(min) => h.max(min),
        LineSpacing::Exact(_) => h,
    }
}

/// Rendered height of a textbox: the stored height, or for `a:spAutoFit` shapes the
/// height of their content (Word shrinks or grows the shape to fit). Shared by body
/// and table-cell textboxes so a cell-anchored autofit box does not keep Word's
/// 110.6pt default and paint over neighbouring shapes (japanese_land_development p1).
pub(super) fn textbox_height(tb: &Textbox, ctx: &RenderContext) -> f32 {
    if matches!(tb.auto_fit, crate::model::AutoFit::Shape) {
        let area_w = tb.text_area(0.0).0.max(0.0);
        let tmp_w = if tb.no_text_wrap { 10000.0 } else { area_w };
        let h = StoryPart::whole(tb).height(tmp_w, area_w, ctx);
        h + tb.margin_top + tb.margin_bottom
    } else {
        tb.height_pt
    }
}

/// The part of a story a box shows: its paragraphs and the tables among
/// them, each with the index of the shown paragraph it comes before (the
/// paragraph count for tables whose paragraph a later box shows).
pub(super) struct StoryPart<'t> {
    pub(super) paragraphs: &'t [Paragraph],
    tables: Vec<(usize, &'t Table)>,
}

#[derive(Clone, Copy)]
enum Block<'t> {
    Table(usize, &'t Table),
    Para(usize),
}

/// A story's tables and paragraphs in reading order, each table before the
/// paragraph whose index it carries.
fn story_blocks<'t>(
    paragraphs: &'t [Paragraph],
    tables: &'t [(usize, Table)],
) -> impl Iterator<Item = Block<'t>> {
    (0..paragraphs.len()).flat_map(move |i| {
        tables
            .iter()
            .filter(move |(at, _)| *at == i)
            .map(|(at, t)| Block::Table(*at, t))
            .chain(std::iter::once(Block::Para(i)))
    })
}

impl<'t> StoryPart<'t> {
    /// The box's own paragraphs and tables.
    pub(super) fn whole(tb: &'t Textbox) -> Self {
        Self::slice(&tb.paragraphs, &tb.tables, 0..usize::MAX)
    }

    /// The blocks in `range` of the story's `story_blocks`.
    fn slice(
        paragraphs: &'t [Paragraph],
        tables: &'t [(usize, Table)],
        range: std::ops::Range<usize>,
    ) -> Self {
        let mut first = None;
        let mut end = 0;
        let mut shown = Vec::new();
        for block in story_blocks(paragraphs, tables)
            .skip(range.start)
            .take(range.len())
        {
            let at = match block {
                Block::Table(at, t) => {
                    shown.push((at, t));
                    at
                }
                Block::Para(i) => {
                    end = i + 1;
                    i
                }
            };
            first.get_or_insert(at);
        }
        let start = first.unwrap_or(0);
        Self {
            tables: shown.into_iter().map(|(at, t)| (at - start, t)).collect(),
            paragraphs: &paragraphs[start..end.max(start)],
        }
    }

    pub(super) fn is_empty(&self) -> bool {
        self.paragraphs.is_empty() && self.tables.is_empty()
    }

    /// The story's height: its paragraphs with their spacing at text width
    /// `text_w`, and its tables in a text area `area_w` wide.
    fn height(&self, text_w: f32, area_w: f32, ctx: &RenderContext) -> f32 {
        let paras: f32 = self
            .paragraphs
            .iter()
            .map(|tp| tp.space_before + paragraph_body_height(tp, text_w, ctx) + tp.space_after)
            .sum();
        let tables: f32 = self
            .tables
            .iter()
            .map(|(_, t)| TextboxTable::layout(t, area_w, ctx, None).height())
            .sum();
        paras + tables
    }
}

/// How far a box's story starts below the top of its text area: Word centres
/// or bottom-aligns it (anchor ctr/b) in the room it leaves in `avail_h`.
pub(super) fn anchor_offset(
    tb: &Textbox,
    story: &StoryPart,
    text_w: f32,
    area_w: f32,
    avail_h: f32,
    ctx: &RenderContext,
) -> f32 {
    let gap = || (avail_h - story.height(text_w, area_w, ctx)).max(0.0);
    match tb.text_anchor {
        TextAnchor::Top => 0.0,
        TextAnchor::Middle => gap() / 2.0,
        TextAnchor::Bottom => gap(),
    }
}

/// A textbox paragraph's height without its spacing, at text width `text_w`.
fn paragraph_body_height(tp: &Paragraph, text_w: f32, ctx: &RenderContext) -> f32 {
    if let Some(img) = box_block_image(tp, ctx) {
        return image_block_height(img, tp, ctx);
    }
    let tw = (text_w - tp.indent_left - tp.indent_right).max(1.0);
    let hang = if !tp.list_label.is_empty() {
        match tp.num_level_tab_stop {
            Some(nts)
                if nts < tp.indent_left && (tp.indent_left - tp.indent_hanging).abs() < 0.5 =>
            {
                (tp.indent_left - nts).max(0.0)
            }
            _ if tp.indent_first_line > 0.0 && tp.indent_hanging == 0.0 => -tp.indent_first_line,
            _ => 0.0,
        }
    } else if tp.indent_hanging > 0.0 {
        tp.indent_hanging
    } else {
        -tp.indent_first_line
    };
    let laid = layout_box_paragraph(tp, &tp.runs, tw, hang, ctx);
    lines_height(&laid.lines, laid.line_h, laid.metrics)
}

/// A textbox paragraph's lines as they render, pictures included: a
/// shape-fitted box is as tall as they are (Word example: a box of only an
/// underlined 100pt picture in a 16pt Tahoma run is 103.8pt tall).
struct BoxParagraph {
    lines: Vec<TextLine>,
    line_h: f32,
    font_size: f32,
    first_ascent: f32,
    /// (ascent, descent) for picture lines, see `inline_line_advance`.
    metrics: (f32, f32),
}

fn layout_box_paragraph(
    tp: &Paragraph,
    runs: &[Run],
    text_w: f32,
    hanging: f32,
    ctx: &RenderContext,
) -> BoxParagraph {
    let ls = tp.line_spacing.unwrap_or(ctx.doc_line_spacing);
    let exact = matches!(ls, LineSpacing::Exact(_));
    let inline_imgs = super::layout::run_image_map(runs, ctx.textbox_image_names);
    let mut lines = build_lines(
        runs,
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
    let (fs, lhr, ar) = tallest_run_metrics(runs, ctx.fonts);
    // A paragraph of only tabs is sized by its break or mark, not by the
    // tab runs (`size_unsized_lines`).
    let (fs, (lhr, ar)) = if runs.iter().all(|r| r.is_tab || r.vanish) {
        super::unsized_line_font(tp, fs, ctx.fonts)
    } else {
        (fs, (lhr, ar))
    };
    let line_h = resolve_line_h(ls, fs, lhr);
    let text_ascent = fs * ar.unwrap_or(0.75);
    let first_ascent = super::layout::boxed_line_ascent(ls, line_h, fs, lhr, ar, runs, ctx.fonts)
        .unwrap_or(text_ascent);
    // Each line is as tall as its own runs, as in the body: air_pollution's
    // 10.5pt dot-leader line before a break steps 18.1 at 1.5 lines, not
    // the 20.7 of the paragraph's 12pt text.
    if !lines.is_empty() && !exact {
        super::layout::size_lines_by_own_runs(
            &mut lines,
            ctx.fonts,
            ls,
            line_h,
            text_ascent,
            1.0,
            tp.text_alignment == TextAlignment::Top,
        );
        super::size_break_lines(&mut lines, runs, ls, line_h, ctx.fonts);
        super::size_unsized_lines(
            &mut lines,
            tp,
            runs,
            ls,
            (line_h, fs, first_ascent),
            ctx.fonts,
        );
        super::size_mark_line(&mut lines, tp, runs, ls, lhr, ctx.fonts);
        super::size_tab_mark_line(&mut lines, tp, ls, lhr, ctx.fonts);
    }
    let descent = if inline_imgs.is_empty() {
        0.0
    } else {
        picture_line_bottom(runs, tp, ctx.fonts, ls)
    };
    let picture_ascent = if exact {
        lines
            .iter()
            .map(line_max_image_h)
            .fold(text_ascent, f32::max)
    } else {
        text_ascent
    };
    BoxParagraph {
        lines,
        line_h,
        font_size: fs,
        first_ascent,
        metrics: (picture_ascent, descent),
    }
}

/// The part of its story a linked box shows: after what the boxes before it
/// hold, and only what fits when a later box takes the rest (torture_test's
/// second paragraph opens the right-hand box). A box always keeps its first
/// table or paragraph, also one taller than the box, and what follows moves
/// on (text-box-table-linked-tall). ponytail: whole paragraphs and tables;
/// Word splits a story between lines and table rows.
pub(super) fn story_window<'t>(tb: &'t Textbox, height: f32, ctx: &RenderContext) -> StoryPart<'t> {
    let Some(link) = &tb.link else {
        return StoryPart::whole(tb);
    };
    // A chain outside the body was never linked: its first box keeps the text.
    let (paras, tables) = if link.story.is_empty() {
        (&tb.paragraphs[..], &tb.tables[..])
    } else {
        (&link.story[..], &link.tables[..])
    };
    let mut start = 0;
    for &area in &link.before {
        start += blocks_fitting(paras, tables, start, area, ctx);
    }
    let end = if link.continues {
        start + blocks_fitting(paras, tables, start, tb.text_area(height), ctx)
    } else {
        usize::MAX
    };
    StoryPart::slice(paras, tables, start..end)
}

/// How many of the story's blocks from `start` on a text area holds, at
/// least one when there are any.
fn blocks_fitting(
    paras: &[Paragraph],
    tables: &[(usize, Table)],
    start: usize,
    (w, h): (f32, f32),
    ctx: &RenderContext,
) -> usize {
    let mut used = 0.0;
    let mut n = 0;
    for block in story_blocks(paras, tables).skip(start) {
        let (height, after) = match block {
            Block::Table(_, t) => (TextboxTable::layout(t, w, ctx, None).height(), 0.0),
            Block::Para(i) => (
                paras[i].space_before + paragraph_body_height(&paras[i], w, ctx),
                paras[i].space_after,
            ),
        };
        if n > 0 && used + height > h + 0.01 {
            break;
        }
        used += height + after;
        n += 1;
    }
    n
}

/// `a:glow` on a shape: the shape grown by up to the glow's radius in faint
/// layers of its colour, so it fades outward from the edge to nothing.
/// ponytail: layered fills; Word draws a blurred raster.
fn draw_shape_glow(
    content: &mut Content,
    glow: &ImageGlow,
    (x, y, w, h): (f32, f32, f32, f32),
    shape: &ShapeGeometry,
    alpha_states: &mut HashSet<u8>,
) {
    const LAYERS: usize = 8;
    content.save_state();
    // Layers stack to the glow's own opacity at the shape's edge.
    let layer_alpha = 1.0 - (1.0 - glow.alpha).powf(1.0 / LAYERS as f32);
    set_fill_alpha(content, layer_alpha, alpha_states);
    fill_rgb(content, glow.color);
    for i in 0..LAYERS {
        let d = glow.radius * (LAYERS - i) as f32 / LAYERS as f32;
        draw_shape_path(content, x - d, y - d, w + 2.0 * d, h + 2.0 * d, shape);
        content.fill_nonzero();
    }
    content.restore_state();
}

/// `a:reflection` on a shape: its fill mirrored under it, fading from
/// `start_alpha` to `end_alpha` over `end_pos` of its height. Layers all start
/// at the top and reach one band further each, so their opacities stack
/// to the fade without seams between bands.
/// ponytail: the fill only (no outline or text).
fn draw_shape_reflection(
    content: &mut Content,
    refl: &ImageReflection,
    color: [u8; 3],
    (x, y, w, h): (f32, f32, f32, f32),
    shape: &ShapeGeometry,
    alpha_states: &mut HashSet<u8>,
) {
    const BANDS: usize = 8;
    let top = y - refl.distance;
    let band_h = h * refl.end_pos / BANDS as f32;
    let target = |k: usize| {
        let t = (k as f32 + 0.5) / BANDS as f32;
        refl.start_alpha + (refl.end_alpha - refl.start_alpha) * t
    };
    content.save_state();
    // The mirror image of a symmetric shape is the shape itself.
    draw_shape_path(content, x, top - h, w, h, shape);
    content.clip_nonzero();
    content.end_path();
    fill_rgb(content, color);
    for k in (0..BANDS).rev() {
        // Band k lies under layers k.., so layer k adds what band k + 1 lacks.
        let below = if k + 1 < BANDS { target(k + 1) } else { 0.0 };
        let alpha = 1.0 - (1.0 - target(k)) / (1.0 - below);
        set_fill_alpha(content, alpha, alpha_states);
        content.rect(x, top - (k + 1) as f32 * band_h, w, (k + 1) as f32 * band_h);
        content.fill_nonzero();
    }
    content.restore_state();
}

/// A text box's outline, dashed by its `a:prstDash`. Word strokes a
/// rectangle as `re`, from its bottom-left corner, which sets where a dash
/// pattern starts (shape-outline-dash).
pub(super) fn stroke_textbox_outline(
    content: &mut Content,
    tb: &Textbox,
    x: f32,
    y: f32,
    width: f32,
    height: f32,
) {
    let Some(stroke) = tb.stroke_color.filter(|_| tb.stroke_width > 0.0) else {
        return;
    };
    content.save_state();
    content.set_line_width(tb.stroke_width);
    if !tb.stroke_dash.is_empty() {
        content.set_dash_pattern(tb.stroke_dash.iter().map(|d| d * tb.stroke_width), 0.0);
    }
    stroke_rgb(content, stroke);
    if tb.shape_type.preset.as_deref() == Some("rect") && tb.shape_type.custom.is_none() {
        content.rect(x, y, width, height);
    } else {
        super::smartart::draw_shape_stroke_path(content, x, y, width, height, &tb.shape_type);
    }
    content.stroke();
    content.restore_state();
}

/// A text box's picture fill, stretched over `frame` (x, bottom, w, h) and
/// clipped to the shape.
pub(super) fn render_picture_fill(
    content: &mut Content,
    tb: &Textbox,
    (x, y, w, h): (f32, f32, f32, f32),
    ctx: &RenderContext,
) {
    if let Some(name) = tb
        .picture_fill
        .as_ref()
        .and_then(|img| ctx.textbox_image_names.get(&img.id))
    {
        super::smartart::render_image_with_clip(content, name, x, y, w, h, Some(&tb.shape_type));
    }
}

pub(super) fn render_single_textbox(
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
    // Tags its text, otherwise an artifact, in a Sect hoisted after the anchor
    // paragraph; `content` must be inside the default artifact.
    tags: &mut Tags,
    page: usize,
    alpha_states: &mut HashSet<u8>,
) {
    // A group's wrap placeholder: nothing to draw or tag.
    if tb.paragraphs.is_empty()
        && tb.link.is_none()
        && tb.fill.is_none()
        && tb.picture_fill.is_none()
        && tb.stroke_color.is_none()
    {
        return;
    }
    let sect = tags.hoist("Sect", tb.anchor_seq);
    let tb_x = resolve_h_position(
        tb.h_relative_from,
        &tb.h_position,
        tb.width_pt,
        sp,
        col_x,
        col_w,
        text_width,
    );
    let tb_height = textbox_height(tb, ctx);
    let story = story_window(tb, tb_height, ctx);
    let tb_y_top = resolve_tb_y_top(tb.v_relative_from, &tb.v_position, tb_height, sp, slot_top);
    log::debug!(
        "TEXTBOX x={tb_x:.1} top={tb_y_top:.1} w={:.1} h={tb_height:.1} rot={} paras={}",
        tb.width_pt,
        tb.rotation_deg,
        story.paragraphs.len()
    );
    let rotated = rotate_about_center(
        content,
        tb.rotation_deg,
        tb_x,
        tb_y_top,
        tb.width_pt,
        tb_height,
    );

    let frame = (tb_x, tb_y_top - tb_height, tb.width_pt, tb_height);
    if let Some(glow) = &tb.glow {
        draw_shape_glow(content, glow, frame, &tb.shape_type, alpha_states);
    }
    if let (Some(refl), Some(ShapeFill::Solid(color))) = (&tb.reflection, &tb.fill) {
        draw_shape_reflection(content, refl, *color, frame, &tb.shape_type, alpha_states);
    }

    if let Some(ref fill) = tb.fill {
        render_shape_fill(
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
    render_picture_fill(content, tb, frame, ctx);

    stroke_textbox_outline(
        content,
        tb,
        tb_x,
        tb_y_top - tb_height,
        tb.width_pt,
        tb_height,
    );

    // Lay text out within the shape's text rectangle (e.g. an arrow's body,
    // not its head). For plain rectangles this is the full box, so non-shape
    // textboxes are unaffected. Body-margin insets apply within the text rect.
    let (txt_l, txt_b, txt_w, txt_h) = super::smartart::shape_text_rect(
        &tb.shape_type,
        tb.width_pt,
        tb_height,
    )
    .unwrap_or((0.0, 0.0, tb.width_pt, tb_height));
    let text_top = tb_y_top - tb_height + txt_b + txt_h;

    let content_x = tb_x + txt_l + tb.margin_left;
    let natural_w = (txt_w - tb.margin_left - tb.margin_right).max(0.0);
    // no_text_wrap uses a huge width for line-breaking to prevent wrapping,
    // but alignment still uses the natural textbox width
    let content_w = if tb.no_text_wrap { 10000.0 } else { natural_w };
    let align_w = natural_w;

    // WordArt uses the full frame height, including the vertical insets.
    let needs_clip = !matches!(tb.auto_fit, crate::model::AutoFit::Shape)
        || story
            .paragraphs
            .iter()
            .any(|tp| exact_picture_line(tp, ctx));
    if needs_clip {
        content.save_state();
        content.rect(tb_x, tb_y_top - tb_height, tb.width_pt, tb_height);
        content.clip_nonzero();
        content.end_path();
    }

    // Text warp: render glyphs as warped paths instead of normal text
    if tb
        .text_warp
        .as_ref()
        .is_some_and(|w| w.preset != "textNoShape")
    {
        if wordart::render_warped_textbox(tb, content, ctx.fonts, tb_x, tb_y_top, align_w)
            || wordart::render_text_on_path(tb, content, ctx.fonts, tb_x, tb_y_top, align_w)
        {
            // The outlines stay artifacts; the text reads as one P (Word:
            // Sect > P with the text).
            let p = tags.add(sect, "P");
            tags.begin(content, page, p);
            let (x, top) = (tb_x + tb.margin_left, tb_y_top - tb.margin_top);
            wordart::invisible_text(tb, content, ctx.fonts, x, top);
            Tags::end(content);
            if needs_clip {
                content.restore_state();
            }
            if rotated {
                content.restore_state();
            }
            return;
        }
    }

    // Word clips a box's text to its height inside the top and bottom insets,
    // and a standalone box's pictures to the text area, insets on all sides
    // (a group draws them out to the box's left and right edges).
    let inner_bottom = tb_y_top - tb_height + txt_b + tb.margin_bottom;
    let inner_h = (text_top - tb.margin_top - inner_bottom).max(0.0);
    if needs_clip {
        content.rect(tb_x, inner_bottom, tb.width_pt, inner_h);
        content.clip_nonzero();
        content.end_path();
    }
    let picture_clip =
        (needs_clip && !tb.in_group).then_some((content_x, inner_bottom, natural_w, inner_h));

    let available = txt_h - tb.margin_top - tb.margin_bottom;
    let anchor_offset = anchor_offset(tb, &story, content_w, natural_w, available, ctx);

    // For fixed-size textboxes, stop rendering content below the textbox bounds
    let clip_bottom = needs_clip.then_some(tb_y_top - tb_height);

    // The glow, shadow and text passes walk the same paragraphs; a pass differs
    // only in its offset, colour override, labels, link sink and tags.
    let mut pass = |content: &mut Content,
                    (dx, dy): (f32, f32),
                    color: Option<[u8; 3]>,
                    labels: bool,
                    links: &mut Vec<LinkAnnotation>,
                    tag: Option<StoryTag>| {
        render_textbox_paragraphs(
            &story,
            content,
            content_x,
            content_w,
            align_w,
            text_top - tb.margin_top - anchor_offset,
            dx,
            dy,
            color,
            labels,
            links,
            ctx,
            clip_bottom,
            gradient_specs,
            tag,
            None,
            picture_clip,
        );
    };
    let mut discard_links: Vec<LinkAnnotation> = Vec::new();

    // Glow pass: render text as thick stroke in glow color behind everything
    if let Some(glow) = wordart::find_text_glow(tb) {
        content.save_state();
        stroke_rgb(content, glow.color);
        content.set_line_width(glow.radius_pt * 2.0);
        content.set_line_join(pdf_writer::types::LineJoinStyle::RoundJoin);
        content.set_text_rendering_mode(pdf_writer::types::TextRenderingMode::Stroke);
        pass(content, (0.0, 0.0), None, false, &mut discard_links, None);
        content.restore_state();
    }

    // Shadow pass: render all text offset by shadow distance in shadow color
    if let Some(shadow) = wordart::find_text_shadow(tb) {
        content.save_state();
        let [sr, sg, sb] = shadow.color;
        let shadow_color = [
            (sr as f32 * shadow.alpha) as u8,
            (sg as f32 * shadow.alpha) as u8,
            (sb as f32 * shadow.alpha) as u8,
        ];
        let offset = (shadow.offset_x, shadow.offset_y);
        pass(
            content,
            offset,
            Some(shadow_color),
            false,
            &mut discard_links,
            None,
        );
        content.restore_state();
    }

    pass(
        content,
        (0.0, 0.0),
        None,
        true,
        page_links,
        Some(StoryTag { tags, page, sect }),
    );

    if needs_clip {
        content.restore_state();
    }
    if rotated {
        content.restore_state();
    }
}

/// Turn the graphics state by a shape's rotation (OOXML: degrees clockwise)
/// about its box centre. Returns whether a state was pushed to restore.
pub(super) fn rotate_about_center(
    content: &mut Content,
    rotation_deg: f32,
    x: f32,
    y_top: f32,
    w: f32,
    h: f32,
) -> bool {
    if rotation_deg.rem_euclid(360.0).abs() < 0.01 {
        return false;
    }
    let (s, c) = (-rotation_deg).to_radians().sin_cos();
    let (cx, cy) = (x + w / 2.0, y_top - h / 2.0);
    content.save_state();
    content.transform([c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]);
    true
}

pub(super) struct StoryTag<'a> {
    pub(super) tags: &'a mut Tags,
    pub(super) page: usize,
    pub(super) sect: usize,
}

/// Shared paragraph iteration loop used by the glow, shadow, and normal text passes
/// within a textbox. Each pass differs only in positional offsets, forced color override,
/// and whether list labels are rendered.
pub(super) fn render_textbox_paragraphs(
    story: &StoryPart,
    content: &mut Content,
    content_x: f32,
    content_w: f32,
    align_w: f32,
    start_y: f32,
    x_offset: f32,
    y_offset: f32,
    force_color: Option<[u8; 3]>,
    render_labels: bool,
    links: &mut Vec<LinkAnnotation>,
    ctx: &RenderContext,
    clip_bottom: Option<f32>,
    gradient_specs: &mut Vec<GradientSpec>,
    // As in `render_single_textbox`; each paragraph becomes a P or H1–H6
    // (a list paragraph an LI) in the Sect.
    mut tag: Option<StoryTag>,
    hf_sub: Option<&super::table_layout::HfSubstitution>,
    // (x, y, w, h) a paragraph holding a picture is clipped to.
    picture_clip: Option<(f32, f32, f32, f32)>,
) {
    let mut cursor_y = start_y;
    let mut prev_space_after = 0.0f32;
    let mut lists = Lists::default();
    for tp_idx in 0..=story.paragraphs.len() {
        for (_, table) in story.tables.iter().filter(|(at, _)| *at == tp_idx) {
            cursor_y -= prev_space_after;
            prev_space_after = 0.0;
            if clip_bottom.is_some_and(|bottom| cursor_y < bottom) {
                return;
            }
            let layout = TextboxTable::layout(table, align_w, ctx, hf_sub);
            if render_labels {
                let mut y = cursor_y - y_offset;
                let tag = tag.as_mut().map(|t| StoryTag {
                    tags: &mut *t.tags,
                    page: t.page,
                    sect: t.sect,
                });
                layout.render(
                    (content_x + x_offset, align_w),
                    &mut y,
                    content,
                    ctx,
                    gradient_specs,
                    links,
                    tag,
                    hf_sub,
                );
                lists = Lists::default();
            }
            cursor_y -= layout.height();
        }
        let Some(tp) = story.paragraphs.get(tp_idx) else {
            break;
        };
        // Collapse adjacent spacing: use max(prev_after, current_before) like body text
        let inter_gap = if tp_idx == 0 {
            tp.space_before
        } else {
            prev_space_after.max(tp.space_before)
        };
        // Stop rendering when content overflows the textbox bounds
        if let Some(bottom) = clip_bottom
            && cursor_y - inter_gap < bottom
        {
            break;
        }
        // Word keeps empty and picture-only paragraphs as empty P elements.
        // (Lbl, text element) as in the body: see `para_tags`.
        let tag_nodes = tag.as_mut().map(|StoryTag { tags, sect, .. }| {
            let item = tp.list_item.filter(|_| tp.outline_level.is_none());
            // Word tags headings in textboxes too (never in table cells).
            let style_name = tp
                .style_id
                .as_ref()
                .and_then(|id| ctx.style_id_to_name.get(id));
            let kind = super::para_tag_kind(tp, style_name);
            tags.para_nodes(&mut lists, *sect, item, !tp.list_label.is_empty(), kind)
        });
        let para_tag = tag_nodes.map(|(_, text)| text);
        let tp_text_w = (content_w - tp.indent_left - tp.indent_right).max(1.0);
        let tp_align_w = (align_w - tp.indent_left - tp.indent_right).max(1.0);
        let text_hanging = if !tp.list_label.is_empty() {
            if let Some(nts) = tp.num_level_tab_stop {
                (tp.indent_left - nts).max(0.0)
            } else if tp.indent_first_line > 0.0 && tp.indent_hanging == 0.0 {
                -tp.indent_first_line
            } else {
                0.0
            }
        } else if tp.indent_hanging > 0.0 {
            tp.indent_hanging
        } else {
            -tp.indent_first_line
        };

        if let Some(img) = box_block_image(tp, ctx) {
            let block_h = image_block_height(img, tp, ctx);
            let foot_drop = match tp.line_spacing.unwrap_or(ctx.doc_line_spacing) {
                LineSpacing::AtLeast(_) => block_h - img.display_height - img.layout_extra_height,
                _ => 0.0,
            };
            if let Some(pdf_name) = ctx.textbox_image_names.get(&img.id) {
                let img_x = content_x
                    + tp.indent_left
                    + x_offset
                    + align_offset(tp.alignment, (tp_align_w - img.box_width()).max(0.0));
                let img_y = cursor_y - inter_gap - foot_drop - img.display_height - y_offset;
                // The paragraph's P stays empty and the picture follows it as
                // a Figure in the Sect, as in the body.
                let figure = tag.as_mut().filter(|_| !img.decorative);
                if let Some(StoryTag { tags, page, sect }) = figure {
                    let figure = tags.add_figure(*sect, img.alt.as_deref());
                    tags.begin(content, *page, figure);
                }
                let clipped = super::layout::push_clip(content, picture_clip);
                super::smartart::render_image_with_clip(
                    content,
                    pdf_name,
                    img_x,
                    img_y,
                    img.display_width,
                    img.display_height,
                    img.clip_geometry.as_ref(),
                );
                if clipped {
                    content.restore_state();
                }
                if tag.is_some() && !img.decorative {
                    Tags::end(content);
                }
            }
            cursor_y -= inter_gap + block_h;
            prev_space_after = tp.space_after;
            continue;
        }

        // Header/footer page numbers may sit in a text box: evaluate the
        // field runs for this page before laying them out. The substituted
        // vector lives for the rest of the iteration.
        let hf_runs: Vec<Run>;
        let runs: &[Run] = match hf_sub {
            Some(sub) => {
                hf_runs = super::header_footer::substitute_hf_runs(
                    &tp.runs,
                    sub.page_num,
                    sub.total_pages,
                    sub.section_pages,
                    sub.styleref_values,
                    sub.page_num_format,
                );
                &hf_runs
            }
            None => &tp.runs,
        };
        let BoxParagraph {
            lines: tb_lines,
            line_h: tb_line_h,
            font_size: tb_fs,
            first_ascent,
            metrics: tb_metrics,
        } = layout_box_paragraph(tp, runs, tp_text_w, text_hanging, ctx);
        if tb_lines.is_empty() {
            cursor_y -= inter_gap + tb_line_h;
            prev_space_after = tp.space_after;
            continue;
        }
        let tb_baseline = cursor_y - inter_gap - first_ascent - y_offset;
        let tp_text_x = content_x + tp.indent_left + x_offset;
        if let Some(c) = force_color {
            fill_rgb(content, c);
        }
        let draw_label = |content: &mut Content| {
            if render_labels {
                let label_x = content_x + tp.indent_left - tp.indent_hanging;
                render_list_label(content, tp, ctx.fonts, label_x, tb_baseline, tb_fs);
            }
        };
        match tag.as_mut().zip(tag_nodes) {
            Some((StoryTag { tags, page, .. }, nodes)) => {
                tags.begin_para(content, *page, nodes, draw_label)
            }
            None => draw_label(content),
        }
        render_paragraph_lines(
            content,
            &tb_lines,
            &tp.alignment,
            tp_text_x,
            tp.indent_left,
            tp_align_w,
            tb_baseline,
            tb_line_h,
            tb_metrics,
            tb_lines.len(),
            0,
            links,
            // The first line's indent also moves centred and right-aligned text
            // (air_pollution's caption: firstLine 709 shifts it 17.6pt).
            text_hanging,
            ctx.fonts,
            None,
            gradient_specs,
            None,
            None,
            tag.as_mut()
                .zip(para_tag)
                .map(|(StoryTag { tags, page, .. }, p)| LinkTagger::new(tags, *page, p)),
            picture_clip,
            None,
        );
        if para_tag.is_some() {
            Tags::end(content);
        }
        cursor_y -= inter_gap + lines_height(&tb_lines, tb_line_h, tb_metrics);
        prev_space_after = tp.space_after;
    }
}
