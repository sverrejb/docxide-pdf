//! Pre-flow pagination: float bands, lifted frames and where the next line can start.

use super::*;

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
pub(super) fn block_para(blocks: &[Block], i: usize) -> Option<&Paragraph> {
    blocks.get(i).and_then(leading_para)
}

pub(super) fn leading_para(block: &Block) -> Option<&Paragraph> {
    match block {
        Block::Paragraph(p) => Some(p),
        Block::Table(t) => first_cell_paragraph(t),
    }
}

/// The frame Word lifts `block` out of the flow into: any page- or
/// margin-anchored framePr (a table's is its first cell paragraph's), which
/// keeps text off its sides or, wrapping, lets text flow on both of them
/// (empty-frame-control), or a paragraph's wrapping frame at an offset
/// (`wrapping_text_frame`), which text then wraps around like a float.
/// ponytail: other text-anchored frames stay in the flow; lift them when a
/// fixture needs it.
pub(super) fn lifted_frame(block: &Block) -> Option<&FrameProperties> {
    let fp = leading_para(block)?.frame_props.as_ref()?;
    if matches!(block, Block::Table(_)) && fp.v_relative_from == VRelativeFrom::Paragraph
        || flow_frame(fp)
    {
        return None;
    }
    Some(fp)
}

/// A paragraph frame laid out in the flow: anchored to the text, neither
/// wrapping at an offset nor a drop cap.
pub(super) fn flow_frame(fp: &FrameProperties) -> bool {
    fp.v_relative_from == VRelativeFrom::Paragraph && !wrapping_text_frame(fp)
}

/// A frame wholly beside the text area, which no text can wrap around
/// (3ec631ca50's address block in the right margin).
pub(super) fn frame_beside_text(fp: &FrameProperties, sp: &SectionProperties) -> bool {
    let (text_x, text_w) = (sp.margin_left, sp.text_width());
    let x = resolve_h_position(
        fp.h_relative_from,
        &fp.h_position,
        fp.width,
        sp,
        text_x,
        text_w,
        text_w,
    );
    fp.width > 0.0 && (x >= text_x + text_w || x + fp.width <= text_x)
}

/// A wrapping paragraph-anchored frame at an offset, zero or not, with or
/// without w:h (fixtures frame-moves-*), or a drop cap: it keeps its own
/// coordinates, and text wraps around it.
fn wrapping_text_frame(fp: &FrameProperties) -> bool {
    fp.drop_cap.is_some()
        || (!fp.text_below
            && fp.v_relative_from == VRelativeFrom::Paragraph
            && matches!(fp.v_position, VerticalPosition::Offset(_)))
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

/// A band of the page that body lines keep out of, down from the page top.
#[derive(Clone, Copy, PartialEq)]
pub(super) struct Band {
    pub(super) top: f32,
    pub(super) bottom: f32,
    /// A tight or through picture with no room beside it: a line reaching
    /// into it steps down by whole line heights until clear, where other
    /// bands move it to their bottom (Word examples, a page-wide picture
    /// ending 270pt down: 13pt lines from 108.7 land at 277.7, 15pt lines
    /// from 116.7 at 281.7; with square wrapping both start at 270).
    pub(super) by_lines: bool,
    /// A picture's left and right edges (its wrap polygon's for tight and
    /// through, `polygon_edges`) with its text distances: the band
    /// holds only in the columns the picture reaches and leaves no usable
    /// strip in (`beside`) or reaches at all (topAndBottom). Word examples, two
    /// columns: a column-wide picture in column one leaves column two's
    /// lines in place, a page-wide one moves them too.
    pub(super) x: (f32, f32),
    pub(super) beside: bool,
}

impl Band {
    pub(super) fn new(top: f32, bottom: f32) -> Self {
        Band {
            top,
            bottom,
            by_lines: false,
            x: (f32::NEG_INFINITY, f32::INFINITY),
            beside: false,
        }
    }

    pub(super) fn in_column(&self, (col_x, col_w): (f32, f32)) -> bool {
        let left = self.x.0 - col_x;
        let right = col_x + col_w - self.x.1;
        left < col_w && right < col_w && (!self.beside || left.max(right) < MIN_EMPTY_STRIP)
    }
}

/// Where a line `need` tall starting at `top` clears the bands of column
/// `col`, stepping `step` at a time through those that move it by lines.
fn below_bands(mut top: f32, need: f32, step: f32, bands: &[Band], col: (f32, f32)) -> f32 {
    while let Some(b) = bands
        .iter()
        .find(|b| b.in_column(col) && top < b.bottom && top + need > b.top)
    {
        top = if b.by_lines && step > 0.0 {
            top + ((b.bottom - top) / step).ceil() * step
        } else {
            b.bottom
        };
    }
    top
}

/// A later line of a paragraph that reaches into a band steps below it on
/// its own, like the paragraph's first (`step_below_existing_bands`): Word
/// examples, a paragraph starting above a page-wide picture keeps its first
/// line there and moves the second below it (square) or by whole lines
/// (tight, through). Judged where the lines are placed: the bands are the
/// page's and column's they land on. Sets each line's `gap_after` and
/// returns their sum.
pub(super) fn step_lines_below_bands(
    lines: &mut [TextLine],
    bands: &[Band],
    col: (f32, f32),
    page_height: f32,
    first_top: f32,
    line_h: f32,
    metrics: (f32, f32),
) -> f32 {
    if bands.is_empty() {
        lines.iter_mut().for_each(|l| l.gap_after = 0.0);
        return 0.0;
    }
    let mut top = page_height - first_top;
    let mut gaps = 0.0;
    for i in 0..lines.len() {
        lines[i].gap_after = 0.0;
        let advance = layout::inline_line_advance(&lines[i], line_h, metrics);
        if i > 0 {
            let gap = below_bands(top, advance, advance, bands, col) - top;
            lines[i - 1].gap_after = gap;
            gaps += gap;
            top += gap;
        }
        top += advance;
    }
    gaps
}

/// Gives the current page to section `sect_idx` as its first page: its
/// first-page header and footer, their margins and bands.
pub(super) fn open_section_first_page(
    state: &mut LayoutState,
    sect_idx: usize,
    sp: &SectionProperties,
    ctx: &RenderContext,
) {
    state.pb.page_hf_section = sect_idx;
    state.pb.is_first_page_of_section = true;
    let page = state.pb.page_num(ctx.sections);
    state.pb.slot_top = effective_slot_top(sp, true, page, ctx);
    state.pb.column_top_y = state.pb.slot_top;
    state.pb.page_top_y = state.pb.slot_top;
    state.effective_margin_bottom =
        compute_effective_margin_bottom(sp, true, page, ctx) + state.pb.col_fn_reserved;
    state.pb.frame_bands = hf_page_bands(sp, true, page, state.effective_margin_bottom, ctx);
}

/// Whether `paragraph_bands` has any, without measuring a textbox.
pub(super) fn has_bands(p: &Paragraph, sp: &SectionProperties, col: (f32, f32), text_width: f32) -> bool {
    p.textboxes
        .iter()
        .any(|tb| textbox_blocks_band(tb, p, sp, col, text_width))
        || p.floating_images
            .iter()
            .any(|fi| image_blocks_band(fi, sp, col.0, col.1, text_width))
}

/// How far below the cursor a paragraph's first line reaches: the gap above
/// it, then the line, sized by its own runs and as tall as its tallest inline
/// drawing unless the spacing is exact, and before compat 15 a one-line
/// paragraph's space after (Word examples, fixtures page-table-line-reach and
/// its compat 14 twin: 24pt before, the space after above it, or an 80pt
/// inline box or picture carry a line into a floating table, 20pt after only
/// under compat 14; an 80pt box in a 13pt exact line keeps it out, and so does
/// an 80pt picture on the second line).
pub(super) fn first_line_reach(
    p: &Paragraph,
    (blocks, i): (&[Block], usize),
    col_w: f32,
    state: &LayoutState,
    sp: &SectionProperties,
    ctx: &RenderContext,
) -> f32 {
    let prev = i.checked_sub(1).and_then(|j| block_para(blocks, j));
    let before = helpers::effective_space_before(p, prev);
    let gap = state
        .pb
        .page_top_gap(sp, before, state.prev_space_after)
        .unwrap_or(inter_gap(state.prev_space_after, before, ctx));
    let mut lines = paragraph_lines(p, col_w, ctx, true);
    let after = if ctx.compat_mode < 15 && lines.len() <= 1 {
        helpers::effective_space_after(p, block_para(blocks, i + 1))
    } else {
        0.0
    };
    let ls = p.line_spacing.unwrap_or(ctx.doc_line_spacing);
    let (fs, lhr, ar) = tallest_run_metrics(&p.runs, ctx.fonts);
    let line_h = resolve_line_h(ls, fs, lhr);
    if matches!(ls, LineSpacing::Exact(_)) {
        return gap + line_h + after;
    }
    let picture_h = p
        .textboxes
        .iter()
        .filter(|tb| tb.on_line)
        .map(crate::model::Textbox::flow_bottom)
        .fold(
            if p.image.is_some() {
                p.content_height
            } else {
                lines.first().map_or(0.0, line_max_image_h)
            },
            f32::max,
        );
    let grid_snapped = p.snap_to_grid && sp.line_grid_pitch().is_some();
    let pitch = if !grid_snapped && let Some(first) = lines.first_mut() {
        let top_aligned = p.text_alignment == crate::model::TextAlignment::Top;
        let ascent = fs * ar.unwrap_or(0.75);
        size_lines_by_own_runs(
            std::slice::from_mut(first),
            ctx.fonts,
            ls,
            line_h,
            ascent,
            auto_ascent_scale(ls),
            top_aligned,
        );
        first.pitch.unwrap_or(line_h)
    } else {
        line_h
    };
    gap + line_h_with_picture(picture_h, &p.runs, p, ctx.fonts, ls, pitch) + after
}

/// The gap between a paragraph's space after and the next one's space
/// before: the larger under Word's HTML spacing rule, both when
/// doNotUseHTMLParagraphAutoSpacing adds them (5pt + 5pt list items step
/// 23.76 on a 13.8pt line).
pub(super) fn inter_gap(prev_after: f32, before: f32, ctx: &RenderContext) -> f32 {
    if ctx.additive_paragraph_spacing {
        prev_after + before
    } else {
        f32::max(prev_after, before)
    }
}

/// A page break still sits on its paragraph's first line, and that line
/// must fit first: one that does not moves to a page of its own before
/// breaking, leaving that page blank (Word probes p29a–c: a 25pt exact
/// break line under 23pt of room makes a blank page, a 15pt one does
/// not). The space above it may run past the page bottom (erasmus_plus:
/// a 17pt line with 22pt left under 12pt of space after stays put).
fn break_line_misses_page(para: &Paragraph, state: &LayoutState, ctx: &RenderContext) -> bool {
    if !para.page_break_before_explicit {
        return false;
    }
    let (mut font_size, mut lhr, _) = tallest_glyph_run_metrics(&para.runs, ctx.fonts);
    if lhr.is_none() {
        (font_size, lhr) = unsized_line_metrics(para, font_size, ctx.fonts);
    }
    let ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
    state.pb.slot_top - resolve_line_h(ls, font_size, lhr) < state.effective_margin_bottom
}

pub(super) fn take_page_break_before(
    state: &mut LayoutState,
    para: &Paragraph,
    sect_idx: usize,
    sp: &SectionProperties,
    ctx: &RenderContext,
) {
    let at_top = state.pb.is_at_page_top(sp);
    if !at_top && break_line_misses_page(para, state, ctx) {
        state.begin_next_page(sect_idx, sp, ctx);
    }
    if !at_top || para.page_break_before_explicit {
        state.begin_next_page(sect_idx, sp, ctx);
        state.current_col = 0;
    }
}

/// Whether the block before `i` is a floating table.
pub(super) fn follows_floating_table(blocks: &[Block], i: usize) -> bool {
    matches!(
        i.checked_sub(1).and_then(|j| blocks.get(j)),
        Some(Block::Table(t)) if t.position.is_some()
    )
}

/// The bands of a paragraph's blocking page- or margin-anchored floats.
fn paragraph_bands(
    p: &Paragraph,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    ctx: &RenderContext,
) -> Vec<Band> {
    let textboxes = p
        .textboxes
        .iter()
        .filter_map(|tb| blocking_textbox_band(tb, p, sp, col, text_width, ctx));
    let pictures = p
        .floating_images
        .iter()
        .filter(|fi| image_blocks_band(fi, sp, col.0, col.1, text_width))
        .map(|fi| {
            let (top, bottom) = image_band(fi, sp);
            let x = resolve_fi_x(fi, sp, col.0, col.1, text_width);
            let (l, r) = polygon_edges(fi);
            Band {
                by_lines: matches!(fi.wrap_type, WrapType::Tight | WrapType::Through),
                x: (x - fi.dist_left + l, x + r + fi.dist_right),
                beside: fi.wrap_type != WrapType::TopAndBottom,
                ..Band::new(top, bottom)
            }
        });
    textboxes.chain(pictures).collect()
}

/// The band of a floating table placed against the page or margin that leaves
/// no strip beside it in the column. Text laid out before the table on its
/// page keeps out of it too (Word examples, fixture float-table-band-before: a
/// paragraph above a page-top table moves below it, and on below the band of
/// a later anchor; with a strip beside the table it wraps there instead).
pub(super) fn floating_table_band(
    t: &crate::model::Table,
    sp: &SectionProperties,
    col: (f32, f32),
    ctx: &RenderContext,
) -> Option<Band> {
    let pos = page_anchored(t)?;
    let fp = FloatingTablePos::resolve(t, pos, sp, col.0, col.1, 0.0, ctx);
    let col_widths = table::resolved_col_widths(t, sp, ctx, None);
    let width: f32 = col_widths.iter().sum();
    let strip = side_strip(
        (fp.x, fp.x + width),
        (fp.left_from_text, fp.right_from_text),
        col,
    );
    if strip >= MIN_EMPTY_STRIP {
        return None;
    }
    let top = sp.page_height - fp.y;
    let height = table::hf_rows_height(t, &col_widths, ctx);
    Some(Band::new(
        top - fp.top_from_text,
        top + height + fp.bottom_from_text,
    ))
}

/// A floating table's position when it is placed against the page or margin,
/// not its text.
pub(super) fn page_anchored(t: &crate::model::Table) -> Option<&crate::model::TablePosition> {
    t.position.as_ref().filter(|pos| pos.v_anchor != "text")
}

/// A paragraph with no text and no drawings: it lays nothing out.
pub(super) fn lays_out_nothing(p: &Paragraph) -> bool {
    is_text_empty(&p.runs)
        && p.floating_images.is_empty()
        && p.textboxes.is_empty()
        && p.connectors.is_empty()
        && p.smartart.is_empty()
}

pub(super) fn add_bands(bands: &mut Vec<Band>, new: Vec<Band>) {
    for b in new {
        if !bands.contains(&b) {
            bands.push(b);
        }
    }
}

/// Before a body block: its paragraph's blocking floats join the page's
/// bands, and its first line steps below any band it would overlap.
#[allow(clippy::too_many_arguments)]
pub(super) fn step_below_bands(
    state: &mut LayoutState,
    block_idx: usize,
    block: &Block,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    retry_first: Option<f32>,
    ctx: &RenderContext,
) {
    let Block::Paragraph(p) = block else {
        // From the table on, its float zone keeps the text out.
        if let Block::Table(t) = block
            && !state.pb.frame_bands.is_empty()
            && let Some(own) = floating_table_band(t, sp, col, ctx)
        {
            state.pb.frame_bands.retain(|&b| b != own);
        }
        step_below_existing_bands(state, block, sp, col, retry_first, ctx);
        return;
    };
    add_bands(
        &mut state.pb.frame_bands,
        paragraph_bands(p, sp, col, text_width, ctx),
    );
    let retried = retry_first.is_some();
    let pending = state.pb.float_anchor_for(block_idx);
    let anchor_y = pending.map_or(
        state.pb.slot_top
            - if retried {
                0.0
            } else {
                space_after_laid_out(state)
            },
        |a| a.top,
    );
    let slot_before = state.pb.slot_top;
    step_below_existing_bands(state, block, sp, col, retry_first, ctx);
    let flowed = state.pb.slot_top - if retried { 0.0 } else { state.prev_space_after };
    // A line-relative box hangs from where the paragraph went (`ParaAnchor::drawing_top`).
    let line_y = pending
        .filter(|a| matches!(a.kind, AnchorKind::LookAhead))
        .map_or(flowed, |a| a.top);
    // After the anchor's own step: these bands hold off what follows, and the
    // anchor's own line only when a band pushed it down into its box (Word
    // examples: the box stays where the anchor stood, the line goes below it).
    let hanging = p
        .textboxes
        .iter()
        .filter(|tb| hangs_below_anchor(tb, p, col.1, ctx))
        .map(|tb| {
            let anchor_y = if tb.v_from_line { line_y } else { anchor_y };
            let top = sp.page_height - anchor_y + tb.v_offset_pt;
            Band::new(
                top,
                top + textbox_render::textbox_height(tb, ctx) + tb.dist_bottom,
            )
        })
        .collect();
    add_bands(&mut state.pb.frame_bands, hanging);
    if state.pb.slot_top != slot_before {
        step_below_existing_bands(state, block, sp, col, retry_first, ctx);
    }
}

/// Word places all of a page's page- or margin-anchored floats before it lays
/// out the page's lines, so text above an anchor keeps out of its band too
/// (issue #148: a letterhead band at the page top anchored in the third
/// paragraph moves the first two below it). On a page's first block, the
/// bands of the paragraphs estimated to start on the page, and of the
/// floating tables between them, join its bands. A paragraph that can't
/// start on the page moves on with its floats (Word example: a keepLines
/// paragraph at a page foot holding a page-wide picture at the page top
/// leaves the page's lines in place).
/// ponytail: plain line counts estimate where the page ends, stopping at the
/// first table in the flow (inline or text-anchored) or page break.
pub(super) fn prescan_page_bands(
    state: &mut LayoutState,
    blocks: &[Block],
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    ctx: &RenderContext,
) {
    let bottom = sp.page_height - state.effective_margin_bottom;
    let top = sp.page_height - state.pb.slot_top;
    let mut y = top;
    let mut bands = std::mem::take(&mut state.pb.frame_bands);
    // (space before, line height, lines, space after) of the paragraphs so far.
    let mut paras: Vec<(f32, f32, f32, f32)> = Vec::new();
    let place = |y: f32, (before, line_h, lines, after): (f32, f32, f32, f32), bands: &[Band]| {
        below_bands(y + before, line_h, line_h, bands, col) + lines * line_h + after
    };
    for (i, block) in blocks.iter().enumerate() {
        let p = match block {
            Block::Paragraph(p) => p,
            Block::Table(t) if page_anchored(t).is_some() => {
                // Text before the table that would no longer fit on the page
                // below it keeps its place: Word moves the table on, and what
                // follows it with it (examples: a 690pt table at the margin top,
                // or one of twenty 40pt rows, goes to the next page, the
                // paragraph above it stays out of a later anchor's band).
                if let Some(band) = floating_table_band(t, sp, col, ctx) {
                    let mut with = bands.clone();
                    add_bands(&mut with, vec![band]);
                    let end = paras.iter().fold(top, |y, &para| place(y, para, &with));
                    if end > bottom {
                        break;
                    }
                    bands = with;
                }
                continue;
            }
            Block::Table(_) => break,
        };
        if i > 0 && p.page_break_before {
            break;
        }
        let line_h = para_line_h(p, ctx);
        let lines = paragraph_line_count(p, col.1, ctx).max(1);
        let kept = if p.keep_lines {
            lines
        } else {
            lines_kept_together(p.widow_control, || lines)
        };
        let line_top = y + p.space_before;
        if i > 0
            && below_bands(line_top, line_h, line_h, &bands, col) + kept as f32 * line_h > bottom
        {
            break;
        }
        add_bands(&mut bands, paragraph_bands(p, sp, col, text_width, ctx));
        let para = (p.space_before, line_h, lines as f32, p.space_after);
        paras.push(para);
        y = place(y, para, &bands);
        if y > bottom || p.page_break_after || p.page_break_at.is_some() {
            break;
        }
    }
    state.pb.frame_bands = bands;
}

/// The width of a paragraph's first word, None for one without text.
pub(super) fn first_word_width(p: &Paragraph, ctx: &RenderContext) -> Option<f32> {
    if is_text_empty(&p.runs) {
        return None;
    }
    layout::opening_word_width(&p.runs, ctx.fonts)
}

/// The bands of the header and footer floating tables of a page opening with
/// section `sp`'s header that body lines keep out of (`hf_table_bands`); one
/// reaching the bottom margin in every column ends the page instead
/// (`compute_effective_margin_bottom`).
pub(super) fn hf_page_bands(
    sp: &SectionProperties,
    is_first: bool,
    page: usize,
    effective_margin_bottom: f32,
    ctx: &RenderContext,
) -> Vec<Band> {
    header_footer::hf_table_bands(sp, is_first, page, ctx)
        .into_iter()
        .filter(|band| band.top < sp.page_height - effective_margin_bottom)
        .collect()
}

/// The first time a page is laid out on, its upcoming float bands join
/// (`prescan_page_bands`), also when a paragraph runs or moves onto it.
pub(super) fn prescan_new_page(
    state: &mut LayoutState,
    blocks: &[Block],
    sect_idx: usize,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    ctx: &RenderContext,
) {
    let page = (sect_idx, state.pb.page_count());
    if state.band_prescan && state.pb.bands_prescanned != Some(page) {
        state.pb.bands_prescanned = Some(page);
        prescan_page_bands(state, blocks, sp, col, text_width, ctx);
    }
}

/// Moves the cursor down until a paragraph's first line, `first` tall
/// without the room planned below it where it was (example: a paragraph moved
/// on from above a second picture still starts 12 lines down), clears the
/// bands of the current column of `cols`, by whole lines through
/// a through picture (Word examples: a paragraph moved or running on into a
/// column or page whose top a page-wide through picture covers starts by
/// whole lines below the column top), first prescanning a new page's bands
/// from `rest`. A page whose bands leave the line no room in any column,
/// though it would fit without them, takes it at its foot in its first
/// column, returning true, and the next line on the next page (Word example:
/// header tables over both columns from the top margin past the bottom one
/// put one line a page, its top on the bottom margin of column one).
#[allow(clippy::too_many_arguments)]
pub(super) fn place_first_line(
    state: &mut LayoutState,
    first: f32,
    rest: &[Block],
    sect_idx: usize,
    sp: &SectionProperties,
    cols: &[(f32, f32)],
    text_width: f32,
    ctx: &RenderContext,
) -> bool {
    loop {
        let col = cols[state.current_col];
        prescan_new_page(state, rest, sect_idx, sp, col, text_width, ctx);
        let top = sp.page_height - state.pb.slot_top;
        let foot = sp.page_height - state.effective_margin_bottom;
        let below = |col| below_bands(top, first, first, &state.pb.frame_bands, col);
        let blocked = state.pb.takes_a_line(sp)
            && top + first <= foot
            && cols.iter().all(|&col| below(col) + first > foot);
        if !blocked {
            state.pb.slot_top -= below(col) - top;
            return false;
        }
        if state.current_col == 0 {
            state.pb.slot_top = state.effective_margin_bottom;
            return true;
        }
        state.advance_column_or_page(cols.len(), sect_idx, sp, ctx);
    }
}

#[allow(clippy::too_many_arguments)]
pub(super) fn place_moved_first_line(
    state: &mut LayoutState,
    first: f32,
    gap: f32,
    rest: &[Block],
    sect_idx: usize,
    sp: &SectionProperties,
    cols: &[(f32, f32)],
    text_width: f32,
    ctx: &RenderContext,
) {
    let stood = (state.current_col, state.pb.page_count(), state.pb.slot_top);
    place_first_line(state, first, rest, sect_idx, sp, cols, text_width, ctx);
    if (state.current_col, state.pb.page_count()) == (stood.0, stood.1)
        && state.pb.slot_top != stood.2
    {
        state.pb.pin_float_anchor(stood.2 + gap);
    }
}

/// Each text column's left edge and width.
pub(super) fn column_geometry(sp: &SectionProperties) -> Vec<(f32, f32)> {
    match sp.columns.as_ref() {
        Some(cfg) => {
            let mut x = sp.margin_left;
            cfg.columns
                .iter()
                .map(|col| {
                    let result = (x, col.width);
                    x += col.width + col.space;
                    result
                })
                .collect()
        }
        None => vec![(sp.margin_left, sp.text_width())],
    }
}

/// How many lines a paragraph takes in a column `col_w` wide.
pub(super) fn paragraph_line_count(p: &Paragraph, col_w: f32, ctx: &RenderContext) -> usize {
    paragraph_lines(p, col_w, ctx, false).len()
}

/// Names only mark the pictures, so their lines measure as tall.
fn picture_marks(p: &Paragraph) -> HashMap<usize, String> {
    p.runs
        .iter()
        .enumerate()
        .filter(|(_, r)| r.inline_image.is_some())
        .map(|(i, _)| (i, String::new()))
        .collect()
}

/// A paragraph's lines in a column `col_w` wide, with its inline pictures
/// when `with_pictures`.
fn paragraph_lines(p: &Paragraph, col_w: f32, ctx: &RenderContext, with_pictures: bool) -> Vec<TextLine> {
    let pictures = with_pictures.then(|| picture_marks(p));
    build_lines(
        &p.runs,
        ctx,
        (col_w - p.indent_left - p.indent_right).max(1.0),
        ctx.cjk(true, p),
        &LineOpts {
            inline_images: pictures.as_ref(),
            tab_stops: &p.tab_stops,
            indent_left: p.indent_left,
            indent_right: p.indent_right,
            hanging: p.indent_hanging - p.indent_first_line,
            ..Default::default()
        },
    )
}

/// The previous paragraph's space after above the next one: none once the
/// next one's page break is taken (Word examples, fixture float-table-band-before:
/// 24pt after and a break keep a 13pt line above a table 18pt down, and a box
/// hanging 40pt below the paragraph after the break holds the text off where
/// it is drawn).
fn space_after_laid_out(state: &LayoutState) -> f32 {
    if state.page_break_taken {
        0.0
    } else {
        state.prev_space_after
    }
}

pub(super) fn step_below_existing_bands(
    state: &mut LayoutState,
    block: &Block,
    sp: &SectionProperties,
    col: (f32, f32),
    retry_first: Option<f32>,
    ctx: &RenderContext,
) {
    let top = sp.page_height - state.pb.slot_top;
    let Some(p) = leading_para(block) else { return };
    if state.pb.frame_bands.iter().all(|b| top >= b.bottom) {
        return;
    }
    // The first line fits above a band with the space before it, and a
    // one-line paragraph only with its space after too (Word probes, 12.65pt
    // line + 6pt after: an 18.0pt gap between two topAndBottom boxes is
    // skipped, 19.0pt holds it; 14.2pt holds it with no space after).
    let line_h = para_line_h(p, ctx);
    let laid_out_after = if retry_first.is_some() {
        0.0
    } else {
        space_after_laid_out(state)
    };
    let (need, line_top, dropped) = if let Some(first) = retry_first {
        (first, top, 0.0)
    } else {
        let one_line = paragraph_line_count(p, col.1, ctx) <= 1;
        // At a page top a paragraph lays out without its space before (Word
        // examples, fixture float-table-band-before: 28pt before a break's line does
        // not hold it above a page-top table).
        let laid_out_gap = laid_out_after.max(p.space_before);
        let gap = state
            .pb
            .page_top_gap(sp, p.space_before, laid_out_after)
            .unwrap_or(laid_out_gap);
        let own = own_space_before(p.space_before, state.prev_space_after);
        (
            line_h + if one_line { p.space_after } else { 0.0 },
            top + gap,
            laid_out_gap - own,
        )
    };
    let below = below_bands(line_top, need, line_h, &state.pb.frame_bands, col);
    if below != line_top {
        // Its floating paragraph-relative drawings still hang from where the
        // paragraph stood (example: a wrapNone box at offset 0 in a paragraph
        // pushed below a band stays above the band; vaccines' circles under
        // their banner); inline ones go with its line (Word examples: an inline
        // text box above a page band or a page-top table).
        state
            .pb
            .pin_float_anchor(state.pb.slot_top - laid_out_after);
        // Below a band the previous paragraph's space after is gone and this
        // one's own space before stays (probe, 6pt after and 12pt before: the
        // line starts 6pt below the box), even across a page break (Word
        // examples, fixture float-table-band-before: 24pt after, a break and
        // 12pt before put the line on the table's bottom, 10pt after 2pt below).
        state.pb.slot_top = sp.page_height - (below - dropped);
    }
}

/// A lifted frame being laid out: its blocks render in the frame's own column
/// from its top, and the flow then resumes where it left off, clear of the
/// frame's band (Word probes: body lines step below it with no gap).
pub(super) struct OpenFrame<'a> {
    pub(super) props: &'a FrameProperties,
    top: f32,
    /// The frame's column: its left edge and width.
    pub(super) geometry: (f32, f32),
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
    pub(super) fn open(
        props: &'a FrameProperties,
        blocks: &[Block],
        state: &mut LayoutState,
        ctx: &RenderContext,
        sp: &SectionProperties,
        (sect_idx, col_geometry): (usize, &[(f32, f32)]),
    ) -> Self {
        // Its break comes first: the frame opens on the new page.
        take_frame_break(blocks, state, sp, ctx, (sect_idx, col_geometry.len()));
        let width = frame_width(props, blocks, ctx, sp.text_width());
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
        let top_y = |slot_top| {
            header_footer::resolve_tb_y_top(props.v_relative_from, &props.v_position, height, sp, slot_top)
        };
        // A frame never splits: one anchored to the text that runs past the
        // bottom moves to the next column or page with its anchor (Word example,
        // fixture frame-moves-exact-height: a 100pt frame with 54pt left opens page 2).
        if props.v_relative_from == VRelativeFrom::Paragraph && props.drop_cap.is_none() {
            let box_h = frame_box_h(props, blocks, ctx, width);
            advance_until_fits(state, sp, ctx, (sect_idx, col_geometry.len()), |state| {
                top_y(state.pb.slot_top) - box_h >= state.effective_margin_bottom
            });
        }
        let (col_x, col_w) = col_geometry[state.current_col];
        let x = match props.drop_cap {
            Some(dc) if dc.in_margin => col_x - props.h_space - width,
            Some(_) => col_x,
            None => resolve_h_position(
                props.h_relative_from,
                &props.h_position,
                width,
                sp,
                col_x,
                col_w,
                sp.text_width(),
            ),
        };
        let top = sp.page_height - top_y(state.pb.slot_top);
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

    pub(super) fn close(self, state: &mut LayoutState, sp: &SectionProperties) {
        let bottom = sp.page_height - state.pb.slot_top;
        // A full-width wrapping frame blocks body lines even if a previous
        // table already owns float_zone. Its overflowing content must not be
        // lost when restoring that table's exclusion zone below.
        let full_width = side_room(
            WrapText::BothSides,
            self.geometry.0 - sp.margin_left,
            sp.page_width - sp.margin_right - self.geometry.0 - self.geometry.1,
        ) < MIN_EMPTY_STRIP;
        if self.props.text_below || full_width {
            let wrap_bottom = if full_width && self.props.height_exact && self.wrap_height > 0.0 {
                self.top + self.wrap_height
            } else if full_width {
                self.top + self.wrap_height.max(bottom - self.top)
            } else {
                bottom
            };
            state.pb.frame_bands.push(Band::new(self.top, wrap_bottom));
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
            });
        }
    }
}

/// A frame in the flow, over the paragraphs that share it. Its w:h holds
/// the text after it off (Word examples, fixture frame-not-beside-height:
/// the next line starts 100pt below a 100pt exact frame's top, not under
/// its 72pt of lines; frame-exact-box-overflows: a 24pt exact box with 72pt
/// of lines stays in 44pt). It never splits: a box (w:h if exact, else at
/// least its content) that runs past the bottom moves to the next column or
/// page whole, and one taller than the page stays where it is
/// (frame-moves-shared-tall). Moved, its space before is
/// suppressed and still counts toward w:h (frame-moved-space-before: a
/// 100pt exact frame with 20pt before holds the next line 80pt below the
/// new page's top).
/// An exact box clips what overflows it, unless the frame holds a table
/// (Word example: a 12pt exact frame on a first cell keeps the table whole).
pub(super) struct FlowFrame<'a> {
    pub(super) props: &'a FrameProperties,
    top: f32,
    page: usize,
    /// An exact w:h around paragraphs only: it clips them and sets where the
    /// next text starts.
    exact: bool,
}

impl<'a> FlowFrame<'a> {
    pub(super) fn open(
        props: &'a FrameProperties,
        blocks: &[Block],
        state: &mut LayoutState,
        ctx: &RenderContext,
        sp: &SectionProperties,
        (sect_idx, col_geometry): (usize, &[(f32, f32)]),
    ) -> Self {
        take_frame_break(blocks, state, sp, ctx, (sect_idx, col_geometry.len()));
        let col_width = |state: &LayoutState| frame_width(props, blocks, ctx, col_geometry[state.current_col].1);
        let box_h = frame_box_h(props, blocks, ctx, col_width(state));
        let (start_page, col) = (state.pb.page_count(), state.current_col);
        let left = state.pb.column_top_y - state.pb.slot_top;
        let fits = |state: &LayoutState| state.pb.slot_top - box_h >= state.effective_margin_bottom;
        advance_until_fits(state, sp, ctx, (sect_idx, col_geometry.len()), fits);
        let moved = (state.pb.page_count(), state.current_col) != (start_page, col);
        let space_before = leading_para(&blocks[0]).map_or(0.0, |p| p.space_before);
        let top = state.pb.slot_top + if moved { space_before } else { 0.0 };
        let moved_column = moved && state.pb.page_count() == start_page;
        state.frame_overflow |= !fits(state)
            || (moved_column && frame_content_height(props, blocks, ctx, col_width(state)) > left);
        let exact = props.height_exact
            && props.height > 0.0
            && !frame_blocks(props, blocks).any(|b| matches!(b, Block::Table(_)));
        state.pb.frame_clip = exact
            && layout::push_clip(
                &mut state.pb.content,
                Some((0.0, top - props.height, sp.page_width, props.height)),
            );
        FlowFrame { props, top, page: state.pb.page_count(), exact }
    }

    pub(super) fn close(self, state: &mut LayoutState) {
        state.pb.end_frame_clip();
        // ponytail: a frame that changed pages (a table splitting) holds no
        // room on the new one.
        if self.page != state.pb.page_count() {
            return;
        }
        let bottom = self.top - self.props.height;
        state.pb.slot_top = if self.exact {
            bottom
        } else {
            state.pb.slot_top.min(bottom)
        };
    }
}

/// The break opening a frame's first paragraph, taken before the frame is
/// placed: a page break before, or a column break (a page break inside a
/// frame acts as one).
fn take_frame_break(
    blocks: &[Block],
    state: &mut LayoutState,
    sp: &SectionProperties,
    ctx: &RenderContext,
    (sect_idx, col_count): (usize, usize),
) {
    let Some(Block::Paragraph(p)) = blocks.first() else {
        return;
    };
    if p.page_break_before && !lays_out_nothing(p) {
        take_page_break_before(state, p, sect_idx, sp, ctx);
        state.page_break_taken = true;
    } else if p.column_break_before {
        state.pb.advance_column_or_page(
            &mut state.current_col,
            col_count,
            sect_idx,
            sp,
            &mut state.effective_margin_bottom,
            ctx,
        );
    }
}

/// Advances column by column until `fits` or a page top: a frame never
/// splits, and a column 2 that starts as high as column 1 may not hold it
/// either (Word example, fixture frame-continuous-columns: the frame opens
/// the next page).
fn advance_until_fits(
    state: &mut LayoutState,
    sp: &SectionProperties,
    ctx: &RenderContext,
    (sect_idx, col_count): (usize, usize),
    fits: impl Fn(&LayoutState) -> bool,
) {
    while !fits(state) && !state.pb.is_at_page_top(sp) {
        state.pb.advance_column_or_page(
            &mut state.current_col,
            col_count,
            sect_idx,
            sp,
            &mut state.effective_margin_bottom,
            ctx,
        );
    }
}

/// A frame without w:w is as wide as its widest block: a paragraph's picture
/// or longest line, a table's grid.
pub(super) fn auto_frame_width(
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
                |img| img.box_width(),
            ),
            Block::Table(t) => t.col_widths.iter().sum(),
        })
        .fold(0.0, f32::max)
}

/// A frame without w:h is as tall as its blocks laid out at its width.
pub(super) fn frame_content_height(
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
            Block::Paragraph(p) if p.paragraph_mark_vanish && lays_out_nothing(p) => {}
            Block::Paragraph(p) => {
                height += header_footer::hf_paragraph_gap(prev, prev_space_after, p);
                let pictures = picture_marks(p);
                let opts = LineOpts {
                    inline_images: Some(&pictures),
                    tab_stops: &p.tab_stops,
                    ..Default::default()
                };
                let line_w = (width - p.indent_left - p.indent_right).max(1.0);
                let lines = build_lines(&p.runs, ctx, line_w, ctx.cjk(true, p), &opts);
                let (fs, lhr, _) = tallest_run_metrics(&p.runs, ctx.fonts);
                let ls = p.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let line_h = resolve_line_h(ls, fs, lhr);
                // A picture makes its line as tall as itself, unless the
                // spacing is exact.
                let exact = matches!(ls, LineSpacing::Exact(_));
                let lines_h = lines
                    .iter()
                    .map(|l| if exact { line_h } else { line_h.max(layout::line_max_image_h(l)) })
                    .sum::<f32>()
                    .max(line_h);
                height += p.image.as_ref().map_or(lines_h, |img| lines_h.max(img.layout_size().1));
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

/// Modern Word keeps a margin-anchored empty frame's box even though its
/// paragraphs consume no body flow. A full-width inline table cannot wrap
/// beside that box, nor can body text above the anchor (Word probes:
/// empty-margin-frame-visible-body-before-frame). Only boxes whose anchor
/// fits this page are collected; frames after forced breaks are left alone.
pub(super) fn empty_frame_table_bands(
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
                    if t.all_paragraphs().any(|q| {
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
                    let widths = table_layout::auto_fit_columns(
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
pub(super) fn is_empty_wrapping_frame(para: &Paragraph) -> bool {
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

#[cfg(test)]
mod tests {
    use super::*;

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
}
