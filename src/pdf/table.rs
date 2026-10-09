use std::collections::HashMap;

use pdf_writer::{Content, Name, Str};

use crate::fonts::FontEntry;
use crate::model::{
    Alignment, Block, BorderStyle, CellBorder, CellMargins, CellVAlign, Paragraph,
    SectionProperties, Table, TableAlignment, TableRow, TextDirection, VMerge,
};

use super::color::{fill_rgb, stroke_rgb};
use super::header_footer::effective_slot_top;

use super::RenderContext;
use super::layout::{LinkAnnotation, LinkTagger, encode_text_for_pdf, render_paragraph_lines};
use super::table_layout::{
    CELL_START, CellContentItem, CellCursor, CellFloatingImageLayout, CellLayout,
    CellParagraphLayout, Chunk, HfSubstitution, RowLayout, RowPiece, apply_pct_width,
    auto_fit_columns, cell_span_width, cell_x_offset, chunk_space_before, compute_merge_spans,
    compute_row_layouts, cursor_chunks, find_cell_split, item_chunk_height, min_content_widths,
    para_block_height, partial_row_height, row_pieces, squeeze_to_width,
};
use super::tagging::{CellTagger, TableTags, Tags};

fn draw_border(content: &mut Content, border: &CellBorder, x1: f32, y1: f32, x2: f32, y2: f32) {
    if !border.present {
        return;
    }
    let w = border.width;
    content.save_state();
    content.set_line_width(w);
    if let Some(c) = border.color {
        stroke_rgb(content, c);
    }
    match border.style {
        BorderStyle::Dotted => {
            content.set_line_cap(pdf_writer::types::LineCapStyle::RoundCap);
            content.set_dash_pattern([0.0, w * 3.0], 0.0);
        }
        BorderStyle::Dashed | BorderStyle::DashSmallGap => {
            let dash = if border.style == BorderStyle::DashSmallGap {
                w * 3.0
            } else {
                w * 4.0
            };
            content.set_dash_pattern([dash, w * 2.0], 0.0);
        }
        BorderStyle::DashDot => {
            content.set_dash_pattern([w * 4.0, w * 2.0, 0.0, w * 2.0], 0.0);
        }
        BorderStyle::DashDotDot => {
            content.set_dash_pattern([w * 4.0, w * 2.0, 0.0, w * 2.0, 0.0, w * 2.0], 0.0);
        }
        BorderStyle::Double => {
            // Word renders each line of a double border at the full specified
            // width, separated by a clear gap of the same width. The gap here is
            // the center-to-center distance, so it must be 2×width: each line's
            // stroke spans ±width/2, and 2×width centers leave a width-wide gap
            // between their edges. (gap == width made the edges touch, so
            // antialiasing merged the pair into one thick line.)
            let thin = w.max(0.25);
            let gap = thin * 2.0;
            content.set_line_width(thin);
            content.move_to(x1, y1);
            content.line_to(x2, y2);
            content.stroke();
            let dx = if (x1 - x2).abs() < 0.01 { gap } else { 0.0 };
            let dy = if (y1 - y2).abs() < 0.01 { gap } else { 0.0 };
            content.move_to(x1 - dx, y1 - dy);
            content.line_to(x2 - dx, y2 - dy);
            content.stroke();
            content.restore_state();
            return;
        }
        BorderStyle::Single => {}
    }
    content.move_to(x1, y1);
    content.line_to(x2, y2);
    content.stroke();
    content.restore_state();
}

fn draw_cell_borders(
    content: &mut Content,
    borders: &crate::model::CellBorders,
    bx: f32,
    top: f32,
    bottom: f32,
    col_w: f32,
    draw_top: bool,
    draw_bottom: bool,
) {
    if draw_top {
        draw_border(content, &borders.top, bx, top, bx + col_w, top);
    }
    if draw_bottom {
        draw_border(content, &borders.bottom, bx, bottom, bx + col_w, bottom);
    }
    draw_border(content, &borders.left, bx, top, bx, bottom);
    draw_border(content, &borders.right, bx + col_w, top, bx + col_w, bottom);
}

/// Paint a cell background: hatching when a line/cross pattern is present,
/// otherwise a solid fill. No-op when both are absent.
fn paint_cell_background(
    content: &mut Content,
    shading: Option<[u8; 3]>,
    hatch: Option<crate::model::HatchPattern>,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
) {
    if let Some(hp) = hatch {
        draw_hatch(content, &hp, x, y, w, h);
    } else if let Some(s) = shading {
        content.save_state();
        fill_rgb(content, s);
        content.rect(x, y, w, h);
        content.fill_nonzero();
        content.restore_state();
    }
}

/// Draw a `w:shd` line/cross pattern as real hatching: a background fill plus
/// thin stroked lines clipped to the cell rect.
fn draw_hatch(
    content: &mut Content,
    hp: &crate::model::HatchPattern,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
) {
    use crate::model::HatchKind::*;
    if w <= 0.0 || h <= 0.0 {
        return;
    }
    content.save_state();
    fill_rgb(content, hp.bg);
    content.rect(x, y, w, h);
    content.fill_nonzero();
    // Clip subsequent strokes to the cell so the diagonal lines don't bleed out.
    content.rect(x, y, w, h);
    content.clip_nonzero();
    content.end_path();
    stroke_rgb(content, hp.fg);
    content.set_line_width(0.5);
    let s = 4.0_f32;
    if matches!(hp.kind, Horz | CrossHorzVert) {
        let mut yy = y + s;
        while yy < y + h {
            content.move_to(x, yy);
            content.line_to(x + w, yy);
            yy += s;
        }
    }
    if matches!(hp.kind, Vert | CrossHorzVert) {
        let mut xx = x + s;
        while xx < x + w {
            content.move_to(xx, y);
            content.line_to(xx, y + h);
            xx += s;
        }
    }
    if matches!(hp.kind, DiagFwd | CrossDiag) {
        let mut off = 0.0;
        while off <= w + h {
            content.move_to(x + off - h, y);
            content.line_to(x + off, y + h);
            off += s;
        }
    }
    if matches!(hp.kind, DiagBack | CrossDiag) {
        let mut off = 0.0;
        while off <= w + h {
            content.move_to(x + off - h, y + h);
            content.line_to(x + off, y);
            off += s;
        }
    }
    content.stroke();
    content.restore_state();
}

fn draw_cell_shading(
    content: &mut Content,
    shading: Option<[u8; 3]>,
    hatch: Option<crate::model::HatchPattern>,
    borders: &crate::model::CellBorders,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
) {
    let bw = |b: &crate::model::CellBorder| if b.present { b.width } else { 0.0 };
    let inset =
        (bw(&borders.top) + bw(&borders.bottom) + bw(&borders.left) + bw(&borders.right)) / 8.0;
    paint_cell_background(
        content,
        shading,
        hatch,
        x + inset,
        y + inset,
        w - 2.0 * inset,
        h - 2.0 * inset,
    );
}

/// Render an inline image within a table cell paragraph, handling alignment,
/// shadow, glow, stroke, and the image transform. Returns the image height
/// consumed so the caller can advance its cursor.
fn render_cell_inline_image(
    content: &mut Content,
    para: &CellParagraphLayout,
    img_name: &str,
    cell_x: f32,
    col_w: f32,
    cursor_y: f32,
    cm: &CellMargins,
) -> f32 {
    // The picture is the paragraph's first line, so it starts at the first-line
    // indent: nabl's logo paragraph (w:ind left=-198) sits 9.9pt into the cell
    // margin in Word.
    let left = cm.left + para.indent_left + para.indent_first_line - para.indent_hanging;
    let text_w = (col_w - left - cm.right - para.indent_right).max(0.0);
    let img_x = cell_x
        + left
        + match para.alignment {
            Alignment::Center => (text_w - para.image_width) / 2.0,
            Alignment::Right => text_w - para.image_width,
            _ => 0.0,
        };
    let img_y = cursor_y - para.image_height;

    // Word clips a picture wider than its cell to the cell's edges (nabl's
    // 90.75pt logo in an 81pt column).
    let clip = img_x < cell_x || img_x + para.image_width > cell_x + col_w;
    if clip {
        content.save_state();
        content
            .rect(cell_x, img_y - 1.0, col_w, para.image_height + 2.0)
            .clip_nonzero()
            .end_path();
    }

    if let Some(ref shadow) = para.image_shadow {
        super::color::draw_image_shadow(
            content,
            shadow,
            img_x,
            img_y,
            para.image_width,
            para.image_height,
            para.image_shadow_xobj.as_deref(),
        );
    }
    if let Some(ref glow) = para.image_glow {
        super::color::draw_image_glow(
            content,
            glow,
            img_x,
            img_y,
            para.image_width,
            para.image_height,
            para.image_glow_xobj.as_deref(),
        );
    }

    super::smartart::render_image_with_clip(
        content,
        img_name,
        img_x,
        img_y,
        para.image_width,
        para.image_height,
        para.image_clip.as_ref(),
    );

    if let Some(sc) = para.image_stroke_color {
        super::smartart::stroke_image_border(
            content,
            img_x,
            img_y,
            para.image_width,
            para.image_height,
            sc,
            para.image_stroke_width,
            para.image_clip.as_ref(),
        );
    }
    if clip {
        content.restore_state();
    }

    para.image_height
}

fn valign_offset(v_align: CellVAlign, available: f32, content_h: f32) -> f32 {
    match v_align {
        CellVAlign::Top => 0.0,
        CellVAlign::Center => ((available - content_h) / 2.0).max(0.0),
        CellVAlign::Bottom => (available - content_h).max(0.0),
    }
}

fn para_has_visible_content(para: &CellParagraphLayout) -> bool {
    !para.list_label.is_empty()
        || (!para.lines.is_empty() && para.lines.iter().any(|l| !l.chunks.is_empty()))
        || !para.floating_images.is_empty()
        || para.image_name.is_some()
        || para.has_textboxes
        || para.has_connectors
}

/// Total content height including trailing space_after, matching Word's vAlign calculation.
fn cell_content_h_for_valign(items: &[CellContentItem]) -> f32 {
    let mut h: f32 = 0.0;
    let mut floating_extent: f32 = 0.0;
    for item in items {
        match item {
            CellContentItem::Paragraph(p) => {
                h += p.space_before;
                // Anchored pictures do not enlarge an automatic row, but Word
                // includes their lower edge when distributing vAlign space.
                for fi in &p.floating_images {
                    floating_extent = floating_extent.max(h + fi.v_offset + fi.display_height);
                }
                h += para_block_height(p);
            }
            nested @ CellContentItem::NestedTable { .. } => h += nested.height(),
        }
    }
    // Word includes the last paragraph's space_after in the content block height
    // used for vertical alignment, so bottom/center-aligned cells position correctly.
    if let Some(CellContentItem::Paragraph(last_para)) = items.last() {
        h += last_para.space_after;
        // Baselines sit at a fixed font_size below each line's top, so a font
        // whose line height exceeds font_size + descent leaves its extra
        // leading dangling below the ink of the last line. Only compensate
        // when the font is a metric-changing fallback (e.g. a CJK substitute
        // with a large lineGap): there the dangling leading is our artifact,
        // whereas with the document's real font the full line box matches
        // Word's own centering.
        if last_para.font_substituted
            && (!last_para.lines.is_empty()
                || (last_para.content_height <= 0.0 && !last_para.paragraph_mark_vanish))
        {
            let ink_bottom = last_para.font_size * (1.0 + last_para.descender_ratio);
            h -= (last_para.line_h - ink_bottom).max(0.0);
        }
    }
    h.max(floating_extent)
}

fn cell_has_visible_content(items: &[CellContentItem]) -> bool {
    items.iter().any(|item| match item {
        CellContentItem::Paragraph(p) => para_has_visible_content(p),
        CellContentItem::NestedTable { rows, .. } => rows.iter().any(|r| r.height > 0.0),
    })
}

/// Column span to tag a cell with. A row that starts late (`w:gridBefore`)
/// lends the skipped columns to its first cell and one that stops short
/// (`w:gridAfter`) the rest to its last cell, so every TR spans the table's
/// columns (PDF/UA 7.2-42/43).
fn row_tag_span(
    row: &TableRow,
    ci: usize,
    span: usize,
    grid_cols: usize,
    grid_col_after: usize,
) -> i32 {
    let lead = if ci == 0 { row.grid_before } else { 0 };
    let span = if ci + 1 == row.cells.len() {
        span + grid_cols.saturating_sub(grid_col_after)
    } else {
        span
    };
    (lead + span) as i32
}

/// The list label of a tagged cell list item goes in its Lbl, the rest of the
/// paragraph in its LBody.
fn draw_tagged_cell_label(
    content: &mut Content,
    tagger: &mut Option<CellTagger<'_>>,
    nodes: Option<(Option<usize>, usize)>,
    para: &CellParagraphLayout,
    label_x: f32,
    baseline_y: f32,
    fonts: &HashMap<String, FontEntry>,
) {
    match (tagger.as_mut(), nodes) {
        (Some(t), Some((Some(label), body))) => {
            t.switch(content, label);
            draw_cell_label(content, para, label_x, baseline_y, fonts);
            t.switch(content, body);
        }
        _ => draw_cell_label(content, para, label_x, baseline_y, fonts),
    }
}

/// Links and footnote references in a tagged cell paragraph nest in its P.
fn cell_link_tagger<'a>(
    tagger: &'a mut Option<CellTagger<'_>>,
    para: Option<usize>,
) -> Option<LinkTagger<'a>> {
    let t = tagger.as_mut()?;
    Some(LinkTagger::new(&mut *t.tags, t.page, para?))
}

fn end_cell_tag(content: &mut Content, tagger: &Option<CellTagger<'_>>) {
    if tagger.is_some() {
        Tags::end(content);
    }
}

fn render_cell_content(
    content: &mut Content,
    items: &[CellContentItem],
    blocks: &[Block],
    cell_x: f32,
    col_w: f32,
    cursor_y_start: f32,
    // vAlign centering shift already folded into cursor_y_start. Word anchors
    // paragraph-relative floats to the cell's content top, unaffected by the
    // vAlign redistribution, so float anchors add this back.
    valign_off: f32,
    cm: &CellMargins,
    ctx: &RenderContext,
    gradient_specs: &mut Vec<super::GradientSpec>,
    links: &mut Vec<LinkAnnotation>,
    mut tagger: Option<CellTagger<'_>>,
) {
    let mut cursor_y = cursor_y_start;
    let mut block_idx = 0;

    for (item_idx, item) in items.iter().enumerate() {
        match item {
            CellContentItem::Paragraph(para) => {
                // Advance block_idx past the corresponding Block::Paragraph
                let mut source_para: Option<&Paragraph> = None;
                while block_idx < blocks.len() {
                    if let Block::Paragraph(p) = &blocks[block_idx] {
                        source_para = Some(p);
                        block_idx += 1;
                        break;
                    }
                    block_idx += 1;
                }

                // Word tags every cell paragraph, empty ones included.
                let cell_nodes = tagger.as_mut().map(|t| {
                    t.begin(
                        content,
                        item_idx,
                        para.list_item,
                        !para.list_label.is_empty(),
                    )
                });
                let cell_para = cell_nodes.map(|(_, body)| body);
                let para_top = cursor_y;
                if !para_has_visible_content(para) && !para.has_textboxes && !para.has_connectors {
                    cursor_y -= para.space_before + para_block_height(para);
                    end_cell_tag(content, &tagger);
                    continue;
                }

                cursor_y -= para.space_before;
                // Word anchors paragraph-relative floats to the cell's content
                // top, unaffected by vAlign (see `valign_off`).
                let float_y = cursor_y + valign_off;

                // Word stacks anchored objects by relativeHeight: a picture
                // above the paragraph's connectors/textboxes must paint after
                // them (annotation #241: an opaque label cutting a vertical
                // arrow). Everything else keeps the picture-before-text order.
                let shape_z = source_para.and_then(|p| {
                    p.connectors
                        .iter()
                        .map(|c| c.z_index)
                        .chain(p.textboxes.iter().map(|t| t.z_index))
                        .max()
                });
                let above_shapes =
                    |fi: &&CellFloatingImageLayout| shape_z.is_some_and(|z| fi.z_index > z);
                for fi in para.floating_images.iter().filter(|fi| !above_shapes(fi)) {
                    cell_figure(content, &mut tagger, cell_para, fi.tagging(), |c| {
                        draw_cell_float(c, fi, cell_x, float_y)
                    });
                }

                if let Some(ref img_name) = para.image_name {
                    for fi in para.floating_images.iter().filter(above_shapes) {
                        cell_figure(content, &mut tagger, cell_para, fi.tagging(), |c| {
                            draw_cell_float(c, fi, cell_x, float_y)
                        });
                    }
                    // distT/distB in layout_extra_height contribute to row
                    // height but don't add spacing between image and text.
                    // The paragraph's P stays empty and the picture follows
                    // it as a Figure, as in the body.
                    let picture = (para.image_alt.as_deref(), para.image_decorative);
                    cursor_y -= cell_figure(content, &mut tagger, None, picture, |c| {
                        render_cell_inline_image(c, para, img_name, cell_x, col_w, cursor_y, cm)
                    });
                    continue;
                }

                let text_x = cell_x + cm.left + para.indent_left + para.float_indent_left;
                // Must match the wrap width in table_layout.rs (incl. indent_right),
                // or centered text shifts right and justify overshoots the border.
                let text_w = (col_w
                    - cm.left
                    - cm.right
                    - para.indent_left
                    - para.indent_right
                    - para.float_indent_left)
                    .max(0.0);
                let baseline_y = cursor_y - para.font_size * para.ascender_ratio;

                // Word clips cell text at the cell's horizontal edges, even
                // when a negative paragraph indent puts a glyph outside it.
                // Keep this around inline text/labels only: anchored shapes
                // and nested floating tables have their own drawing bounds.
                let page_h = ctx
                    .sections
                    .iter()
                    .map(|s| s.properties.page_height)
                    .fold(cursor_y_start.max(0.0), f32::max);
                content.save_state();
                content
                    .rect(cell_x, 0.0, col_w, page_h)
                    .clip_nonzero()
                    .end_path();

                let first_line_hanging = if para.list_label.is_empty() {
                    para.text_hanging
                } else {
                    let label_x = cell_x + cm.left + para.indent_left - para.indent_hanging
                        + para.indent_first_line;
                    draw_tagged_cell_label(
                        content,
                        &mut tagger,
                        cell_nodes,
                        para,
                        label_x,
                        baseline_y,
                        ctx.fonts,
                    );
                    para.text_hanging
                };

                render_paragraph_lines(
                    content,
                    &para.lines,
                    &para.alignment,
                    text_x,
                    text_w,
                    baseline_y,
                    para.line_h,
                    (para.font_size, para.font_size * para.descender_ratio),
                    para.lines.len(),
                    0,
                    links,
                    first_line_hanging,
                    ctx.fonts,
                    None,
                    gradient_specs,
                    None,
                    None,
                    cell_link_tagger(&mut tagger, cell_para),
                );
                content.restore_state();
                end_cell_tag(content, &tagger);

                cursor_y -= super::table_layout::para_block_height(para);

                if let Some(src) = source_para {
                    render_cell_floating_shapes(
                        content,
                        src,
                        cell_x,
                        col_w,
                        para_top + valign_off,
                        ctx,
                        gradient_specs,
                        links,
                        &mut tagger,
                    );
                }
                for fi in para.floating_images.iter().filter(above_shapes) {
                    cell_figure(content, &mut tagger, None, fi.tagging(), |c| {
                        draw_cell_float(c, fi, cell_x, float_y)
                    });
                }
            }
            item @ CellContentItem::NestedTable {
                col_widths,
                rows,
                space_before,
                floating_offset,
                floating_x_offset,
            } => {
                // Find the corresponding Block::Table
                let table = loop {
                    if block_idx >= blocks.len() {
                        break None;
                    }
                    if let Block::Table(t) = &blocks[block_idx] {
                        block_idx += 1;
                        break Some(t);
                    }
                    block_idx += 1;
                };
                if let Some(table) = table {
                    let saved_y = cursor_y;
                    cursor_y -= space_before + floating_offset.unwrap_or(0.0);
                    render_nested_table(
                        table,
                        content,
                        cell_x + cm.left + floating_x_offset.unwrap_or(0.0),
                        col_w - cm.left - cm.right,
                        &mut cursor_y,
                        ctx,
                        gradient_specs,
                        links,
                        &mut tagger,
                        (col_widths, rows),
                        &Chunk {
                            item: item_idx,
                            l0: 0,
                            l1: None,
                            from: &[],
                            to: &[],
                        },
                    );
                    if floating_offset.is_some() {
                        cursor_y = saved_y;
                    }
                } else {
                    cursor_y -= item.height();
                }
            }
        }
    }
}

/// Draw a picture in a cell as a Figure inside the cell element, as Word tags
/// one (an artifact when decorative), then go back into the cell paragraph
/// `resume` while it is still open.
fn cell_figure<R>(
    content: &mut Content,
    tagger: &mut Option<CellTagger<'_>>,
    resume: Option<usize>,
    (alt, decorative): (Option<&str>, bool),
    draw: impl FnOnce(&mut Content) -> R,
) -> R {
    let Some(t) = tagger.as_mut() else {
        return draw(content);
    };
    if decorative {
        Tags::end(content);
    } else {
        let figure = t.figure(alt);
        t.switch(content, figure);
    }
    let drawn = draw(content);
    match resume {
        Some(para) => t.switch(content, para),
        None => Tags::end(content),
    }
    drawn
}

/// Draw one floating picture anchored to a cell paragraph whose top is `para_y`.
fn draw_cell_float(content: &mut Content, fi: &CellFloatingImageLayout, cell_x: f32, para_y: f32) {
    let fi_x = cell_x + fi.h_offset;
    let fi_y_bottom = para_y - fi.v_offset - fi.display_height;
    content.save_state();
    super::positioning::push_center_rotation(
        content,
        fi_x,
        fi_y_bottom,
        fi.display_width,
        fi.display_height,
        fi.rotation_deg,
    );
    content.transform([
        fi.display_width,
        0.0,
        0.0,
        fi.display_height,
        fi_x,
        fi_y_bottom,
    ]);
    content.x_object(Name(fi.pdf_name.as_bytes()));
    content.restore_state();
}

/// Render floating textboxes and connectors anchored to a cell paragraph.
fn render_cell_floating_shapes(
    content: &mut Content,
    para: &Paragraph,
    cell_x: f32,
    col_w: f32,
    para_top: f32,
    ctx: &RenderContext,
    gradient_specs: &mut Vec<super::GradientSpec>,
    links: &mut Vec<LinkAnnotation>,
    tagger: &mut Option<CellTagger<'_>>,
) {
    use super::positioning::render_connector;
    use crate::model::HorizontalPosition;

    // Positioned relative to the cell column origin and paragraph top.
    for conn in &para.connectors {
        let x = conn.h_position.place(cell_x, col_w, conn.width);
        render_connector(
            conn,
            content,
            x,
            para_top - conn.v_position.offset_or_zero(),
        );
    }

    for tb in &para.textboxes {
        let h_off = match tb.h_position {
            HorizontalPosition::Offset(o) => o,
            HorizontalPosition::AlignCenter => (col_w - tb.width_pt) / 2.0,
            HorizontalPosition::AlignRight => col_w - tb.width_pt,
            HorizontalPosition::AlignLeft => 0.0,
        };
        let tb_x = cell_x + h_off;
        let tb_y_top = para_top - tb.v_offset_pt;
        let tag = tagger
            .as_mut()
            .filter(|_| !tb.paragraphs.is_empty())
            .map(|t| {
                let sect = t.sect();
                (&mut *t.tags, t.page, sect)
            });
        render_simple_textbox(content, tb, tb_x, tb_y_top, ctx, gradient_specs, links, tag);
    }
}

/// Minimal textbox rendering for cell contexts: stroke/fill the shape and
/// render text content. Doesn't support all features that body textboxes do.
fn render_simple_textbox(
    content: &mut Content,
    tb: &crate::model::Textbox,
    tb_x: f32,
    tb_y_top: f32,
    ctx: &RenderContext,
    gradient_specs: &mut Vec<super::GradientSpec>,
    links: &mut Vec<LinkAnnotation>,
    // As in `render_textbox_paragraphs`: the Sect the text is tagged in.
    tag: Option<(&mut Tags, usize, usize)>,
) {
    use super::smartart::{draw_shape_path, draw_shape_stroke_path};
    use super::textbox_render::render_textbox_paragraphs;
    use crate::model::ShapeFill;

    let tb_width = tb.width_pt;
    let tb_height = super::textbox_render::textbox_height(tb, ctx);

    // Fill
    if let Some(ShapeFill::Solid(color)) = &tb.fill {
        content.save_state();
        fill_rgb(content, *color);
        draw_shape_path(
            content,
            tb_x,
            tb_y_top - tb_height,
            tb_width,
            tb_height,
            &tb.shape_type,
        );
        content.fill_nonzero();
        content.restore_state();
    }

    // Stroke
    if let Some(stroke) = tb.stroke_color
        && tb.stroke_width > 0.0
    {
        content.save_state();
        content.set_line_width(tb.stroke_width);
        stroke_rgb(content, stroke);
        // Honor per-subpath stroke flags (brace/bracket pairs have a
        // fill-only outline subpath), same as body-anchored textboxes.
        draw_shape_stroke_path(
            content,
            tb_x,
            tb_y_top - tb_height,
            tb_width,
            tb_height,
            &tb.shape_type,
        );
        content.stroke();
        content.restore_state();
    }

    // Text: route through the shared textbox renderer (same path as body
    // textboxes).
    let content_x = tb_x + tb.margin_left;
    let content_w = (tb_width - tb.margin_left - tb.margin_right).max(0.0);
    render_textbox_paragraphs(
        &tb.paragraphs,
        content,
        content_x,
        content_w,
        content_w,
        tb_y_top - tb.margin_top,
        0.0,
        0.0,
        None,
        true,
        links,
        ctx,
        None,
        gradient_specs,
        tag,
    );
}

/// Render a nested table inline within a parent cell at the given cursor position.
/// Render every row of `table` — cell backgrounds and content, then borders —
/// advancing `cursor_y`. Shared by nested and header/footer table rendering.
fn render_table_rows(
    table: &Table,
    row_layouts: &[RowLayout],
    col_widths: &[f32],
    table_left: f32,
    merge_spans: &HashMap<(usize, usize), f32>,
    content: &mut Content,
    cursor_y: &mut f32,
    ctx: &RenderContext,
    gradient_specs: &mut Vec<super::GradientSpec>,
    links: &mut Vec<LinkAnnotation>,
    // (tags, this table's structure, page) for a nested table in a tagged
    // cell; header/footer tables stay artifacts.
    mut tag: Option<(&mut Tags, &mut TableTags, usize)>,
    rows: std::ops::Range<usize>,
) {
    for (ri, (row, layout)) in table
        .rows
        .iter()
        .zip(row_layouts.iter())
        .enumerate()
        .skip(rows.start)
        .take(rows.len())
    {
        let row_h = layout.height;
        let row_top = *cursor_y;
        let row_bottom = row_top - row_h;

        for (ci, ((grid_col, span, cell), cell_layout)) in
            row.grid_cells().zip(layout.cells.iter()).enumerate()
        {
            let col_w = cell_span_width(col_widths, grid_col, span);
            let cx = cell_x_offset(col_widths, table_left, grid_col);
            let tagger = tag.as_mut().map(|(tags, table_tags, page)| CellTagger {
                tags,
                table: table_tags,
                page: *page,
                row: ri,
                cell: ci,
                col_span: row_tag_span(row, ci, span, col_widths.len(), grid_col + span),
            });

            if cell.v_merge == VMerge::Continue {
                if let Some(mut t) = tagger {
                    t.empty_cell();
                }
                continue;
            }

            let merge_extra = merge_spans.get(&(ri, grid_col)).copied().unwrap_or(0.0);
            let effective_h = row_h + merge_extra;

            paint_cell_background(
                content,
                cell.shading,
                cell.hatch,
                cx,
                row_bottom,
                col_w,
                row_h,
            );

            if cell_has_visible_content(&cell_layout.items) {
                let ecm = &cell_layout.cm;
                let content_h = cell_content_h_for_valign(&cell_layout.items);

                let avail = effective_h - ecm.top - ecm.bottom;
                let v_offset = valign_offset(cell.v_align, avail, content_h);
                let cell_cursor_y = row_top - ecm.top - v_offset;

                render_cell_content(
                    content,
                    &cell_layout.items,
                    &cell.content,
                    cx,
                    col_w,
                    cell_cursor_y,
                    v_offset,
                    ecm,
                    ctx,
                    gradient_specs,
                    links,
                    tagger,
                );
            } else if let Some(t) = tagger {
                t.empty_para(content);
            }
        }

        for (grid_col, span, cell) in row.grid_cells() {
            let col_w = cell_span_width(col_widths, grid_col, span);
            let bx = cell_x_offset(col_widths, table_left, grid_col);

            if cell.v_merge == VMerge::Continue {
                continue;
            }

            let merge_extra = merge_spans.get(&(ri, grid_col)).copied().unwrap_or(0.0);
            let effective_bottom = row_bottom - merge_extra;

            draw_cell_borders(
                content,
                &cell.borders,
                bx,
                row_top,
                effective_bottom,
                col_w,
                ri == 0,
                true,
            );
        }

        *cursor_y = row_bottom;
    }
}

/// Reborrow a row renderer's tagging context for one more call.
fn reborrow<'a>(
    tag: &'a mut Option<(&mut Tags, &mut TableTags, usize)>,
) -> Option<(&'a mut Tags, &'a mut TableTags, usize)> {
    tag.as_mut().map(|(t, n, p)| (&mut **t, &mut **n, *p))
}

/// Draw `chunk` of a nested table (all of it, or one page's piece of a row
/// split) from the cell layout's own `(col_widths, rows)`.
fn render_nested_table(
    table: &Table,
    content: &mut Content,
    available_x: f32,
    available_w: f32,
    cursor_y: &mut f32,
    ctx: &RenderContext,
    gradient_specs: &mut Vec<super::GradientSpec>,
    links: &mut Vec<LinkAnnotation>,
    // The parent cell's tagger: the nested table is tagged inside that cell.
    tagger: &mut Option<CellTagger<'_>>,
    (col_widths, row_layouts): (&[f32], &[RowLayout]),
    chunk: &Chunk,
) {
    let table_total_w: f32 = col_widths.iter().sum();
    let table_left = match table.alignment {
        TableAlignment::Center => available_x + (available_w - table_total_w) / 2.0,
        TableAlignment::Right => available_x + available_w - table_total_w,
        TableAlignment::Left => available_x + table.table_indent,
    };

    let merge_spans = compute_merge_spans(table, row_layouts);

    let mut tag = tagger.as_mut().map(|t| t.nested_table(table, chunk.item));
    let n_rows = table.rows.len().min(row_layouts.len());
    for piece in row_pieces(n_rows, chunk) {
        match piece {
            RowPiece::Rows(range) => render_table_rows(
                table,
                row_layouts,
                col_widths,
                table_left,
                &merge_spans,
                content,
                cursor_y,
                ctx,
                gradient_specs,
                links,
                reborrow(&mut tag),
                range,
            ),
            RowPiece::Partial { row, starts, ends } => render_partial_row(
                &table.rows[row],
                &row_layouts[row],
                col_widths,
                table_left,
                content,
                cursor_y,
                ctx,
                gradient_specs,
                links,
                reborrow(&mut tag),
                starts,
                ends,
                row,
            ),
        }
    }
    if let Some((tags, nested, _)) = tag {
        nested.finish(tags);
    }
}

fn render_partial_cell_content(
    content: &mut Content,
    items: &[CellContentItem],
    blocks: &[Block],
    start: &CellCursor,
    end: &CellCursor,
    cell_x: f32,
    col_w: f32,
    cursor_y_start: f32,
    cm: &CellMargins,
    ctx: &RenderContext,
    gradient_specs: &mut Vec<super::GradientSpec>,
    links: &mut Vec<LinkAnnotation>,
    mut tagger: Option<CellTagger<'_>>,
) {
    let mut cursor_y = cursor_y_start;
    let mut reanchor_shift = 0.0;
    // Build a mapping from item index to block index
    let mut block_idx = 0usize;
    let mut item_to_block: Vec<usize> = Vec::new();
    for item in items {
        item_to_block.push(block_idx);
        match item {
            CellContentItem::Paragraph(_) => {
                while block_idx < blocks.len() {
                    if matches!(&blocks[block_idx], Block::Paragraph(_)) {
                        block_idx += 1;
                        break;
                    }
                    block_idx += 1;
                }
            }
            CellContentItem::NestedTable { .. } => {
                while block_idx < blocks.len() {
                    if matches!(&blocks[block_idx], Block::Table(_)) {
                        block_idx += 1;
                        break;
                    }
                    block_idx += 1;
                }
            }
        }
    }

    for chunk in cursor_chunks(items, start, end) {
        let (pi, l0, l1) = (chunk.item, chunk.l0, chunk.l1);
        match &items[pi] {
            CellContentItem::Paragraph(para) => {
                let sb = chunk_space_before(&items[pi], pi, start);

                let cell_nodes = tagger
                    .as_mut()
                    .map(|t| t.begin(content, pi, para.list_item, !para.list_label.is_empty()));
                let cell_para = cell_nodes.map(|(_, body)| body);
                if !para_has_visible_content(para) {
                    cursor_y -= sb + para_block_height(para);
                    end_cell_tag(content, &tagger);
                    continue;
                }

                cursor_y -= sb;

                // Anchored pictures and the list label belong to the
                // paragraph's first line; a continuation chunk has neither.
                if l0 == 0 {
                    for fi in &para.floating_images {
                        cell_figure(content, &mut tagger, cell_para, fi.tagging(), |c| {
                            draw_cell_float(c, fi, cell_x, cursor_y)
                        });
                    }
                }

                if let Some(ref img_name) = para.image_name {
                    let picture = (para.image_alt.as_deref(), para.image_decorative);
                    cursor_y -= cell_figure(content, &mut tagger, None, picture, |c| {
                        render_cell_inline_image(c, para, img_name, cell_x, col_w, cursor_y, cm)
                    });
                    continue;
                }

                let text_x = cell_x + cm.left + para.indent_left + para.float_indent_left;
                // Must match the wrap width in table_layout.rs (incl. indent_right),
                // or centered text shifts right and justify overshoots the border.
                let text_w = (col_w
                    - cm.left
                    - cm.right
                    - para.indent_left
                    - para.indent_right
                    - para.float_indent_left)
                    .max(0.0);
                let baseline_y = cursor_y - para.font_size * para.ascender_ratio;

                let first_line_hanging = if para.list_label.is_empty() {
                    para.text_hanging
                } else {
                    if l0 == 0 {
                        let label_x = cell_x + cm.left + para.indent_left - para.indent_hanging
                            + para.indent_first_line;
                        draw_tagged_cell_label(
                            content,
                            &mut tagger,
                            cell_nodes,
                            para,
                            label_x,
                            baseline_y,
                            ctx.fonts,
                        );
                    }
                    para.text_hanging
                };

                let l1 = l1.unwrap_or(para.lines.len());
                render_paragraph_lines(
                    content,
                    &para.lines[l0..l1],
                    &para.alignment,
                    text_x,
                    text_w,
                    baseline_y,
                    para.line_h,
                    (para.font_size, para.font_size * para.descender_ratio),
                    para.lines.len(),
                    l0,
                    links,
                    first_line_hanging,
                    ctx.fonts,
                    None,
                    gradient_specs,
                    None,
                    None,
                    cell_link_tagger(&mut tagger, cell_para),
                );
                end_cell_tag(content, &tagger);

                cursor_y -= if para.lines.is_empty() {
                    para_block_height(para)
                } else {
                    super::table_layout::cell_lines_h(para, l0..l1)
                };
            }
            CellContentItem::NestedTable {
                col_widths,
                rows,
                space_before,
                floating_offset,
                floating_x_offset,
            } => {
                let bi = item_to_block.get(pi).copied().unwrap_or(0);
                if let Some(Block::Table(table)) = blocks.get(bi) {
                    let saved_y = cursor_y;
                    if chunk.l0 == 0 && chunk.from.is_empty() {
                        let drop_anchor = pi == start.item && start.item > 0 && start.line == 0;
                        let gap = if let Some(offset) = floating_offset {
                            super::table_layout::floating_anchor_gap(
                                *space_before,
                                *offset,
                                floating_x_offset.is_some(),
                                drop_anchor,
                                &mut reanchor_shift,
                            )
                        } else if drop_anchor {
                            0.0
                        } else {
                            *space_before
                        };
                        cursor_y -= gap;
                    }
                    render_nested_table(
                        table,
                        content,
                        cell_x
                            + cm.left
                            + if chunk.l0 == 0
                                && chunk.from.is_empty()
                                && pi == start.item
                                && start.item > 0
                                && start.line == 0
                            {
                                0.0
                            } else {
                                floating_x_offset.unwrap_or(0.0)
                            },
                        col_w - cm.left - cm.right,
                        &mut cursor_y,
                        ctx,
                        gradient_specs,
                        links,
                        &mut tagger,
                        (col_widths, rows),
                        &chunk,
                    );
                    if floating_offset.is_some() {
                        cursor_y = saved_y;
                    }
                } else {
                    cursor_y -= item_chunk_height(&items[pi], &chunk);
                }
            }
        }
    }
}

fn draw_cell_label(
    content: &mut Content,
    para: &CellParagraphLayout,
    label_x: f32,
    baseline_y: f32,
    fonts: &HashMap<String, FontEntry>,
) {
    let font_key = para
        .list_label_font
        .as_deref()
        .unwrap_or(para.first_run_font_key.as_str());
    let Some(entry) = fonts.get(font_key) else {
        return;
    };
    let bytes = super::list_label::encode_label(entry, &para.list_label)
        .unwrap_or_else(|| entry.encode(&para.list_label));

    if let Some(c) = para.label_color {
        fill_rgb(content, c);
    }
    content
        .begin_text()
        .set_font(Name(entry.pdf_name.as_bytes()), para.font_size)
        .next_line(label_x + para.label_shift, baseline_y)
        .show(Str(&bytes))
        .end_text();
    if para.label_color.is_some() {
        content.set_fill_gray(0.0);
    }
}

fn render_table_row(
    row: &TableRow,
    layout: &RowLayout,
    col_widths: &[f32],
    table_left: f32,
    pb: &mut super::PageBuilder,
    ctx: &RenderContext,
    row_idx: usize,
    merge_spans: &HashMap<(usize, usize), f32>,
) {
    let row_h = layout.height;
    let row_top = pb.slot_top;
    let row_bottom = row_top - row_h;

    for (ci, ((grid_col, span, cell), cell_layout)) in
        row.grid_cells().zip(layout.cells.iter()).enumerate()
    {
        let col_w = cell_span_width(col_widths, grid_col, span);
        let cell_x = cell_x_offset(col_widths, table_left, grid_col);
        // None while repeated header rows are drawn (render_header_rows).
        let page = pb.all_contents.len();
        let tagger = pb.table_tags.as_mut().map(|table| CellTagger {
            tags: &mut pb.tags,
            table,
            page,
            row: row_idx,
            cell: ci,
            col_span: row_tag_span(row, ci, span, col_widths.len(), grid_col + span),
        });

        if cell.v_merge == VMerge::Continue {
            if let Some(mut t) = tagger {
                t.empty_cell();
            }
            continue;
        }

        let merge_extra = merge_spans
            .get(&(row_idx, grid_col))
            .copied()
            .unwrap_or(0.0);
        let effective_h = row_h + merge_extra;

        draw_cell_shading(
            &mut pb.content,
            cell.shading,
            cell.hatch,
            &cell.borders,
            cell_x,
            row_top - effective_h,
            col_w,
            effective_h,
        );

        let has_content = cell_has_visible_content(&cell_layout.items);
        let ecm = &cell_layout.cm;
        let mut no_links = Vec::new();

        if !has_content {
            if let Some(t) = tagger {
                t.empty_para(&mut pb.content);
            }
        } else if cell_layout.text_direction == TextDirection::TbRl {
            render_vertical_cjk_cell(
                &mut pb.content,
                cell_layout,
                cell,
                cell_x,
                row_top,
                effective_h,
                col_w,
                ecm,
                ctx,
                tagger,
            );
        } else {
            let content_h = cell_content_h_for_valign(&cell_layout.items);

            let avail = effective_h - ecm.top - ecm.bottom;
            let v_offset = valign_offset(cell.v_align, avail, content_h);
            let cursor_y = row_top - ecm.top - v_offset;

            render_cell_content(
                &mut pb.content,
                &cell_layout.items,
                &cell.content,
                cell_x,
                col_w,
                cursor_y,
                v_offset,
                ecm,
                ctx,
                &mut pb.gradient_specs,
                // Repeated header rows are untagged artifacts: no link annotations.
                if tagger.is_some() {
                    &mut pb.links
                } else {
                    &mut no_links
                },
                tagger,
            );
        }
    }

    for (grid_col, span, cell) in row.grid_cells() {
        let col_w = cell_span_width(col_widths, grid_col, span);
        let bx = cell_x_offset(col_widths, table_left, grid_col);

        // Each row draws its slice of a vertically merged cell, so a span
        // that crosses a page break ends at the page bottom and resumes on
        // the next page (nabl's "Sl." column, pages 10-11).
        let span_below = merge_spans
            .get(&(row_idx, grid_col))
            .copied()
            .unwrap_or(0.0);
        draw_cell_borders(
            &mut pb.content,
            &cell.borders,
            bx,
            row_top,
            row_bottom,
            col_w,
            cell.v_merge != VMerge::Continue,
            span_below == 0.0,
        );
    }

    pb.slot_top = row_bottom;
}

fn render_vertical_cjk_cell(
    content: &mut Content,
    cell_layout: &CellLayout,
    cell: &crate::model::TableCell,
    cell_x: f32,
    row_top: f32,
    row_h: f32,
    col_w: f32,
    cm: &CellMargins,
    ctx: &RenderContext,
    // Tags each paragraph as the cell's P, like a horizontal cell's.
    mut tagger: Option<CellTagger<'_>>,
) {
    let pdf_name_to_entry: HashMap<&str, &FontEntry> = ctx
        .fonts
        .values()
        .map(|e| (e.pdf_name.as_str(), e))
        .collect();

    let paras: Vec<&CellParagraphLayout> = cell_layout
        .items
        .iter()
        .filter_map(|item| match item {
            CellContentItem::Paragraph(p) => Some(p),
            _ => None,
        })
        .collect();
    let total_char_h: f32 = paras
        .iter()
        .flat_map(|p| &p.lines)
        .flat_map(|l| &l.chunks)
        .filter(|c| !c.text.is_empty())
        .map(|c| c.text.chars().count() as f32 * c.font_size)
        .sum();

    let avail_h = row_h - cm.top - cm.bottom;
    // In vertical text cells, paragraph jc controls vertical positioning
    let effective_v_align = if cell.v_align == CellVAlign::Top {
        match paras.first().map(|p| p.alignment) {
            Some(Alignment::Center) => CellVAlign::Center,
            Some(Alignment::Right) => CellVAlign::Bottom,
            _ => CellVAlign::Center,
        }
    } else {
        cell.v_align
    };
    let v_offset = valign_offset(effective_v_align, avail_h, total_char_h);

    let avail_w = col_w - cm.left - cm.right;
    let mut char_y = row_top - cm.top - v_offset;
    let mut char_buf = [0u8; 4];

    for (item, content_item) in cell_layout.items.iter().enumerate() {
        let CellContentItem::Paragraph(para) = content_item else {
            continue;
        };
        let nodes = tagger
            .as_mut()
            .map(|t| t.begin(content, item, para.list_item, false));
        let mut link_tags = cell_link_tagger(&mut tagger, nodes.map(|(_, body)| body));
        for line in &para.lines {
            for chunk in &line.chunks {
                if chunk.text.is_empty() {
                    continue;
                }
                let fs = chunk.font_size;
                let entry = pdf_name_to_entry.get(chunk.pdf_font.as_str());
                let ascender_ratio = entry.and_then(|e| e.ascender_ratio).unwrap_or(0.75);
                let widths = entry.and_then(|e| e.char_widths_1000.as_ref());

                if let Some(c) = chunk.color {
                    fill_rgb(content, c);
                }

                let boundary_space = chunk.boundary_space(entry.copied());
                content.begin_text();
                // ponytail: no link annotations in vertical text, so no Link
                // elements either; add both if a document has one
                if let Some(lt) = link_tags.as_mut() {
                    lt.chunk(content, chunk, None, boundary_space);
                }
                content.set_font(Name(chunk.pdf_font.as_bytes()), fs);

                let mut td_x = 0.0f32;
                let mut td_y = 0.0f32;
                for ch in chunk.text.chars() {
                    let baseline_y = char_y - fs * ascender_ratio;
                    let char_w = widths
                        .and_then(|m| m.get(&ch))
                        .map(|w| w * fs / 1000.0)
                        .unwrap_or(fs);
                    let cx = cell_x + cm.left + (avail_w - char_w) / 2.0;

                    let ch_str = ch.encode_utf8(&mut char_buf);
                    let bytes = encode_text_for_pdf(ch_str, &chunk.pdf_font, &pdf_name_to_entry);
                    content.next_line(cx - td_x, baseline_y - td_y);
                    td_x = cx;
                    td_y = baseline_y;
                    content.show(Str(&bytes));

                    char_y -= fs;
                }
                if boundary_space {
                    let space = encode_text_for_pdf(" ", &chunk.pdf_font, &pdf_name_to_entry);
                    content.show(Str(&space));
                }
                content.end_text();

                if chunk.color.is_some() {
                    content.set_fill_gray(0.0);
                }
            }
        }
        if let Some(lt) = link_tags {
            lt.finish(content);
        }
        end_cell_tag(content, &tagger);
    }
}

/// Render a subset of each cell's paragraphs for a split row.
/// `starts[ci]..ends[ci]` gives the paragraph range for cell `ci`.
///
/// Word closes every fragment of a split row as a complete box: the fragment
/// ends right after its last fitted item (not at the page's content bottom)
/// with the cell's bottom border drawn there, and the continuation on the
/// next page starts with the cell's top border again.
fn render_partial_row(
    row: &TableRow,
    layout: &RowLayout,
    col_widths: &[f32],
    table_left: f32,
    content: &mut Content,
    cursor_y: &mut f32,
    ctx: &RenderContext,
    gradient_specs: &mut Vec<super::GradientSpec>,
    links: &mut Vec<LinkAnnotation>,
    // (tags, this table's structure, page), as in render_table_rows.
    mut tag: Option<(&mut Tags, &mut TableTags, usize)>,
    starts: &[CellCursor],
    ends: &[CellCursor],
    row_idx: usize,
) {
    let row_top = *cursor_y;
    let row_h = partial_row_height(layout, starts, ends);
    let row_bottom = row_top - row_h;

    for (ci, ((grid_col, span, cell), cell_layout)) in
        row.grid_cells().zip(layout.cells.iter()).enumerate()
    {
        let col_w = cell_span_width(col_widths, grid_col, span);
        let cell_x = cell_x_offset(col_widths, table_left, grid_col);
        let tagger = tag.as_mut().map(|(tags, table, page)| CellTagger {
            tags,
            table,
            page: *page,
            row: row_idx,
            cell: ci,
            col_span: row_tag_span(row, ci, span, col_widths.len(), grid_col + span),
        });

        if cell.v_merge == VMerge::Continue {
            if let Some(mut t) = tagger {
                t.empty_cell();
            }
            continue;
        }

        let start = starts.get(ci).unwrap_or(&CELL_START);
        let done = CellCursor::at(cell_layout.items.len(), 0);
        let end = ends.get(ci).unwrap_or(&done);

        draw_cell_shading(
            content,
            cell.shading,
            cell.hatch,
            &cell.borders,
            cell_x,
            row_bottom,
            col_w,
            row_h,
        );

        let has_content = cursor_chunks(&cell_layout.items, start, end).any(|c| match &cell_layout
            .items[c.item]
        {
            CellContentItem::Paragraph(p) => para_has_visible_content(p),
            CellContentItem::NestedTable { rows, .. } => rows.iter().any(|r| r.height > 0.0),
        });

        if has_content {
            render_partial_cell_content(
                content,
                &cell_layout.items,
                &cell.content,
                start,
                end,
                cell_x,
                col_w,
                row_top - cell_layout.cm.top,
                &cell_layout.cm,
                ctx,
                gradient_specs,
                links,
                tagger,
            );
        } else if let Some(t) = tagger.filter(|_| *start == CELL_START) {
            t.empty_para(content);
        }
    }

    // A continued row tops its page with each cell's own top border, not the
    // edge it shares with the row above (radiographer's nested row, `nil` on
    // top, starts page 2 with no line).
    let continued = starts.iter().any(|s| *s != CELL_START);
    for (grid_col, span, cell) in row.grid_cells() {
        let col_w = cell_span_width(col_widths, grid_col, span);
        let bx = cell_x_offset(col_widths, table_left, grid_col);
        let mut borders = cell.borders;
        if continued && let Some(own) = borders.own_top {
            borders.top = own;
        }

        // A merged cell's slice, as in render_table_row; every chunk of a
        // split row closes at its page bottom.
        draw_cell_borders(
            content,
            &borders,
            bx,
            row_top,
            row_bottom,
            col_w,
            cell.v_merge != VMerge::Continue,
            true,
        );
    }

    *cursor_y = row_bottom;
}

fn render_header_rows(
    table: &Table,
    row_layouts: &[RowLayout],
    col_widths: &[f32],
    table_left: f32,
    pb: &mut super::PageBuilder,
    ctx: &RenderContext,
    merge_spans: &HashMap<(usize, usize), f32>,
    header_count: usize,
) {
    // Repeated header rows are page furniture: tagged once, where they first appear.
    let table_tags = pb.table_tags.take();
    for hi in 0..header_count {
        render_table_row(
            &table.rows[hi],
            &row_layouts[hi],
            col_widths,
            table_left,
            pb,
            ctx,
            hi,
            merge_spans,
        );
    }
    pb.table_tags = table_tags;
}

/// Word 2013+ layout puts the outer edge of the table's left border band at
/// the indent, so the border grid (drawn centred) sits half a band further
/// right, and cell text follows it: rehab_centre's 2.25pt bands span
/// 72.0–74.25 from a 72pt margin and text starts 5.4 + 1.125 in.
fn word2013_border_shift(table: &Table) -> f32 {
    table
        .first_cell()
        .map_or(0.0, |c| c.borders.left.band() / 2.0)
}

/// A non-floating table's left edge in the area it is aligned in.
fn aligned_table_left(
    table: &Table,
    (area_left, area_width): (f32, f32),
    total_width: f32,
    compat_mode: u32,
) -> f32 {
    use crate::model::TableAlignment;
    match table.alignment {
        TableAlignment::Center => area_left + (area_width - total_width) / 2.0,
        TableAlignment::Right => area_left + area_width - total_width,
        // Before Word 2013 layout, tblInd positions the first cell's text, so
        // the edge sits one cell margin further out, even for an explicit
        // tblInd of 0 (chinese_costume). Word 2013+ (compat 15) never outdents:
        // the border sits at the margin and the text inside it.
        TableAlignment::Left if compat_mode >= 15 => {
            area_left + table.table_indent + word2013_border_shift(table)
        }
        TableAlignment::Left => area_left + table.table_indent - table.first_cell_left_margin(),
    }
}

/// `override_pos`: positioning info for floating tables.
pub(super) fn render_table(
    table: &Table,
    sp: &SectionProperties,
    ctx: &RenderContext,
    pb: &mut super::PageBuilder,
    sect_idx: usize,
    prev_space_after: f32,
    mut override_pos: Option<super::FloatingTablePos>,
    footnotes: &std::collections::HashMap<u32, crate::model::Footnote>,
    effective_margin_bottom: &mut f32,
    column_bounds: Option<(f32, f32)>,
) {
    let available_w = column_bounds.map(|(_, w)| w);
    // Top-level table: an AutoFit-to-Window table fills the available width
    // (page text column or newspaper column), so feed the content width even
    // when no column bound is set. Passed as the dedicated fill target so it
    // only steers the window-fill path, not the nested-table shrink path.
    let fit_w = available_w.unwrap_or(sp.page_width - sp.margin_left - sp.margin_right);
    let mut col_widths = if table.fixed_layout {
        table.col_widths.clone()
    } else if let Some(col_w) = available_w {
        // A table in a newspaper column sizes like one on the page, then is
        // squeezed into the column (the nested-table path shrank it to content).
        let mut w = auto_fit_columns(table, ctx.fonts, None, Some(fit_w));
        squeeze_to_width(&mut w, &min_content_widths(table, ctx.fonts), col_w);
        w
    } else {
        auto_fit_columns(table, ctx.fonts, available_w, Some(fit_w))
    };
    apply_pct_width(
        table,
        &mut col_widths,
        available_w.unwrap_or(sp.page_width - sp.margin_left - sp.margin_right),
    );
    let row_layouts = compute_row_layouts(table, &col_widths, ctx, None);
    let merge_spans = compute_merge_spans(table, &row_layouts);

    let preceding_float_zone = table
        .position
        .as_ref()
        .filter(|p| !p.allow_overlap)
        .and_then(|_| pb.float_zone.clone());
    if table.position.as_ref().is_some_and(|p| !p.allow_overlap)
        && let Some(ref zone) = pb.float_zone
        && let Some(ref mut fp) = override_pos
    {
        let width: f32 = col_widths.iter().sum();
        let height: f32 = row_layouts.iter().map(|r| r.height).sum();
        if fp.x < zone.obj_right
            && fp.x + width > zone.obj_left
            && fp.y > zone.bottom_y
            && fp.y - height < zone.top_y
        {
            fp.y = zone.bottom_y - fp.top_from_text;
            if let Some(pos) = table.position.as_ref() {
                fp.constrain_nonoverlap_left(
                    pos,
                    sp,
                    column_bounds.map_or(sp.margin_left, |(left, _)| left),
                    ctx.compat_mode,
                );
            }
        }
    }
    let is_floating = override_pos.is_some();
    // A text-anchored floating table sitting at or below its anchor paginates like
    // an inline table: Word starts it in the room left on the page and breaks it
    // across pages, between rows (indigenous_innovation p1-2, trHeight rows) or
    // inside a row (croatian_grant_guidelines p4-5) — annotations #232 #200. Only
    // a table hoisted above its anchor (negative tblpY, pendulum_mechanics) moves
    // whole to the next page with the anchor, and only then may a row overflow be
    // split regardless of its own splittability.
    let flows_inline = override_pos
        .as_ref()
        .is_some_and(|fp| fp.v_anchor_text && fp.v_offset_pt >= 0.0);
    let keep_with_anchor = is_floating && !flows_inline;
    let (table_left, saved_slot_top, text_margins) = if let Some(ref fp) = override_pos {
        // The float zone starts at the table's drawn top, which sits below fp.y
        // by the space after the previous paragraph (see below).
        let saved = Some((pb.slot_top - prev_space_after, fp.y - prev_space_after));
        pb.slot_top = fp.y;
        (fp.x, saved, (fp.top_from_text, fp.bottom_from_text))
    } else {
        let area = column_bounds.unwrap_or((
            sp.margin_left,
            sp.page_width - sp.margin_left - sp.margin_right,
        ));
        let total_w: f32 = col_widths.iter().sum();
        let left = aligned_table_left(table, area, total_w, ctx.compat_mode);
        (left, None, (0.0, 0.0))
    };

    // For non-floating tables, prev_space_after offsets the table start.
    // For floating tables, it was already consumed into the saved cursor position.
    // At a page top a table follows the paragraphs' rule with no space before:
    // radiographer's section 2 table starts at its header's bottom, not 10pt below.
    pb.slot_top -= if is_floating {
        prev_space_after
    } else {
        pb.page_top_gap(sp, 0.0, prev_space_after)
            .unwrap_or(prev_space_after)
    };
    // The outer border bands sit inside the table's flow height: the top band
    // starts where the previous text ends and the next paragraph starts below
    // the bottom band. The row insets hold the inner halves (docx::tables).
    let outer_band = |row: Option<&crate::model::TableRow>, top: bool| {
        row.map_or(0.0, |r| {
            r.cells
                .iter()
                .map(|c| {
                    if top {
                        c.borders.top.band()
                    } else {
                        c.borders.bottom.band()
                    }
                })
                .fold(0.0f32, f32::max)
        })
    };
    pb.slot_top -= outer_band(table.rows.first(), true) / 2.0;

    // Count contiguous header rows from the start of the table (per OOXML spec,
    // only contiguous header rows starting from row 0 are repeated).
    let header_count = table.rows.iter().take_while(|r| r.is_header).count();

    // Pre-scan each row for footnote and endnote references so we can reserve
    // space as rows containing notes are rendered.
    let row_footnote_ids: Vec<Vec<u32>> = table
        .rows
        .iter()
        .map(|row| {
            let mut ids = Vec::new();
            for cell in &row.cells {
                for p in cell.all_paragraphs() {
                    for run in &p.runs {
                        if let Some(id) = run.footnote_id {
                            ids.push(id);
                        }
                    }
                }
            }
            ids
        })
        .collect();
    let row_endnote_ids: Vec<Vec<u32>> = table
        .rows
        .iter()
        .map(|row| {
            let mut ids = Vec::new();
            for cell in &row.cells {
                for p in cell.all_paragraphs() {
                    for run in &p.runs {
                        if let Some(id) = run.endnote_id {
                            ids.push(id);
                        }
                    }
                }
            }
            ids
        })
        .collect();

    // Text width for footnote height computation (same as paragraph layout uses).
    // A table's notes go to the foot of its column (the page in one column).
    let fn_col = column_bounds.unwrap_or((
        sp.margin_left,
        sp.page_width - sp.margin_left - sp.margin_right,
    ));
    let fn_text_width = fn_col.1;

    let flush_and_render_headers = |pb: &mut super::PageBuilder, ri: usize, emb: &mut f32| {
        pb.begin_next_page(sect_idx, sp, emb, ctx);
        if header_count > 0 && ri >= header_count {
            render_header_rows(
                table,
                &row_layouts,
                &col_widths,
                table_left,
                pb,
                ctx,
                &merge_spans,
                header_count,
            );
        }
    };

    let mut did_flush_while_floating = false;
    // Set to the fresh page's body top when a floating table is pushed whole
    // onto it. The table's anchor paragraph flows from here even though the
    // cursor is left below the table, so its paragraph-relative shapes anchor
    // correctly. `None` while the table stays on its original page.
    let mut flushed_page_top: Option<f32> = None;
    // tblpY to apply to the render start of a vertAnchor="text" table (0 for
    // page/margin anchors). Raised before rendering, restored after, so the
    // table sits where Word puts it without shifting the following flow.
    let text_anchor_offset = override_pos
        .as_ref()
        .filter(|fp| fp.v_anchor_text)
        .map_or(0.0, |fp| fp.v_offset_pt);

    // Split an oversized row across pages, rendering partial rows and
    // flushing pages between chunks until every cell is fully emitted.
    let split_row_across_pages = |row: &TableRow,
                                  layout: &RowLayout,
                                  pb: &mut super::PageBuilder,
                                  ri: usize,
                                  did_flush: &mut bool,
                                  emb: &mut f32| {
        let ncells = layout.cells.len();
        let mut starts = vec![CellCursor::default(); ncells];
        loop {
            let avail = pb.slot_top - *emb;
            let mut ends = Vec::with_capacity(ncells);
            let mut all_done = true;

            for ci in 0..ncells {
                let end = find_cell_split(&layout.cells[ci], &starts[ci], avail);
                if end.item < layout.cells[ci].items.len() {
                    all_done = false;
                }
                ends.push(end);
            }

            let page = pb.all_contents.len();
            render_partial_row(
                row,
                layout,
                &col_widths,
                table_left,
                &mut pb.content,
                &mut pb.slot_top,
                ctx,
                &mut pb.gradient_specs,
                &mut pb.links,
                pb.table_tags.as_mut().map(|t| (&mut pb.tags, t, page)),
                &starts,
                &ends,
                ri,
            );

            if all_done {
                break;
            }

            starts = ends;
            if is_floating {
                *did_flush = true;
            }
            flush_and_render_headers(pb, ri, emb);
        }
    };

    // A floating table whose first row doesn't fit in the space left on the
    // current page is pushed whole to the next page rather than overlapping the
    // bottom content. vertAnchor="text" tables flow with their anchor text, so
    // Word starts them on the next page; the per-row split path below can't
    // rescue this when the first row's cells aren't splittable (e.g. headers
    // that are OMML math). Re-anchor to the top of the fresh page. Setting
    // did_flush_while_floating leaves the cursor on the new page so the
    // following content flows below the table instead of behind it.
    // (Pendulum #172/#173: the data table was landing on page 1 over list
    // item 10, which also pushed the following illustration off-page.)
    if keep_with_anchor && !row_layouts.is_empty() {
        let eff_top = effective_slot_top(sp, pb.is_first_page_of_section, pb.page_count(), ctx);
        let at_page_top = (pb.slot_top - eff_top).abs() < 1.0;
        let available = pb.slot_top - *effective_margin_bottom;
        let total_h: f32 = row_layouts.iter().map(|r| r.height).sum();
        let page_content_h = eff_top - *effective_margin_bottom;
        // Keep the floating table together: when it doesn't fit in the space
        // left here but would fit on a fresh page, move the whole table to the
        // next page (re-anchoring to the top). Using the full height (not just
        // the first row) is robust to short header rows. Tables taller than a
        // full page fall through to the per-row split path below.
        if !at_page_top && total_h > available && total_h <= page_content_h {
            pb.begin_next_page(sect_idx, sp, effective_margin_bottom, ctx);
            did_flush_while_floating = true;
            // The body top is where the anchor paragraph flows (and where its
            // paragraph-relative shapes anchor). The table itself, however,
            // honors tblpY relative to that anchor — for vertAnchor="text" a
            // negative tblpY lifts it above the top margin. Render the table from
            // the offset position; the flow cursor is restored to the un-offset
            // bottom afterwards so following content is unaffected.
            flushed_page_top = Some(pb.slot_top);
            pb.slot_top -= text_anchor_offset;
        }
    }

    for (ri, (row, layout)) in table.rows.iter().zip(row_layouts.iter()).enumerate() {
        // Pre-compute extra footnote space this row introduces, so the
        // page-break check below accounts for it before rendering.
        let mut row_fn_extra = 0.0f32;
        for &fn_id in &row_footnote_ids[ri] {
            if !pb.footnote_ids_set.contains(&fn_id) {
                row_fn_extra +=
                    super::footnotes::footnote_height(fn_id, footnotes, ctx, fn_text_width);
            }
        }
        if row_fn_extra > 0.0 && pb.col_fn_reserved == 0.0 {
            row_fn_extra += ctx.note_separator.height;
        }

        let row_h = layout.height;
        log::debug!(
            "TABLE row={} row_h={:.2} cells={} slot_top={:.2}",
            ri,
            row_h,
            layout.cells.len(),
            pb.slot_top
        );
        let eff_top = effective_slot_top(sp, pb.is_first_page_of_section, pb.page_count(), ctx);
        let eff_bottom = *effective_margin_bottom + row_fn_extra;
        let at_page_top = (pb.slot_top - eff_top).abs() < 1.0;
        let available_h = pb.slot_top - eff_bottom;
        let page_content_h = eff_top - eff_bottom;
        // A row whose first cell holds a keepNext paragraph stays on the page
        // where the next row starts, and such rows chain: italian_academic's
        // Titolo2 rows move to page 2 together although the first would fit.
        // Rows with keepNext only in later cells split as usual (a 199-page
        // corpus document's syllabus tables). ponytail: first-cell rule fits
        // the three measured documents; probe if one disagrees.
        let keeps_next = |r: &TableRow| {
            r.cells.first().is_some_and(|c| {
                c.content
                    .iter()
                    .any(|b| matches!(b, Block::Paragraph(p) if p.keep_next))
            })
        };
        let chain_moves = !at_page_top && keeps_next(row) && {
            let mut k = ri;
            let mut h = 0.0;
            while k < table.rows.len() && keeps_next(&table.rows[k]) {
                h += row_layouts[k].height;
                k += 1;
            }
            // ponytail: a splittable next row's start = a 14pt line, as the
            // split guard below; measure its first chunk if one misjudges.
            if let Some(next) = row_layouts.get(k) {
                h += next.split_min.map_or(next.height, |m| m.max(14.0));
            }
            h > available_h && h <= page_content_h
        };

        // Word splits any non-cantSplit row that overflows the page remainder,
        // filling the current page before continuing on the next — there is no
        // minimum row height for splitting. But a row with an explicit
        // trHeight (exact or atLeast) never breaks in Word; it migrates whole
        // (arizona_physical / traditional_skills vs isla / master_thesis).
        // A cell must have somewhere to break — several items, or a paragraph
        // long enough to leave two lines on each side — for the split to
        // produce anything, and every cell's first chunk must genuinely fit
        // in the remaining space (otherwise find_cell_split's force-included
        // first item would overflow the footer and the row migrates whole
        // instead).
        let any_cell_multi_item = layout.cells.iter().any(|c| {
            c.items.len() > 1
                || c.items.iter().any(|it| match it {
                    CellContentItem::Paragraph(p) => p.lines.len() >= 4,
                    CellContentItem::NestedTable { rows, .. } => {
                        rows.len() >= 2 || rows.iter().any(|r| r.split_min.is_some())
                    }
                })
        });
        // Whether every cell's first chunk fits in the room left: its first
        // paragraph (two lines of a long one) or first nested row, with the
        // opening space before that find_cell_split charges.
        let first_chunks_fit = layout.cells.iter().all(|c| {
            c.items.first().is_none_or(|it| {
                let end = match it {
                    CellContentItem::Paragraph(p) if p.lines.len() >= 4 => Some(2),
                    CellContentItem::Paragraph(_) => None,
                    CellContentItem::NestedTable { .. } => Some(1),
                };
                c.cm.top
                    + c.cm.bottom
                    + chunk_space_before(it, 0, &CELL_START)
                    + item_chunk_height(
                        it,
                        &Chunk {
                            item: 0,
                            l0: 0,
                            l1: end,
                            from: &[],
                            to: &[],
                        },
                    )
                    <= available_h
            })
        });
        // Word splits with one line of room: nabl's "Remarks" row breaks
        // between its paragraphs with 30pt left. ponytail: 14pt (a line) guard
        // so a near-boundary rounding error can't split off nothing; drop it
        // if a reference ever splits with less.
        // A long first paragraph needs only its first two lines in the room
        // left, as split rows break between lines (bulgarian_road's row ends
        // page 2 with three lines of a six-line cell paragraph).
        let can_meaningfully_split = layout.split_min.is_some_and(|m| available_h >= m)
            && any_cell_multi_item
            && !at_page_top
            && available_h > 14.0
            && first_chunks_fit;

        // A row taller than a page must split, but not where its cells can't
        // start: away from the page top each cell needs its first paragraph,
        // or two lines of a long one (widow control), in the room left —
        // croatian_grant's 776pt row starts on the next page rather than in
        // the 28pt above a footnote.
        let must_split = (row_h > page_content_h || keep_with_anchor)
            && !row.cant_split
            && (at_page_top || first_chunks_fit);
        if row_h > available_h && (must_split || can_meaningfully_split) && !chain_moves {
            split_row_across_pages(
                row,
                layout,
                pb,
                ri,
                &mut did_flush_while_floating,
                effective_margin_bottom,
            );
        } else if !at_page_top && (row_h > available_h || chain_moves) {
            if is_floating {
                did_flush_while_floating = true;
            }
            // Word closes a merged cell that runs on past the page break with
            // its bottom border (nabl page 10's "26." column).
            if ri > 0 {
                for (grid_col, span, cell) in table.rows[ri - 1].grid_cells() {
                    if merge_spans
                        .get(&(ri - 1, grid_col))
                        .is_some_and(|&below| below > 0.0)
                    {
                        let bx = cell_x_offset(&col_widths, table_left, grid_col);
                        let right = bx + cell_span_width(&col_widths, grid_col, span);
                        let y = pb.slot_top;
                        draw_border(&mut pb.content, &cell.borders.bottom, bx, y, right, y);
                    }
                }
            }
            flush_and_render_headers(pb, ri, effective_margin_bottom);
            // After flushing + rendering header rows, re-check if the row
            // fits on the fresh page.  If repeating headers consumed enough
            // space that the row no longer fits, split it across pages
            // instead of rendering blindly (which would overflow the footer).
            let new_eff_top =
                effective_slot_top(sp, pb.is_first_page_of_section, pb.page_count(), ctx);
            let new_eff_bot = *effective_margin_bottom;
            let new_available = pb.slot_top - new_eff_bot;
            let new_page_h = new_eff_top - new_eff_bot;
            if row_h > new_available && (row_h > new_page_h || keep_with_anchor) && !row.cant_split
            {
                split_row_across_pages(
                    row,
                    layout,
                    pb,
                    ri,
                    &mut did_flush_while_floating,
                    effective_margin_bottom,
                );
            } else {
                render_table_row(
                    row,
                    layout,
                    &col_widths,
                    table_left,
                    pb,
                    ctx,
                    ri,
                    &merge_spans,
                );
            }
        } else {
            render_table_row(
                row,
                layout,
                &col_widths,
                table_left,
                pb,
                ctx,
                ri,
                &merge_spans,
            );
        }

        // Register footnotes from this row and reserve space for them.
        for &fn_id in &row_footnote_ids[ri] {
            let fn_h = footnotes
                .contains_key(&fn_id)
                .then(|| super::footnotes::footnote_height(fn_id, footnotes, ctx, fn_text_width));
            pb.book_footnote(
                fn_id,
                fn_col,
                fn_h,
                ctx.note_separator.height,
                effective_margin_bottom,
            );
        }

        // Register endnotes from this row; they render at end of document
        // (last page), so just collect IDs in encounter order.
        for &en_id in &row_endnote_ids[ri] {
            if pb.endnote_ids_set.insert(en_id) {
                pb.endnote_ids.push(en_id);
            }
        }

        if pb.slot_top < *effective_margin_bottom - 1.0 {
            log::warn!(
                "Table overflow: row={} slot_top={:.2} < eff_margin_bottom={:.2} row_h={:.2} page={}",
                ri,
                pb.slot_top,
                *effective_margin_bottom,
                row_h,
                pb.all_contents.len(),
            );
        }
    }
    pb.slot_top -= outer_band(table.rows.last(), false) / 2.0;

    if let Some((saved, table_top_y)) = saved_slot_top {
        if did_flush_while_floating {
            // Table spanned multiple pages — cursor is already on the new page,
            // don't restore to the pre-table position or register a float zone.
            // The cursor stays below the table so following text flows below it
            // (not behind it). But the table's anchor paragraph — the next block,
            // which carries vertAnchor="text" shapes — must anchor its
            // paragraph-relative shapes from the top of the page body where the
            // anchor naturally flows, not from the cursor below the table. Hand
            // that top down as a one-shot override (pendulum graph + pendulum
            // illustrations were dropping by the table's height). Only the
            // whole-table-to-next-page path sets flushed_page_top; the per-row
            // split path (taller-than-page tables) leaves it None, so the
            // override is skipped there rather than anchoring shapes off-page.
            if let Some(top) = flushed_page_top {
                pb.pending_float_anchor = Some(top);
                // Undo the tblpY offset applied to the render start so the flow
                // cursor sits at the un-offset table bottom — following content
                // (which is correctly placed) must not shift with the table.
                pb.slot_top += text_anchor_offset;
            }
        } else {
            let table_total_w: f32 = col_widths.iter().sum();
            let (top_margin, bottom_margin) = text_margins;
            let table_bottom = pb.slot_top;

            // Always restore cursor to body text position — the float zone lets
            // paragraph layout decide whether to wrap beside or push below.
            pb.slot_top = saved;

            if table_bottom < saved {
                let fp = override_pos.as_ref().unwrap();
                // Use BothSides wrapping when the table has >=72pt of space on each side
                let text_area_left = sp.margin_left;
                let text_area_right = sp.page_width - sp.margin_right;
                let space_left = table_left - text_area_left;
                let space_right = text_area_right - (table_left + table_total_w);
                let wrap_text = if space_left >= 72.0 && space_right >= 72.0 {
                    crate::model::WrapText::BothSides
                } else {
                    crate::model::WrapText::Largest
                };
                pb.float_zone = Some(super::FloatZone {
                    top_y: table_top_y + top_margin,
                    bottom_y: table_bottom - bottom_margin,
                    obj_left: table_left,
                    obj_right: table_left + table_total_w,
                    left_from_text: fp.left_from_text,
                    right_from_text: fp.right_from_text,
                    polygon_pts: None,
                    wrap_text,
                    para_relative: false,
                    from_table: true,
                });
                if let Some(old) = preceding_float_zone
                    && let Some(ref mut current) = pb.float_zone
                    && (old.obj_left - current.obj_left).abs() < 0.5
                    && (old.obj_right - current.obj_right).abs() < 0.5
                {
                    current.top_y = current.top_y.max(old.top_y);
                    current.bottom_y = current.bottom_y.min(old.bottom_y);
                    current.left_from_text = current.left_from_text.max(old.left_from_text);
                    current.right_from_text = current.right_from_text.max(old.right_from_text);
                }
            }
        }
    }
}

pub(super) fn compute_hf_table_height(table: &Table, ctx: &RenderContext, content_w: f32) -> f32 {
    let mut col_widths = auto_fit_columns(table, ctx.fonts, None, None);
    apply_pct_width(table, &mut col_widths, content_w);
    let row_layouts = compute_row_layouts(table, &col_widths, ctx, None);
    row_layouts.iter().map(|r| r.height).sum()
}

pub(super) fn render_header_footer_table(
    table: &Table,
    sp: &SectionProperties,
    ctx: &RenderContext,
    content: &mut Content,
    cursor_y: &mut f32,
    page_num: usize,
    total_pages: usize,
    styleref_values: &HashMap<String, String>,
    page_num_format: Option<&str>,
    gradient_specs: &mut Vec<super::GradientSpec>,
    links: &mut Vec<LinkAnnotation>,
) {
    let mut col_widths = auto_fit_columns(table, ctx.fonts, None, None);
    apply_pct_width(
        table,
        &mut col_widths,
        sp.page_width - sp.margin_left - sp.margin_right,
    );
    let hf_sub = HfSubstitution {
        page_num,
        total_pages,
        styleref_values,
        page_num_format,
    };
    let row_layouts = compute_row_layouts(table, &col_widths, ctx, Some(&hf_sub));
    // A floating table sits at its own position and leaves the header's flow
    // alone (french_sexual: the logo paragraph after it starts at the header top).
    let mut float_y: f32;
    let (table_left, y) = if let Some(pos) = &table.position {
        let fp = super::FloatingTablePos::resolve(
            table,
            pos,
            sp,
            sp.margin_left,
            sp.text_width(),
            *cursor_y,
            ctx,
        );
        float_y = fp.y;
        (fp.x, &mut float_y)
    } else {
        let total_w: f32 = col_widths.iter().sum();
        let area = (sp.margin_left, sp.text_width());
        (
            aligned_table_left(table, area, total_w, ctx.compat_mode),
            cursor_y,
        )
    };

    let merge_spans = compute_merge_spans(table, &row_layouts);

    render_table_rows(
        table,
        &row_layouts,
        &col_widths,
        table_left,
        &merge_spans,
        content,
        y,
        ctx,
        gradient_specs,
        links,
        None,
        0..usize::MAX,
    );
}

#[cfg(test)]
mod floating_nested_visibility_tests {
    use super::*;

    #[test]
    fn zero_flow_height_does_not_hide_nested_fields() {
        let items = vec![CellContentItem::NestedTable {
            col_widths: vec![80.0],
            rows: vec![RowLayout {
                height: 12.0,
                cells: vec![],
                split_min: None,
            }],
            space_before: 0.0,
            floating_offset: Some(0.0),
            floating_x_offset: None,
        }];
        assert_eq!(items[0].height(), 0.0);
        assert!(cell_has_visible_content(&items));
    }
}
