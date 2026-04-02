use std::collections::HashMap;

use pdf_writer::{Content, Name, Str};

use crate::fonts::{FontEntry, encode_as_gids, to_winansi_bytes};
use crate::model::{
    Alignment, Block, BorderStyle, CellBorder, CellBorders, CellMargins, CellVAlign,
    SectionProperties, Table, TableAlignment, TableRow, TextDirection, VMerge,
};

use super::color::{fill_rgb, stroke_rgb};
use super::header_footer::{compute_effective_margin_bottom, effective_slot_top};

use super::RenderContext;
use super::layout::{
    encode_text_for_pdf, render_paragraph_lines,
};
use super::table_layout::{
    CellContentItem, CellLayout, CellParagraphLayout, HfSubstitution,
    RowLayout, auto_fit_columns, cell_span_width, cell_x_offset, compute_merge_spans,
    compute_row_layouts, find_cell_split, para_block_height,
};

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
            content.set_dash_pattern(
                [w * 4.0, w * 2.0, 0.0, w * 2.0, 0.0, w * 2.0],
                0.0,
            );
        }
        BorderStyle::Double => {
            // Word renders each line of a double border at the full
            // specified width, separated by a gap equal to the width.
            let thin = w.max(0.25);
            let gap = thin;
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

fn draw_cell_shading(
    content: &mut Content,
    shading: [u8; 3],
    borders: &crate::model::CellBorders,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
) {
    let bw = |b: &crate::model::CellBorder| if b.present { b.width } else { 0.0 };
    let inset = (bw(&borders.top) + bw(&borders.bottom) + bw(&borders.left) + bw(&borders.right)) / 8.0;
    content.save_state();
    fill_rgb(content, shading);
    content.rect(x + inset, y + inset, w - 2.0 * inset, h - 2.0 * inset);
    content.fill_nonzero();
    content.restore_state();
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
}

/// Total content height including trailing space_after, matching Word's vAlign calculation.
fn cell_content_h_for_valign(items: &[CellContentItem]) -> f32 {
    let mut h: f32 = items
        .iter()
        .map(|item| match item {
            CellContentItem::Paragraph(p) => p.space_before + para_block_height(p),
            CellContentItem::NestedTable { height } => *height,
        })
        .sum();
    // Word includes the last paragraph's space_after in the content block height
    // used for vertical alignment, so bottom/center-aligned cells position correctly.
    if let Some(CellContentItem::Paragraph(last_para)) = items.last() {
        h += last_para.space_after;
    }
    h
}

fn cell_has_visible_content(items: &[CellContentItem]) -> bool {
    items.iter().any(|item| match item {
        CellContentItem::Paragraph(p) => para_has_visible_content(p),
        CellContentItem::NestedTable { height } => *height > 0.0,
    })
}

fn render_cell_content(
    content: &mut Content,
    items: &[CellContentItem],
    blocks: &[Block],
    cell_x: f32,
    col_w: f32,
    cursor_y_start: f32,
    cm: &CellMargins,
    ctx: &RenderContext,
) {
    let mut cursor_y = cursor_y_start;
    let mut block_idx = 0;

    for item in items {
        match item {
            CellContentItem::Paragraph(para) => {
                // Advance block_idx past the corresponding Block::Paragraph
                while block_idx < blocks.len() {
                    if matches!(&blocks[block_idx], Block::Paragraph(_)) {
                        block_idx += 1;
                        break;
                    }
                    block_idx += 1;
                }

                let has_floats = !para.floating_images.is_empty();

                if !para_has_visible_content(para)
                    && para.image_name.is_none()
                    && !has_floats
                {
                    cursor_y -= para.space_before + para_block_height(para);
                    continue;
                }

                cursor_y -= para.space_before;

                // Render floating images positioned relative to this paragraph
                for fi in &para.floating_images {
                    let fi_x = cell_x + fi.h_offset;
                    let fi_y_top = cursor_y - fi.v_offset;
                    let fi_y_bottom = fi_y_top - fi.display_height;
                    content.save_state();
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

                if let Some(ref img_name) = para.image_name {
                    let img_x = cell_x + cm.left;
                    let img_y = cursor_y - para.image_height;

                    if let Some(ref shadow) = para.image_shadow {
                        super::color::draw_image_shadow(
                            content, shadow, img_x, img_y,
                            para.image_width, para.image_height, None,
                        );
                    }

                    content.save_state();
                    content.transform([
                        para.image_width,
                        0.0,
                        0.0,
                        para.image_height,
                        img_x,
                        img_y,
                    ]);
                    content.x_object(Name(img_name.as_bytes()));
                    content.restore_state();
                    if let Some(sc) = para.image_stroke_color {
                        content.save_state();
                        stroke_rgb(content, sc);
                        content.set_line_width(para.image_stroke_width);
                        content.rect(img_x, img_y, para.image_width, para.image_height);
                        content.stroke();
                        content.restore_state();
                    }
                    // Advance cursor by display height only; distT/distB in
                    // layout_extra_height contribute to row height calculation
                    // but don't add spacing between image and following text.
                    cursor_y -= para.image_height;
                    continue;
                }

                let text_x = cell_x + cm.left + para.indent_left + para.float_indent_left;
                let text_w =
                    (col_w - cm.left - cm.right - para.indent_left - para.float_indent_left)
                        .max(0.0);
                // Word positions first baseline at cell_top - font_size (full em)
                let baseline_y = cursor_y - para.font_size;

                let first_line_hanging = if para.list_label.is_empty() {
                    para.indent_hanging
                } else {
                    let label_x = cell_x + cm.left + para.indent_left - para.indent_hanging;
                    draw_cell_label(content, para, label_x, baseline_y, ctx.fonts);
                    if para.indent_first_line > 0.0 && para.indent_hanging == 0.0 {
                        -para.indent_first_line
                    } else {
                        0.0
                    }
                };

                render_paragraph_lines(
                    content,
                    &para.lines,
                    &para.alignment,
                    text_x,
                    text_w,
                    baseline_y,
                    para.line_h,
                    para.lines.len(),
                    0,
                    &mut Vec::new(),
                    first_line_hanging,
                    ctx.fonts,
                    None,
                );

                cursor_y -= para.lines.len() as f32 * para.line_h;
            }
            CellContentItem::NestedTable { height } => {
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
                    render_nested_table(table, content, cell_x + cm.left, col_w - cm.left - cm.right, &mut cursor_y, ctx);
                } else {
                    cursor_y -= height;
                }
            }
        }
    }
}

/// Render a nested table inline within a parent cell at the given cursor position.
fn render_nested_table(
    table: &Table,
    content: &mut Content,
    available_x: f32,
    available_w: f32,
    cursor_y: &mut f32,
    ctx: &RenderContext,
) {
    let col_widths = auto_fit_columns(table, ctx.fonts, Some(available_w));
    let row_layouts = compute_row_layouts(table, &col_widths, ctx, None);
    let cm = &table.cell_margins;
    let table_total_w: f32 = col_widths.iter().sum();
    let table_left = match table.alignment {
        TableAlignment::Center => available_x + (available_w - table_total_w) / 2.0,
        TableAlignment::Right => available_x + available_w - table_total_w,
        TableAlignment::Left => available_x + table.table_indent - cm.left,
    };

    let merge_spans = compute_merge_spans(table, &row_layouts);

    for (ri, (row, layout)) in table.rows.iter().zip(row_layouts.iter()).enumerate() {
        let row_h = layout.height;
        let row_top = *cursor_y;
        let row_bottom = row_top - row_h;

        let mut grid_col = 0usize;
        for (cell, cell_layout) in row.cells.iter().zip(layout.cells.iter()) {
            let span = cell.grid_span.max(1) as usize;
            let col_w = cell_span_width(&col_widths, grid_col, span);
            let cx = cell_x_offset(&col_widths, table_left, grid_col);
            let cell_grid_col = grid_col;
            grid_col += span;

            if cell.v_merge == VMerge::Continue {
                continue;
            }

            let merge_extra = merge_spans
                .get(&(ri, cell_grid_col))
                .copied()
                .unwrap_or(0.0);
            let effective_h = row_h + merge_extra;

            if let Some(shading) = cell.shading {
                content.save_state();
                fill_rgb(content, shading);
                content.rect(cx, row_bottom, col_w, row_h);
                content.fill_nonzero();
                content.restore_state();
            }

            if cell_has_visible_content(&cell_layout.items) {
                let ecm = cell.cell_margins.as_ref().unwrap_or(cm);
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
                    ecm,
                    ctx,
                );
            }
        }

        let mut grid_col = 0usize;
        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            let col_w = cell_span_width(&col_widths, grid_col, span);
            let bx = cell_x_offset(&col_widths, table_left, grid_col);
            let cell_grid_col = grid_col;
            grid_col += span;

            if cell.v_merge == VMerge::Continue {
                continue;
            }

            let merge_extra = merge_spans
                .get(&(ri, cell_grid_col))
                .copied()
                .unwrap_or(0.0);
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

fn render_partial_cell_content(
    content: &mut Content,
    items: &[CellContentItem],
    blocks: &[Block],
    start: usize,
    end: usize,
    cell_x: f32,
    col_w: f32,
    cursor_y_start: f32,
    cm: &CellMargins,
    ctx: &RenderContext,
) {
    let mut cursor_y = cursor_y_start;
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

    for pi in start..end {
        match &items[pi] {
            CellContentItem::Paragraph(para) => {
                let sb = if pi == start { 0.0 } else { para.space_before };

                if !para_has_visible_content(para) && para.floating_images.is_empty() {
                    cursor_y -= sb + para_block_height(para);
                    continue;
                }

                cursor_y -= sb;

                for fi in &para.floating_images {
                    let fi_x = cell_x + fi.h_offset;
                    let fi_y_top = cursor_y - fi.v_offset;
                    let fi_y_bottom = fi_y_top - fi.display_height;
                    content.save_state();
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

                let text_x = cell_x + cm.left + para.indent_left + para.float_indent_left;
                let text_w =
                    (col_w - cm.left - cm.right - para.indent_left - para.float_indent_left)
                        .max(0.0);
                // Word positions first baseline at cell_top - font_size (full em)
                let baseline_y = cursor_y - para.font_size;

                let first_line_hanging = if para.list_label.is_empty() {
                    para.indent_hanging
                } else {
                    let label_x = cell_x + cm.left + para.indent_left - para.indent_hanging;
                    draw_cell_label(content, para, label_x, baseline_y, ctx.fonts);
                    if para.indent_first_line > 0.0 && para.indent_hanging == 0.0 {
                        -para.indent_first_line
                    } else {
                        0.0
                    }
                };

                render_paragraph_lines(
                    content,
                    &para.lines,
                    &para.alignment,
                    text_x,
                    text_w,
                    baseline_y,
                    para.line_h,
                    para.lines.len(),
                    0,
                    &mut Vec::new(),
                    first_line_hanging,
                    ctx.fonts,
                    None,
                );

                cursor_y -= para.lines.len() as f32 * para.line_h;
            }
            CellContentItem::NestedTable { height } => {
                let bi = item_to_block.get(pi).copied().unwrap_or(0);
                if let Some(Block::Table(table)) = blocks.get(bi) {
                    render_nested_table(
                        table, content, cell_x + cm.left, col_w - cm.left - cm.right,
                        &mut cursor_y, ctx,
                    );
                } else {
                    cursor_y -= height;
                }
            }
        }
    }
}

fn encode_label(label: &str, entry: &FontEntry) -> Vec<u8> {
    match &entry.char_to_gid {
        Some(map) => encode_as_gids(label, map),
        None => to_winansi_bytes(label),
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
    let bytes = encode_label(&para.list_label, entry);

    if let Some(c) = para.label_color {
        fill_rgb(content, c);
    }
    content
        .begin_text()
        .set_font(Name(entry.pdf_name.as_bytes()), para.font_size)
        .next_line(label_x, baseline_y)
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
    cm: &CellMargins,
    table_left: f32,
    pb: &mut super::PageBuilder,
    ctx: &RenderContext,
    row_idx: usize,
    merge_spans: &HashMap<(usize, usize), f32>,
) {
    let row_h = layout.height;
    let row_top = pb.slot_top;
    let row_bottom = row_top - row_h;

    let mut grid_col = 0usize;
    for (cell, cell_layout) in row.cells.iter().zip(layout.cells.iter()) {
        let span = cell.grid_span.max(1) as usize;
        let col_w = cell_span_width(col_widths, grid_col, span);
        let cell_x = cell_x_offset(col_widths, table_left, grid_col);
        let cell_grid_col = grid_col;
        grid_col += span;

        if cell.v_merge == VMerge::Continue {
            continue;
        }

        let merge_extra = merge_spans
            .get(&(row_idx, cell_grid_col))
            .copied()
            .unwrap_or(0.0);
        let effective_h = row_h + merge_extra;

        if let Some(shading) = cell.shading {
            draw_cell_shading(
                &mut pb.content,
                shading,
                &cell.borders,
                cell_x,
                row_bottom,
                col_w,
                row_h,
            );
        }

        let has_content = cell_has_visible_content(&cell_layout.items);
        let ecm = cell.cell_margins.as_ref().unwrap_or(cm);

        if has_content && cell_layout.text_direction == TextDirection::TbRl {
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
            );
        } else if has_content {
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
                ecm,
                ctx,
            );
        }
    }

    let mut grid_col = 0usize;
    for cell in &row.cells {
        let span = cell.grid_span.max(1) as usize;
        let col_w = cell_span_width(col_widths, grid_col, span);
        let bx = cell_x_offset(col_widths, table_left, grid_col);

        if cell.v_merge == VMerge::Continue {
            grid_col += span;
            continue;
        }

        let merge_extra = merge_spans
            .get(&(row_idx, grid_col))
            .copied()
            .unwrap_or(0.0);
        let effective_bottom = row_bottom - merge_extra;

        draw_cell_borders(
            &mut pb.content,
            &cell.borders,
            bx,
            row_top,
            effective_bottom,
            col_w,
            true,
            true,
        );

        grid_col += span;
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

    for para in &paras {
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

                content.begin_text();
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
                content.end_text();

                if chunk.color.is_some() {
                    content.set_fill_gray(0.0);
                }
            }
        }
    }
}

/// Render a subset of each cell's paragraphs for a split row.
/// `starts[ci]..ends[ci]` gives the paragraph range for cell `ci`.
/// `is_first`/`is_last` control top/bottom border drawing.
fn render_partial_row(
    row: &TableRow,
    layout: &RowLayout,
    col_widths: &[f32],
    cm: &CellMargins,
    table_left: f32,
    pb: &mut super::PageBuilder,
    ctx: &RenderContext,
    starts: &[usize],
    ends: &[usize],
    is_first: bool,
    is_last: bool,
) {
    let mut max_h: f32 = cm.top + cm.bottom;
    for (ci, cell_layout) in layout.cells.iter().enumerate() {
        let start = starts[ci];
        let end = ends[ci];
        let mut h = cm.top + cm.bottom;
        for pi in start..end {
            let item_h = match &cell_layout.items[pi] {
                CellContentItem::Paragraph(para) => {
                    let sb = if pi == start { 0.0 } else { para.space_before };
                    sb + para_block_height(para)
                }
                CellContentItem::NestedTable { height } => *height,
            };
            h += item_h;
        }
        max_h = max_h.max(h);
    }

    let row_h = max_h;
    let row_top = pb.slot_top;
    let row_bottom = row_top - row_h;

    let mut grid_col = 0usize;
    for (ci, (cell, cell_layout)) in row.cells.iter().zip(layout.cells.iter()).enumerate() {
        let span = cell.grid_span.max(1) as usize;
        let col_w = cell_span_width(col_widths, grid_col, span);
        let cell_x = cell_x_offset(col_widths, table_left, grid_col);
        grid_col += span;

        if cell.v_merge == VMerge::Continue {
            continue;
        }

        let start = starts[ci];
        let end = ends[ci];

        if let Some(shading) = cell.shading {
            draw_cell_shading(
                &mut pb.content,
                shading,
                &cell.borders,
                cell_x,
                row_bottom,
                col_w,
                row_h,
            );
        }

        let has_content = (start..end).any(|pi| match &cell_layout.items[pi] {
            CellContentItem::Paragraph(p) => para_has_visible_content(p),
            CellContentItem::NestedTable { height } => *height > 0.0,
        });

        if has_content {
            render_partial_cell_content(
                &mut pb.content,
                &cell_layout.items,
                &cell.content,
                start,
                end,
                cell_x,
                col_w,
                row_top - cm.top,
                cm,
                ctx,
            );
        }
    }

    let mut grid_col = 0usize;
    for cell in &row.cells {
        let span = cell.grid_span.max(1) as usize;
        let col_w = cell_span_width(col_widths, grid_col, span);
        let bx = cell_x_offset(col_widths, table_left, grid_col);
        grid_col += span;

        if cell.v_merge == VMerge::Continue {
            continue;
        }

        let draw_top = is_first;
        draw_cell_borders(
            &mut pb.content,
            &cell.borders,
            bx,
            row_top,
            row_bottom,
            col_w,
            draw_top,
            is_last,
        );
    }

    pb.slot_top = row_bottom;
}

fn render_header_rows(
    table: &Table,
    row_layouts: &[RowLayout],
    col_widths: &[f32],
    cm: &CellMargins,
    table_left: f32,
    pb: &mut super::PageBuilder,
    ctx: &RenderContext,
    merge_spans: &HashMap<(usize, usize), f32>,
    header_count: usize,
) {
    for hi in 0..header_count {
        render_table_row(
            &table.rows[hi],
            &row_layouts[hi],
            col_widths,
            cm,
            table_left,
            pb,
            ctx,
            hi,
            merge_spans,
        );
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
    override_pos: Option<super::FloatingTablePos>,
) {
    let col_widths = auto_fit_columns(table, ctx.fonts, None);
    let row_layouts = compute_row_layouts(table, &col_widths, ctx, None);
    let merge_spans = compute_merge_spans(table, &row_layouts);
    let cm = &table.cell_margins;

    let is_floating = override_pos.is_some();
    let (table_left, saved_slot_top, text_margins) =
        if let Some(ref fp) = override_pos {
            let saved = Some((pb.slot_top - prev_space_after, fp.y));
            pb.slot_top = fp.y;
            (fp.x, saved, (fp.top_from_text, fp.bottom_from_text))
        } else {
        use crate::model::TableAlignment;
        let text_width = sp.page_width - sp.margin_left - sp.margin_right;
        let table_total_w: f32 = col_widths.iter().sum();
        let left = match table.alignment {
            TableAlignment::Center => sp.margin_left + (text_width - table_total_w) / 2.0,
            TableAlignment::Right => sp.margin_left + text_width - table_total_w,
            TableAlignment::Left => sp.margin_left + table.table_indent - cm.left,
        };
        (left, None, (0.0, 0.0))
    };

    // For non-floating tables, prev_space_after offsets the table start.
    // For floating tables, it was already consumed into the saved cursor position.
    pb.slot_top -= prev_space_after;

    // Count contiguous header rows from the start of the table (per OOXML spec,
    // only contiguous header rows starting from row 0 are repeated).
    let header_count = table.rows.iter().take_while(|r| r.is_header).count();

    let flush_and_render_headers = |pb: &mut super::PageBuilder, ri: usize| {
        pb.flush_page(sect_idx);
        pb.is_first_page_of_section = false;
        pb.slot_top = effective_slot_top(sp, false, ctx);
        if header_count > 0 && ri >= header_count {
            render_header_rows(
                table,
                &row_layouts,
                &col_widths,
                cm,
                table_left,
                pb,
                ctx,
                &merge_spans,
                header_count,
            );
        }
    };

    let mut did_flush_while_floating = false;

    for (ri, (row, layout)) in table.rows.iter().zip(row_layouts.iter()).enumerate() {
        let row_h = layout.height;
        log::debug!(
            "TABLE row={} row_h={:.2} cells={} slot_top={:.2}",
            ri,
            row_h,
            layout.cells.len(),
            pb.slot_top
        );
        let eff_top = effective_slot_top(sp, pb.is_first_page_of_section, ctx);
        let eff_bottom = compute_effective_margin_bottom(sp, pb.is_first_page_of_section, ctx);
        let at_page_top = (pb.slot_top - eff_top).abs() < 1.0;
        let available_h = pb.slot_top - eff_bottom;
        let page_content_h = eff_top - eff_bottom;

        if row_h > available_h && (row_h > page_content_h || is_floating) {
            // Row is too tall for any single page, or floating table overflows -- split across pages
            let ncells = layout.cells.len();
            let mut starts = vec![0usize; ncells];
            let mut is_first_chunk = true;

            loop {
                let avail = pb.slot_top - compute_effective_margin_bottom(sp, pb.is_first_page_of_section, ctx);
                let mut ends = Vec::with_capacity(ncells);
                let mut all_done = true;

                for ci in 0..ncells {
                    let end = find_cell_split(&layout.cells[ci], starts[ci], avail, cm);
                    if end < layout.cells[ci].items.len() {
                        all_done = false;
                    }
                    ends.push(end);
                }

                render_partial_row(
                    row,
                    layout,
                    &col_widths,
                    cm,
                    table_left,
                    pb,
                    ctx,
                    &starts,
                    &ends,
                    is_first_chunk,
                    all_done,
                );

                if all_done {
                    break;
                }

                starts = ends;
                is_first_chunk = false;
                if is_floating {
                    did_flush_while_floating = true;
                }
                flush_and_render_headers(pb, ri);
            }
        } else if !at_page_top && row_h > available_h {
            if is_floating {
                did_flush_while_floating = true;
            }
            flush_and_render_headers(pb, ri);
            render_table_row(
                row,
                layout,
                &col_widths,
                cm,
                table_left,
                pb,
                ctx,
                ri,
                &merge_spans,
            );
        } else {
            render_table_row(
                row,
                layout,
                &col_widths,
                cm,
                table_left,
                pb,
                ctx,
                ri,
                &merge_spans,
            );
        }
    }

    if let Some((saved, table_top_y)) = saved_slot_top {
        if did_flush_while_floating {
            // Table spanned multiple pages — cursor is already on the new page,
            // don't restore to the pre-table position or register a float zone.
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
            });
        }
        }
    }
}

pub(super) fn compute_hf_table_height(table: &Table, ctx: &RenderContext) -> f32 {
    let col_widths = auto_fit_columns(table, ctx.fonts, None);
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
) {
    let col_widths = auto_fit_columns(table, ctx.fonts, None);
    let hf_sub = HfSubstitution {
        page_num,
        total_pages,
        styleref_values,
        page_num_format: sp.page_num_format.as_deref(),
    };
    let row_layouts = compute_row_layouts(table, &col_widths, ctx, Some(&hf_sub));
    let cm = &table.cell_margins;
    let table_left = {
        use crate::model::TableAlignment;
        let text_width = sp.page_width - sp.margin_left - sp.margin_right;
        let table_total_w: f32 = col_widths.iter().sum();
        match table.alignment {
            TableAlignment::Center => sp.margin_left + (text_width - table_total_w) / 2.0,
            TableAlignment::Right => sp.margin_left + text_width - table_total_w,
            TableAlignment::Left => sp.margin_left + table.table_indent - cm.left,
        }
    };

    let merge_spans = compute_merge_spans(table, &row_layouts);

    for (ri, (row, layout)) in table.rows.iter().zip(row_layouts.iter()).enumerate() {
        let row_h = layout.height;
        let row_top = *cursor_y;
        let row_bottom = row_top - row_h;

        let mut grid_col = 0usize;
        for (cell, cell_layout) in row.cells.iter().zip(layout.cells.iter()) {
            let span = cell.grid_span.max(1) as usize;
            let col_w = cell_span_width(&col_widths, grid_col, span);
            let cell_x = cell_x_offset(&col_widths, table_left, grid_col);
            let cell_grid_col = grid_col;
            grid_col += span;

            if cell.v_merge == VMerge::Continue {
                continue;
            }

            let merge_extra = merge_spans
                .get(&(ri, cell_grid_col))
                .copied()
                .unwrap_or(0.0);
            let effective_h = row_h + merge_extra;

            if let Some(shading) = cell.shading {
                content.save_state();
                fill_rgb(content, shading);
                content.rect(cell_x, row_bottom, col_w, row_h);
                content.fill_nonzero();
                content.restore_state();
            }

            if cell_has_visible_content(&cell_layout.items) {
                let ecm = cell.cell_margins.as_ref().unwrap_or(cm);
                let content_h = cell_content_h_for_valign(&cell_layout.items);

                let avail = effective_h - ecm.top - ecm.bottom;
                let v_offset = valign_offset(cell.v_align, avail, content_h);
                let cell_cursor_y = row_top - ecm.top - v_offset;

                render_cell_content(
                    content,
                    &cell_layout.items,
                    &cell.content,
                    cell_x,
                    col_w,
                    cell_cursor_y,
                    ecm,
                    ctx,
                );
            }
        }

        let mut grid_col = 0usize;
        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            let col_w = cell_span_width(&col_widths, grid_col, span);
            let bx = cell_x_offset(&col_widths, table_left, grid_col);
            let cell_grid_col = grid_col;
            grid_col += span;

            if cell.v_merge == VMerge::Continue {
                continue;
            }

            let merge_extra = merge_spans
                .get(&(ri, cell_grid_col))
                .copied()
                .unwrap_or(0.0);
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
