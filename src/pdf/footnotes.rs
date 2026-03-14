use std::collections::HashMap;

use pdf_writer::Content;

use crate::model::{Footnote, LineSpacing, Run};

use super::RenderContext;
use super::layout::{
    TextLine, build_paragraph_lines, is_text_empty, render_paragraph_lines, tallest_run_metrics,
};
use super::resolve_line_h;

fn substitute_ref_marks(runs: &[Run], display_num: u32) -> Vec<Run> {
    runs.iter()
        .map(|run| {
            if run.is_footnote_ref_mark {
                let mut r = run.clone();
                r.text = display_num.to_string();
                r
            } else {
                run.clone()
            }
        })
        .collect()
}

struct ParagraphLayout {
    font_size: f32,
    line_height: f32,
    ascender_ratio: f32,
    lines: Vec<TextLine>,
}

fn layout_paragraph(
    runs: &[Run],
    line_spacing: LineSpacing,
    ctx: &RenderContext,
    text_width: f32,
) -> Option<ParagraphLayout> {
    if is_text_empty(runs) {
        return None;
    }
    let (fs, tallest_lhr, tallest_ar) = tallest_run_metrics(runs, ctx.fonts);
    let lh = resolve_line_h(line_spacing, fs, tallest_lhr);
    let lines = build_paragraph_lines(runs, ctx.fonts, text_width, 0.0, &HashMap::new());
    if lines.is_empty() {
        return None;
    }
    Some(ParagraphLayout {
        font_size: fs,
        line_height: lh,
        ascender_ratio: tallest_ar.unwrap_or(0.75),
        lines,
    })
}

pub(super) fn compute_footnote_height(
    footnote: &Footnote,
    ctx: &RenderContext,
    text_width: f32,
) -> f32 {
    footnote
        .paragraphs
        .iter()
        .filter_map(|para| {
            let ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
            layout_paragraph(&para.runs, ls, ctx, text_width)
        })
        .map(|layout| layout.lines.len().max(1) as f32 * layout.line_height)
        .sum()
}

pub(super) fn render_page_footnotes(
    content: &mut Content,
    fn_ids: &[u32],
    footnotes: &HashMap<u32, Footnote>,
    footnote_display_order: &HashMap<u32, u32>,
    ctx: &RenderContext,
    margin_left: f32,
    margin_bottom: f32,
    text_width: f32,
) {
    if fn_ids.is_empty() {
        return;
    }

    let total_fn_height: f32 = fn_ids
        .iter()
        .filter_map(|id| footnotes.get(id))
        .map(|fn_note| compute_footnote_height(fn_note, ctx, text_width))
        .sum();

    let separator_gap = 12.0f32;
    let block_top = margin_bottom + total_fn_height + separator_gap;

    // Draw separator line: 0.5pt black, ~1/3 page width
    let sep_y = block_top - 3.0;
    let sep_width = 144.0f32.min(text_width);
    content.save_state();
    content.set_line_width(0.5);
    content.move_to(margin_left, sep_y);
    content.line_to(margin_left + sep_width, sep_y);
    content.stroke();
    content.restore_state();

    let mut fn_y = sep_y - 9.0;

    for fn_id in fn_ids {
        let Some(footnote) = footnotes.get(fn_id) else {
            continue;
        };
        let display_num = footnote_display_order.get(fn_id).copied().unwrap_or(0);

        for para in &footnote.paragraphs {
            let runs = substitute_ref_marks(&para.runs, display_num);
            let ls = para.line_spacing.unwrap_or(LineSpacing::Auto(1.0));

            let Some(layout) = layout_paragraph(&runs, ls, ctx, text_width) else {
                continue;
            };

            let baseline_y = fn_y - layout.font_size * layout.ascender_ratio;
            let line_count = layout.lines.len();

            render_paragraph_lines(
                content,
                &layout.lines,
                &para.alignment,
                margin_left,
                text_width,
                baseline_y,
                layout.line_height,
                line_count,
                0,
                &mut Vec::new(),
                0.0,
                ctx.fonts,
            );

            fn_y -= line_count as f32 * layout.line_height;
        }
    }
}
