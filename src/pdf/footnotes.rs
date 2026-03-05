use std::collections::HashMap;

use pdf_writer::Content;

use crate::fonts::FontEntry;
use crate::model::{Footnote, LineSpacing, Run};

use super::layout::{
    build_paragraph_lines, is_text_empty, render_paragraph_lines, tallest_run_metrics,
};
use super::resolve_line_h;

pub(super) fn compute_footnote_height(
    footnote: &Footnote,
    seen_fonts: &HashMap<String, FontEntry>,
    text_width: f32,
    doc_line_spacing: LineSpacing,
) -> f32 {
    let mut height = 0.0f32;
    for para in &footnote.paragraphs {
        if is_text_empty(&para.runs) {
            continue;
        }
        let (fs, tallest_lhr, _) = tallest_run_metrics(&para.runs, seen_fonts);
        let effective_ls = para.line_spacing.unwrap_or(doc_line_spacing);
        let lh = resolve_line_h(effective_ls, fs, tallest_lhr);
        let lines = build_paragraph_lines(&para.runs, seen_fonts, text_width, 0.0, &HashMap::new());
        height += lines.len().max(1) as f32 * lh;
    }
    height
}

pub(super) fn render_page_footnotes(
    content: &mut Content,
    fn_ids: &[u32],
    footnotes: &HashMap<u32, Footnote>,
    footnote_display_order: &HashMap<u32, u32>,
    seen_fonts: &HashMap<String, FontEntry>,
    margin_left: f32,
    margin_bottom: f32,
    text_width: f32,
    doc_line_spacing: LineSpacing,
) {
    if fn_ids.is_empty() {
        return;
    }

    // Compute total footnote block height
    let mut total_fn_height = 0.0f32;
    for fn_id in fn_ids {
        if let Some(footnote) = footnotes.get(fn_id) {
            total_fn_height +=
                compute_footnote_height(footnote, seen_fonts, text_width, doc_line_spacing);
        }
    }
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

    // Render footnote paragraphs top-down from below separator
    let mut fn_y = sep_y - 9.0;
    for fn_id in fn_ids {
        let Some(footnote) = footnotes.get(fn_id) else {
            continue;
        };
        let display_num = footnote_display_order.get(fn_id).copied().unwrap_or(0);

        for para in &footnote.paragraphs {
            let substituted_runs: Vec<Run> = para
                .runs
                .iter()
                .map(|run| {
                    if run.is_footnote_ref_mark {
                        let mut r = run.clone();
                        r.text = display_num.to_string();
                        r
                    } else {
                        run.clone()
                    }
                })
                .collect();

            if is_text_empty(&substituted_runs) {
                continue;
            }

            let (fs, tallest_lhr, tallest_ar) =
                tallest_run_metrics(&substituted_runs, seen_fonts);
            let effective_ls = para.line_spacing.unwrap_or(LineSpacing::Auto(1.0));
            let lh = resolve_line_h(effective_ls, fs, tallest_lhr);

            let lines = build_paragraph_lines(
                &substituted_runs,
                seen_fonts,
                text_width,
                0.0,
                &HashMap::new(),
            );

            if lines.is_empty() {
                continue;
            }

            let ascender_ratio = tallest_ar.unwrap_or(0.75);
            let baseline_y = fn_y - fs * ascender_ratio;

            render_paragraph_lines(
                content,
                &lines,
                &para.alignment,
                margin_left,
                text_width,
                baseline_y,
                lh,
                lines.len(),
                0,
                &mut Vec::new(),
                0.0,
                seen_fonts,
            );

            fn_y -= lines.len() as f32 * lh;
        }
    }
}
