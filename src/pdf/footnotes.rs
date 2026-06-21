use std::collections::HashMap;

use pdf_writer::Content;

use crate::model::{Footnote, LineSpacing, Run};

use super::RenderContext;
use super::layout::{
    TextLine, build_paragraph_lines, is_text_empty, render_paragraph_lines, tallest_run_metrics,
};
use super::list_label::render_list_label;
use super::resolve_line_h;

fn substitute_ref_marks(runs: &[Run], display_num: &str) -> Vec<Run> {
    runs.iter()
        .map(|run| {
            if run.is_footnote_ref_mark || run.is_endnote_ref_mark {
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
    first_line_hanging: f32,
) -> Option<ParagraphLayout> {
    if is_text_empty(runs) {
        return None;
    }
    let (fs, tallest_lhr, tallest_ar) = tallest_run_metrics(runs, ctx.fonts);
    let lh = resolve_line_h(line_spacing, fs, tallest_lhr);
    let lines = build_paragraph_lines(runs, ctx.fonts, text_width, first_line_hanging, &HashMap::new(), &HashMap::new(), None, None, None, true);
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
    let mut total = 0.0f32;
    let mut prev_space_after = 0.0f32;
    let mut prev_contextual = false;
    for (i, para) in footnote.paragraphs.iter().enumerate() {
        let ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
        let para_text_width =
            (text_width - para.indent_left - para.indent_right).max(1.0);
        let hanging = super::compute_text_hanging(para, 0.0);
        let Some(layout) = layout_paragraph(&para.runs, ls, ctx, para_text_width, hanging) else {
            continue;
        };
        if i > 0 {
            let effective_sb = if para.contextual_spacing && prev_contextual {
                0.0
            } else {
                para.space_before
            };
            total += f32::max(prev_space_after, effective_sb);
        }
        total += layout.lines.len().max(1) as f32 * layout.line_height;
        prev_space_after = if para.contextual_spacing
            && footnote.paragraphs.get(i + 1).is_some_and(|p| p.contextual_spacing)
        {
            0.0
        } else {
            para.space_after
        };
        prev_contextual = para.contextual_spacing;
    }
    total
}

pub(super) fn render_page_footnotes(
    content: &mut Content,
    fn_ids: &[u32],
    footnotes: &HashMap<u32, Footnote>,
    footnote_display_order: &HashMap<u32, String>,
    ctx: &RenderContext,
    margin_left: f32,
    margin_bottom: f32,
    text_width: f32,
    gradient_specs: &mut Vec<super::GradientSpec>,
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
    draw_note_separator(content, margin_left, sep_y, text_width);

    let fn_y = sep_y - 9.0;
    render_notes_downward(
        content,
        fn_y,
        fn_ids,
        footnotes,
        footnote_display_order,
        ctx,
        margin_left,
        text_width,
        gradient_specs,
    );
}

fn draw_note_separator(content: &mut Content, margin_left: f32, sep_y: f32, text_width: f32) {
    // 0.5pt black rule, ~1/3 page width — matches Word's footnote/endnote separator
    let sep_width = 144.0f32.min(text_width);
    content.save_state();
    content.set_line_width(0.5);
    content.move_to(margin_left, sep_y);
    content.line_to(margin_left + sep_width, sep_y);
    content.stroke();
    content.restore_state();
}

/// Endnotes default to `pos=docEnd`: Word flows them in the normal content
/// stream right after the last body block (NOT pinned to the page bottom like
/// footnotes). `top_y` is the body cursor below the last block.
/// ponytail: single-page flow only — endnotes that overflow the bottom margin
/// are not paginated to a new page; add when a fixture needs it.
pub(super) fn render_endnotes_inline(
    content: &mut Content,
    top_y: f32,
    en_ids: &[u32],
    endnotes: &HashMap<u32, Footnote>,
    endnote_display_order: &HashMap<u32, String>,
    ctx: &RenderContext,
    margin_left: f32,
    text_width: f32,
    gradient_specs: &mut Vec<super::GradientSpec>,
) {
    if en_ids.is_empty() {
        return;
    }
    let sep_y = top_y;
    draw_note_separator(content, margin_left, sep_y, text_width);
    let fn_y = sep_y - 9.0;
    render_notes_downward(
        content,
        fn_y,
        en_ids,
        endnotes,
        endnote_display_order,
        ctx,
        margin_left,
        text_width,
        gradient_specs,
    );
}

#[allow(clippy::too_many_arguments)]
fn render_notes_downward(
    content: &mut Content,
    mut fn_y: f32,
    fn_ids: &[u32],
    footnotes: &HashMap<u32, Footnote>,
    footnote_display_order: &HashMap<u32, String>,
    ctx: &RenderContext,
    margin_left: f32,
    text_width: f32,
    gradient_specs: &mut Vec<super::GradientSpec>,
) {
    for fn_id in fn_ids {
        let Some(footnote) = footnotes.get(fn_id) else {
            continue;
        };
        let display_num = footnote_display_order
            .get(fn_id)
            .cloned()
            .unwrap_or_else(|| "1".to_string());

        let mut prev_space_after = 0.0f32;
        let mut prev_contextual = false;
        for (pi, para) in footnote.paragraphs.iter().enumerate() {
            let runs = substitute_ref_marks(&para.runs, &display_num);
            let ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);

            let para_text_x = margin_left + para.indent_left;
            let para_text_width =
                (text_width - para.indent_left - para.indent_right).max(1.0);

            let hanging = super::compute_text_hanging(para, 0.0);
            let Some(layout) = layout_paragraph(&runs, ls, ctx, para_text_width, hanging) else {
                continue;
            };

            // Inter-paragraph spacing within the footnote
            if pi > 0 {
                let effective_sb = if para.contextual_spacing && prev_contextual {
                    0.0
                } else {
                    para.space_before
                };
                fn_y -= f32::max(prev_space_after, effective_sb);
            }

            let baseline_y = fn_y - layout.font_size * layout.ascender_ratio;
            let line_count = layout.lines.len();

            render_list_label(
                content,
                para,
                ctx.fonts,
                para_text_x - para.indent_hanging,
                baseline_y,
                layout.font_size,
            );

            render_paragraph_lines(
                content,
                &layout.lines,
                &para.alignment,
                para_text_x,
                para_text_width,
                baseline_y,
                layout.line_height,
                line_count,
                0,
                &mut Vec::new(),
                hanging,
                ctx.fonts,
                None,
                gradient_specs,
                None,
                None,
            );

            fn_y -= line_count as f32 * layout.line_height;
            prev_space_after = if para.contextual_spacing
                && footnote.paragraphs.get(pi + 1).is_some_and(|p| p.contextual_spacing)
            {
                0.0
            } else {
                para.space_after
            };
            prev_contextual = para.contextual_spacing;
        }
    }
}
