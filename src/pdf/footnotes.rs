use std::collections::HashMap;

use pdf_writer::Content;

use crate::model::{Footnote, LineSpacing, Paragraph, Run};

use super::RenderContext;
use super::color::stroke_segment;
use super::helpers::{effective_space_after, effective_space_before};
use super::layout::{
    LineOpts, LinkTagger, TextLine, boxed_line_ascent, build_lines, is_text_empty, lines_height,
    os2_strike, render_paragraph_lines, size_lines_by_own_runs, tallest_run_metrics,
};
use super::list_label::render_list_label;
use super::resolve_line_h;
use super::tagging::NoteTagger;

/// Destination name a note's reference mark links to. Word bookmark names
/// can't contain spaces, so it never collides with a real bookmark.
pub(super) fn note_anchor(endnote: bool, id: u32) -> String {
    format!("{} {id}", if endnote { "endnote" } else { "footnote" })
}

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

impl ParagraphLayout {
    fn height(&self) -> f32 {
        lines_height(
            &self.lines,
            self.line_height,
            (self.font_size * self.ascender_ratio, 0.0),
        )
    }
}

/// `runs` are `para`'s with the reference mark filled in; a tab goes to
/// the paragraph's stops as in the body (uk_commercial_lease's notes tab to
/// their 567-twip stop on the mark's line).
fn layout_paragraph(
    runs: &[Run],
    para: &Paragraph,
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
    let mut lines = build_lines(
        runs,
        ctx,
        text_width,
        ctx.cjk(true, para),
        &LineOpts {
            tab_stops: &para.tab_stops,
            indent_left: para.indent_left,
            indent_right: para.indent_right,
            hanging: first_line_hanging,
            ..Default::default()
        },
    );
    // Each line as tall as its own runs, as in the body: a note's larger
    // reference mark raises only its own line (uk_commercial_lease: 10pt
    // marks over 8pt text, the following lines step 9.5 in Word).
    let ascender_ratio = tallest_ar.unwrap_or(0.75);
    if !matches!(line_spacing, LineSpacing::Exact(_)) {
        size_lines_by_own_runs(&mut lines, ctx.fonts, line_spacing, lh, fs * ascender_ratio, para.text_alignment == crate::model::TextAlignment::Top);
    }
    if lines.is_empty() {
        return None;
    }
    Some(ParagraphLayout {
        font_size: fs,
        line_height: lh,
        ascender_ratio,
        lines,
    })
}

/// Word keeps one line for an empty paragraph, sized by its paragraph mark —
/// the synthetic run the parser leaves in an empty paragraph carries that font.
fn empty_paragraph_line_h(para: &Paragraph, ls: LineSpacing, ctx: &RenderContext) -> f32 {
    let (fs, lhr, _) = tallest_run_metrics(&para.runs, ctx.fonts);
    resolve_line_h(ls, fs, lhr)
}

/// Footnote `id` laid out as `render_page_footnotes` draws it, the note's
/// mark included (a mark larger than the text raises its line); 0 if missing.
pub(super) fn footnote_height(
    id: u32,
    footnotes: &HashMap<u32, Footnote>,
    ctx: &RenderContext,
    text_width: f32,
) -> f32 {
    let Some(footnote) = footnotes.get(&id) else {
        return 0.0;
    };
    let mark = ctx.footnote_marks.get(&id).map_or("1", String::as_str);
    let mut total = 0.0f32;
    let mut prev_space_after = 0.0f32;
    let mut prev_para = None;
    for (i, para) in footnote.paragraphs.iter().enumerate() {
        let ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
        let para_text_width = (text_width - para.indent_left - para.indent_right).max(1.0);
        let hanging = super::compute_text_hanging(para, 0.0, ctx.fonts);
        let runs = substitute_ref_marks(&para.runs, mark);
        let layout = layout_paragraph(&runs, para, ls, ctx, para_text_width, hanging);
        if layout.is_none() && para.paragraph_mark_vanish {
            continue;
        }
        if i > 0 {
            total += f32::max(prev_space_after, effective_space_before(para, prev_para));
        }
        total += layout.as_ref().map_or_else(
            || empty_paragraph_line_h(para, ls, ctx),
            ParagraphLayout::height,
        );
        prev_space_after = effective_space_after(para, footnote.paragraphs.get(i + 1));
        prev_para = Some(para);
    }
    // Word puts a note's last space-after before the next note (erasmus_plus
    // endnotes 5pt, master_thesis footnotes 3pt). ponytail: the last note is
    // charged too, so a bottom-anchored block ends that far above the margin;
    // the next note's space-before is untested.
    total + prev_space_after
}

/// The separator above a page's footnotes, laid out as Word lays out
/// footnotes.xml's separator paragraph (Word probes: Calibri 6-20pt; single,
/// double and exact lines; space before and after): one line of its font and
/// spacing with its space after below, right on the first note, and the rule
/// that line's strikethrough, 144pt long. Space before goes nowhere.
pub(crate) struct NoteSeparator {
    pub(super) height: f32,
    /// The rule's centre below the separator's top, and its thickness.
    rule: (f32, f32),
}

impl NoteSeparator {
    pub(super) fn new(
        separator: Option<&Paragraph>,
        fonts: &HashMap<String, crate::fonts::FontEntry>,
        doc_line_spacing: LineSpacing,
    ) -> Self {
        let Some(p) = separator else {
            return Self {
                height: 12.0,
                rule: (3.0, 0.5),
            };
        };
        let ls = p.line_spacing.unwrap_or(doc_line_spacing);
        let (fs, lhr, ar) = tallest_run_metrics(&p.runs, fonts);
        let line_h = resolve_line_h(ls, fs, lhr);
        let baseline = boxed_line_ascent(ls, line_h, fs, lhr, ar, &p.runs, fonts)
            .unwrap_or(fs * ar.unwrap_or(0.8));
        let (top, thickness) = p
            .runs
            .first()
            .and_then(|r| fonts.get(&crate::fonts::font_key(r)))
            .and_then(|e| os2_strike(e, fs))
            .unwrap_or((fs * 0.25, 0.5));
        Self {
            height: line_h + p.space_after,
            rule: (baseline - top + thickness / 2.0, thickness),
        }
    }
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
    notes: NoteTagger<'_>,
) -> Vec<(u32, f32)> {
    if fn_ids.is_empty() {
        return Vec::new();
    }

    let total_fn_height: f32 = fn_ids
        .iter()
        .map(|&id| footnote_height(id, footnotes, ctx, text_width))
        .sum();

    let fn_y = margin_bottom + total_fn_height;
    let separator = &ctx.note_separator;
    let (rule_centre, thickness) = separator.rule;
    let sep_y = fn_y + separator.height - rule_centre;
    draw_note_separator(content, margin_left, sep_y, text_width, thickness);

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
        notes,
    )
}

fn draw_note_separator(
    content: &mut Content,
    margin_left: f32,
    sep_y: f32,
    text_width: f32,
    thickness: f32,
) {
    // Black rule, ~1/3 page width — matches Word's footnote/endnote separator
    let sep_width = 144.0f32.min(text_width);
    stroke_segment(
        content,
        (margin_left, sep_y),
        (margin_left + sep_width, sep_y),
        thickness,
        None,
    );
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
    notes: NoteTagger<'_>,
) -> Vec<(u32, f32)> {
    if en_ids.is_empty() {
        return Vec::new();
    }
    // endnotes.xml's separator paragraph, laid out as the footnote one
    // (case75: Word's rule 20.8pt below the last body baseline, the first
    // note 13.5pt below the rule).
    let separator = &ctx.endnote_separator;
    let (rule_centre, thickness) = separator.rule;
    draw_note_separator(
        content,
        margin_left,
        top_y - rule_centre,
        text_width,
        thickness,
    );
    let fn_y = top_y - separator.height;
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
        notes,
    )
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
    notes: NoteTagger<'_>,
) -> Vec<(u32, f32)> {
    // Top of each note drawn, for the link from its reference mark.
    let mut tops = Vec::new();
    for fn_id in fn_ids {
        let Some(footnote) = footnotes.get(fn_id) else {
            continue;
        };
        tops.push((*fn_id, fn_y));
        let note = notes.tags.note(notes.endnote, *fn_id, super::tagging::ROOT);
        let display_num = footnote_display_order
            .get(fn_id)
            .cloned()
            .unwrap_or_else(|| "1".to_string());

        let mut prev_space_after = 0.0f32;
        let mut prev_para = None;
        for (pi, para) in footnote.paragraphs.iter().enumerate() {
            let runs = substitute_ref_marks(&para.runs, &display_num);
            let ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);

            let para_text_x = margin_left + para.indent_left;
            let para_text_width = (text_width - para.indent_left - para.indent_right).max(1.0);

            let hanging = super::compute_text_hanging(para, 0.0, ctx.fonts);
            let layout = layout_paragraph(&runs, para, ls, ctx, para_text_width, hanging);
            if layout.is_none() && para.paragraph_mark_vanish {
                continue;
            }

            // Inter-paragraph spacing within the footnote
            if pi > 0 {
                fn_y -= f32::max(prev_space_after, effective_space_before(para, prev_para));
            }

            if let Some(layout) = layout {
                let baseline_y = fn_y - layout.font_size * layout.ascender_ratio;
                let line_count = layout.lines.len();

                let p = notes.tags.add(note, "P");
                notes.tags.begin(content, notes.page, p);
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
                    // Footnote lines carry no inline pictures, so no descent is needed.
                    (layout.font_size * layout.ascender_ratio, 0.0),
                    line_count,
                    0,
                    notes.links,
                    hanging,
                    ctx.fonts,
                    None,
                    gradient_specs,
                    None,
                    None,
                    Some(LinkTagger::new(notes.tags, notes.page, p)),
                );
                super::tagging::Tags::end(content);

                fn_y -= layout.height();
            } else {
                fn_y -= empty_paragraph_line_h(para, ls, ctx);
            }
            prev_space_after = effective_space_after(para, footnote.paragraphs.get(pi + 1));
            prev_para = Some(para);
        }
        // The note's trailing space, as `footnote_height` charges it.
        fn_y -= prev_space_after;
    }
    tops
}
