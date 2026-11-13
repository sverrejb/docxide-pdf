//! Line metrics: heights and baselines of paragraph marks, breaks and unsized lines.

use super::*;

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
pub(super) fn label_boosted_baseline_offset(
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
pub(super) fn unsized_line_metrics(
    para: &Paragraph,
    font_size: f32,
    fonts: &HashMap<String, FontEntry>,
) -> (f32, Option<f32>) {
    let (size, metrics) = unsized_line_font(para, font_size, fonts);
    (size, metrics.0)
}

/// The font that sizes a line no glyph run sizes (only tabs and spaces): the
/// first break's, else the paragraph mark's, as (size, run_line_metrics).
pub(super) fn unsized_line_font(
    para: &Paragraph,
    font_size: f32,
    fonts: &HashMap<String, FontEntry>,
) -> (f32, (Option<f32>, Option<f32>)) {
    match para.runs.iter().find(|r| r.is_line_break) {
        Some(br) => break_font(br, fonts),
        None => mark_font(para, font_size, fonts),
    }
}

pub(super) fn break_font(br: &Run, fonts: &HashMap<String, FontEntry>) -> (f32, (Option<f32>, Option<f32>)) {
    let metrics = fonts
        .get(&font_key(br))
        .map_or((None, None), |e| run_line_metrics(e, ""));
    (br.font_size, metrics)
}

pub(super) fn mark_font(
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
pub(super) fn size_unsized_lines(
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
        let br = if line.ends_with_break {
            breaks.next()
        } else {
            None
        };
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
pub(super) fn mark_position_stretch(
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
pub(super) fn size_break_lines(
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
pub(super) fn mark_line_h(
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
pub(super) fn size_mark_line(
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
pub(super) fn size_tab_mark_line(
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
pub(super) fn break_run_lhr(runs: &[Run], break_fs: f32, fonts: &HashMap<String, FontEntry>) -> Option<f32> {
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

#[cfg(test)]
mod tests {
    use super::*;

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
