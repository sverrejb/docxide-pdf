//! Whole-document prepass: language, PAGE field text and bookmark pages, before layout.

use super::*;

/// Every run of the body: its paragraphs and table cells.
pub(super) fn body_runs<'a>(doc: &'a Document) -> impl Iterator<Item = &'a Run> + 'a {
    doc.sections.iter().flat_map(|s| s.blocks.iter()).flat_map(
        |block| -> Box<dyn Iterator<Item = &'a Run> + 'a> {
            match block {
                Block::Paragraph(p) => Box::new(p.runs.iter()),
                Block::Table(t) => Box::new(t.all_paragraphs().flat_map(|p| p.runs.iter())),
            }
        },
    )
}

/// The catalog `/Lang`: the language most of the body text is in (by letters,
/// East Asian ones by their own language), so only passages in another one
/// need a `/Lang` Span. The most common primary subtag wins (en-US and en-GB
/// together outvote fr-FR), then its most common tag; else the declared
/// default, else Word's en-US.
pub(super) fn document_lang(doc: &Document) -> String {
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
pub(super) fn body_page_field_text(
    run: &Run,
    ctx: &RenderContext,
    page_idx: usize,
    sect_idx: usize,
) -> Option<String> {
    let fields = ctx.body_page_fields?;
    match run.field_code {
        Some(FieldCode::Page(fmt)) => {
            let (num, section_fmt) = fields.pages.get(page_idx)?;
            Some(hf_fields::page_field_text(
                *num,
                fmt.or(section_fmt.as_deref()),
            ))
        }
        Some(FieldCode::NumPages(fmt)) => Some(hf_fields::page_field_text(fields.total, fmt)),
        Some(FieldCode::SectionPages(fmt)) => Some(hf_fields::page_field_text(
            fields.section_pages[sect_idx],
            fmt,
        )),
        _ => None,
    }
}

/// Pre-compute bookmark page positions so PAGEREF fields (e.g. TOC) can
/// show correct page numbers, and each simulated page's label. Simulates
/// page layout without rendering.
pub(super) fn compute_bookmark_positions(
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
    // The section each simulated page opens in, shaped like
    // `PageBuilder::page_section_indices` for `page_number_formats`; a
    // bookmark's page is the last one opened.
    let mut pages = vec![(0, false, 0)];
    // Opens the next simulated page of section `si`: its slot top and
    // effective bottom margin.
    let open_page =
        |pages: &mut Vec<(usize, bool, usize)>, si: usize, sp: &SectionProperties, first: bool| {
            pages.push((si, false, si));
            let page = last_page_number(&doc.sections, pages.iter());
            (
                effective_slot_top(sp, first, page, ctx),
                compute_effective_margin_bottom(sp, first, page, ctx),
            )
        };
    let first_sp = &doc.sections[0].properties;
    let mut sp = first_sp;
    let page = last_page_number(&doc.sections, &pages);
    let mut slot_top = effective_slot_top(sp, true, page, ctx);
    let mut margin_bottom = compute_effective_margin_bottom(sp, true, page, ctx);
    let mut prev_space_after: f32 = 0.0;
    let mut prev_para: Option<&Paragraph> = None;

    for (si, section) in doc.sections.iter().enumerate() {
        sp = &section.properties;
        if si > 0 {
            match sp.break_type {
                SectionBreakType::NextPage
                | SectionBreakType::OddPage
                | SectionBreakType::EvenPage => {
                    (slot_top, margin_bottom) = open_page(&mut pages, si, sp, true);
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
                        && slot_top
                            < effective_slot_top(
                                sp,
                                false,
                                last_page_number(&doc.sections, &pages),
                                ctx,
                            )
                    {
                        (slot_top, margin_bottom) = open_page(&mut pages, si, sp, false);
                        prev_space_after = 0.0;
                    }
                    for bm in &para.bookmarks {
                        bookmark_positions.insert(bm.clone(), (pages.len() - 1, slot_top));
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
                    let hanging = list_label::text_hanging(para, ctx.default_tab_stop, ctx.fonts);
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
                    let next_para = block_para(blocks, bi + 1);
                    let effective_sa = helpers::effective_space_after(para, next_para);
                    let inter_gap = f32::max(prev_space_after, effective_sb);
                    let needed = inter_gap + content_h;
                    if slot_top - needed < margin_bottom
                        && slot_top
                            < effective_slot_top(
                                sp,
                                false,
                                last_page_number(&doc.sections, &pages),
                                ctx,
                            )
                    {
                        (slot_top, margin_bottom) = open_page(&mut pages, si, sp, false);
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
                        (slot_top, margin_bottom) = open_page(&mut pages, si, sp, false);
                    }
                    // Captions in cells count as on the table's page (issue
                    // #146: a panel table's TOC entries had no page numbers).
                    // ponytail: a table that splits keeps its first page.
                    for bm in cells.iter().flatten().flat_map(|p| &p.bookmarks) {
                        bookmark_positions.insert(bm.clone(), (pages.len() - 1, slot_top));
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
        .map(|(n, fmt)| hf_fields::page_field_text(n, fmt))
        .collect();
    (bookmark_positions, labels)
}

/// An empty section mark keeps its line after an inline table (the rule is
/// explained in `render_paragraph_block`); bookmark positions follow it too.
pub(super) fn keeps_section_mark_line(
    compat_mode: u32,
    next_continuous: bool,
    prev: Option<&Block>,
) -> bool {
    compat_mode >= 15
        && next_continuous
        && matches!(prev, Some(Block::Table(t)) if t.position.is_none())
}

pub(super) fn cell_grid_pitch(doc: &Document, sp: &crate::model::SectionProperties) -> f32 {
    sp.line_grid_pitch()
        .filter(|_| doc.adjust_line_height_in_table)
        .unwrap_or(0.0)
}
