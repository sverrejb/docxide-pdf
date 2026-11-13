//! Body paragraphs: `render_paragraph_block` and the helpers only it uses.

use super::*;

/// Where a split paragraph's final fragment goes around its trailing float.
enum FragmentPlan {
    /// Lines `from..to` sit beside the float in a narrowed box.
    Beside {
        from: usize,
        to: usize,
        text_x: f32,
        text_w: f32,
    },
    /// Lines from `from` on start below the float, at `top`.
    Below { from: usize, top: f32 },
    /// Below is off the page: the lines above the band keep their place
    /// (the page's first line stays even inside the band) and the rest
    /// move on, the float with them.
    NextPage { keep: usize },
    /// The anchor line alone on a page, below the float's off-page bottom:
    /// Word sets it at the page bottom above the footer.
    PageBottom,
}

/// The float's anchor follows every run with visible text, a break, a tab or
/// an inline picture: its line is the paragraph's last.
/// No run from `index` on shows anything. ponytail: the index counts the
/// runs as parsed; a run dropped since (a deleted page break) only moves it
/// past the end.
fn blank_after(runs: &[crate::model::Run], index: usize) -> bool {
    runs[index.min(runs.len())..].iter().all(|r| {
        r.text.trim().is_empty() && !r.is_line_break && !r.is_tab && r.inline_image.is_none()
    })
}

/// The floats `keep` selects, with their XObject names re-keyed to their
/// positions in the subset (`render_floating_images` keys by index).
fn float_subset(
    fis: &[crate::model::FloatingImage],
    pdf_names: &HashMap<(usize, usize), String>,
    effect_names: &HashMap<(usize, usize), images::EffectXObjs>,
    block_idx: usize,
    keep: impl Fn(&crate::model::FloatingImage) -> bool,
) -> (
    Vec<crate::model::FloatingImage>,
    HashMap<(usize, usize), String>,
    HashMap<(usize, usize), images::EffectXObjs>,
) {
    let mut sub = Vec::new();
    let mut names = HashMap::new();
    let mut effects = HashMap::new();
    for (i, fi) in fis.iter().enumerate() {
        if !keep(fi) {
            continue;
        }
        let j = sub.len();
        if let Some(n) = pdf_names.get(&(block_idx, i)) {
            names.insert((block_idx, j), n.clone());
        }
        if let Some(e) = effect_names.get(&(block_idx, i)) {
            effects.insert((block_idx, j), e.clone());
        }
        sub.push(fi.clone());
    }
    (sub, names, effects)
}

/// Whether a picture offset below its paragraph's top leaves no usable text
/// strip on either side of it.
fn offset_picture_without_strip(
    fi: &crate::model::FloatingImage,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> bool {
    if fi.v_relative_from != VRelativeFrom::Paragraph
        || !matches!(fi.v_position, VerticalPosition::Offset(y) if y > 0.0)
    {
        return false;
    }
    let x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
    (x - fi.dist_left - col_x).max(col_x + col_w - x - fi.image.display_width - fi.dist_right)
        < MIN_EMPTY_STRIP
}

/// Whether a leading picture leaves no usable text strip on either side.
fn leading_full_width_picture(
    fi: &crate::model::FloatingImage,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> bool {
    fi.anchor_at_start
        && matches!(fi.wrap_type, WrapType::Square | WrapType::Tight)
        && offset_picture_without_strip(fi, sp, col_x, col_w, text_width)
}

#[allow(clippy::too_many_arguments)]
fn render_para_textbox(
    tb: &crate::model::Textbox,
    para: &Paragraph,
    anchor: ParaAnchor,
    sp: &SectionProperties,
    (col_x, col_w): (f32, f32),
    text_width: f32,
    ctx: &RenderContext,
    pb: &mut PageBuilder,
) {
    let tb_x = col_x + if tb.on_line { para.indent_left } else { 0.0 };
    let page = pb.all_contents.len();
    // Each box renders into its own buffer; flush_page paints these behind or
    // above the page's text layer sorted by relativeHeight (Word's z-order for
    // floating shapes spans paragraphs).
    let mut shape_content = tagging::artifact_content();
    render_single_textbox(
        tb,
        sp,
        tb_x,
        col_w,
        text_width,
        anchor.drawing_top(tb.on_line, tb.v_from_line),
        &mut shape_content,
        &mut pb.gradient_specs,
        ctx,
        &mut pb.links,
        &mut pb.tags,
        page,
        &mut pb.alpha_states,
    );
    let shapes = if tb.behind_doc {
        &mut pb.behind_shapes
    } else {
        &mut pb.deferred_shapes
    };
    shapes.push((tb.z_index, shape_content));
}

/// Render a single paragraph block. Returns `true` if the block was skipped
/// (the caller should `continue` the block loop).
#[allow(clippy::too_many_arguments)]
pub(super) fn render_paragraph_block(
    para: &Paragraph,
    state: &mut LayoutState,
    ctx: &RenderContext,
    sp: &SectionProperties,
    col_geometry: &[(f32, f32)],
    col_count: usize,
    text_width: f32,
    sect_idx: usize,
    block_idx: usize,
    section_blocks: &[Block],
    images: &EmbeddedImages,
    doc: &Document,
    smartart_font_key: &str,
    debug_wrap: bool,
    table_connector_anchor: Option<(usize, f32)>,
) -> bool {
    let prev_between_below = std::mem::take(&mut state.prev_between_below);
    let reflowed = std::mem::take(&mut state.reflowed);
    // One wholly beside the text keeps its (empty) paragraph, which Word
    // tags (empty-margin-frame-off-page).
    if is_empty_wrapping_frame(para)
        && !para
            .frame_props
            .as_ref()
            .is_some_and(|fp| frame_beside_text(fp, sp))
    {
        state.global_block_idx += 1;
        return true;
    }

    // §17.6.8: per-section line-number config (None if disabled). Holds no borrow
    // of `state`, so each render call can freshly borrow the shared counter.
    // Word numbers neither frame text (a drop cap's letter) nor suppressed lines.
    let numbered = para.frame_props.is_none() && !para.suppress_line_numbers;
    let ln_cfg: Option<(i32, u32, f32)> =
        sp.line_numbering.as_ref().filter(|_| numbered).map(|ln| {
            (
                ln.start,
                ln.count_by,
                sp.margin_left - ln.distance.unwrap_or(18.0),
            )
        });
    let ln_font: Option<(String, f32)> =
        sp.line_numbering
            .as_ref()
            .filter(|_| numbered)
            .and_then(|ln| {
                ctx.fonts
                    .get(&crate::fonts::font_key_for(&ln.font_name, false, false))
                    .map(|e| (e.pdf_name.clone(), ln.font_size))
            });
    if para.frame_props.is_none() && para.runs.iter().any(|run| !run.text.trim().is_empty()) {
        if let Some(bottom) = state.pb.next_text_below_picture.take() {
            state.pb.slot_top = state.pb.slot_top.min(bottom);
        }
    }
    let adjacent_para = |idx: usize| block_para(section_blocks, idx);

    // Skip empty section-break paragraphs — Word gives these zero height, also
    // before a continuous section that changes the columns (Word probes in
    // compat 14 and 15; covid_insomnia's two columns start 12pt higher) —
    // unless the paragraph is the first block of its section, including a
    // continuous empty section:
    // transition_to_work's contents start a line and 8pt below the top of the
    // page that empty section opens.
    let next_continuous = doc
        .sections
        .get(sect_idx + 1)
        .is_some_and(|next| next.properties.break_type == SectionBreakType::Continuous);
    // In modern compatibility mode, Word preserves the section mark's own
    // paragraph line after an inline table. Floating tables do not reserve
    // that extra line. Legacy mode still collapses it for both kinds.
    // Only before a continuous section: before a new page the line could only
    // push itself onto a blank page, which Word doesn't do (indigenous_innovation
    // ends a page with its signature table and starts Schedule A on the next).
    let after_float_table = follows_floating_table(section_blocks, block_idx);
    let keeps_line = block_idx == 0
        || keeps_section_mark_line(
            ctx.compat_mode,
            next_continuous,
            block_idx.checked_sub(1).and_then(|i| section_blocks.get(i)),
        )
        // A modern continuous section mark can occupy a side strip beside
        // a floating table. Full-width floats still absorb that mark line.
        || (ctx.compat_mode >= 15 && next_continuous
            && after_float_table
            && state.pb.float_zone.as_ref().is_some_and(|zone| {
                zone.side_strip(col_geometry[state.current_col]) >= MIN_EMPTY_STRIP
            }));

    if para.is_section_break
        && !keeps_line
        && lays_out_nothing(para)
        && para.image.is_none()
        && para.inline_chart.is_none()
    {
        // Its space after still meets the next section's first space before:
        // case25's sections (break paragraph after=10pt) start their 24pt
        // heading 14pt down, victorian's (after=0) its 26pt heading 26pt down.
        // At a page top nothing moves (both this branch's and main's probes,
        // 2026-10-05: gap = prev after + max(0, next before − break after)).
        let drop;
        (drop, state.prev_space_after) =
            section_break_spacing(state.prev_space_after, para.space_after, next_continuous);
        if !state.pb.is_at_page_top(sp) {
            state.pb.slot_top -= drop;
        }
        state.global_block_idx += 1;
        return true;
    }

    // Handle explicit page breaks.
    // `<w:br w:type="page"/>` (page_break_before_explicit) advances
    // unconditionally — even at the top of a page, Word emits a blank
    // page when the explicit break follows a section break. The
    // `<w:pageBreakBefore/>` style property is idempotent and skipped
    // when already at the top.
    if para.page_break_before {
        if !state.page_break_taken {
            take_page_break_before(state, para, sect_idx, sp, ctx);
        }
        state.prev_space_after = 0.0;
        if lays_out_nothing(para) {
            state.global_block_idx += 1;
            return true;
        }
    }

    // Handle explicit column breaks
    // In a one-column section Word breaks the page (bosch's page 2 ends there).
    // A frame's column break was taken as it opened (`take_frame_break`).
    if para.column_break_before && para.frame_props.is_none() {
        // The column ends below the space after its last paragraph (its
        // separator reaches there, Word probe).
        if col_count > 1 {
            state.pb.slot_top -= state.prev_space_after;
        }
        state.advance_column_or_page(col_count, sect_idx, sp, ctx);
        state.prev_space_after = 0.0;
    }

    let next_para = adjacent_para(block_idx + 1);
    let prev_para = block_idx.checked_sub(1).and_then(adjacent_para);

    // A continuous section boundary does not interrupt contextual spacing.
    // Border grouping still uses the section-local neighbours above.
    let context_prev = prev_para.or_else(|| {
        (block_idx == 0 && sp.break_type == SectionBreakType::Continuous)
            .then(|| sect_idx.checked_sub(1).and_then(|i| doc.sections.get(i)))
            .flatten()
            .and_then(|section| section.blocks.last())
            .and_then(leading_para)
    });
    let context_next = next_para.or_else(|| {
        (block_idx + 1 == section_blocks.len())
            .then(|| doc.sections.get(sect_idx + 1))
            .flatten()
            .filter(|section| section.properties.break_type == SectionBreakType::Continuous)
            .and_then(|section| section.blocks.first())
            .and_then(leading_para)
    });

    // A leading floating table is already the first body block. Word drops
    // the following empty anchor's before-spacing at the page top, while a
    // document's ordinary first paragraph retains its own before-spacing.
    let empty_anchor_after_float =
        is_text_empty(&para.runs) && state.pb.is_at_page_top(sp) && after_float_table;
    let effective_space_before = if empty_anchor_after_float {
        0.0
    } else {
        helpers::effective_space_before(para, context_prev)
    };
    let effective_space_after = helpers::effective_space_after(para, context_next);

    let mut inter_gap = if reflowed {
        0.0
    } else {
        inter_gap(state.prev_space_after, effective_space_before, ctx)
    };
    let (mut font_size, mut tallest_lhr, mut tallest_ar) =
        tallest_glyph_run_metrics(&para.runs, ctx.fonts);
    // A line of only tabs and spaces sits on the mark's (or break's) font,
    // its ascent included: an underlined tab hangs from that baseline.
    if tallest_lhr.is_none() {
        (font_size, (tallest_lhr, tallest_ar)) = unsized_line_font(para, font_size, ctx.fonts);
    }
    let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
    let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
    let grid_snapped = para.snap_to_grid
        && sp.line_grid_pitch().is_some()
        && !matches!(effective_ls, LineSpacing::Exact(_));
    let line_h = if grid_snapped {
        grid_snapped_line_h(&para.runs, ctx.fonts, effective_ls, line_h, sp.line_pitch)
    } else {
        line_h
    };
    let grid_baseline = grid_snapped
        .then(|| grid_baseline_offset(&para.runs, ctx.fonts, line_h))
        .flatten()
        .unwrap_or(sp.line_pitch);

    // The anchor the wrap and reservation passes below see, before a page
    // break or a page top's suppressed space before is known (peeked here,
    // taken when the drawings render).
    let pre_anchor = ParaAnchor::new(
        state.pb.float_anchor_for(block_idx),
        state.pb.anchor_slot_top(sp, ctx) - inter_gap,
        own_space_before(inter_gap, state.prev_space_after),
        sp,
        ctx,
    );

    // Word bottom-aligns text within an exact-height line box: the baseline
    // sits winDescent above the box bottom (identity: line_h_ratio −
    // ascender_ratio = winDescent/upm), however large the font's ascent is.
    // Placing the baseline at font_size * ascender_ratio instead pushes a
    // large-lineGap CJK substitute's descenders out of the fixed box and
    // into whatever follows (annotation #219: heading into table border).
    let exact_baseline_base = layout::boxed_line_ascent(
        effective_ls,
        line_h,
        font_size,
        tallest_lhr,
        tallest_ar,
        &para.runs,
        ctx.fonts,
    );

    // A single leading clearing break keeps its empty anchor line beside the
    // immediately preceding floating table, but its visible text resumes below.
    // Clearing only after painting the paragraph misplaces that first text line.
    let leading_table_clear = if para.clears_floats
        && after_float_table
        && matches!(para.runs.as_slice(), [first, rest @ ..]
            if first.is_line_break && !rest.iter().any(|r| r.is_line_break))
        && para
            .runs
            .iter()
            .any(|r| !r.text.trim().is_empty() || r.inline_image.is_some())
        && let Some(zone) = state
            .pb
            .float_zone
            .take_if(|zone| zone.from_table && state.pb.slot_top > zone.bottom_y)
    {
        state.pb.slot_top = zone.bottom_y + inter_gap;
        true
    } else {
        false
    };

    let (col_x, col_w) = col_geometry[state.current_col];
    let mut para_text_x = col_x + para.indent_left;
    let mut para_text_width = (col_w - para.indent_left - para.indent_right).max(1.0);
    let mut label_x = list_label::label_x(para, col_x, col_w, ctx.fonts);

    // When inside a floating object zone, narrow the paragraph to
    // fit beside the object rather than overlapping it.
    // The paragraph's first line starts below the gap: a centred line under
    // a logo sits clear of it once the 8pt after-space is counted.
    let first_line_top = state.pb.slot_top - inter_gap;
    if let Some(ref fz) = state.pb.float_zone {
        fz.narrow_paragraph(
            first_line_top,
            col_x,
            col_w,
            para,
            &mut para_text_x,
            &mut para_text_width,
            &mut label_x,
            ctx.compat_mode < 15,
        );
    }

    let text_hanging = compute_text_hanging(para, ctx.default_tab_stop, ctx.fonts);

    // Substitute footnote/endnote refs and resolve PAGEREF fields
    let has_footnote_refs = para.runs.iter().any(|r| r.footnote_id.is_some());
    let has_endnote_refs = para.runs.iter().any(|r| r.endnote_id.is_some());
    let has_pageref = para.runs.iter().any(|r| {
        r.field_code
            .as_ref()
            .is_some_and(|fc| fc.is_page_count() || matches!(fc, FieldCode::PageRef(_)))
    });
    let effective_runs: std::borrow::Cow<'_, Vec<Run>> =
        if has_footnote_refs || has_endnote_refs || has_pageref {
            let substituted: Vec<Run> = para
                .runs
                .iter()
                .map(|run| {
                    if let Some(text) = ctx.note_mark_text(run) {
                        Run {
                            text,
                            ..run.clone()
                        }
                    } else if let Some(FieldCode::PageRef(ref bookmark)) = run.field_code {
                        let mut r = run.clone();
                        // Word prints PAGEREF's cached result unless fields are updated
                        // before printing; only an empty result needs our estimate.
                        if r.text.trim().is_empty()
                            && let Some(&(page_idx, _)) = state.bookmark_positions.get(bookmark)
                        {
                            r.text = ctx
                                .page_labels
                                .get(page_idx)
                                .cloned()
                                .unwrap_or_else(|| (page_idx + 1).to_string());
                        }
                        r
                    } else if let Some(text) =
                        body_page_field_text(run, ctx, state.pb.all_contents.len(), sect_idx)
                    {
                        Run {
                            text,
                            field_code: None,
                            ..run.clone()
                        }
                    } else {
                        run.clone()
                    }
                })
                .collect();
            std::borrow::Cow::Owned(substituted)
        } else {
            std::borrow::Cow::Borrowed(&para.runs)
        };

    let text_empty = is_text_empty(&effective_runs);
    let has_tabs = effective_runs.iter().any(|r| r.is_tab);
    let block_inline_images: HashMap<usize, String> = images
        .inline_image_pdf_names
        .iter()
        .filter(|((bi, _), _)| *bi == state.global_block_idx)
        .map(|((_, ri), name)| (*ri, name.clone()))
        .collect();
    let block_effect_inlines: HashMap<usize, images::EffectXObjs> = images
        .effect_inline_names
        .iter()
        .filter(|((bi, _), _)| *bi == state.global_block_idx)
        .map(|((_, ri), fx)| (*ri, fx.clone()))
        .collect();
    // A float anchored after the paragraph's text belongs to the page its
    // last line lands on (the split path draws it there); when the forced
    // breaks alone outrun this page the paragraph will split, so that float
    // wraps none of the lines laid out here (an external tracked corpus
    // cover: five pages of break-separated lines, a page-wide picture after
    // them, set one letter per line beside a picture Word prints on page 5).
    let forced_lines = para.runs.iter().filter(|r| r.is_line_break).count() as f32 + 1.0;
    let splits_for_sure = forced_lines * line_h > first_line_top - state.effective_margin_bottom;
    let wraps_here = |fi: &crate::model::FloatingImage| {
        !(splits_for_sure && blank_after(&para.runs, fi.anchor_run))
    };
    // Self-wrapping: if this paragraph anchors a wrapping float
    // and has text, set up the float zone NOW so width-narrowing
    // applies to this paragraph's own lines.  Always replace any
    // previous float zone — the paragraph's own image takes priority.
    let mut own_textbox_zones: Vec<FloatZone> = if text_empty {
        Vec::new()
    } else {
        textbox_zones(para, sp, (col_x, col_w), text_width, pre_anchor, ctx)
    };
    let own_picture = para
        .floating_images
        .iter()
        .filter(|fi| !text_empty && wraps_here(fi))
        .find(|fi| wraps_in_column(fi, sp, col_x, col_w, text_width));
    // The previous paragraph may have re-wrapped around this float and
    // grown; the float stays where that look-ahead anchored it
    // (peeked here, taken below).
    let own_zone = match own_picture {
        Some(fi) => Some(FloatZone::for_float(
            fi,
            resolve_fi_x(fi, sp, col_x, col_w, text_width),
            pre_anchor.fi_y_top(fi, sp),
        )),
        None if !own_textbox_zones.is_empty() => Some(own_textbox_zones.remove(0)),
        None => None,
    };
    let own_zone_set = own_zone.is_some();
    let entry_zones = (
        state.pb.float_zone.clone(),
        state.pb.more_float_zones.clone(),
    );
    if let Some(zone) = own_zone {
        // Re-narrow para_text_x / para_text_width using the
        // new float zone (same logic as the block above).
        zone.narrow_paragraph(
            first_line_top,
            col_x,
            col_w,
            para,
            &mut para_text_x,
            &mut para_text_width,
            &mut label_x,
            ctx.compat_mode < 15,
        );
        state.pb.float_zone = Some(zone);
    }

    let cjk = ctx.cjk(para.auto_space_de || para.auto_space_dn, para);

    // Look-ahead: a wrapping float anchored in the *next* block (an image-only
    // paragraph) sits at that block's top, which Word computes from this
    // paragraph laid out at full width — and then re-wraps this paragraph's
    // lines around the float without moving it (case41 p3: the paragraph before
    // a centred 4.5in picture wraps beside it from its second line, annotation
    // #152). Install that zone now so the per-line geometry below narrows the
    // lines it reaches, and hand the anchor position to the next paragraph so
    // it draws the picture there rather than where it now flows.
    // (anchor paragraph top, the zone's real top edge) once installed.
    let mut lookahead: Option<(f32, f32, usize)> = None;
    if state.pb.float_zone.is_none() && !text_empty && !has_tabs {
        // The anchor paragraph may carry text of its own (case41 p6, annotation
        // #240): the anchor is its top either way. It may also sit up to three
        // paragraphs ahead: Word places the float from the paragraphs between
        // laid out at full width, then re-wraps them all beside it (Word for
        // Mac probes p26a and p26b: two short titles before the anchor both
        // wrap beside a page-tall picture at y 0 and beside a 300pt one
        // hoisted 60pt above the third paragraph). ponytail: plain paragraphs
        // between only, three at most.
        let qualifies = |fi: &crate::model::FloatingImage| {
            // Only floats that hang off the anchor paragraph itself;
            // a page- or margin-relative float does not move with it.
            fi.v_relative_from == VRelativeFrom::Paragraph
                && matches!(
                    fi.v_position,
                    VerticalPosition::Offset(_) | VerticalPosition::AlignTop
                )
                && wraps_in_column(fi, sp, col_x, col_w, text_width)
                // Only a float text can sit beside: Word wraps the
                // preceding paragraph next to case41 p3's centred
                // picture (64.8pt free on each side) but leaves
                // brazilian p9's caption alone above a figure with
                // 37.5pt beside it. ponytail: 48pt threshold, two
                // calibration points.
                && positioning::side_gaps(fi, sp, col_x, col_w, text_width)
                    .is_some_and(|(left, right)| left.max(right) >= 48.0)
        };
        // Find the anchor first: the paragraphs between are only measured when
        // there is one, so a plain paragraph costs no line breaking ahead.
        let mut found = None;
        let mut gaps = [0.0f32; 3];
        let mut prev_after = para.space_after;
        for k in 1..=3 {
            let Some(Block::Paragraph(np)) = section_blocks.get(block_idx + k) else {
                break;
            };
            if np.image.is_some() || np.inline_chart.is_some() || np.page_break_before {
                break;
            }
            gaps[k - 1] = self::inter_gap(prev_after, np.space_before, ctx);
            if let Some(fi) = np.floating_images.iter().find(|fi| qualifies(fi)) {
                found = Some((fi, k));
                break;
            }
            if !np.floating_images.is_empty() || !np.textboxes.is_empty() || np.page_break_after {
                break;
            }
            prev_after = np.space_after;
        }
        let first_gap = gaps[0];
        let found = found.map(|(fi, k)| {
            let mut between = 0.0f32;
            for j in 1..k {
                let Some(Block::Paragraph(np)) = section_blocks.get(block_idx + j) else {
                    unreachable!("the scan above only passes paragraphs");
                };
                between +=
                    gaps[j - 1] + line_count(np, ctx, col_w).max(1) as f32 * plain_line_h(np, ctx);
            }
            (fi, between + gaps[k - 1], block_idx + k)
        });
        if let Some((fi, gap, anchor_block)) = found {
            let full_lines = build_lines(
                &effective_runs,
                ctx,
                para_text_width,
                cjk,
                &LineOpts {
                    inline_images: Some(&block_inline_images),
                    effects: Some(&block_effect_inlines),
                    hanging: text_hanging,
                    ..Default::default()
                },
            );
            let anchor_top = state.pb.slot_top - inter_gap - full_lines.len() as f32 * line_h - gap;
            let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
            let fi_y_top = match fi.v_position {
                VerticalPosition::Offset(o) => anchor_top - o,
                _ => anchor_top,
            };
            let mut zone = FloatZone::for_float(fi, fi_x, fi_y_top);
            // Only when the float's top and its anchor land on this page: a
            // paragraph that breaks before its anchor would hand the next
            // page a stale anchor. (A page-tall picture's foot lies below the
            // margin and still wraps this page: p26a.)
            if zone.top_y > state.effective_margin_bottom
                && anchor_top > state.effective_margin_bottom
            {
                let true_top = zone.top_y;
                // Word treats the last line's space-after as part of that line
                // when testing overlap, so the zone reaches up through the gap
                // for this paragraph's geometry (restored after the lines are
                // built).
                zone.top_y += first_gap;
                state.pb.float_zone = Some(zone);
                lookahead = Some((anchor_top, true_top, anchor_block));
            }
        }
    }

    // Additional wrapping floats anchored to this same paragraph beyond the
    // first (which became `float_zone` above). The single-zone geometry below
    // can't model e.g. a logo on each side of a centered title, so when these
    // exist we compute per-line free intervals across all zones instead.
    let extra_float_zones: Vec<FloatZone> = if !text_empty {
        para.floating_images
            .iter()
            .filter(|fi| fi.wrap_type.wraps_beside() && wraps_here(fi))
            .filter(|fi| !image_blocks_band(fi, sp, col_x, col_w, text_width))
            .skip(usize::from(own_picture.is_some()))
            .map(|fi| {
                let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
                let fi_y_top = pre_anchor.fi_y_top(fi, sp);
                FloatZone::for_float(fi, fi_x, fi_y_top)
            })
            .chain(own_textbox_zones)
            .chain(if own_zone_set {
                Vec::new()
            } else {
                state.pb.more_float_zones.clone()
            })
            .collect()
    } else {
        Vec::new()
    };

    // Lines beside floats with no room on a side text may take (floating
    // tables keep their own rule).
    let mut blocked_lines = Vec::new();
    // Build per-line geometry and dual-region geometry
    let (mut poly_line_geom, mut poly_dual_geom): (
        Option<Vec<(f32, f32)>>,
        Option<Vec<DualRegion>>,
    ) = if let Some(fz) = state.pb.float_zone.as_ref() {
        let eff_top = state.pb.slot_top - inter_gap;
        if !extra_float_zones.is_empty() {
            // Multiple wrapping floats on one paragraph: subtract every
            // float's exclusion span per line, keep to the sides each
            // float's wrapText allows, and lay the text in the gaps left,
            // the first two from the left as a dual region (Word example:
            // between two boxes and right of the second).
            let zones: Vec<&FloatZone> = std::iter::once(fz)
                .chain(extra_float_zones.iter())
                .collect();
            if zones.iter().all(|z| eff_top <= z.bottom_y) {
                (None, None)
            } else {
                let full_w = (col_w - para.indent_left - para.indent_right).max(1.0);
                let col_right = col_x + col_w;
                let lowest_bottom = zones
                    .iter()
                    .map(|z| z.bottom_y)
                    .fold(f32::INFINITY, f32::min);
                let max_lines =
                    ((((eff_top - lowest_bottom) / line_h).ceil() as usize) + 5).max(50);
                let mut geom = Vec::with_capacity(max_lines);
                let mut dual: Vec<DualRegion> = Vec::with_capacity(max_lines);
                let mut two_regions = false;
                let legacy = ctx.compat_mode < 15;
                for i in 0..max_lines {
                    let line_top = eff_top - i as f32 * line_h;
                    let line_bottom = line_top - line_h;
                    let para_left = col_x + para.indent_left;
                    let mut intervals: Vec<(f32, f32)> =
                        vec![(para_left, col_right - para.indent_right)];
                    let mut float_edges: Vec<f32> = Vec::new();
                    let mut overlapped = false;
                    for z in &zones {
                        let partial_overlap = z.para_relative || z.polygon_pts.is_some();
                        // Require the line to dip >20% of its height below
                        // the zone top before counting as in-zone, matching
                        // the symmetric bottom_threshold. Without this, a
                        // caption line directly above a float offset down by
                        // ~a line height got flagged in-zone and wrapped to
                        // one word per line (annotation #120).
                        let in_zone = if partial_overlap {
                            line_bottom < z.top_y - line_h * 0.2
                        } else {
                            line_top <= z.top_y
                        };
                        let bottom_threshold = if z.from_table {
                            z.bottom_y + line_h * 0.2
                        } else {
                            z.bottom_y
                        };
                        if !(in_zone && line_top > bottom_threshold) {
                            continue;
                        }
                        let (ex_left, ex_right) =
                            z.exclusion_in_band(line_top.min(z.top_y), line_bottom);
                        let sl = ex_left - z.left_from_text;
                        let sr = ex_right + z.right_from_text;
                        float_edges.push(sr);
                        overlapped |= !z.from_table;
                        // Text keeps to the sides the float's wrapText allows.
                        let side = (z.wrap_text != WrapText::BothSides)
                            .then(|| wrap_side(z.wrap_text, sl - col_x, col_right - sr));
                        let (lo, hi) = match side {
                            None => (f32::MIN, f32::MAX),
                            Some(Some(true)) => (sr, f32::MAX),
                            Some(Some(false)) => (f32::MIN, sl),
                            Some(None) => (0.0, 0.0),
                        };
                        intervals = intervals
                            .into_iter()
                            .flat_map(|(a, b)| [(a, b.min(sl)), (a.max(sr), b)])
                            .map(|(a, b)| (a.max(lo), b.min(hi)))
                            .filter(|(a, b)| b > a)
                            .collect();
                    }
                    let min_w = if overlapped { MIN_EMPTY_STRIP } else { 1.0 };
                    let open: Vec<(f32, f32)> = intervals
                        .into_iter()
                        .filter(|(a, b)| b - a >= min_w)
                        .collect();
                    let region = |(a, b): (f32, f32)| {
                        let edge = float_edges
                            .iter()
                            .copied()
                            .filter(|&e| e <= a + 0.01)
                            .reduce(f32::max);
                        let a = match edge {
                            Some(e) if i == 0 => {
                                first_line_region_x(a, e, para, text_hanging, legacy)
                            }
                            _ => a,
                        };
                        (a, (b - a).max(1.0))
                    };
                    let pick = match open.len() {
                        n if para.bidi && n > 2 => &open[n - 2..],
                        _ => &open[..],
                    };
                    let (x, w, rx, rw) = match pick {
                        [] => {
                            blocked_lines.resize(i, false);
                            blocked_lines.push(overlapped);
                            (para_left, full_w, 0.0, 0.0)
                        }
                        [one] => {
                            let (x, w) = region(*one);
                            (x, w, 0.0, 0.0)
                        }
                        [first, second, ..] => {
                            two_regions = true;
                            let ((x, w), (rx, rw)) = (region(*first), region(*second));
                            (x, w, rx, rw)
                        }
                    };
                    geom.push((x, w));
                    dual.push((x, w, rx, rw));
                }
                (Some(geom), two_regions.then_some(dual))
            }
        } else if eff_top <= fz.bottom_y {
            (None, None)
        } else {
            let is_both_sides = fz.wrap_text == WrapText::BothSides;
            let full_w = (col_w - para.indent_left - para.indent_right).max(1.0);
            let col_right = col_x + col_w;
            let max_lines = ((eff_top - fz.bottom_y) / line_h).ceil() as usize + 5;
            let max_lines = max_lines.max(50);
            let mut geom = Vec::with_capacity(max_lines);
            let mut dual = if is_both_sides {
                Some(Vec::with_capacity(max_lines))
            } else {
                None
            };
            // A line whose box reaches into a picture's band wraps however
            // little: Word for Mac probe wrap_h* (a 24pt at-least heading
            // whose line top the band undercuts by 0.5pt wraps beside the
            // picture, one it clears by 1.5pt does not; a corpus title the
            // band undercut by 3.3pt wrapped in Word and ran full width in
            // ours under a 20% threshold). Floating tables keep that
            // threshold: probe tblwrap_d* wraps a line the band (margin
            // included) enters by 0.8pt, yet case40's line 1.6pt into a
            // right table's margin runs full width in Word — unresolved.
            let bottom_threshold = if fz.from_table {
                fz.bottom_y + line_h * 0.2
            } else {
                fz.bottom_y
            };
            // A line starting just above a float can still intersect it.
            // Apply the same overlap threshold to rectangular table zones
            // as to paragraph-relative images and polygon zones.
            for i in 0..max_lines {
                let line_top = eff_top - i as f32 * line_h;
                let line_bottom = line_top - line_h;
                // See annotation #120: require >20% line-height overlap
                // (symmetric with bottom_threshold) so a caption line
                // directly above a downward-offset float is not wrongly
                // squeezed into the float's side margin.
                let in_zone = line_bottom < fz.top_y - line_h * 0.2;
                if in_zone && line_top > bottom_threshold {
                    let (ex_left, ex_right) =
                        fz.exclusion_in_band(line_top.min(fz.top_y), line_bottom);
                    let float_left = ex_left - fz.left_from_text;
                    let float_right = ex_right + fz.right_from_text;
                    let sr = col_right - float_right;
                    let sl = float_left - col_x;
                    let (lx, mut lw) = left_text_region(col_x, col_right, float_left, para);
                    let mut right_of_float = right_text_region(col_x, col_right, float_right, para);
                    if i == 0 {
                        let (x, w) = right_of_float;
                        let legacy = ctx.compat_mode < 15;
                        let first = first_line_region_x(x, float_right, para, text_hanging, legacy);
                        right_of_float = (first, (w - (first - x)).max(0.0));
                        // The renderer widens the first line leftwards by a
                        // hanging indent: not past a float's edge.
                        if text_hanging > 0.0 && lw == 0.0 {
                            lw = (float_left - lx).max(-text_hanging);
                        }
                    }

                    if is_both_sides {
                        // A strip under the minimum stays empty on its own
                        // (Word example: 17pt left of the box, text right only).
                        let narrow = |room: f32| !fz.from_table && room < MIN_EMPTY_STRIP;
                        if narrow(sl) && narrow(sr) {
                            blocked_lines.resize(i, false);
                            blocked_lines.push(true);
                        }
                        if narrow(sl) {
                            lw = lw.min(0.0);
                        }
                        // BothSides: provide both regions
                        let (rx, rw) = right_of_float;
                        let rw = if narrow(sr) { 0.0 } else { rw };
                        if let Some(ref mut d) = dual {
                            d.push((lx, lw, rx, rw));
                        }
                        // Single-region geometry always stores the
                        // LEFT region — render_paragraph_lines uses
                        // this for left-chunk x positioning.
                        geom.push((lx, lw));
                    } else {
                        match wrap_side(fz.wrap_text, sl, sr) {
                            Some(true) => {
                                let (x, w) = right_of_float;
                                geom.push((x, w.max(1.0)));
                            }
                            Some(false) => geom.push((lx, lw.max(1.0))),
                            None => {
                                blocked_lines.resize(i, false);
                                blocked_lines.push(!fz.from_table);
                                geom.push((col_x + para.indent_left, full_w));
                            }
                        }
                    }
                } else {
                    geom.push((col_x + para.indent_left, full_w));
                    if let Some(ref mut d) = dual {
                        d.push((col_x + para.indent_left, full_w, 0.0, 0.0));
                    }
                }
            }
            (Some(geom), dual)
        }
    } else {
        (None, None)
    };
    // A paragraph whose first word fits no strip beside a float starts below
    // it, the word whole (Word examples: "Internationalization" and "Hi" +
    // no-break space + "Internationalization" beside a 40pt strip, the anchor's
    // own box or an earlier one).
    let from_table = state
        .pb
        .float_zone
        .iter()
        .chain(&extra_float_zones)
        .any(|z| z.from_table);
    if let Some(word) = layout::opening_word_width(&effective_runs, ctx.fonts)
        .filter(|_| !text_empty && !from_table)
    {
        let full_w = (col_w - para.indent_left - para.indent_right).max(1.0);
        let room = |i: usize| match (&poly_dual_geom, &poly_line_geom) {
            (Some(d), _) => d.get(i).map(|&(_, w, _, rw)| w.max(rw)),
            (None, Some(g)) => g.get(i).map(|&(_, w)| w),
            _ => None,
        };
        for i in 0.. {
            let Some(room) = room(i) else { break };
            if blocked_lines.get(i).copied().unwrap_or(false) {
                continue;
            }
            if room >= word || room >= full_w - 0.5 {
                break;
            }
            blocked_lines.resize(blocked_lines.len().max(i + 1), false);
            blocked_lines[i] = true;
        }
    }
    // A right-to-left paragraph fills the strips beside floats from the right
    // (Word example: Hebrew between a wrapText=right box and a bothSides one
    // starts in the strip right of the second).
    if para.bidi
        && let Some(dual) = poly_dual_geom.as_mut()
    {
        for (i, d) in dual.iter_mut().enumerate() {
            if d.3 > 0.0 {
                *d = (d.2, d.3, d.0, d.1);
                if let Some(g) = poly_line_geom.as_mut().and_then(|g| g.get_mut(i)) {
                    *g = (d.0, d.1);
                }
            }
        }
    }
    // A first line with no room beside a float moves to the float's foot, not
    // a whole line lower (Word example: a 17pt strip beside a wrapText=largest
    // picture puts the first line right under it).
    let blocked_top = blocked_lines.iter().take_while(|&&b| b).count();
    let float_drop = if blocked_top > 0 {
        let eff_top = state.pb.slot_top - inter_gap;
        let reach = (blocked_top as f32 + 0.2) * line_h;
        state
            .pb
            .float_zone
            .iter()
            .chain(extra_float_zones.iter())
            .map(|z| eff_top - z.bottom_y)
            .filter(|&d| d <= reach)
            .fold((blocked_top - 1) as f32 * line_h, f32::max)
    } else {
        0.0
    };
    blocked_lines.drain(..blocked_top);
    if let Some(g) = poly_line_geom.as_mut() {
        g.drain(..blocked_top.min(g.len()));
    }
    if let Some(d) = poly_dual_geom.as_mut() {
        d.drain(..blocked_top.min(d.len()));
    }
    if blocked_top > 0 {
        // The first line sits below the float: the lines past the per-line
        // geometry take the whole measure, not the strip beside it.
        para_text_x = col_x + para.indent_left;
        para_text_width = (col_w - para.indent_left - para.indent_right).max(1.0);
        label_x = list_label::label_x(para, col_x, col_w, ctx.fonts);
    }

    let poly_line_widths: Option<Vec<f32>> = poly_line_geom
        .as_ref()
        .map(|g| g.iter().map(|&(_, w)| w).collect());

    let has_inline_image_runs = effective_runs.iter().any(|r| r.inline_image.is_some());
    // Word advances a left tab past any floating image whose body sits on the
    // line, snapping to the first stop clear of the image's right edge. Collect
    // the horizontal spans (in from-text-margin coords) of wrapping images that
    // overlap this paragraph's first line so build_tabbed_line can skip them.
    let tab_exclusions: Vec<(f32, f32)> = if has_tabs && !text_empty {
        let slot_top = state.pb.slot_top;
        para.floating_images
            .iter()
            .filter(|fi| fi.wrap_type.wraps_beside() && wraps_here(fi))
            .filter_map(|fi| {
                let fi_y_top = pre_anchor.fi_y_top(fi, sp);
                let fi_y_bottom = fi_y_top - fi.image.display_height;
                // Only images vertically overlapping the first line band.
                if fi_y_bottom > slot_top + 2.0 || fi_y_top < slot_top - 40.0 {
                    return None;
                }
                let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
                Some((fi_x - col_x, fi_x + fi.image.display_width - col_x))
            })
            .collect()
    } else {
        Vec::new()
    };
    let mut lines = if para.image.is_some() || (text_empty && !has_inline_image_runs) {
        vec![]
    } else {
        // Per-line geometry handles narrow→wide transitions;
        // dual geometry takes priority over single-region widths.
        let plw: Option<&[f32]> = if poly_dual_geom.is_some() {
            None
        } else {
            poly_line_widths.as_deref()
        };
        let build = |per_line_widths: Option<&[f32]>, dual: Option<&[DualRegion]>| {
            build_lines(
                &effective_runs,
                ctx,
                para_text_width,
                cjk,
                &LineOpts {
                    inline_images: Some(&block_inline_images),
                    effects: Some(&block_effect_inlines),
                    tab_stops: &para.tab_stops,
                    indent_left: para.indent_left,
                    indent_right: para.indent_right,
                    hanging: text_hanging,
                    tab_exclusions: &tab_exclusions,
                    per_line_widths,
                    dual,
                },
            )
        };
        if blocked_lines.contains(&true) {
            let open = |v: &[_]| keep_open(v, &blocked_lines);
            let widths = plw.map(open);
            let dual = poly_dual_geom
                .as_deref()
                .map(|d| keep_open(d, &blocked_lines));
            skip_blocked_lines(&blocked_lines, build(widths.as_deref(), dual.as_deref()))
        } else {
            build(plw, poly_dual_geom.as_deref())
        }
    };
    // A pure one-break anchor can be taller than a short floating table.
    // In that case its first line extends below the float and occupies flow
    // height; dropping it also drops a full line before the paragraph mark.
    let short_empty_clear = || {
        let mut breaks = para.runs.iter().filter(|r| r.is_line_break);
        let (Some(br), None) = (breaks.next(), breaks.next()) else {
            return false;
        };
        para.runs
            .iter()
            .all(|r| r.text.is_empty() && !r.is_tab && r.inline_image.is_none())
            && para.image.is_none()
            && para.inline_chart.is_none()
            && para.smartart.is_empty()
            && para.floating_images.is_empty()
            && para.textboxes.is_empty()
            && state.pb.float_zone.as_ref().is_some_and(|zone| {
                let lhr = break_run_lhr(&effective_runs, br.font_size, ctx.fonts);
                let first_pitch = resolve_line_h(effective_ls, br.font_size, lhr);
                zone.from_table && state.pb.slot_top - inter_gap - first_pitch < zone.bottom_y
            })
    };
    // A clearing break that opens the anchor paragraph of the floating table
    // above it sits beside that table's top; only what follows the break comes
    // below the table (indigenous_innovation's defined terms: one line, not two).
    // The zone is gone once the table broke across pages or left no strip
    // (case129's full-width table: both lines below it), so ask the table's
    // own extent.
    if (leading_table_clear
        || para.clears_floats
            && after_float_table
            && !short_empty_clear()
            && state.pb.float_table_extent.as_ref().is_some_and(|t| {
                side_strip(
                    (t.left, t.right),
                    (t.left_from_text, t.right_from_text),
                    (col_x, col_w),
                ) >= MIN_EMPTY_STRIP
            }))
        && lines.len() > 1
        && lines[0].ends_with_break
        && lines[0].chunks.is_empty()
    {
        lines.remove(0);
    }
    // The look-ahead zone reached up through this paragraph's space-after only
    // for its own geometry; following paragraphs see the float's real edge.
    if let (Some((_, top, _)), Some(fz)) = (lookahead, state.pb.float_zone.as_mut()) {
        fz.top_y = top;
    }

    let max_inline_img_h = lines.iter().map(line_max_image_h).fold(0.0f32, f32::max);

    // Paragraph ascent/descent in points: the first baseline sits `para_ascent`
    // below the paragraph top, and picture lines are sized from both (see
    // `inline_line_advance` and `picture_line_bottom`). The bottom part is only
    // read for picture lines, so most paragraphs skip its run scan.
    let mark_ar = || {
        lines
            .iter()
            .any(|l| l.chunks.is_empty() && !l.end_spaces.is_empty())
            .then(|| mark_font(para, font_size, ctx.fonts).1.1)
            .flatten()
    };
    let para_ascent =
        exact_baseline_base.unwrap_or(font_size * tallest_ar.or_else(mark_ar).unwrap_or(0.75));
    let ascent_scale = auto_ascent_scale(effective_ls);
    let first_baseline_offset =
        label_boosted_baseline_offset(para, ctx.fonts, para_ascent, font_size) * ascent_scale;
    let line_grid = |ascent: f32, page: usize| {
        let text_h = font_size * tallest_lhr.unwrap_or(1.2);
        Some(layout::LineGrid {
            snapped: grid_snapped,
            shift: ctx.valign_offsets.get(page).copied().unwrap_or(0.0),
            label_boost: (ascent - para_ascent * ascent_scale).max(0.0),
            ..layout::LineGrid::new(
                sp.page_height,
                ascent,
                effective_ls,
                line_h,
                text_h,
                para_ascent,
                effective_space_after,
            )
        })
    };
    // A grid or an exact rule gives every line the same box.
    if !grid_snapped && !matches!(effective_ls, LineSpacing::Exact(_)) {
        size_lines_by_own_runs(
            &mut lines,
            ctx.fonts,
            effective_ls,
            line_h,
            para_ascent,
            ascent_scale,
            para.text_alignment == crate::model::TextAlignment::Top,
        );
        size_break_lines(&mut lines, &effective_runs, effective_ls, line_h, ctx.fonts);
        size_unsized_lines(
            &mut lines,
            para,
            &effective_runs,
            effective_ls,
            (line_h, font_size, para_ascent),
            ctx.fonts,
        );
        size_tab_mark_line(&mut lines, para, effective_ls, tallest_lhr, ctx.fonts);
    } else if grid_snapped {
        layout::size_grid_lines(
            &mut lines,
            ctx.fonts,
            effective_ls,
            line_h,
            sp.line_pitch,
            grid_baseline,
        );
    }
    // Only glyphs on a picture's own line hang a descent below it: when the
    // text after leading spaces and a picture wraps to the next line, Word
    // closes the picture line at the picture (leading-picture-image-fits).
    let picture_lines_glyphless = lines.iter().filter(|l| line_max_image_h(l) > 0.0).all(|l| {
        l.chunks
            .iter()
            .all(|c| c.inline_image_name.is_some() || c.text.trim().is_empty())
    });
    let picture_runs: Vec<Run> = if picture_lines_glyphless {
        effective_runs
            .iter()
            .filter(|r| r.inline_image.is_some() || r.text.trim().is_empty())
            .cloned()
            .collect()
    } else {
        Vec::new()
    };
    let mut para_metrics = (
        para_ascent,
        if max_inline_img_h > 0.0 {
            let runs = if picture_lines_glyphless {
                &picture_runs[..]
            } else {
                &effective_runs[..]
            };
            picture_line_bottom(runs, para, ctx.fonts, effective_ls)
        } else {
            0.0
        },
    );

    // Legacy Word puts the mark's extra automatic leading only after the
    // last picture line. Modern Word adds it after every picture line.
    let legacy_picture_leading = if ctx.compat_mode < 15
        && max_inline_img_h > 0.0
        && matches!(effective_ls, LineSpacing::Auto(factor) if factor > 1.0)
        && !effective_runs.is_empty()
        && effective_runs
            .iter()
            .all(|r| r.vanish || r.inline_image.is_some())
    {
        std::mem::replace(&mut para_metrics.1, 0.0)
    } else {
        0.0
    };
    // Whole, split across pages or beside a float, the lines are drawn the
    // same way; only the slice and its frame differ. `pb` is a parameter so
    // the callers keep their own split borrows of `state`.
    let draw = |pb: &mut PageBuilder,
                lines: &[TextLine],
                x: f32,
                w: f32,
                baseline: f32,
                first_idx: usize,
                geom: Option<&[(f32, f32)]>,
                total: usize,
                tag: usize,
                grid: Option<layout::LineGrid>| {
        render_paragraph_lines(
            &mut pb.content,
            lines,
            &para.alignment,
            x,
            para.indent_left,
            w,
            baseline,
            line_h,
            para_metrics,
            total,
            first_idx,
            &mut pb.links,
            text_hanging,
            ctx.fonts,
            geom,
            &mut pb.gradient_specs,
            Some(&mut pb.comment_anchors),
            ln_cfg.map(|(start, count_by, right_x)| LineNumberArg {
                counter: &mut pb.line_number_counter,
                start,
                count_by,
                right_x,
                font: ln_font.clone(),
            }),
            Some(LinkTagger::new(&mut pb.tags, pb.all_contents.len(), tag)),
            None,
            grid,
        );
    };
    let mut planned_gaps = step_lines_below_floats(
        &mut lines,
        &state.pb,
        (col_x, col_w),
        sp.page_height,
        first_line_top,
        line_h,
        para_metrics,
    );

    // One line of the mark's font at single spacing: a picture-only line's
    // baseline below its top (and the base its leading is measured from).
    let natural_line_h = font_size * tallest_lhr.unwrap_or(1.2);
    // A picture taller than the text line takes its paragraph's own
    // line-spacing leading below it, sized by the paragraph mark
    // (dental_amalgam: a 68.25pt logo under Normal's 1.15 lines is
    // 70.1pt tall in Word although the next paragraph is single-spaced).
    // The mark's, not the picture run's: Word probe at 3× spacing, 12pt
    // Times New Roman, 12pt Calibri, 24pt Times New Roman and 11pt Cambria
    // marks add 2 × their line whatever the picture run's font and size.
    // Lazy: most paragraphs hold no picture.
    let mark_line = || {
        let (mark_fs, mark_lhr) = unsized_line_metrics(para, font_size, ctx.fonts);
        let mark_line_h = resolve_line_h(effective_ls, mark_fs, mark_lhr);
        (mark_line_h, spacing_leading(mark_line_h, mark_fs, mark_lhr))
    };
    // A shorter one sits on the baseline of a full line of the mark's
    // font (indigenous_innovation's 2.25pt rule under the title). Taller
    // than a natural line is enough for the leading: a 30pt picture under
    // 3 × 13.8pt lines is 57.6pt in Word, not 41.4. Charts and an inline
    // shape alone on its line are pictures too (torture_test: Word's next
    // line is 2.09pt lower than the chart's bottom).
    // Markup view: an inserted picture keeps its run's descent below it like
    // the underlined text it is; a short one then sits on the text baseline
    // instead of the line's foot (Word for Mac probe p22d). An at-least
    // minimum is no leading: the line only grows to it (`picture_leading`).
    let inserted_descent = inserted_picture_descent(&effective_runs, ctx.fonts);
    let picture_line_h = |bottom: f32| match effective_ls {
        _ if bottom + inserted_descent <= natural_line_h => line_h,
        LineSpacing::AtLeast(min) => (bottom + inserted_descent).max(min),
        _ => bottom + inserted_descent + mark_line().1,
    };
    // Modern Word swaps the inline layout axes in the nearest quarter-turn
    // quadrant; it does not reserve the exact oblique bounding rectangle.
    let swap_axes = ctx.compat_mode >= 15
        && para.image.as_ref().is_some_and(|img| {
            (img.rotation_deg.rem_euclid(360.0) + 45.0).rem_euclid(180.0) >= 90.0
        });
    let mut content_h = if para.inline_chart.is_some() {
        picture_line_h(para.content_height)
    } else if let Some(img) = &para.image {
        let image_height = if swap_axes {
            img.display_width + img.layout_extra_height
        } else {
            para.content_height
        };
        // Under a lines grid a lone picture takes the cells its own height
        // needs, whatever its run's font: a Taiwanese press release's 17.8pt
        // logo in a Times New Roman 16 run (an 18.4pt line, two 18pt cells)
        // takes one cell in Word.
        let picture_h = if grid_snapped {
            layout::grid_cells(image_height + inserted_descent, sp.line_pitch)
        } else {
            picture_line_h(image_height)
        };
        // An OLE preview wider than the column leaves its line no room for the
        // paragraph mark, which wraps onto a line of its own (alfies_arc: an
        // 11pt line follows a 1014pt-wide OLE logo strip, annotation #186).
        // The paragraph's own indents do not count: learning_cultures keeps
        // the mark beside a column-wide picture in a right-indented paragraph.
        // Compared in whole twips: a picture sized to the column (8616.99
        // twips from its EMU extent against an 8617 twip column) fits.
        let twips = |pt: f32| (pt * 20.0).round();
        // A line break after the picture leaves the mark on a line of its own
        // too, sized by the mark's rPr like any trailing break's empty line
        // (a corpus logo paragraph: 7pt mark under a 16pt break run, Word's
        // next paragraph 8.54pt lower). ponytail: one line however many
        // breaks follow the picture; extra ones would be break-run lines.
        let trailing_break = effective_runs
            .iter()
            .rev()
            .find(|r| !r.text.is_empty() || r.is_line_break || r.inline_image.is_some())
            .is_some_and(|r| r.is_line_break);
        if trailing_break {
            picture_h
                + mark_line_h(para, effective_ls, tallest_lhr, ctx.fonts)
                    .unwrap_or_else(|| mark_line().0)
        } else if img.is_ole_preview && twips(img.layout_size().0) > twips(col_w) {
            picture_h + mark_line().0
        } else {
            picture_h
        }
    } else if max_inline_img_h > 0.0 {
        lines_height(&lines, line_h, para_metrics)
    } else if text_empty {
        if para.paragraph_mark_vanish {
            0.0
        } else if para.content_height > 0.0 {
            para.content_height
        } else {
            line_h + mark_position_stretch(para, effective_ls, grid_snapped, ctx.fonts)
        }
    } else {
        let num_lines = lines.len();
        // The numbering label is a run on the first line, so its font
        // metrics participate in that line's height.
        let first_line_h = label_boosted_line_h(
            para,
            ctx.fonts,
            line_h,
            effective_ls,
            font_size,
            tallest_lhr,
            tallest_ar,
        ) - lines
            .first()
            .and_then(|l| l.pitch)
            .map_or(0.0, |p| line_h - p);
        (if num_lines <= 1 {
            // If the single line was created by a break, use its font size
            if let Some(bfs) = lines.first().and_then(|l| l.break_font_size) {
                let blhr = break_run_lhr(&effective_runs, bfs, ctx.fonts);
                resolve_line_h(effective_ls, bfs, blhr)
            } else {
                first_line_h
            }
        } else {
            // Per-line height: break-created lines use the break
            // run's font metrics instead of the paragraph's text metrics.
            let mut h = first_line_h;
            let last = lines.len() - 1;
            for (i, line) in lines.iter().enumerate().skip(1) {
                if line.leading_tab_wrap
                    && let Some(pitch) = line.pitch
                {
                    h += pitch;
                } else if let Some(bfs) = line.break_font_size {
                    // The empty line a trailing break leaves holds only the
                    // paragraph mark, whose rPr sizes it — the break char
                    // sizes the line it terminates, not this one (samtale:
                    // 26pt br before a 12pt mark, annotation #121).
                    let mark_h = (i == last && line.chunks.is_empty())
                        .then(|| mark_line_h(para, effective_ls, tallest_lhr, ctx.fonts))
                        .flatten();
                    if let Some(mark_h) = mark_h {
                        h += mark_h;
                    } else {
                        let blhr = break_run_lhr(&effective_runs, bfs, ctx.fonts);
                        h += resolve_line_h(effective_ls, bfs, blhr);
                    }
                } else {
                    h += line.pitch.unwrap_or(line_h);
                }
            }
            h
        }) + planned_gaps
    };

    if legacy_picture_leading > 0.0
        && lines
            .last()
            .is_some_and(|line| line_max_image_h(line) > para_ascent)
    {
        content_h += legacy_picture_leading;
    }

    // The shading covers the lines, not the room reserved below them for a
    // float: indonesian_school's "Format 12" box shows between its anchor's
    // white shading and the next paragraph's.
    let mut lines_h = content_h;

    // Extra height from floating images that extends beyond
    // the text content — used only for page-break decisions,
    // not for cursor advancement (text wraps beside the image).
    let mut float_overflow_h = 0.0f32;
    // How far a topAndBottom float over the first line pushes the text down.
    let mut text_drop = float_drop;
    content_h += float_drop;

    // A band keeps the lines out by itself (`step_below_bands`).
    for fi in para
        .floating_images
        .iter()
        .filter(|fi| wraps_here(fi) && !image_blocks_band(fi, sp, col_x, col_w, text_width))
    {
        let reserve = match fi.wrap_type {
            WrapType::TopAndBottom => true,
            // wrapSquare/Tight with a usable side strip: content (even empty
            // spacer paragraphs) flows beside via the float zone — reserving
            // here would double-count the image height (brazilian, ~42pt
            // strips). With no usable strip (sample500kB: image width ==
            // column width) Word stacks everything below the float, which
            // reserving the image height in the anchor reproduces.
            WrapType::Square | WrapType::Tight => no_strip_beside(fi, sp, col_x, col_w, text_width),
            WrapType::Through => false,
            WrapType::None => false,
        };
        let fi_h = match fi.v_position {
            VerticalPosition::Offset(o) => {
                o + fi.dist_top + fi.image.display_height + fi.dist_bottom
            }
            _ => fi.dist_top + fi.image.display_height + fi.dist_bottom,
        };
        if reserve && !text_empty && leading_full_width_picture(fi, sp, col_x, col_w, text_width) {
            float_overflow_h = float_overflow_h.max(fi_h);
        } else if reserve {
            // Wide images block all text — add to content_h. An empty anchor
            // paragraph's own line moves below such a float when the float
            // covers it: cyprus' page-wide letterhead (margin-relative, top
            // 83.6pt above the margin) keeps the mark's 14.65pt line under its
            // bottom, and Word probes with paragraph- and margin-relative
            // offsets, square and topAndBottom wrapping all do the same.
            let content_top = state.pb.slot_top - inter_gap;
            let zone = FloatZone::for_float(
                fi,
                resolve_fi_x(fi, sp, col_x, col_w, text_width),
                pre_anchor.fi_y_top(fi, sp),
            );
            // ponytail: only a float that starts at or below the paragraph top,
            // or a paragraph opening the page (the probed cases); a float
            // reaching over earlier lines on the page (sample500kB) keeps the
            // plain reserve.
            let covers_line = zone.top_y > content_top - content_h
                && zone.bottom_y < content_top
                && (zone.top_y <= state.pb.slot_top + 0.5 || state.pb.is_at_page_top(sp));
            // A topAndBottom picture over a text paragraph's first line moves
            // its lines below it (Word probe: offset 0 puts the first baseline
            // under the picture; +30pt leaves the first line above it).
            let covers_first_line = matches!(fi.wrap_type, WrapType::TopAndBottom)
                && zone.top_y > content_top - line_h
                && zone.bottom_y < content_top
                && (zone.top_y <= state.pb.slot_top + 0.5 || state.pb.is_at_page_top(sp));
            if text_empty && !para.paragraph_mark_vanish && covers_line {
                content_h = content_h.max(content_top - zone.bottom_y + content_h);
            } else if !text_empty && covers_first_line {
                text_drop = text_drop.max(content_top - zone.bottom_y);
                content_h = content_h.max(fi_h).max(lines_h + text_drop);
            } else if fi.v_relative_from == VRelativeFrom::Paragraph {
                content_h = content_h.max(fi_h);
            } else {
                // A page- or margin-placed float's offset is not from this
                // paragraph: it reaches down to the float's own bottom (a
                // 45pt rubric banner at 484.8pt from the page top made its
                // one-line anchor 530pt tall and sent it to the next page).
                // The space after overlaps the float: Word starts the next
                // paragraph at the banner's foot. Only a float the next line
                // would run into counts; one further down the page leaves the
                // text flowing (a page-foot logo 680pt below its anchor).
                // ponytail: lines that reach a lower float later are not moved.
                let next_line_bottom = content_top - content_h - para.space_after - line_h;
                if zone.top_y > next_line_bottom {
                    content_h = content_h.max(content_top - zone.bottom_y - para.space_after);
                }
            }
        } else if fi.v_relative_from == VRelativeFrom::Paragraph
            && fi.wrap_type.wraps_beside()
            && fi_h <= sp.page_height - sp.margin_top - sp.margin_bottom
        {
            // Paragraph-relative wrapping images: track overflow
            // for page-break check only (text wraps beside them). One taller
            // than the page body fits no page: it stays, moved up to the
            // page top (`ParaAnchor::fi_y_top`).
            float_overflow_h = float_overflow_h.max(fi_h);
        }
    }

    for tb in &para.textboxes {
        // A band keeps the lines out by itself (`step_below_bands`).
        let reserve = tb.v_relative_from == VRelativeFrom::Paragraph
            && !hangs_below_anchor(tb, para, col_w, ctx)
            && match tb.wrap_type {
                WrapType::TopAndBottom => true,
                WrapType::Square => textbox_room_beside(tb, para, sp, (col_x, col_w), text_width)
                    .is_some_and(|room| room < MIN_EMPTY_STRIP),
                _ => false,
            };
        if reserve {
            let tb_bottom =
                tb.v_offset_pt + textbox_render::textbox_height(tb, ctx) + tb.dist_bottom;
            match tb.v_relative_from {
                // The offset counts from above the paragraph's own space
                // before, which already sits above the content. An empty
                // anchor's own line that would overlap the box moves below it,
                // keeping that space (indonesian_school's "Format 12" with
                // before=200 twips: Word +22.5pt = 10 + box bottom 33.18 +
                // line - 33.18).
                // Drawings on the line count from it (isla_language's lesson
                // canvas).
                // An inline shape alone on its line is a picture there.
                _ if tb.on_line => {
                    let h = if text_empty {
                        picture_line_h(tb_bottom)
                    } else {
                        tb_bottom
                    };
                    content_h = content_h.max(h);
                }
                _ => {
                    let before = own_space_before(inter_gap, state.prev_space_after);
                    let line_overlaps = tb.v_offset_pt < before + lines_h && tb_bottom > before;
                    // Only a mark line moves below the box: a picture-only
                    // paragraph keeps its picture line beside it (a 65pt logo
                    // beside a square-wrapped title box ends the paragraph at
                    // the logo's foot in Word, not the two stacked).
                    // Side-by-side boxes move the line below the lower one, not
                    // below their stacked heights (torture_test's linked pair).
                    // A box over a text paragraph's first line moves its lines
                    // below it, as a picture does (Word examples: topAndBottom,
                    // and square with a 17pt strip beside it).
                    let first_line_overlaps =
                        tb.v_offset_pt < before + line_h && tb_bottom > before;
                    content_h = if text_empty && para.image.is_none() && line_overlaps {
                        content_h.max(tb_bottom + lines_h)
                    } else if !text_empty && first_line_overlaps {
                        text_drop = text_drop.max(tb_bottom - before);
                        content_h.max(lines_h + text_drop)
                    } else {
                        content_h.max(tb_bottom - before)
                    };
                }
            }
        } else if tb.v_relative_from == VRelativeFrom::Paragraph
            && textbox_wraps_beside(tb, para, sp, (col_x, col_w), text_width)
        {
            // Like a picture text wraps beside: only the page break counts
            // it (Word example: an empty anchor 40pt above the bottom margin
            // goes to the next page with its 180pt box).
            let offset = match tb.v_position {
                VerticalPosition::Offset(o) => o,
                _ => 0.0,
            };
            let h = textbox_render::textbox_height(tb, ctx);
            let tb_h = offset + tb.dist_top + h + tb.dist_bottom;
            if tb_h <= sp.page_height - sp.margin_top - sp.margin_bottom {
                float_overflow_h = float_overflow_h.max(tb_h);
            }
        }
    }
    // Inline shapes alone on an at-least line sit at its foot, like pictures.
    let on_line_drop = if text_empty && matches!(effective_ls, LineSpacing::AtLeast(_)) {
        para.textboxes
            .iter()
            .filter(|tb| tb.on_line && tb.v_relative_from == VRelativeFrom::Paragraph)
            .map(crate::model::Textbox::flow_bottom)
            .reduce(f32::max)
            .map_or(0.0, |bottom| (picture_line_h(bottom) - bottom).max(0.0))
    } else {
        0.0
    };

    // Vanished paragraph mark: zero out height and spacing
    if text_empty && para.paragraph_mark_vanish {
        content_h = 0.0;
        inter_gap = 0.0;
    }

    // Word treats consecutive paragraphs with identical borders and matching
    // box edges as one border group: top padding/rule only on the first, bottom
    // padding/rule only on the last, and a `between` rule (if any) at the joins.
    let prev_borders_match = prev_para.is_some_and(|pp| joins_border_group(pp, para));
    let next_borders_match = next_para.is_some_and(|np| joins_border_group(para, np));
    let bottom_collapses = next_borders_match;

    // The top border's band sits inside the paragraph, like the bottom one
    // (`bdr_bottom_extent`): case17's 0.5pt and 1pt boxes start their band at
    // the paragraph top and their text `space` below its lower edge.
    let bdr_top_pad = if prev_borders_match {
        0.0
    } else {
        border_band(para.borders.top.as_ref())
    };
    let bdr_top_half_band = if prev_borders_match {
        0.0
    } else {
        para.borders.top.as_ref().map_or(0.0, |b| b.extent() / 2.0)
    };
    let (bottom_bdr, bottom_band) = bottom_border(para, next_para.filter(|_| bottom_collapses));
    let bdr_bottom_pad = bottom_band.0 + bottom_bdr.map_or(0.0, |b| b.extent() / 2.0);
    // Full extent of bottom border below content (to border bottom edge)
    let bdr_bottom_extent = bottom_band.1;

    // Word measures the bottom border `space` attribute from
    // the full line-height content bottom, not from the text
    // descent.  No trailing-lead adjustment is needed.

    let needed = inter_gap + bdr_top_pad + content_h + bdr_bottom_extent;
    // For page-break decisions, also account for floating
    // images that extend below the text content.
    let needed_with_floats = needed.max(inter_gap + float_overflow_h);
    let at_page_top = state.pb.is_at_page_top(sp);

    // Pre-compute footnote space for this paragraph so the
    // page-break check accounts for footnotes the paragraph
    // introduces (otherwise they're only tracked after
    // rendering, which can cause body/footnote overlap).
    let line_refs: Vec<Vec<u32>> = lines
        .iter()
        .map(|l| l.chunks.iter().filter_map(|c| c.footnote_id).collect())
        .collect();
    let run_refs: Vec<u32> = para.runs.iter().filter_map(|r| r.footnote_id).collect();
    let (mut line_fn_extra, para_fn_extra) =
        page_footnote_extra(state, doc, ctx, col_geometry, &line_refs, &run_refs);

    // Word allows the last line's trailing inter-line
    // spacing to extend past the bottom margin — only the
    // text (ascent + descent) must fit inside the content
    // area.  Compute the excess leading so the page-break
    // check can tolerate it.
    // A lone paragraph mark is a last line too: Word keeps an empty 1.15-spaced
    // Calibri paragraph whose leading overhangs the margin by 0.8pt
    // (sao_paulo_procurement_contract p2) and an empty 1.5-spaced Arial one
    // (czech_wastewater_discharge_permit p1).
    let mark_only = text_empty && !para.paragraph_mark_vanish && para.content_height <= 0.0;
    // Not into a footnote area, though, as in the per-line check below
    // (zimbabwe_gold p3: Word moves a double-spaced paragraph whose last line
    // would end 0.6pt into it; examples: no tolerance). A line an at-least
    // minimum sets keeps its surplus above the text and pictures, so it has
    // none to hang (Word examples: 10pt Arial under 24pt needs all 24pt).
    let min_sets_line = matches!(effective_ls, LineSpacing::AtLeast(min) if min >= line_h);
    let last_line_lead = if (!lines.is_empty() || mark_only)
        && state.pb.col_fn_reserved == 0.0
        && para_fn_extra == 0.0
        && para.image.is_none()
        && para.inline_chart.is_none()
        && para.smartart.is_empty()
        && !matches!(effective_ls, LineSpacing::Exact(_))
        && !min_sets_line
    {
        let single_h = tallest_lhr
            .map(|r| font_size * r)
            .unwrap_or(font_size * 1.2);
        (line_h - single_h).max(0.0)
    } else {
        0.0
    };

    let keep_next_extra = if para.keep_next && !at_page_top {
        let mut extra = 0.0;
        // Footnotes the kept paragraphs bring come along too
        // (uk_commercial_lease's "Break Date" moves with its definition's note 8).
        let mut chain_notes: Vec<u32> = Vec::new();
        let mut prev_sa = effective_space_after;
        let mut first_link: Option<f32> = None;
        let mut i = block_idx + 1;
        loop {
            let next = match section_blocks.get(i) {
                Some(Block::Paragraph(p)) => p,
                // The chain reaches into a following in-flow table: the paragraph
                // stays with its first row (czech_wastewater's "8. Seznam…"
                // heading moves to page 4 with its table).
                Some(Block::Table(t)) if t.position.is_none() => {
                    extra += prev_sa + first_row_height(t, &ctx, col_geometry[state.current_col].1);
                    break;
                }
                _ => break,
            };
            if next.page_break_before {
                // A page break opening the next paragraph, a `<w:br>` or
                // pageBreakBefore, ends the chain where it is: Word keeps the
                // kept paragraphs on this page (online export probes,
                // 2026-10-03; run-borders' items before a pageBreakBefore
                // Heading 3).
                break;
            }
            chain_notes.extend(next.runs.iter().filter_map(|r| r.footnote_id));
            let (nfs, nlhr, _) = tallest_run_metrics(&next.runs, ctx.fonts);
            let next_inter = f32::max(prev_sa, next.space_before);
            let next_ls = next.line_spacing.unwrap_or(ctx.doc_line_spacing);
            // An exact or at-least box is the whole first line (issue #145's
            // heading moves with its 40pt-exact successor); an auto multiple's
            // extra leading may hang past the margin like `last_line_lead`.
            let first_ls = match next_ls {
                LineSpacing::Auto(_) => LineSpacing::Auto(1.0),
                boxed => boxed,
            };
            // A picture makes its line as tall as itself (issue #146: a
            // heading kept with an 80pt picture under a 14pt line moves with
            // it); a paragraph of one picture holds it in `image`, not a run.
            let picture_h = if next.image.is_some() {
                next.content_height
            } else {
                layout::runs_max_image_h(&next.runs)
            };
            let next_first_line_h = line_h_with_picture(
                picture_h,
                &next.runs,
                next,
                ctx.fonts,
                next_ls,
                resolve_line_h(first_ls, nfs, nlhr),
            );
            // A frame in the flow comes along whole; a long chain keeps only
            // its first line (Word examples, fixture frame-keep-with-next: a
            // heading kept with a 100pt exact frame moves with it; two kept
            // headings before a frame taller than the page stay together).
            if let Some(fp) = flow_frame_of(next) {
                first_link.get_or_insert(extra + next_inter + next_first_line_h);
                let col_w = col_geometry[state.current_col].1;
                let blocks = &section_blocks[i..];
                extra +=
                    next_inter + frame_box_h(fp, blocks, ctx, frame_width(fp, blocks, ctx, col_w));
                break;
            }
            // The chain needs as many of the next paragraph's lines as must
            // stay together on this page: all of a keepLines paragraph; else
            // one without widow control, with it two, or all of a paragraph
            // of three or fewer (it can't split without leaving a lone line).
            // lithuanian's headings end on an empty paragraph and fit at the
            // foot; western_australia's end on a three-line item and move.
            // A kept paragraph that stays whole passes the chain on to its
            // successor; one that may split ends it (australian_higher's
            // keepLines items move with heading 6.5.1).
            let mut n = None;
            let mut count = || {
                *n.get_or_insert_with(|| line_count(next, ctx, col_geometry[state.current_col].1))
            };
            let needed = if next.keep_lines {
                count()
            } else {
                lines_kept_together(next.widow_control, &mut count)
            };
            extra += next_inter
                + next_first_line_h
                + needed.saturating_sub(1) as f32 * resolve_line_h(next_ls, nfs, nlhr);
            first_link.get_or_insert(extra);
            if !next.keep_next || needed < count() {
                break;
            }
            if next.page_break_after {
                extra = f32::MAX;
                break;
            }
            prev_sa = next.space_after;
            i += 1;
        }
        let opens_note_area = state.pb.col_fn_reserved == 0.0 && para_fn_extra == 0.0;
        let (_, notes_h) = per_line_footnote_extra(
            &[],
            &chain_notes,
            &state.pb.footnote_ids_set,
            if opens_note_area {
                ctx.note_separator.height
            } else {
                0.0
            },
            |id| footnote_height(id, &doc.footnotes, ctx, col_geometry[state.current_col].1),
        );
        // A chain no page can hold starts on a new page and flows from
        // there: its later paragraphs keep only their own link to the next,
        // or each would strand on a page of its own (a bibliography of
        // Heading 3 entries opens page 3 in Word, the rest follows, each
        // entry still with its successor's first lines).
        let page_h = effective_slot_top(sp, false, state.pb.page_num(ctx.sections), ctx)
            - state.effective_margin_bottom;
        let inside_chain = block_idx
            .checked_sub(1)
            .and_then(|i| section_blocks.get(i))
            .is_some_and(|b| matches!(b, Block::Paragraph(p) if p.keep_next));
        // Judged once, where the chain starts: from a later member the rest
        // may fit a page and would move again.
        if !inside_chain {
            state.long_keep_chain = needed_with_floats + extra + notes_h > page_h;
        }
        if inside_chain && state.long_keep_chain {
            first_link.unwrap_or(extra)
        } else {
            extra + notes_h
        }
    } else {
        0.0
    };

    // A multi-line paragraph that opens the page and outruns it splits like
    // any other: one of 108 manual line breaks, double spaced, took one page
    // instead of five (corpus 2026-11-05; probe p18b, 70 break lines). A
    // single oversized line stays put (case150's 700pt exact spacer).
    // A paragraph moved for want of a fitting line splits on its new page
    // too (Word example: eight 10pt lines under 24pt after a 50pt filler on a
    // 70pt body move and take two a page).
    let first_line_h = tallest_lhr
        .map(|ratio| font_size * ratio)
        .unwrap_or(font_size);
    // Each line must fit together with the footnotes it introduces.
    let fit_in = |lines: &[TextLine], fn_extra: &[f32], room: f32, notes_on_page: bool| {
        if line_h > 0.0 {
            let whole_lines =
                notes_on_page || min_sets_line || matches!(effective_ls, LineSpacing::Exact(_));
            lines_that_fit(
                lines,
                fn_extra,
                room,
                whole_lines,
                line_h,
                first_line_h,
                para_metrics,
            )
        } else {
            0
        }
    };
    // A frame in the flow never splits: `FlowFrame` has placed it, and it
    // runs past the bottom if it must (Word example, fixture
    // frame-moves-shared-tall: the next block opens the next page).
    let in_flow_frame = flow_frame_of(para).is_some();
    let mut moved = false;
    while !in_flow_frame
        && (!at_page_top || lines.len() > 1)
        && (moved
            || state.pb.slot_top - needed_with_floats - keep_next_extra + last_line_lead
                < state.effective_margin_bottom + para_fn_extra)
    {
        if moved {
            let col = col_geometry[state.current_col];
            let rest = &section_blocks[block_idx..];
            let page = (sect_idx, state.pb.page_count());
            if state.pb.bands_prescanned != Some(page) {
                state.pb.bands_prescanned = Some(page);
                prescan_page_bands(state, rest, sp, col, text_width, ctx);
            }
            let first = lines.first().map_or(line_h, |l| {
                layout::line_own_advance(l, line_h, para_metrics)
            });
            place_moved_first_line(
                state,
                first,
                0.0,
                rest,
                sect_idx,
                sp,
                col_geometry,
                text_width,
                ctx,
            );
            step_below_bands(
                state,
                block_idx,
                &section_blocks[block_idx],
                sp,
                col,
                text_width,
                Some(first),
                ctx,
            );
            let gaps = step_lines_below_floats(
                &mut lines,
                &state.pb,
                col_geometry[state.current_col],
                sp.page_height,
                state.pb.slot_top,
                line_h,
                para_metrics,
            );
            content_h += gaps - planned_gaps;
            lines_h += gaps - planned_gaps;
            planned_gaps = gaps;
        }
        let available = state.pb.slot_top - inter_gap - text_drop - state.effective_margin_bottom;
        let mut lines_that_fit = fit_in(
            &lines,
            &line_fn_extra,
            available,
            state.pb.col_fn_reserved > 0.0,
        );
        if moved && lines_that_fit >= lines.len() {
            break;
        }

        let text_would_split = lines_h > available;
        if para.widow_control {
            // Ensure at least 2 lines remain on next page (orphan prevention)
            if lines_that_fit > 0 && lines.len().saturating_sub(lines_that_fit) < 2 {
                lines_that_fit = lines.len().saturating_sub(2);
            }
        }

        // keepLines: don't split — move entire paragraph to next column/page
        let leading_picture_keeps_text = content_h
            <= sp.page_height - sp.margin_top - sp.margin_bottom
            && text_would_split
            && para
                .floating_images
                .iter()
                .any(|fi| leading_full_width_picture(fi, sp, col_x, col_w, text_width));
        if para.keep_lines || leading_picture_keeps_text {
            lines_that_fit = 0;
        }

        let min_split_lines = if para.widow_control { 2 } else { 1 };
        if lines_that_fit >= min_split_lines && lines_that_fit < lines.len() {
            let first_part = &lines[..lines_that_fit];
            state.pb.slot_top -= inter_gap;
            layout::set_change_bar_extension(
                own_space_before(inter_gap, state.prev_space_after),
                effective_space_after,
            );
            let split_anchor = ParaAnchor::new(
                state.pb.take_float_anchor_for(block_idx),
                state.pb.anchor_slot_top(sp, ctx),
                own_space_before(inter_gap, state.prev_space_after),
                sp,
                ctx,
            );
            let baseline_offset = if grid_snapped {
                grid_baseline
            } else {
                first_baseline_offset
            };
            let baseline_y = state.pb.slot_top - baseline_offset - text_drop;

            // Its text boxes hang from its top on this page, as an unsplit
            // paragraph's do.
            for tb in &para.textboxes {
                render_para_textbox(
                    tb,
                    para,
                    split_anchor,
                    sp,
                    (col_x, col_w),
                    text_width,
                    ctx,
                    &mut state.pb,
                );
            }

            // One element for both halves: its content continues on the next page.
            let tags = state.pb.para_tags(para, doc);
            let tag = tags.1;
            let grid = line_grid(baseline_offset, state.pb.all_contents.len());
            let label_y =
                layout::first_drawn_baseline(baseline_y, first_part, para_ascent, grid, false);
            state.pb.begin_para_tags(tags, |content| {
                render_list_label(content, para, ctx.fonts, label_x, label_y, font_size)
            });

            draw(
                &mut state.pb,
                first_part,
                para_text_x,
                para_text_width,
                baseline_y,
                0,
                poly_line_geom.as_deref(),
                lines.len(),
                tag,
                grid,
            );
            state.pb.end_tag();

            // Floats anchored before the paragraph's last text hang from its
            // top on this page, like an unsplit paragraph's (the split path
            // returns before the ordinary float rendering below).
            let (leading, leading_names, leading_effects) = float_subset(
                &para.floating_images,
                &images.floating_image_pdf_names,
                &images.effect_floating_names,
                state.global_block_idx,
                |fi| !blank_after(&para.runs, fi.anchor_run),
            );
            if !leading.is_empty() {
                let page = state.pb.all_contents.len();
                render_floating_images(
                    &leading,
                    state.global_block_idx,
                    &leading_names,
                    &leading_effects,
                    sp,
                    col_x,
                    col_w,
                    text_width,
                    split_anchor,
                    &mut state.pb.behind_shapes,
                    &mut state.pb.deferred_shapes,
                    &mut state.pb.tags,
                    page,
                );
            }

            // Footnotes referenced on the lines that stay here belong to this
            // page's footnote area; the flush below would otherwise carry them
            // to the continuation page while the space stays reserved here.
            for id in line_footnote_ids(first_part) {
                track_page_footnote(state, doc, ctx, col_geometry[state.current_col], id);
            }

            // The column ends below the lines that stay (its separator reaches them).
            state.pb.slot_top -= text_drop + lines_height(first_part, line_h, para_metrics);
            state.advance_column_or_page(col_count, sect_idx, sp, ctx);

            let baseline_offset2 = if grid_snapped {
                grid_baseline
            } else {
                para_ascent * auto_ascent_scale(effective_ls)
            };
            // The rest may itself outrun the column: a long paragraph spans as
            // many columns or pages as it needs (case80 probes: one paragraph
            // balanced over three columns).
            // The trailing pictures of a split paragraph belong to its
            // final fragment: a float goes on the page its anchor run's
            // line lands on, hung from the paragraph's top on that page
            // (an external tracked corpus cover: a deleted page-wide
            // picture after five pages of inserted text prints on the
            // fifth page). Full-width pictures also keep following visible
            // text below them, while intervening empty marks may occupy
            // the space above.
            let (trailing, trailing_names, trailing_effects) = float_subset(
                &para.floating_images,
                &images.floating_image_pdf_names,
                &images.effect_floating_names,
                state.global_block_idx,
                |fi| blank_after(&para.runs, fi.anchor_run),
            );
            let unlined_refs: Vec<u32> = run_refs
                .iter()
                .copied()
                .filter(|id| !line_refs.iter().flatten().any(|l| l == id))
                .collect();
            let mut start = lines_that_fit;
            loop {
                let line = &lines[start];
                let first = layout::line_own_advance(line, line_h, para_metrics);
                let rest = &section_blocks[block_idx + 1..];
                let at_foot = place_first_line(
                    state,
                    first,
                    rest,
                    sect_idx,
                    sp,
                    col_geometry,
                    text_width,
                    ctx,
                );
                step_lines_below_floats(
                    &mut lines[start..],
                    &state.pb,
                    col_geometry[state.current_col],
                    sp.page_height,
                    state.pb.slot_top,
                    line_h,
                    para_metrics,
                );
                let remaining = &lines[start..];
                let (fn_extra, _) = page_footnote_extra(
                    state,
                    doc,
                    ctx,
                    col_geometry,
                    &line_refs[start..],
                    &unlined_refs,
                );
                let mut fit = fit_in(
                    remaining,
                    &fn_extra,
                    state.pb.slot_top - state.effective_margin_bottom,
                    state.pb.col_fn_reserved > 0.0,
                );
                if para.widow_control && fit < remaining.len() {
                    if remaining.len() - fit < 2 {
                        fit = remaining.len().saturating_sub(2);
                    }
                    // Widow control keeps two lines together here too; with
                    // room for fewer, the column is skipped.
                    if fit < 2 {
                        fit = 0;
                    }
                }
                let fresh_page = at_foot || state.pb.takes_a_line(sp);
                if fit == 0 && !fresh_page {
                    state.advance_column_or_page(col_count, sect_idx, sp, ctx);
                    continue;
                }
                let chunk = &remaining[..fit.clamp(1, remaining.len())];
                let baseline_y2 = state.pb.slot_top - baseline_offset2;
                let (rest_col_x, rest_col_w) = col_geometry[state.current_col];
                let rest_text_x = rest_col_x + para.indent_left;
                let rest_text_width = (rest_col_w - para.indent_left - para.indent_right).max(1.0);

                let last_fragment = fit >= remaining.len() && !trailing.is_empty();
                let anchor = ParaAnchor::new(None, state.pb.slot_top, 0.0, sp, ctx);
                // Where the fragment's lines go around the first wrapping
                // trailing float (Word probes p20c/p20e: lines whose box meets
                // the picture's band sit beside it, x 249 by a 150pt picture
                // at 90; the corpus cover's page-wide picture leaves no side,
                // its lines go below, and when below is off the page the
                // lines move to the next page, one line staying per page, and
                // the anchor line alone sits at the page bottom above the
                // footer: the corpus cover's "Montering", an empty line and
                // "Med forbehold" take one page each).
                let mut plan: Option<FragmentPlan> = None;
                if last_fragment
                    && let Some(fi) = trailing.iter().find(|fi| fi.wrap_type.wraps_beside())
                {
                    let fi_x = resolve_fi_x(fi, sp, rest_col_x, rest_col_w, text_width);
                    let zone = FloatZone::for_float(fi, fi_x, anchor.fi_y_top(fi, sp));
                    let mut y = state.pb.slot_top;
                    let mut first_in = None;
                    let mut after = chunk.len();
                    for (i, l) in chunk.iter().enumerate() {
                        let adv = lines_height(std::slice::from_ref(l), line_h, para_metrics);
                        let in_band = y > zone.bottom_y && y - adv < zone.top_y;
                        if in_band && first_in.is_none() {
                            first_in = Some((i, y));
                        }
                        if !in_band && first_in.is_some() && after == chunk.len() {
                            after = i;
                        }
                        y -= adv;
                    }
                    if let Some((from, y_b)) = first_in {
                        let (mut tx, mut tw, mut lx) = (rest_text_x, rest_text_width, 0.0);
                        // The line's box meets the band though its top may sit above it.
                        let legacy = ctx.compat_mode < 15;
                        zone.narrow_paragraph(
                            y_b.min(zone.top_y),
                            rest_col_x,
                            rest_col_w,
                            para,
                            &mut tx,
                            &mut tw,
                            &mut lx,
                            legacy,
                        );
                        plan = Some(if tw < rest_text_width - 0.5 {
                            FragmentPlan::Beside {
                                from,
                                to: after,
                                text_x: tx,
                                text_w: tw,
                            }
                        } else if zone.bottom_y - lines_height(&chunk[from..], line_h, para_metrics)
                            >= state.effective_margin_bottom
                        {
                            FragmentPlan::Below {
                                from,
                                top: zone.bottom_y,
                            }
                        } else if chunk.len() > from.max(1) {
                            FragmentPlan::NextPage { keep: from.max(1) }
                        } else {
                            FragmentPlan::PageBottom
                        });
                    }
                }
                if let Some(FragmentPlan::NextPage { keep }) = plan {
                    fit = keep;
                }
                let chunk = &remaining[..fit.clamp(1, remaining.len())];
                let last_fragment = last_fragment && fit >= remaining.len();
                if last_fragment {
                    let page = state.pb.all_contents.len();
                    render_floating_images(
                        &trailing,
                        state.global_block_idx,
                        &trailing_names,
                        &trailing_effects,
                        sp,
                        rest_col_x,
                        rest_col_w,
                        text_width,
                        anchor,
                        &mut state.pb.behind_shapes,
                        &mut state.pb.deferred_shapes,
                        &mut state.pb.tags,
                        page,
                    );
                }
                // Full-width pictures keep following visible text below them,
                // while intervening empty marks may occupy the space above.
                let trailing_images_fit = last_fragment
                    && trailing.iter().all(|fi| {
                        fi.wrap_type.wraps_beside()
                            && offset_picture_without_strip(
                                fi, sp, rest_col_x, rest_col_w, text_width,
                            )
                            && resolve_fi_y_top(fi, sp, state.pb.slot_top)
                                - fi.image.display_height
                                - fi.dist_bottom
                                >= state.effective_margin_bottom
                    });
                if trailing_images_fit {
                    for fi in &trailing {
                        let x = resolve_fi_x(fi, sp, rest_col_x, rest_col_w, text_width);
                        let top = resolve_fi_y_top(
                            fi,
                            sp,
                            anchor.drawing_top(fi.on_line, fi.v_from_line),
                        );
                        state.pb.float_zone = Some(FloatZone::for_float(fi, x, top));
                        let bottom = top - fi.image.display_height - fi.dist_bottom;
                        state.pb.lower_next_text_to(bottom);
                    }
                }
                // (from, to, text_x, text_w, baseline of the segment's first line)
                let mut segments: Vec<(usize, usize, f32, f32, f32)> = Vec::new();
                let mut extra_drop = 0.0f32;
                // (`lines_height` of no lines is one pitch.)
                let seg_h = |a: usize, b: usize| {
                    if a < b {
                        lines_height(&chunk[a..b], line_h, para_metrics)
                    } else {
                        0.0
                    }
                };
                match plan {
                    Some(FragmentPlan::Beside {
                        from,
                        to,
                        text_x,
                        text_w,
                    }) if last_fragment => {
                        let h_a = seg_h(0, from);
                        let h_b = seg_h(from, to);
                        segments.push((0, from, rest_text_x, rest_text_width, baseline_y2));
                        segments.push((from, to, text_x, text_w, baseline_y2 - h_a));
                        segments.push((
                            to,
                            chunk.len(),
                            rest_text_x,
                            rest_text_width,
                            baseline_y2 - h_a - h_b,
                        ));
                    }
                    Some(FragmentPlan::Below { from, top }) if last_fragment => {
                        let h_a = seg_h(0, from);
                        segments.push((0, from, rest_text_x, rest_text_width, baseline_y2));
                        segments.push((
                            from,
                            chunk.len(),
                            rest_text_x,
                            rest_text_width,
                            top - baseline_offset2,
                        ));
                        extra_drop = (state.pb.slot_top - h_a) - top;
                    }
                    Some(FragmentPlan::PageBottom) if last_fragment => {
                        let top = state.pb.footer_top(sp, ctx)
                            + lines_height(chunk, line_h, para_metrics);
                        segments.push((
                            0,
                            chunk.len(),
                            rest_text_x,
                            rest_text_width,
                            top - baseline_offset2,
                        ));
                        extra_drop = state.pb.slot_top - top;
                    }
                    _ => segments.push((0, chunk.len(), rest_text_x, rest_text_width, baseline_y2)),
                }
                state.pb.begin_tag(tag);
                for &(from, to, seg_x, seg_w, seg_baseline) in segments.iter().filter(|s| s.1 > s.0)
                {
                    let grid = line_grid(baseline_offset2, state.pb.all_contents.len());
                    draw(
                        &mut state.pb,
                        &chunk[from..to],
                        seg_x,
                        seg_w,
                        seg_baseline,
                        start + from,
                        None,
                        lines.len(),
                        tag,
                        grid,
                    );
                }
                state.pb.end_tag();

                state.pb.slot_top -= lines_height(chunk, line_h, para_metrics) + extra_drop;
                start += chunk.len();
                let last_chunk = start >= lines.len();
                let mut ids = line_footnote_ids(chunk);
                if last_chunk {
                    ids.extend(&unlined_refs);
                }
                for id in ids {
                    track_page_footnote(state, doc, ctx, col_geometry[state.current_col], id);
                }
                if last_chunk {
                    break;
                }
                state.advance_column_or_page(col_count, sect_idx, sp, ctx);
            }
            state.prev_space_after = effective_space_after;
            layout::set_change_bar_extension(0.0, 0.0);

            for id in para.runs.iter().filter_map(|r| r.endnote_id) {
                state.pb.track_endnote(id);
            }

            state.global_block_idx += 1;
            state.take_page_break_after(para, sect_idx, sp, ctx);
            return true;
        }

        if moved {
            break;
        }
        let column = col_geometry[state.current_col];
        let page = state.pb.page_count();
        state.pb.overflow_column_or_page(
            &mut state.current_col,
            col_count,
            sect_idx,
            sp,
            &mut state.effective_margin_bottom,
            ctx,
            state.prev_space_after,
        );
        inter_gap = 0.0;
        // Its own floats move to the new column with it, an earlier anchor's
        // stay behind: lay it out again there, its gap the overflow's (Word
        // examples: a 100pt box and its anchor's text start 280pt right in the
        // second of two 200pt columns, at its top after a 24pt space after; a
        // paragraph beside an earlier box moved on by widow control takes the
        // next column's or page's full measure).
        let new_page = state.pb.page_count() != page;
        let earlier_floats = poly_line_geom.is_some() && !own_zone_set;
        if (own_zone_set || earlier_floats)
            && ((col_geometry[state.current_col].0 - col_x).abs() > 0.01
                || (new_page && earlier_floats))
            && !para.page_break_before
            && !para.column_break_before
            && !leading_table_clear
        {
            if !new_page {
                (state.pb.float_zone, state.pb.more_float_zones) = entry_zones;
            }
            state.reflowed = true;
            return render_paragraph_block(
                para,
                state,
                ctx,
                sp,
                col_geometry,
                col_count,
                text_width,
                sect_idx,
                block_idx,
                section_blocks,
                images,
                doc,
                smartart_font_key,
                debug_wrap,
                table_connector_anchor,
            );
        }
        moved = true;
        if col_geometry[state.current_col] != column {
            break;
        }
        line_fn_extra = page_footnote_extra(state, doc, ctx, col_geometry, &line_refs, &run_refs).0;
    }

    // Suppress space_before at the top of a page
    if let Some(gap) = state
        .pb
        .page_top_gap(sp, effective_space_before, state.prev_space_after)
    {
        state.pb.top_suppressed = (effective_space_before - gap).max(0.0);
        inter_gap = gap;
    }

    let applied_inter_gap = inter_gap;
    // The part of the band above that lies below the previous paragraph's
    // `between` border takes this paragraph's shading.
    let prev_band_below = if prev_borders_match && !state.pb.is_at_page_top(sp) {
        prev_between_below
    } else {
        0.0
    };
    state.pb.slot_top -= inter_gap;
    if moved {
        let first = lines.first().map_or(line_h, |l| {
            layout::line_own_advance(l, line_h, para_metrics)
        });
        let rest = &section_blocks[block_idx..];
        place_moved_first_line(
            state,
            first,
            applied_inter_gap,
            rest,
            sect_idx,
            sp,
            col_geometry,
            text_width,
            ctx,
        );
    }
    let d = step_lines_below_floats(
        &mut lines,
        &state.pb,
        col_geometry[state.current_col],
        sp.page_height,
        state.pb.slot_top,
        line_h,
        para_metrics,
    ) - planned_gaps;
    content_h += d;
    lines_h += d;

    for bookmark in &para.bookmarks {
        state.bookmark_positions.insert(
            bookmark.clone(),
            (state.pb.all_contents.len(), state.pb.slot_top),
        );
    }

    if let Some(level) = para.outline_level {
        let title: String = para.runs.iter().map(|r| r.text.as_str()).collect();
        if !title.trim().is_empty() {
            state.heading_entries.push(HeadingEntry {
                title: title.trim().to_string(),
                level,
                page_idx: state.pb.all_contents.len(),
                y_position: state.pb.slot_top,
            });
        }
    }

    // Re-fetch column geometry (may have changed after overflow)
    let (col_x, col_w) = col_geometry[state.current_col];
    para_text_x = col_x + para.indent_left;
    para_text_width = (col_w - para.indent_left - para.indent_right).max(1.0);
    label_x = list_label::label_x(para, col_x, col_w, ctx.fonts);

    // Re-apply float zone adjustment after potential column change
    let first_line_top = state.pb.slot_top - inter_gap;
    if let Some(ref fz) = state.pb.float_zone {
        fz.narrow_paragraph(
            first_line_top,
            col_x,
            col_w,
            para,
            &mut para_text_x,
            &mut para_text_width,
            &mut label_x,
            ctx.compat_mode < 15,
        );
    }
    // A centred or right-aligned paragraph carries its label with the text:
    // Word lays out label, tab gap and first line as one left-aligned unit
    // and shifts the unit by the line's slack (issue #145 probe: bullet at
    // 244.7, text at 266.6 on a 453.85pt measure, the 21.85pt gap the same
    // as the left-aligned control's); `cell_label_alignment_offset` is the
    // same rule for cells. ponytail: left-to-right paragraphs only.
    if !para.bidi {
        let first_width = lines.first().map_or(0.0, |l| l.left_content_width());
        let slack = (para_text_width + text_hanging - first_width).max(0.0);
        label_x += align_offset(para.alignment, slack);
    }

    // A floating table pushed onto a fresh page hands its anchor paragraph the
    // page-body top so paragraph-relative shapes anchor there (the flow cursor
    // stays below the table). One-shot: consumed by this, the next, paragraph.
    let mut anchor = ParaAnchor::new(
        state.pb.take_float_anchor_for(block_idx),
        state.pb.anchor_slot_top(sp, ctx),
        own_space_before(applied_inter_gap, state.prev_space_after),
        sp,
        ctx,
    );
    anchor.on_line -= on_line_drop;
    // Hand the look-ahead float's anchor to the paragraph that actually
    // carries it only after this paragraph has taken its own.
    if let Some((anchor_top, _, anchor_block)) = lookahead {
        state.pb.pending_float_anchor = Some((
            PendingAnchor {
                top: anchor_top,
                kind: AnchorKind::LookAhead,
            },
            Some(anchor_block),
        ));
    } else if state
        .pb
        .pending_float_anchor
        .is_none_or(|(_, b)| b.is_none_or(|b| b <= block_idx))
    {
        // One-shot: an anchor meant for this paragraph (or any next one) is
        // spent; one for a later block rides through the paragraphs between.
        state.pb.pending_float_anchor = None;
    }

    // Margin anchors belong to the page that already started. A continuous
    // section's new margins apply on its next sheet, not to this page's float.
    let page_sp = &doc.sections[state.pb.page_hf_section].properties;
    let render_images = if state.pb.page_hf_section != sect_idx
        && (page_sp.margin_left != sp.margin_left || page_sp.margin_right != sp.margin_right)
    {
        let mut images = para.floating_images.clone();
        for fi in &mut images {
            if fi.h_relative_from == crate::model::HRelativeFrom::Margin
                && fi.wrap_type == WrapType::None
            {
                let x = fi.h_position.place(
                    page_sp.margin_left,
                    page_sp.text_width(),
                    fi.image.display_width,
                );
                fi.h_relative_from = crate::model::HRelativeFrom::Page;
                fi.h_position = crate::model::HorizontalPosition::Offset(x);
            }
        }
        std::borrow::Cow::Owned(images)
    } else {
        std::borrow::Cow::Borrowed(&para.floating_images)
    };

    // Floating images + textboxes, each into a per-shape buffer on its layer:
    // behind-doc ones paint under the page's text, the rest above it. Both
    // layers are sorted by relativeHeight in flush_page, so foreground images
    // interleave with foreground shapes by z-order rather than always
    // painting beneath them (annotation #191), and Word's z-order for
    // floating shapes spans paragraphs.
    let page = state.pb.all_contents.len();
    render_floating_images(
        &render_images,
        state.global_block_idx,
        &images.floating_image_pdf_names,
        &images.effect_floating_names,
        sp,
        col_x,
        col_w,
        text_width,
        anchor,
        &mut state.pb.behind_shapes,
        &mut state.pb.deferred_shapes,
        &mut state.pb.tags,
        page,
    );
    for tb in &para.textboxes {
        render_para_textbox(
            tb,
            para,
            anchor,
            sp,
            (col_x, col_w),
            text_width,
            ctx,
            &mut state.pb,
        );
    }

    // Draw paragraph shading (background), extending outward to match borders
    if let Some(shd_color) = para.shading {
        // Up to the side borders' inner edges: stopping at their space left a
        // white seam inside english_town_council's blue section bars.
        let (shd_left_outset, shd_right_outset) = side_border_outsets(&para.borders, false);
        let (box_left_in, box_right_in) = border_box_indents(para);
        let shd_left = col_x + box_left_in - shd_left_outset;
        let shd_right = col_x + col_w - box_right_in + shd_right_outset;
        // Only a group member with the same fill closes the gap above:
        // renewable_dispatch's blue headings join the unshaded list item
        // above them in a (borderless) group, and Word leaves that gap white.
        let shd_top = state.pb.slot_top
            + if prev_borders_match && prev_para.is_some_and(|pp| pp.shading == para.shading) {
                applied_inter_gap
            } else {
                0.0
            }
            + prev_band_below;
        // min: a vanished paragraph mark zeroes content_h after `lines_h`.
        let shd_bottom_pad = if bottom_collapses {
            bottom_band.0
        } else {
            bdr_bottom_pad
        };
        let shd_bottom = state.pb.slot_top - bdr_top_pad - content_h.min(lines_h) - shd_bottom_pad;
        state.pb.content.save_state();
        fill_rgb(&mut state.pb.content, shd_color);
        state.pb.content.rect(
            shd_left,
            shd_bottom,
            shd_right - shd_left,
            shd_top - shd_bottom,
        );
        state.pb.content.fill_nonzero();
        state.pb.content.restore_state();
    }

    // Set FloatZone for wrapping floating images
    // (may already be set by self-wrapping above; overwrite
    // to ensure polygon data is included).
    // A float entirely outside the text column (e.g. a QR code in the left
    // margin) never narrows text — installing its zone would only clobber a
    // still-active in-column zone from an earlier paragraph's float.
    let mut zones = Vec::new();
    for fi in para
        .floating_images
        .iter()
        .filter(|fi| wraps_in_column(fi, sp, col_x, col_w, text_width))
    {
        let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
        let fi_y_top = anchor.fi_y_top(fi, sp);
        zones.push(FloatZone::for_float(fi, fi_x, fi_y_top));
        if !text_empty && leading_full_width_picture(fi, sp, col_x, col_w, text_width) {
            let bottom = fi_y_top - fi.image.display_height - fi.dist_bottom;
            state.pb.lower_next_text_to(bottom);
        }
    }
    zones.extend(textbox_zones(
        para,
        sp,
        (col_x, col_w),
        text_width,
        anchor,
        ctx,
    ));
    state.pb.set_float_zones(zones);

    if debug_wrap && let Some(ref fz) = state.pb.float_zone {
        draw_debug_wrap_overlay(&mut state.pb.content, fz);
    }

    for conn in &para.connectors {
        // Same page-level z-stack as textboxes — anchored connectors must
        // interleave with shapes by relativeHeight (e.g. letter strokes
        // drawn over gradient circles)
        let mut shape_content = tagging::artifact_content();
        // An inline canvas's connector goes with its line like the canvas's
        // other shapes.
        let top = table_connector_anchor
            .filter(|(page, _)| !conn.on_line && *page == state.pb.page_count())
            .map_or(
                anchor.drawing_top(conn.on_line, conn.v_from_line),
                |(_, top)| top,
            );
        let (x, y) = positioning::connector_top_left(conn, sp, col_x, col_w, text_width, top);
        render_connector(conn, &mut shape_content, x, y);
        state.pb.deferred_shapes.push((conn.z_index, shape_content));
    }

    for diagram in &para.floating_smartart {
        let Some(anchor) = &diagram.anchor else {
            continue;
        };
        let x = positioning::resolve_h_position(
            anchor.h_relative_from,
            &anchor.h_position,
            anchor.width,
            sp,
            col_x,
            col_w,
            text_width,
        );
        let y = header_footer::resolve_tb_y_top(
            anchor.v_relative_from,
            &anchor.v_position,
            diagram.display_height,
            sp,
            state.pb.slot_top,
        );
        // A content-less Figure after the paragraph, as for an inline diagram.
        let alt = smartart_alt(std::slice::from_ref(diagram));
        state.pb.tags.hoist_figure(alt.as_deref(), 0);
        let mut shape_content = tagging::artifact_content();
        smartart::render_smartart(
            &mut shape_content,
            diagram,
            x,
            y,
            ctx.fonts,
            smartart_font_key,
            &images.smartart_image_names,
        );
        state
            .pb
            .deferred_shapes
            .push((anchor.z_index, shape_content));
    }

    if let Some(ref ic) = para.inline_chart {
        let chart_x = col_x + align_offset(para.alignment, (col_w - ic.display_width).max(0.0));
        state.pb.figure_without_content(para, doc, None);
        charts::render_chart(
            ic,
            &mut state.pb.content,
            chart_x,
            state.pb.slot_top,
            ctx.fonts,
            ctx.chart_font_name,
            &mut state.pb.alpha_states,
        );
    } else if !para.smartart.is_empty() {
        let alt = smartart_alt(&para.smartart);
        state.pb.figure_without_content(para, doc, alt.as_deref());
        for (i, diagram) in para.smartart.iter().enumerate() {
            if i > 0 {
                state.pb.slot_top -= diagram.display_height;
            }
            smartart::render_smartart(
                &mut state.pb.content,
                diagram,
                col_x,
                state.pb.slot_top,
                ctx.fonts,
                smartart_font_key,
                &images.smartart_image_names,
            );
        }
    } else if let Some(ref hr) = para.horizontal_rule {
        let line_bottom = state.pb.slot_top - content_h;
        draw_horizontal_rule(&mut state.pb.content, para, hr, col_x, col_w, line_bottom);
    } else if para.image.is_some() && para.content_height > 0.0 {
        if let Some(pdf_name) = images.image_pdf_names.get(&state.global_block_idx) {
            let img = para.image.as_ref().unwrap();
            // Decorative pictures stay artifacts, as in Word's export.
            if img.decorative {
                // The paragraph mark still has its element (a heading's H3).
                state.pb.tag_empty_para(para, doc);
            } else {
                state.pb.begin_figure(para, doc, img.alt.as_deref());
            }
            // A picture shorter than the line sits, effect extents included, on
            // a baseline one natural line of the mark's font down, whatever the
            // spacing (Word probe: 2.25-12pt pictures under Arial 12 at single
            // and 1.5 lines).
            let turned = img.rotation_deg.abs() > 0.01;
            let trim = if swap_axes { 0.0 } else { img.trimmed_width };
            let (box_w, box_h) = if swap_axes {
                (img.display_height, img.display_width)
            } else {
                (img.display_width + trim, img.display_height)
            };
            let image_height = box_h + img.layout_extra_height;
            let bottom_depth = if image_height + inserted_descent > line_h {
                img.layout_extra_top + box_h
            } else {
                // An at-least line's room above the picture stays above it.
                let baseline = if matches!(effective_ls, LineSpacing::AtLeast(_)) {
                    line_h
                } else {
                    (image_height + inserted_descent).max(natural_line_h) - inserted_descent
                };
                baseline - (img.layout_extra_height - img.layout_extra_top)
            };
            let box_bottom = state.pb.slot_top - bottom_depth;
            // Keep the picture paragraph's existing first-line and side indents.
            let first = if para.list_label.is_empty() {
                para.indent_first_line - para.indent_hanging
            } else {
                0.0
            };
            let start = para.indent_left + first;
            let room = col_w - start - para.indent_right - box_w;
            let box_x = col_x + start + align_offset(para.alignment, room.max(0.0));
            let x = box_x + (box_w - trim - img.display_width) / 2.0;
            let y_bottom = box_bottom + (box_h - img.display_height) / 2.0;
            if turned {
                state.pb.content.save_state();
                // Legacy inline layout clips a turned picture to its original
                // frame instead of growing the occupied line box.
                if ctx.compat_mode < 15 {
                    state.pb.content.rect(box_x, box_bottom, box_w, box_h);
                    state.pb.content.clip_nonzero();
                    state.pb.content.end_path();
                }
                positioning::push_center_rotation(
                    &mut state.pb.content,
                    box_x,
                    box_bottom,
                    box_w,
                    box_h,
                    img.rotation_deg,
                );
            }
            let img_fx = images.effect_names.get(&state.global_block_idx);
            if let Some(ref shadow) = img.shadow {
                color::draw_image_shadow(
                    &mut state.pb.content,
                    shadow,
                    x,
                    y_bottom,
                    img.display_width,
                    img.display_height,
                    img_fx.and_then(|fx| fx.shadow.as_deref()),
                );
            }
            if let Some(ref glow) = img.glow {
                color::draw_image_glow(
                    &mut state.pb.content,
                    glow,
                    x,
                    y_bottom,
                    img.display_width,
                    img.display_height,
                    img_fx.and_then(|fx| fx.glow.as_deref()),
                );
            }
            smartart::render_image_with_clip(
                &mut state.pb.content,
                pdf_name,
                x,
                y_bottom,
                img.display_width,
                img.display_height,
                img.clip_geometry.as_ref(),
            );
            if let Some(sc) = img.stroke_color {
                smartart::stroke_image_border(
                    &mut state.pb.content,
                    x,
                    y_bottom,
                    img.display_width,
                    img.display_height,
                    sc,
                    img.stroke_width,
                    img.clip_geometry.as_ref(),
                );
            }
            // Post-image effects: inner shadow, reflection (drawn on top / below)
            if let Some(ref inner) = img.inner_shadow {
                color::draw_inner_shadow(
                    &mut state.pb.content,
                    inner,
                    x,
                    y_bottom,
                    img.display_width,
                    img.display_height,
                    img_fx.and_then(|fx| fx.inner_shadow.as_deref()),
                );
            }
            if let Some(ref refl) = img.reflection {
                color::draw_reflection(
                    &mut state.pb.content,
                    refl,
                    x,
                    y_bottom,
                    img.display_width,
                    img.display_height,
                    img_fx.and_then(|fx| fx.reflection.as_deref()),
                );
            }
            if turned {
                state.pb.content.restore_state();
            }
            // Markup view: the inserted picture's revision underline, the run's
            // post-table geometry on a 0.25pt grid like the text path's.
            if inserted_descent > 0.0
                && let Some(run) = para.runs.iter().find(|r| r.revised && r.underline)
                && let Some((offset, thick)) = ctx
                    .fonts
                    .get(&font_key(run))
                    .and_then(|e| layout::post_underline(e, run.font_size))
            {
                color::fill_color_or_black(&mut state.pb.content, run.color);
                state
                    .pb
                    .content
                    .rect(box_x, box_bottom - offset - thick, box_w, thick)
                    .fill_nonzero();
            }
            if !img.decorative {
                state.pb.end_tag();
            }
        } else if para.image.is_some() {
            state
                .pb
                .content
                .set_fill_gray(0.5)
                .rect(col_x, state.pb.slot_top - content_h, col_w, content_h)
                .fill_nonzero()
                .set_fill_gray(0.0);
        }
    } else if lines.is_empty()
        && para.list_label.is_empty()
        && para.mark_revised
        && !para.paragraph_mark_vanish
    {
        layout::set_change_bar_extension(
            own_space_before(applied_inter_gap, state.prev_space_after),
            effective_space_after,
        );
        // Markup view: an empty paragraph with a revised mark gets its bar
        // (Word's bar runs unbroken over a decree's empty inserted paragraphs).
        let top = state.pb.slot_top - bdr_top_pad;
        layout::record_change_bar(top - content_h, top);
    } else if !lines.is_empty() || !para.list_label.is_empty() {
        // An empty numbered paragraph still shows its label: uk_commercial's
        // "SCHEDULE 1" headings are empty paragraphs whose style numbers them.
        // When the document grid snaps line heights, align the first
        // baseline one linePitch below the slot top so text sits on
        // the grid rather than at a font-metric-dependent offset.
        let baseline_offset = if grid_snapped {
            grid_baseline
        } else {
            first_baseline_offset
        };
        let baseline_y = state.pb.slot_top - bdr_top_pad - baseline_offset - text_drop;
        layout::set_change_bar_extension(
            own_space_before(applied_inter_gap, state.prev_space_after),
            effective_space_after,
        );

        let tags = state.pb.para_tags(para, doc);
        let grid = line_grid(baseline_offset, state.pb.all_contents.len());
        let label_y =
            layout::first_drawn_baseline(baseline_y, &lines, para_ascent, grid, lines.len() <= 1);
        state.pb.begin_para_tags(tags, |content| {
            render_list_label(content, para, ctx.fonts, label_x, label_y, font_size)
        });

        if !lines.is_empty() {
            draw(
                &mut state.pb,
                &lines,
                para_text_x,
                para_text_width,
                baseline_y,
                0,
                poly_line_geom.as_deref(),
                lines.len(),
                tags.1,
                grid,
            );
        }
        state.pb.end_tag();
    } else {
        // Word tags empty paragraphs too; keeping them keeps the P sequence aligned.
        state.pb.tag_empty_para(para, doc);
    }

    // Draw paragraph borders — left/right borders extend outward
    // from the text area so text inside stays aligned with text outside
    {
        let bdr = &para.borders;
        let box_top = state.pb.slot_top - bdr_top_half_band;
        let box_bottom = state.pb.slot_top - bdr_top_pad - content_h - bdr_bottom_pad;
        // Word puts a side border's inner edge its space plus 1.47pt (left) or
        // 1.73pt (right) outside the text, whatever its width (online export
        // probes: sz 4/12/24 × space 0/4/12, with and without indents).
        let (bdr_left_outset, bdr_right_outset) = side_border_outsets(bdr, true);
        let (box_left_in, box_right_in) = border_box_indents(para);
        let box_left = col_x + box_left_in - bdr_left_outset;
        let box_right = col_x + col_w - box_right_in + bdr_right_outset;

        // Extend horizontal borders past the corners so they cover the
        // corner gap left by butt-capped vertical border strokes.
        let (h_left_ext, h_right_ext) = horizontal_border_ext(bdr);
        // `inward`: +1 for a border below the text, -1 for one above.
        let page_top = sp.page_height
            + ctx
                .valign_offsets
                .get(state.pb.all_contents.len())
                .copied()
                .unwrap_or(0.0);
        let draw_h_border = |content: &mut Content, b: &ParagraphBorder, y: f32, inward: f32| {
            let y = layout::rule_on_grid(y, page_top, b.extent());
            let (x0, x1) = (box_left - h_left_ext, box_right + h_right_ext);
            stroke_border(content, b, (x0, y), (x1, y), inward);
        };
        // When this paragraph continues a border group, extend vertical
        // borders upward through the inter-paragraph gap so there is no
        // visible break between consecutive paragraphs' left/right borders.
        let v_border_top = box_top
            + if prev_borders_match {
                applied_inter_gap
            } else {
                0.0
            };
        // `inward`: +1 for the left border, -1 for the right.
        // Down through the band of a `between` border the group goes on below.
        let v_border_bottom = box_bottom
            - if bottom_collapses {
                bdr_bottom_extent - bdr_bottom_pad
            } else {
                0.0
            };
        let draw_v_border = |content: &mut Content, b: &ParagraphBorder, x: f32, inward: f32| {
            let (top, bottom) = (
                layout::rule_on_grid(v_border_top, page_top, 0.0),
                layout::rule_on_grid(v_border_bottom, page_top, 0.0),
            );
            stroke_border(content, b, (x, top), (x, bottom), inward);
        };

        if !prev_borders_match && let Some(b) = &bdr.top {
            draw_h_border(&mut state.pb.content, b, box_top, -1.0);
        }
        if let Some(b) = bottom_bdr {
            draw_h_border(&mut state.pb.content, b, box_bottom, 1.0);
        }
        if let Some(b) = &bdr.left {
            draw_v_border(&mut state.pb.content, b, box_left, 1.0);
        }
        if let Some(b) = &bdr.right {
            draw_v_border(&mut state.pb.content, b, box_right, -1.0);
        }
    }

    // Word numbers a paragraph without text lines too (an empty anchor under
    // its boxes, a chart or a picture), on its mark line at the foot.
    if lines.is_empty()
        && content_h > 0.0
        && let Some((start, count_by, right_x)) = ln_cfg
    {
        // ponytail: no grid snap; a grid-snapped empty anchor's number may sit off its line.
        // A picture's, chart's or inline shape's bottom sits on its line's
        // baseline.
        let picture_bottom = if para.image.is_some() || para.inline_chart.is_some() {
            Some(para.content_height)
        } else {
            para.textboxes
                .iter()
                .filter(|tb| tb.on_line)
                .map(crate::model::Textbox::flow_bottom)
                .reduce(f32::max)
        };
        let mark_baseline = if let Some(bottom) = picture_bottom {
            state.pb.slot_top - bdr_top_pad - bottom
        } else {
            state.pb.slot_top - bdr_top_pad - content_h + (line_h - first_baseline_offset).max(0.0)
        };
        let fonts = fonts::by_pdf_name(ctx.fonts);
        LineNumberArg {
            counter: &mut state.pb.line_number_counter,
            start,
            count_by,
            right_x,
            font: ln_font.clone(),
        }
        .number(
            &mut state.pb.content,
            0.0,
            mark_baseline,
            || None,
            &fonts,
            None,
        );
    }

    state.pb.slot_top -= content_h + bdr_top_pad + bdr_bottom_extent;
    state.prev_between_below = if bottom_collapses {
        between_space(para)
    } else {
        0.0
    };
    if state.pb.slot_top < state.effective_margin_bottom - 1.0 {
        log::warn!(
            "Body overflow: slot_top={:.2} < eff_margin_bottom={:.2} after paragraph on page {}",
            state.pb.slot_top,
            state.effective_margin_bottom,
            state.pb.all_contents.len(),
        );
    }
    if !(text_empty && para.paragraph_mark_vanish) {
        state.prev_space_after = effective_space_after;
        layout::set_change_bar_extension(0.0, 0.0);
    }

    // Track footnotes referenced on this page
    for run in para.runs.iter() {
        if let Some(id) = run.footnote_id {
            track_page_footnote(state, doc, ctx, col_geometry[state.current_col], id);
        }
        if let Some(id) = run.endnote_id {
            state.pb.track_endnote(id);
        }
    }

    update_styleref_from_para(
        &mut state.pb.styleref_running,
        &mut state.pb.styleref_page_first,
        para,
        &doc.style_id_to_name,
    );

    state.take_page_break_after(para, sect_idx, sp, ctx);
    if para.column_break_after {
        state.advance_column_or_page(col_count, sect_idx, sp, ctx);
        state.prev_space_after = 0.0;
    }

    false
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_trailing_anchor_has_only_blank_runs_after_it() {
        let text = |t: &str| crate::model::Run {
            text: t.to_string(),
            ..Default::default()
        };
        let brk = crate::model::Run {
            is_line_break: true,
            ..Default::default()
        };
        let runs = [text("Med forbehold"), text(" "), brk, text("")];
        assert!(blank_after(&runs, 4));
        assert!(blank_after(&runs, 3));
        assert!(
            blank_after(&runs, 9),
            "an index past the end is a trailing anchor"
        );
        assert!(!blank_after(&runs, 2), "a break after the anchor is a line");
        assert!(!blank_after(&runs, 0));
    }
}
