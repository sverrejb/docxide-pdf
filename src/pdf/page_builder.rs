//! Page state: the page being filled, float anchors, layout state and page numbering.

use super::*;

/// A one-shot override of where the next paragraph's drawings hang from.
#[derive(Clone, Copy)]
pub(super) struct PendingAnchor {
    pub(super) top: f32,
    pub(super) kind: AnchorKind,
}

#[derive(Clone, Copy)]
pub(super) enum AnchorKind {
    /// Drawings on the line (inline canvases, inline text boxes) follow the
    /// line wherever it goes.
    Floats,
    /// Drawings on the line take it too.
    TableMoved,
    /// Drawings on and positioned relative to the line take it too (see
    /// `ParaAnchor::drawing_top`).
    LookAhead,
}

/// Where a paragraph's drawings hang from. Paragraph-relative offsets count
/// from below the previous paragraph's space after but above this one's own
/// space before: Word keeps indonesian_school's "Format 12" box in place for
/// before = 0, 10 and 200 twips, and case124's stamp below the section-break
/// paragraph's 8pt. Drawings on the line (inline canvases, inline text boxes)
/// count from the line top, below the space before.
#[derive(Clone, Copy)]
pub(super) struct ParaAnchor {
    float: f32,
    pub(super) on_line: f32,
    /// The paragraph's top where it flowed, space before included, or the
    /// look-ahead float's anchor.
    flowed: f32,
    /// The highest a wrapping paragraph-relative picture may sit: the top
    /// margin under compat 15 (Word probes: -30pt from a first paragraph
    /// lands on the margin; compat 14 and older keep the -30pt).
    ceiling: f32,
}

impl ParaAnchor {
    /// `pending`: a floating table's one-shot override for its anchor
    /// paragraph; `line_top`: the first line's top, below the space before.
    pub(super) fn new(
        pending: Option<PendingAnchor>,
        line_top: f32,
        own_before: f32,
        sp: &SectionProperties,
        ctx: &RenderContext,
    ) -> Self {
        let below_before = line_top + own_before;
        let (float, on_line, flowed) = match pending {
            None => (below_before, line_top, below_before),
            Some(PendingAnchor {
                top,
                kind: AnchorKind::Floats,
            }) => (top, line_top, below_before),
            Some(PendingAnchor {
                top,
                kind: AnchorKind::TableMoved,
            }) => (top, top, below_before),
            Some(PendingAnchor {
                top,
                kind: AnchorKind::LookAhead,
            }) => (top, top, top),
        };
        Self {
            float,
            on_line,
            flowed,
            // Compat 14 keeps a negative offset only down to the page edge
            // (Word probes p20c/p20d: -82.7pt from a page-top paragraph puts
            // the picture at y 0).
            ceiling: if ctx.compat_mode >= 15 {
                sp.page_height - sp.margin_top
            } else {
                sp.page_height
            },
        }
    }

    /// Where a drawing's vertical offset counts from. One placed relative to
    /// the line follows its paragraph below a band or floating table that
    /// pushed it, where a paragraph-relative one stays where the paragraph
    /// stood (Word examples, fixture line-anchor-follows-push: offsets -30, 0
    /// and 10pt; with 24pt before, the line one counts from above it). It
    /// follows the line below a floating table moved whole to the next page
    /// too, but keeps a look-ahead float's anchor like a paragraph-relative
    /// one (fixture line-anchor-overrides).
    pub(super) fn drawing_top(self, on_line: bool, v_from_line: bool) -> f32 {
        if on_line {
            self.on_line
        } else if v_from_line {
            self.flowed
        } else {
            self.float
        }
    }

    /// A body float's top (`resolve_fi_y_top`) under the anchor's ceiling.
    pub(super) fn fi_y_top(self, fi: &crate::model::FloatingImage, sp: &SectionProperties) -> f32 {
        let y = resolve_fi_y_top(fi, sp, self.drawing_top(fi.on_line, fi.v_from_line));
        let wrapping_para = fi.wrap_type != WrapType::None
            && matches!(fi.v_relative_from, VRelativeFrom::Paragraph);
        // A wrapping picture taller than the page body fits no page: Word
        // keeps it on its anchor's sheet, moved up to the page top (an
        // external corpus cover: a 797pt picture anchored -82.7pt above the
        // third paragraph of a 595pt landscape page prints at y 0, the title
        // wrapped beside it; its A4 sibling, where the picture fits, keeps
        // 18.85). One that fits a page is not moved: its anchor paragraph
        // goes to the next page instead (leading-anchor fixtures).
        let h = fi.image.display_height;
        let y = if wrapping_para
            && h > sp.page_height - sp.margin_top - sp.margin_bottom
            && y - h < 0.0
        {
            y.max(h.min(sp.page_height))
        } else {
            y
        };
        if wrapping_para {
            y.min(self.ceiling)
        } else {
            y
        }
    }
}

impl PageBuilder {
    /// The pending float anchor, when it is section block `block_idx`'s (or
    /// the next paragraph's, when no block was named).
    pub(super) fn float_anchor_for(&self, block_idx: usize) -> Option<PendingAnchor> {
        self.pending_float_anchor
            .filter(|&(_, b)| b.is_none_or(|b| b == block_idx))
            .map(|(a, _)| a)
    }

    pub(super) fn take_float_anchor_for(&mut self, block_idx: usize) -> Option<PendingAnchor> {
        let anchor = self.float_anchor_for(block_idx)?;
        self.pending_float_anchor = None;
        Some(anchor)
    }
}

pub(super) struct PageBuilder {
    /// The body table being rendered, as `(left, right)`: a page it opens
    /// starts below the header floats it would overlap.
    pub(super) table_extent: Option<(f32, f32)>,
    // Current page state
    pub(super) content: Content,
    pub(super) links: Vec<LinkAnnotation>,
    /// (comment_id, anchor_x, anchor_y) — anchor is the end of the last chunk
    /// covered by that comment on this page. Used by the comment-pane renderer
    /// to draw the connector line from highlighted phrase to callout.
    pub(super) comment_anchors: Vec<(u32, f32, f32, f32)>,
    pub(super) footnote_ids: Vec<u32>,
    pub(super) footnote_ids_set: HashSet<u32>,
    /// Endnote IDs encountered across the whole document, in encounter order.
    /// Endnotes default to `pos=docEnd`, so they all render on the final page
    /// rather than on the page where the reference appears.
    pub(super) endnote_ids: Vec<u32>,
    pub(super) endnote_ids_set: HashSet<u32>,
    pub(super) alpha_states: HashSet<u8>,
    pub(super) gradient_specs: Vec<GradientSpec>,

    // Cross-page running state
    pub(super) styleref_running: HashMap<String, String>,
    pub(super) styleref_page_first: HashMap<String, String>,

    // Layout position state
    pub(super) slot_top: f32,
    /// Y-position at which the current column on this page begins. For a
    /// continuous 2-column section that starts mid-page, both the left and
    /// right columns share this top-y so advancing from left to right returns
    /// to where the section started rather than the top of the page.
    pub(super) column_top_y: f32,
    /// Where this page's body starts: below a header taller than the top
    /// margin, not at the margin (`is_at_page_top`).
    pub(super) page_top_y: f32,
    pub(super) is_first_page_of_section: bool,
    /// A continuous section whose first block has not flowed yet.
    pub(super) pending_section_first_page: Option<(usize, usize, f32)>,
    /// Section that owns the current page for header/footer purposes.
    /// For continuous section breaks, this stays as the section that started
    /// the page, not the section that continues mid-page.
    pub(super) page_hf_section: usize,
    /// Floating table exclusion zone on this page; paragraph layout
    /// uses horizontal bounds to decide wrap-beside vs push-below.
    pub(super) float_zone: Option<FloatZone>,
    /// The other wrapping floats of the paragraph that set `float_zone`,
    /// next in line once the cursor passes it.
    pub(super) more_float_zones: Vec<FloatZone>,
    /// The last floating table's horizontal extent, kept after its zone is
    /// consumed or lost to a page break.
    pub(super) float_table_extent: Option<FloatTableExtent>,
    /// The bands of this page's lifted frames and blocking floats: body
    /// lines may not overlap them (see `OpenFrame`, `step_below_bands`).
    pub(super) frame_bands: Vec<Band>,
    /// The (section, page) whose upcoming float bands `prescan_page_bands`
    /// collected.
    pub(super) bands_prescanned: Option<(usize, usize)>,
    /// One-shot anchor override for the paragraph that follows a floating table
    /// which was pushed whole onto a fresh page. That paragraph (the table's
    /// vertAnchor="text" anchor) must position its paragraph-relative shapes
    /// from the top of the page body — where the anchor naturally flows — even
    /// though the flow cursor stays below the table so following text doesn't
    /// overlap it. Without this the shapes drop by the table's height.
    /// With it, the section block the anchor is for, when a look-ahead found
    /// its float more than one paragraph ahead; `None` means the next one.
    pub(super) pending_float_anchor: Option<(PendingAnchor, Option<usize>)>,
    pub(super) next_text_below_picture: Option<f32>,
    /// Anchored shapes for this page, painted above the text layer sorted by
    /// relativeHeight: Word stacks floating shapes in z order regardless of
    /// which paragraph anchors them.
    pub(super) deferred_shapes: Vec<(u32, Content)>,
    /// Behind-text drawings for this page, painted under all of its body text
    /// (shading included) sorted by relativeHeight, whatever paragraph or
    /// table cell anchors them.
    pub(super) behind_shapes: Vec<(u32, Content)>,

    // Accumulated pages
    pub(super) all_contents: Vec<Content>,
    /// An exact frame in the flow clips this page's content (`FlowFrame`).
    pub(super) frame_clip: bool,
    /// Per-page y of the cursor at flush time (bottom of the last body block).
    /// Used to compute the §17.6.23 `w:vAlign` center/bottom offset.
    pub(super) all_content_bottom: Vec<f32>,
    pub(super) all_deferred_shapes: Vec<Vec<(u32, Content)>>,
    pub(super) all_behind_shapes: Vec<Vec<(u32, Content)>>,
    pub(super) all_links: Vec<Vec<LinkAnnotation>>,
    pub(super) all_comment_anchors: Vec<Vec<(u32, f32, f32, f32)>>,
    pub(super) all_footnote_ids: Vec<Vec<u32>>,
    /// Each booked footnote's column (x, width), parallel to `footnote_ids`:
    /// Word sets a multi-column section's footnotes at the foot of the column
    /// that cites them, in its width.
    footnote_cols: Vec<(f32, f32)>,
    pub(super) all_footnote_cols: Vec<Vec<(f32, f32)>>,
    /// Footnote space booked in the current column.
    pub(super) col_fn_reserved: f32,
    /// Footnotes (id, column, height, separator) cited on this page with no
    /// room left for them: the next page books them.
    carried_footnotes: Vec<(u32, (f32, f32), Option<f32>, f32)>,
    /// §17.6.8 body line-number counter (0-based count of lines seen).
    pub(super) line_number_counter: u32,
    /// The current section's `lnNumType restart` is newPage (the default):
    /// the counter goes back to zero with every page flushed.
    pub(super) line_numbers_restart_each_page: bool,
    pub(super) all_alpha_states: Vec<HashSet<u8>>,
    pub(super) all_gradient_specs: Vec<Vec<GradientSpec>>,
    /// Pages inserted by odd/even section breaks; Word prints them without header or footer.
    pub(super) filler_pages: Vec<usize>,
    /// Per-page tuples: (hf_section, is_first_page, content_section).
    /// hf_section: which section provides headers/footers.
    /// content_section: which section is being rendered (for page numbering, geometry).
    pub(super) page_section_indices: Vec<(usize, bool, usize)>,
    /// Highest column index the current column region reached on this page;
    /// Word draws a column separator only up to the last column holding text.
    last_col: usize,
    /// Lowest point a finished column of the current region reached on this page.
    pub(super) col_bottom: f32,
    /// Separator x positions of the current column region (empty without w:sep).
    pub(super) region_sep_xs: Vec<f32>,
    /// Separator lines on the current page: (x positions, top y, bottom y).
    col_seps: Vec<(Vec<f32>, f32, f32)>,
    pub(super) all_col_seps: Vec<Vec<(Vec<f32>, f32, f32)>>,
    /// Per page, (bottom, top) of each revised line (markup view).
    pub(super) all_change_bars: Vec<Vec<(f32, f32)>>,
    /// Summed heights of the current region's finished columns on this page.
    pub(super) col_heights: f32,
    /// Space before dropped at the top of this page (balancing still counts it).
    pub(super) top_suppressed: f32,
    /// (page index, y): columns on that page end no lower than y, to balance
    /// the last page of a region spanning pages.
    pub(super) balance_floor: Option<(usize, f32)>,
    pub(super) all_styleref: Vec<HashMap<String, String>>,
    pub(super) all_first_styleref: Vec<HashMap<String, String>>,
    pub(super) tags: tagging::Tags,
    pub(super) lists: tagging::Lists,
    /// The open TOC element while consecutive "toc N" paragraphs are tagged.
    pub(super) toc: Option<usize>,
    /// A TOC field has begun and its entries haven't ended yet.
    toc_field: bool,
    /// Structure of the body table being rendered (see `table::render_table`).
    pub(super) table_tags: Option<tagging::TableTags>,
}

impl PageBuilder {
    /// The cursor left `float_zone` behind: the next of its paragraph's
    /// floats still below the cursor takes its place.
    pub(super) fn next_float_zone(&mut self) {
        let cursor = self.slot_top;
        self.more_float_zones.retain(|z| z.bottom_y < cursor);
        self.float_zone =
            (!self.more_float_zones.is_empty()).then(|| self.more_float_zones.remove(0));
    }

    /// The wrapping floats a paragraph just placed replace the page's zones,
    /// its last one first.
    pub(super) fn set_float_zones(&mut self, mut zones: Vec<FloatZone>) {
        if let Some(last) = zones.pop() {
            self.float_zone = Some(last);
            self.more_float_zones = zones;
        }
    }

    /// A paragraph moved below a band or a floating table: its floating
    /// paragraph-relative drawings still hang from where it stood, while inline
    /// ones go with its line.
    pub(super) fn pin_float_anchor(&mut self, top: f32) {
        self.pending_float_anchor.get_or_insert((
            PendingAnchor {
                top,
                kind: AnchorKind::Floats,
            },
            None,
        ));
    }

    pub(super) fn new(slot_top: f32) -> Self {
        PageBuilder {
            content: tagging::artifact_content(),
            links: Vec::new(),
            comment_anchors: Vec::new(),
            footnote_ids: Vec::new(),
            footnote_ids_set: HashSet::new(),
            endnote_ids: Vec::new(),
            endnote_ids_set: HashSet::new(),
            alpha_states: HashSet::new(),
            gradient_specs: Vec::new(),
            styleref_running: HashMap::new(),
            styleref_page_first: HashMap::new(),
            slot_top,
            column_top_y: slot_top,
            page_top_y: slot_top,
            is_first_page_of_section: true,
            pending_section_first_page: None,
            page_hf_section: 0,
            float_zone: None,
            more_float_zones: Vec::new(),
            float_table_extent: None,
            frame_bands: Vec::new(),
            bands_prescanned: None,
            pending_float_anchor: None,
            table_extent: None,
            next_text_below_picture: None,
            deferred_shapes: Vec::new(),
            behind_shapes: Vec::new(),
            all_contents: Vec::new(),
            frame_clip: false,
            all_content_bottom: Vec::new(),
            all_deferred_shapes: Vec::new(),
            all_behind_shapes: Vec::new(),
            all_links: Vec::new(),
            all_comment_anchors: Vec::new(),
            all_footnote_ids: Vec::new(),
            footnote_cols: Vec::new(),
            all_footnote_cols: Vec::new(),
            col_fn_reserved: 0.0,
            carried_footnotes: Vec::new(),
            line_number_counter: 0,
            line_numbers_restart_each_page: false,
            all_alpha_states: Vec::new(),
            all_gradient_specs: Vec::new(),
            filler_pages: Vec::new(),
            page_section_indices: Vec::new(),
            last_col: 0,
            col_bottom: f32::INFINITY,
            region_sep_xs: Vec::new(),
            col_seps: Vec::new(),
            all_col_seps: Vec::new(),
            all_change_bars: Vec::new(),
            col_heights: 0.0,
            top_suppressed: 0.0,
            balance_floor: None,
            all_styleref: Vec::new(),
            all_first_styleref: Vec::new(),
            tags: tagging::Tags::new(),
            lists: tagging::Lists::default(),
            toc: None,
            toc_field: false,
            table_tags: None,
        }
    }

    /// Start `node`'s content on the current page (see `tagging`).
    pub(super) fn begin_tag(&mut self, node: usize) {
        let page = self.all_contents.len();
        self.tags.begin(&mut self.content, page, node);
    }

    pub(super) fn end_tag(&mut self) {
        tagging::Tags::end(&mut self.content);
    }

    /// Structure nodes for a body paragraph: (Lbl, element for its text). List
    /// items become LI > Lbl + LBody; numbered headings stay headings.
    pub(super) fn para_tags(&mut self, para: &Paragraph, doc: &Document) -> (Option<usize>, usize) {
        // Word tags a table of contents as one flat TOC holding a TOCI per entry.
        let style_name = para
            .style_id
            .as_ref()
            .and_then(|id| doc.style_id_to_name.get(id));
        let is_toc_entry = style_name.is_some_and(|n| {
            n.get(..4).is_some_and(|p| p.eq_ignore_ascii_case("toc "))
                && n[4..].parse::<u8>().is_ok()
        });
        // Only inside a TOC field: its begin can sit in the entry or a heading above.
        self.toc_field |= para.starts_toc_field;
        if is_toc_entry && self.toc_field {
            self.lists.close();
            let toc = *self
                .toc
                .get_or_insert_with(|| self.tags.add(tagging::ROOT, "TOC"));
            return (None, self.tags.add(toc, "TOCI"));
        }
        self.toc = None;
        self.toc_field = para.starts_toc_field;
        self.tags.para_nodes(
            &mut self.lists,
            tagging::ROOT,
            para.list_item.filter(|_| para.outline_level.is_none()),
            !para.list_label.is_empty(),
            para_tag_kind(para, style_name),
        )
    }

    /// Tag a picture, chart or diagram paragraph the way Word does: an empty
    /// element for the paragraph mark, then the Figure hoisted beside it.
    pub(super) fn begin_figure(&mut self, para: &Paragraph, doc: &Document, alt: Option<&str>) {
        self.tag_empty_para(para, doc);
        let figure = self.tags.add_figure(tagging::ROOT, alt);
        self.begin_tag(figure);
    }

    /// A chart or SmartArt paragraph as Word tags it: the paragraph mark, then
    /// a Figure whose drawing (labels included) stays an artifact, so only the
    /// alt text speaks for it.
    // ponytail: content-less Figure; tag the drawing itself if a validator asks for content
    pub(super) fn figure_without_content(&mut self, para: &Paragraph, doc: &Document, alt: Option<&str>) {
        self.tag_empty_para(para, doc);
        self.tags.add_figure(tagging::ROOT, alt);
    }

    /// The paragraph's elements with nothing drawn inside.
    pub(super) fn tag_empty_para(&mut self, para: &Paragraph, doc: &Document) {
        let mark = self.para_tags(para, doc);
        self.begin_para_tags(mark, |_| {});
        self.end_tag();
    }

    /// Draw the list label as its own Lbl (or inside the paragraph's element
    /// when it has none) and leave the paragraph text's tag open.
    pub(super) fn begin_para_tags(
        &mut self,
        nodes: (Option<usize>, usize),
        draw_label: impl FnOnce(&mut Content),
    ) {
        let page = self.all_contents.len();
        self.tags
            .begin_para(&mut self.content, page, nodes, draw_label);
    }

    pub(super) fn end_frame_clip(&mut self) {
        if std::mem::take(&mut self.frame_clip) {
            self.content.restore_state();
        }
    }

    pub(super) fn flush_page(&mut self, sect_idx: usize) {
        // A frame's clip ends with its page.
        self.end_frame_clip();
        self.all_contents.push(std::mem::replace(
            &mut self.content,
            tagging::artifact_content(),
        ));
        self.all_content_bottom.push(self.slot_top);
        // Stable sort: equal relativeHeight keeps document order
        self.deferred_shapes.sort_by_key(|(z, _)| *z);
        self.all_deferred_shapes
            .push(std::mem::take(&mut self.deferred_shapes));
        self.behind_shapes.sort_by_key(|(z, _)| *z);
        self.all_behind_shapes
            .push(std::mem::take(&mut self.behind_shapes));
        self.all_links.push(std::mem::take(&mut self.links));
        self.all_comment_anchors
            .push(std::mem::take(&mut self.comment_anchors));
        self.all_footnote_ids
            .push(std::mem::take(&mut self.footnote_ids));
        self.all_footnote_cols
            .push(std::mem::take(&mut self.footnote_cols));
        self.col_fn_reserved = 0.0;
        if self.line_numbers_restart_each_page {
            self.line_number_counter = 0;
        }
        self.footnote_ids_set.clear();
        let mut reserved = 0.0;
        for (id, col, height, separator_h) in std::mem::take(&mut self.carried_footnotes) {
            self.book_footnote(id, col, height, separator_h, &mut reserved);
        }
        self.all_alpha_states
            .push(std::mem::take(&mut self.alpha_states));
        self.all_gradient_specs
            .push(std::mem::take(&mut self.gradient_specs));
        self.page_section_indices.push((
            self.page_hf_section,
            self.is_first_page_of_section,
            sect_idx,
        ));
        self.push_col_seps(self.slot_top);
        self.all_col_seps.push(std::mem::take(&mut self.col_seps));
        self.all_change_bars.push(layout::take_change_bars());
        self.top_suppressed = 0.0;
        self.all_styleref.push(self.styleref_running.clone());
        self.all_first_styleref
            .push(std::mem::take(&mut self.styleref_page_first));
        self.float_zone = None;
        self.more_float_zones.clear();
        self.frame_bands.clear();
        // An anchor measured on the old page must not place a float on this one
        self.pending_float_anchor = None;
        self.next_text_below_picture = None;
        // After flush, the new page starts with the current section
        self.page_hf_section = sect_idx;
    }

    pub(super) fn push_blank_page(&mut self, sect_idx: usize) {
        self.filler_pages.push(self.all_contents.len());
        self.all_contents.push(tagging::artifact_content());
        // Blank page has no body content; record top so vAlign yields no shift.
        self.all_content_bottom.push(self.slot_top);
        self.all_deferred_shapes.push(Vec::new());
        self.all_behind_shapes.push(Vec::new());
        self.all_links.push(Vec::new());
        self.all_comment_anchors.push(Vec::new());
        self.all_footnote_ids.push(Vec::new());
        self.all_footnote_cols.push(Vec::new());
        self.all_alpha_states.push(HashSet::new());
        self.all_gradient_specs.push(Vec::new());
        self.page_section_indices
            .push((self.page_hf_section, false, sect_idx));
        self.all_col_seps.push(Vec::new());
        self.all_change_bars.push(Vec::new());
        self.all_styleref.push(self.styleref_running.clone());
        self.all_first_styleref
            .push(std::mem::take(&mut self.styleref_page_first));
        self.page_hf_section = sect_idx;
    }

    pub(super) fn page_count(&self) -> usize {
        self.all_contents.len()
    }

    /// The number the open page shows (`page_numbers`), which picks its even
    /// or odd header.
    pub(super) fn page_num(&self, sections: &[crate::model::Section]) -> usize {
        let open = (self.page_hf_section, false, self.page_hf_section);
        last_page_number(sections, self.page_section_indices.iter().chain([&open]))
    }

    /// Book footnote `id` (height None when it doesn't exist) into the foot
    /// of column `col` once per page, shrinking the column by its height and,
    /// for the column's first note, the separator.
    pub(super) fn book_footnote(
        &mut self,
        id: u32,
        col: (f32, f32),
        height: Option<f32>,
        separator_h: f32,
        effective_margin_bottom: &mut f32,
    ) {
        if !self.footnote_ids_set.insert(id) {
            return;
        }
        self.footnote_ids.push(id);
        self.footnote_cols.push(col);
        if let Some(h) = height {
            let sep = if self.col_fn_reserved == 0.0 {
                separator_h
            } else {
                0.0
            };
            *effective_margin_bottom += sep + h;
            self.col_fn_reserved += sep + h;
        }
    }

    /// Leave footnote `id` for the next page, as `book_footnote` takes it.
    pub(super) fn carry_footnote(
        &mut self,
        id: u32,
        col: (f32, f32),
        height: Option<f32>,
        separator_h: f32,
    ) {
        if !self.footnote_ids_set.contains(&id) {
            self.carried_footnotes.push((id, col, height, separator_h));
        }
    }

    /// Close the current column region's separators on this page: Word draws
    /// them from the region's top down to its deepest column (`bottom` is the
    /// column in progress), and none when the text never left column 1.
    pub(super) fn push_col_seps(&mut self, bottom: f32) {
        let n = self.last_col.min(self.region_sep_xs.len());
        if n > 0 {
            let bottom = bottom.min(self.col_bottom);
            self.col_seps
                .push((self.region_sep_xs[..n].to_vec(), self.column_top_y, bottom));
        }
        self.last_col = 0;
        self.col_bottom = f32::INFINITY;
        self.col_heights = 0.0;
    }

    /// Endnotes render at the end of the document: collect the id once, in
    /// encounter order, for the last page (Phase 2c).
    pub(super) fn track_endnote(&mut self, id: u32) {
        if self.endnote_ids_set.insert(id) {
            self.endnote_ids.push(id);
        }
    }

    /// A page opened (by a break) that nothing has been placed on yet.
    pub(super) fn is_empty_page(&self) -> bool {
        self.page_count() > 0
            && (self.slot_top - self.page_top_y).abs() < 0.01
            && self.deferred_shapes.is_empty()
            && self.footnote_ids.is_empty()
    }

    /// Where a paragraph's floats anchor: its slot top, except that a page's
    /// first paragraph pushed down by header floats anchors where the header's
    /// text alone ends (Word probes under a 133pt floating letterhead: offsets
    /// 0 and +30pt land at the 56.7pt margin and 30pt below it, while the
    /// text sits under the letterhead; a header of flowing text moves both).
    pub(super) fn footer_top(&self, sp: &SectionProperties, ctx: &RenderContext) -> f32 {
        header_footer::footer_top(
            sp,
            self.is_first_page_of_section,
            self.page_num(ctx.sections),
            ctx,
        )
    }

    pub(super) fn anchor_slot_top(&self, sp: &SectionProperties, ctx: &RenderContext) -> f32 {
        match self.page_top_header(sp, ctx) {
            Some(is_first) => self.slot_top.max(header_footer::flow_slot_top(
                sp,
                is_first,
                self.page_num(ctx.sections),
                ctx,
            )),
            None => self.slot_top,
        }
    }

    /// Which header variant (first page or not) opened this page, when the
    /// cursor still sits at its top: the one whose body top matches.
    fn page_top_header(&self, sp: &SectionProperties, ctx: &RenderContext) -> Option<bool> {
        if !self.is_at_page_top(sp) {
            return None;
        }
        let page = self.page_num(ctx.sections);
        [true, false].into_iter().find(|&is_first| {
            (effective_slot_top(sp, is_first, page, ctx) - self.page_top_y).abs() < 0.01
        })
    }

    /// Where a table spanning `left..right` starts: a page-top table drops
    /// below the header's wrapping floats it would overlap, see
    /// `header_float_floor`.
    pub(super) fn below_header_floats(
        &self,
        sp: &SectionProperties,
        ctx: &RenderContext,
        left: f32,
        right: f32,
    ) -> f32 {
        self.page_top_header(sp, ctx)
            .and_then(|is_first| {
                header_footer::header_float_floor(
                    sp,
                    is_first,
                    self.page_num(ctx.sections),
                    ctx,
                    left,
                    right,
                )
            })
            .map_or(self.slot_top, |floor| self.slot_top.min(floor))
    }

    pub(super) fn is_at_page_top(&self, sp: &SectionProperties) -> bool {
        // nabl's 47.49pt header ends below its 36pt top margin: Word still
        // drops the space before the first paragraph under it.
        (self.slot_top - (sp.page_height - sp.margin_top)).abs() < 1.0
            || (self.slot_top - self.page_top_y).abs() < 0.01
    }

    /// A fresh page always takes a line, except a balanced one whose short
    /// columns must make the trial fail instead.
    pub(super) fn takes_a_line(&self, sp: &SectionProperties) -> bool {
        self.is_at_page_top(sp)
            && self
                .balance_floor
                .is_none_or(|(page, _)| page != self.page_count())
    }

    /// The gap above a block at the top of a new page (None elsewhere): none,
    /// except on a section's first page, where the block's space before
    /// collapses with the previous section's trailing space after.
    pub(super) fn page_top_gap(
        &self,
        sp: &SectionProperties,
        space_before: f32,
        prev_space_after: f32,
    ) -> Option<f32> {
        (!self.all_contents.is_empty() && self.is_at_page_top(sp)).then(|| {
            if self.is_first_page_of_section {
                own_space_before(space_before, prev_space_after)
            } else {
                0.0
            }
        })
    }

    /// Text that didn't fit: on to the next column, but when this column is
    /// still empty the others are no taller, so to the next page (case80's
    /// five-column region starts on page 2). The column's separator reaches
    /// the `trailing_space` after its last paragraph (case80 page 5).
    #[allow(clippy::too_many_arguments)]
    pub(super) fn overflow_column_or_page(
        &mut self,
        current_col: &mut usize,
        col_count: usize,
        sect_idx: usize,
        sp: &SectionProperties,
        effective_margin_bottom: &mut f32,
        ctx: &RenderContext,
        trailing_space: f32,
    ) {
        let col_count = if self.slot_top >= self.column_top_y - 0.01 {
            1
        } else {
            self.slot_top -= trailing_space;
            col_count
        };
        self.advance_column_or_page(
            current_col,
            col_count,
            sect_idx,
            sp,
            effective_margin_bottom,
            ctx,
        );
    }

    /// Advance to the next column if available, otherwise flush the current page.
    pub(super) fn advance_column_or_page(
        &mut self,
        current_col: &mut usize,
        col_count: usize,
        sect_idx: usize,
        sp: &SectionProperties,
        effective_margin_bottom: &mut f32,
        ctx: &RenderContext,
    ) {
        if *current_col + 1 < col_count {
            *current_col += 1;
            self.last_col = self.last_col.max(*current_col);
            self.col_bottom = self.col_bottom.min(self.slot_top);
            self.col_heights += self.column_top_y - self.slot_top;
            *effective_margin_bottom -= self.col_fn_reserved;
            self.col_fn_reserved = 0.0;
            self.slot_top = self.column_top_y;
            self.pending_float_anchor = None;
        } else {
            *current_col = 0;
            self.begin_next_page(sect_idx, sp, effective_margin_bottom, ctx);
            if let Some((page, y)) = self.balance_floor
                && page == self.page_count()
            {
                *effective_margin_bottom = effective_margin_bottom.max(y);
            }
        }
    }

    /// The text after a full-width picture resumes below the lowest one.
    pub(super) fn lower_next_text_to(&mut self, bottom: f32) {
        self.next_text_below_picture =
            Some(self.next_text_below_picture.map_or(bottom, |old| old.min(bottom)));
    }

    fn take_pending_first_page(&mut self, sect_idx: usize) -> bool {
        self.pending_section_first_page
            .take()
            .is_some_and(|(section, page, y)| {
                section == sect_idx && page == self.page_count() && (self.slot_top - y).abs() < 0.01
            })
    }

    /// A table opening the section's pending first page moves the cursor
    /// from `from` to its position, space before and top border; nothing has
    /// flowed yet.
    pub(super) fn keep_pending_first_page(&mut self, from: f32) {
        if let Some((_, _, y)) = &mut self.pending_section_first_page
            && (*y - from).abs() < 0.01
        {
            *y = self.slot_top;
        }
    }

    /// Footnotes carried past the end of a section get a page of their own
    /// under the first-page header before the next section starts (Word
    /// examples: a 630pt row whose note does not fit ends section 1;
    /// the note fills page 2 alone under section 1's first header for a
    /// next-page break, section 2's for a continuous one, whose text starts
    /// page 3).
    pub(super) fn flush_note_page(&mut self, sect_idx: usize) {
        if !self.footnote_ids.is_empty() {
            self.is_first_page_of_section = true;
            self.flush_page(sect_idx);
        }
    }

    /// Flush this page and open the next one of section `sp`.
    pub(super) fn begin_next_page(
        &mut self,
        sect_idx: usize,
        sp: &SectionProperties,
        effective_margin_bottom: &mut f32,
        ctx: &RenderContext,
    ) {
        let first_page = self.take_pending_first_page(sect_idx);
        // A section that moved off the page in full leaves it to the one before:
        // the page neither counts for it nor starts its numbering.
        let ends_in = if first_page { sect_idx - 1 } else { sect_idx };
        self.flush_page(ends_in);
        self.page_hf_section = sect_idx;
        if first_page {
            self.flush_note_page(sect_idx);
        }
        let page = self.page_num(ctx.sections);
        self.slot_top = effective_slot_top(sp, first_page, page, ctx);
        self.page_top_y = self.slot_top;
        if let Some((left, right)) = self.table_extent {
            self.slot_top = self.below_header_floats(sp, ctx, left, right);
        }
        self.column_top_y = self.slot_top;
        self.page_top_y = self.slot_top;
        *effective_margin_bottom =
            compute_effective_margin_bottom(sp, first_page, page, ctx) + self.col_fn_reserved;
        self.is_first_page_of_section = first_page;
        let bands = hf_page_bands(sp, first_page, page, *effective_margin_bottom, ctx);
        add_bands(&mut self.frame_bands, bands);
    }
}

/// Bundles the mutable render loop state so it can be passed to extracted functions.
pub(super) struct LayoutState {
    pub(super) pb: PageBuilder,
    pub(super) prev_space_after: f32,
    pub(super) prev_between_below: f32,
    pub(super) effective_margin_bottom: f32,
    pub(super) current_col: usize,
    pub(super) global_block_idx: usize,
    pub(super) heading_entries: Vec<HeadingEntry>,
    pub(super) bookmark_positions: HashMap<String, (usize, f32)>,
    /// The keep-with-next chain being laid out is too long for a page: its
    /// paragraphs keep only their own link to the next.
    pub(super) long_keep_chain: bool,
    /// The gap above a paragraph laid out again in the column it overflowed
    /// into: the overflow's, not its spacing recomputed.
    pub(super) reflowed: bool,
    /// The page's upcoming float bands are collected ahead of its lines
    /// (`prescan_page_bands`): a single-column section with bands, outside
    /// frames.
    pub(super) band_prescan: bool,
    /// The next paragraph's page break before was taken ahead of the page's
    /// float bands, which may have moved it off the page top since.
    pub(super) page_break_taken: bool,
    /// A frame in the flow runs past the bottom, or moved on to a column
    /// with more lines than it left in the one before: a balancing trial at
    /// this height fails (Word examples: a frame of four lines after one
    /// stays in column 1, fixture frame-balanced-columns; six lines after
    /// seven open column 2, frame-balanced-opens-column).
    pub(super) frame_overflow: bool,
}

impl LayoutState {
    pub(super) fn advance_column_or_page(
        &mut self,
        col_count: usize,
        sect_idx: usize,
        sp: &SectionProperties,
        ctx: &RenderContext,
    ) {
        self.pb.advance_column_or_page(
            &mut self.current_col,
            col_count,
            sect_idx,
            sp,
            &mut self.effective_margin_bottom,
            ctx,
        );
    }

    pub(super) fn begin_next_page(&mut self, sect_idx: usize, sp: &SectionProperties, ctx: &RenderContext) {
        self.pb
            .begin_next_page(sect_idx, sp, &mut self.effective_margin_bottom, ctx);
    }

    pub(super) fn take_page_break_after(
        &mut self,
        para: &Paragraph,
        sect_idx: usize,
        sp: &SectionProperties,
        ctx: &RenderContext,
    ) {
        if para.page_break_after {
            self.begin_next_page(sect_idx, sp, ctx);
            self.prev_space_after = 0.0;
            self.current_col = 0;
        }
    }

    /// A throwaway copy for laying out a column region again with its
    /// columns ending at `bottom`; its output is discarded.
    pub(super) fn trial(&self, bottom: f32) -> LayoutState {
        let mut pb = PageBuilder::new(self.pb.slot_top);
        pb.column_top_y = self.pb.column_top_y;
        pb.page_top_y = self.pb.page_top_y;
        pb.is_first_page_of_section = self.pb.is_first_page_of_section;
        pb.page_hf_section = self.pb.page_hf_section;
        pb.page_section_indices = self.pb.page_section_indices.clone();
        pb.float_zone = self.pb.float_zone.clone();
        pb.more_float_zones = self.pb.more_float_zones.clone();
        pb.pending_float_anchor = self.pb.pending_float_anchor;
        pb.next_text_below_picture = self.pb.next_text_below_picture;
        pb.footnote_ids_set = self.pb.footnote_ids_set.clone();
        pb.col_fn_reserved = self.pb.col_fn_reserved;
        pb.line_number_counter = self.pb.line_number_counter;
        pb.line_numbers_restart_each_page = self.pb.line_numbers_restart_each_page;
        // Page-top rules look at whether a page was already flushed.
        if !self.pb.all_contents.is_empty() {
            pb.all_contents.push(Content::new());
        }
        LayoutState {
            pb,
            prev_space_after: self.prev_space_after,
            prev_between_below: self.prev_between_below,
            effective_margin_bottom: bottom,
            current_col: self.current_col,
            global_block_idx: self.global_block_idx,
            heading_entries: Vec::new(),
            bookmark_positions: self.bookmark_positions.clone(),
            long_keep_chain: self.long_keep_chain,
            reflowed: false,
            band_prescan: self.band_prescan,
            page_break_taken: false,
            frame_overflow: false,
        }
    }
}

/// Each page's number and number format (None = decimal), each section's
/// page count, and the page count.
pub(crate) struct BodyPageFields {
    pub(super) pages: Vec<(usize, Option<String>)>,
    pub(super) section_pages: Vec<usize>,
    pub(super) total: usize,
}

/// How many pages each section is on (SECTIONPAGES). A page where a
/// continuous section starts counts for both sections.
pub(super) fn section_page_counts(sections: usize, pages: &[(usize, bool, usize)]) -> Vec<usize> {
    let mut counts = vec![0; sections];
    for &(first, _, last) in pages {
        for count in &mut counts[first..=last] {
            *count += 1;
        }
    }
    counts
}

/// The number the last of `page_section_indices` shows (`page_numbers`), 1
/// for none.
pub(super) fn last_page_number<'a>(
    sections: &[crate::model::Section],
    page_section_indices: impl IntoIterator<Item = &'a (usize, bool, usize)>,
) -> usize {
    page_numbers(sections, page_section_indices)
        .last()
        .copied()
        .unwrap_or(1)
}

/// Logical page number of each page: a section with `w:pgNumType @start` restarts,
/// others continue. A restart after an odd/even section break with the wrong
/// parity takes the next number instead (Word skips a number, not a sheet).
/// A page shows the number of the section at its top; a restart in a section
/// that begins further down counts that page as its start, so the next page
/// shows start + 1 (Word probes, 2026-10-03).
pub(super) fn page_numbers<'a>(
    sections: &[crate::model::Section],
    page_section_indices: impl IntoIterator<Item = &'a (usize, bool, usize)>,
) -> Vec<usize> {
    let mut numbers = Vec::new();
    let mut running = 0;
    let mut prev_si = None;
    for &(si, _, last_si) in page_section_indices {
        let sp = &sections[si].properties;
        running = match sp.page_num_start {
            Some(start) if prev_si != Some(si) => {
                let start = start as usize;
                let wrong_parity = si > 0
                    && match sp.break_type {
                        SectionBreakType::OddPage => start % 2 == 0,
                        SectionBreakType::EvenPage => start % 2 == 1,
                        _ => false,
                    };
                start + wrong_parity as usize
            }
            _ => running + 1,
        };
        numbers.push(running);
        if let Some(start) = (si + 1..=last_si)
            .filter_map(|s| sections[s].properties.page_num_start)
            .last()
        {
            running = start as usize;
        }
        prev_si = Some(last_si);
    }
    numbers
}

/// Each page's logical number and number format (None = decimal): the format
/// of the section a page ends in, which never inherits from the one before
/// (OOXML §17.6.12; Word renders arabic after roman front matter).
pub(super) fn page_number_formats<'a>(
    doc: &'a Document,
    page_section_indices: &[(usize, bool, usize)],
) -> Vec<(usize, Option<&'a str>)> {
    page_numbers(&doc.sections, page_section_indices)
        .into_iter()
        .zip(page_section_indices)
        .map(|(n, &(_, _, si))| (n, doc.sections[si].properties.page_num_format.as_deref()))
        .collect()
}

/// The docGrid pitch table-cell lines snap to in a section: only under
/// `w:compat/w:adjustLineHeightInTable`, else 0.
/// How an empty, zero-height section-break paragraph passes spacing on: the
/// drop below the cursor and the space after left pending. Before a
/// continuous section the paragraph before it keeps its whole space after,
/// and the break's own space after only absorbs that much of the next
/// paragraph's space before (Word probes: 12pt after, 8pt on the break, 10pt
/// before give 14pt; covid_insomnia's two columns start 12pt below its
/// keywords). Before a new page the break's own space after counts.
pub(super) fn section_break_spacing(prev_after: f32, break_after: f32, next_continuous: bool) -> (f32, f32) {
    if next_continuous {
        (prev_after - break_after, break_after)
    } else {
        (0.0, break_after)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn continuous_first_page_is_pending_only_before_the_first_block_flows() {
        let mut page = PageBuilder::new(97.9);
        page.pending_section_first_page = Some((2, 0, 97.9));
        assert!(page.take_pending_first_page(2));
        assert!(!page.take_pending_first_page(2));
        page.pending_section_first_page = Some((2, 0, 97.9));
        page.slot_top -= 16.0; // A first paragraph has flowed on this sheet.
        assert!(!page.take_pending_first_page(2));
        page.slot_top = 97.9;
        page.pending_section_first_page = Some((2, 0, 97.9));
        assert!(!page.take_pending_first_page(3));
        page.pending_section_first_page = Some((2, 0, 97.9));
        page.all_contents.push(Content::new());
        assert!(!page.take_pending_first_page(2));
    }
}
