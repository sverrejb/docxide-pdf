//! Tagged PDF (PDF/UA): a structure tree built while rendering and written at
//! assembly.
//!
//! Body page streams are artifacts by default: each opens `/Artifact BMC` when
//! it is created and assembly closes it. A tagged region closes the artifact,
//! opens its own marked content and reopens the artifact when it ends, so
//! everything drawn between regions (shading, borders, rules, separators) is
//! never untagged content.

use std::collections::HashMap;

use pdf_writer::types::{ListNumbering, TableHeaderScope};
use pdf_writer::writers::StructTreeRoot;
use pdf_writer::{Content, Name, Pdf, Ref, Str, TextStr};

use crate::model::ListItem;

pub(super) const ROOT: usize = 0;

enum Kid {
    Node(usize),
    Mcid { page: usize, mcid: i32 },
}

#[derive(Default)]
struct Node {
    kind: &'static str,
    parent: usize,
    kids: Vec<Kid>,
    /// Table cell attributes: header scope and column span.
    cell: Option<(Option<TableHeaderScope>, i32)>,
    alt: Option<String>,
    /// Structure element ID; PDF/UA requires one on every Note.
    id: Option<String>,
    /// `/ActualText`: the source letters of a caps/small-caps Span.
    actual: Option<String>,
    /// `/Lang` of a Span in another language than the document's.
    lang: Option<String>,
    /// An L's `/ListNumbering`.
    numbering: Option<ListNumbering>,
}

impl Node {
    fn new(kind: &'static str, parent: usize) -> Self {
        Self {
            kind,
            parent,
            ..Default::default()
        }
    }

    fn child_nodes(&self) -> impl DoubleEndedIterator<Item = usize> + '_ {
        self.kids.iter().filter_map(|k| match k {
            Kid::Node(c) => Some(*c),
            Kid::Mcid { .. } => None,
        })
    }
}

pub(crate) struct Tags {
    nodes: Vec<Node>,
    /// Marked-content ids are unique per page.
    next_mcid: Vec<i32>,
    /// (endnote?, note id) → its Note element.
    notes: HashMap<(bool, u32), usize>,
    /// Document-level elements for floating content (textboxes, pictures), drawn while
    /// their anchor paragraph renders but placed after its element, like Word.
    hoisted: Vec<(u32, usize)>,
    /// The document's language (the catalog `/Lang`); text in another one
    /// gets a Span with `/Lang`.
    pub(super) lang: String,
}

/// Open list levels for L/LI nesting. Word nests a deeper level's L inside the
/// LBody of the item above it and ends the list at any non-list block.
#[derive(Default)]
pub(crate) struct Lists {
    /// (level, its L, LBody of that level's latest item)
    stack: Vec<(u8, usize, usize)>,
    id: Option<u32>,
}

impl Lists {
    pub(super) fn close(&mut self) {
        self.stack.clear();
        self.id = None;
    }
}

/// Structure of the table being rendered. A row split across pages looks its
/// TR/TH/TD/P up again, so a cell paragraph continued on the next page stays
/// one element. Repeated header rows are drawn with tagging suspended.
#[derive(Default)]
pub(crate) struct TableTags {
    table: usize,
    head_rows: usize,
    first_col_header: bool,
    head: Option<usize>,
    body: Option<usize>,
    rows: HashMap<usize, usize>,
    cells: HashMap<(usize, usize), usize>,
    paras: HashMap<(usize, usize, usize), usize>,
    lists: HashMap<(usize, usize), Lists>,
    /// Tables nested in a cell, by (row, cell, item), kept between the pieces
    /// of a row split across pages.
    nested: HashMap<(usize, usize, usize), TableTags>,
}

impl TableTags {
    /// Word puts the repeated `tblHeader` rows (else the first row, when
    /// tblLook marks it) in THead as TH cells, and makes the first cell of the
    /// other rows a TH when the first column is a header.
    pub(super) fn for_table(tags: &mut Tags, parent: usize, table: &crate::model::Table) -> Self {
        let repeated = table.rows.iter().take_while(|r| r.is_header).count();
        Self {
            table: tags.add(parent, "Table"),
            head_rows: if repeated > 0 {
                repeated
            } else {
                usize::from(table.header_first_row)
            },
            first_col_header: table.header_first_col,
            ..Default::default()
        }
    }

    /// Word gives a table whose rows are all headers an empty TBody (PDF/UA 7.2-14).
    pub(super) fn finish(&mut self, tags: &mut Tags) {
        if self.head.is_some() && self.body.is_none() {
            self.body = Some(tags.add(self.table, "TBody"));
        }
    }

    fn cell(&mut self, tags: &mut Tags, row: usize, cell: usize, col_span: i32) -> usize {
        if let Some(&c) = self.cells.get(&(row, cell)) {
            return c;
        }
        let in_head = row < self.head_rows;
        let tr = match self.rows.get(&row) {
            Some(&tr) => tr,
            None => {
                let table = self.table;
                let (slot, kind) = if in_head {
                    (&mut self.head, "THead")
                } else {
                    (&mut self.body, "TBody")
                };
                let section = *slot.get_or_insert_with(|| tags.add(table, kind));
                let tr = tags.add(section, "TR");
                self.rows.insert(row, tr);
                tr
            }
        };
        // Word leaves TH without /Scope (PDF/UA 7.5-1); the scope is known here.
        let (kind, scope) = if in_head {
            ("TH", Some(TableHeaderScope::Column))
        } else if cell == 0 && self.first_col_header {
            ("TH", Some(TableHeaderScope::Row))
        } else {
            ("TD", None)
        };
        let c = tags.add(tr, kind);
        tags.nodes[c].cell = Some((scope, col_span));
        self.cells.insert((row, cell), c);
        c
    }

    /// The cell paragraph's elements: a P, or for a list item its Lbl (when
    /// the label is drawn separately) and LBody, with the cell's own L/LI
    /// nesting. A continued paragraph gets its body back.
    fn para(
        &mut self,
        tags: &mut Tags,
        (row, cell, col_span): (usize, usize, i32),
        item: usize,
        list_item: Option<ListItem>,
        labelled: bool,
        kind: &'static str,
    ) -> (Option<usize>, usize) {
        if let Some(&p) = self.paras.get(&(row, cell, item)) {
            return (None, p);
        }
        let c = self.cell(tags, row, cell, col_span);
        let lists = self.lists.entry((row, cell)).or_default();
        let nodes = tags.para_nodes(lists, c, list_item, labelled, kind);
        self.paras.insert((row, cell, item), nodes.1);
        nodes
    }
}

/// Tags one cell's paragraphs as they are drawn into a body stream.
pub(super) struct CellTagger<'a> {
    pub(super) tags: &'a mut Tags,
    pub(super) table: &'a mut TableTags,
    pub(super) page: usize,
    pub(super) row: usize,
    pub(super) cell: usize,
    pub(super) col_span: i32,
}

impl CellTagger<'_> {
    /// Open the cell paragraph's text element (see `TableTags::para`);
    /// returns (Lbl, text element) for the label, links and notes inside.
    pub(super) fn begin(
        &mut self,
        content: &mut Content,
        item: usize,
        list_item: Option<ListItem>,
        labelled: bool,
        kind: &'static str,
    ) -> (Option<usize>, usize) {
        let at = (self.row, self.cell, self.col_span);
        let nodes = self
            .table
            .para(self.tags, at, item, list_item, labelled, kind);
        self.tags.begin(content, self.page, nodes.1);
        nodes
    }

    /// Continue the open content as `node`'s (e.g. the list label's Lbl).
    pub(super) fn switch(&mut self, content: &mut Content, node: usize) {
        self.tags.begin(content, self.page, node);
    }

    /// Structure for a table nested in this cell: a Table inside the cell
    /// element, between the paragraphs around it. It ends the cell's list.
    /// The table at `item` carries on in the Table its first piece started
    /// when its row is split across pages, so a row split partway stays one
    /// TR with all its cells (7.2-42). Returns (tags, the nested table's
    /// structure, page), as `render_table_rows` takes them.
    pub(super) fn nested_table(
        &mut self,
        table: &crate::model::Table,
        item: usize,
    ) -> (&mut Tags, &mut TableTags, usize) {
        let cell = self.close_list();
        let tags = &mut *self.tags;
        let nested = self
            .table
            .nested
            .entry((self.row, self.cell, item))
            .or_insert_with(|| TableTags::for_table(tags, cell, table));
        (tags, nested, self.page)
    }

    /// A Figure inside the cell element for a picture floating in it, after
    /// its anchor paragraph, as Word tags one. It ends the cell's list.
    pub(super) fn figure(&mut self, alt: Option<&str>) -> usize {
        let cell = self.close_list();
        self.tags.add_figure(cell, alt)
    }

    /// A Sect inside the cell element for a textbox anchored in it, after its
    /// anchor paragraph (Word: `TD > Sect > P`). It ends the cell's list.
    pub(super) fn sect(&mut self) -> usize {
        let cell = self.close_list();
        self.tags.add(cell, "Sect")
    }

    /// End the cell's list and return the cell element.
    fn close_list(&mut self) -> usize {
        if let Some(lists) = self.table.lists.get_mut(&(self.row, self.cell)) {
            lists.close();
        }
        self.table
            .cell(self.tags, self.row, self.cell, self.col_span)
    }

    /// The cell element alone: Word keeps a vertically merged cell's
    /// continuation as an empty cell so every row has all its columns.
    pub(super) fn empty_cell(&mut self) {
        self.table
            .cell(self.tags, self.row, self.cell, self.col_span);
    }

    /// A cell with nothing tagged inside still gets its element with an empty
    /// P, like Word, so every row has all its columns.
    pub(super) fn empty_para(mut self, content: &mut Content) {
        self.begin(content, 0, None, false, "P");
        Tags::end(content);
    }
}

/// Tags the footnotes (or endnotes) drawn into a body page's stream.
pub(super) struct NoteTagger<'a> {
    pub(super) tags: &'a mut Tags,
    pub(super) page: usize,
    pub(super) endnote: bool,
    /// The page's link annotations, for links in the note text.
    pub(super) links: &'a mut Vec<super::layout::LinkAnnotation>,
}

/// The language part of a language tag ("en" of "en-GB").
pub(super) fn primary_subtag(lang: &str) -> &str {
    lang.split('-').next().unwrap_or(lang)
}

/// A fresh body page stream, inside the default artifact.
pub(super) fn artifact_content() -> Content {
    let mut content = Content::new();
    content.begin_marked_content(Name(b"Artifact"));
    content
}

/// Wrap a finished stream that holds no real content (headers, footers, page
/// borders, the comment pane) as one artifact.
pub(super) fn wrap_artifact(out: &mut Vec<u8>, raw: &[u8], pagination: bool) {
    out.extend_from_slice(if pagination {
        b"/Artifact <</Type /Pagination>> BDC\n"
    } else {
        b"/Artifact BMC\n"
    });
    out.extend_from_slice(raw);
    out.extend_from_slice(b"\nEMC\n");
}

/// Close a stream opened by `artifact_content`, dropping the `/Artifact BMC
/// EMC` pairs left with nothing inside when tagged regions follow each other.
pub(super) fn finish_artifact_content(content: Content) -> Vec<u8> {
    const EMPTY: &[u8] = b"/Artifact BMC\nEMC\n";
    let mut raw = content.finish().into_vec();
    raw.extend_from_slice(b"\nEMC\n");
    let mut out = Vec::with_capacity(raw.len());
    let mut rest = &raw[..];
    while let Some(at) = rest.windows(EMPTY.len()).position(|w| w == EMPTY) {
        out.extend_from_slice(&rest[..at]);
        rest = &rest[at + EMPTY.len()..];
    }
    out.extend_from_slice(rest);
    out
}

impl Tags {
    pub(super) fn new() -> Self {
        Self {
            nodes: vec![Node::new("Document", ROOT)],
            next_mcid: Vec::new(),
            notes: HashMap::new(),
            hoisted: Vec::new(),
            // Word's own default when a document declares no language.
            lang: "en-US".to_string(),
        }
    }

    pub(super) fn is_empty(&self) -> bool {
        self.nodes.len() == 1
    }

    fn push(&mut self, kind: &'static str, parent: usize) -> usize {
        self.nodes.push(Node::new(kind, parent));
        self.nodes.len() - 1
    }

    pub(super) fn add(&mut self, parent: usize, kind: &'static str) -> usize {
        let id = self.push(kind, parent);
        self.nodes[parent].kids.push(Kid::Node(id));
        id
    }

    /// A document-level element that joins the tree at `attach_hoisted`, in
    /// `anchor_seq` order.
    pub(super) fn hoist(&mut self, kind: &'static str, anchor_seq: u32) -> usize {
        let id = self.push(kind, ROOT);
        self.hoisted.push((anchor_seq, id));
        id
    }

    /// A hoisted Figure with the picture's alt text when it has one.
    pub(super) fn hoist_figure(&mut self, alt: Option<&str>, anchor_seq: u32) -> usize {
        let id = self.hoist("Figure", anchor_seq);
        self.nodes[id].alt = alt.map(str::to_string);
        id
    }

    /// Place the hoisted elements after everything their anchor added. They
    /// were drawn layer by layer (behind-text first, pictures before
    /// textboxes); Word orders them like their anchors in the XML.
    pub(super) fn attach_hoisted(&mut self) {
        self.hoisted.sort_by_key(|&(seq, _)| seq);
        self.nodes[ROOT]
            .kids
            .extend(self.hoisted.drain(..).map(|(_, id)| Kid::Node(id)));
    }

    /// The Note for a footnote (or endnote), created under `parent` the first
    /// time it is seen: Word nests a note where its reference mark sits. Notes
    /// drawn before any tagged reference fall back to the document level.
    pub(super) fn note(&mut self, endnote: bool, id: u32, parent: usize) -> usize {
        if let Some(&n) = self.notes.get(&(endnote, id)) {
            return n;
        }
        let n = self.add(parent, "Note");
        self.nodes[n].id = Some(format!("Note {}", self.notes.len() + 1));
        self.notes.insert((endnote, id), n);
        n
    }

    /// A Figure under `parent`, with the picture's alt text when it has one.
    pub(super) fn add_figure(&mut self, parent: usize, alt: Option<&str>) -> usize {
        let id = self.add(parent, "Figure");
        self.nodes[id].alt = alt.map(str::to_string);
        id
    }

    /// A Formula under `parent`, read by its `/Alt` (PDF/UA 7.7-1); a
    /// paragraph that is itself the Formula (`sole_formula`) takes the `/Alt`.
    pub(super) fn add_formula(&mut self, parent: usize, alt: &str) -> usize {
        let id = if self.nodes[parent].kind == "Formula" {
            parent
        } else {
            self.add(parent, "Formula")
        };
        self.nodes[id].alt = Some(alt.to_string());
        id
    }

    /// A Span with `/Lang` for text in another language, and/or `/ActualText`
    /// replacing its glyphs for text extraction and screen readers;
    /// `push_actual` adds to the latter as the Span's content is drawn.
    pub(super) fn add_span(
        &mut self,
        parent: usize,
        lang: Option<&str>,
        actual: Option<&str>,
    ) -> usize {
        let span = self.add(parent, "Span");
        self.nodes[span].lang = lang.map(str::to_string);
        self.nodes[span].actual = actual.map(str::to_string);
        span
    }

    pub(super) fn push_actual(&mut self, node: usize, text: &str) {
        self.nodes[node]
            .actual
            .get_or_insert_with(String::new)
            .push_str(text);
    }

    /// True when `lang` is the document's language, by primary subtag: en-GB
    /// text in an en-US document needs no Span, French text does.
    pub(super) fn is_document_lang(&self, lang: &str) -> bool {
        primary_subtag(&self.lang).eq_ignore_ascii_case(primary_subtag(lang))
    }

    /// A paragraph's elements under `parent`: for a list item (see
    /// `list_item`) its (Lbl, LBody), else a `kind` element, which ends the
    /// open list.
    pub(super) fn para_nodes(
        &mut self,
        lists: &mut Lists,
        parent: usize,
        list_item: Option<ListItem>,
        labelled: bool,
        kind: &'static str,
    ) -> (Option<usize>, usize) {
        match list_item {
            Some(item) => self.list_item(lists, parent, item, labelled),
            None => {
                lists.close();
                (None, self.add(parent, kind))
            }
        }
    }

    /// LI for the item's list and level (a new L under `parent`, with the
    /// level's numbering, when the list starts); returns (Lbl when the label
    /// is drawn separately, LBody).
    pub(super) fn list_item(
        &mut self,
        lists: &mut Lists,
        parent: usize,
        item: ListItem,
        labelled: bool,
    ) -> (Option<usize>, usize) {
        if lists.id != Some(item.list_id) {
            lists.close();
            lists.id = Some(item.list_id);
        }
        while lists.stack.last().is_some_and(|&(l, ..)| l > item.level) {
            lists.stack.pop();
        }
        let list = match lists.stack.last() {
            Some(&(l, list, _)) if l == item.level => {
                lists.stack.pop();
                list
            }
            top => {
                let list = self.add(top.map_or(parent, |&(.., body)| body), "L");
                self.nodes[list].numbering = Some(item.numbering);
                list
            }
        };
        let li = self.add(list, "LI");
        let label = labelled.then(|| self.add(li, "Lbl"));
        let body = self.add(li, "LBody");
        lists.stack.push((item.level, list, body));
        (label, body)
    }

    /// Open a paragraph's elements from `para_nodes`: the label is drawn in
    /// its Lbl, or in the text element when it has none; the text element
    /// stays open.
    pub(super) fn begin_para(
        &mut self,
        content: &mut Content,
        page: usize,
        (label, text): (Option<usize>, usize),
        draw_label: impl FnOnce(&mut Content),
    ) {
        self.begin(content, page, label.unwrap_or(text));
        draw_label(content);
        if label.is_some() {
            self.begin(content, page, text);
        }
    }

    /// Start a piece of `node`'s content on `page`; it runs until `end`.
    /// `content` must be a body stream (inside the default artifact).
    pub(super) fn begin(&mut self, content: &mut Content, page: usize, node: usize) {
        if self.next_mcid.len() <= page {
            self.next_mcid.resize(page + 1, 0);
        }
        let mcid = self.next_mcid[page];
        self.next_mcid[page] += 1;
        self.nodes[node].kids.push(Kid::Mcid { page, mcid });
        content.end_marked_content();
        content
            .begin_marked_content_with_properties(Name(self.nodes[node].kind.as_bytes()))
            .properties()
            .identify(mcid);
    }

    pub(super) fn end(content: &mut Content) {
        content.end_marked_content();
        content.begin_marked_content(Name(b"Artifact"));
    }

    /// The PDF/UA structure rules that depend on the document itself: every
    /// Figure has alt text (7.3-1), and in reading order the first heading is
    /// H1 and going deeper never skips a level (7.4.2-1).
    pub(super) fn ua_structure_ok(&self) -> bool {
        let mut level = 0;
        let mut stack = vec![ROOT];
        while let Some(n) = stack.pop() {
            let node = &self.nodes[n];
            if node.kind == "Figure" && node.alt.is_none() {
                return false;
            }
            if let Some(l) = node
                .kind
                .strip_prefix('H')
                .and_then(|d| d.parse::<u8>().ok())
            {
                if l > level + 1 {
                    return false;
                }
                level = l;
            }
            stack.extend(node.child_nodes().rev());
        }
        true
    }

    /// Page `p`'s `/StructParents` key, if it has tagged content.
    pub(super) fn struct_parents(&self, page: usize) -> Option<i32> {
        (self.next_mcid.get(page).copied().unwrap_or(0) > 0).then_some(page as i32)
    }

    /// Write the elements and parent tree; returns the StructTreeRoot ref.
    /// `annots`: (Link element, page, annotation) in `/StructParent` key order,
    /// keys starting after the pages' keys.
    pub(super) fn write(
        &self,
        pdf: &mut Pdf,
        alloc: &mut impl FnMut() -> Ref,
        page_ids: &[Ref],
        annots: &[(usize, usize, Ref)],
    ) -> Ref {
        // An unattached element would name a parent that doesn't list it.
        debug_assert!(self.hoisted.is_empty(), "hoisted elements never attached");
        let root = alloc();
        let refs: Vec<Ref> = self.nodes.iter().map(|_| alloc()).collect();
        let mut owners: Vec<Vec<Option<Ref>>> = self
            .next_mcid
            .iter()
            .map(|&n| vec![None; n as usize])
            .collect();
        let mut node_annots: Vec<Vec<(usize, Ref)>> = vec![Vec::new(); self.nodes.len()];
        for &(node, page, annot) in annots {
            node_annots[node].push((page, annot));
        }

        for (i, node) in self.nodes.iter().enumerate() {
            let mut elem = pdf.struct_element(refs[i]);
            elem.custom_kind(Name(node.kind.as_bytes()));
            elem.parent(if i == ROOT { root } else { refs[node.parent] });
            if let Some(alt) = &node.alt {
                elem.alt(TextStr(alt));
            }
            if let Some(actual) = &node.actual {
                elem.actual_text(TextStr(actual));
            }
            if let Some(lang) = &node.lang {
                elem.lang(TextStr(lang));
            }
            if let Some(id) = &node.id {
                elem.id(Str(id.as_bytes()));
            }
            if let Some((scope, col_span)) = node.cell.filter(|&(s, span)| s.is_some() || span > 1)
            {
                let mut attrs = elem.attributes();
                let mut table = attrs.push().table();
                if let Some(scope) = scope {
                    table.scope(scope);
                }
                if col_span > 1 {
                    table.col_span(col_span);
                }
            }
            // None is the default; writing it says nothing.
            if let Some(numbering) = node.numbering.filter(|&n| n != ListNumbering::None) {
                elem.attributes().push().list().list_numbering(numbering);
            }
            // Content on the element's own /Pg is a bare MCID; only a paragraph
            // continued on the next page needs full marked-content references.
            let first_page = node.kids.iter().find_map(|k| match k {
                Kid::Mcid { page, .. } => Some(*page),
                Kid::Node(_) => None,
            });
            if let Some(p) = first_page {
                elem.page(page_ids[p]);
            }
            let mut kids = elem.children();
            for kid in &node.kids {
                match *kid {
                    Kid::Node(c) => {
                        kids.struct_element(refs[c]);
                    }
                    Kid::Mcid { page, mcid } => {
                        owners[page][mcid as usize] = Some(refs[i]);
                        if Some(page) == first_page {
                            kids.marked_content_id(mcid);
                        } else {
                            kids.marked_content_ref()
                                .marked_content_id(mcid)
                                .page(page_ids[page]);
                        }
                    }
                }
            }
            for &(page, annot) in &node_annots[i] {
                kids.object_ref().object(annot).page(page_ids[page]);
            }
        }

        let arrays: Vec<(i32, Ref)> = owners
            .iter()
            .enumerate()
            .filter(|(_, o)| !o.is_empty())
            .map(|(page, o)| {
                let id = alloc();
                pdf.indirect(id)
                    .array()
                    .items(o.iter().map(|r| r.expect("every MCID has an owner")));
                (page as i32, id)
            })
            .collect();
        let mut tree = pdf.indirect(root).start::<StructTreeRoot>();
        tree.child(refs[ROOT]);
        {
            let mut parent_tree = tree.parent_tree();
            let mut nums = parent_tree.nums();
            for (key, id) in arrays {
                nums.insert(key, id);
            }
            for (k, &(node, ..)) in annots.iter().enumerate() {
                nums.insert((page_ids.len() + k) as i32, refs[node]);
            }
        }
        tree.parent_tree_next_key((page_ids.len() + annots.len()) as i32);
        root
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tagged_regions_leave_no_empty_artifacts() {
        let mut tags = Tags::new();
        let p = tags.add(ROOT, "P");
        let mut content = artifact_content();
        tags.begin(&mut content, 0, p);
        Tags::end(&mut content);
        content.rect(0.0, 0.0, 1.0, 1.0).fill_nonzero();
        tags.begin(&mut content, 0, p);
        Tags::end(&mut content);
        let out = String::from_utf8(finish_artifact_content(content)).unwrap();
        // The rectangle keeps its artifact; the empty ones around the tags go.
        assert_eq!(out.matches("/Artifact BMC").count(), 1);
        assert_eq!(
            out.matches("BMC").count() + out.matches("BDC").count(),
            out.matches("EMC").count()
        );
        assert_eq!(tags.struct_parents(0), Some(0));
    }

    #[test]
    fn hoisted_elements_follow_their_anchor() {
        let mut tags = Tags::new();
        let before = tags.add(ROOT, "P");
        // Drawn picture-first, but the textbox's anchor comes first in the XML.
        let figure = tags.hoist_figure(None, 1);
        let sect = tags.hoist("Sect", 0);
        tags.add(sect, "P");
        let anchor = tags.add(ROOT, "P");
        tags.attach_hoisted();
        let order: Vec<usize> = tags.nodes[ROOT].child_nodes().collect();
        assert_eq!(order, [before, anchor, sect, figure]);
        tags.attach_hoisted();
        assert_eq!(tags.nodes[ROOT].kids.len(), 4);
    }

    #[test]
    fn ua_structure_needs_alt_and_unskipped_headings() {
        let doc = |kinds: &[&'static str]| {
            let mut tags = Tags::new();
            let sect = tags.add(ROOT, "Sect");
            for &k in kinds {
                tags.add(sect, k);
            }
            tags
        };
        assert!(doc(&["H1", "P", "H2", "H3", "H1", "H2"]).ua_structure_ok());
        assert!(doc(&["P"]).ua_structure_ok());
        assert!(
            !doc(&["H2", "H3"]).ua_structure_ok(),
            "first heading must be H1"
        );
        assert!(!doc(&["H1", "H3"]).ua_structure_ok(), "skipped H2");
        let mut tags = doc(&["H1"]);
        let figure = tags.add_figure(ROOT, None);
        assert!(!tags.ua_structure_ok(), "Figure without alt");
        tags.nodes[figure].alt = Some("A chart".into());
        assert!(tags.ua_structure_ok());
    }

    #[test]
    fn list_items_nest_like_word() {
        let mut tags = Tags::new();
        let mut lists = Lists::default();
        let disc = ListNumbering::Disc;
        let item = |level, list_id, numbering| ListItem {
            level,
            list_id,
            numbering,
        };
        let (label, first_body) = tags.list_item(&mut lists, ROOT, item(0, 7, disc), true);
        tags.list_item(&mut lists, ROOT, item(1, 7, ListNumbering::Circle), false);
        tags.list_item(&mut lists, ROOT, item(0, 7, disc), true);
        tags.list_item(&mut lists, ROOT, item(0, 8, ListNumbering::Decimal), true);
        let kids = |n: usize| -> Vec<&str> {
            tags.nodes[n]
                .child_nodes()
                .map(|c| tags.nodes[c].kind)
                .collect()
        };
        // A new list id starts a new L; the sub-list sits in the first item's body.
        assert_eq!(kids(ROOT), ["L", "L"]);
        assert_eq!(kids(tags.nodes[label.unwrap()].parent), ["Lbl", "LBody"]);
        assert_eq!(kids(first_body), ["L"]);
        let first_list = tags.nodes[tags.nodes[first_body].parent].parent;
        assert_eq!(kids(first_list), ["LI", "LI"]);
        // Each L carries the numbering of the level that opened it.
        let sub_list = tags.nodes[first_body].child_nodes().next().unwrap();
        assert_eq!(tags.nodes[first_list].numbering, Some(disc));
        assert_eq!(tags.nodes[sub_list].numbering, Some(ListNumbering::Circle));
    }
}
