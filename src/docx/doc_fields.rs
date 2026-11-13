//! Passes over the parsed body: SEQ, TOC and STYLEREF fields, linked text boxes, grid pitch.

use std::collections::HashMap;
use std::io::Read;

use crate::model::{Block, Document, Section};

use super::{ParseContext, WML_NS, is_east_asian_char, numbering, paragraph, runs, styles};

/// A `linesAndChars` grid adds its `charSpace` to every character a run
/// snapped to it draws, Latin or East Asian, at any size (Word probes
/// 2026-11-08: +1pt under 4096, −0.49 under −2000; tokyo_welfare's 10.5pt
/// ideographs step 10.58). ponytail: body only; `snapToChars` snaps to
/// whole cells instead and is not modelled.
pub(super) fn grid_char_pitch(doc: &mut Document, grid_deltas: &[f32]) {
    for (section, &delta) in doc.sections.iter_mut().zip(grid_deltas) {
        let grid = section.properties.grid_type;
        if delta == 0.0 || grid != crate::model::DocGridType::LinesAndChars {
            continue;
        }
        let mut paras = Vec::new();
        body_paragraphs(&mut section.blocks, &mut paras);
        for run in paras.into_iter().flat_map(|p| p.runs.iter_mut()) {
            if run.snap_to_grid {
                run.char_spacing += delta;
            }
        }
    }
}

/// Under `balanceSingleByteDoubleByteWidth` Word spaces a full-width
/// character as two bytes: its run's `w:spacing` counts twice
/// (chinese_student_union: −1.6pt in 14pt text draws 10.75pt ideographs,
/// 12.5 without the flag; Word probes 2026-11-08). ponytail: body runs of
/// East Asian characters only; mixed runs and headers keep single spacing.
pub(super) fn double_wide_char_spacing(doc: &mut Document) {
    let mut paras = Vec::new();
    for section in &mut doc.sections {
        body_paragraphs(&mut section.blocks, &mut paras);
    }
    for run in paras.into_iter().flat_map(|p| p.runs.iter_mut()) {
        if !run.text.is_empty() && run.text.chars().all(is_east_asian_char) {
            run.char_spacing *= 2.0;
        }
    }
}

/// Number the body's SEQ fields in document order, one counter per name
/// (issue #145: a stale "old" cache shows 1). `\s n` restarts a counter
/// after a built-in "heading 1"…"heading n" paragraph that follows the
/// sequence's previous field, so a heading before an intervening `\c` does
/// not count; custom styles with an outline level never do (Word online
/// probes, 2026-11-03; issue #146). ponytail: text-box captions are not
/// counted.
pub(super) fn resolve_seq_fields(doc: &mut Document) {
    use crate::model::{FieldCode, SeqStep};
    let mut paras = Vec::new();
    for section in &mut doc.sections {
        body_paragraphs(&mut section.blocks, &mut paras);
    }
    // Per name: the count and the paragraph holding the last field.
    let mut counters: HashMap<String, (u32, usize)> = HashMap::new();
    // The latest built-in heading paragraph per 0-based level (0 = none,
    // harmless because `last` starts at 0 too).
    let mut last_heading = [0usize; 9];
    for (i, para) in paras.into_iter().enumerate() {
        if let Some(level) = para
            .style_id
            .as_deref()
            .and_then(|id| doc.style_id_to_name.get(id))
            .and_then(|name| styles::builtin_heading_level(name))
        {
            last_heading[level as usize] = i;
        }
        for run in &mut para.runs {
            let Some(FieldCode::Seq {
                name,
                format,
                step,
                hidden,
            }) = &run.field_code
            else {
                continue;
            };
            let (counter, last) = counters.entry(name.clone()).or_insert((0, 0));
            match step {
                SeqStep::Next => *counter += 1,
                // A repeat before any number shows the first.
                SeqStep::Repeat => *counter = (*counter).max(1),
                SeqStep::Restart(n) => *counter = *n,
                SeqStep::Heading(n) => {
                    let restarts = last_heading[..(*n as usize).min(9)]
                        .iter()
                        .any(|&at| at > *last);
                    *counter = if restarts { 1 } else { *counter + 1 };
                }
            }
            *last = i;
            run.text = if *hidden {
                String::new()
            } else {
                numbering::format_number(*counter, format.unwrap_or("decimal"))
            };
        }
    }
}

/// What a TOC field collects: `\o "a-b"` (and `\u`) the heading outline
/// levels, `\t "Style,level,…"` paragraphs of a style, `\c "Name"` paragraphs
/// holding a SEQ of that name; `\n` drops the page numbers. Without any of
/// them Word takes headings 1–3.
#[derive(Debug, PartialEq)]
struct TocSpec {
    outline: Option<std::ops::RangeInclusive<u8>>,
    styles: Vec<(String, u8)>,
    seq: Option<String>,
    page_numbers: bool,
    /// `\c`, with or without a name: a table of figures, whose entries all
    /// take the "table of figures" style instead of "toc N".
    figures: bool,
}

impl TocSpec {
    fn parse(instr: &str) -> TocSpec {
        let has = |s| runs::has_switch(instr, s);
        let mut outline = toc_switch_arg(instr, "\\o")
            .and_then(|r| r.split_once('-'))
            .and_then(|(a, b)| Some(a.trim().parse().ok()?..=b.trim().parse().ok()?))
            .or_else(|| (has("\\o") || has("\\u")).then_some(1..=9));
        let styles: Vec<(String, u8)> = toc_switch_arg(instr, "\\t")
            .map(|list| {
                let items: Vec<&str> = list.split([',', ';']).map(str::trim).collect();
                items
                    .chunks(2)
                    .filter_map(|c| Some((c.first()?.to_string(), c.get(1)?.parse().ok()?)))
                    .collect()
            })
            .unwrap_or_default();
        let seq = toc_switch_arg(instr, "\\c").map(str::to_string);
        if outline.is_none() && styles.is_empty() && seq.is_none() {
            outline = Some(1..=3);
        }
        TocSpec {
            outline,
            styles,
            seq,
            page_numbers: !has("\\n"),
            figures: has("\\c"),
        }
    }

    /// The TOC level `p` enters at, if any.
    fn level_of(&self, p: &crate::model::Paragraph, names: &HashMap<String, String>) -> Option<u8> {
        use crate::model::FieldCode;
        if let (Some(range), Some(level)) = (&self.outline, p.outline_level)
            && range.contains(&(level + 1))
        {
            return Some(level + 1);
        }
        if let Some(id) = &p.style_id
            && let Some((_, level)) = self.styles.iter().find(|(name, _)| {
                id.eq_ignore_ascii_case(name)
                    || names.get(id).is_some_and(|n| n.eq_ignore_ascii_case(name))
            })
        {
            return Some(*level);
        }
        let seq = self.seq.as_deref()?;
        p.runs
            .iter()
            .any(|r| matches!(&r.field_code, Some(FieldCode::Seq { name, .. }) if name.eq_ignore_ascii_case(seq)))
            .then_some(1)
    }
}

/// The argument after a TOC switch (`\o "1-3"`), if the switch is present
/// and followed by one (not by the next switch).
fn toc_switch_arg<'a>(instr: &'a str, switch: &'a str) -> Option<&'a str> {
    runs::switch_args(instr, switch)
        .next()
        .filter(|arg| !arg.starts_with('\\'))
}

fn xml_text(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
}

/// Regenerate the TOC fields whose cache is a single paragraph (issue #145's
/// "Update field in Word."), as Word's field update does: one "toc N"
/// paragraph per entry with its text, a dot-leader right tab 10 twips inside
/// the margin and a PAGEREF to the entry's paragraph (a bookmark added here),
/// then the field's own paragraph left empty. Later TOCs are expanded first
/// so the earlier ones' block indices hold. ponytail: `\p` separators, `\b`
/// scopes and TC fields are ignored; without a "toc N" style a level N entry
/// is indented 11pt × (N − 1).
pub(super) fn expand_toc_fields<R: Read + std::io::Seek>(
    sections: &mut [Section],
    ctx: &mut ParseContext<'_, R>,
) {
    let mut hosts = Vec::new();
    for (s, section) in sections.iter().enumerate() {
        for (b, block) in section.blocks.iter().enumerate() {
            if matches!(block, Block::Paragraph(p) if p.toc_field.is_some()) {
                hosts.push((s, b));
            }
        }
    }
    let names = &ctx.styles.style_id_to_name;
    let mut bookmarks = 0;
    for (s, b) in hosts.into_iter().rev() {
        let Block::Paragraph(host) = &mut sections[s].blocks[b] else {
            continue;
        };
        let spec = TocSpec::parse(&host.toc_field.take().unwrap_or_default());
        host.runs.clear();
        host.starts_toc_field = false;

        // (level, bookmark, list label, text) of each entry, in document order.
        let mut entries = Vec::new();
        let mut paras = Vec::new();
        for section in sections.iter_mut() {
            body_paragraphs(&mut section.blocks, &mut paras);
        }
        for p in paras {
            if p.starts_toc_field || p.toc_field.is_some() {
                continue;
            }
            let Some(level) = spec.level_of(p, names) else {
                continue;
            };
            let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
            let bookmark = format!("_Toc145_{bookmarks}");
            bookmarks += 1;
            p.bookmarks.push(bookmark.clone());
            let label = p.list_label.clone();
            entries.push((level, bookmark, label, text.trim().to_string()));
        }

        let sp = &sections[s].properties;
        let text_width = sp.page_width - sp.margin_left - sp.margin_right;
        let tab_pos = (text_width * 20.0).round() as i32 - 10;
        let mut xmls: Vec<String> = entries
            .iter()
            .map(|(level, bookmark, label, text)| {
                // A table of figures styles every entry "table of figures",
                // whatever its level; a contents table "toc N". Without the
                // style Word's stock definition applies: 24pt in (issue #146,
                // Windows Word 16.0.20527), 11pt × (N − 1) (unprobed).
                let (name, stock_indent) = if spec.figures {
                    ("table of figures".to_string(), 480)
                } else {
                    (format!("toc {level}"), 220 * (*level as i32 - 1))
                };
                let style_id =
                    names.iter().find(|(_, n)| n.eq_ignore_ascii_case(&name)).map(|(id, _)| id);
                let mut ppr = match style_id {
                    Some(id) => format!(r#"<w:pStyle w:val="{}"/>"#, xml_text(id)),
                    None if stock_indent > 0 => format!(r#"<w:ind w:left="{stock_indent}"/>"#),
                    None => String::new(),
                };
                let mut body = String::new();
                if !label.is_empty() {
                    body += &format!(r#"<w:r><w:t xml:space="preserve">{}</w:t></w:r><w:r><w:tab/></w:r>"#, xml_text(label));
                }
                body += &format!(r#"<w:r><w:t xml:space="preserve">{}</w:t></w:r>"#, xml_text(text));
                if spec.page_numbers {
                    ppr += &format!(r#"<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{tab_pos}"/></w:tabs>"#);
                    body += &format!(
                        r#"<w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF {bookmark} \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
                    );
                }
                format!(r#"<w:p xmlns:w="{WML_NS}"><w:pPr>{ppr}</w:pPr>{body}</w:p>"#)
            })
            .collect();
        if xmls.is_empty() {
            xmls.push(format!(
                r#"<w:p xmlns:w="{WML_NS}"><w:r><w:t>No table of contents entries found.</w:t></w:r></w:p>"#
            ));
        }
        let opts = paragraph::ParagraphOptions {
            resolve_bookmarks: false,
            resolve_outline_level: false,
            resolve_drawings: false,
            collect_extra_textboxes: false,
            style_num_id: None,
            style_num_ilvl: None,
        };
        let mut generated: Vec<Block> = xmls
            .iter()
            .filter_map(|xml| roxmltree::Document::parse(xml).ok())
            .map(|doc| {
                ctx.fields = Default::default();
                let mut lists = numbering::ListCounters::default();
                Block::Paragraph(paragraph::build_paragraph(
                    doc.root_element(),
                    ctx,
                    &mut lists,
                    &opts,
                ))
            })
            .collect();
        if let Some(Block::Paragraph(first)) = generated.first_mut() {
            first.starts_toc_field = true;
        }
        sections[s].blocks.splice(b..b, generated);
    }
}

/// Every body paragraph, table cells included, in document order.
fn body_paragraphs<'a>(blocks: &'a mut [Block], out: &mut Vec<&'a mut crate::model::Paragraph>) {
    for block in blocks {
        match block {
            Block::Paragraph(p) => out.push(p),
            Block::Table(t) => {
                for cell in t.rows.iter_mut().flat_map(|r| r.cells.iter_mut()) {
                    body_paragraphs(&mut cell.content, out);
                }
            }
        }
    }
}

/// Each box of a linked chain shares the story its first box holds and
/// learns the text areas of the boxes before it, from which the renderer works
/// out the part it shows. A box alone in its chain (a standalone `wps:txbx`
/// with an id) keeps its story as an unlinked box. ponytail: body text boxes
/// only, not headers or groups.
pub(super) fn link_textbox_stories(doc: &mut Document) {
    let mut paras = Vec::new();
    for section in &mut doc.sections {
        body_paragraphs(&mut section.blocks, &mut paras);
    }
    let mut boxes: Vec<(u32, u32, &mut crate::model::Textbox)> = paras
        .into_iter()
        .flat_map(|p| p.textboxes.iter_mut())
        .filter_map(|tb| Some((tb.link.as_ref()?.id, tb.link.as_ref()?.seq, tb)))
        .collect();
    boxes.sort_by_key(|&(id, seq, _)| (id, seq));
    for chain in boxes.chunk_by_mut(|a, b| a.0 == b.0) {
        if let [(_, _, tb)] = chain {
            tb.link = None;
            continue;
        }
        let story = std::sync::Arc::new(std::mem::take(&mut chain[0].2.paragraphs));
        let tables = std::sync::Arc::new(std::mem::take(&mut chain[0].2.tables));
        let mut before = Vec::new();
        for (_, _, tb) in chain.iter_mut() {
            let area = tb.text_area(tb.height_pt);
            if let Some(link) = &mut tb.link {
                link.story = story.clone();
                link.tables = tables.clone();
                link.before = before.clone();
                link.continues = true;
            }
            before.push(area);
        }
        if let Some(link) = chain.last_mut().and_then(|(_, _, tb)| tb.link.as_mut()) {
            link.continues = false;
        }
    }
}

/// Body STYLEREF fields show the nearest paragraph in the named style before
/// the field, else the first one after it (§17.16.5.59). Word updates them on
/// export, so the cached text is stale (issue #142: the body copy of a header's
/// STYLEREF kept "Heading pending"). Headers resolve theirs per page instead.
pub(super) fn resolve_body_stylerefs(doc: &mut Document) {
    use crate::model::{FieldCode, styleref_key};
    let names = std::mem::take(&mut doc.style_id_to_name);
    let mut paras = Vec::new();
    for section in &mut doc.sections {
        body_paragraphs(&mut section.blocks, &mut paras);
    }
    // Per paragraph, the keys it defines (paragraph style and character
    // styles, by id and by name) and their values.
    let defined: Vec<Vec<(String, String)>> = paras
        .iter()
        .map(|p| {
            let mut defs = Vec::new();
            let mut define = |id: &str, text: String| {
                for name in std::iter::once(id).chain(names.get(id).map(String::as_str)) {
                    for (form, value) in p.styleref_values(&text) {
                        defs.push((styleref_key(name, form), value.to_string()));
                    }
                }
            };
            if let Some(sid) = &p.style_id {
                let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
                if !text.is_empty() {
                    define(sid, text);
                }
            }
            for group in p.runs.chunk_by(|a, b| a.char_style_id == b.char_style_id) {
                if let Some(csid) = &group[0].char_style_id {
                    let text: String = group.iter().map(|r| r.text.as_str()).collect();
                    if !text.is_empty() {
                        define(csid, text);
                    }
                }
            }
            defs
        })
        .collect();
    let lookup = |key: &str, at: usize| {
        let before = (0..at).rev();
        let after = at + 1..defined.len();
        before.chain(after).find_map(|i| {
            defined[i]
                .iter()
                .find(|(k, _)| k == key)
                .map(|(_, v)| v.clone())
        })
    };
    for (i, p) in paras.iter_mut().enumerate() {
        for run in &mut p.runs {
            if let Some(FieldCode::StyleRef { name, form }) = run.field_code.clone()
                && let Some(text) = lookup(&styleref_key(&name, form), i)
            {
                run.text = text;
            }
        }
    }
    doc.style_id_to_name = names;
}

#[cfg(test)]
mod tests {
    #[test]
    fn toc_spec_reads_its_switches() {
        let spec = super::TocSpec::parse(r#" TOC \o "1-3" \h \z \u "#);
        assert_eq!((spec.outline, spec.page_numbers), (Some(1..=3), true));
        let spec = super::TocSpec::parse(r#" TOC \t "Figure Caption,1,Table Caption,2" \n "#);
        let styles = [
            ("Figure Caption".to_string(), 1),
            ("Table Caption".to_string(), 2),
        ];
        assert_eq!(spec.styles, styles);
        assert!(!spec.page_numbers && spec.outline.is_none());
        let spec = super::TocSpec::parse(r#" TOC \c "Figure" "#);
        assert_eq!(spec.seq.as_deref(), Some("Figure"));
        assert!(spec.figures);
        assert!(super::TocSpec::parse(r#" TOC \t "Figure Caption,1" \h \z \c "#).figures);
        let spec = super::TocSpec::parse(" TOC \\h ");
        assert_eq!(spec.outline, Some(1..=3));
        assert!(!spec.figures);
    }
}
