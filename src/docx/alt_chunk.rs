use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::model::{
    Alignment, Block, CellBorder, CellBorders, CellMargins, Paragraph, Run, Table, TableCell,
    TableRow,
};

use super::parse_hex_color;

pub(super) fn parse_alt_chunk<R: Read + Seek>(
    rel_id: &str,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
) -> Vec<Block> {
    let Some(target) = rels.get(rel_id) else {
        return vec![];
    };
    let zip_path = target.trim_start_matches('/');
    let raw = match super::read_zip_text(zip, zip_path) {
        Some(s) => s,
        None => return vec![],
    };

    let html = if raw.starts_with("MIME-Version:") || raw.starts_with("Content-Type:") {
        match extract_html_from_mht(&raw) {
            Some(h) => h,
            None => return vec![],
        }
    } else {
        raw
    };

    let fixed = fix_xhtml_void_tags(&html);
    let Ok(doc) = roxmltree::Document::parse(&fixed) else {
        return vec![];
    };

    let css = extract_css(&doc);
    convert_html_to_blocks(&doc, &css)
}

// --- MHT / MIME envelope ---

fn extract_html_from_mht(raw: &str) -> Option<String> {
    let boundary = raw.lines().find_map(|line| {
        let line = line.trim();
        if let Some(rest) = line.strip_prefix("Content-Type:") {
            rest.split(';')
                .find_map(|part| part.trim().strip_prefix("boundary="))
                .map(|b| b.trim_matches('"').to_string())
        } else {
            None
        }
    })?;

    let delimiter = format!("--{boundary}");
    for part in raw.split(&delimiter) {
        let Some(header_end) = part.find("\r\n\r\n").or_else(|| part.find("\n\n")) else {
            continue;
        };
        let headers = &part[..header_end];
        if !headers.lines().any(|l| l.contains("text/html")) {
            continue;
        }

        let body_start = if part[header_end..].starts_with("\r\n\r\n") {
            header_end + 4
        } else {
            header_end + 2
        };
        let body = &part[body_start..];

        let is_qp = headers
            .lines()
            .any(|l| l.to_ascii_lowercase().contains("quoted-printable"));

        return Some(if is_qp {
            decode_quoted_printable(body)
        } else {
            body.to_string()
        });
    }
    None
}

fn decode_quoted_printable(input: &str) -> String {
    let mut out = Vec::with_capacity(input.len());
    let bytes = input.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'=' {
            if i + 2 < bytes.len() {
                let hi = bytes[i + 1];
                let lo = bytes[i + 2];
                if hi == b'\r' || hi == b'\n' {
                    i += 2;
                    if i < bytes.len() && bytes[i] == b'\n' {
                        i += 1;
                    }
                    continue;
                }
                if let (Some(h), Some(l)) = (hex_val(hi), hex_val(lo)) {
                    out.push(h << 4 | l);
                    i += 3;
                    continue;
                }
            }
            out.push(b'=');
            i += 1;
        } else {
            out.push(bytes[i]);
            i += 1;
        }
    }
    String::from_utf8_lossy(&out).into_owned()
}

fn hex_val(b: u8) -> Option<u8> {
    match b {
        b'0'..=b'9' => Some(b - b'0'),
        b'A'..=b'F' => Some(b - b'A' + 10),
        b'a'..=b'f' => Some(b - b'a' + 10),
        _ => None,
    }
}

// --- XHTML fixup ---

fn fix_xhtml_void_tags(html: &str) -> String {
    let mut result = String::with_capacity(html.len());
    let mut rest = html;
    while let Some(idx) = rest.find('<') {
        result.push_str(&rest[..idx]);
        rest = &rest[idx..];

        let Some(end) = rest.find('>') else {
            result.push_str(rest);
            return result;
        };

        let tag_content = &rest[1..end];
        let is_void = ["meta", "br", "hr", "img", "link", "input"]
            .iter()
            .any(|t| {
                let tc = tag_content.trim_start();
                tc.starts_with(t)
                    && tc.as_bytes()[t.len()..]
                        .first()
                        .is_none_or(|&b| b == b' ' || b == b'/' || b == b'>')
            });

        if is_void && !tag_content.ends_with('/') && !tag_content.starts_with('/') {
            result.push('<');
            result.push_str(tag_content);
            result.push_str("/>");
        } else {
            result.push_str(&rest[..=end]);
        }
        rest = &rest[end + 1..];
    }
    result.push_str(rest);
    result
}

// --- CSS parsing ---

#[derive(Default, Clone)]
struct CssProperties {
    font_size_pt: Option<f32>,
    font_family: Option<String>,
    bold: Option<bool>,
    text_align: Option<String>,
    text_indent_pt: Option<f32>,
    margin_top_pt: Option<f32>,
    margin_bottom_pt: Option<f32>,
    margin_left_pt: Option<f32>,
    line_height_pct: Option<f32>,
    color: Option<[u8; 3]>,
    width_px: Option<f32>,
    vertical_align: Option<String>,
    border_top: Option<CellBorder>,
    border_right: Option<CellBorder>,
    border_bottom: Option<CellBorder>,
    border_left: Option<CellBorder>,
}

fn parse_css_numeric(val: &str) -> f32 {
    let numeric: String = val
        .trim()
        .replace(',', ".")
        .chars()
        .take_while(|c| c.is_ascii_digit() || *c == '.' || *c == '-')
        .collect();
    numeric.parse().unwrap_or(0.0)
}

fn parse_css_length_pt(val: &str) -> f32 {
    let val = val.trim();
    let n = parse_css_numeric(val);
    if val.ends_with("pt") {
        n
    } else if val.ends_with("px") {
        n * 0.75
    } else if val.ends_with("in") {
        n * 72.0
    } else {
        n
    }
}

fn parse_font_weight_bold(val: &str) -> bool {
    val == "bold" || val.parse::<u32>().is_ok_and(|n| n >= 700)
}

fn parse_css_border(val: &str) -> Option<CellBorder> {
    let parts: Vec<&str> = val.split_whitespace().collect();
    if parts.is_empty() || parts[0] == "none" {
        return None;
    }
    let mut width = 0.5f32;
    for p in &parts {
        if p.ends_with("px") || p.ends_with("pt") {
            width = parse_css_length_pt(p);
        }
    }
    Some(CellBorder::visible(Some([0, 0, 0]), width))
}

fn parse_css_properties(decl_block: &str) -> CssProperties {
    let mut props = CssProperties::default();
    for decl in decl_block.split(';') {
        let decl = decl.trim();
        let Some((key, val)) = decl.split_once(':') else {
            continue;
        };
        let key = key.trim().to_ascii_lowercase();
        let val = val.trim();
        match key.as_str() {
            "font-size" => props.font_size_pt = Some(parse_css_length_pt(val)),
            "font-family" => {
                let first = val.split(',').next().unwrap_or(val);
                props.font_family =
                    Some(first.trim().trim_matches('\'').trim_matches('"').to_string());
            }
            "font-weight" => props.bold = Some(parse_font_weight_bold(val)),
            "text-align" => props.text_align = Some(val.to_string()),
            "text-indent" => props.text_indent_pt = Some(parse_css_length_pt(val)),
            "margin-top" => props.margin_top_pt = Some(parse_css_length_pt(val)),
            "margin-bottom" => props.margin_bottom_pt = Some(parse_css_length_pt(val)),
            "margin-left" => props.margin_left_pt = Some(parse_css_length_pt(val)),
            "line-height" => {
                let n = val.replace(',', ".");
                if let Some(pct) = n.strip_suffix('%') {
                    props.line_height_pct = pct.trim().parse().ok();
                }
            }
            "color" => {
                let c = val.trim_start_matches('#');
                props.color = parse_hex_color(c);
            }
            "width" => props.width_px = Some(parse_css_numeric(val)),
            "vertical-align" => props.vertical_align = Some(val.to_string()),
            "border-top" => props.border_top = parse_css_border(val),
            "border-right" => props.border_right = parse_css_border(val),
            "border-bottom" => props.border_bottom = parse_css_border(val),
            "border-left" => props.border_left = parse_css_border(val),
            _ => {}
        }
    }
    props
}

fn extract_css(doc: &roxmltree::Document) -> HashMap<String, CssProperties> {
    let mut map = HashMap::new();
    for node in doc.descendants() {
        if node.tag_name().name() == "style" {
            if let Some(text) = node.text() {
                parse_css_block(text, &mut map);
            }
            for child in node.children() {
                if child.is_text() && let Some(t) = child.text() {
                    parse_css_block(t, &mut map);
                }
            }
        }
    }
    map
}

fn parse_css_block(text: &str, map: &mut HashMap<String, CssProperties>) {
    let mut rest = text;
    while let Some(brace) = rest.find('{') {
        let selector = rest[..brace].trim();
        let Some(end_brace) = rest[brace..].find('}') else {
            break;
        };
        let body = &rest[brace + 1..brace + end_brace];
        let props = parse_css_properties(body);
        map.insert(selector.to_string(), props);
        rest = &rest[brace + end_brace + 1..];
    }
}

// --- HTML → IR conversion ---

fn convert_html_to_blocks(
    doc: &roxmltree::Document,
    css: &HashMap<String, CssProperties>,
) -> Vec<Block> {
    let body = match find_element(doc.root(), "body") {
        Some(b) => b,
        None => doc.root_element(),
    };

    let mut blocks = Vec::new();
    convert_children_to_blocks(body, css, &mut blocks);
    blocks
}

fn convert_children_to_blocks(
    parent: roxmltree::Node,
    css: &HashMap<String, CssProperties>,
    blocks: &mut Vec<Block>,
) {
    for child in parent.children() {
        if !child.is_element() {
            continue;
        }
        match child.tag_name().name() {
            "p" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6" => {
                blocks.push(Block::Paragraph(convert_paragraph(child, css)));
            }
            "table" => {
                if let Some(tbl) = convert_table(child, css) {
                    blocks.push(Block::Table(tbl));
                }
            }
            "div" | "section" | "article" | "main" => {
                convert_children_to_blocks(child, css, blocks);
            }
            _ => {}
        }
    }
}

fn resolve_css(
    node: roxmltree::Node,
    css: &HashMap<String, CssProperties>,
) -> CssProperties {
    let tag = node.tag_name().name();
    let class = node.attribute("class").unwrap_or("");

    let class_props = if !class.is_empty() {
        css.get(&format!("{tag}.{class}"))
            .or_else(|| css.get(&format!(".{class}")))
    } else {
        css.get(tag)
    };

    let mut merged = class_props.cloned().unwrap_or_default();

    if let Some(style) = node.attribute("style") {
        let inline = parse_css_properties(style);
        merge_css(&mut merged, &inline);
    }

    merged
}

fn merge_css(base: &mut CssProperties, over: &CssProperties) {
    if over.font_size_pt.is_some() {
        base.font_size_pt = over.font_size_pt;
    }
    if over.font_family.is_some() {
        base.font_family.clone_from(&over.font_family);
    }
    if over.bold.is_some() {
        base.bold = over.bold;
    }
    if over.text_align.is_some() {
        base.text_align.clone_from(&over.text_align);
    }
    if over.text_indent_pt.is_some() {
        base.text_indent_pt = over.text_indent_pt;
    }
    if over.margin_top_pt.is_some() {
        base.margin_top_pt = over.margin_top_pt;
    }
    if over.margin_bottom_pt.is_some() {
        base.margin_bottom_pt = over.margin_bottom_pt;
    }
    if over.margin_left_pt.is_some() {
        base.margin_left_pt = over.margin_left_pt;
    }
    if over.line_height_pct.is_some() {
        base.line_height_pct = over.line_height_pct;
    }
    if over.color.is_some() {
        base.color = over.color;
    }
    if over.width_px.is_some() {
        base.width_px = over.width_px;
    }
    if over.vertical_align.is_some() {
        base.vertical_align.clone_from(&over.vertical_align);
    }
    if over.border_top.is_some() {
        base.border_top = over.border_top;
    }
    if over.border_right.is_some() {
        base.border_right = over.border_right;
    }
    if over.border_bottom.is_some() {
        base.border_bottom = over.border_bottom;
    }
    if over.border_left.is_some() {
        base.border_left = over.border_left;
    }
}

struct RunContext<'a> {
    css: &'a HashMap<String, CssProperties>,
    font_size: f32,
    font_name: &'a str,
    bold: bool,
    italic: bool,
    underline: bool,
    color: Option<[u8; 3]>,
}

fn convert_paragraph(
    node: roxmltree::Node,
    css: &HashMap<String, CssProperties>,
) -> Paragraph {
    let props = resolve_css(node, css);

    let alignment = match props.text_align.as_deref() {
        Some("center") => Alignment::Center,
        Some("right") => Alignment::Right,
        Some("justify") => Alignment::Justify,
        _ => Alignment::Left,
    };

    let font_size = props.font_size_pt.unwrap_or(12.0);
    let font_name = props
        .font_family
        .clone()
        .unwrap_or_else(|| "Times New Roman".to_string());
    let bold = props.bold.unwrap_or(false);

    let mut runs = Vec::new();
    let ctx = RunContext {
        css,
        font_size,
        font_name: &font_name,
        bold,
        italic: false,
        underline: false,
        color: props.color,
    };
    collect_runs(node, &ctx, &mut runs);
    trim_block_whitespace(&mut runs);

    let max_run_fs = runs.iter().map(|r| r.font_size).fold(0.0f32, f32::max);

    let line_spacing = if max_run_fs > 0.0 && font_size > max_run_fs + 0.1 {
        // Paragraph CSS font-size exceeds tallest run (e.g. 30pt paragraph
        // with 16pt spans) — use it as minimum line height.
        Some(crate::model::LineSpacing::AtLeast(font_size))
    } else if let Some(pct) = props.line_height_pct {
        Some(crate::model::LineSpacing::Auto(pct / 100.0))
    } else {
        // Default for HTML content without explicit line-height
        Some(crate::model::LineSpacing::Auto(1.1))
    };

    Paragraph {
        runs,
        space_before: props.margin_top_pt.unwrap_or(0.0),
        space_after: props.margin_bottom_pt.unwrap_or(0.0),
        alignment,
        indent_left: props.margin_left_pt.unwrap_or(0.0),
        indent_first_line: props.text_indent_pt.unwrap_or(0.0),
        line_spacing,
        ..Paragraph::default()
    }
}

fn collect_runs(node: roxmltree::Node, ctx: &RunContext, runs: &mut Vec<Run>) {
    for child in node.children() {
        if child.is_text() {
            let text = collapse_whitespace(child.text().unwrap_or(""));
            if !text.is_empty() {
                runs.push(Run {
                    text,
                    font_size: ctx.font_size,
                    font_name: ctx.font_name.to_string(),
                    bold: ctx.bold,
                    italic: ctx.italic,
                    underline: ctx.underline,
                    color: ctx.color,
                    ..Run::default()
                });
            }
            continue;
        }
        if !child.is_element() {
            continue;
        }

        let tag = child.tag_name().name();
        match tag {
            "span" | "b" | "strong" | "i" | "em" | "a" | "u" => {
                let span_css = resolve_css(child, ctx.css);
                let child_ctx = RunContext {
                    css: ctx.css,
                    font_size: span_css.font_size_pt.unwrap_or(ctx.font_size),
                    font_name: span_css.font_family.as_deref().unwrap_or(ctx.font_name),
                    bold: span_css.bold.unwrap_or(ctx.bold)
                        || tag == "b"
                        || tag == "strong",
                    italic: ctx.italic || tag == "i" || tag == "em",
                    underline: ctx.underline || tag == "u",
                    color: span_css.color.or(ctx.color),
                };
                collect_runs(child, &child_ctx, runs);
            }
            "br" => {
                runs.push(Run {
                    text: "\n".to_string(),
                    font_size: ctx.font_size,
                    font_name: ctx.font_name.to_string(),
                    ..Run::default()
                });
            }
            _ => {
                collect_runs(child, ctx, runs);
            }
        }
    }
}

fn collect_text(node: roxmltree::Node) -> String {
    let mut out = String::new();
    for child in node.children() {
        if child.is_text() {
            out.push_str(child.text().unwrap_or(""));
        } else if child.is_element() {
            out.push_str(&collect_text(child));
        }
    }
    out
}

/// Strip leading/trailing whitespace-only runs from a block element's run list.
/// Matches HTML rendering: whitespace at the start/end of block elements is ignored.
fn trim_block_whitespace(runs: &mut Vec<Run>) {
    while runs.first().is_some_and(|r| r.text.trim().is_empty()) {
        runs.remove(0);
    }
    while runs.last().is_some_and(|r| r.text.trim().is_empty()) {
        runs.pop();
    }
}

fn collapse_whitespace(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut prev_ws = false;
    for ch in s.chars() {
        if ch.is_ascii_whitespace() {
            if !prev_ws {
                out.push(' ');
            }
            prev_ws = true;
        } else {
            out.push(ch);
            prev_ws = false;
        }
    }
    out
}

fn find_element<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    for child in node.children() {
        if child.is_element() && child.tag_name().name() == name {
            return Some(child);
        }
        if let Some(found) = find_element(child, name) {
            return Some(found);
        }
    }
    None
}

// --- Table conversion ---

fn is_table_cell(n: &roxmltree::Node) -> bool {
    n.is_element() && (n.tag_name().name() == "td" || n.tag_name().name() == "th")
}

fn cell_colspan(td: &roxmltree::Node) -> usize {
    td.attribute("colspan")
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(1)
}

fn convert_table(
    table_node: roxmltree::Node,
    css: &HashMap<String, CssProperties>,
) -> Option<Table> {
    let tbody = find_element(table_node, "tbody").unwrap_or(table_node);

    let tr_nodes: Vec<_> = tbody
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tr")
        .collect();

    if tr_nodes.is_empty() {
        return None;
    }

    // Collect td/th nodes per row once, reused across passes.
    let rows_tds: Vec<Vec<_>> = tr_nodes
        .iter()
        .map(|tr| tr.children().filter(|n| is_table_cell(n)).collect())
        .collect();

    // First pass: determine max column count and extract column widths from
    // the row with the most individual (non-colspan) cells.
    let mut max_cols = 0usize;
    let mut col_widths: Vec<f32> = Vec::new();
    let mut best_individual_cells = 0usize;

    for tds in &rows_tds {
        let total_cols: usize = tds.iter().map(|td| cell_colspan(td)).sum();
        if total_cols > max_cols {
            max_cols = total_cols;
        }

        let individual = tds.iter().filter(|td| cell_colspan(td) == 1).count();
        if individual > best_individual_cells {
            best_individual_cells = individual;
            col_widths = vec![72.0; max_cols.max(total_cols)];
            let mut col_idx = 0;
            for td in tds {
                let td_css = resolve_css(*td, css);
                let colspan = cell_colspan(td);
                let w_pt = td_css.width_px.unwrap_or(100.0) * 0.75;
                if colspan == 1 {
                    if col_idx < col_widths.len() {
                        col_widths[col_idx] = w_pt;
                    }
                } else {
                    let per_col = w_pt / colspan as f32;
                    for i in 0..colspan {
                        if col_idx + i < col_widths.len() {
                            col_widths[col_idx + i] = per_col;
                        }
                    }
                }
                col_idx += colspan;
            }
        }
    }

    // Ensure col_widths covers all columns
    col_widths.resize(max_cols, 72.0);

    // Build rows
    let mut rows = Vec::new();
    for tds in &rows_tds {
        let mut cells = Vec::new();
        for td in tds {
            let td_css = resolve_css(*td, css);
            let colspan = cell_colspan(td);
            let w_pt = td_css.width_px.unwrap_or(100.0) * 0.75;

            let borders = CellBorders {
                top: td_css.border_top.unwrap_or_default(),
                right: td_css.border_right.unwrap_or_default(),
                bottom: td_css.border_bottom.unwrap_or_default(),
                left: td_css.border_left.unwrap_or_default(),
            };

            let v_align = match td_css.vertical_align.as_deref() {
                Some("middle") => crate::model::CellVAlign::Center,
                Some("bottom") => crate::model::CellVAlign::Bottom,
                _ => crate::model::CellVAlign::Top,
            };

            let mut cell_paras = Vec::new();
            for child in td.children() {
                if child.is_element()
                    && matches!(
                        child.tag_name().name(),
                        "p" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6"
                    )
                {
                    cell_paras.push(convert_paragraph(child, css));
                }
            }
            if cell_paras.is_empty() {
                let text = collect_text(*td);
                let text = collapse_whitespace(&text);
                cell_paras.push(Paragraph {
                    runs: vec![Run {
                        text,
                        font_size: td_css.font_size_pt.unwrap_or(12.0),
                        font_name: td_css
                            .font_family
                            .clone()
                            .unwrap_or_else(|| "Times New Roman".to_string()),
                        ..Run::default()
                    }],
                    ..Paragraph::default()
                });
            }

            cells.push(TableCell {
                width: w_pt,
                paragraphs: cell_paras,
                borders,
                shading: None,
                grid_span: colspan as u16,
                v_merge: crate::model::VMerge::None,
                v_align,
            });
        }

        rows.push(TableRow {
            cells,
            height: None,
            height_exact: false,
        });
    }

    Some(Table {
        col_widths,
        rows,
        table_indent: 0.0,
        cell_margins: CellMargins::default(),
        position: None,
    })
}
