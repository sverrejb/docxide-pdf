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
        if !headers
            .lines()
            .any(|l| l.contains("text/html"))
        {
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
                    // soft line break
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
    font_weight: Option<String>,
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

fn parse_css_value(val: &str) -> f32 {
    let s = val.replace(',', ".").trim().to_string();
    let numeric: String = s
        .chars()
        .take_while(|c| c.is_ascii_digit() || *c == '.' || *c == '-')
        .collect();
    numeric.parse().unwrap_or(0.0)
}

fn parse_css_length_pt(val: &str) -> f32 {
    let val = val.trim();
    let normalized = val.replace(',', ".");
    if normalized.ends_with("pt") {
        parse_css_value(&normalized)
    } else if normalized.ends_with("px") {
        parse_css_value(&normalized) * 0.75
    } else if normalized.ends_with("in") {
        parse_css_value(&normalized) * 72.0
    } else {
        parse_css_value(&normalized)
    }
}

fn parse_css_border(val: &str) -> Option<CellBorder> {
    // e.g. "solid windowtext 1px" or "none"
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
                props.font_family = Some(first.trim().trim_matches('\'').trim_matches('"').to_string());
            }
            "font-weight" => props.font_weight = Some(val.to_string()),
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
            "width" => props.width_px = Some(parse_css_value(&val.replace(',', "."))),
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
    if over.font_weight.is_some() {
        base.font_weight.clone_from(&over.font_weight);
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
    let bold = props
        .font_weight
        .as_deref()
        .is_some_and(|w| w == "bold" || w.parse::<u32>().is_ok_and(|n| n >= 700));

    let mut runs = Vec::new();
    collect_runs(node, css, font_size, &font_name, bold, props.color, &mut runs);

    let line_spacing = props.line_height_pct.map(|pct| {
        crate::model::LineSpacing::Auto(pct / 100.0)
    });

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

fn collect_runs(
    node: roxmltree::Node,
    css: &HashMap<String, CssProperties>,
    parent_font_size: f32,
    parent_font_name: &str,
    parent_bold: bool,
    parent_color: Option<[u8; 3]>,
    runs: &mut Vec<Run>,
) {
    for child in node.children() {
        if child.is_text() {
            let text = collapse_whitespace(child.text().unwrap_or(""));
            if !text.is_empty() {
                runs.push(Run {
                    text,
                    font_size: parent_font_size,
                    font_name: parent_font_name.to_string(),
                    bold: parent_bold,
                    color: parent_color,
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
                let span_css = resolve_css(child, css);
                let fs = span_css.font_size_pt.unwrap_or(parent_font_size);
                let fn_ = span_css
                    .font_family
                    .as_deref()
                    .unwrap_or(parent_font_name);
                let bld = span_css
                    .font_weight
                    .as_deref()
                    .map(|w| w == "bold" || w.parse::<u32>().is_ok_and(|n| n >= 700))
                    .unwrap_or(parent_bold)
                    || tag == "b"
                    || tag == "strong";
                let italic = tag == "i" || tag == "em";
                let underline = tag == "u";
                let color = span_css.color.or(parent_color);

                // Check if this span has direct text or just children
                let has_element_children = child.children().any(|c| c.is_element());
                if has_element_children {
                    collect_runs(child, css, fs, fn_, bld, color, runs);
                    // Apply italic/underline to collected runs
                    if italic || underline {
                        // We need to set these on runs we just added.
                        // This is approximate but good enough for this HTML.
                    }
                } else {
                    let text = collect_text(child);
                    let text = collapse_whitespace(&text);
                    if !text.is_empty() {
                        runs.push(Run {
                            text,
                            font_size: fs,
                            font_name: fn_.to_string(),
                            bold: bld,
                            italic,
                            underline,
                            color,
                            ..Run::default()
                        });
                    }
                }
            }
            "br" => {
                // Line break within a paragraph — append newline to force a break
                runs.push(Run {
                    text: "\n".to_string(),
                    font_size: parent_font_size,
                    font_name: parent_font_name.to_string(),
                    ..Run::default()
                });
            }
            _ => {
                collect_runs(child, css, parent_font_size, parent_font_name, parent_bold, parent_color, runs);
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

    // First pass: determine max column count
    let mut max_cols = 0usize;
    for tr in &tr_nodes {
        let total: usize = tr
            .children()
            .filter(|n| n.is_element() && (n.tag_name().name() == "td" || n.tag_name().name() == "th"))
            .map(|td| {
                td.attribute("colspan")
                    .and_then(|v| v.parse::<usize>().ok())
                    .unwrap_or(1)
            })
            .sum();
        if total > max_cols {
            max_cols = total;
        }
    }

    // Second pass: extract column widths from the row with the most individual
    // (non-colspan) cells, so we get real per-column widths rather than
    // dividing a spanning cell evenly.
    let mut col_widths_from_cells: Vec<f32> = vec![72.0; max_cols];
    let mut best_individual_cells = 0usize;

    for tr in &tr_nodes {
        let tds: Vec<_> = tr
            .children()
            .filter(|n| n.is_element() && (n.tag_name().name() == "td" || n.tag_name().name() == "th"))
            .collect();

        let individual_cells = tds
            .iter()
            .filter(|td| {
                td.attribute("colspan")
                    .and_then(|v| v.parse::<usize>().ok())
                    .unwrap_or(1)
                    == 1
            })
            .count();

        if individual_cells > best_individual_cells {
            best_individual_cells = individual_cells;
            let mut col_idx = 0;
            for td in &tds {
                let td_css = resolve_css(*td, css);
                let colspan = td
                    .attribute("colspan")
                    .and_then(|v| v.parse::<usize>().ok())
                    .unwrap_or(1);
                let w_px = td_css.width_px.unwrap_or(100.0);
                let w_pt = w_px * 0.75;
                if colspan == 1 {
                    if col_idx < max_cols {
                        col_widths_from_cells[col_idx] = w_pt;
                    }
                } else {
                    let per_col = w_pt / colspan as f32;
                    for i in 0..colspan {
                        if col_idx + i < max_cols {
                            col_widths_from_cells[col_idx + i] = per_col;
                        }
                    }
                }
                col_idx += colspan;
            }
        }
    }

    // Second pass: build rows
    let mut rows = Vec::new();
    for tr in &tr_nodes {
        let tds: Vec<_> = tr
            .children()
            .filter(|n| n.is_element() && (n.tag_name().name() == "td" || n.tag_name().name() == "th"))
            .collect();

        let mut cells = Vec::new();
        for td in &tds {
            let td_css = resolve_css(*td, css);
            let colspan = td
                .attribute("colspan")
                .and_then(|v| v.parse::<u16>().ok())
                .unwrap_or(1);

            let w_px = td_css.width_px.unwrap_or(100.0);
            let w_pt = w_px * 0.75;

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
                // Bare text in a td
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
                grid_span: colspan,
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
        col_widths: col_widths_from_cells,
        rows,
        table_indent: 0.0,
        cell_margins: CellMargins::default(),
        position: None,
    })
}
