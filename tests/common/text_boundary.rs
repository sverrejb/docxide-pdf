//! Text-boundary metric: do lines start and end on the same words as the reference?
//! Shared by `tests/text_boundary.rs` and `tools/src/bin/page_metrics.rs` so the
//! engine-comparison viewer reports exactly the harness number.
use std::path::Path;
use std::process::Command;

/// Words per page from one `mutool draw -F text` run over the whole document. mutool ends every
/// page with a form feed, and the whole-document output is byte-identical to the per-page outputs
/// concatenated, so this is the same tokenisation as one run per page at a fraction of the cost:
/// mutool parses the PDF once instead of once per page (on a 200-page document the per-page
/// spawns were two thirds of the metric's time).
fn extract_all_pages(pdf: &Path) -> Vec<Vec<String>> {
    let output = Command::new("mutool")
        .args(["draw", "-F", "text", pdf.to_str().unwrap()])
        .output()
        .expect("Failed to run mutool draw");
    assert!(
        output.status.success(),
        "mutool draw -F text failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let text = String::from_utf8_lossy(&output.stdout);
    let mut pages: Vec<Vec<String>> = text
        .split('\x0c')
        .map(|page| page.split_whitespace().map(String::from).collect())
        .collect();
    if pages.last().is_some_and(|p| p.is_empty()) {
        pages.pop(); // after the last page's form feed
    }
    pages
}

/// Text lines per page from one `mutool draw -F stext` run over the whole document, each page
/// parsed by [`stext_lines`].
fn extract_all_page_lines(pdf: &Path) -> Vec<Vec<String>> {
    let output = Command::new("mutool")
        .args(["draw", "-F", "stext", pdf.to_str().unwrap()])
        .output()
        .expect("Failed to run mutool draw -F stext");
    assert!(
        output.status.success(),
        "mutool draw -F stext failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    parse_stext_pages(&String::from_utf8_lossy(&output.stdout))
        .expect("mutool returned invalid structured-text XML")
}

/// MuPDF emits character nodes inside each line. Also accept the compact
/// line-level text attribute used by older metric inputs.
fn parse_stext_pages(xml: &str) -> Result<Vec<Vec<String>>, roxmltree::Error> {
    let doc = roxmltree::Document::parse(xml)?;
    Ok(doc
        .descendants()
        .filter(|n| n.has_tag_name("page"))
        .map(|page| {
            let lines = page
                .descendants()
                .filter(|n| n.has_tag_name("line"))
                .filter_map(|line| {
                    let mut bbox = line.attribute("bbox").unwrap_or("").split_whitespace();
                    let x = bbox.next().and_then(|v| v.parse().ok()).unwrap_or(0.0);
                    let y = bbox.next().and_then(|v| v.parse().ok()).unwrap_or(0.0);
                    let text = line
                        .attribute("text")
                        .map(str::to_owned)
                        .unwrap_or_else(|| {
                            line.descendants()
                                .filter(|n| n.has_tag_name("char"))
                                .filter_map(|n| n.attribute("c"))
                                .collect()
                        });
                    let text = text.trim();
                    (!text.is_empty()).then(|| (x, y, text.to_owned()))
                })
                .collect();
            merge_stext_lines(lines)
        })
        .collect())
}

/// Sort line fragments by y, cluster within 8pt, then order each cluster by x
/// so superscripts/subscripts and adjacent font runs retain reading order.
fn merge_stext_lines(mut lines: Vec<(f64, f64, String)>) -> Vec<String> {
    lines.sort_by(|a, b| a.1.partial_cmp(&b.1).unwrap_or(std::cmp::Ordering::Equal));
    let mut clusters: Vec<Vec<(f64, f64, String)>> = Vec::new();
    for item in lines {
        if let Some(last) = clusters.last_mut()
            && (item.1 - last[0].1).abs() < 8.0
        {
            last.push(item);
            continue;
        }
        clusters.push(vec![item]);
    }
    clusters
        .into_iter()
        .map(|mut cluster| {
            cluster.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));
            cluster
                .into_iter()
                .map(|(_, _, t)| t)
                .collect::<Vec<_>>()
                .join(" ")
        })
        .collect()
}

fn break_positions(pages: &[Vec<String>]) -> Vec<usize> {
    let mut pos = Vec::with_capacity(pages.len());
    let mut cumulative = 0;
    for page in pages {
        cumulative += page.len();
        pos.push(cumulative);
    }
    pos
}

/// Replace tab-leader dots (runs of 3+) with a space so comparison focuses on text content.
fn normalize_leaders(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut dots = 0usize;
    for c in s.chars() {
        if c == '.' {
            dots += 1;
        } else {
            if dots >= 3 {
                out.push(' ');
            } else {
                for _ in 0..dots {
                    out.push('.');
                }
            }
            dots = 0;
            out.push(c);
        }
    }
    if dots > 0 && dots < 3 {
        for _ in 0..dots {
            out.push('.');
        }
    }
    out
}

fn first_word(s: &str) -> String {
    let n = normalize_leaders(s);
    n.split_whitespace().next().unwrap_or_default().to_string()
}

fn last_word(s: &str) -> String {
    let n = normalize_leaders(s);
    n.split_whitespace().last().unwrap_or_default().to_string()
}

#[derive(serde::Serialize, serde::Deserialize)]
pub struct TextBoundary {
    pub ref_pages: usize,
    pub gen_pages: usize,
    pub max_break_drift: i64,
    pub total_words: usize,
    pub total_lines: usize,
    pub matching_lines: usize,
}

impl TextBoundary {
    /// None when no page was comparable line by line (n/a, not a 0% match).
    pub fn line_match_pct(&self) -> Option<f64> {
        (self.total_lines > 0).then(|| self.matching_lines as f64 / self.total_lines as f64)
    }
}

/// Compare `generated_pdf` against `reference_pdf` page by page.
pub fn analyze(reference_pdf: &Path, generated_pdf: &Path) -> TextBoundary {
    let ref_word_pages = extract_all_pages(reference_pdf);
    let gen_word_pages = extract_all_pages(generated_pdf);
    let common_pages = ref_word_pages.len().min(gen_word_pages.len());

    let ref_breaks = break_positions(&ref_word_pages);
    let gen_breaks = break_positions(&gen_word_pages);
    let total_words = ref_breaks.last().copied().unwrap_or(0);
    let break_count = (ref_breaks.len().saturating_sub(1)).min(gen_breaks.len().saturating_sub(1));
    let max_break_drift = (0..break_count)
        .map(|i| gen_breaks[i] as i64 - ref_breaks[i] as i64)
        .max_by_key(|d| d.unsigned_abs())
        .unwrap_or(0);

    let ref_line_pages = extract_all_page_lines(reference_pdf);
    let gen_line_pages = extract_all_page_lines(generated_pdf);
    let mut total_lines = 0;
    let mut matching_lines = 0;
    for p in 0..common_pages {
        let (Some(ref_lines), Some(gen_lines)) = (ref_line_pages.get(p), gen_line_pages.get(p))
        else {
            break;
        };

        let max_count = ref_lines.len().max(gen_lines.len());
        let min_count = ref_lines.len().min(gen_lines.len());
        if max_count > 0 && (max_count - min_count) as f64 / max_count as f64 > 0.15 {
            continue;
        }

        for l in 0..min_count {
            total_lines += 1;
            if first_word(&ref_lines[l]) == first_word(&gen_lines[l])
                && last_word(&ref_lines[l]) == last_word(&gen_lines[l])
            {
                matching_lines += 1;
            }
        }
    }

    TextBoundary {
        ref_pages: ref_word_pages.len(),
        gen_pages: gen_word_pages.len(),
        max_break_drift,
        total_words,
        total_lines,
        matching_lines,
    }
}

#[cfg(test)]
mod tests {
    use super::parse_stext_pages;

    #[test]
    fn character_nodes_decode_entities_and_font_fragments() {
        let xml = r#"<document><page id="p1"><line dir="1 0" bbox="10 20 80 40">
            <font><char c="A"/><char c="&amp;"/></font><font><char c="Б"/><char c="&#x20;"/><char c="4"/></font>
            </line></page></document>"#;
        assert_eq!(parse_stext_pages(xml).unwrap(), vec![vec!["A&Б 4"]]);
    }

    #[test]
    fn compact_line_attributes_keep_existing_support() {
        let xml = r#"<document><page id="p1"><line text=" A &amp; B " bbox="10 20 80 40"/></page></document>"#;
        assert_eq!(parse_stext_pages(xml).unwrap(), vec![vec!["A & B"]]);
    }

    #[test]
    fn page_order_and_empty_pages_are_preserved() {
        let xml = r#"<document><page id="p1"><line text="first"/></page><page id="p2"/><page id="p3"><line text="last"/></page></document>"#;
        assert_eq!(
            parse_stext_pages(xml).unwrap(),
            vec![vec!["first".to_owned()], vec![], vec!["last".to_owned()]]
        );
    }

    #[test]
    fn superscript_fragments_are_sorted_left_to_right() {
        let xml = r#"<document><page id="p1"><line bbox="30 6 40 16" text="2"/><line bbox="10 10 20 20" text="x"/><line bbox="10 30 40 40" text="next"/></page></document>"#;
        assert_eq!(parse_stext_pages(xml).unwrap(), vec![vec!["x 2", "next"]]);
    }

    #[test]
    fn invalid_xml_is_an_error_instead_of_a_zero_line_measurement() {
        assert!(parse_stext_pages("<document><page>").is_err());
    }
}
