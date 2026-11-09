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

/// Text-line fragments per page from one `mutool draw -F stext` run over the whole document.
fn extract_all_page_fragments(pdf: &Path) -> Vec<Vec<Fragment>> {
    let output = Command::new("mutool")
        .args(["draw", "-F", "stext", pdf.to_str().unwrap()])
        .output()
        .expect("Failed to run mutool draw -F stext");
    assert!(
        output.status.success(),
        "mutool draw -F stext failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    parse_stext_fragments(&String::from_utf8_lossy(&output.stdout))
        .expect("mutool returned invalid structured-text XML")
}

/// One mutool `line`: its left and right edge, top and text.
#[derive(Clone)]
struct Fragment {
    x0: f64,
    x1: f64,
    top: f64,
    text: String,
}

/// MuPDF emits character nodes inside each line. Also accept the compact
/// line-level text attribute used by older metric inputs.
fn parse_stext_fragments(xml: &str) -> Result<Vec<Vec<Fragment>>, roxmltree::Error> {
    let doc = roxmltree::Document::parse(xml)?;
    Ok(doc
        .descendants()
        .filter(|n| n.has_tag_name("page"))
        .map(|page| {
            page.descendants()
                .filter(|n| n.has_tag_name("line"))
                .filter_map(|line| {
                    let bbox: Vec<f64> = line
                        .attribute("bbox")
                        .unwrap_or("")
                        .split_whitespace()
                        .filter_map(|v| v.parse().ok())
                        .collect();
                    let at = |i: usize| bbox.get(i).copied().unwrap_or(0.0);
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
                    (!text.is_empty()).then(|| Fragment {
                        x0: at(0),
                        x1: at(2),
                        top: at(1),
                        text: text.to_owned(),
                    })
                })
                .collect()
        })
        .collect())
}

/// Sort fragments by top and cluster within 8pt of each cluster's first fragment, so
/// superscripts/subscripts and side-by-side runs of one visual line stay together.
fn group_by_top(mut frags: Vec<Fragment>) -> Vec<Vec<Fragment>> {
    frags.sort_by(|a, b| a.top.total_cmp(&b.top));
    let mut clusters: Vec<Vec<Fragment>> = Vec::new();
    for frag in frags {
        if let Some(last) = clusters.last_mut()
            && frag.top - last[0].top < 8.0
        {
            last.push(frag);
            continue;
        }
        clusters.push(vec![frag]);
    }
    clusters
}

/// A line's fragments in x order (reading order for superscripts and adjacent font runs).
fn line_text(mut line: Vec<Fragment>) -> String {
    line.sort_by(|a, b| a.x0.total_cmp(&b.x0));
    line.into_iter()
        .map(|f| f.text)
        .collect::<Vec<_>>()
        .join(" ")
}

/// 1pt slack so touching runs don't overlap.
fn overlap_x(a: &Fragment, b: &Fragment) -> bool {
    a.x0.max(b.x0) + 1.0 < a.x1.min(b.x1)
}

/// Overlapping in x at different tops: two lines one above the other, never one visual line.
/// 1pt slack for text drawn twice at one top (simulated bold).
fn stack(a: &Fragment, b: &Fragment) -> bool {
    overlap_x(a, b) && (a.top - b.top).abs() > 1.0
}

/// How far (pt) a generated fragment's top may sit from a reference fragment's and still join it.
const ASSIGN_TOLERANCE: f64 = 3.0;

/// Our fragments grouped into the reference's lines, so a sub-point shift can't carry a fragment
/// across the 8pt cluster limit (text-box-table-anchor: Word's two lines 8.00pt apart, ours 7.84pt
/// apart would merge). Each fragment joins the line of the reference fragment whose top is
/// nearest, within [`ASSIGN_TOLERANCE`], preferring one it overlaps in x: Word's own 8pt grouping
/// can split one table row (label 7.95pt below the header joins it, the row's numbers 8.15pt
/// below don't), and ours halfway between must follow its own counterpart. Two of ours that
/// [`stack`] can't both be that one reference fragment, so the farther (a reflowed extra line)
/// stays unassigned. An unassigned fragment (side-by-side content drifted off Word's row: a right
/// cell, a comment pane) joins the nearest assigned line within 8pt unless it stacks with one of
/// that line's fragments; the rest group by the 8pt rule.
fn group_like_reference(reference: &[Vec<Fragment>], generated: Vec<Fragment>) -> Vec<String> {
    let targets: Vec<(&Fragment, usize)> = reference
        .iter()
        .enumerate()
        .flat_map(|(line, frags)| frags.iter().map(move |f| (f, line)))
        .collect();
    let mut candidates = Vec::new();
    let mut unassigned = Vec::new();
    for frag in generated {
        let nearest = targets
            .iter()
            .map(|&(r, _)| (!overlap_x(r, &frag), (r.top - frag.top).abs()))
            .enumerate()
            .filter(|(_, (_, d))| *d <= ASSIGN_TOLERANCE)
            .min_by(|a, b| a.1.0.cmp(&b.1.0).then(a.1.1.total_cmp(&b.1.1)));
        match nearest {
            Some((target, (_, d))) => candidates.push((d, target, frag)),
            None => unassigned.push(frag),
        }
    }
    candidates.sort_by(|a, b| a.0.total_cmp(&b.0));
    let mut slots: Vec<Vec<Fragment>> = vec![Vec::new(); targets.len()];
    for (_, target, frag) in candidates {
        if slots[target].iter().any(|kept| stack(&frag, kept)) {
            unassigned.push(frag);
        } else {
            slots[target].push(frag);
        }
    }
    let mut assigned: Vec<Vec<Fragment>> = vec![Vec::new(); reference.len()];
    for (slot, &(_, line)) in slots.into_iter().zip(&targets) {
        assigned[line].extend(slot);
    }
    let mut loose = Vec::new();
    for frag in unassigned {
        let nearest = assigned
            .iter()
            .enumerate()
            .flat_map(|(line, frags)| frags.iter().map(move |f| ((f.top - frag.top).abs(), line)))
            .min_by(|a, b| a.0.total_cmp(&b.0));
        match nearest {
            Some((d, line)) if d < 8.0 && !assigned[line].iter().any(|f| stack(&frag, f)) => {
                assigned[line].push(frag)
            }
            _ => loose.push(frag),
        }
    }
    let mut lines: Vec<(f64, Vec<Fragment>)> = assigned
        .into_iter()
        .zip(reference)
        .filter(|(frags, _)| !frags.is_empty())
        .map(|(frags, line)| (line[0].top, frags))
        .chain(group_by_top(loose).into_iter().map(|c| (c[0].top, c)))
        .collect();
    lines.sort_by(|a, b| a.0.total_cmp(&b.0));
    lines.into_iter().map(|(_, line)| line_text(line)).collect()
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

/// Lines whose first and last words match, paired by longest common subsequence rather than by
/// index: two fragments merging into one line (a table cell 7.95pt below its neighbour where
/// Word's is 8.2pt, either side of the 8pt cluster limit) would otherwise shift every later
/// line against the reference and fail them all.
fn aligned_matches(ref_lines: &[String], gen_lines: &[String]) -> usize {
    let key = |s: &String| (first_word(s), last_word(s));
    let (r, g): (Vec<_>, Vec<_>) = (
        ref_lines.iter().map(key).collect(),
        gen_lines.iter().map(key).collect(),
    );
    let mut row = vec![0; g.len() + 1];
    for rk in &r {
        let mut diag = 0;
        for (j, gk) in g.iter().enumerate() {
            let up = row[j + 1];
            row[j + 1] = if rk == gk { diag + 1 } else { up.max(row[j]) };
            diag = up;
        }
    }
    row[g.len()]
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

    let ref_line_groups: Vec<_> = extract_all_page_fragments(reference_pdf)
        .into_iter()
        .map(group_by_top)
        .collect();
    let gen_line_pages: Vec<Vec<String>> = ref_line_groups
        .iter()
        .zip(extract_all_page_fragments(generated_pdf))
        .map(|(reference, generated)| group_like_reference(reference, generated))
        .collect();
    let ref_line_pages: Vec<Vec<String>> = ref_line_groups
        .into_iter()
        .map(|page| page.into_iter().map(line_text).collect())
        .collect();
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

        total_lines += min_count;
        matching_lines += aligned_matches(ref_lines, gen_lines);
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
    use super::{Fragment, group_by_top, group_like_reference, line_text, parse_stext_fragments};

    fn parse_stext_pages(xml: &str) -> Result<Vec<Vec<String>>, roxmltree::Error> {
        Ok(parse_stext_fragments(xml)?
            .into_iter()
            .map(|page| group_by_top(page).into_iter().map(line_text).collect())
            .collect())
    }

    fn frag(x0: f64, x1: f64, top: f64, text: &str) -> Fragment {
        Fragment {
            x0,
            x1,
            top,
            text: text.to_owned(),
        }
    }

    #[test]
    fn sub_point_shift_across_8pt_keeps_reference_lines() {
        let reference = group_by_top(vec![
            frag(53.0, 78.0, 291.94, "Date:"),
            frag(327.6, 380.0, 299.94, "3pt borders"),
        ]);
        let generated = vec![
            frag(53.0, 78.0, 291.95, "Date:"),
            frag(327.6, 380.0, 299.79, "3pt borders"),
        ];
        assert_eq!(group_by_top(generated.clone()).len(), 1);
        assert_eq!(
            group_like_reference(&reference, generated),
            ["Date:", "3pt borders"]
        );
    }

    #[test]
    fn side_by_side_cell_drifted_past_the_tolerance_stays_on_its_line() {
        let reference = group_by_top(vec![
            frag(70.0, 100.0, 200.0, "(PPT)"),
            frag(150.0, 300.0, 200.0, "[yes, no]"),
        ]);
        let generated = vec![
            frag(70.0, 100.0, 200.0, "(PPT)"),
            frag(150.0, 300.0, 195.0, "[yes, no]"),
        ];
        assert_eq!(
            group_like_reference(&reference, generated),
            ["(PPT) [yes, no]"]
        );
    }

    #[test]
    fn reflowed_extra_line_is_not_glued_to_the_reference_line() {
        let reference = group_by_top(vec![
            frag(0.0, 300.0, 100.0, "a b c"),
            frag(0.0, 50.0, 130.0, "next"),
        ]);
        let wrapped = vec![
            frag(0.0, 200.0, 99.0, "a b"),
            frag(0.0, 100.0, 106.0, "c"),
            frag(0.0, 50.0, 130.0, "next"),
        ];
        assert_eq!(
            group_like_reference(&reference, wrapped),
            ["a b", "c", "next"]
        );
        // Both within the tolerance of the one reference line: the closer keeps it.
        let squeezed = vec![frag(0.0, 200.0, 102.5, "a b"), frag(0.0, 100.0, 99.5, "c")];
        assert_eq!(group_like_reference(&reference, squeezed), ["c", "a b"]);
    }

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
    fn one_merged_line_costs_only_its_own_lines() {
        let lines = |s: &[&str]| s.iter().map(|l| l.to_string()).collect::<Vec<_>>();
        let reference = lines(&["KELAYAKAN", "NO. DOKUMEN", "YA TIDAK", "1. Akta"]);
        let generated = lines(&["NO. DOKUMEN KELAYAKAN", "YA TIDAK", "1. Akta"]);
        assert_eq!(super::aligned_matches(&reference, &generated), 2);
    }

    #[test]
    fn invalid_xml_is_an_error_instead_of_a_zero_line_measurement() {
        assert!(parse_stext_pages("<document><page>").is_err());
    }
}
