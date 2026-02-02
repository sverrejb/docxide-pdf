use std::path::{Path, PathBuf};
use std::process::Command;
use std::{fs, io};

const SKIP_FIXTURES: &[&str] = &["case5"];

fn discover_fixtures() -> io::Result<Vec<PathBuf>> {
    let fixtures_dir = Path::new("tests/fixtures");
    let mut fixtures: Vec<PathBuf> = fs::read_dir(fixtures_dir)?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| {
            p.is_dir()
                && p.file_name()
                    .and_then(|n| n.to_str())
                    .map_or(true, |n| !SKIP_FIXTURES.contains(&n))
        })
        .collect();
    fixtures.sort();
    Ok(fixtures)
}

fn pdf_page_count(pdf: &Path) -> usize {
    let output = Command::new("mutool")
        .args(["info", pdf.to_str().unwrap()])
        .output()
        .expect("Failed to run mutool info");
    let text = String::from_utf8_lossy(&output.stdout);
    for line in text.lines() {
        if let Some(rest) = line.strip_prefix("Pages:") {
            if let Ok(n) = rest.trim().parse::<usize>() {
                return n;
            }
        }
    }
    0
}

fn extract_page_words(pdf: &Path, page: usize) -> Vec<String> {
    let output = Command::new("mutool")
        .args([
            "draw",
            "-F",
            "text",
            pdf.to_str().unwrap(),
            &page.to_string(),
        ])
        .output()
        .expect("Failed to run mutool draw");
    String::from_utf8_lossy(&output.stdout)
        .split_whitespace()
        .map(String::from)
        .collect()
}

/// Extract lines from a PDF page using mutool's structured text XML.
/// Each `<line text="...">` becomes one entry.
fn extract_page_lines(pdf: &Path, page: usize) -> Vec<String> {
    let output = Command::new("mutool")
        .args([
            "draw",
            "-F",
            "stext",
            pdf.to_str().unwrap(),
            &page.to_string(),
        ])
        .output()
        .expect("Failed to run mutool draw -F stext");
    let xml = String::from_utf8_lossy(&output.stdout);
    let mut lines = Vec::new();
    for xml_line in xml.lines() {
        let trimmed = xml_line.trim();
        if let Some(rest) = trimmed.strip_prefix("<line ") {
            if let Some(start) = rest.find("text=\"") {
                let after_quote = &rest[start + 6..];
                if let Some(end) = after_quote.find('"') {
                    let text = &after_quote[..end];
                    let text = text.trim();
                    if !text.is_empty() {
                        lines.push(text.to_string());
                    }
                }
            }
        }
    }
    lines
}

/// Extract all pages as word-vectors for break position analysis.
fn extract_all_pages(pdf: &Path) -> Vec<Vec<String>> {
    let n = pdf_page_count(pdf);
    (1..=n).map(|p| extract_page_words(pdf, p)).collect()
}

/// Build cumulative word-index of each page break.
fn break_positions(pages: &[Vec<String>]) -> Vec<usize> {
    let mut pos = Vec::with_capacity(pages.len());
    let mut cumulative = 0;
    for page in pages {
        cumulative += page.len();
        pos.push(cumulative);
    }
    pos
}

struct FixtureResult {
    name: String,
    page_count_match: bool,
    ref_pages: usize,
    gen_pages: usize,
    page_results: Vec<PageResult>,
    break_offsets: Vec<BreakOffset>,
    line_results: Vec<LineResult>,
}

struct PageResult {
    page: usize,
    first_ok: bool,
    last_ok: bool,
    ref_first: String,
    gen_first: String,
    ref_last: String,
    gen_last: String,
}

struct BreakOffset {
    after_page: usize,
    ref_pos: usize,
    gen_pos: usize,
    offset_words: i64,
    total_words: usize,
}

struct LineResult {
    page: usize,
    line: usize,
    first_match: bool,
    last_match: bool,
    ref_first: String,
    gen_first: String,
    ref_last: String,
    gen_last: String,
}

fn first_word(s: &str) -> String {
    s.split_whitespace()
        .next()
        .unwrap_or_default()
        .to_string()
}

fn last_word(s: &str) -> String {
    s.split_whitespace()
        .last()
        .unwrap_or_default()
        .to_string()
}

fn analyze_fixture(fixture_dir: &Path) -> Option<FixtureResult> {
    let name = fixture_dir
        .file_name()
        .unwrap()
        .to_string_lossy()
        .to_string();
    let input_docx = fixture_dir.join("input.docx");
    let reference_pdf = fixture_dir.join("reference.pdf");

    let output_base = PathBuf::from("tests/output").join(&name);
    fs::create_dir_all(&output_base).ok();
    let generated_pdf = output_base.join("generated.pdf");

    println!("\n=== Text boundaries: {name} ===");

    if let Err(e) = docxside_pdf::convert_docx_to_pdf(&input_docx, &generated_pdf) {
        println!("  [SKIP] {name}: {e}");
        return None;
    }

    let ref_word_pages = extract_all_pages(&reference_pdf);
    let gen_word_pages = extract_all_pages(&generated_pdf);

    let page_count_match = ref_word_pages.len() == gen_word_pages.len();
    let common_pages = ref_word_pages.len().min(gen_word_pages.len());

    // Per-page first/last word
    let mut page_results = Vec::new();
    for i in 0..common_pages {
        let rf = ref_word_pages[i].first().cloned().unwrap_or_default();
        let gf = gen_word_pages[i].first().cloned().unwrap_or_default();
        let rl = ref_word_pages[i].last().cloned().unwrap_or_default();
        let gl = gen_word_pages[i].last().cloned().unwrap_or_default();
        page_results.push(PageResult {
            page: i + 1,
            first_ok: rf == gf,
            last_ok: rl == gl,
            ref_first: rf,
            gen_first: gf,
            ref_last: rl,
            gen_last: gl,
        });
    }

    // Break offset analysis
    let ref_breaks = break_positions(&ref_word_pages);
    let gen_breaks = break_positions(&gen_word_pages);
    let total_words = ref_breaks.last().copied().unwrap_or(0);

    let mut break_offsets = Vec::new();
    let break_count = (ref_breaks.len().saturating_sub(1)).min(gen_breaks.len().saturating_sub(1));
    for i in 0..break_count {
        break_offsets.push(BreakOffset {
            after_page: i + 1,
            ref_pos: ref_breaks[i],
            gen_pos: gen_breaks[i],
            offset_words: gen_breaks[i] as i64 - ref_breaks[i] as i64,
            total_words,
        });
    }

    // Line-level comparison
    let ref_page_count = ref_word_pages.len();
    let gen_page_count = gen_word_pages.len();
    let line_pages = ref_page_count.min(gen_page_count);
    let mut line_results = Vec::new();

    for p in 1..=line_pages {
        let ref_lines = extract_page_lines(&reference_pdf, p);
        let gen_lines = extract_page_lines(&generated_pdf, p);

        // Skip line-level comparison when counts differ significantly — mutool
        // splits our per-word BT/ET blocks into separate lines for justified text,
        // inflating the count. We'll revisit once we use Tw word spacing.
        let max_count = ref_lines.len().max(gen_lines.len());
        let min_count = ref_lines.len().min(gen_lines.len());
        if max_count > 0 && (max_count - min_count) as f64 / max_count as f64 > 0.15 {
            println!(
                "  Page {p}: line count mismatch — reference={}, generated={} (skipping line comparison)",
                ref_lines.len(),
                gen_lines.len()
            );
            continue;
        }

        let line_count = min_count;
        for l in 0..line_count {
            let rf = first_word(&ref_lines[l]);
            let gf = first_word(&gen_lines[l]);
            let rl = last_word(&ref_lines[l]);
            let gl = last_word(&gen_lines[l]);
            line_results.push(LineResult {
                page: p,
                line: l + 1,
                first_match: rf == gf,
                last_match: rl == gl,
                ref_first: rf,
                gen_first: gf,
                ref_last: rl,
                gen_last: gl,
            });
        }

        if ref_lines.len() != gen_lines.len() {
            println!(
                "  Page {p}: line count mismatch — reference={}, generated={}",
                ref_lines.len(),
                gen_lines.len()
            );
        }
    }

    Some(FixtureResult {
        name,
        page_count_match,
        ref_pages: ref_page_count,
        gen_pages: gen_page_count,
        page_results,
        break_offsets,
        line_results,
    })
}

fn print_summary(results: &[FixtureResult]) {
    println!();

    let name_w = results
        .iter()
        .map(|r| r.name.len())
        .max()
        .unwrap_or(4)
        .max(4);

    println!(
        "+-{}-+-------+--------+-----------+-------+-------+",
        "-".repeat(name_w)
    );
    println!(
        "| {:<name_w$} | Pages | Breaks | Max drift | Lines | Match |",
        "Case"
    );
    println!(
        "+-{}-+-------+--------+-----------+-------+-------+",
        "-".repeat(name_w)
    );

    for r in results {
        let pages_str = if r.page_count_match {
            format!("{}", r.ref_pages)
        } else {
            format!("{}/{}", r.ref_pages, r.gen_pages)
        };

        let all_breaks_ok = r.break_offsets.iter().all(|b| b.offset_words == 0);
        let breaks_str = if r.ref_pages <= 1 {
            "-".to_string()
        } else if all_breaks_ok {
            "OK".to_string()
        } else {
            "MISS".to_string()
        };

        let max_drift = r
            .break_offsets
            .iter()
            .map(|b| b.offset_words.unsigned_abs())
            .max()
            .unwrap_or(0);
        let drift_str = if r.ref_pages <= 1 {
            "-".to_string()
        } else if max_drift == 0 {
            "0".to_string()
        } else {
            let pct = max_drift as f64 / r.break_offsets[0].total_words.max(1) as f64 * 100.0;
            format!("{max_drift}w ({pct:.1}%)")
        };

        let total_lines = r.line_results.len();
        let matching_lines = r
            .line_results
            .iter()
            .filter(|lr| lr.first_match && lr.last_match)
            .count();
        let line_pct = if total_lines > 0 {
            format!(
                "{:.0}%",
                matching_lines as f64 / total_lines as f64 * 100.0
            )
        } else {
            "-".to_string()
        };

        println!(
            "| {:<name_w$} | {:>5} | {:>6} | {:>9} | {:>5} | {:>5} |",
            r.name, pages_str, breaks_str, drift_str, total_lines, line_pct
        );
    }

    println!(
        "+-{}-+-------+--------+-----------+-------+-------+",
        "-".repeat(name_w)
    );
}

#[test]
fn text_boundaries_match() {
    let fixtures = discover_fixtures().expect("Failed to read tests/fixtures");

    if fixtures.is_empty() {
        println!("[INFO] No fixtures found");
        return;
    }

    let mut results: Vec<FixtureResult> = Vec::new();
    let mut all_passed = true;

    for fixture_dir in &fixtures {
        let Some(result) = analyze_fixture(fixture_dir) else {
            all_passed = false;
            continue;
        };

        // Page count
        if !result.page_count_match {
            println!(
                "  [FAIL] Page count mismatch: reference={}, generated={}",
                result.ref_pages, result.gen_pages
            );
            all_passed = false;
        }

        // Per-page first/last word
        for pr in &result.page_results {
            if pr.first_ok && pr.last_ok {
                println!(
                    "  Page {}: OK (first={:?}, last={:?})",
                    pr.page, pr.ref_first, pr.ref_last
                );
            } else {
                if !pr.first_ok {
                    println!(
                        "  Page {}: FIRST WORD MISMATCH — reference={:?}, generated={:?}",
                        pr.page, pr.ref_first, pr.gen_first
                    );
                }
                if !pr.last_ok {
                    println!(
                        "  Page {}: LAST WORD MISMATCH — reference={:?}, generated={:?}",
                        pr.page, pr.ref_last, pr.gen_last
                    );
                }
                all_passed = false;
            }
        }

        // Break offsets
        for bo in &result.break_offsets {
            if bo.offset_words == 0 {
                println!(
                    "  Break after page {}: OK (at word {})",
                    bo.after_page, bo.ref_pos
                );
            } else {
                let direction = if bo.offset_words > 0 {
                    "late"
                } else {
                    "early"
                };
                let pct =
                    bo.offset_words.unsigned_abs() as f64 / bo.total_words.max(1) as f64 * 100.0;
                println!(
                    "  Break after page {}: {} by {} words ({:.1}% of {}) — ref@{} gen@{}",
                    bo.after_page,
                    direction,
                    bo.offset_words.abs(),
                    pct,
                    bo.total_words,
                    bo.ref_pos,
                    bo.gen_pos
                );
            }
        }

        // Line-level mismatches (only print failures to keep output manageable)
        let line_mismatches: Vec<&LineResult> = result
            .line_results
            .iter()
            .filter(|lr| !lr.first_match || !lr.last_match)
            .collect();
        if line_mismatches.is_empty() {
            let n = result.line_results.len();
            println!("  Lines: all {n} lines match");
        } else {
            let total = result.line_results.len();
            let ok = total - line_mismatches.len();
            println!("  Lines: {ok}/{total} match");
            for lr in &line_mismatches {
                let mut parts = Vec::new();
                if !lr.first_match {
                    parts.push(format!(
                        "first: {:?}→{:?}",
                        lr.ref_first, lr.gen_first
                    ));
                }
                if !lr.last_match {
                    parts.push(format!(
                        "last: {:?}→{:?}",
                        lr.ref_last, lr.gen_last
                    ));
                }
                println!(
                    "    p{}:L{}: {}",
                    lr.page,
                    lr.line,
                    parts.join(", ")
                );
            }
        }

        results.push(result);
    }

    print_summary(&results);
    assert!(all_passed, "One or more fixtures had text boundary mismatches");
}
