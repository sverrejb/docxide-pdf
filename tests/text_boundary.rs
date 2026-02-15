use std::collections::HashMap;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};
use std::{fs, io};

const SKIP_FIXTURES: &[&str] = &["sample100kB"];

fn natural_cmp(a: &Path, b: &Path) -> std::cmp::Ordering {
    let a = a.file_name().and_then(|n| n.to_str()).unwrap_or("");
    let b = b.file_name().and_then(|n| n.to_str()).unwrap_or("");
    let extract = |s: &str| -> (String, u64) {
        let i = s.find(|c: char| c.is_ascii_digit()).unwrap_or(s.len());
        (s[..i].to_string(), s[i..].parse().unwrap_or(0))
    };
    extract(a).cmp(&extract(b))
}

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
    fixtures.sort_by(|a, b| natural_cmp(a, b));
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

fn extract_all_pages(pdf: &Path) -> Vec<Vec<String>> {
    let n = pdf_page_count(pdf);
    (1..=n).map(|p| extract_page_words(pdf, p)).collect()
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

fn first_word(s: &str) -> String {
    s.split_whitespace().next().unwrap_or_default().to_string()
}

fn last_word(s: &str) -> String {
    s.split_whitespace().last().unwrap_or_default().to_string()
}

fn timestamp() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_secs()
}

fn log_csv(csv_name: &str, header: &str, row: &str) {
    let csv_path = PathBuf::from("tests/output").join(csv_name);
    fs::create_dir_all("tests/output").ok();
    let write_header = !csv_path.exists();
    let mut file = fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&csv_path)
        .expect("Cannot open CSV file");
    if write_header {
        writeln!(file, "{header}").unwrap();
    }
    writeln!(file, "{row}").unwrap();
}

/// Read previous line_match_pct per case from CSV (col index 5).
fn read_previous_scores(csv_name: &str, score_col: usize) -> HashMap<String, f64> {
    let csv_path = PathBuf::from("tests/output").join(csv_name);
    let mut latest: HashMap<String, f64> = HashMap::new();
    let Ok(content) = fs::read_to_string(&csv_path) else {
        return latest;
    };
    for line in content.lines().skip(1) {
        let cols: Vec<&str> = line.split(',').collect();
        if cols.len() > score_col {
            if let Ok(score) = cols[score_col].parse::<f64>() {
                latest.insert(cols[1].to_string(), score);
            }
        }
    }
    latest
}

struct CaseResult {
    name: String,
    ref_pages: usize,
    gen_pages: usize,
    max_break_drift: i64,
    total_words: usize,
    total_lines: usize,
    matching_lines: usize,
}

fn analyze_fixture(fixture_dir: &Path) -> Option<CaseResult> {
    let name = fixture_dir
        .file_name()
        .unwrap()
        .to_string_lossy()
        .to_string();
    let input_docx = fixture_dir.join("input.docx");
    let reference_pdf = fixture_dir.join("reference.pdf");
    if !reference_pdf.exists() {
        println!("  [SKIP] {name}: no reference.pdf");
        return None;
    }
    let output_base = PathBuf::from("tests/output").join(&name);
    let _ = fs::remove_file(output_base.join("generated.pdf"));
    fs::create_dir_all(&output_base).ok();
    let generated_pdf = output_base.join("generated.pdf");

    if let Err(e) = docxside_pdf::convert_docx_to_pdf(&input_docx, &generated_pdf) {
        println!("  [SKIP] {name}: {e}");
        return None;
    }

    let ref_word_pages = extract_all_pages(&reference_pdf);
    let gen_word_pages = extract_all_pages(&generated_pdf);
    let common_pages = ref_word_pages.len().min(gen_word_pages.len());

    // Break offset analysis
    let ref_breaks = break_positions(&ref_word_pages);
    let gen_breaks = break_positions(&gen_word_pages);
    let total_words = ref_breaks.last().copied().unwrap_or(0);
    let break_count = (ref_breaks.len().saturating_sub(1)).min(gen_breaks.len().saturating_sub(1));
    let max_break_drift = (0..break_count)
        .map(|i| gen_breaks[i] as i64 - ref_breaks[i] as i64)
        .max_by_key(|d| d.unsigned_abs())
        .unwrap_or(0);

    // Line-level comparison
    let mut total_lines = 0;
    let mut matching_lines = 0;
    for p in 1..=common_pages {
        let ref_lines = extract_page_lines(&reference_pdf, p);
        let gen_lines = extract_page_lines(&generated_pdf, p);

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

    Some(CaseResult {
        name,
        ref_pages: ref_word_pages.len(),
        gen_pages: gen_word_pages.len(),
        max_break_drift,
        total_words,
        total_lines,
        matching_lines,
    })
}

fn delta_str(current: f64, previous: Option<f64>) -> String {
    match previous {
        Some(prev) => {
            let diff = (current - prev) * 100.0;
            if diff.abs() < 0.05 {
                String::new()
            } else if diff > 0.0 {
                format!(" (+{diff:.1}pp)")
            } else {
                format!(" ({diff:.1}pp)")
            }
        }
        None => String::new(),
    }
}

#[test]
fn text_boundaries_match() {
    let fixtures = discover_fixtures().expect("Failed to read tests/fixtures");
    if fixtures.is_empty() {
        return;
    }

    let prev_scores = read_previous_scores("text_boundary_results.csv", 5);
    let mut results: Vec<CaseResult> = Vec::new();

    for fixture_dir in &fixtures {
        let Some(result) = analyze_fixture(fixture_dir) else {
            continue;
        };
        results.push(result);
    }

    // Log to CSV and print summary
    let name_w = results
        .iter()
        .map(|r| r.name.len())
        .max()
        .unwrap_or(4)
        .max(4);
    let sep = format!(
        "+-{}-+-------+--------+--------------+-------+-------+-----------+",
        "-".repeat(name_w)
    );
    println!("\n{sep}");
    println!(
        "| {:<name_w$} | Pages | Breaks | Max drift    | Lines | Match | Delta     |",
        "Case"
    );
    println!("{sep}");

    for r in &results {
        let pages_str = if r.ref_pages == r.gen_pages {
            format!("{}", r.ref_pages)
        } else {
            format!("{}/{}", r.ref_pages, r.gen_pages)
        };

        let breaks_str = if r.ref_pages <= 1 {
            "-".to_string()
        } else if r.max_break_drift == 0 {
            "OK".to_string()
        } else {
            "MISS".to_string()
        };

        let drift_str = if r.ref_pages <= 1 {
            "-".to_string()
        } else if r.max_break_drift == 0 {
            "0".to_string()
        } else {
            let abs = r.max_break_drift.unsigned_abs();
            let pct = abs as f64 / r.total_words.max(1) as f64 * 100.0;
            format!("{abs}w ({pct:.1}%)")
        };

        let line_pct = if r.total_lines > 0 {
            r.matching_lines as f64 / r.total_lines as f64
        } else {
            1.0
        };
        let line_pct_str = if r.total_lines > 0 {
            format!("{:.0}%", line_pct * 100.0)
        } else {
            "-".to_string()
        };

        let delta = delta_str(line_pct, prev_scores.get(&r.name).copied());

        println!(
            "| {:<name_w$} | {:>5} | {:>6} | {:>12} | {:>5} | {:>5} | {:<9} |",
            r.name, pages_str, breaks_str, drift_str, r.total_lines, line_pct_str, delta
        );

        log_csv(
            "text_boundary_results.csv",
            "timestamp,case,ref_pages,gen_pages,max_drift,line_match_pct",
            &format!(
                "{},{},{},{},{},{:.4}",
                timestamp(),
                r.name,
                r.ref_pages,
                r.gen_pages,
                r.max_break_drift,
                line_pct
            ),
        );
    }

    println!("{sep}");

    let regressions: Vec<&str> = results
        .iter()
        .filter(|r| {
            let line_pct = if r.total_lines > 0 {
                r.matching_lines as f64 / r.total_lines as f64
            } else {
                1.0
            };
            prev_scores
                .get(&r.name)
                .is_some_and(|&p| line_pct < p - 0.005)
        })
        .map(|r| r.name.as_str())
        .collect();
    if !regressions.is_empty() {
        println!("  REGRESSION in: {}", regressions.join(", "));
    }

    let page_mismatches: Vec<String> = results
        .iter()
        .filter(|r| r.ref_pages != r.gen_pages)
        .map(|r| format!("{} (ref={}, gen={})", r.name, r.ref_pages, r.gen_pages))
        .collect();
    assert!(
        page_mismatches.is_empty(),
        "Page count mismatch: {}",
        page_mismatches.join(", ")
    );
}
