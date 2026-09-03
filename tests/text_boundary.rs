mod common;

use common::text_boundary::TextBoundary;
use rayon::prelude::*;
use std::collections::HashMap;
use std::path::Path;

struct CaseResult {
    name: String,
    tb: TextBoundary,
}

fn analyze_fixture(fixture_dir: &Path) -> Option<CaseResult> {
    let name = common::display_name(fixture_dir);
    let reference_pdf = fixture_dir.join("reference.pdf");
    if !reference_pdf.exists() {
        println!("  [SKIP] {name}: no reference.pdf");
        return None;
    }
    let generated_pdf = match common::ensure_generated_pdf(fixture_dir) {
        Ok(p) => p,
        Err(e) => {
            println!("  [SKIP] {name}: {e}");
            return None;
        }
    };
    let tb = common::text_boundary::analyze(&reference_pdf, &generated_pdf);
    Some(CaseResult { name, tb })
}

fn text_boundaries_match() {
    let _ = env_logger::try_init();
    let fixtures = common::discover_fixtures().expect("Failed to read tests/fixtures");
    if fixtures.is_empty() {
        return;
    }

    let baselines = common::read_baselines();
    let prev_scores: HashMap<String, f64> = baselines
        .iter()
        .filter_map(|(k, v)| v.text_boundary.map(|t| (k.clone(), t)))
        .collect();

    let mut results: Vec<CaseResult> = fixtures
        .par_iter()
        .filter_map(|fixture_dir| analyze_fixture(fixture_dir))
        .collect();
    results.sort_by(|a, b| a.name.cmp(&b.name));

    let name_w = results
        .iter()
        .map(|r| r.name.len())
        .max()
        .unwrap_or(4)
        .max(4);
    println!(
        "\n  {:<name_w$}  Pages  Breaks  Max drift      Lines  Match  Delta",
        "Case"
    );

    for r in &results {
        let pages_str = if r.tb.ref_pages == r.tb.gen_pages {
            format!("{}", r.tb.ref_pages)
        } else {
            format!("{}/{}", r.tb.ref_pages, r.tb.gen_pages)
        };

        let breaks_str = if r.tb.ref_pages <= 1 {
            "-".to_string()
        } else if r.tb.max_break_drift == 0 {
            "OK".to_string()
        } else {
            "MISS".to_string()
        };

        let drift_str = if r.tb.ref_pages <= 1 {
            "-".to_string()
        } else if r.tb.max_break_drift == 0 {
            "0".to_string()
        } else {
            let abs = r.tb.max_break_drift.unsigned_abs();
            let pct = abs as f64 / r.tb.total_words.max(1) as f64 * 100.0;
            format!("{abs}w ({pct:.1}%)")
        };

        let line_pct = r.tb.line_match_pct();
        let line_pct_str = if r.tb.total_lines > 0 {
            format!("{:.0}%", line_pct * 100.0)
        } else {
            "-".to_string()
        };

        let delta = common::delta_str(line_pct, prev_scores.get(&r.name).copied());

        println!(
            "  {:<name_w$}  {:>5}  {:>6}  {:>12}  {:>5}  {:>5}  {:<9}",
            r.name, pages_str, breaks_str, drift_str, r.tb.total_lines, line_pct_str, delta
        );

        common::log_csv(
            "text_boundary_results.csv",
            "timestamp,case,ref_pages,gen_pages,max_drift,line_match_pct",
            &format!(
                "{},{},{},{},{},{:.4}",
                common::timestamp(),
                r.name,
                r.tb.ref_pages,
                r.tb.gen_pages,
                r.tb.max_break_drift,
                line_pct
            ),
        );
    }

    let mut baseline_updates: HashMap<String, common::Baselines> = HashMap::new();
    for r in &results {
        baseline_updates.insert(
            r.name.clone(),
            common::Baselines {
                jaccard: None,
                ssim: None,
                text_boundary: Some(r.tb.line_match_pct()),
                convert_ms: None,
            },
        );
    }
    common::write_latest_scores(&baseline_updates);

    let regressions: Vec<&str> = results
        .iter()
        .filter(|r| {
            prev_scores
                .get(&r.name)
                .is_some_and(|&p| r.tb.line_match_pct() < p - common::REGRESSION_SLACK)
        })
        .map(|r| r.name.as_str())
        .collect();
    if !regressions.is_empty() {
        println!("  REGRESSION in: {}", regressions.join(", "));
    }

    let page_mismatches: Vec<String> = results
        .iter()
        .filter(|r| r.tb.ref_pages != r.tb.gen_pages)
        .map(|r| format!("{} (ref={}, gen={})", r.name, r.tb.ref_pages, r.tb.gen_pages))
        .collect();
    if !page_mismatches.is_empty() {
        println!("  PAGE COUNT MISMATCH: {}", page_mismatches.join(", "));
    }
}
