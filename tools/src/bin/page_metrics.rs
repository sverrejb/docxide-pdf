//! Harness-identical metrics for one PDF pair: Jaccard, SSIM, text boundary, accessibility.
//!
//! Usage:
//!   page-metrics <ref.pdf> <other.pdf> <ref_png_dir> <other_png_dir>
//!
//! Prints one JSON object. Jaccard and SSIM are per-page means over the pages both
//! PDFs have, exactly like tests/suite/visual_comparison.rs; the metric code itself is the
//! harness's own (tests/common/metrics.rs, a11y.rs, text_boundary.rs), included by path so
//! the numbers cannot drift. Not tests/common/mod.rs: its convert() links the whole library,
//! which had this binary rebuilt after every src/ change.
//! Accessibility scores are those of tests/suite/accessibility.rs, null without veraPDF and pdfinfo.
#[path = "../../../tests/common/metrics.rs"]
#[allow(dead_code)]
mod metrics;
#[path = "../../../tests/common/a11y.rs"]
#[allow(dead_code)]
mod a11y;
#[path = "../../../tests/common/text_boundary.rs"]
#[allow(dead_code)]
mod text_boundary;

use rayon::prelude::*;
use std::path::Path;

/// A PNG dir is named after its PDF, so for the harness's own output dirs the
/// analysis cache is the very file tests/suite/accessibility.rs writes and either side
/// reuses the other's work.
fn a11y_metrics(ref_pdf: &Path, other_pdf: &Path, ref_dir: &Path, other_dir: &Path) -> Option<serde_json::Value> {
    let r = a11y::analyze_cached(ref_pdf, &a11y::cache_path(ref_dir)).ok()?;
    let g = a11y::analyze_cached(other_pdf, &a11y::cache_path(other_dir)).ok()?;
    let s = a11y::scores(&r, &g);
    let v = s.vs_word.as_ref();
    Some(serde_json::json!({
        "ua_fail": s.ua_fail,
        "ua_deficit": v.map(|v| v.deficit.len()),
        "a11y_struct": v.map(|v| v.struct_score),
        "a11y_text": v.map(|v| v.text_score),
    }))
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() != 5 {
        eprintln!("usage: page-metrics <ref.pdf> <other.pdf> <ref_png_dir> <other_png_dir>");
        std::process::exit(1);
    }
    let (ref_pdf, other_pdf) = (Path::new(&args[1]), Path::new(&args[2]));
    let ref_pngs = metrics::collect_page_pngs(Path::new(&args[3])).unwrap_or_default();
    let other_pngs = metrics::collect_page_pngs(Path::new(&args[4])).unwrap_or_default();

    let n = ref_pngs.len().min(other_pngs.len());
    let pages: Vec<(f64, f64)> = (0..n)
        .into_par_iter()
        .filter_map(|i| {
            let a = image::open(&ref_pngs[i]).ok()?;
            let b = image::open(&other_pngs[i]).ok()?;
            let jaccard = metrics::compare_and_diff(&a, &b).ok()?.jaccard;
            let ssim = metrics::ssim_score(&a, &b).ok()?;
            Some((jaccard, ssim))
        })
        .collect();
    let mean = |pick: fn(&(f64, f64)) -> f64| {
        (!pages.is_empty()).then(|| pages.iter().map(pick).sum::<f64>() / pages.len() as f64)
    };

    let tb = text_boundary::analyze(ref_pdf, other_pdf);
    let out = serde_json::json!({
        "jaccard": mean(|p| p.0),
        "ssim": mean(|p| p.1),
        "text_boundary": tb.line_match_pct(),
        "ref_pages": ref_pngs.len(),
        "pages": other_pngs.len(),
        "scored_pages": pages.len(),
        "max_break_drift": tb.max_break_drift,
        "a11y": a11y_metrics(ref_pdf, other_pdf, Path::new(&args[3]), Path::new(&args[4])),
    });
    println!("{out}");
}
