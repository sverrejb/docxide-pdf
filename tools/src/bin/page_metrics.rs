//! Harness-identical metrics for one PDF pair: Jaccard, SSIM, text boundary, accessibility.
//!
//! Usage:
//!   page-metrics <ref.pdf> <other.pdf> <ref_png_dir> <other_png_dir>
//!   page-metrics --a11y <pdf> <png_dir> [<pdf> <png_dir>...]
//!
//! The second form only runs the accessibility analyses, all in one veraPDF process, and leaves
//! their caches for the pair calls that follow: engine_compare.py calls it once per fixture so a
//! JVM starts per fixture rather than per PDF.
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
use std::path::{Path, PathBuf};

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

fn prefetch_a11y(pairs: &[String]) {
    let pairs: Vec<(&Path, PathBuf)> = pairs
        .chunks(2)
        .map(|c| (Path::new(&c[0]), a11y::cache_path(Path::new(&c[1]))))
        .collect();
    let stale: Vec<PathBuf> = pairs
        .iter()
        .filter(|(pdf, cache)| !a11y::cache_fresh(pdf, cache))
        .map(|(pdf, _)| pdf.to_path_buf())
        .collect();
    a11y::prefetch_ua_rules(&stale);
    pairs.par_iter().for_each(|(pdf, cache)| {
        let _ = a11y::analyze_cached(pdf, cache);
    });
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.get(1).map(String::as_str) == Some("--a11y") && args.len().is_multiple_of(2) {
        prefetch_a11y(&args[2..]);
        return;
    }
    if args.len() != 5 {
        eprintln!("usage: page-metrics <ref.pdf> <other.pdf> <ref_png_dir> <other_png_dir>");
        eprintln!("       page-metrics --a11y <pdf> <png_dir> [<pdf> <png_dir>...]");
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
