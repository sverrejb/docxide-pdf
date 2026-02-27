mod common;

use rayon::prelude::*;
use std::fs;
use std::path::Path;

const SIZE_RATIO_THRESHOLD: f64 = 10.0;

struct SizeResult {
    name: String,
    gen_bytes: u64,
    ref_bytes: u64,
    ratio: f64,
    pass: bool,
}

fn analyze_fixture(fixture_dir: &Path) -> Option<SizeResult> {
    let name = common::display_name(fixture_dir);
    let input_docx = fixture_dir.join("input.docx");
    let reference_pdf = fixture_dir.join("reference.pdf");
    if !input_docx.exists() || !reference_pdf.exists() {
        return None;
    }

    let output_dir = common::output_dir(fixture_dir);
    fs::create_dir_all(&output_dir).ok();
    let generated_pdf = output_dir.join("generated.pdf");

    let needs_convert = !generated_pdf.exists() || {
        let docx_mtime = fs::metadata(&input_docx)
            .and_then(|m| m.modified())
            .unwrap_or(std::time::SystemTime::UNIX_EPOCH);
        let pdf_mtime = fs::metadata(&generated_pdf)
            .and_then(|m| m.modified())
            .unwrap_or(std::time::SystemTime::UNIX_EPOCH);
        pdf_mtime < docx_mtime
    };
    if needs_convert {
        if let Err(e) = docxide_pdf::convert_docx_to_pdf(&input_docx, &generated_pdf) {
            println!("  [SKIP] {name}: {e}");
            return None;
        }
    }

    let gen_bytes = fs::metadata(&generated_pdf).map(|m| m.len()).unwrap_or(0);
    let ref_bytes = fs::metadata(&reference_pdf).map(|m| m.len()).unwrap_or(0);
    let ratio = if ref_bytes > 0 {
        gen_bytes as f64 / ref_bytes as f64
    } else {
        0.0
    };
    let pass = ratio <= SIZE_RATIO_THRESHOLD;

    Some(SizeResult {
        name,
        gen_bytes,
        ref_bytes,
        ratio,
        pass,
    })
}

/// ANSI 256-color gradient from green (ratio ≤ 1) to dark red (ratio ≥ 10).
fn color_ratio(ratio: f64, text: &str) -> String {
    let t = ((ratio - 1.0) / 9.0).clamp(0.0, 1.0);
    let r = (80.0 + 175.0 * t) as u8;
    let g = (200.0 * (1.0 - t)) as u8;
    let b = (80.0 * (1.0 - t)) as u8;
    format!("\x1b[38;2;{r};{g};{b}m{text}\x1b[0m")
}

fn human_size(bytes: u64) -> String {
    if bytes >= 1_000_000 {
        format!("{:.1}MB", bytes as f64 / 1_000_000.0)
    } else if bytes >= 1_000 {
        format!("{:.0}kB", bytes as f64 / 1_000.0)
    } else {
        format!("{bytes}B")
    }
}

#[test]
fn file_size_within_threshold() {
    let _ = env_logger::try_init();
    let fixtures = common::discover_fixtures().expect("Failed to read tests/fixtures");
    if fixtures.is_empty() {
        return;
    }

    let mut results: Vec<SizeResult> = fixtures
        .par_iter()
        .filter_map(|f| analyze_fixture(f))
        .collect();
    results.sort_by(|a, b| a.name.cmp(&b.name));

    let ts = common::timestamp();
    let name_w = results
        .iter()
        .map(|r| r.name.len())
        .max()
        .unwrap_or(4)
        .max(4);

    let sep = format!(
        "+-{}-+------+-----------+-----------+-------+",
        "-".repeat(name_w)
    );

    println!("\n{sep}");
    println!(
        "| {:<name_w$} | Pass | {:<9} | {:<9} | Ratio |",
        "Case", "Generated", "Reference"
    );
    println!("{sep}");

    let mut all_pass = true;
    for r in &results {
        let status = if r.pass { "Y" } else { "N" };
        let ratio_str = format!("{:>5.1}", r.ratio);
        let colored_ratio = color_ratio(r.ratio, &ratio_str);
        println!(
            "| {:<name_w$} | {:<4} | {:>9} | {:>9} | {} |",
            r.name,
            status,
            human_size(r.gen_bytes),
            human_size(r.ref_bytes),
            colored_ratio
        );
        if !r.pass {
            all_pass = false;
        }

        common::log_csv(
            "file_size_results.csv",
            "timestamp,case,gen_bytes,ref_bytes,ratio,pass",
            &format!(
                "{},{},{},{},{:.2},{}",
                ts, r.name, r.gen_bytes, r.ref_bytes, r.ratio, r.pass
            ),
        );
    }

    println!("{sep}");
    println!("  threshold: generated <= {SIZE_RATIO_THRESHOLD:.0}x reference");
    assert!(
        all_pass,
        "Some fixtures exceed the file size threshold — see details above"
    );
}
