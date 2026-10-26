use crate::common;

use rayon::prelude::*;
use std::path::Path;

fn pdf_mediabox(pdf: &Path) -> Option<(f32, f32)> {
    let text = common::mutool_info(pdf).ok()?;
    let mut in_mediaboxes = false;
    for line in text.lines() {
        if line.starts_with("Mediaboxes") {
            in_mediaboxes = true;
            continue;
        }
        if in_mediaboxes {
            if let Some(bracket_start) = line.find('[') {
                let bracket_end = line.find(']')?;
                let nums: Vec<f32> = line[bracket_start + 1..bracket_end]
                    .split_whitespace()
                    .filter_map(|s| s.parse().ok())
                    .collect();
                if nums.len() == 4 {
                    return Some((nums[2] - nums[0], nums[3] - nums[1]));
                }
            }
            break;
        }
    }
    None
}

/// `pdf_mediabox`, kept in `cache` while the PDF is unchanged: the mutool spawn is the
/// whole cost of this suite.
fn mediabox_cached(pdf: &Path, cache: &Path) -> Option<(f32, f32)> {
    common::cached(cache, &[pdf], || Some(pdf_mediabox(pdf))).flatten()
}

/// One fixture's table row and, when the page sizes differ, the mismatch to report.
fn check(fixture: &Path) -> Option<(String, Option<String>)> {
    let name = common::case_key(fixture);
    let input = fixture.join("input.docx");
    let reference = fixture.join("reference.pdf");

    if !input.exists() || !reference.exists() {
        return None;
    }
    let out = common::output_dir(fixture);
    let _ = std::fs::create_dir_all(&out);

    let Some(ref_dims) = mediabox_cached(&reference, &out.join("reference.mediabox.json")) else {
        return Some((
            format!("  {:<64}  {:<14}  {:<14}  {:<5}", name, "?", "?", "SKIP"),
            None,
        ));
    };

    let gen_pdf = match common::ensure_generated_pdf(fixture) {
        Ok(p) => p,
        Err(e) => {
            return Some((
                format!(
                    "  {:<64}  {:>5.1}x{:<6.1}  {:<14}  {:<5}",
                    name,
                    ref_dims.0,
                    ref_dims.1,
                    format!("ERR:{e}"),
                    "FAIL"
                ),
                Some(format!("{name}: conversion error: {e}")),
            ));
        }
    };

    let Some(gen_dims) = mediabox_cached(&gen_pdf, &out.join("generated.mediabox.json")) else {
        return Some((
            format!(
                "  {:<64}  {:>5.1}x{:<6.1}  {:<14}  {:<5}",
                name, ref_dims.0, ref_dims.1, "?", "FAIL"
            ),
            Some(format!("{name}: could not read generated mediabox")),
        ));
    };

    let w_ok = (ref_dims.0 - gen_dims.0).abs() < 1.0;
    let h_ok = (ref_dims.1 - gen_dims.1).abs() < 1.0;
    let ok = w_ok && h_ok;

    let status = if ok { "OK" } else { "FAIL" };
    let mismatch = (!ok).then(|| {
        format!(
            "{name}: ref={:.1}x{:.1} gen={:.1}x{:.1} (dW={:+.1} dH={:+.1})",
            ref_dims.0,
            ref_dims.1,
            gen_dims.0,
            gen_dims.1,
            gen_dims.0 - ref_dims.0,
            gen_dims.1 - ref_dims.1,
        )
    });

    Some((
        format!(
            "  {:<64}  {:>5.1}x{:<6.1}  {:>5.1}x{:<6.1}  {:<5}",
            name, ref_dims.0, ref_dims.1, gen_dims.0, gen_dims.1, status
        ),
        mismatch,
    ))
}

#[test]
fn page_geometry_comparison() {
    let fixtures = common::discover_fixtures().expect("discover fixtures");

    println!();
    println!(
        "  {:<64}  {:<14}  {:<14}  {:<5}",
        "Case", "Reference", "Generated", "Match"
    );

    let rows: Vec<(String, Option<String>)> =
        fixtures.par_iter().filter_map(|f| check(f)).collect();
    for (row, _) in &rows {
        println!("{row}");
    }

    let mismatches: Vec<&String> = rows.iter().filter_map(|(_, m)| m.as_ref()).collect();
    if !mismatches.is_empty() {
        println!("\nPage size mismatches ({}):", mismatches.len());
        for m in &mismatches {
            println!("  - {m}");
        }
    }
}
