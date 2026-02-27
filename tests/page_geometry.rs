mod common;

use std::path::Path;
use std::process::Command;

fn pdf_mediabox(pdf: &Path) -> Option<(f32, f32)> {
    let output = Command::new("mutool")
        .args(["info", pdf.to_str().unwrap()])
        .output()
        .ok()?;
    let text = String::from_utf8_lossy(&output.stdout);
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

#[test]
fn page_geometry_comparison() {
    let fixtures = common::discover_fixtures().expect("discover fixtures");

    println!();
    println!(
        "+{:-<66}+{:-<16}+{:-<16}+{:-<7}+",
        "", "", "", ""
    );
    println!(
        "| {:<64} | {:<14} | {:<14} | {:<5} |",
        "Case", "Reference", "Generated", "Match"
    );
    println!(
        "+{:-<66}+{:-<16}+{:-<16}+{:-<7}+",
        "", "", "", ""
    );

    let mut mismatches = Vec::new();

    for fixture in &fixtures {
        let name = common::display_name(fixture);
        let input = fixture.join("input.docx");
        let reference = fixture.join("reference.pdf");

        if !input.exists() || !reference.exists() {
            continue;
        }

        let ref_dims = match pdf_mediabox(&reference) {
            Some(d) => d,
            None => {
                println!("| {:<64} | {:<14} | {:<14} | {:<5} |", name, "?", "?", "SKIP");
                continue;
            }
        };

        let case_out = common::output_dir(fixture);
        std::fs::create_dir_all(&case_out).ok();
        let gen_pdf = case_out.join("generated.pdf");

        if let Err(e) = docxide_pdf::convert_docx_to_pdf(&input, &gen_pdf) {
            println!(
                "| {:<64} | {:>5.1}x{:<6.1} | {:<14} | {:<5} |",
                name,
                ref_dims.0,
                ref_dims.1,
                format!("ERR:{e}"),
                "FAIL"
            );
            mismatches.push(format!("{name}: conversion error: {e}"));
            continue;
        }

        let gen_dims = match pdf_mediabox(&gen_pdf) {
            Some(d) => d,
            None => {
                println!(
                    "| {:<64} | {:>5.1}x{:<6.1} | {:<14} | {:<5} |",
                    name, ref_dims.0, ref_dims.1, "?", "FAIL"
                );
                mismatches.push(format!("{name}: could not read generated mediabox"));
                continue;
            }
        };

        let w_ok = (ref_dims.0 - gen_dims.0).abs() < 1.0;
        let h_ok = (ref_dims.1 - gen_dims.1).abs() < 1.0;
        let ok = w_ok && h_ok;

        let status = if ok { "OK" } else { "FAIL" };
        if !ok {
            mismatches.push(format!(
                "{name}: ref={:.1}x{:.1} gen={:.1}x{:.1} (dW={:+.1} dH={:+.1})",
                ref_dims.0,
                ref_dims.1,
                gen_dims.0,
                gen_dims.1,
                gen_dims.0 - ref_dims.0,
                gen_dims.1 - ref_dims.1,
            ));
        }

        println!(
            "| {:<64} | {:>5.1}x{:<6.1} | {:>5.1}x{:<6.1} | {:<5} |",
            name, ref_dims.0, ref_dims.1, gen_dims.0, gen_dims.1, status
        );
    }

    println!(
        "+{:-<66}+{:-<16}+{:-<16}+{:-<7}+",
        "", "", "", ""
    );

    if !mismatches.is_empty() {
        println!("\nPage size mismatches ({}):", mismatches.len());
        for m in &mismatches {
            println!("  - {m}");
        }
    }
}
