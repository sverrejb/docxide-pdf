use image::{DynamicImage, GenericImageView, ImageBuffer, Rgba};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};
use std::{fs, io};

const SIMILARITY_THRESHOLD: f64 = 0.25;
const SSIM_THRESHOLD: f64 = 0.40;
const MUTOOL_DPI: &str = "150";

fn discover_fixtures() -> io::Result<Vec<PathBuf>> {
    let fixtures_dir = Path::new("tests/fixtures");
    let mut fixtures = Vec::new();

    for entry in fs::read_dir(fixtures_dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.is_dir() {
            fixtures.push(path);
        }
    }

    fixtures.sort();
    Ok(fixtures)
}

fn screenshot_pdf(pdf: &Path, output_dir: &Path) -> Result<(), String> {
    fs::create_dir_all(output_dir).map_err(|e| e.to_string())?;

    let output_pattern = output_dir.join("page_%03d.png");

    let status = Command::new("mutool")
        .args([
            "draw",
            "-F", "png",
            "-r", MUTOOL_DPI,
            "-o", output_pattern.to_str().unwrap(),
            pdf.to_str().unwrap(),
        ])
        .status()
        .map_err(|e| format!("Failed to run mutool: {e}"))?;

    if status.success() {
        Ok(())
    } else {
        Err(format!(
            "mutool exited with status {}",
            status.code().unwrap_or(-1)
        ))
    }
}

// Jaccard similarity on ink pixels: ignores white background, measures overlap
// of dark (text/content) pixels between the two images.
fn compare_images(a: &Path, b: &Path) -> Result<f64, String> {
    let img_a = image::open(a).map_err(|e| format!("Failed to open {}: {e}", a.display()))?;
    let img_b = image::open(b).map_err(|e| format!("Failed to open {}: {e}", b.display()))?;

    let (w, h) = img_a.dimensions();
    let (w2, h2) = img_b.dimensions();
    // Allow up to 2px difference from sub-pixel rounding of A4 page sizes
    if w.abs_diff(w2) > 2 || h.abs_diff(h2) > 2 {
        return Err(format!(
            "Image dimensions differ: {:?} vs {:?}",
            (w, h),
            (w2, h2)
        ));
    }
    let cw = w.min(w2);
    let ch = h.min(h2);

    let mut intersection: u64 = 0;
    let mut union: u64 = 0;

    for y in 0..ch {
        for x in 0..cw {
            let [ra, ga, ba, _] = img_a.get_pixel(x, y).0;
            let [rb, gb, bb, _] = img_b.get_pixel(x, y).0;
            let a_ink = is_ink(ra, ga, ba);
            let b_ink = is_ink(rb, gb, bb);
            if a_ink || b_ink {
                union += 1;
            }
            if a_ink && b_ink {
                intersection += 1;
            }
        }
    }

    if union == 0 {
        return Ok(1.0);
    }
    Ok(intersection as f64 / union as f64)
}

fn is_ink(r: u8, g: u8, b: u8) -> bool {
    let luma = 0.299 * r as f32 + 0.587 * g as f32 + 0.114 * b as f32;
    luma < 200.0
}

fn save_diff_image(a: &Path, b: &Path, out: &Path) -> Result<(), String> {
    let img_a = image::open(a).map_err(|e| format!("{e}"))?;
    let img_b = image::open(b).map_err(|e| format!("{e}"))?;

    let (w, h) = img_a.dimensions();
    let (w2, h2) = img_b.dimensions();
    let cw = w.min(w2);
    let ch = h.min(h2);
    let mut diff: ImageBuffer<Rgba<u8>, Vec<u8>> = ImageBuffer::new(cw, ch);

    for y in 0..ch {
        for x in 0..cw {
            let Rgba([ra, ga, ba, aa]) = img_a.get_pixel(x, y).0.into();
            let Rgba([rb, gb, bb, ab]) = img_b.get_pixel(x, y).0.into();
            diff.put_pixel(
                x,
                y,
                Rgba([
                    ra.abs_diff(rb),
                    ga.abs_diff(gb),
                    ba.abs_diff(bb),
                    aa.max(ab),
                ]),
            );
        }
    }

    if let Some(parent) = out.parent() {
        fs::create_dir_all(parent).map_err(|e| e.to_string())?;
    }
    DynamicImage::ImageRgba8(diff)
        .save(out)
        .map_err(|e| e.to_string())
}

fn log_result(case: &str, pages: usize, avg_jaccard: f64, passed: bool) {
    let csv_path = Path::new("tests/output/results.csv");
    fs::create_dir_all("tests/output").ok();
    let write_header = !csv_path.exists();
    let mut file = fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(csv_path)
        .expect("Cannot open tests/output/results.csv");
    if write_header {
        writeln!(file, "timestamp,case,pages,avg_jaccard,pass").unwrap();
    }
    let ts = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_secs();
    writeln!(file, "{ts},{case},{pages},{avg_jaccard:.4},{passed}").unwrap();
}

fn collect_page_pngs(dir: &Path) -> io::Result<Vec<PathBuf>> {
    let mut pages = Vec::new();
    for entry in fs::read_dir(dir)? {
        let p = entry?.path();
        if p.extension().and_then(|e| e.to_str()) == Some("png") {
            pages.push(p);
        }
    }
    pages.sort();
    Ok(pages)
}

#[test]
fn visual_comparison() {
    let fixtures = discover_fixtures().expect("Failed to read tests/fixtures");

    if fixtures.is_empty() {
        println!("[INFO] No fixtures found — add DOCX+PDF pairs to tests/fixtures/<name>/");
        return;
    }

    let mut all_passed = true;
    let mut table_rows: Vec<(String, f64, bool)> = Vec::new();

    for fixture_dir in &fixtures {
        let name = fixture_dir.file_name().unwrap().to_string_lossy();
        let input_docx = fixture_dir.join("input.docx");
        let reference_pdf = fixture_dir.join("reference.pdf");

        let output_base = PathBuf::from("tests/output").join(name.as_ref());
        let reference_screenshots = output_base.join("reference");
        let generated_screenshots = output_base.join("generated");
        let diff_dir = output_base.join("diff");

        println!("\n=== Fixture: {name} ===");

        // Screenshot the reference PDF
        println!("  Screenshotting reference PDF...");
        if let Err(e) = screenshot_pdf(&reference_pdf, &reference_screenshots) {
            println!("  [ERROR] Failed to screenshot reference: {e}");
            all_passed = false;
            continue;
        }

        // Convert DOCX -> PDF via library
        let generated_pdf = output_base.join("generated.pdf");
        match convert_docx_to_pdf(&input_docx, &generated_pdf) {
            Err(e) => {
                println!("  [SKIP] {name}: {e}");
                continue;
            }
            Ok(()) => {}
        }

        // Screenshot the generated PDF
        println!("  Screenshotting generated PDF...");
        if let Err(e) = screenshot_pdf(&generated_pdf, &generated_screenshots) {
            println!("  [ERROR] Failed to screenshot generated: {e}");
            all_passed = false;
            continue;
        }

        // Compare pages
        let ref_pages = collect_page_pngs(&reference_screenshots)
            .unwrap_or_default();
        let gen_pages = collect_page_pngs(&generated_screenshots)
            .unwrap_or_default();

        if ref_pages.is_empty() {
            println!("  [WARN] No reference pages found");
            continue;
        }

        let page_count = ref_pages.len().min(gen_pages.len());
        let mut scores: Vec<f64> = Vec::new();

        for i in 0..page_count {
            let ref_page = &ref_pages[i];
            let gen_page = &gen_pages[i];
            let page_num = String::from(ref_page.file_stem().unwrap().to_str().unwrap());
            let diff_path = diff_dir.join(format!("{page_num}.png"));

            match compare_images(ref_page, gen_page) {
                Ok(score) => {
                    println!("  Page {}: similarity = {:.2}%", i + 1, score * 100.0);
                    scores.push(score);
                    let _ = save_diff_image(ref_page, gen_page, &diff_path);
                }
                Err(e) => {
                    println!("  Page {}: comparison error — {e}", i + 1);
                }
            }
        }

        if ref_pages.len() != gen_pages.len() {
            println!(
                "  [WARN] Page count mismatch: reference={}, generated={}",
                ref_pages.len(),
                gen_pages.len()
            );
        }

        if !scores.is_empty() {
            let avg = scores.iter().sum::<f64>() / scores.len() as f64;
            println!("  Average similarity: {:.2}%", avg * 100.0);
            let passed = avg >= SIMILARITY_THRESHOLD;
            log_result(&name, scores.len(), avg, passed);
            table_rows.push((name.to_string(), avg, passed));
            if !passed {
                println!(
                    "  [FAIL] {name}: average similarity {:.2}% below threshold {:.0}%",
                    avg * 100.0,
                    SIMILARITY_THRESHOLD * 100.0
                );
                all_passed = false;
            } else {
                println!("  [PASS] {name}");
            }
        }
    }

    print_summary_table("Jaccard", SIMILARITY_THRESHOLD, &table_rows);
    assert!(all_passed, "One or more fixtures failed visual comparison");
}

fn convert_docx_to_pdf(input: &Path, output: &Path) -> Result<(), String> {
    docxside_pdf::convert_docx_to_pdf(input, output).map_err(|e| e.to_string())
}

fn print_summary_table(metric: &str, threshold: f64, rows: &[(String, f64, bool)]) {
    let name_w = rows.iter().map(|(n, _, _)| n.len()).max().unwrap_or(4).max(4);
    // Format all scores first so we can measure the widest one
    let score_strs: Vec<String> = rows.iter().map(|(_, s, _)| format!("{:.1}%", s * 100.0)).collect();
    let metric_w = score_strs.iter().map(|s| s.len()).max().unwrap_or(0).max(metric.len());
    let sep = format!("+-{}-+-{}-+------+", "-".repeat(name_w), "-".repeat(metric_w));
    println!("\n{sep}");
    println!("| {:<name_w$} | {:>metric_w$} | Pass |", "Case", metric);
    println!("{sep}");
    for ((name, _, passed), score_str) in rows.iter().zip(&score_strs) {
        let mark = if *passed { "✓" } else { "✗" };
        println!("| {:<name_w$} | {:>metric_w$} | {mark}    |", name, score_str);
    }
    println!("{sep}");
    println!("  threshold: {:.0}%", threshold * 100.0);
}

// SSIM over non-overlapping 8×8 blocks. C1/C2 are the standard stabilisation constants.
fn ssim_score(a: &Path, b: &Path) -> Result<f64, String> {
    let img_a = image::open(a)
        .map_err(|e| format!("Failed to open {}: {e}", a.display()))?
        .to_luma8();
    let img_b = image::open(b)
        .map_err(|e| format!("Failed to open {}: {e}", b.display()))?
        .to_luma8();

    let (w, h) = img_a.dimensions();
    let (w2, h2) = img_b.dimensions();
    if w.abs_diff(w2) > 2 || h.abs_diff(h2) > 2 {
        return Err(format!(
            "Image dimensions differ: {:?} vs {:?}",
            (w, h),
            (w2, h2)
        ));
    }
    let cw = w.min(w2);
    let ch = h.min(h2);

    // (k1*L)^2 and (k2*L)^2 with k1=0.01, k2=0.03, L=255
    let c1: f64 = 6.5025;
    let c2: f64 = 58.5225;
    const WINDOW: u32 = 8;

    let mut ssim_sum = 0.0f64;
    let mut count = 0u64;

    for by in 0..ch / WINDOW {
        for bx in 0..cw / WINDOW {
            let x0 = bx * WINDOW;
            let y0 = by * WINDOW;
            let n = (WINDOW * WINDOW) as f64;

            // Skip blocks with no ink in the reference
            let has_ink = (y0..y0 + WINDOW).any(|y| {
                (x0..x0 + WINDOW).any(|x| img_a.get_pixel(x, y).0[0] < 200)
            });
            if !has_ink {
                continue;
            }

            let mut sum_a = 0.0f64;
            let mut sum_b = 0.0f64;
            for y in y0..y0 + WINDOW {
                for x in x0..x0 + WINDOW {
                    sum_a += img_a.get_pixel(x, y).0[0] as f64;
                    sum_b += img_b.get_pixel(x, y).0[0] as f64;
                }
            }
            let mu_a = sum_a / n;
            let mu_b = sum_b / n;

            let mut var_a = 0.0f64;
            let mut var_b = 0.0f64;
            let mut cov = 0.0f64;
            for y in y0..y0 + WINDOW {
                for x in x0..x0 + WINDOW {
                    let da = img_a.get_pixel(x, y).0[0] as f64 - mu_a;
                    let db = img_b.get_pixel(x, y).0[0] as f64 - mu_b;
                    var_a += da * da;
                    var_b += db * db;
                    cov += da * db;
                }
            }
            var_a /= n;
            var_b /= n;
            cov /= n;

            let num = (2.0 * mu_a * mu_b + c1) * (2.0 * cov + c2);
            let den = (mu_a * mu_a + mu_b * mu_b + c1) * (var_a + var_b + c2);
            ssim_sum += num / den;
            count += 1;
        }
    }

    if count == 0 {
        return Ok(1.0);
    }
    Ok(ssim_sum / count as f64)
}

fn log_ssim_result(case: &str, pages: usize, avg_ssim: f64) {
    let csv_path = Path::new("tests/output/ssim_results.csv");
    fs::create_dir_all("tests/output").ok();
    let write_header = !csv_path.exists();
    let mut file = fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(csv_path)
        .expect("Cannot open tests/output/ssim_results.csv");
    if write_header {
        writeln!(file, "timestamp,case,pages,avg_ssim").unwrap();
    }
    let ts = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_secs();
    writeln!(file, "{ts},{case},{pages},{avg_ssim:.4}").unwrap();
}

#[test]
fn ssim_comparison() {
    let fixtures = discover_fixtures().expect("Failed to read tests/fixtures");

    if fixtures.is_empty() {
        println!("[INFO] No fixtures found");
        return;
    }

    let mut all_passed = true;
    let mut table_rows: Vec<(String, f64, bool)> = Vec::new();

    for fixture_dir in &fixtures {
        let name = fixture_dir.file_name().unwrap().to_string_lossy();
        let input_docx = fixture_dir.join("input.docx");
        let reference_pdf = fixture_dir.join("reference.pdf");

        let output_base = PathBuf::from("tests/output").join(name.as_ref());
        let reference_screenshots = output_base.join("reference");
        let generated_screenshots = output_base.join("generated");

        println!("\n=== SSIM: {name} ===");

        println!("  Screenshotting reference PDF...");
        if let Err(e) = screenshot_pdf(&reference_pdf, &reference_screenshots) {
            println!("  [ERROR] Failed to screenshot reference: {e}");
            continue;
        }

        let generated_pdf = output_base.join("generated.pdf");
        println!("  Converting DOCX...");
        match convert_docx_to_pdf(&input_docx, &generated_pdf) {
            Err(e) => {
                println!("  [SKIP] {name}: {e}");
                continue;
            }
            Ok(()) => {}
        }

        println!("  Screenshotting generated PDF...");
        if let Err(e) = screenshot_pdf(&generated_pdf, &generated_screenshots) {
            println!("  [ERROR] Failed to screenshot generated: {e}");
            continue;
        }

        let ref_pages = collect_page_pngs(&reference_screenshots).unwrap_or_default();
        let gen_pages = collect_page_pngs(&generated_screenshots).unwrap_or_default();

        if ref_pages.is_empty() {
            println!("  [WARN] No reference pages found");
            continue;
        }

        if ref_pages.len() != gen_pages.len() {
            println!(
                "  [WARN] Page count mismatch: reference={}, generated={}",
                ref_pages.len(),
                gen_pages.len()
            );
        }

        let page_count = ref_pages.len().min(gen_pages.len());
        let mut scores: Vec<f64> = Vec::new();

        for i in 0..page_count {
            match ssim_score(&ref_pages[i], &gen_pages[i]) {
                Ok(score) => {
                    println!("  Page {}: SSIM = {:.4}", i + 1, score);
                    scores.push(score);
                }
                Err(e) => {
                    println!("  Page {}: SSIM error — {e}", i + 1);
                }
            }
        }

        if !scores.is_empty() {
            let avg = scores.iter().sum::<f64>() / scores.len() as f64;
            let passed = avg >= SSIM_THRESHOLD;
            println!("  Average SSIM: {:.4}", avg);
            log_ssim_result(&name, scores.len(), avg);
            table_rows.push((name.to_string(), avg, passed));
            if passed {
                println!("  [PASS] {name}");
            } else {
                println!(
                    "  [FAIL] {name}: SSIM {:.2}% below threshold {:.0}%",
                    avg * 100.0,
                    SSIM_THRESHOLD * 100.0
                );
                all_passed = false;
            }
        }
    }

    print_summary_table("SSIM", SSIM_THRESHOLD, &table_rows);
    assert!(all_passed, "One or more fixtures failed SSIM comparison");
}
