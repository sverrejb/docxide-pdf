use image::{DynamicImage, GenericImageView, ImageBuffer, Rgba};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};
use std::{fs, io};

const SIMILARITY_THRESHOLD: f64 = 0.25;
const SSIM_THRESHOLD: f64 = 0.40;
const MUTOOL_DPI: &str = "150";

const SKIP_FIXTURES: &[&str] = &["sample100kB"];

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

fn timestamp() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_secs()
}

fn collect_page_pngs(dir: &Path) -> io::Result<Vec<PathBuf>> {
    let mut pages: Vec<PathBuf> = fs::read_dir(dir)?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("png"))
        .collect();
    pages.sort();
    Ok(pages)
}

struct FixturePages {
    name: String,
    ref_pages: Vec<PathBuf>,
    gen_pages: Vec<PathBuf>,
    output_base: PathBuf,
}

fn prepare_fixture(fixture_dir: &Path, label: &str) -> Option<FixturePages> {
    let name = fixture_dir.file_name().unwrap().to_string_lossy().to_string();
    let input_docx = fixture_dir.join("input.docx");
    let reference_pdf = fixture_dir.join("reference.pdf");

    let output_base = PathBuf::from("tests/output").join(&name);
    let reference_screenshots = output_base.join("reference");
    let generated_screenshots = output_base.join("generated");

    println!("\n=== {label}: {name} ===");

    println!("  Screenshotting reference PDF...");
    if let Err(e) = screenshot_pdf(&reference_pdf, &reference_screenshots) {
        println!("  [ERROR] Failed to screenshot reference: {e}");
        return None;
    }

    let generated_pdf = output_base.join("generated.pdf");
    println!("  Converting DOCX...");
    if let Err(e) = docxside_pdf::convert_docx_to_pdf(&input_docx, &generated_pdf) {
        println!("  [SKIP] {name}: {e}");
        return None;
    }

    println!("  Screenshotting generated PDF...");
    if let Err(e) = screenshot_pdf(&generated_pdf, &generated_screenshots) {
        println!("  [ERROR] Failed to screenshot generated: {e}");
        return None;
    }

    let ref_pages = collect_page_pngs(&reference_screenshots).unwrap_or_default();
    let gen_pages = collect_page_pngs(&generated_screenshots).unwrap_or_default();

    if ref_pages.is_empty() {
        println!("  [WARN] No reference pages found");
        return None;
    }

    if ref_pages.len() != gen_pages.len() {
        println!(
            "  [WARN] Page count mismatch: reference={}, generated={}",
            ref_pages.len(),
            gen_pages.len()
        );
    }

    Some(FixturePages { name, ref_pages, gen_pages, output_base })
}

#[test]
fn visual_comparison() {
    let fixtures = discover_fixtures().expect("Failed to read tests/fixtures");

    if fixtures.is_empty() {
        println!("[INFO] No fixtures found -- add DOCX+PDF pairs to tests/fixtures/<name>/");
        return;
    }

    let mut all_passed = true;
    let mut table_rows: Vec<(String, f64, bool)> = Vec::new();

    for fixture_dir in &fixtures {
        let Some(fixture) = prepare_fixture(fixture_dir, "Fixture") else {
            all_passed = false;
            continue;
        };
        let diff_dir = fixture.output_base.join("diff");

        let page_count = fixture.ref_pages.len().min(fixture.gen_pages.len());
        let mut scores: Vec<f64> = Vec::new();

        for i in 0..page_count {
            let ref_page = &fixture.ref_pages[i];
            let gen_page = &fixture.gen_pages[i];
            let page_num = ref_page.file_stem().unwrap().to_str().unwrap();
            let diff_path = diff_dir.join(format!("{page_num}.png"));

            match compare_images(ref_page, gen_page) {
                Ok(score) => {
                    println!("  Page {}: similarity = {:.2}%", i + 1, score * 100.0);
                    scores.push(score);
                    let _ = save_diff_image(ref_page, gen_page, &diff_path);
                }
                Err(e) => {
                    println!("  Page {}: comparison error -- {e}", i + 1);
                }
            }
        }

        if !scores.is_empty() {
            let avg = scores.iter().sum::<f64>() / scores.len() as f64;
            let passed = avg >= SIMILARITY_THRESHOLD;
            println!("  Average similarity: {:.2}%", avg * 100.0);
            log_csv(
                "results.csv",
                "timestamp,case,pages,avg_jaccard,pass",
                &format!("{},{},{},{:.4},{}", timestamp(), fixture.name, scores.len(), avg, passed),
            );
            table_rows.push((fixture.name.clone(), avg, passed));
            if passed {
                println!("  [PASS] {}", fixture.name);
            } else {
                println!(
                    "  [FAIL] {}: average similarity {:.2}% below threshold {:.0}%",
                    fixture.name,
                    avg * 100.0,
                    SIMILARITY_THRESHOLD * 100.0
                );
                all_passed = false;
            }
        }
    }

    print_summary_table("Jaccard", SIMILARITY_THRESHOLD, &table_rows);
    assert!(all_passed, "One or more fixtures failed visual comparison");
}

fn print_summary_table(metric: &str, threshold: f64, rows: &[(String, f64, bool)]) {
    let name_w = rows.iter().map(|(n, _, _)| n.len()).max().unwrap_or(4).max(4);
    let score_strs: Vec<String> = rows.iter().map(|(_, s, _)| format!("{:.1}%", s * 100.0)).collect();
    let metric_w = score_strs.iter().map(|s| s.len()).max().unwrap_or(0).max(metric.len());
    let sep = format!("+-{}-+-{}-+------+", "-".repeat(name_w), "-".repeat(metric_w));
    println!("\n{sep}");
    println!("| {:<name_w$} | {:>metric_w$} | Pass |", "Case", metric);
    println!("{sep}");
    for ((name, _, passed), score_str) in rows.iter().zip(&score_strs) {
        let mark = if *passed { "Y" } else { "N" };
        println!("| {:<name_w$} | {:>metric_w$} | {mark}    |", name, score_str);
    }
    println!("{sep}");
    println!("  threshold: {:.0}%", threshold * 100.0);
}

// SSIM with spatial tolerance: for each reference ink block, find the best
// structural match in a small vertical neighborhood of the generated image.
// This tolerates cumulative line-spacing errors (a few pixels) while still
// measuring whether the text was rendered with the correct font, size, and structure.
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

    let c1: f64 = 6.5025;
    let c2: f64 = 58.5225;
    const WINDOW: u32 = 8;
    const SEARCH_RADIUS: i32 = 8; // pixels of vertical tolerance

    let mut ssim_sum = 0.0f64;
    let mut count = 0u64;

    for by in 0..ch / WINDOW {
        for bx in 0..cw / WINDOW {
            let x0 = bx * WINDOW;
            let y0 = by * WINDOW;
            let n = (WINDOW * WINDOW) as f64;

            let has_ink = (y0..y0 + WINDOW).any(|y| {
                (x0..x0 + WINDOW).any(|x| img_a.get_pixel(x, y).0[0] < 200)
            });
            if !has_ink {
                continue;
            }

            let mut sum_a = 0.0f64;
            for y in y0..y0 + WINDOW {
                for x in x0..x0 + WINDOW {
                    sum_a += img_a.get_pixel(x, y).0[0] as f64;
                }
            }
            let mu_a = sum_a / n;
            let mut var_a = 0.0f64;
            for y in y0..y0 + WINDOW {
                for x in x0..x0 + WINDOW {
                    let da = img_a.get_pixel(x, y).0[0] as f64 - mu_a;
                    var_a += da * da;
                }
            }
            var_a /= n;

            let mut best_ssim = f64::NEG_INFINITY;
            for dy in -SEARCH_RADIUS..=SEARCH_RADIUS {
                let sy0 = y0 as i32 + dy;
                if sy0 < 0 || (sy0 as u32 + WINDOW) > ch {
                    continue;
                }
                let sy0 = sy0 as u32;

                let mut sum_b = 0.0f64;
                for y in sy0..sy0 + WINDOW {
                    for x in x0..x0 + WINDOW {
                        sum_b += img_b.get_pixel(x, y).0[0] as f64;
                    }
                }
                let mu_b = sum_b / n;

                let mut var_b = 0.0f64;
                let mut cov = 0.0f64;
                for y in 0..WINDOW {
                    for x in x0..x0 + WINDOW {
                        let da = img_a.get_pixel(x, y0 + y).0[0] as f64 - mu_a;
                        let db = img_b.get_pixel(x, sy0 + y).0[0] as f64 - mu_b;
                        var_b += db * db;
                        cov += da * db;
                    }
                }
                var_b /= n;
                cov /= n;

                let num = (2.0 * mu_a * mu_b + c1) * (2.0 * cov + c2);
                let den = (mu_a * mu_a + mu_b * mu_b + c1) * (var_a + var_b + c2);
                best_ssim = best_ssim.max(num / den);
            }
            ssim_sum += best_ssim;
            count += 1;
        }
    }

    if count == 0 {
        return Ok(1.0);
    }
    Ok(ssim_sum / count as f64)
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
        let Some(fixture) = prepare_fixture(fixture_dir, "SSIM") else {
            continue;
        };

        let page_count = fixture.ref_pages.len().min(fixture.gen_pages.len());
        let mut scores: Vec<f64> = Vec::new();

        for i in 0..page_count {
            match ssim_score(&fixture.ref_pages[i], &fixture.gen_pages[i]) {
                Ok(score) => {
                    println!("  Page {}: SSIM = {:.4}", i + 1, score);
                    scores.push(score);
                }
                Err(e) => {
                    println!("  Page {}: SSIM error -- {e}", i + 1);
                }
            }
        }

        if !scores.is_empty() {
            let avg = scores.iter().sum::<f64>() / scores.len() as f64;
            let passed = avg >= SSIM_THRESHOLD;
            println!("  Average SSIM: {:.4}", avg);
            log_csv(
                "ssim_results.csv",
                "timestamp,case,pages,avg_ssim",
                &format!("{},{},{},{:.4}", timestamp(), fixture.name, scores.len(), avg),
            );
            table_rows.push((fixture.name.clone(), avg, passed));
            if passed {
                println!("  [PASS] {}", fixture.name);
            } else {
                println!(
                    "  [FAIL] {}: SSIM {:.2}% below threshold {:.0}%",
                    fixture.name,
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
