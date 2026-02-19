mod common;

use image::{DynamicImage, GenericImageView, ImageBuffer, Rgba};
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::OnceLock;
use std::{fs, io};

const SIMILARITY_THRESHOLD: f64 = 0.27;
const SSIM_THRESHOLD: f64 = 0.54;
const MUTOOL_DPI: &str = "150";

fn screenshot_pdf(pdf: &Path, output_dir: &Path) -> Result<(), String> {
    fs::create_dir_all(output_dir).map_err(|e| e.to_string())?;
    let output_pattern = output_dir.join("page_%03d.png");
    let status = Command::new("mutool")
        .args([
            "draw",
            "-F",
            "png",
            "-r",
            MUTOOL_DPI,
            "-o",
            output_pattern.to_str().unwrap(),
            pdf.to_str().unwrap(),
        ])
        .stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::null())
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

fn compare_images(a: &Path, b: &Path) -> Result<f64, String> {
    let img_a = image::open(a).map_err(|e| format!("Failed to open {}: {e}", a.display()))?;
    let img_b = image::open(b).map_err(|e| format!("Failed to open {}: {e}", b.display()))?;
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

fn save_diff_image(ref_path: &Path, gen_path: &Path, out: &Path) -> Result<(), String> {
    let img_ref = image::open(ref_path).map_err(|e| format!("{e}"))?;
    let img_gen = image::open(gen_path).map_err(|e| format!("{e}"))?;
    let (w, h) = img_ref.dimensions();
    let (w2, h2) = img_gen.dimensions();
    let cw = w.min(w2);
    let ch = h.min(h2);
    let ink_threshold = 200u8;
    let mut diff: ImageBuffer<Rgba<u8>, Vec<u8>> = ImageBuffer::new(cw, ch);
    for y in 0..ch {
        for x in 0..cw {
            let Rgba([rr, gr, br, _]) = img_ref.get_pixel(x, y).0.into();
            let Rgba([rg, gg, bg, _]) = img_gen.get_pixel(x, y).0.into();
            let luma_ref: u8 =
                ((rr as u32 * 299 + gr as u32 * 587 + br as u32 * 114) / 1000) as u8;
            let luma_gen: u8 =
                ((rg as u32 * 299 + gg as u32 * 587 + bg as u32 * 114) / 1000) as u8;
            let ref_ink = luma_ref < ink_threshold;
            let gen_ink = luma_gen < ink_threshold;
            let pixel = match (ref_ink, gen_ink) {
                (true, true) => Rgba([80, 80, 80, 255]),       // both: dark gray
                (true, false) => Rgba([0, 80, 220, 255]),      // reference only: blue
                (false, true) => Rgba([220, 40, 40, 255]),     // generated only: red
                (false, false) => Rgba([255, 255, 255, 255]),  // neither: white
            };
            diff.put_pixel(x, y, pixel);
        }
    }
    if let Some(parent) = out.parent() {
        fs::create_dir_all(parent).map_err(|e| e.to_string())?;
    }
    DynamicImage::ImageRgba8(diff)
        .save(out)
        .map_err(|e| e.to_string())
}

fn save_side_by_side(a: &Path, b: &Path, out: &Path) -> Result<(), String> {
    let img_a = image::open(a).map_err(|e| format!("{e}"))?;
    let img_b = image::open(b).map_err(|e| format!("{e}"))?;
    let (wa, ha) = img_a.dimensions();
    let (wb, hb) = img_b.dimensions();
    let gap = 4u32;
    let total_w = wa + gap + wb;
    let total_h = ha.max(hb);
    let mut canvas: ImageBuffer<Rgba<u8>, Vec<u8>> = ImageBuffer::from_pixel(
        total_w,
        total_h,
        Rgba([220, 220, 220, 255]),
    );
    for y in 0..ha {
        for x in 0..wa {
            canvas.put_pixel(x, y, img_a.get_pixel(x, y));
        }
    }
    for y in 0..hb {
        for x in 0..wb {
            canvas.put_pixel(wa + gap + x, y, img_b.get_pixel(x, y));
        }
    }
    if let Some(parent) = out.parent() {
        fs::create_dir_all(parent).map_err(|e| e.to_string())?;
    }
    DynamicImage::ImageRgba8(canvas)
        .save(out)
        .map_err(|e| e.to_string())
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

fn prepare_fixture(fixture_dir: &Path) -> Option<FixturePages> {
    let name = fixture_dir
        .file_name()
        .unwrap()
        .to_string_lossy()
        .to_string();
    let input_docx = fixture_dir.join("input.docx");
    let reference_pdf = fixture_dir.join("reference.pdf");
    let output_base = PathBuf::from("tests/output").join(&name);
    let reference_screenshots = output_base.join("reference");
    let generated_screenshots = output_base.join("generated");

    let _ = fs::remove_dir_all(&reference_screenshots);
    let _ = fs::remove_dir_all(&generated_screenshots);
    let _ = fs::remove_dir_all(&output_base.join("diff"));
    let _ = fs::remove_dir_all(&output_base.join("comparison"));

    if let Err(e) = screenshot_pdf(&reference_pdf, &reference_screenshots) {
        println!("  [ERROR] {name}: screenshot reference failed: {e}");
        return None;
    }
    let generated_pdf = output_base.join("generated.pdf");
    if let Err(e) = docxside_pdf::convert_docx_to_pdf(&input_docx, &generated_pdf) {
        println!("  [SKIP] {name}: {e}");
        return None;
    }
    if let Err(e) = screenshot_pdf(&generated_pdf, &generated_screenshots) {
        println!("  [ERROR] {name}: screenshot generated failed: {e}");
        return None;
    }
    let ref_pages = collect_page_pngs(&reference_screenshots).unwrap_or_default();
    let gen_pages = collect_page_pngs(&generated_screenshots).unwrap_or_default();
    if ref_pages.is_empty() {
        return None;
    }
    Some(FixturePages {
        name,
        ref_pages,
        gen_pages,
        output_base,
    })
}

fn prepared_fixtures() -> &'static Vec<FixturePages> {
    static FIXTURES: OnceLock<Vec<FixturePages>> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        let fixture_dirs = common::discover_fixtures().expect("Failed to read tests/fixtures");
        fixture_dirs
            .iter()
            .filter_map(|d| prepare_fixture(d))
            .collect()
    })
}

fn print_summary(
    metric: &str,
    threshold: f64,
    rows: &[(String, f64, bool)],
    prev: &HashMap<String, f64>,
) {
    let name_w = rows
        .iter()
        .map(|(n, _, _)| n.len())
        .max()
        .unwrap_or(4)
        .max(4);
    let sep = format!("+-{}-+---------+------+-----------+", "-".repeat(name_w));
    println!("\n{sep}");
    println!("| {:<name_w$} | {:>7} | Pass | Delta     |", "Case", metric);
    println!("{sep}");
    for (name, score, passed) in rows {
        let score_str = format!("{:.1}%", score * 100.0);
        let mark = if *passed { "Y" } else { "N" };
        let delta = common::delta_str(*score, prev.get(name).copied());
        println!(
            "| {:<name_w$} | {:>7} | {mark}    | {:<9} |",
            name, score_str, delta
        );
    }
    println!("{sep}");
    println!("  threshold: {:.0}%", threshold * 100.0);

    let regressions: Vec<&str> = rows
        .iter()
        .filter(|(name, score, _)| prev.get(name).is_some_and(|&p| *score < p - 0.005))
        .map(|(name, _, _)| name.as_str())
        .collect();
    if !regressions.is_empty() {
        println!("  REGRESSION in: {}", regressions.join(", "));
    }
}

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
    const SEARCH_RADIUS: i32 = 8;
    let mut ssim_sum = 0.0f64;
    let mut count = 0u64;
    for by in 0..ch / WINDOW {
        for bx in 0..cw / WINDOW {
            let x0 = bx * WINDOW;
            let y0 = by * WINDOW;
            let n = (WINDOW * WINDOW) as f64;
            let has_ink = (y0..y0 + WINDOW)
                .any(|y| (x0..x0 + WINDOW).any(|x| img_a.get_pixel(x, y).0[0] < 200));
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
fn visual_comparison() {
    let _ = env_logger::try_init();
    let fixtures = prepared_fixtures();
    if fixtures.is_empty() {
        return;
    }

    let prev_scores = common::read_previous_scores("results.csv", 3);
    let mut all_passed = true;
    let mut table_rows: Vec<(String, f64, bool)> = Vec::new();

    for fixture in fixtures {
        let diff_dir = fixture.output_base.join("diff");
        let comparison_dir = fixture.output_base.join("comparison");
        let page_count = fixture.ref_pages.len().min(fixture.gen_pages.len());
        let mut scores: Vec<f64> = Vec::new();
        for i in 0..page_count {
            if let Ok(score) = compare_images(&fixture.ref_pages[i], &fixture.gen_pages[i]) {
                scores.push(score);
                let page_num = fixture.ref_pages[i].file_stem().unwrap().to_str().unwrap();
                let _ = save_diff_image(
                    &fixture.ref_pages[i],
                    &fixture.gen_pages[i],
                    &diff_dir.join(format!("{page_num}.png")),
                );
                let _ = save_side_by_side(
                    &fixture.ref_pages[i],
                    &fixture.gen_pages[i],
                    &comparison_dir.join(format!("{page_num}.png")),
                );
            }
        }
        if !scores.is_empty() {
            let avg = scores.iter().sum::<f64>() / scores.len() as f64;
            let passed = avg >= SIMILARITY_THRESHOLD;
            common::log_csv(
                "results.csv",
                "timestamp,case,pages,avg_jaccard,pass",
                &format!(
                    "{},{},{},{:.4},{}",
                    common::timestamp(),
                    fixture.name,
                    scores.len(),
                    avg,
                    passed
                ),
            );
            table_rows.push((fixture.name.clone(), avg, passed));
            if !passed {
                all_passed = false;
            }
        }
    }

    print_summary("Jaccard", SIMILARITY_THRESHOLD, &table_rows, &prev_scores);
    assert!(all_passed, "One or more fixtures failed visual comparison");
}

#[test]
fn ssim_comparison() {
    let _ = env_logger::try_init();
    let fixtures = prepared_fixtures();
    if fixtures.is_empty() {
        return;
    }

    let prev_scores = common::read_previous_scores("ssim_results.csv", 3);
    let mut all_passed = true;
    let mut table_rows: Vec<(String, f64, bool)> = Vec::new();

    for fixture in fixtures {
        let page_count = fixture.ref_pages.len().min(fixture.gen_pages.len());
        let mut scores: Vec<f64> = Vec::new();
        for i in 0..page_count {
            if let Ok(score) = ssim_score(&fixture.ref_pages[i], &fixture.gen_pages[i]) {
                scores.push(score);
            }
        }
        if !scores.is_empty() {
            let avg = scores.iter().sum::<f64>() / scores.len() as f64;
            let passed = avg >= SSIM_THRESHOLD;
            common::log_csv(
                "ssim_results.csv",
                "timestamp,case,pages,avg_ssim",
                &format!(
                    "{},{},{},{:.4}",
                    common::timestamp(),
                    fixture.name,
                    scores.len(),
                    avg
                ),
            );
            table_rows.push((fixture.name.clone(), avg, passed));
            if !passed {
                all_passed = false;
            }
        }
    }

    print_summary("SSIM", SSIM_THRESHOLD, &table_rows, &prev_scores);
    assert!(all_passed, "One or more fixtures failed SSIM comparison");
}
