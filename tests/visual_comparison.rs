mod common;

use image::{DynamicImage, GenericImageView, ImageBuffer, Rgba};
use rayon::prelude::*;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::OnceLock;
use std::time::Instant;
use std::{fs, io};

const SIMILARITY_THRESHOLD: f64 = 0.205;
const SSIM_THRESHOLD: f64 = 0.75;
const MUTOOL_DPI: &str = "150";

fn pdf_page_count(pdf: &Path) -> Result<usize, String> {
    let output = Command::new("mutool")
        .args(["info", pdf.to_str().unwrap()])
        .output()
        .map_err(|e| format!("Failed to run mutool info: {e}"))?;
    let text = String::from_utf8_lossy(&output.stdout);
    for line in text.lines() {
        if let Some(rest) = line.strip_prefix("Pages:") {
            if let Ok(n) = rest.trim().parse::<usize>() {
                return Ok(n);
            }
        }
    }
    Err("Could not determine page count".to_string())
}

fn screenshot_pdf(pdf: &Path, output_dir: &Path) -> Result<(), String> {
    fs::create_dir_all(output_dir).map_err(|e| e.to_string())?;
    let n = pdf_page_count(pdf)?;
    let errors: Vec<String> = (1..=n)
        .into_par_iter()
        .filter_map(|page| {
            let out_file = output_dir.join(format!("page_{:03}.png", page));
            let status = Command::new("mutool")
                .args([
                    "draw",
                    "-F",
                    "png",
                    "-r",
                    MUTOOL_DPI,
                    "-o",
                    out_file.to_str().unwrap(),
                    pdf.to_str().unwrap(),
                    &page.to_string(),
                ])
                .stdout(std::process::Stdio::null())
                .stderr(std::process::Stdio::null())
                .status();
            match status {
                Ok(s) if s.success() => None,
                Ok(s) => Some(format!("page {page}: exit {}", s.code().unwrap_or(-1))),
                Err(e) => Some(format!("page {page}: {e}")),
            }
        })
        .collect();
    if errors.is_empty() {
        Ok(())
    } else {
        Err(errors.join("; "))
    }
}

fn is_ink_luma(r: u8, g: u8, b: u8) -> bool {
    (r as u32 * 299 + g as u32 * 587 + b as u32 * 114) < 200_000
}

struct PageResult {
    jaccard: f64,
    diff_img: ImageBuffer<Rgba<u8>, Vec<u8>>,
}

fn compare_and_diff(img_ref: &DynamicImage, img_gen: &DynamicImage) -> Result<PageResult, String> {
    let (w, h) = img_ref.dimensions();
    let (w2, h2) = img_gen.dimensions();
    if w.abs_diff(w2) > 2 || h.abs_diff(h2) > 2 {
        return Err(format!(
            "Image dimensions differ: {:?} vs {:?}",
            (w, h),
            (w2, h2)
        ));
    }
    let cw = w.min(w2);
    let ch = h.min(h2);
    let ref_rgba = img_ref.to_rgba8();
    let gen_rgba = img_gen.to_rgba8();
    let ref_buf = ref_rgba.as_raw();
    let gen_buf = gen_rgba.as_raw();
    let stride_ref = (w * 4) as usize;
    let stride_gen = (w2 * 4) as usize;

    let mut intersection: u64 = 0;
    let mut union: u64 = 0;
    let mut diff_buf: Vec<u8> = vec![255; (cw * ch * 4) as usize];

    for y in 0..ch as usize {
        let ref_row = &ref_buf[y * stride_ref..];
        let gen_row = &gen_buf[y * stride_gen..];
        let diff_row = &mut diff_buf[y * (cw as usize * 4)..];
        for x in 0..cw as usize {
            let ri = x * 4;
            let (rr, gr, br) = (ref_row[ri], ref_row[ri + 1], ref_row[ri + 2]);
            let (rg, gg, bg) = (gen_row[ri], gen_row[ri + 1], gen_row[ri + 2]);
            let ref_ink = is_ink_luma(rr, gr, br);
            let gen_ink = is_ink_luma(rg, gg, bg);
            if ref_ink || gen_ink {
                union += 1;
            }
            if ref_ink && gen_ink {
                intersection += 1;
            }
            let pixel = match (ref_ink, gen_ink) {
                (true, true) => [80, 80, 80, 255],
                (true, false) => [0, 80, 220, 255],
                (false, true) => [220, 40, 40, 255],
                (false, false) => [255, 255, 255, 255],
            };
            diff_row[ri..ri + 4].copy_from_slice(&pixel);
        }
    }

    let jaccard = if union == 0 {
        1.0
    } else {
        intersection as f64 / union as f64
    };
    let diff_img = ImageBuffer::from_raw(cw, ch, diff_buf)
        .ok_or_else(|| "failed to create diff image".to_string())?;
    Ok(PageResult { jaccard, diff_img })
}

fn save_side_by_side(img_a: &DynamicImage, img_b: &DynamicImage, out: &Path) -> Result<(), String> {
    let (wa, ha) = img_a.dimensions();
    let (wb, hb) = img_b.dimensions();
    let buf_a = img_a.to_rgba8();
    let buf_b = img_b.to_rgba8();
    let gap = 4u32;
    let total_w = wa + gap + wb;
    let total_h = ha.max(hb);
    let bpp = 4usize;
    let row_bytes = total_w as usize * bpp;
    let mut canvas = vec![220u8; total_h as usize * row_bytes];
    // fill alpha channel to 255
    for px in canvas.chunks_exact_mut(4) {
        px[3] = 255;
    }
    let stride_a = wa as usize * bpp;
    let stride_b = wb as usize * bpp;
    let a_raw = buf_a.as_raw();
    let b_raw = buf_b.as_raw();
    for y in 0..ha as usize {
        let dst_offset = y * row_bytes;
        let src_offset = y * stride_a;
        canvas[dst_offset..dst_offset + stride_a]
            .copy_from_slice(&a_raw[src_offset..src_offset + stride_a]);
    }
    let x_offset = (wa + gap) as usize * bpp;
    for y in 0..hb as usize {
        let dst_offset = y * row_bytes + x_offset;
        let src_offset = y * stride_b;
        canvas[dst_offset..dst_offset + stride_b]
            .copy_from_slice(&b_raw[src_offset..src_offset + stride_b]);
    }
    if let Some(parent) = out.parent() {
        fs::create_dir_all(parent).map_err(|e| e.to_string())?;
    }
    let img: ImageBuffer<Rgba<u8>, _> =
        ImageBuffer::from_raw(total_w, total_h, canvas).ok_or("canvas alloc")?;
    DynamicImage::ImageRgba8(img)
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
    convert_ms: u64,
    screenshot_ms: u64,
}

fn ref_screenshots_fresh(reference_pdf: &Path, screenshot_dir: &Path) -> bool {
    let Ok(pdf_meta) = fs::metadata(reference_pdf) else {
        return false;
    };
    let Ok(pdf_mtime) = pdf_meta.modified() else {
        return false;
    };
    let Ok(entries) = fs::read_dir(screenshot_dir) else {
        return false;
    };
    let pngs: Vec<_> = entries
        .filter_map(|e| e.ok())
        .filter(|e| e.path().extension().and_then(|x| x.to_str()) == Some("png"))
        .collect();
    if pngs.is_empty() {
        return false;
    }
    pngs.iter().all(|e| {
        e.metadata()
            .and_then(|m| m.modified())
            .map_or(false, |t| t >= pdf_mtime)
    })
}

fn prepare_fixture(fixture_dir: &Path) -> Option<FixturePages> {
    let name = common::display_name(fixture_dir);
    let reference_pdf = fixture_dir.join("reference.pdf");
    let output_base = common::output_dir(fixture_dir);
    let reference_screenshots = output_base.join("reference");
    let generated_screenshots = output_base.join("generated");

    let _ = fs::remove_dir_all(&generated_screenshots);
    let _ = fs::remove_dir_all(&output_base.join("diff"));
    if save_side_by_side_images() {
        let _ = fs::remove_dir_all(&output_base.join("comparison"));
    }

    if !ref_screenshots_fresh(&reference_pdf, &reference_screenshots) {
        let _ = fs::remove_dir_all(&reference_screenshots);
        if let Err(e) = screenshot_pdf(&reference_pdf, &reference_screenshots) {
            println!("  [ERROR] {name}: screenshot reference failed: {e}");
            return None;
        }
    }
    let t0 = Instant::now();
    let generated_pdf = match common::ensure_generated_pdf(fixture_dir) {
        Ok(p) => p,
        Err(e) => {
            println!("  [SKIP] {name}: {e}");
            return None;
        }
    };
    let convert_ms = t0.elapsed().as_millis() as u64;
    let t1 = Instant::now();
    if let Err(e) = screenshot_pdf(&generated_pdf, &generated_screenshots) {
        println!("  [ERROR] {name}: screenshot generated failed: {e}");
        return None;
    }
    let screenshot_ms = t1.elapsed().as_millis() as u64;
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
        convert_ms,
        screenshot_ms,
    })
}

fn prepared_fixtures() -> &'static Vec<FixturePages> {
    static FIXTURES: OnceLock<Vec<FixturePages>> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        let fixture_dirs = common::discover_fixtures().expect("Failed to read tests/fixtures");
        fixture_dirs
            .par_iter()
            .filter_map(|d| prepare_fixture(d))
            .collect()
    })
}

/// ANSI color gradient from red (0%) to green (100%).
fn color_score(score: f64, text: &str) -> String {
    let t = score.clamp(0.0, 1.0);
    let r = (220.0 * (1.0 - t) + 80.0 * t) as u8;
    let g = (40.0 * (1.0 - t) + 200.0 * t) as u8;
    let b = (40.0 * (1.0 - t) + 80.0 * t) as u8;
    format!("\x1b[38;2;{r};{g};{b}m{text}\x1b[0m")
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
    println!("\n  {:<name_w$}  {:>7}  Pass  Delta", "Case", metric);
    for (name, score, passed) in rows {
        let score_str = format!("{:.1}%", score * 100.0);
        let colored_score = color_score(*score, &format!("{:>7}", score_str));
        let mark = if *passed { "Y" } else { "N" };
        let delta = common::delta_str(*score, prev.get(name).copied());
        println!(
            "  {:<name_w$}  {}  {mark}     {:<9}",
            name, colored_score, delta
        );
    }
    println!("  threshold: {:.0}%", threshold * 100.0);

    let regressions: Vec<&str> = rows
        .iter()
        .filter(|(name, score, _)| {
            prev.get(name)
                .is_some_and(|&p| *score < p - common::REGRESSION_SLACK)
        })
        .map(|(name, _, _)| name.as_str())
        .collect();
    if !regressions.is_empty() {
        println!("  REGRESSION in: {}", regressions.join(", "));
    }
    assert!(
        regressions.is_empty(),
        "{metric} regression in: {}",
        regressions.join(", ")
    );
}

fn ssim_score(img_a_dyn: &DynamicImage, img_b_dyn: &DynamicImage) -> Result<f64, String> {
    let img_a = img_a_dyn.to_luma8();
    let img_b = img_b_dyn.to_luma8();
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

struct FixtureResult {
    name: String,
    jaccard: f64,
    ssim: f64,
    page_count: usize,
    jaccard_ms: u64,
    ssim_ms: u64,
    diff_save_ms: u64,
}

fn save_side_by_side_images() -> bool {
    std::env::var("DOCXIDE_IMAGES").is_ok()
}

struct PageTiming {
    jaccard: f64,
    ssim: f64,
    jaccard_ms: u64,
    ssim_ms: u64,
    diff_save_ms: u64,
}

fn score_fixture(fixture: &FixturePages) -> Option<FixtureResult> {
    let save_comparison = save_side_by_side_images();
    let diff_dir = fixture.output_base.join("diff");
    let comparison_dir = fixture.output_base.join("comparison");
    let _ = fs::create_dir_all(&diff_dir);
    if save_comparison {
        let _ = fs::create_dir_all(&comparison_dir);
    }
    let page_count = fixture.ref_pages.len().min(fixture.gen_pages.len());

    let page_timings: Vec<PageTiming> = (0..page_count)
        .into_par_iter()
        .filter_map(|i| {
            let img_ref = image::open(&fixture.ref_pages[i]).ok()?;
            let img_gen = image::open(&fixture.gen_pages[i]).ok()?;
            let page_num = fixture.ref_pages[i].file_stem()?.to_str()?.to_string();

            let t0 = Instant::now();
            let result = compare_and_diff(&img_ref, &img_gen).ok()?;
            let jaccard = result.jaccard;
            let jaccard_ms = t0.elapsed().as_millis() as u64;

            let t1 = Instant::now();
            let _ = DynamicImage::ImageRgba8(result.diff_img)
                .save(diff_dir.join(format!("{page_num}.png")));
            if save_comparison {
                let _ = save_side_by_side(
                    &img_ref,
                    &img_gen,
                    &comparison_dir.join(format!("{page_num}.png")),
                );
            }
            let diff_save_ms = t1.elapsed().as_millis() as u64;

            let t2 = Instant::now();
            let ssim = ssim_score(&img_ref, &img_gen).ok()?;
            let ssim_ms = t2.elapsed().as_millis() as u64;

            Some(PageTiming {
                jaccard,
                ssim,
                jaccard_ms,
                ssim_ms,
                diff_save_ms,
            })
        })
        .collect();

    if page_timings.is_empty() {
        return None;
    }
    let n = page_timings.len() as f64;
    let avg_jaccard = page_timings.iter().map(|t| t.jaccard).sum::<f64>() / n;
    let avg_ssim = page_timings.iter().map(|t| t.ssim).sum::<f64>() / n;
    let jaccard_ms = page_timings.iter().map(|t| t.jaccard_ms).sum();
    let ssim_ms = page_timings.iter().map(|t| t.ssim_ms).sum();
    let diff_save_ms = page_timings.iter().map(|t| t.diff_save_ms).sum();
    Some(FixtureResult {
        name: fixture.name.clone(),
        jaccard: avg_jaccard,
        ssim: avg_ssim,
        page_count: page_timings.len(),
        jaccard_ms,
        ssim_ms,
        diff_save_ms,
    })
}

#[test]
fn visual_comparison() {
    let _ = env_logger::try_init();
    let t_prep = Instant::now();
    let fixtures = prepared_fixtures();
    let _prep_ms = t_prep.elapsed().as_millis() as u64;
    if fixtures.is_empty() {
        return;
    }

    if !save_side_by_side_images() {
        let has_stale = fixtures
            .iter()
            .any(|f| f.output_base.join("comparison").exists());
        if has_stale {
            println!(
                "  [STALE] comparison/ dirs exist from a previous run (set DOCXIDE_IMAGES=1 to regenerate)"
            );
        }
    }

    let baselines = common::read_baselines();
    let prev_jaccard: HashMap<String, f64> = baselines
        .iter()
        .filter_map(|(k, v)| v.jaccard.map(|j| (k.clone(), j)))
        .collect();
    let prev_ssim: HashMap<String, f64> = baselines
        .iter()
        .filter_map(|(k, v)| v.ssim.map(|s| (k.clone(), s)))
        .collect();

    let t_score = Instant::now();
    let mut results: Vec<FixtureResult> = fixtures
        .par_iter()
        .filter_map(|fixture| score_fixture(fixture))
        .collect();
    let _score_ms = t_score.elapsed().as_millis() as u64;
    results.sort_by(|a, b| a.name.cmp(&b.name));

    for r in &results {
        let jaccard_pass = r.jaccard >= SIMILARITY_THRESHOLD;
        common::log_csv(
            "results.csv",
            "timestamp,case,pages,avg_jaccard,pass",
            &format!(
                "{},{},{},{:.4},{}",
                common::timestamp(),
                r.name,
                r.page_count,
                r.jaccard,
                jaccard_pass
            ),
        );
        common::log_csv(
            "ssim_results.csv",
            "timestamp,case,pages,avg_ssim",
            &format!(
                "{},{},{},{:.4}",
                common::timestamp(),
                r.name,
                r.page_count,
                r.ssim
            ),
        );
    }

    let mut baseline_updates: HashMap<String, common::Baselines> = HashMap::new();
    for r in &results {
        baseline_updates.insert(
            r.name.clone(),
            common::Baselines {
                jaccard: Some(r.jaccard),
                ssim: Some(r.ssim),
                text_boundary: None,
                convert_ms: None,
            },
        );
    }
    common::update_baselines(&baseline_updates);

    let jaccard_rows: Vec<(String, f64, bool)> = results
        .iter()
        .map(|r| (r.name.clone(), r.jaccard, r.jaccard >= SIMILARITY_THRESHOLD))
        .collect();
    print_summary(
        "Jaccard",
        SIMILARITY_THRESHOLD,
        &jaccard_rows,
        &prev_jaccard,
    );

    let ssim_rows: Vec<(String, f64, bool)> = results
        .iter()
        .map(|r| (r.name.clone(), r.ssim, r.ssim >= SSIM_THRESHOLD))
        .collect();
    print_summary("SSIM", SSIM_THRESHOLD, &ssim_rows, &prev_ssim);

    // // Timing breakdown
    // let name_w = results
    //     .iter()
    //     .map(|r| r.name.len())
    //     .max()
    //     .unwrap_or(4)
    //     .max(4);
    // let total_convert: u64 = fixtures.iter().map(|f| f.convert_ms).sum();
    // let total_screenshot: u64 = fixtures.iter().map(|f| f.screenshot_ms).sum();
    // let total_jaccard: u64 = results.iter().map(|r| r.jaccard_ms).sum();
    // let total_ssim: u64 = results.iter().map(|r| r.ssim_ms).sum();
    // let total_diff_save: u64 = results.iter().map(|r| r.diff_save_ms).sum();
    //
    // println!("\n  Timing (wall: prep {:.1}s, score {:.1}s)", prep_ms as f64 / 1000.0, score_ms as f64 / 1000.0);
    // println!(
    //     "  {:<name_w$}  {:>6}  {:>6}  {:>6}  {:>6}  {:>6}",
    //     "Case", "Conv", "Scrn", "Jacc", "Diff", "SSIM"
    // );
    // for fixture in fixtures.iter() {
    //     let r = results.iter().find(|r| r.name == fixture.name);
    //     let (jms, dms, sms) = r.map_or((0, 0, 0), |r| (r.jaccard_ms, r.diff_save_ms, r.ssim_ms));
    //     println!(
    //         "  {:<name_w$}  {:>5}ms {:>5}ms {:>5}ms {:>5}ms {:>5}ms",
    //         fixture.name, fixture.convert_ms, fixture.screenshot_ms, jms, dms, sms
    //     );
    // }
    // println!(
    //     "  {:<name_w$}  {:>5}ms {:>5}ms {:>5}ms {:>5}ms {:>5}ms  (cpu totals)",
    //     "TOTAL", total_convert, total_screenshot, total_jaccard, total_diff_save, total_ssim
    // );
}

#[test]
fn ssim_comparison() {
    // Merged into visual_comparison — this test is kept for backwards compatibility
    // with `cargo test ssim` filtering. It shares prepared_fixtures() via OnceLock
    // so no duplicate work if visual_comparison already ran.
    let _ = env_logger::try_init();
    let fixtures = prepared_fixtures();
    if fixtures.is_empty() {
        return;
    }
}
