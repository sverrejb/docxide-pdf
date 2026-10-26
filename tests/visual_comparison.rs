mod common;

use common::{compare_and_diff, ssim_score};
use image::{DynamicImage, GenericImageView, ImageBuffer, Rgba};
use rayon::prelude::*;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::OnceLock;

const SIMILARITY_THRESHOLD: f64 = 0.205;
const SSIM_THRESHOLD: f64 = 0.75;

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

struct FixturePages {
    name: String,
    ref_pages: Vec<PathBuf>,
    gen_pages: Vec<PathBuf>,
    output_base: PathBuf,
    diffs_fresh: bool,
}

fn prepare_fixture(fixture_dir: &Path) -> Option<FixturePages> {
    let name = common::case_key(fixture_dir);
    let reference_pdf = fixture_dir.join("reference.pdf");
    let output_base = common::output_dir(fixture_dir);
    let reference_screenshots = output_base.join("reference");
    let generated_screenshots = output_base.join("generated");

    if save_side_by_side_images() {
        let _ = fs::remove_dir_all(output_base.join("comparison"));
    }

    let ref_fresh = match common::ensure_screenshots(&reference_pdf, &reference_screenshots) {
        Ok(fresh) => fresh,
        Err(e) => {
            println!("  [ERROR] {name}: screenshot reference failed: {e}");
            return None;
        }
    };
    let generated_pdf = common::generated_pdf_or_skip(fixture_dir, &name)?;
    // An unchanged generated.pdf keeps its screenshots (see ensure_generated_pdf).
    let gen_fresh = match common::ensure_screenshots(&generated_pdf, &generated_screenshots) {
        Ok(fresh) => fresh,
        Err(e) => {
            println!("  [ERROR] {name}: screenshot generated failed: {e}");
            return None;
        }
    };
    let ref_pages = common::collect_page_pngs(&reference_screenshots).unwrap_or_default();
    let gen_pages = common::collect_page_pngs(&generated_screenshots).unwrap_or_default();
    if ref_pages.is_empty() {
        return None;
    }
    // The diff images derive from both sets of pages; keep them while neither set was re-rendered.
    let diff_dir = output_base.join("diff");
    let diffs_fresh = ref_fresh
        && gen_fresh
        && common::collect_page_pngs(&diff_dir)
            .is_ok_and(|d| d.len() == ref_pages.len().min(gen_pages.len()));
    if !diffs_fresh {
        let _ = fs::remove_dir_all(&diff_dir);
    }
    Some(FixturePages {
        name,
        ref_pages,
        gen_pages,
        output_base,
        diffs_fresh,
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

fn print_summary(
    metric: &str,
    threshold: f64,
    rows: &[(String, f64, bool)],
    prev: &HashMap<String, f64>,
) {
    let name_w = common::name_width(rows.iter().map(|(n, _, _)| n.as_str()), 4);
    println!("\n  {:<name_w$}  {:>7}  Pass  Delta", "Case", metric);
    for (name, score, passed) in rows {
        let score_str = format!("{:.1}%", score * 100.0);
        let colored_score = common::color_score(*score, &format!("{:>7}", score_str));
        let mark = if *passed { "Y" } else { "N" };
        let delta = common::delta_str(*score, prev.get(name).copied());
        println!(
            "  {:<name_w$}  {}  {mark}     {:<9}",
            name, colored_score, delta
        );
    }
    println!("  threshold: {:.0}%", threshold * 100.0);

    let regressions = common::regressions(rows.iter().map(|(n, s, _)| (n.as_str(), *s)), prev);
    if !regressions.is_empty() {
        println!("  REGRESSION in: {}", regressions.join(", "));
    }
    assert!(
        regressions.is_empty(),
        "{metric} regression in: {}",
        regressions.join(", ")
    );
}

#[derive(Serialize, Deserialize)]
struct FixtureResult {
    name: String,
    jaccard: f64,
    ssim: f64,
    ref_pages: usize,
    gen_pages: usize,
    /// SHA-256 of each generated page's pixels, for visual change detection.
    hashes: Vec<String>,
}

fn save_side_by_side_images() -> bool {
    std::env::var("DOCXIDE_IMAGES").is_ok()
}

/// Scores from `visual.json` while the page PNGs are unchanged; a missing diff/ or a
/// side-by-side request needs the images written, so those score afresh.
fn score_fixture(fixture: &FixturePages) -> Option<FixtureResult> {
    let save_comparison = save_side_by_side_images();
    let cache = fixture.output_base.join("visual.json");
    if save_comparison || !fixture.diffs_fresh {
        let _ = fs::remove_file(&cache);
    }
    let inputs: Vec<&Path> = fixture
        .ref_pages
        .iter()
        .chain(&fixture.gen_pages)
        .map(PathBuf::as_path)
        .collect();
    common::cached(&cache, &inputs, || compute_scores(fixture, save_comparison))
}

fn compute_scores(fixture: &FixturePages, save_comparison: bool) -> Option<FixtureResult> {
    let diff_dir = fixture.output_base.join("diff");
    let comparison_dir = fixture.output_base.join("comparison");
    let _ = fs::create_dir_all(&diff_dir);
    if save_comparison {
        let _ = fs::create_dir_all(&comparison_dir);
    }
    let page_count = fixture.ref_pages.len().min(fixture.gen_pages.len());

    // (jaccard, ssim) per page
    let page_scores: Vec<(f64, f64)> = (0..page_count)
        .into_par_iter()
        .filter_map(|i| {
            let img_ref = image::open(&fixture.ref_pages[i]).ok()?;
            let img_gen = image::open(&fixture.gen_pages[i]).ok()?;
            let page_num = fixture.ref_pages[i].file_stem()?.to_str()?.to_string();

            let result = compare_and_diff(&img_ref, &img_gen).ok()?;
            if !fixture.diffs_fresh {
                let _ = DynamicImage::ImageRgba8(result.diff_img)
                    .save(diff_dir.join(format!("{page_num}.png")));
            }
            if save_comparison {
                let _ = save_side_by_side(
                    &img_ref,
                    &img_gen,
                    &comparison_dir.join(format!("{page_num}.png")),
                );
            }
            Some((result.jaccard, ssim_score(&img_ref, &img_gen).ok()?))
        })
        .collect();

    if page_scores.is_empty() {
        return None;
    }
    let n = page_scores.len() as f64;
    Some(FixtureResult {
        name: fixture.name.clone(),
        jaccard: page_scores.iter().map(|s| s.0).sum::<f64>() / n,
        ssim: page_scores.iter().map(|s| s.1).sum::<f64>() / n,
        ref_pages: fixture.ref_pages.len(),
        gen_pages: fixture.gen_pages.len(),
        hashes: common::compute_page_hashes(&fixture.gen_pages),
    })
}

#[test]
fn visual_comparison() {
    let _ = env_logger::try_init();
    let fixtures = prepared_fixtures();
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

    let mut results: Vec<FixtureResult> = fixtures.par_iter().filter_map(score_fixture).collect();
    results.sort_by(|a, b| a.name.cmp(&b.name));

    let mut baseline_updates: HashMap<String, common::Baselines> = HashMap::new();
    for r in &results {
        baseline_updates.insert(
            r.name.clone(),
            common::Baselines {
                jaccard: Some(r.jaccard),
                ssim: Some(r.ssim),
                ref_pages: Some(r.ref_pages),
                gen_pages: Some(r.gen_pages),
                ..Default::default()
            },
        );
    }
    common::write_latest_scores(&baseline_updates);

    // Compute SHA-256 hashes of generated page PNGs for visual change detection
    let mut hash_updates: std::collections::BTreeMap<String, Vec<String>> =
        std::collections::BTreeMap::new();
    for r in &results {
        if !r.hashes.is_empty() {
            hash_updates.insert(r.name.clone(), r.hashes.clone());
        }
    }
    common::write_latest_hashes(&hash_updates);

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
}
