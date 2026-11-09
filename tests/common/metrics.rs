//! PDF rasterization and the similarity metrics (Jaccard, SSIM), shared with tools/page-metrics,
//! which includes this file by path. Nothing here may depend on the library or on mod.rs: that
//! would make page-metrics link docxide-pdf and rebuild after every src/ change.
use image::{DynamicImage, GenericImageView, ImageBuffer, Rgba};
use rayon::prelude::*;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::SystemTime;
use std::{fs, io};

pub fn mtime(p: &Path) -> Option<SystemTime> {
    fs::metadata(p).and_then(|m| m.modified()).ok()
}

pub const MUTOOL_DPI: &str = "150";

/// The output of `mutool info` for a PDF.
pub fn mutool_info(pdf: &Path) -> Result<String, String> {
    let output = Command::new("mutool")
        .args(["info", pdf.to_str().unwrap()])
        .output()
        .map_err(|e| format!("Failed to run mutool info: {e}"))?;
    Ok(String::from_utf8_lossy(&output.stdout).into_owned())
}

pub fn pdf_page_count(pdf: &Path) -> Result<usize, String> {
    let text = mutool_info(pdf)?;
    for line in text.lines() {
        if let Some(rest) = line.strip_prefix("Pages:")
            && let Ok(n) = rest.trim().parse::<usize>()
        {
            return Ok(n);
        }
    }
    Err("Could not determine page count".to_string())
}

pub fn screenshot_pdf(pdf: &Path, output_dir: &Path) -> Result<(), String> {
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

pub fn collect_page_pngs(dir: &Path) -> io::Result<Vec<PathBuf>> {
    let mut pages: Vec<PathBuf> = fs::read_dir(dir)?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("png"))
        .collect();
    pages.sort();
    Ok(pages)
}

/// Rasterize `pdf` into `dir` unless the PNGs there are newer than it. Returns
/// whether the cached pages were kept.
pub fn ensure_screenshots(pdf: &Path, dir: &Path) -> Result<bool, String> {
    if pngs_fresh(pdf, dir) {
        return Ok(true);
    }
    let _ = fs::remove_dir_all(dir);
    screenshot_pdf(pdf, dir).map(|()| false)
}

/// True if `screenshot_dir` contains PNGs all newer than `pdf` — i.e. the
/// cached rasterization is still valid for the current PDF.
pub fn pngs_fresh(pdf: &Path, screenshot_dir: &Path) -> bool {
    let Some(pdf_mtime) = mtime(pdf) else {
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
            .is_ok_and(|t| t >= pdf_mtime)
    })
}

pub fn is_ink_luma(r: u8, g: u8, b: u8) -> bool {
    (r as u32 * 299 + g as u32 * 587 + b as u32 * 114) < 200_000
}

/// Diff image palette: gray=both, blue=ref-only, red=gen-only, white=neither.
pub fn diff_pixel(ref_ink: bool, gen_ink: bool) -> [u8; 4] {
    match (ref_ink, gen_ink) {
        (true, true) => [80, 80, 80, 255],
        (true, false) => [0, 80, 220, 255],
        (false, true) => [220, 40, 40, 255],
        (false, false) => [255, 255, 255, 255],
    }
}

pub struct PageResult {
    pub jaccard: f64,
    pub diff_img: ImageBuffer<Rgba<u8>, Vec<u8>>,
}

/// Jaccard similarity on ink pixels (luma < ~200). Also produces a color-coded
/// diff image (`diff_pixel`).
pub fn compare_and_diff(
    img_ref: &DynamicImage,
    img_gen: &DynamicImage,
) -> Result<PageResult, String> {
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
            diff_row[ri..ri + 4].copy_from_slice(&diff_pixel(ref_ink, gen_ink));
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

/// SSIM with 8×8 windows and ±8px vertical search (compensates for small
/// vertical baseline drift between renderers). Skips white windows.
pub fn ssim_score(img_a_dyn: &DynamicImage, img_b_dyn: &DynamicImage) -> Result<f64, String> {
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
    const WN: usize = WINDOW as usize;
    const SEARCH_RADIUS: i32 = 8;
    let n = (WINDOW * WINDOW) as f64;

    let raw_a = img_a.as_raw();
    let raw_b = img_b.as_raw();
    let stride_a = w as usize;
    let stride_b = w2 as usize;

    let mut ssim_sum = 0.0f64;
    let mut count = 0u64;
    for by in 0..ch / WINDOW {
        for bx in 0..cw / WINDOW {
            let x0 = (bx * WINDOW) as usize;
            let y0 = (by * WINDOW) as usize;

            let mut has_ink = false;
            let mut sum_a = 0.0f64;
            let mut win_a = [0.0f64; WN * WN];
            for wy in 0..WN {
                let row_off = (y0 + wy) * stride_a + x0;
                for wx in 0..WN {
                    let v = raw_a[row_off + wx] as f64;
                    win_a[wy * WN + wx] = v;
                    sum_a += v;
                    if !has_ink && v < 200.0 {
                        has_ink = true;
                    }
                }
            }
            if !has_ink {
                continue;
            }

            let mu_a = sum_a / n;
            let mut var_a = 0.0f64;
            for &v in &win_a {
                let da = v - mu_a;
                var_a += da * da;
            }
            var_a /= n;

            let mut best_ssim = f64::NEG_INFINITY;
            for dy in -SEARCH_RADIUS..=SEARCH_RADIUS {
                let sy0 = y0 as i32 + dy;
                if sy0 < 0 || (sy0 as u32 + WINDOW) > ch {
                    continue;
                }
                let sy0 = sy0 as usize;

                let mut sum_b = 0.0f64;
                for wy in 0..WN {
                    let row_off = (sy0 + wy) * stride_b + x0;
                    for wx in 0..WN {
                        sum_b += raw_b[row_off + wx] as f64;
                    }
                }
                let mu_b = sum_b / n;

                let mut var_b = 0.0f64;
                let mut cov = 0.0f64;
                for wy in 0..WN {
                    let row_off = (sy0 + wy) * stride_b + x0;
                    for wx in 0..WN {
                        let da = win_a[wy * WN + wx] - mu_a;
                        let db = raw_b[row_off + wx] as f64 - mu_b;
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
