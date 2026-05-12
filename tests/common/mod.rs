#![allow(dead_code)]
use image::{DynamicImage, GenericImageView, ImageBuffer, Rgba};
use rayon::prelude::*;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};
use std::{fs, io};

pub const REGRESSION_SLACK: f64 = 0.02;

#[derive(Clone, Default, Serialize, Deserialize)]
pub struct Baselines {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub jaccard: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ssim: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_boundary: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub convert_ms: Option<f64>,
}

fn load_skiplist() -> HashSet<String> {
    let path = Path::new("tests/fixtures/SKIPLIST");
    let Ok(content) = fs::read_to_string(path) else {
        return HashSet::new();
    };
    content
        .lines()
        .map(|l| l.trim())
        .filter(|l| !l.is_empty() && !l.starts_with('#'))
        .map(|l| l.to_string())
        .collect()
}

pub fn group_name(fixture: &Path) -> String {
    fixture
        .parent()
        .and_then(|p| p.file_name())
        .and_then(|n| n.to_str())
        .unwrap_or("")
        .to_string()
}

/// Output directory: tests/output/<group>/<case>/
pub fn output_dir(fixture: &Path) -> PathBuf {
    let case = fixture.file_name().unwrap().to_string_lossy();
    PathBuf::from("tests/output")
        .join(group_name(fixture))
        .join(case.as_ref())
}

/// Display name for tables: group/case (hashes truncated to 16 chars)
pub fn display_name(fixture: &Path) -> String {
    let case = fixture.file_name().unwrap().to_string_lossy();
    let short = if case.len() > 16 {
        format!("{}..", &case[..16])
    } else {
        case.to_string()
    };
    format!("{}/{}", group_name(fixture), short)
}

fn natural_cmp(a: &Path, b: &Path) -> std::cmp::Ordering {
    let ag = group_name(a);
    let bg = group_name(b);
    let a_name = a.file_name().and_then(|n| n.to_str()).unwrap_or("");
    let b_name = b.file_name().and_then(|n| n.to_str()).unwrap_or("");
    let extract = |s: &str| -> (String, u64) {
        let i = s.find(|c: char| c.is_ascii_digit()).unwrap_or(s.len());
        (s[..i].to_string(), s[i..].parse().unwrap_or(0))
    };
    ag.cmp(&bg)
        .then_with(|| extract(a_name).cmp(&extract(b_name)))
        .then_with(|| a_name.cmp(b_name))
}

/// Discover fixtures. Filter with DOCXSIDE_CASE (case name) and DOCXSIDE_GROUP (folder name).
pub fn discover_fixtures() -> io::Result<Vec<PathBuf>> {
    let fixtures_dir = Path::new("tests/fixtures");
    let case_filter = std::env::var("DOCXIDE_CASE").ok();
    let group_filter = std::env::var("DOCXSIDE_GROUP").ok();
    let skiplist = load_skiplist();
    let mut fixtures: Vec<PathBuf> = Vec::new();
    for group_entry in fs::read_dir(fixtures_dir)? {
        let group = group_entry?.path();
        if !group.is_dir() {
            continue;
        }
        let gname = group.file_name().and_then(|n| n.to_str()).unwrap_or("");
        if let Some(ref gf) = group_filter {
            if gname != gf.as_str() {
                continue;
            }
        }
        for entry in fs::read_dir(&group)? {
            let path = entry?.path();
            if !path.is_dir() {
                continue;
            }
            let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("");
            if let Some(ref filter) = case_filter {
                if name == filter.as_str() {
                    fixtures.push(path);
                }
            } else if !skiplist.contains(name) && !skiplist.contains(gname) {
                fixtures.push(path);
            }
        }
    }
    fixtures.sort_by(|a, b| natural_cmp(a, b));
    Ok(fixtures)
}

pub fn timestamp() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_secs()
}

pub fn log_csv(csv_name: &str, header: &str, row: &str) {
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

pub fn read_baselines() -> HashMap<String, Baselines> {
    let path = Path::new("tests/baselines.json");
    let Ok(content) = fs::read_to_string(path) else {
        return HashMap::new();
    };
    serde_json::from_str(&content).unwrap_or_default()
}

fn round4(v: f64) -> f64 {
    (v * 10000.0).round() / 10000.0
}

/// Write scores from the current test run to tests/output/latest_scores.json.
/// Merges per-field so different tests (visual, text_boundary, speed) can each
/// contribute their metrics without overwriting each other.
/// This file is gitignored — use `accept-baselines` to promote into baselines.json.
pub fn write_latest_scores(updates: &HashMap<String, Baselines>) {
    fs::create_dir_all("tests/output").ok();
    let path = Path::new("tests/output/latest_scores.json");
    let mut scores: BTreeMap<String, Baselines> = fs::read_to_string(path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default();
    for (name, new) in updates {
        let entry = scores.entry(name.clone()).or_default();
        if let Some(v) = new.jaccard {
            entry.jaccard = Some(round4(v));
        }
        if let Some(v) = new.ssim {
            entry.ssim = Some(round4(v));
        }
        if let Some(v) = new.text_boundary {
            entry.text_boundary = Some(round4(v));
        }
        if let Some(v) = new.convert_ms {
            entry.convert_ms = Some(round4(v));
        }
    }
    let json = serde_json::to_string_pretty(&scores).expect("Failed to serialize latest scores");
    fs::write(path, json + "\n").expect("Failed to write latest_scores.json");
}

/// Returns the newest mtime of any file found by recursively walking `dir`.
fn dir_newest_mtime(dir: &Path) -> std::time::SystemTime {
    let mut newest = std::time::SystemTime::UNIX_EPOCH;
    let Ok(entries) = fs::read_dir(dir) else {
        return newest;
    };
    for entry in entries.filter_map(|e| e.ok()) {
        let path = entry.path();
        if path.is_dir() {
            let sub = dir_newest_mtime(&path);
            if sub > newest {
                newest = sub;
            }
        } else if let Ok(mtime) = fs::metadata(&path).and_then(|m| m.modified()) {
            if mtime > newest {
                newest = mtime;
            }
        }
    }
    newest
}

/// The newest mtime of any file under `src/`, cached for the process lifetime.
fn src_newest_mtime() -> std::time::SystemTime {
    static SRC_MTIME: std::sync::OnceLock<std::time::SystemTime> = std::sync::OnceLock::new();
    *SRC_MTIME.get_or_init(|| dir_newest_mtime(Path::new("src")))
}

/// Convert DOCX→PDF only if the generated PDF is missing or older than input.docx or src/.
pub fn ensure_generated_pdf(fixture_dir: &Path) -> Result<PathBuf, String> {
    let input_docx = fixture_dir.join("input.docx");
    let out = output_dir(fixture_dir);
    fs::create_dir_all(&out).map_err(|e| e.to_string())?;
    let generated_pdf = out.join("generated.pdf");

    let needs_convert = !generated_pdf.exists() || {
        let docx_mtime = fs::metadata(&input_docx)
            .and_then(|m| m.modified())
            .unwrap_or(std::time::SystemTime::UNIX_EPOCH);
        let src_mtime = src_newest_mtime();
        let newest_input = docx_mtime.max(src_mtime);
        let pdf_mtime = fs::metadata(&generated_pdf)
            .and_then(|m| m.modified())
            .unwrap_or(std::time::SystemTime::UNIX_EPOCH);
        pdf_mtime < newest_input
    };

    if needs_convert {
        let input = input_docx.clone();
        let output = generated_pdf.clone();
        let result =
            std::panic::catch_unwind(move || docxide_pdf::convert_docx_to_pdf(&input, &output));
        match result {
            Ok(Ok(())) => {}
            Ok(Err(e)) => return Err(e.to_string()),
            Err(_) => return Err("conversion panicked".to_string()),
        }
    }

    Ok(generated_pdf)
}

pub fn delta_str(current: f64, previous: Option<f64>) -> String {
    match previous {
        Some(prev) => {
            let diff = (current - prev) * 100.0;
            if diff.abs() < 0.05 {
                String::new()
            } else if diff > 0.0 {
                format!(" (+{diff:.1}pp)")
            } else {
                format!(" ({diff:.1}pp)")
            }
        }
        None => String::new(),
    }
}

/// SHA-256 hash of decoded RGBA pixel data for each generated page PNG.
/// Hashing pixel data (not file bytes) ensures stability across mutool versions.
pub fn compute_page_hashes(gen_pages: &[PathBuf]) -> Vec<String> {
    gen_pages
        .iter()
        .filter_map(|p| {
            let img = image::open(p).ok()?.to_rgba8();
            let mut hasher = Sha256::new();
            hasher.update(img.as_raw());
            Some(format!("{:x}", hasher.finalize()))
        })
        .collect()
}

/// Write per-case page hashes to tests/output/latest_hashes.json.
/// Simple overwrite (only visual_comparison writes hashes).
pub fn write_latest_hashes(hashes: &BTreeMap<String, Vec<String>>) {
    fs::create_dir_all("tests/output").ok();
    let path = Path::new("tests/output/latest_hashes.json");
    let json = serde_json::to_string_pretty(hashes).expect("Failed to serialize latest hashes");
    fs::write(path, json + "\n").expect("Failed to write latest_hashes.json");
}

// ----------------------------------------------------------------------------
// PDF rasterization + similarity metrics
// ----------------------------------------------------------------------------

pub const MUTOOL_DPI: &str = "150";

pub fn pdf_page_count(pdf: &Path) -> Result<usize, String> {
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

/// True if `screenshot_dir` contains PNGs all newer than `pdf` — i.e. the
/// cached rasterization is still valid for the current PDF.
pub fn pngs_fresh(pdf: &Path, screenshot_dir: &Path) -> bool {
    let Ok(pdf_meta) = fs::metadata(pdf) else {
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

pub fn is_ink_luma(r: u8, g: u8, b: u8) -> bool {
    (r as u32 * 299 + g as u32 * 587 + b as u32 * 114) < 200_000
}

pub struct PageResult {
    pub jaccard: f64,
    pub diff_img: ImageBuffer<Rgba<u8>, Vec<u8>>,
}

/// Jaccard similarity on ink pixels (luma < ~200). Also produces a color-coded
/// diff image: gray=both, blue=ref-only, red=gen-only, white=neither.
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

// ----------------------------------------------------------------------------
// LibreOffice headless conversion
// ----------------------------------------------------------------------------

/// Find the LibreOffice headless binary. Order:
/// 1. $LIBREOFFICE_PATH if set and executable
/// 2. /Applications/LibreOffice.app/Contents/MacOS/soffice (macOS default)
/// 3. `soffice` on $PATH
pub fn find_libreoffice() -> Option<PathBuf> {
    if let Ok(p) = std::env::var("LIBREOFFICE_PATH") {
        let path = PathBuf::from(p);
        if path.is_file() {
            return Some(path);
        }
    }
    let mac_default = PathBuf::from("/Applications/LibreOffice.app/Contents/MacOS/soffice");
    if mac_default.is_file() {
        return Some(mac_default);
    }
    if Command::new("soffice")
        .arg("--version")
        .stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::null())
        .status()
        .map_or(false, |s| s.success())
    {
        return Some(PathBuf::from("soffice"));
    }
    None
}

/// Convert fixture's input.docx via LibreOffice headless.
/// Caches at tests/output/<group>/<case>/libreoffice.pdf. Per-case user profile
/// dir avoids LibreOffice's single-instance lock so calls can run in parallel.
pub fn ensure_libreoffice_pdf(fixture_dir: &Path, soffice: &Path) -> Result<PathBuf, String> {
    let input_docx = fixture_dir.join("input.docx");
    let out = output_dir(fixture_dir);
    fs::create_dir_all(&out).map_err(|e| e.to_string())?;
    let lo_pdf = out.join("libreoffice.pdf");

    let needs_convert = !lo_pdf.exists() || {
        let docx_mtime = fs::metadata(&input_docx)
            .and_then(|m| m.modified())
            .unwrap_or(std::time::SystemTime::UNIX_EPOCH);
        let pdf_mtime = fs::metadata(&lo_pdf)
            .and_then(|m| m.modified())
            .unwrap_or(std::time::SystemTime::UNIX_EPOCH);
        pdf_mtime < docx_mtime
    };
    if !needs_convert {
        return Ok(lo_pdf);
    }

    let profile = out.join("lo_profile");
    fs::create_dir_all(&profile).map_err(|e| e.to_string())?;
    let abs_profile = fs::canonicalize(&profile).map_err(|e| e.to_string())?;
    let profile_uri = format!("file://{}", abs_profile.display());

    let output = Command::new(soffice)
        .arg(format!("-env:UserInstallation={profile_uri}"))
        .arg("--headless")
        .arg("--convert-to")
        .arg("pdf")
        .arg("--outdir")
        .arg(&out)
        .arg(&input_docx)
        .output()
        .map_err(|e| format!("Failed to run soffice: {e}"))?;
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(format!(
            "soffice exited with status {}: {}",
            output.status,
            stderr.trim()
        ));
    }

    let lo_default = out.join("input.pdf");
    if !lo_default.exists() {
        return Err(format!(
            "soffice succeeded but {} was not produced",
            lo_default.display()
        ));
    }
    fs::rename(&lo_default, &lo_pdf).map_err(|e| format!("rename: {e}"))?;
    Ok(lo_pdf)
}
