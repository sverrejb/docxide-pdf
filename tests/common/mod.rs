#![allow(dead_code)]
use rayon::prelude::*;
use serde::de::DeserializeOwned;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};
use std::{fs, io};

pub mod a11y;
pub mod metrics;
pub mod text_boundary;
pub use metrics::*;

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
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ref_pages: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gen_pages: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ua_fail: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub a11y_missing: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ua_deficit: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub a11y_struct: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub a11y_text: Option<f64>,
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

/// group/case: the key for baselines, hashes and score tables. The full case
/// name, since cases sharing a prefix must not share a key.
pub fn case_key(fixture: &Path) -> String {
    let case = fixture.file_name().unwrap().to_string_lossy();
    format!("{}/{}", group_name(fixture), case)
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

/// Discover fixtures. Filter with DOCXIDE_CASE (case name) and DOCXIDE_GROUP (folder name).
pub fn discover_fixtures() -> io::Result<Vec<PathBuf>> {
    let fixtures_dir = Path::new("tests/fixtures");
    let case_filter = std::env::var("DOCXIDE_CASE").ok();
    let group_filter = std::env::var("DOCXIDE_GROUP").ok();
    let skiplist = load_skiplist();
    let mut fixtures: Vec<PathBuf> = Vec::new();
    for group_entry in fs::read_dir(fixtures_dir)? {
        let group = group_entry?.path();
        if !group.is_dir() {
            continue;
        }
        let gname = group.file_name().and_then(|n| n.to_str()).unwrap_or("");
        if let Some(ref gf) = group_filter
            && gname != gf.as_str()
        {
            continue;
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
        if let Some(v) = new.ref_pages {
            entry.ref_pages = Some(v);
        }
        if let Some(v) = new.gen_pages {
            entry.gen_pages = Some(v);
        }
        if let Some(v) = new.ua_fail {
            entry.ua_fail = Some(v);
        }
        if let Some(v) = new.a11y_missing {
            entry.a11y_missing = Some(v);
        }
        if let Some(v) = new.ua_deficit {
            entry.ua_deficit = Some(v);
        }
        if let Some(v) = new.a11y_struct {
            entry.a11y_struct = Some(round4(v));
        }
        if let Some(v) = new.a11y_text {
            entry.a11y_text = Some(round4(v));
        }
    }
    let json = serde_json::to_string_pretty(&scores).expect("Failed to serialize latest scores");
    fs::write(path, json + "\n").expect("Failed to write latest_scores.json");
}

/// The per-case metric cache: `cache` is reused while it is newer than every input, since the
/// metrics of an unchanged PDF pair cannot change and an unchanged PDF keeps its mtime (see
/// `ensure_generated_pdf`). Otherwise `compute` runs and its result is written back.
// ponytail: like the a11y cache this ignores the metric code's own version; delete
// tests/output/*/*/<cache name> after changing a metric
pub fn cached<T: Serialize + DeserializeOwned>(
    cache: &Path,
    inputs: &[&Path],
    compute: impl FnOnce() -> Option<T>,
) -> Option<T> {
    let fresh =
        mtime(cache).is_some_and(|c| inputs.iter().all(|i| mtime(i).is_some_and(|m| m <= c)));
    if fresh
        && let Some(v) = fs::read_to_string(cache)
            .ok()
            .and_then(|s| serde_json::from_str(&s).ok())
    {
        return Some(v);
    }
    let v = compute()?;
    if let Ok(json) = serde_json::to_string(&v) {
        let _ = fs::write(cache, json);
    }
    Some(v)
}

/// Returns the newest mtime of any file found by recursively walking `dir`.
fn dir_newest_mtime(dir: &Path) -> SystemTime {
    let mut newest = UNIX_EPOCH;
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
        } else if let Some(mtime) = mtime(&path)
            && mtime > newest
        {
            newest = mtime;
        }
    }
    newest
}

/// The newest mtime of any file under `src/`, cached for the process lifetime.
fn src_newest_mtime() -> SystemTime {
    static SRC_MTIME: std::sync::OnceLock<SystemTime> = std::sync::OnceLock::new();
    *SRC_MTIME.get_or_init(|| dir_newest_mtime(Path::new("src")))
}

/// Convert `input` to `output`, turning a conversion panic into an error.
pub fn convert(input: &Path, output: &Path) -> Result<(), String> {
    match std::panic::catch_unwind(|| docxide_pdf::convert_docx_to_pdf(input, output)) {
        Ok(r) => r.map_err(|e| e.to_string()),
        Err(_) => Err("conversion panicked".to_string()),
    }
}

/// Convert DOCX→PDF only if the generated PDF is missing or older than input.docx or src/.
///
/// Conversion is deterministic, so when the new PDF is byte-identical to the existing one the
/// existing file is kept, mtime and all: the screenshots, diffs and accessibility analyses
/// derived from it stay fresh, and a src/ edit only re-renders the fixtures it changed. A stamp
/// beside the PDF records which inputs it was last converted against, so an unchanged PDF is
/// not reconverted on every run either.
pub fn ensure_generated_pdf(fixture_dir: &Path) -> Result<PathBuf, String> {
    let input_docx = fixture_dir.join("input.docx");
    let out = output_dir(fixture_dir);
    fs::create_dir_all(&out).map_err(|e| e.to_string())?;
    let generated_pdf = out.join("generated.pdf");
    let stamp = out.join("generated.stamp");

    let mtime_or_epoch = |p: &Path| mtime(p).unwrap_or(UNIX_EPOCH);
    let newest_input = mtime_or_epoch(&input_docx).max(src_newest_mtime());
    let stamp_value = format!(
        "{:?}",
        newest_input.duration_since(UNIX_EPOCH).unwrap_or_default()
    );
    let up_to_date = generated_pdf.exists()
        && (mtime_or_epoch(&generated_pdf) >= newest_input
            || fs::read_to_string(&stamp).is_ok_and(|s| s == stamp_value));

    if !up_to_date {
        let tmp = out.join("generated.tmp.pdf");
        if let Err(e) = convert(&input_docx, &tmp) {
            let _ = fs::remove_file(&tmp);
            return Err(e);
        }
        let new_bytes = fs::read(&tmp).map_err(|e| e.to_string())?;
        let unchanged = fs::read(&generated_pdf).is_ok_and(|old| old == new_bytes);
        if unchanged {
            let _ = fs::remove_file(&tmp);
        } else {
            fs::rename(&tmp, &generated_pdf).map_err(|e| e.to_string())?;
        }
        let _ = fs::write(&stamp, stamp_value);
    }

    Ok(generated_pdf)
}

/// `ensure_generated_pdf` for a test that prints a `[SKIP]` row and drops the fixture
/// when the conversion fails.
pub fn generated_pdf_or_skip(fixture_dir: &Path, name: &str) -> Option<PathBuf> {
    match ensure_generated_pdf(fixture_dir) {
        Ok(p) => Some(p),
        Err(e) => {
            println!("  [SKIP] {name}: {e}");
            None
        }
    }
}

/// Column width for a table of case names: the longest name, at least `min`.
pub fn name_width<'a>(names: impl IntoIterator<Item = &'a str>, min: usize) -> usize {
    names.into_iter().map(str::len).fold(min, usize::max)
}

/// ANSI color gradient from red (0%) to green (100%).
pub fn color_score(score: f64, text: &str) -> String {
    let t = score.clamp(0.0, 1.0);
    let r = (220.0 * (1.0 - t) + 80.0 * t) as u8;
    let g = (40.0 * (1.0 - t) + 200.0 * t) as u8;
    let b = (40.0 * (1.0 - t) + 80.0 * t) as u8;
    format!("\x1b[38;2;{r};{g};{b}m{text}\x1b[0m")
}

/// The cases whose score fell more than `REGRESSION_SLACK` below their baseline.
pub fn regressions<'a>(
    scores: impl IntoIterator<Item = (&'a str, f64)>,
    prev: &HashMap<String, f64>,
) -> Vec<&'a str> {
    scores
        .into_iter()
        .filter(|(name, score)| {
            prev.get(*name)
                .is_some_and(|&p| *score < p - REGRESSION_SLACK)
        })
        .map(|(name, _)| name)
        .collect()
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
        .par_iter()
        .filter_map(|p| {
            let img = image::open(p).ok()?.to_rgba8();
            let mut hasher = Sha256::new();
            hasher.update(img.as_raw());
            Some(format!("{:x}", hasher.finalize()))
        })
        .collect()
}

/// Write per-case page hashes to tests/output/latest_hashes.json.
/// Merges by case so a single-case run (the case browser's re-generate button)
/// updates that case without dropping every other case from the change-detection
/// set. A full run still replaces every entry; run-tests.sh rm's the file first.
pub fn write_latest_hashes(hashes: &BTreeMap<String, Vec<String>>) {
    fs::create_dir_all("tests/output").ok();
    let path = Path::new("tests/output/latest_hashes.json");
    let mut merged: BTreeMap<String, Vec<String>> = fs::read_to_string(path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default();
    merged.extend(hashes.iter().map(|(k, v)| (k.clone(), v.clone())));
    let json = serde_json::to_string_pretty(&merged).expect("Failed to serialize latest hashes");
    fs::write(path, json + "\n").expect("Failed to write latest_hashes.json");
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
        .is_ok_and(|s| s.success())
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

    let needs_convert = !lo_pdf.exists()
        || mtime(&lo_pdf).unwrap_or(UNIX_EPOCH) < mtime(&input_docx).unwrap_or(UNIX_EPOCH);
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
