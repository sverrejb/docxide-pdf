//! Fixture discovery under tests/output, plus the committed scores and page hashes.

use serde::Deserialize;
use std::collections::{HashMap, HashSet};
use std::path::{Path, PathBuf};
use std::time::SystemTime;

pub fn find_output_dir() -> PathBuf {
    // Try relative to CWD first, then walk up
    let candidates = [
        PathBuf::from("tests/output"),
        PathBuf::from("../tests/output"),
    ];
    for c in &candidates {
        if c.is_dir() {
            return c.clone();
        }
    }
    eprintln!("Could not find tests/output directory. Run from the project root.");
    std::process::exit(1);
}

#[derive(Clone)]
pub struct CaseInfo {
    pub name: String,
    pub dir: PathBuf,
    pub page_count: usize,
}

pub fn discover_cases(output_dir: &Path) -> Vec<CaseInfo> {
    let mut cases = Vec::new();

    // Scan all subdirs (cases/, scraped/, fonts/, new/, etc.)
    let Ok(top_entries) = std::fs::read_dir(output_dir) else {
        return cases;
    };
    for top_entry in top_entries.flatten() {
        let subdir_path = top_entry.path();
        if !subdir_path.is_dir() {
            continue;
        }
        let subdir = top_entry.file_name().to_string_lossy().into_owned();
        let Ok(entries) = std::fs::read_dir(&subdir_path) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if !path.is_dir() {
                continue;
            }
            let ref_dir = path.join("reference");
            let gen_dir = path.join("generated");
            if !ref_dir.is_dir() && !gen_dir.is_dir() {
                continue;
            }
            let page_count = count_pages(&path);
            if page_count == 0 {
                continue;
            }
            let name = if subdir == "cases" {
                entry.file_name().to_string_lossy().into_owned()
            } else {
                format!("{}/{}", subdir, entry.file_name().to_string_lossy())
            };
            cases.push(CaseInfo {
                name,
                dir: path,
                page_count,
            });
        }
    }

    cases.sort_by(|a, b| natural_sort_key(&a.name).cmp(&natural_sort_key(&b.name)));
    cases
}

pub fn natural_sort_key(s: &str) -> Vec<NatPart> {
    let mut parts = Vec::new();
    let mut chars = s.chars().peekable();
    while chars.peek().is_some() {
        if chars.peek().unwrap().is_ascii_digit() {
            let mut num = String::new();
            while chars.peek().is_some() && chars.peek().unwrap().is_ascii_digit() {
                num.push(chars.next().unwrap());
            }
            parts.push(NatPart::Num(num.parse().unwrap_or(0)));
        } else {
            let mut text = String::new();
            while chars.peek().is_some() && !chars.peek().unwrap().is_ascii_digit() {
                text.push(chars.next().unwrap());
            }
            parts.push(NatPart::Str(text));
        }
    }
    parts
}

#[derive(Clone, PartialEq, Eq, PartialOrd, Ord)]
pub enum NatPart {
    Str(String),
    Num(u64),
}

// --- Baseline scores ---

#[derive(Default, Deserialize)]
pub struct BaselineScores {
    pub jaccard: Option<f64>,
    pub ssim: Option<f64>,
}

/// Canonical form for case-insensitive, space-tolerant substring matching.
pub fn normalize_name(s: &str) -> String {
    s.to_lowercase().replace(' ', "_")
}

pub fn baseline_key(name: &str) -> String {
    if name.contains('/') {
        name.to_string()
    } else {
        format!("cases/{name}")
    }
}

pub fn load_baselines(output_dir: &Path) -> HashMap<String, BaselineScores> {
    let path = output_dir
        .parent()
        .unwrap_or(output_dir)
        .join("baselines.json");
    std::fs::read_to_string(&path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

pub fn load_hashes(path: &Path) -> HashMap<String, Vec<String>> {
    std::fs::read_to_string(path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

pub fn compute_changed_cases(
    committed: &HashMap<String, Vec<String>>,
    latest: &HashMap<String, Vec<String>>,
) -> HashSet<String> {
    latest
        .iter()
        .filter(|(name, hashes)| committed.get(*name) != Some(hashes))
        .map(|(name, _)| name.clone())
        .collect()
}

/// Copy generated page PNGs to acknowledged/ as a snapshot of the current output.
pub fn snapshot_acknowledged(case_dir: &Path) {
    let gen_dir = case_dir.join("generated");
    let ack_dir = case_dir.join("acknowledged");
    let Ok(entries) = std::fs::read_dir(&gen_dir) else {
        return;
    };
    let _ = std::fs::create_dir_all(&ack_dir);
    // Remove old acknowledged files first
    if let Ok(old) = std::fs::read_dir(&ack_dir) {
        for e in old.flatten() {
            let _ = std::fs::remove_file(e.path());
        }
    }
    for entry in entries.flatten() {
        let path = entry.path();
        if path
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("png"))
        {
            let dest = ack_dir.join(path.file_name().unwrap());
            let _ = std::fs::copy(&path, &dest);
        }
    }
}

pub fn count_pages(case_dir: &Path) -> usize {
    // Count from whichever has more pages (reference or generated)
    let r = count_pngs(&case_dir.join("reference"));
    let g = count_pngs(&case_dir.join("generated"));
    r.max(g)
}

pub fn count_pngs(dir: &Path) -> usize {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return 0;
    };
    entries
        .flatten()
        .filter(|e| {
            e.path()
                .extension()
                .is_some_and(|ext| ext.eq_ignore_ascii_case("png"))
        })
        .count()
}

pub fn file_age(path: &Path) -> Option<String> {
    let modified = std::fs::metadata(path).ok()?.modified().ok()?;
    let elapsed = SystemTime::now().duration_since(modified).ok()?;
    let total_mins = elapsed.as_secs() / 60;
    let hours = total_mins / 60;
    let mins = total_mins % 60;
    if hours > 0 {
        Some(format!("{}h {}m ago", hours, mins))
    } else {
        Some(format!("{}m ago", mins))
    }
}

impl CaseInfo {
    /// Fixture group and bare case name; a name without a slash lives in `cases/`.
    pub fn group_and_name(&self) -> (&str, &str) {
        self.name.split_once('/').unwrap_or(("cases", &self.name))
    }
}
