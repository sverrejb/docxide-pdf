use std::collections::HashMap;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};
use std::{fs, io};

const SKIP_FIXTURES: &[&str] = &[
    "sample100kB",
    // Wrong font — DOCX requires fonts not available on this system.
    // Re-enable when font resolution/fallback improves.
    "1bb388ab50997e6c545fdfda84f4eb0ab2869de68111b1dd3c3e952d368d379c", // missing TimesNewRoman+Calibri combo, falls back to Helvetica/Tahoma
    "0ad33844870c6ad7993f3778ed6a75ffba7c28da28a2bd5b033c26572dc6c398", // missing ＭＳゴシック (Japanese), falls back to Helvetica
    "7c4438cbb8d2b6439c70c182d86cd4a8b8cc355fc490aca3ef282cc51cea3468", // missing Times, uses Arial/TimesNewRoman instead
    "50d8bc19a389d4e8b9a9a9b9fcc8d28c58ac5a9f4e5d9d4206909b78919ac7fc", // missing MSSansSerif, falls back to Calibri/Helvetica
    "6a43b590ebf28dfc516277c90b288c2d222c7aaf3f48e97d0b893ece583406e6", // missing TimesNewRoman, falls back to Arial
    "c47b46a10e31ef720d9d54f60a7899ff2546672b91ad0d8af5e9c2cd0a2a0022", // missing Times, falls back to Aptos/Helvetica
    "f271d69a2fca4461c732d9431b6cc1e59e27e86cfba6cf06859863645146bb35", // missing TimesNewRomanPSMT, falls back to Helvetica
    "3513b16a0fb47b39e900b4a9bf4b3f6806445ab261df280940587fddd840935c", // missing Calibri (uses TimesNewRoman only)
    "5ba6b6915b4c096de47db896b65dc7484f23d2578d868e0a741b377eed19f8bb", // no matched fonts, falls back to Helvetica
    "347593422a3890bcb721834b94978c528bd6cc4bfd59e4f9f3e8bc647c3be5b9", // missing lemming font, falls back to Calibri/Helvetica
];
const SKIP_GROUPS: &[&str] = &[];

fn natural_cmp(a: &Path, b: &Path) -> std::cmp::Ordering {
    let a = a.file_name().and_then(|n| n.to_str()).unwrap_or("");
    let b = b.file_name().and_then(|n| n.to_str()).unwrap_or("");
    let extract = |s: &str| -> (String, u64) {
        let i = s.find(|c: char| c.is_ascii_digit()).unwrap_or(s.len());
        (s[..i].to_string(), s[i..].parse().unwrap_or(0))
    };
    extract(a).cmp(&extract(b))
}

pub fn discover_fixtures() -> io::Result<Vec<PathBuf>> {
    let fixtures_dir = Path::new("tests/fixtures");
    let case_filter = std::env::var("DOCXIDE_CASE").ok();
    let mut fixtures: Vec<PathBuf> = Vec::new();
    for group_entry in fs::read_dir(fixtures_dir)? {
        let group = group_entry?.path();
        if !group.is_dir() {
            continue;
        }
        let group_name = group.file_name().and_then(|n| n.to_str()).unwrap_or("");
        if SKIP_GROUPS.contains(&group_name) {
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
            } else if !SKIP_FIXTURES.contains(&name) {
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

pub fn read_previous_scores(csv_name: &str, score_col: usize) -> HashMap<String, f64> {
    let csv_path = PathBuf::from("tests/output").join(csv_name);
    let mut latest: HashMap<String, f64> = HashMap::new();
    let Ok(content) = fs::read_to_string(&csv_path) else {
        return latest;
    };
    for line in content.lines().skip(1) {
        let cols: Vec<&str> = line.split(',').collect();
        if cols.len() > score_col {
            if let Ok(score) = cols[score_col].parse::<f64>() {
                latest.insert(cols[1].to_string(), score);
            }
        }
    }
    latest
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
