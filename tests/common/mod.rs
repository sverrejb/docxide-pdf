use std::collections::{HashMap, HashSet};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};
use std::{fs, io};

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
