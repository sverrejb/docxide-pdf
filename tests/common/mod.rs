use std::collections::HashMap;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};
use std::{fs, io};

const SKIP_FIXTURES: &[&str] = &["sample100kB"];

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
