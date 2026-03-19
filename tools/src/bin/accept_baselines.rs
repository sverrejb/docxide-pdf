use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::{env, fs};

#[derive(Clone, Default, Serialize, Deserialize)]
struct Scores {
    #[serde(skip_serializing_if = "Option::is_none")]
    jaccard: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    ssim: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    text_boundary: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    convert_ms: Option<f64>,
}

fn find_project_root() -> PathBuf {
    for candidate in &[PathBuf::from("."), PathBuf::from("..")] {
        if candidate.join("tests/baselines.json").exists() {
            return candidate.clone();
        }
    }
    eprintln!("Could not find tests/baselines.json. Run from the project root or tools/.");
    std::process::exit(1);
}

fn read_json(path: &Path) -> BTreeMap<String, Scores> {
    fs::read_to_string(path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

fn fmt_score(v: Option<f64>) -> String {
    v.map(|x| format!("{:.1}", x * 100.0))
        .unwrap_or_else(|| "-".into())
}

fn fmt_delta(old: Option<f64>, new: Option<f64>) -> String {
    match (old, new) {
        (Some(o), Some(n)) => {
            let diff = (n - o) * 100.0;
            if diff.abs() < 0.05 {
                String::new()
            } else if diff > 0.0 {
                format!(" \x1b[32m(+{diff:.1})\x1b[0m")
            } else {
                format!(" \x1b[31m({diff:.1})\x1b[0m")
            }
        }
        (None, Some(_)) => " (new)".into(),
        _ => String::new(),
    }
}

fn main() {
    let args: Vec<String> = env::args().skip(1).collect();
    let dry_run = args.iter().any(|a| a == "--dry-run");
    let filters: Vec<&str> = args
        .iter()
        .filter(|a| !a.starts_with("--"))
        .map(|s| s.as_str())
        .collect();

    let root = find_project_root();
    let latest_path = root.join("tests/output/latest_scores.json");
    let baselines_path = root.join("tests/baselines.json");

    let latest = read_json(&latest_path);
    let mut baselines = read_json(&baselines_path);

    if latest.is_empty() {
        eprintln!("No latest scores found at {}.", latest_path.display());
        eprintln!("Run tests first: cargo test -- --nocapture");
        std::process::exit(1);
    }

    let mut accepted = 0;
    let mut unchanged = 0;

    for (name, new) in &latest {
        if !filters.is_empty() && !filters.iter().any(|f| name.contains(f)) {
            continue;
        }

        let existing = baselines.get(name);
        let changed = is_changed(existing, new);

        if changed {
            print_change(name, existing, new);
            // Merge per-field so we don't lose fields not present in latest
            let entry = baselines.entry(name.clone()).or_default();
            if let Some(v) = new.jaccard {
                entry.jaccard = Some(v);
            }
            if let Some(v) = new.ssim {
                entry.ssim = Some(v);
            }
            if let Some(v) = new.text_boundary {
                entry.text_boundary = Some(v);
            }
            if let Some(v) = new.convert_ms {
                entry.convert_ms = Some(v);
            }
            accepted += 1;
        } else {
            unchanged += 1;
        }
    }

    if accepted == 0 {
        println!("No changes to accept ({unchanged} unchanged).");
        return;
    }

    println!("\n{accepted} accepted, {unchanged} unchanged.");

    if dry_run {
        println!("(dry run — baselines.json not modified)");
    } else {
        let json = serde_json::to_string_pretty(&baselines).expect("serialize");
        fs::write(&baselines_path, json + "\n").expect("write baselines.json");
        println!("Updated {}", baselines_path.display());
    }
}

fn is_changed(existing: Option<&Scores>, new: &Scores) -> bool {
    let Some(old) = existing else {
        return true;
    };
    field_changed(old.jaccard, new.jaccard)
        || field_changed(old.ssim, new.ssim)
        || field_changed(old.text_boundary, new.text_boundary)
        || field_changed(old.convert_ms, new.convert_ms)
}

fn field_changed(old: Option<f64>, new: Option<f64>) -> bool {
    match (old, new) {
        (Some(o), Some(n)) => (o - n).abs() > 0.00005,
        (None, Some(_)) => true,
        _ => false,
    }
}

fn print_change(name: &str, existing: Option<&Scores>, new: &Scores) {
    let empty = Scores::default();
    let old = existing.unwrap_or(&empty);

    let mut parts = Vec::new();
    if field_changed(old.jaccard, new.jaccard) {
        parts.push(format!(
            "J:{}→{}{}",
            fmt_score(old.jaccard),
            fmt_score(new.jaccard),
            fmt_delta(old.jaccard, new.jaccard)
        ));
    }
    if field_changed(old.ssim, new.ssim) {
        parts.push(format!(
            "S:{}→{}{}",
            fmt_score(old.ssim),
            fmt_score(new.ssim),
            fmt_delta(old.ssim, new.ssim)
        ));
    }
    if field_changed(old.text_boundary, new.text_boundary) {
        parts.push(format!(
            "T:{}→{}{}",
            fmt_score(old.text_boundary),
            fmt_score(new.text_boundary),
            fmt_delta(old.text_boundary, new.text_boundary)
        ));
    }
    if field_changed(old.convert_ms, new.convert_ms) {
        let old_ms = old
            .convert_ms
            .map(|v| format!("{v:.0}"))
            .unwrap_or("-".into());
        let new_ms = new
            .convert_ms
            .map(|v| format!("{v:.0}"))
            .unwrap_or("-".into());
        parts.push(format!("ms:{old_ms}→{new_ms}"));
    }

    println!("  {name}  {}", parts.join("  "));
}
