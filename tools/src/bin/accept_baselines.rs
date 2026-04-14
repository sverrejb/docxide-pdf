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

fn read_hashes(path: &Path) -> BTreeMap<String, Vec<String>> {
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
    let latest_hashes_path = root.join("tests/output/latest_hashes.json");
    let visual_hashes_path = root.join("tests/visual_hashes.json");

    let latest = read_json(&latest_path);
    let mut baselines = read_json(&baselines_path);
    let latest_hashes = read_hashes(&latest_hashes_path);
    let visual_hashes_before = read_hashes(&visual_hashes_path);
    let mut visual_hashes = visual_hashes_before.clone();

    if latest.is_empty() && latest_hashes.is_empty() {
        eprintln!("No latest scores or hashes found.");
        eprintln!("Run tests first: cargo test -- --nocapture");
        std::process::exit(1);
    }

    let mut accepted = 0;
    let mut unchanged = 0;
    let mut hash_accepted = 0;
    let mut hash_unchanged = 0;

    // Collect all case names from both scores and hashes
    let mut all_cases: Vec<String> = latest.keys().chain(latest_hashes.keys()).cloned().collect();
    all_cases.sort();
    all_cases.dedup();

    for name in &all_cases {
        if !filters.is_empty() && !filters.iter().any(|f| name.contains(f)) {
            continue;
        }

        let score_changed = if let Some(new) = latest.get(name) {
            let existing = baselines.get(name);
            is_changed(existing, new)
        } else {
            false
        };

        let hash_changed = if let Some(new_h) = latest_hashes.get(name) {
            visual_hashes.get(name) != Some(new_h)
        } else {
            false
        };

        if score_changed || hash_changed {
            let existing_scores = baselines.get(name);
            let new_scores = latest.get(name);
            print_change(name, existing_scores, new_scores, hash_changed);

            if score_changed {
                if let Some(new) = new_scores {
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
                }
                accepted += 1;
            }

            if hash_changed {
                if let Some(new_h) = latest_hashes.get(name) {
                    visual_hashes.insert(name.clone(), new_h.clone());
                }
                hash_accepted += 1;
            }
        } else {
            if latest.contains_key(name) {
                unchanged += 1;
            }
            if latest_hashes.contains_key(name) {
                hash_unchanged += 1;
            }
        }
    }

    if accepted == 0 && hash_accepted == 0 {
        println!(
            "No changes to accept ({unchanged} scores unchanged, {hash_unchanged} hashes unchanged)."
        );
        return;
    }

    let mut summary_parts = Vec::new();
    if accepted > 0 {
        summary_parts.push(format!("{accepted} scores accepted"));
    }
    if hash_accepted > 0 {
        summary_parts.push(format!("{hash_accepted} hashes accepted"));
    }
    if unchanged > 0 {
        summary_parts.push(format!("{unchanged} scores unchanged"));
    }
    if hash_unchanged > 0 {
        summary_parts.push(format!("{hash_unchanged} hashes unchanged"));
    }
    println!("\n{}", summary_parts.join(", "));

    if dry_run {
        println!("(dry run — no files modified)");
    } else {
        if accepted > 0 {
            let json = serde_json::to_string_pretty(&baselines).expect("serialize");
            fs::write(&baselines_path, json + "\n").expect("write baselines.json");
            println!("Updated {}", baselines_path.display());
        }
        if hash_accepted > 0 {
            let json = serde_json::to_string_pretty(&visual_hashes).expect("serialize");
            fs::write(&visual_hashes_path, json + "\n").expect("write visual_hashes.json");
            println!("Updated {}", visual_hashes_path.display());
            // Snapshot generated PNGs to acknowledged/ for delta view
            let hash_accepted_cases: Vec<String> = all_cases
                .iter()
                .filter(|name| {
                    if !filters.is_empty() && !filters.iter().any(|f| name.contains(f)) {
                        return false;
                    }
                    latest_hashes
                        .get(*name)
                        .is_some_and(|h| visual_hashes_before.get(*name) != Some(h))
                })
                .cloned()
                .collect();
            snapshot_accepted_cases(&root, &hash_accepted_cases);
        }
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

fn print_change(
    name: &str,
    existing: Option<&Scores>,
    new: Option<&Scores>,
    hash_changed: bool,
) {
    let empty = Scores::default();
    let old = existing.unwrap_or(&empty);
    let new = new.unwrap_or(&empty);

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
    if hash_changed {
        if existing.is_none() {
            parts.push("H:new".into());
        } else {
            parts.push("H:changed".into());
        }
    }

    println!("  {name}  {}", parts.join("  "));
}

/// Copy generated PNGs to acknowledged/ for each accepted case, enabling delta view.
/// Case names are display_name format ("group/truncated.."), resolve to output dirs by scanning.
fn snapshot_accepted_cases(root: &Path, accepted_names: &[String]) {
    if accepted_names.is_empty() {
        return;
    }
    let output_dir = root.join("tests/output");
    let Ok(groups) = fs::read_dir(&output_dir) else {
        return;
    };
    for group_entry in groups.flatten() {
        let group_path = group_entry.path();
        if !group_path.is_dir() {
            continue;
        }
        let group_name = group_path
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("");
        let Ok(cases) = fs::read_dir(&group_path) else {
            continue;
        };
        for case_entry in cases.flatten() {
            let case_path = case_entry.path();
            if !case_path.is_dir() {
                continue;
            }
            let case_name = case_path
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("");
            // Build the display_name key (same truncation as tests/common/mod.rs)
            let short = if case_name.len() > 16 {
                format!("{}..", &case_name[..16])
            } else {
                case_name.to_string()
            };
            let key = format!("{group_name}/{short}");
            if !accepted_names.contains(&key) {
                continue;
            }
            let gen_dir = case_path.join("generated");
            let ack_dir = case_path.join("acknowledged");
            let Ok(entries) = fs::read_dir(&gen_dir) else {
                continue;
            };
            let _ = fs::create_dir_all(&ack_dir);
            // Clear old acknowledged files
            if let Ok(old) = fs::read_dir(&ack_dir) {
                for e in old.flatten() {
                    let _ = fs::remove_file(e.path());
                }
            }
            for entry in entries.flatten() {
                let path = entry.path();
                if path
                    .extension()
                    .is_some_and(|e| e.eq_ignore_ascii_case("png"))
                {
                    let dest = ack_dir.join(path.file_name().unwrap());
                    let _ = fs::copy(&path, &dest);
                }
            }
        }
    }
}
