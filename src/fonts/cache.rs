use std::collections::HashMap;
use std::fmt::Write;
use std::path::{Path, PathBuf};
use std::time::UNIX_EPOCH;
use std::{env, fs};

#[derive(Clone)]
pub(super) struct CachedFace {
    pub(super) family: String,
    pub(super) bold: bool,
    pub(super) italic: bool,
    pub(super) face_index: u32,
    pub(super) mac_roman_only: bool,
    pub(super) off_weight: bool,
}

#[derive(Clone, Default)]
pub(super) struct CachedFile {
    pub(super) faces: Vec<CachedFace>,
}

#[derive(Default)]
pub(super) struct FontCache {
    pub(super) dir_mtimes: HashMap<PathBuf, i64>,
    pub(super) files: HashMap<PathBuf, CachedFile>,
}

const CACHE_VERSION: &str = "v4";

fn cache_path() -> Option<PathBuf> {
    let dir = if cfg!(target_os = "macos") {
        env::var("HOME")
            .ok()
            .map(|h| PathBuf::from(h).join("Library/Caches/docxide-pdf"))
    } else if cfg!(target_os = "windows") {
        env::var("LOCALAPPDATA")
            .ok()
            .map(|d| PathBuf::from(d).join("docxide-pdf/cache"))
    } else {
        env::var("XDG_CACHE_HOME")
            .ok()
            .map(PathBuf::from)
            .or_else(|| {
                env::var("HOME")
                    .ok()
                    .map(|h| PathBuf::from(h).join(".cache"))
            })
            .map(|d| d.join("docxide-pdf"))
    };
    dir.map(|d| d.join("font-index.tsv"))
}

pub(super) fn load_cache() -> FontCache {
    let mut fc = FontCache::default();
    let Some(path) = cache_path() else {
        return fc;
    };
    let Ok(content) = fs::read_to_string(&path) else {
        return fc;
    };
    let mut lines = content.lines();
    if lines.next() != Some(CACHE_VERSION) {
        return fc;
    }
    for line in lines {
        let parts: Vec<&str> = line.split('\t').collect();
        match parts.first().copied() {
            Some("D") if parts.len() == 3 => {
                let Ok(mtime) = parts[2].parse::<i64>() else {
                    continue;
                };
                fc.dir_mtimes.insert(PathBuf::from(parts[1]), mtime);
            }
            Some("F") if parts.len() == 8 => {
                let Ok(face_index) = parts[5].parse::<u32>() else {
                    continue;
                };
                fc.files
                    .entry(PathBuf::from(parts[1]))
                    .or_default()
                    .faces
                    .push(CachedFace {
                        family: parts[2].to_string(),
                        bold: parts[3] == "1",
                        italic: parts[4] == "1",
                        face_index,
                        mac_roman_only: parts[6] == "1",
                        off_weight: parts[7] == "1",
                    });
            }
            Some("F") if parts.len() == 3 && parts[2] == "-" => {
                fc.files.entry(PathBuf::from(parts[1])).or_default();
            }
            _ => {}
        }
    }
    fc
}

pub(super) fn save_cache(cache: &FontCache) {
    let Some(path) = cache_path() else {
        return;
    };
    if let Some(dir) = path.parent() {
        let _ = fs::create_dir_all(dir);
    }
    let mut out = String::from(CACHE_VERSION);
    out.push('\n');
    for (dir_path, mtime) in &cache.dir_mtimes {
        let _ = writeln!(out, "D\t{}\t{mtime}", dir_path.to_string_lossy());
    }
    for (file_path, cached) in &cache.files {
        let path_str = file_path.to_string_lossy();
        if cached.faces.is_empty() {
            let _ = writeln!(out, "F\t{path_str}\t-");
        } else {
            for face in &cached.faces {
                let _ = writeln!(
                    out,
                    "F\t{path_str}\t{}\t{}\t{}\t{}\t{}\t{}",
                    face.family,
                    u8::from(face.bold),
                    u8::from(face.italic),
                    face.face_index,
                    u8::from(face.mac_roman_only),
                    u8::from(face.off_weight),
                );
            }
        }
    }
    // Concurrent conversions share this file; a reader that caught a plain
    // fs::write half-done saw a truncated index and fell back to Helvetica.
    // rename is atomic on one filesystem, and the pid keeps writers apart.
    let tmp = path.with_extension(format!("tsv.{}.tmp", std::process::id()));
    if fs::write(&tmp, out).is_ok() && fs::rename(&tmp, &path).is_err() {
        let _ = fs::remove_file(&tmp);
    }
}

pub(super) fn dir_mtime(path: &Path) -> i64 {
    fs::metadata(path)
        .ok()
        .and_then(|m| m.modified().ok())
        .and_then(|t| t.duration_since(UNIX_EPOCH).ok())
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0)
}
