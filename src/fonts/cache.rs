use std::collections::HashMap;
use std::fmt::Write;
use std::path::PathBuf;

#[derive(Clone)]
pub(super) struct CachedFace {
    pub(super) family: String,
    pub(super) bold: bool,
    pub(super) italic: bool,
    pub(super) face_index: u32,
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

pub(super) const CACHE_VERSION: &str = "v1";

pub(super) fn cache_path() -> Option<PathBuf> {
    let dir = if cfg!(target_os = "macos") {
        std::env::var("HOME")
            .ok()
            .map(|h| PathBuf::from(h).join("Library/Caches/docxide-pdf"))
    } else if cfg!(target_os = "windows") {
        std::env::var("LOCALAPPDATA")
            .ok()
            .map(|d| PathBuf::from(d).join("docxide-pdf/cache"))
    } else {
        std::env::var("XDG_CACHE_HOME")
            .ok()
            .map(PathBuf::from)
            .or_else(|| {
                std::env::var("HOME")
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
    let Ok(content) = std::fs::read_to_string(&path) else {
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
            Some("F") if parts.len() == 6 => {
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
        let _ = std::fs::create_dir_all(dir);
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
                    "F\t{path_str}\t{}\t{}\t{}\t{}",
                    face.family,
                    u8::from(face.bold),
                    u8::from(face.italic),
                    face.face_index,
                );
            }
        }
    }
    let _ = std::fs::write(&path, out);
}

pub(super) fn dir_mtime(path: &std::path::Path) -> i64 {
    std::fs::metadata(path)
        .ok()
        .and_then(|m| m.modified().ok())
        .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0)
}
