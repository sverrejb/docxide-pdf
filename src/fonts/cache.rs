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

const CACHE_VERSION: &str = "v5";
// A file without this last line was cut short or written by something else.
const END_MARKER: &str = "end";

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
    cache_path().map_or_else(FontCache::default, |p| load_cache_from(&p))
}

fn load_cache_from(path: &Path) -> FontCache {
    let mut fc = FontCache::default();
    let Ok(content) = fs::read_to_string(path) else {
        return fc;
    };
    let mut lines: Vec<&str> = content.lines().collect();
    if lines.first() != Some(&CACHE_VERSION) || lines.pop() != Some(END_MARKER) {
        return fc;
    }
    for line in &lines[1..] {
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
    if let Some(path) = cache_path() {
        save_cache_to(&path, cache);
    }
}

// Best effort: a failed write only costs the next run a rescan.
fn save_cache_to(path: &Path, cache: &FontCache) {
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
    out.push_str(END_MARKER);
    out.push('\n');
    // Concurrent conversions share this file; a reader that caught a plain
    // fs::write half-done saw a truncated index and fell back to Helvetica.
    // rename is atomic on one filesystem, and the pid keeps writers apart.
    let tmp = path.with_extension(format!("tsv.{}.tmp", std::process::id()));
    if fs::write(&tmp, out).is_ok() && fs::rename(&tmp, path).is_err() {
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

#[cfg(test)]
mod tests {
    use super::*;

    fn temp_cache(name: &str) -> PathBuf {
        let dir = env::temp_dir().join(format!("docxide-cache-test-{}-{name}", std::process::id()));
        let _ = fs::remove_dir_all(&dir);
        dir.join("font-index.tsv")
    }

    fn sample() -> FontCache {
        let mut fc = FontCache::default();
        fc.dir_mtimes.insert(PathBuf::from("/fonts"), 1_700_000_000);
        fc.files.insert(
            PathBuf::from("/fonts/calibri.ttf"),
            CachedFile {
                faces: vec![CachedFace {
                    family: "Calibri".into(),
                    bold: true,
                    italic: false,
                    face_index: 2,
                    mac_roman_only: false,
                    off_weight: true,
                }],
            },
        );
        fc.files
            .insert(PathBuf::from("/fonts/broken.ttf"), CachedFile::default());
        fc
    }

    #[test]
    fn round_trip() {
        let path = temp_cache("round-trip");
        save_cache_to(&path, &sample());
        let fc = load_cache_from(&path);
        assert_eq!(fc.dir_mtimes[Path::new("/fonts")], 1_700_000_000);
        let face = &fc.files[Path::new("/fonts/calibri.ttf")].faces[0];
        assert_eq!(
            (&*face.family, face.bold, face.italic, face.face_index),
            ("Calibri", true, false, 2)
        );
        assert!(!face.mac_roman_only && face.off_weight);
        assert!(fc.files[Path::new("/fonts/broken.ttf")].faces.is_empty());
        let leftovers = fs::read_dir(path.parent().unwrap()).unwrap().count();
        assert_eq!(leftovers, 1, "temp file left behind");
        let _ = fs::remove_dir_all(path.parent().unwrap());
    }

    #[test]
    fn torn_or_old_files_load_empty() {
        let path = temp_cache("torn");
        save_cache_to(&path, &sample());
        let full = fs::read_to_string(&path).unwrap();
        let cut_mid_line = &full[..full.find("Calibri").unwrap() + 3];
        let no_marker = full.strip_suffix("end\n").unwrap();
        // Keeps the end marker, so only the version rejects it.
        let v4 = format!("v4{}", &full[CACHE_VERSION.len()..]);
        for content in [cut_mid_line, no_marker, &v4] {
            fs::write(&path, content).unwrap();
            let fc = load_cache_from(&path);
            assert!(
                fc.dir_mtimes.is_empty() && fc.files.is_empty(),
                "{content:?}"
            );
        }
        let _ = fs::remove_dir_all(path.parent().unwrap());
    }
}
