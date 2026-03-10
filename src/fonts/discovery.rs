use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::OnceLock;

use memmap2::Mmap;
use ttf_parser::Face;

use super::cache::{CachedFace, CachedFile, FontCache, dir_mtime, load_cache, save_cache};

/// (lowercase family name, bold, italic) -> (file path, face index within TTC)
type FontLookup = HashMap<(String, bool, bool), (PathBuf, u32)>;

static FONT_INDEX: OnceLock<FontLookup> = OnceLock::new();

fn font_family_name(face: &Face) -> Option<String> {
    // Use ID 1 (Family) — matches what DOCX references and distinguishes
    // "Aptos Display" from "Aptos" from "Aptos Narrow".
    // ID 16 (Typographic Family) groups all these under one name, causing collisions.
    for name in face.names() {
        if name.name_id == ttf_parser::name_id::FAMILY
            && name.is_unicode()
            && let Some(s) = name.to_string()
        {
            return Some(s);
        }
    }
    None
}

fn read_font_style(data: &[u8], face_index: u32) -> Option<(String, bool, bool)> {
    let face = Face::parse(data, face_index).ok()?;
    let family = font_family_name(&face)?;
    Some((family, face.is_bold(), face.is_italic()))
}

fn font_directories() -> Vec<PathBuf> {
    let mut dirs: Vec<PathBuf> = Vec::new();

    // 1. User-configured directories via DOCXSIDE_FONTS env var
    if let Ok(val) = std::env::var("DOCXSIDE_FONTS") {
        let sep = if cfg!(windows) { ';' } else { ':' };
        for part in val.split(sep) {
            let trimmed = part.trim();
            if !trimmed.is_empty() {
                dirs.push(PathBuf::from(trimmed));
            }
        }
    }

    // 2. Platform-specific system font directories
    #[cfg(target_os = "macos")]
    {
        dirs.extend([
            "/Applications/Microsoft Word.app/Contents/Resources/DFonts".into(),
            "/Library/Fonts".into(),
            "/Library/Fonts/Microsoft".into(),
            "/System/Library/Fonts".into(),
            "/System/Library/Fonts/Supplemental".into(),
        ]);
        if let Ok(home) = std::env::var("HOME") {
            dirs.push(PathBuf::from(&home).join("Library/Fonts"));
            let cloud = PathBuf::from(&home)
                .join("Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts");
            if let Ok(families) = std::fs::read_dir(&cloud) {
                for entry in families.flatten() {
                    if entry.path().is_dir() {
                        dirs.push(entry.path());
                    }
                }
            }
        }
    }

    #[cfg(target_os = "linux")]
    {
        dirs.extend(["/usr/share/fonts".into(), "/usr/local/share/fonts".into()]);
        if let Ok(home) = std::env::var("HOME") {
            dirs.push(PathBuf::from(home).join(".local/share/fonts"));
        }
    }

    #[cfg(target_os = "windows")]
    {
        if let Ok(windir) = std::env::var("WINDIR") {
            dirs.push(PathBuf::from(windir).join("Fonts"));
        } else {
            dirs.push("C:\\Windows\\Fonts".into());
        }
    }

    dirs
}

fn is_font_file(path: &std::path::Path) -> bool {
    matches!(
        path.extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_ascii_lowercase())
            .as_deref(),
        Some("ttf" | "otf" | "ttc")
    )
}

fn is_font_collection(path: &std::path::Path) -> bool {
    path.extension()
        .and_then(|e| e.to_str())
        .is_some_and(|e| e.eq_ignore_ascii_case("ttc"))
}

fn scan_font_dirs() -> FontLookup {
    let t0 = std::time::Instant::now();
    let mut index = FontLookup::new();
    let dirs = font_directories();

    let no_cache = std::env::var("DOCXSIDE_NO_FONT_CACHE").is_ok();

    let cache = if no_cache {
        FontCache {
            dir_mtimes: HashMap::new(),
            files: HashMap::new(),
        }
    } else {
        load_cache()
    };
    let mut new_cache = FontCache {
        dir_mtimes: HashMap::new(),
        files: HashMap::new(),
    };
    let mut files_scanned = 0u32;
    let mut dirs_cached = 0u32;
    let mut dirs_scanned = 0u32;
    let mut visited_dirs: std::collections::HashSet<PathBuf> = std::collections::HashSet::new();

    let mut stack: Vec<PathBuf> = dirs;
    while let Some(dir) = stack.pop() {
        if !visited_dirs.insert(dir.clone()) {
            continue;
        }

        let Ok(entries) = std::fs::read_dir(&dir) else {
            continue;
        };

        // Collect directory listing: subdirs to recurse, font files to process
        let mut subdirs = Vec::new();
        let mut font_files = Vec::new();
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                subdirs.push(path);
            } else if is_font_file(&path) {
                font_files.push(path);
            }
        }
        stack.extend(subdirs);

        if font_files.is_empty() {
            continue;
        }

        let current_mtime = dir_mtime(&dir);

        if let Some(&cached_mtime) = cache.dir_mtimes.get(&dir)
            && cached_mtime == current_mtime
        {
            dirs_cached += 1;
            new_cache.dir_mtimes.insert(dir.clone(), current_mtime);
            for file_path in &font_files {
                if let Some(cached_file) = cache.files.get(file_path) {
                    for face in &cached_file.faces {
                        index
                            .entry((face.family.to_lowercase(), face.bold, face.italic))
                            .or_insert((file_path.clone(), face.face_index));
                    }
                    new_cache.files.insert(
                        file_path.clone(),
                        CachedFile {
                            faces: cached_file
                                .faces
                                .iter()
                                .map(|f| CachedFace {
                                    family: f.family.clone(),
                                    bold: f.bold,
                                    italic: f.italic,
                                    face_index: f.face_index,
                                })
                                .collect(),
                        },
                    );
                }
            }
            continue;
        }

        // Directory changed — scan all font files in it
        dirs_scanned += 1;
        new_cache.dir_mtimes.insert(dir, current_mtime);
        for file_path in font_files {
            files_scanned += 1;
            let Ok(file) = std::fs::File::open(&file_path) else {
                continue;
            };
            let Ok(data) = (unsafe { Mmap::map(&file) }) else {
                continue;
            };
            let face_count = if is_font_collection(&file_path) {
                ttf_parser::fonts_in_collection(&data).unwrap_or(1)
            } else {
                1
            };
            let mut faces = Vec::new();
            for face_idx in 0..face_count {
                if let Some((family, bold, italic)) = read_font_style(&data, face_idx) {
                    index
                        .entry((family.to_lowercase(), bold, italic))
                        .or_insert((file_path.clone(), face_idx));
                    faces.push(CachedFace {
                        family,
                        bold,
                        italic,
                        face_index: face_idx,
                    });
                }
            }
            new_cache.files.insert(file_path, CachedFile { faces });
        }
    }

    if !no_cache {
        save_cache(&new_cache);
    }

    log::info!(
        "Font scan: {:.1}ms, {} dirs cached / {} scanned, {} files parsed → {} entries",
        t0.elapsed().as_secs_f64() * 1000.0,
        dirs_cached,
        dirs_scanned,
        files_scanned,
        index.len(),
    );

    index
}

fn get_font_index() -> &'static FontLookup {
    FONT_INDEX.get_or_init(scan_font_dirs)
}

/// Look up a font file by family name and style using the OS/2 table metadata index.
/// Falls back to the regular variant if the requested bold/italic is not available.
/// Returns `(path, face_index, exact_style_match)`.
pub(super) fn find_font_file(
    font_name: &str,
    bold: bool,
    italic: bool,
) -> Option<(PathBuf, u32, bool)> {
    let index = get_font_index();
    let key = font_name.to_lowercase();
    if let Some((path, face_index)) = index.get(&(key.clone(), bold, italic)) {
        return Some((path.clone(), *face_index, true));
    }
    if bold || italic {
        index
            .get(&(key, false, false))
            .map(|(path, face_index)| (path.clone(), *face_index, false))
    } else {
        None
    }
}
