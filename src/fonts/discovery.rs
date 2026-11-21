use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::Instant;

use memmap2::Mmap;
use ttf_parser::Face;

use super::cache::{CachedFace, CachedFile, FontCache, dir_mtime, load_cache, save_cache};

/// (lowercase family name, bold, italic) -> (file path, face index within TTC)
type FontLookup = HashMap<(String, bool, bool), (PathBuf, u32)>;

/// What a scan is built from: the directories a conversion adds to the
/// system ones (the first listed wins) and whether the disk cache is used.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct FontSetup {
    pub(crate) dirs: Vec<PathBuf>,
    pub(crate) cache: bool,
    /// Whether the system font directories are scanned too.
    pub(crate) system: bool,
}

#[derive(Default)]
pub(crate) struct FontIndex {
    lookup: FontLookup,
    /// Lowercase families that only Mac Roman–named macOS system faces provide.
    mac_only: HashSet<String>,
}

impl FontIndex {
    /// One scan per distinct setup for the life of the process.
    pub(crate) fn for_setup(setup: &FontSetup) -> &'static FontIndex {
        // ponytail: leaked once per distinct setup; a server cycling through
        // many font-directory sets would keep an index per set.
        static INDEXES: Mutex<Vec<(FontSetup, &'static FontIndex)>> = Mutex::new(Vec::new());
        let mut indexes = INDEXES.lock().unwrap_or_else(|e| e.into_inner());
        if let Some((_, index)) = indexes.iter().find(|(s, _)| s == setup) {
            return index;
        }
        let index: &'static FontIndex = Box::leak(Box::new(scan_font_dirs(setup)));
        indexes.push((setup.clone(), index));
        index
    }
}

/// Return all localized family names for a font face (deduplicated by lowercase),
/// and whether they came only from Mac Roman records: macOS system faces such as
/// Helvetica, Times, Courier, Optima and Apple's Symbol carry no Unicode family
/// name. Those rank below every other face (see `scan_font_dirs`).
fn font_family_names(face: &Face) -> (Vec<String>, bool) {
    let mut seen = HashSet::new();
    let mut names = Vec::new();
    let family = || {
        face.names()
            .into_iter()
            .filter(|n| n.name_id == ttf_parser::name_id::FAMILY)
    };
    for s in family()
        .filter(|n| n.is_unicode())
        .filter_map(|n| n.to_string())
    {
        if seen.insert(s.to_lowercase()) {
            names.push(s);
        }
    }
    if !names.is_empty() {
        return (names, false);
    }
    // ponytail: ASCII-only Mac Roman decode; every such family name seen is ASCII.
    for n in family().filter(|n| {
        n.platform_id == ttf_parser::PlatformId::Macintosh
            && n.encoding_id == 0
            && n.name.is_ascii()
    }) {
        let s = String::from_utf8_lossy(n.name).into_owned();
        if seen.insert(s.to_lowercase()) {
            names.push(s);
        }
    }
    (names, true)
}

/// One face under each of its localized family names, so a lookup by any of
/// them finds it.
pub(super) fn read_faces(data: &[u8], face_index: u32) -> Vec<CachedFace> {
    let Ok(face) = Face::parse(data, face_index) else {
        return Vec::new();
    };
    let (names, mac_roman_only) = font_family_names(&face);
    // Thin, Light, Black… faces that share the family name of the regular
    // and bold ones (Apple's HelveticaNeue.ttc: only its Thin face carries a
    // Unicode family name, so it outranked the Mac-Roman-only Regular).
    let off_weight = !(350..=750).contains(&face.weight().to_number());
    names
        .into_iter()
        .map(|family| CachedFace {
            family,
            bold: face.is_bold(),
            italic: face.is_italic(),
            face_index,
            mac_roman_only,
            off_weight,
        })
        .collect()
}

fn font_directories(extra: &[PathBuf], system: bool) -> Vec<PathBuf> {
    if !system {
        return extra.iter().rev().cloned().collect();
    }
    let mut dirs = Vec::new();

    // The scan is a LIFO stack and the first-scanned face wins, so later entries
    // here outrank earlier ones. Platform-specific system font directories first.
    #[cfg(target_os = "macos")]
    {
        dirs.extend([
            "/Library/Fonts".into(),
            "/System/Library/Fonts".into(),
            "/System/Library/Fonts/Supplemental".into(),
        ]);
        if let Ok(home) = std::env::var("HOME") {
            dirs.push(PathBuf::from(&home).join("Library/Fonts"));
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

    // The conversion's own directories (`ConvertOptions::font_dirs`) pushed last
    // = HIGHEST priority. A directory supplied explicitly pins the font build
    // regardless of what the host has installed (#105); the system directories
    // still fill in the families and styles it lacks. The first directory listed
    // is pushed last, so the scan pops it first and it wins over the ones after
    // it (like PATH).
    dirs.extend(extra.iter().rev().cloned());

    dirs
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn first_listed_custom_dir_is_popped_first() {
        let dirs = font_directories(&["a".into(), "b".into(), "c".into()], true);
        assert_eq!(
            dirs[dirs.len() - 3..],
            [PathBuf::from("c"), "b".into(), "a".into()]
        );
        let own = font_directories(&["a".into(), "b".into()], false);
        assert_eq!(own, [PathBuf::from("b"), "a".into()]);
    }
}

#[cfg(test)]
pub(crate) fn test_index() -> &'static FontIndex {
    // The fonts/ mirror that .cargo/config.toml names, as the harness reads it.
    let dirs = std::env::var_os("DOCXIDE_FONTS").unwrap_or_default();
    let options = crate::ConvertOptions::new()
        .font_dirs(std::env::split_paths(&dirs).filter(|d| !d.as_os_str().is_empty()))
        .font_cache(std::env::var("DOCXIDE_NO_FONT_CACHE").map_or(true, |v| {
            matches!(v.trim().to_ascii_lowercase().as_str(), "" | "0" | "false")
        }));
    FontIndex::for_setup(&options.font_setup())
}

fn is_font_file(path: &Path) -> bool {
    path.extension()
        .and_then(std::ffi::OsStr::to_str)
        .is_some_and(|e| {
            ["ttf", "otf", "ttc"]
                .iter()
                .any(|x| e.eq_ignore_ascii_case(x))
        })
}

fn scan_font_dirs(setup: &FontSetup) -> FontIndex {
    // A bare wasm host has no font directories, and no clock to time a scan.
    if cfg!(all(target_arch = "wasm32", target_os = "unknown")) {
        return FontIndex::default();
    }
    if !setup.system && setup.dirs.is_empty() {
        return FontIndex::default();
    }
    let t0 = Instant::now();
    let mut index = FontLookup::new();
    // Mac-Roman-only faces fill in only the families nothing else provides, so
    // Apple's Symbol never shadows the vendored Microsoft Symbol (bullets).
    let mut mac_roman_index = FontLookup::new();
    // Off-weight faces fill in only the styles no regular-weight face provides.
    let mut off_weight_index = FontLookup::new();
    let dirs = font_directories(&setup.dirs, setup.system);

    let no_cache = !setup.cache;

    let cache = if no_cache {
        FontCache::default()
    } else {
        load_cache()
    };
    let mut new_cache = FontCache::default();
    let mut files_scanned: u32 = 0;
    let mut dirs_cached: u32 = 0;
    let mut dirs_scanned: u32 = 0;
    let mut visited_dirs = HashSet::new();

    let mut index_faces = |faces: &[CachedFace], file_path: &PathBuf| {
        for face in faces {
            let tier = if face.off_weight {
                &mut off_weight_index
            } else if face.mac_roman_only {
                &mut mac_roman_index
            } else {
                &mut index
            };
            tier.entry((face.family.to_lowercase(), face.bold, face.italic))
                .or_insert((file_path.clone(), face.face_index));
        }
    };

    let mut stack: Vec<PathBuf> = dirs;
    while let Some(dir) = stack.pop() {
        if !visited_dirs.insert(dir.clone()) {
            continue;
        }

        let Ok(entries) = fs::read_dir(&dir) else {
            continue;
        };

        let mut subdirs = Vec::new();
        let mut font_files = Vec::new();
        for entry in entries.flatten() {
            let path = entry.path();
            // readdir's file type spares a stat per font file; only symlinks need one.
            let is_dir = match entry.file_type() {
                Ok(t) if t.is_symlink() => path.is_dir(),
                Ok(t) => t.is_dir(),
                Err(_) => false,
            };
            if is_dir {
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
                    index_faces(&cached_file.faces, file_path);
                    new_cache
                        .files
                        .insert(file_path.clone(), cached_file.clone());
                }
            }
            continue;
        }

        // Directory changed — scan all font files in it
        dirs_scanned += 1;
        new_cache.dir_mtimes.insert(dir, current_mtime);
        for file_path in font_files {
            files_scanned += 1;
            let Ok(file) = fs::File::open(&file_path) else {
                continue;
            };
            let Ok(data) = (unsafe { Mmap::map(&file) }) else {
                continue;
            };
            let face_count = ttf_parser::fonts_in_collection(&data).unwrap_or(1);
            let faces: Vec<CachedFace> = (0..face_count)
                .flat_map(|face_idx| read_faces(&data, face_idx))
                .collect();
            index_faces(&faces, &file_path);
            new_cache.files.insert(file_path, CachedFile { faces });
        }
    }

    let mut mac_only = HashSet::new();
    for (key, value) in mac_roman_index {
        if let std::collections::hash_map::Entry::Vacant(slot) = index.entry(key) {
            mac_only.insert(slot.key().0.clone());
            slot.insert(value);
        }
    }
    for (key, value) in off_weight_index {
        index.entry(key).or_insert(value);
    }

    // The cache holds what one scan visited: a scan without the system
    // directories would drop them, and the next full scan would re-parse them.
    if !no_cache && setup.system {
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

    FontIndex {
        lookup: index,
        mac_only,
    }
}

impl FontIndex {
    /// True if only a macOS system face provides this family (Helvetica, Optima…),
    /// which Windows Word, including Mac Word's online PDF export, never has.
    pub(super) fn is_mac_only_family(&self, name: &str) -> bool {
        let key = name.to_lowercase();
        super::memory::lookup(&(key.clone(), false, false)).is_none()
            && self.mac_only.contains(&key)
    }
}

impl super::FontSource {
    /// Parse one face (a file is memory-mapped, so only the tables touched are read).
    pub(super) fn probe<T>(&self, face_index: u32, f: impl FnOnce(&Face) -> T) -> Option<T> {
        let parse = |data: &[u8]| Face::parse(data, face_index).ok().map(|face| f(&face));
        match self {
            Self::Data(data) => parse(data),
            Self::File(path) => {
                let file = fs::File::open(path).ok()?;
                parse(&unsafe { Mmap::map(&file) }.ok()?)
            }
        }
    }
}

impl FontIndex {
    /// Look up a font file by family name and style using the OS/2 table metadata index.
    /// Falls back to the regular variant if the requested bold/italic is not available.
    /// Returns `(source, face_index, face_bold, face_italic)`.
    pub(super) fn find_font_file(
        &self,
        font_name: &str,
        bold: bool,
        italic: bool,
    ) -> Option<(super::FontSource, u32, bool, bool)> {
        let key = font_name.to_lowercase();
        let lookup = |bold, italic| self.exact_face(&(key.clone(), bold, italic));
        if let Some((source, face_index)) = lookup(bold, italic) {
            return Some((source, face_index, bold, italic));
        }
        // A plain regular request that misses falls through to a generic family
        // fallback rather than borrowing this family's bold/italic face, which would
        // render unexpectedly heavy/slanted.
        if !bold && !italic {
            return None;
        }
        // A styled variant was requested but the exact (bold, italic) face is absent.
        // Relax the style axes and accept any same-family face: keeping the family's
        // real glyphs beats dropping to a different family, and synthetic bold/oblique
        // covers the style gap (we report exact_match=false). Some families ship only a
        // single styled face — e.g. Vivaldi is an italic-flagged script with no upright
        // or bold variant, so "Vivaldi bold" only resolves once we try (regular,italic).
        for (b, i) in [(bold, !italic), (!bold, italic), (!bold, !italic)] {
            if let Some((source, face_index)) = lookup(b, i) {
                return Some((source, face_index, b, i));
            }
        }
        None
    }

    /// The face of exactly this family and style, supplied or installed.
    pub(super) fn exact_face(
        &self,
        key: &(String, bool, bool),
    ) -> Option<(super::FontSource, u32)> {
        super::memory::lookup(key).or_else(|| {
            let (path, face_index) = self.lookup.get(key)?;
            Some((super::FontSource::File(path.clone()), *face_index))
        })
    }
}
