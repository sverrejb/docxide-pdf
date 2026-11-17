//! Host-provided fonts and lookup diagnostics. State belongs to the converting
//! thread (to the wasm instance on the bare wasm target), never another caller.
use std::cell::{Cell, RefCell};
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::Arc;

type FontKey = (String, bool, bool);

#[derive(Default)]
struct MemoryFonts {
    // The rank preserves the directory scanner's regular/Mac Roman/off-weight order.
    index: HashMap<FontKey, (PathBuf, u32, u8)>,
    data: HashMap<PathBuf, Arc<[u8]>>,
}

/// Family names absent from the host/system index during one resolution, in
/// lookup order, and the requested style. Includes missing exact styles even
/// when a regular face can be used with synthetic bold or italic.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontRequest {
    pub names: Vec<String>,
    pub bold: bool,
    pub italic: bool,
}

#[derive(Default)]
struct Requests {
    current: Vec<FontRequest>,
    completed: Vec<FontRequest>,
}

thread_local! {
    static FONTS: RefCell<MemoryFonts> = RefCell::new(MemoryFonts::default());
    static REQUESTS: RefCell<Requests> = RefCell::new(Requests::default());
    static SYSTEM_FONTS: Cell<bool> = const { Cell::new(!cfg!(all(target_arch = "wasm32", target_os = "unknown"))) };
}

/// Register TrueType/OpenType bytes (including all faces of a TTC) on this
/// thread. Registered faces take precedence over system faces of the same
/// family and style. Returns the number of valid faces with usable names.
/// Fonts persist across conversions until [`clear_fonts`] on this thread.
pub fn register_font(data: Vec<u8>) -> usize {
    let data: Arc<[u8]> = data.into();
    FONTS.with_borrow_mut(|fonts| {
        let path = PathBuf::from(format!("memory:{}", fonts.data.len()));
        let mut count = 0;
        for face_index in 0..ttf_parser::fonts_in_collection(&data).unwrap_or(1) {
            let faces = super::discovery::read_faces(&data, face_index);
            count += usize::from(!faces.is_empty());
            for face in faces {
                let key = (face.family.to_lowercase(), face.bold, face.italic);
                let rank = if face.off_weight {
                    2
                } else {
                    u8::from(face.mac_roman_only)
                };
                if fonts.index.get(&key).is_none_or(|entry| rank < entry.2) {
                    fonts.index.insert(key, (path.clone(), face_index, rank));
                }
            }
        }
        if count > 0 {
            fonts.data.insert(path, data);
        }
        count
    })
}

/// Forget this thread's registered fonts. Does not change the system-font policy.
pub fn clear_fonts() {
    FONTS.with_borrow_mut(|fonts| *fonts = MemoryFonts::default());
}

/// Enable discovery of configured and system font directories on this thread.
/// Defaults to false on bare wasm and true elsewhere. Embedded DOCX fonts are
/// still available when directory discovery is disabled.
pub fn use_system_fonts(enabled: bool) {
    SYSTEM_FONTS.set(enabled);
}

pub(super) fn system_fonts_enabled() -> bool {
    SYSTEM_FONTS.get()
}

pub(super) fn lookup(key: &FontKey) -> Option<(PathBuf, u32)> {
    FONTS.with_borrow(|fonts| {
        fonts
            .index
            .get(key)
            .map(|(path, face, _)| (path.clone(), *face))
    })
}

pub(super) fn font_bytes(path: &Path) -> Option<Arc<[u8]>> {
    FONTS.with_borrow(|fonts| fonts.data.get(path).cloned())
}

pub(crate) fn read_font(path: &Path) -> Option<Vec<u8>> {
    font_bytes(path)
        .map(|data| data.to_vec())
        .or_else(|| std::fs::read(path).ok())
}

/// Drain this thread's missing-font requests from its latest conversion. A new
/// conversion clears old requests, including when parsing fails. Requests are
/// deduplicated across render passes. Register a requested face and retry to
/// replace a fallback; the list describes the index, not embedded DOCX fonts.
pub fn take_font_requests() -> Vec<FontRequest> {
    REQUESTS.with_borrow_mut(|requests| std::mem::take(&mut requests.completed))
}

pub(crate) fn reset_requests() {
    REQUESTS.with_borrow_mut(|requests| *requests = Requests::default());
}

pub(super) struct RequestScope;

pub(super) fn begin_request(bold: bool, italic: bool) -> RequestScope {
    REQUESTS.with_borrow_mut(|requests| {
        requests.current.push(FontRequest {
            names: Vec::new(),
            bold,
            italic,
        })
    });
    RequestScope
}

fn record(requests: &mut Requests, request: FontRequest) {
    if !request.names.is_empty() && !requests.completed.contains(&request) {
        requests.completed.push(request);
    }
}

impl Drop for RequestScope {
    fn drop(&mut self) {
        REQUESTS.with_borrow_mut(|requests| {
            if let Some(request) = requests.current.pop() {
                record(requests, request);
            }
        });
    }
}

pub(super) fn note_miss(name: &str, bold: bool, italic: bool) {
    REQUESTS.with_borrow_mut(|requests| {
        if let Some(current) = requests.current.last_mut() {
            if !current.names.iter().any(|n| n.eq_ignore_ascii_case(name)) {
                current.names.push(name.to_owned());
            }
        } else {
            record(
                requests,
                FontRequest {
                    names: vec![name.to_owned()],
                    bold,
                    italic,
                },
            );
        }
    });
}
