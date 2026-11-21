//! Fonts a host supplies in memory (`ConvertOptions::font`), for hosts
//! without font directories (a bare wasm instance, a sandbox) and for fonts
//! that are not installed. `convert` installs them on the converting thread
//! for the call.
use std::cell::RefCell;
use std::collections::HashMap;
use std::sync::Arc;

type FontKey = (String, bool, bool);

/// Faces a caller supplies with `ConvertOptions::font`, parsed once.
#[derive(Clone, Default)]
pub(crate) struct MemoryFonts {
    // The rank preserves the directory scanner's regular/Mac Roman/off-weight order.
    index: HashMap<FontKey, (Arc<Vec<u8>>, u32, u8)>,
    files: usize,
}

impl std::fmt::Debug for MemoryFonts {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("MemoryFonts")
            .field("files", &self.files)
            .field("faces", &self.index.len())
            .finish()
    }
}

thread_local! {
    static FONTS: RefCell<Arc<MemoryFonts>> = RefCell::new(Arc::default());
}

impl MemoryFonts {
    /// Adds TrueType/OpenType bytes (all faces of a TTC). Supplied faces take
    /// precedence over installed faces of the same family and style. Returns
    /// the number of valid faces with usable names.
    pub(crate) fn add(&mut self, data: Vec<u8>) -> usize {
        let data = Arc::new(data);
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
                if self.index.get(&key).is_none_or(|entry| rank < entry.2) {
                    self.index.insert(key, (data.clone(), face_index, rank));
                }
            }
        }
        self.files += usize::from(count > 0);
        count
    }
}

/// The supplied fonts, visible to this thread's font lookups while it lives.
pub(crate) struct Installed;

pub(crate) fn install(fonts: &Arc<MemoryFonts>) -> Installed {
    FONTS.set(fonts.clone());
    Installed
}

impl Drop for Installed {
    fn drop(&mut self) {
        FONTS.set(Arc::default());
    }
}

pub(super) fn lookup(key: &FontKey) -> Option<(super::FontSource, u32)> {
    FONTS.with_borrow(|fonts| {
        let (data, face_index, _) = fonts.index.get(key)?;
        Some((super::FontSource::Data(data.clone()), *face_index))
    })
}
