//! How hard the PDF's streams are deflated: page content, object and xref
//! streams, fonts and pictures. The default writes Word-sized files; `Fast`
//! trades size for time, for a PDF that never leaves its host (a preview
//! rendered and shown on the same machine).
//!
//! A conversion sets its compression for its own duration ([`scoped`]) and
//! every stream asks for it ([`zlib`]), like the font registry: per thread,
//! so concurrent conversions on other threads keep their own.
use std::cell::Cell;

/// How hard the PDF's streams are compressed.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum Compression {
    /// The smallest file (the default): zlib level 9 for page content and
    /// object streams, 6 for fonts and pictures.
    #[default]
    Smallest,
    /// Every stream at zlib level 1: the same content once decompressed,
    /// with less compression work and a larger file.
    Fast,
}

thread_local! {
    static CURRENT: Cell<Compression> = const { Cell::new(Compression::Smallest) };
}

/// Restores the compression in effect before [`scoped`].
pub(crate) struct Scope(Compression);

impl Drop for Scope {
    fn drop(&mut self) {
        CURRENT.with(|current| current.set(self.0));
    }
}

/// Use `compression` on this thread until the returned guard drops.
pub(crate) fn scoped(compression: Compression) -> Scope {
    Scope(CURRENT.with(|current| current.replace(compression)))
}

/// `data` zlib-compressed at `smallest_level` (the level this stream uses for
/// the smallest file), or at level 1 under [`Compression::Fast`].
pub(crate) fn zlib(data: &[u8], smallest_level: u8) -> Vec<u8> {
    let level = match CURRENT.with(Cell::get) {
        Compression::Smallest => smallest_level,
        Compression::Fast => 1,
    };
    miniz_oxide::deflate::compress_to_vec_zlib(data, level)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn text() -> Vec<u8> {
        (0..4000)
            .flat_map(|i| format!("BT /F1 12 Tf {i} 700 Td (line {i}) Tj ET\n").into_bytes())
            .collect()
    }

    #[test]
    fn fast_writes_the_same_bytes_once_inflated() {
        let data = text();
        let smallest = zlib(&data, 9);
        let fast = {
            let _scope = scoped(Compression::Fast);
            zlib(&data, 9)
        };
        assert_ne!(smallest, fast);
        for compressed in [smallest, fast] {
            assert_eq!(
                miniz_oxide::inflate::decompress_to_vec_zlib(&compressed).unwrap(),
                data
            );
        }
    }

    #[test]
    fn the_scope_restores_the_previous_compression() {
        {
            let _outer = scoped(Compression::Fast);
            {
                let _inner = scoped(Compression::Smallest);
                assert_eq!(CURRENT.with(Cell::get), Compression::Smallest);
            }
            assert_eq!(CURRENT.with(Cell::get), Compression::Fast);
        }
        assert_eq!(CURRENT.with(Cell::get), Compression::Smallest);
    }
}
