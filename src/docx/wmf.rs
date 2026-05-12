//! Minimal WMF (Windows Metafile) handling.
//!
//! Word frequently embeds bitmap content as an OLE "StaticMetafile" — a WMF
//! container whose actual payload is a single `META_DIBBITBLT`,
//! `META_DIBSTRETCHBLT`, or `META_STRETCHDIB` record carrying a Device-
//! Independent Bitmap. For these (the dominant case in real-world DOCX
//! files), we can extract the DIB and turn it into a valid BMP byte stream
//! by prepending a 14-byte `BITMAPFILEHEADER`, then feed it through the
//! existing BMP image path.
//!
//! True vector WMFs (with `LineTo`, `Polygon`, `ExtTextOut`, etc.) are not
//! handled — `wmf_to_bmp` returns `None` and the caller falls back to
//! reserving layout space only.
//!
//! WMF binary layout reference: [MS-WMF].

const PLACEABLE_MAGIC: [u8; 4] = [0xD7, 0xCD, 0xC6, 0x9A];
const STANDARD_MAGIC: [u8; 4] = [0x01, 0x00, 0x09, 0x00];

const META_DIBBITBLT: u16 = 0x0940; // params before DIB: rasterOp(4) + ySrc(2) + xSrc(2) + h(2) + w(2) + yDest(2) + xDest(2)
const META_DIBSTRETCHBLT: u16 = 0x0B41; // adds srcH/srcW: 20 param bytes
const META_STRETCHDIB: u16 = 0x0F43; // adds colorUsage: 22 param bytes

pub(super) fn is_wmf(data: &[u8]) -> bool {
    data.len() >= 4 && (data[..4] == PLACEABLE_MAGIC || data[..4] == STANDARD_MAGIC)
}

pub(super) fn wmf_to_bmp(data: &[u8]) -> Option<Vec<u8>> {
    let header_skip = if data.starts_with(&PLACEABLE_MAGIC) {
        22 + 18 // placeable header + standard WMF header
    } else if data.starts_with(&STANDARD_MAGIC) {
        18
    } else {
        return None;
    };

    let mut i = header_skip;
    while i + 6 <= data.len() {
        let size_words = u32::from_le_bytes(data[i..i + 4].try_into().ok()?) as usize;
        let func = u16::from_le_bytes(data[i + 4..i + 6].try_into().ok()?);
        let size_bytes = size_words.checked_mul(2)?;
        if size_bytes < 6 || i.checked_add(size_bytes)? > data.len() {
            return None;
        }

        let params_before_dib = match func {
            META_DIBBITBLT => 16,
            META_DIBSTRETCHBLT => 20,
            META_STRETCHDIB => 22,
            _ => {
                i += size_bytes;
                continue;
            }
        };

        let dib_start = i + 6 + params_before_dib;
        let dib_end = i + size_bytes;
        if dib_start >= dib_end {
            return None;
        }
        return dib_to_bmp(&data[dib_start..dib_end]);
    }
    None
}

fn dib_to_bmp(dib: &[u8]) -> Option<Vec<u8>> {
    if dib.len() < 20 {
        return None;
    }
    let bih_size = u32::from_le_bytes(dib[0..4].try_into().ok()?) as usize;
    if bih_size < 12 || bih_size > dib.len() {
        return None;
    }
    let bpp = u16::from_le_bytes(dib[14..16].try_into().ok()?) as usize;
    let palette_bytes = if bpp <= 8 { (1usize << bpp) * 4 } else { 0 };
    if bih_size + palette_bytes > dib.len() {
        return None;
    }
    let pixel_off = 14 + bih_size + palette_bytes;
    let file_size = 14 + dib.len();

    let mut bmp = Vec::with_capacity(file_size);
    bmp.extend_from_slice(b"BM");
    bmp.extend_from_slice(&(file_size as u32).to_le_bytes());
    bmp.extend_from_slice(&0u16.to_le_bytes()); // reserved1
    bmp.extend_from_slice(&0u16.to_le_bytes()); // reserved2
    bmp.extend_from_slice(&(pixel_off as u32).to_le_bytes());
    bmp.extend_from_slice(dib);
    Some(bmp)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_non_wmf() {
        assert!(!is_wmf(b"\x89PNG\r\n\x1a\n"));
        assert!(!is_wmf(b""));
        assert!(wmf_to_bmp(b"\x89PNG\r\n\x1a\n").is_none());
    }

    #[test]
    fn dib_to_bmp_24bpp_roundtrip() {
        // Minimal 2x1 24-bpp DIB: BITMAPINFOHEADER (40 bytes) + 6 bytes pixel data + 2 bytes padding
        let mut dib = Vec::new();
        dib.extend_from_slice(&40u32.to_le_bytes()); // size
        dib.extend_from_slice(&2i32.to_le_bytes()); // width
        dib.extend_from_slice(&1i32.to_le_bytes()); // height
        dib.extend_from_slice(&1u16.to_le_bytes()); // planes
        dib.extend_from_slice(&24u16.to_le_bytes()); // bpp
        dib.extend_from_slice(&0u32.to_le_bytes()); // compression
        dib.extend_from_slice(&8u32.to_le_bytes()); // image size
        dib.extend_from_slice(&0i32.to_le_bytes()); // xppm
        dib.extend_from_slice(&0i32.to_le_bytes()); // yppm
        dib.extend_from_slice(&0u32.to_le_bytes()); // colors used
        dib.extend_from_slice(&0u32.to_le_bytes()); // important colors
        dib.extend_from_slice(&[0xFF, 0, 0, 0, 0xFF, 0, 0, 0]); // 2 BGR pixels + padding

        let bmp = dib_to_bmp(&dib).expect("converts");
        assert_eq!(&bmp[..2], b"BM");
        let pixel_off = u32::from_le_bytes(bmp[10..14].try_into().unwrap());
        assert_eq!(pixel_off as usize, 14 + 40); // no palette at 24-bpp
        let file_size = u32::from_le_bytes(bmp[2..6].try_into().unwrap());
        assert_eq!(file_size as usize, bmp.len());
    }
}
