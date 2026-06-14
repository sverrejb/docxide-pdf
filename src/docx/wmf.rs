//! Minimal WMF (Windows Metafile) handling.
//!
//! Word frequently embeds bitmap content as an OLE "StaticMetafile" — a WMF
//! container whose actual payload is one or more `META_DIBBITBLT`,
//! `META_DIBSTRETCHBLT`, or `META_STRETCHDIB` records, each carrying a Device-
//! Independent Bitmap positioned at a destination rectangle. For a single
//! full-frame DIB (the simplest case) we just extract it and prepend a 14-byte
//! `BITMAPFILEHEADER` to make a valid BMP. When several DIB records tile the
//! metafile window (e.g. the Australian coat-of-arms emblem, painted as three
//! stacked monochrome bands), we composite them onto a single canvas at their
//! destination rectangles — otherwise only the first band would render,
//! stretched to fill the whole box.
//!
//! True vector WMFs (with `LineTo`, `Polygon`, `ExtTextOut`, etc.) are not
//! handled — `wmf_to_raster` returns `None` and the caller falls back to
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

/// One DIB-blit record: a Device-Independent Bitmap and the destination
/// rectangle (in metafile logical units) it is painted into.
struct DibBlit<'a> {
    dib: &'a [u8],
    x: i32,
    y: i32,
    w: i32,
    h: i32,
}

/// Walk the metafile record stream and collect every DIB-blit record with its
/// destination rectangle. Returns `None` if the stream is malformed; an empty
/// vec means no DIB records (a true vector WMF we don't handle).
fn collect_dib_blits(data: &[u8]) -> Option<Vec<DibBlit<'_>>> {
    let header_skip = if data.starts_with(&PLACEABLE_MAGIC) {
        22 + 18 // placeable header + standard WMF header
    } else if data.starts_with(&STANDARD_MAGIC) {
        18
    } else {
        return None;
    };

    let mut blits = Vec::new();
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
        let params = &data[i + 6..dib_start];
        // INT16 dest rect offsets within the param block differ per record.
        let i16_at = |off: usize| -> i32 {
            i16::from_le_bytes([params[off], params[off + 1]]) as i32
        };
        let (h, w, y, x) = match func {
            // DIBBITBLT: RasterOp(4), YSrc, XSrc, Height, Width, YDest, XDest
            META_DIBBITBLT => (i16_at(8), i16_at(10), i16_at(12), i16_at(14)),
            // DIBSTRETCHBLT: RasterOp(4), SrcH, SrcW, YSrc, XSrc, DestH, DestW, YDest, XDest
            META_DIBSTRETCHBLT => (i16_at(12), i16_at(14), i16_at(16), i16_at(18)),
            // STRETCHDIB: RasterOp(4), Usage, SrcH, SrcW, YSrc, XSrc, DestH, DestW, YDest, XDest
            _ => (i16_at(14), i16_at(16), i16_at(18), i16_at(20)),
        };
        blits.push(DibBlit {
            dib: &data[dib_start..dib_end],
            x,
            y,
            w,
            h,
        });
        i += size_bytes;
    }
    Some(blits)
}

/// Convert a WMF into a raster image (BMP for a single DIB, PNG for a composited
/// multi-DIB metafile). Returns `None` for vector WMFs and malformed input.
pub(super) fn wmf_to_raster(data: &[u8]) -> Option<Vec<u8>> {
    let blits = collect_dib_blits(data)?;
    match blits.as_slice() {
        [] => None,
        // Single DIB: keep the cheap, lossless BMP path (the common case).
        [only] => dib_to_bmp(only.dib),
        many => composite_blits(many),
    }
}

/// Composite multiple DIB bands onto one white canvas at their destination
/// rectangles, returning PNG bytes. Used when a metafile tiles its window with
/// several DIBs (so the whole image renders, not just the first band).
fn composite_blits(blits: &[DibBlit]) -> Option<Vec<u8>> {
    let mut decoded: Vec<(image::RgbaImage, &DibBlit)> = Vec::new();
    let (mut min_x, mut min_y, mut max_x, mut max_y) = (i32::MAX, i32::MAX, i32::MIN, i32::MIN);
    for b in blits {
        if b.w <= 0 || b.h <= 0 {
            continue;
        }
        let bmp = dib_to_bmp(b.dib)?;
        let img = image::load_from_memory_with_format(&bmp, image::ImageFormat::Bmp)
            .ok()?
            .to_rgba8();
        min_x = min_x.min(b.x);
        min_y = min_y.min(b.y);
        max_x = max_x.max(b.x + b.w);
        max_y = max_y.max(b.y + b.h);
        decoded.push((img, b));
    }
    if decoded.is_empty() {
        return None;
    }

    let span_w = (max_x - min_x).max(1) as f32;
    let span_h = (max_y - min_y).max(1) as f32;
    // Scale the canvas so no band is downscaled (preserve the finest DIB's
    // resolution), capped to avoid pathological allocations.
    let mut scale = 1.0_f32;
    for (img, b) in &decoded {
        scale = scale.max(img.width() as f32 / b.w as f32);
        scale = scale.max(img.height() as f32 / b.h as f32);
    }
    scale = scale.clamp(1.0, 4.0);

    let canvas_w = ((span_w * scale).round() as u32).clamp(1, 5000);
    let canvas_h = ((span_h * scale).round() as u32).clamp(1, 5000);
    let mut canvas = image::RgbaImage::from_pixel(canvas_w, canvas_h, image::Rgba([255, 255, 255, 255]));

    for (img, b) in &decoded {
        let dw = ((b.w as f32 * scale).round() as u32).max(1);
        let dh = ((b.h as f32 * scale).round() as u32).max(1);
        let resized = image::imageops::resize(img, dw, dh, image::imageops::FilterType::Triangle);
        let px = ((b.x - min_x) as f32 * scale).round() as i64;
        let py = ((b.y - min_y) as f32 * scale).round() as i64;
        // Opaque bands tile without overlap, so a straight copy reproduces the
        // metafile's SRCCOPY/SRCAND-onto-white compositing.
        image::imageops::replace(&mut canvas, &resized, px, py);
    }

    let mut out = std::io::Cursor::new(Vec::new());
    image::DynamicImage::ImageRgba8(canvas)
        .write_to(&mut out, image::ImageFormat::Png)
        .ok()?;
    Some(out.into_inner())
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
        assert!(wmf_to_raster(b"\x89PNG\r\n\x1a\n").is_none());
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
