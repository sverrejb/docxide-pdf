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

        // The record's no-bitmap form (size = (function >> 8) + 3 words) is a
        // pattern fill, not a picture.
        if size_words == (func >> 8) as usize + 3 {
            i += size_bytes;
            continue;
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
        let i16_at =
            |off: usize| -> i32 { i16::from_le_bytes([params[off], params[off + 1]]) as i32 };
        let (h, w, y, x) = match func {
            // DIBBITBLT: RasterOp(4), YSrc, XSrc, Height, Width, YDest, XDest
            META_DIBBITBLT => (i16_at(8), i16_at(10), i16_at(12), i16_at(14)),
            // DIBSTRETCHBLT: RasterOp(4), SrcH, SrcW, YSrc, XSrc, DestH, DestW, YDest, XDest
            META_DIBSTRETCHBLT => (i16_at(12), i16_at(14), i16_at(16), i16_at(18)),
            // STRETCHDIB: RasterOp(4), Usage, SrcH, SrcW, YSrc, XSrc, DestH, DestW, YDest, XDest
            _ => (i16_at(14), i16_at(16), i16_at(18), i16_at(20)),
        };
        let dib = &data[dib_start..dib_end];
        // A brush-sized DIB stretched over the frame paints a solid box where
        // Word draws the metafile's vector content.
        if dib_pixel_count(dib).is_some_and(|n| n <= 64) {
            i += size_bytes;
            continue;
        }
        blits.push(DibBlit { dib, x, y, w, h });
        i += size_bytes;
    }
    Some(blits)
}

fn dib_pixel_count(dib: &[u8]) -> Option<u64> {
    let header = u32::from_le_bytes(dib.get(0..4)?.try_into().ok()?);
    let (w, h) = if header == 12 {
        let w = u16::from_le_bytes(dib.get(4..6)?.try_into().ok()?) as i64;
        let h = u16::from_le_bytes(dib.get(6..8)?.try_into().ok()?) as i64;
        (w, h)
    } else {
        let w = i32::from_le_bytes(dib.get(4..8)?.try_into().ok()?) as i64;
        let h = i32::from_le_bytes(dib.get(8..12)?.try_into().ok()?) as i64;
        (w, h)
    };
    Some(w.unsigned_abs() * h.unsigned_abs())
}

const META_ESCAPE: u16 = 0x0626;
const MFCOMMENT: u16 = 0x000F;
/// Fixed part of a META_ESCAPE_ENHANCED_METAFILE comment after EscapeFunction
/// and ByteCount: identifier, type, version, checksum, flags, record count,
/// this chunk's size, remaining bytes, total EMF size ([MS-WMF]).
const WMFC_HEADER: usize = 34;

/// The EMF that Word stores alongside a WMF's own records, split over
/// "WMFC" comment escapes. When present it is the full-fidelity picture.
pub(super) fn embedded_emf(data: &[u8]) -> Option<Vec<u8>> {
    let mut i = if data.starts_with(&PLACEABLE_MAGIC) {
        40
    } else {
        18
    };
    let mut emf = Vec::new();
    let mut total = 0usize;
    while i + 6 <= data.len() {
        let size_bytes =
            (u32::from_le_bytes(data[i..i + 4].try_into().ok()?) as usize).checked_mul(2)?;
        let func = u16::from_le_bytes(data[i + 4..i + 6].try_into().ok()?);
        if size_bytes < 6 || i + size_bytes > data.len() {
            break;
        }
        let rec = &data[i + 6..i + size_bytes];
        if func == META_ESCAPE
            && rec.len() >= 4 + WMFC_HEADER
            && u16::from_le_bytes([rec[0], rec[1]]) == MFCOMMENT
            && &rec[4..8] == b"WMFC"
        {
            let u32_at =
                |o: usize| u32::from_le_bytes(rec[4 + o..8 + o].try_into().unwrap()) as usize;
            let chunk = u32_at(22);
            total = u32_at(30);
            let start = 4 + WMFC_HEADER;
            emf.extend_from_slice(rec.get(start..start + chunk)?);
        }
        i += size_bytes;
    }
    (total > 0 && emf.len() == total && super::emf::is_emf(&emf)).then_some(emf)
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
    let mut canvas =
        image::RgbaImage::from_pixel(canvas_w, canvas_h, image::Rgba([255, 255, 255, 255]));

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
    let info_len = bih_size + palette_bytes;
    if info_len > dib.len() {
        return None;
    }
    Some(bmp_from_parts(&dib[..info_len], &dib[info_len..]))
}

/// Prepend a BITMAPFILEHEADER to a DIB given as its BITMAPINFO (header + colour
/// table) and pixel bits. Shared with the EMF EMR_STRETCHDIBITS path.
pub(super) fn bmp_from_parts(info: &[u8], bits: &[u8]) -> Vec<u8> {
    let pixel_off = 14 + info.len();
    let mut bmp = Vec::with_capacity(pixel_off + bits.len());
    bmp.extend_from_slice(b"BM");
    bmp.extend_from_slice(&((pixel_off + bits.len()) as u32).to_le_bytes());
    bmp.extend_from_slice(&[0u8; 4]); // reserved1 + reserved2
    bmp.extend_from_slice(&(pixel_off as u32).to_le_bytes());
    bmp.extend_from_slice(info);
    bmp.extend_from_slice(bits);
    bmp
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
    fn embedded_emf_joins_wmfc_chunks() {
        let mut emf = vec![0u8; 92];
        emf[0..4].copy_from_slice(&1u32.to_le_bytes());
        emf[40..44].copy_from_slice(b" EMF");
        let mut wmf = vec![0u8; 18];
        for chunk in [&emf[..50], &emf[50..]] {
            let mut rec = vec![0x0F, 0x00, 0, 0];
            rec.extend_from_slice(b"WMFC");
            rec.extend_from_slice(&[0u8; 18]);
            rec.extend_from_slice(&(chunk.len() as u32).to_le_bytes());
            rec.extend_from_slice(&0u32.to_le_bytes());
            rec.extend_from_slice(&(emf.len() as u32).to_le_bytes());
            rec.extend_from_slice(chunk);
            wmf.extend_from_slice(&((6 + rec.len()) as u32 / 2).to_le_bytes());
            wmf.extend_from_slice(&META_ESCAPE.to_le_bytes());
            wmf.extend_from_slice(&rec);
        }
        wmf.extend_from_slice(&3u32.to_le_bytes());
        wmf.extend_from_slice(&0u16.to_le_bytes());
        assert_eq!(embedded_emf(&wmf), Some(emf.clone()));
        // A missing chunk leaves the WMF on its own records.
        assert_eq!(embedded_emf(&wmf[..18 + 94]), None);
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

/// Translate a vector WMF (polygons, pens and brushes) into an EMF byte
/// stream for the EMF translator: the stamp in swiss_materials' header is 47
/// `META_POLYPOLYGON`s, which Word rasterises. Returns `None` when no record
/// paints. ponytail: text, arcs and bitmaps in vector WMFs stay undrawn.
pub(super) fn wmf_to_emf(data: &[u8]) -> Option<Vec<u8>> {
    let i16_at = |o: usize| i16::from_le_bytes(data[o..o + 2].try_into().unwrap()) as i32;
    let (skip, bbox, units_per_inch) = if data.starts_with(&PLACEABLE_MAGIC) {
        if data.len() < 40 {
            return None;
        }
        let upi = u16::from_le_bytes(data[14..16].try_into().unwrap()) as i32;
        (40, Some((i16_at(6), i16_at(8), i16_at(10), i16_at(12))), upi.max(1))
    } else if data.starts_with(&STANDARD_MAGIC) {
        (18, None, 96)
    } else {
        return None;
    };

    let mut recs: Vec<Vec<u8>> = Vec::new();
    let rec = |ty: u32, body: &[u32]| -> Vec<u8> {
        let mut v = Vec::with_capacity(8 + body.len() * 4);
        v.extend_from_slice(&ty.to_le_bytes());
        v.extend_from_slice(&((8 + body.len() * 4) as u32).to_le_bytes());
        for w in body {
            v.extend_from_slice(&w.to_le_bytes());
        }
        v
    };
    // WMF objects take the first free slot; EMF handles are numbered from 1.
    let mut slots: Vec<Option<u32>> = Vec::new();
    let mut next_handle = 1u32;
    let mut alloc = |slots: &mut Vec<Option<u32>>| -> u32 {
        let h = next_handle;
        next_handle += 1;
        match slots.iter().position(Option::is_none) {
            Some(i) => slots[i] = Some(h),
            None => slots.push(Some(h)),
        }
        h
    };
    let mut window: (Option<(i32, i32)>, Option<(i32, i32)>) = (None, None);
    let mut painted = false;
    // META_SETROP2 R2_NOP (11) draws nothing: bulgarian_farmland's emblem
    // first lays its silhouette under R2_NOP, then the real figure.
    let mut rop2_nop = false;
    let mut i = skip;
    while i + 6 <= data.len() {
        let size_bytes = (u32::from_le_bytes(data[i..i + 4].try_into().ok()?) as usize).checked_mul(2)?;
        let func = u16::from_le_bytes(data[i + 4..i + 6].try_into().ok()?);
        if size_bytes < 6 || i + size_bytes > data.len() {
            return None;
        }
        let p = &data[i + 6..i + size_bytes];
        let pi = |k: usize| -> Option<i32> {
            p.get(2 * k..2 * k + 2)
                .map(|s| i16::from_le_bytes(s.try_into().unwrap()) as i32)
        };
        let pu32 = |o: usize| -> Option<u32> {
            p.get(o..o + 4).map(|s| u32::from_le_bytes(s.try_into().unwrap()))
        };
        let path = |recs: &mut Vec<Vec<u8>>, polys: &[&[(i32, i32)]], close: bool| {
            recs.push(rec(59, &[]));
            for pts in polys {
                for (k, (x, y)) in pts.iter().enumerate() {
                    recs.push(rec(if k == 0 { 27 } else { 54 }, &[*x as u32, *y as u32]));
                }
                if close {
                    recs.push(rec(61, &[]));
                }
            }
            recs.push(rec(60, &[]));
            recs.push(rec(if close { 63 } else { 64 }, &[0, 0, 0, 0]));
        };
        match func {
            0x0000 => break,
            // META_SETWINDOWORG / EXT: parameters are (y, x).
            0x020b => {
                let (y, x) = (pi(0)?, pi(1)?);
                window.0 = Some((x, y));
                recs.push(rec(10, &[x as u32, y as u32]));
            }
            0x020c => {
                let (y, x) = (pi(0)?, pi(1)?);
                window.1 = Some((x, y));
                recs.push(rec(9, &[x as u32, y as u32]));
            }
            0x0106 => recs.push(rec(19, &[pi(0)? as u32])),
            0x0104 => rop2_nop = pi(0)? == 11,
            0x02fc => {
                let h = alloc(&mut slots);
                recs.push(rec(39, &[h, pi(0)? as u32, pu32(2)?, pi(3)? as u32]));
            }
            0x02fa => {
                let h = alloc(&mut slots);
                // A width-0 pen is a device hairline; Word's export of these
                // pictures draws no outline at all (bulgarian_farmland's
                // emblem lost its white detail under 1pt strokes), so it is null.
                let (style, width) = (pi(0)?, pi(1)?);
                let style = if width == 0 { 5 } else { style };
                recs.push(rec(38, &[h, style as u32, width as u32, pi(2)? as u32, pu32(6)?]));
            }
            // Other object kinds still take a slot.
            0x02fb | 0x0142 | 0x01f9 | 0x0322 | 0x0b23 => {
                alloc(&mut slots);
            }
            0x012d => {
                if let Some(Some(h)) = slots.get(pi(0)? as usize) {
                    recs.push(rec(37, &[*h]));
                }
            }
            0x01f0 => {
                let s = pi(0)? as usize;
                if let Some(Some(h)) = slots.get(s) {
                    recs.push(rec(40, &[*h]));
                    slots[s] = None;
                }
            }
            0x0538 if rop2_nop => {}
            0x0324 | 0x0325 if rop2_nop => {}
            0x0538 => {
                let n = pi(0)? as usize;
                let counts: Vec<usize> = (0..n).map(|k| pi(1 + k).map(|c| c as usize)).collect::<Option<_>>()?;
                let mut at = 1 + n;
                let mut polys: Vec<Vec<(i32, i32)>> = Vec::new();
                for c in counts {
                    let pts = (0..c)
                        .map(|k| Some((pi(at + 2 * k)?, pi(at + 2 * k + 1)?)))
                        .collect::<Option<Vec<_>>>()?;
                    at += 2 * c;
                    polys.push(pts);
                }
                let refs: Vec<&[(i32, i32)]> = polys.iter().map(Vec::as_slice).collect();
                path(&mut recs, &refs, true);
                painted = true;
            }
            0x0324 | 0x0325 => {
                let c = pi(0)? as usize;
                let pts = (0..c)
                    .map(|k| Some((pi(1 + 2 * k)?, pi(2 + 2 * k)?)))
                    .collect::<Option<Vec<_>>>()?;
                path(&mut recs, &[pts.as_slice()], func == 0x0324);
                painted = true;
            }
            _ => {}
        }
        i += size_bytes;
    }
    if !painted {
        return None;
    }
    // Logical units are device units: the placeable box (or the window) is
    // the frame, mapped 1:1.
    let (l, t, r, b) = bbox.or_else(|| {
        let (ox, oy) = window.0.unwrap_or((0, 0));
        let (ex, ey) = window.1?;
        Some((ox, oy, ox + ex, oy + ey))
    })?;
    let (w, h) = (r - l, b - t);
    if w <= 0 || h <= 0 {
        return None;
    }
    // The window maps onto the frame: a 1000-unit window on a 763-unit box
    // scales down (bulgarian_farmland's emblem drew 20x too large without).
    if window.1.is_some() {
        recs.insert(0, rec(11, &[w as u32, h as u32]));
        recs.insert(1, rec(12, &[0, 0]));
    }
    recs.push(rec(14, &[0, 16, 20]));
    let body_len: usize = recs.iter().map(Vec::len).sum();
    let mm = |v: i32| (v as f64 * 25.4 / units_per_inch as f64).round() as i32;
    let header: Vec<u32> = vec![
        1,
        108,
        l as u32,
        t as u32,
        (r - 1) as u32,
        (b - 1) as u32,
        0,
        0,
        (w as f64 * 2540.0 / units_per_inch as f64) as u32,
        (h as f64 * 2540.0 / units_per_inch as f64) as u32,
        0x464D4520,
        0x10000,
        (108 + body_len) as u32,
        (recs.len() + 1) as u32,
        (next_handle & 0xffff) | 0u32,
        0,
        0,
        0,
        w as u32,
        h as u32,
        mm(w).max(1) as u32,
        mm(h).max(1) as u32,
        0,
        0,
        0,
        (mm(w).max(1) * 1000) as u32,
        (mm(h).max(1) * 1000) as u32,
    ];
    let mut out = Vec::with_capacity(108 + body_len);
    for v in header {
        out.extend_from_slice(&v.to_le_bytes());
    }
    for r in recs {
        out.extend_from_slice(&r);
    }
    Some(out)
}
