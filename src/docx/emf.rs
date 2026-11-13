//! Minimal EMF (Enhanced Metafile) parsing.
//!
//! Walks the 8-byte-header record stream and decodes the subset of records
//! that real-world DOCX EMF logos use (paths + simple objects + viewport
//! transforms). Unsupported records are reported as `Skip` so the caller can
//! keep walking — never panic on a record we don't know.
//!
//! Binary layout reference: [MS-EMF].

use std::convert::TryInto;

const EMF_MAGIC: [u8; 4] = [0x20, 0x45, 0x4D, 0x46]; // " EMF"

/// EMR_HEADER fields (subset relevant to rendering).
#[derive(Debug, Clone, Copy)]
pub(crate) struct EmfHeader {
    /// Inclusive logical bounds of all drawing in device coordinates.
    pub bounds: (i32, i32, i32, i32),
    /// The picture frame (rclFrame, 0.01mm) in device pixels: the rectangle
    /// GDI maps onto the picture box. Ink outside it is clipped, and a frame
    /// larger than the ink leaves margins.
    pub frame: (f64, f64, f64, f64),
}

impl EmfHeader {
    pub(crate) fn bounds_size(&self) -> (i32, i32) {
        (self.bounds.2 - self.bounds.0, self.bounds.3 - self.bounds.1)
    }
}

pub(crate) fn is_emf(data: &[u8]) -> bool {
    data.len() >= 44
        && u32::from_le_bytes(data[0..4].try_into().unwrap()) == 1
        && data[40..44] == EMF_MAGIC
}

pub(crate) fn parse_header(data: &[u8]) -> Option<EmfHeader> {
    if !is_emf(data) || data.len() < 92 {
        return None;
    }
    let i32_at = |off: usize| i32::from_le_bytes(data[off..off + 4].try_into().unwrap());
    let bounds = (i32_at(8), i32_at(12), i32_at(16), i32_at(20));
    // Device pixels per 0.01mm from szlDevice / szlMillimeters.
    let (px_w, px_h, mm_w, mm_h) = (i32_at(72), i32_at(76), i32_at(80), i32_at(84));
    let frame = if px_w > 0 && px_h > 0 && mm_w > 0 && mm_h > 0 && i32_at(32) > i32_at(24) {
        let sx = px_w as f64 / (mm_w as f64 * 100.0);
        let sy = px_h as f64 / (mm_h as f64 * 100.0);
        (
            i32_at(24) as f64 * sx,
            i32_at(28) as f64 * sy,
            i32_at(32) as f64 * sx,
            i32_at(36) as f64 * sy,
        )
    } else {
        let (l, t, r, b) = bounds;
        (l as f64, t as f64, r as f64, b as f64)
    };
    Some(EmfHeader { bounds, frame })
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum FillRule {
    Alternate, // ALTERNATE = 1
    Winding,   // WINDING = 2
}

/// A subset of EMF records — the ones the translator can render. Other record
/// types come through as `Skip`.
#[derive(Debug, Clone)]
pub(crate) enum EmfRecord {
    Header,
    Eof,
    SaveDc,
    RestoreDc,
    SetMapMode,
    SetBkMode,
    SetPolyFillMode(FillRule),
    SetWindowExtEx(i32, i32),
    SetWindowOrgEx(i32, i32),
    SetViewportExtEx(i32, i32),
    SetViewportOrgEx(i32, i32),
    MoveToEx(i32, i32),
    LineTo(i32, i32),
    /// Cubic Bezier in 16-bit point form, continuing from the current point.
    /// Coordinates are flat triples `(ctl1, ctl2, end)`.
    PolyBezierTo16(Vec<(i16, i16)>),
    PolyLineTo16(Vec<(i16, i16)>),
    /// EMR_POLYGON / POLYLINE / POLYPOLYGON / POLYPOLYLINE, 16- or 32-bit:
    /// figures that do not use or move the current point.
    Poly {
        polys: Vec<Vec<(i32, i32)>>,
        close: bool,
    },
    BeginPath,
    EndPath,
    CloseFigure,
    /// Fill the current path using the current brush + fill rule.
    FillPath,
    StrokePath,
    StrokeAndFillPath,
    /// Make the current path the clip region (EMR_SELECTCLIPPATH).
    SelectClipPath,
    /// Discard the current path (EMR_ABORTPATH).
    AbortPath,
    /// A brush; `color` is `None` for BS_NULL.
    CreateBrushIndirect {
        handle: u32,
        color: Option<[u8; 3]>,
    },
    /// EMR_CREATEPEN or EMR_EXTCREATEPEN; `color` is `None` for PS_NULL.
    CreatePen {
        handle: u32,
        width: i32,
        color: Option<[u8; 3]>,
    },
    CreateFont {
        handle: u32,
        /// LOGFONT height: negative is the em size, positive the cell height.
        height: i32,
        /// Text angle, tenths of a degree counter-clockwise.
        escapement: i32,
        bold: bool,
        italic: bool,
        face: String,
    },
    SetTextColor([u8; 3]),
    SetTextAlign(u32),
    /// EMR_EXTTEXTOUTW: reference point, text and per-character advances.
    ExtTextOut {
        x: i32,
        y: i32,
        text: String,
        dx: Vec<i32>,
    },
    /// EMR_BITBLT without a source bitmap: fill the rectangle per `rop`.
    PatBlt {
        x: i32,
        y: i32,
        w: i32,
        h: i32,
        rop: u32,
    },
    Rectangle(i32, i32, i32, i32),
    SelectObject(u32),
    DeleteObject(u32),
    /// Any record we don't decode — `(record_type, payload_bytes)`.
    Skip,
}

/// Word wraps pasted bitmaps (scanned signatures, stamps, logos) in an EMF
/// whose only drawing record is one EMR_STRETCHDIBITS, EMR_STRETCHBLT or
/// EMR_BITBLT. Return that bitmap as a BMP so the
/// raster pipeline can embed it; `None` for anything else, including EMFs that
/// mix vector drawing with a bitmap (those keep the vector translator).
// ponytail: the DIB is stretched over the whole picture frame, ignoring its
// destination rectangle; placing it as an image XObject inside the translated
// form is the general fix for inset and mixed bitmaps.
pub(crate) fn emf_to_raster(data: &[u8]) -> Option<Vec<u8>> {
    if !is_emf(data) {
        return None;
    }
    let mut bmp = None;
    let mut vector = false;
    for_each_raw_record(data, |rec_type, payload| {
        match rec_type {
            81 if bmp.is_none() => bmp = decode_stretchdibits(payload),
            76 | 77 if bmp.is_none() => bmp = decode_dib_at(payload, 76),
            t if paints_vector(t) => vector = true,
            _ => {}
        }
        !vector
    });
    if vector { None } else { bmp }
}

/// Records that paint vector geometry: polygon primitives, shapes, LINETO,
/// path painting and text ([MS-EMF] §2.3.5).
fn paints_vector(rec_type: u32) -> bool {
    matches!(rec_type, 2..=8 | 42..=47 | 54 | 62..=64 | 83..=92)
}

/// The text an EMF draws with the font selected for it: `(face, bold, italic, text)`.
pub(crate) fn text_with_fonts(data: &[u8]) -> Vec<(String, bool, bool, String)> {
    let mut fonts = std::collections::HashMap::new();
    let mut current = None;
    let mut out = Vec::new();
    for_each_record(data, |rec| {
        match rec {
            EmfRecord::CreateFont {
                handle,
                bold,
                italic,
                face,
                ..
            } => {
                fonts.insert(*handle, (face.clone(), *bold, *italic));
            }
            EmfRecord::SelectObject(h) => {
                if let Some(f) = fonts.get(h) {
                    current = Some(f.clone());
                }
            }
            EmfRecord::ExtTextOut { text, .. } => {
                if let Some((face, bold, italic)) = &current {
                    out.push((face.clone(), *bold, *italic, text.clone()));
                }
            }
            _ => {}
        }
        true
    });
    out
}

/// Walk all records after the header, calling `f` for each. Stops at EOF, on
/// a malformed record (zero/negative size), or when `f` returns `false`.
pub(crate) fn for_each_record(data: &[u8], mut f: impl FnMut(&EmfRecord) -> bool) {
    for_each_raw_record(data, |rec_type, payload| f(&decode(rec_type, payload)));
}

/// Same walk without decoding: `f(record_type, payload)`.
fn for_each_raw_record(data: &[u8], mut f: impl FnMut(u32, &[u8]) -> bool) {
    let mut i = 0usize;
    while i + 8 <= data.len() {
        let rec_type = u32::from_le_bytes(data[i..i + 4].try_into().unwrap());
        let rec_size = u32::from_le_bytes(data[i + 4..i + 8].try_into().unwrap()) as usize;
        if rec_size < 8 || i + rec_size > data.len() {
            return;
        }
        if !f(rec_type, &data[i + 8..i + rec_size]) || rec_type == 14 {
            return;
        }
        i += rec_size;
    }
}

fn decode(rec_type: u32, payload: &[u8]) -> EmfRecord {
    use EmfRecord::*;
    let i32_at = |off: usize| -> Option<i32> {
        payload
            .get(off..off + 4)
            .map(|s| i32::from_le_bytes(s.try_into().unwrap()))
    };
    let u32_at = |off: usize| -> Option<u32> {
        payload
            .get(off..off + 4)
            .map(|s| u32::from_le_bytes(s.try_into().unwrap()))
    };
    match rec_type {
        1 => Header,
        14 => Eof,
        33 => SaveDc,
        34 => RestoreDc,
        17 => SetMapMode,
        18 => SetBkMode,
        19 => match u32_at(0).unwrap_or(0) {
            1 => SetPolyFillMode(FillRule::Alternate),
            _ => SetPolyFillMode(FillRule::Winding),
        },
        9 => match (i32_at(0), i32_at(4)) {
            (Some(x), Some(y)) => SetWindowExtEx(x, y),
            _ => Skip,
        },
        10 => match (i32_at(0), i32_at(4)) {
            (Some(x), Some(y)) => SetWindowOrgEx(x, y),
            _ => Skip,
        },
        11 => match (i32_at(0), i32_at(4)) {
            (Some(x), Some(y)) => SetViewportExtEx(x, y),
            _ => Skip,
        },
        12 => match (i32_at(0), i32_at(4)) {
            (Some(x), Some(y)) => SetViewportOrgEx(x, y),
            _ => Skip,
        },
        27 => match (i32_at(0), i32_at(4)) {
            (Some(x), Some(y)) => MoveToEx(x, y),
            _ => Skip,
        },
        54 => match (i32_at(0), i32_at(4)) {
            (Some(x), Some(y)) => LineTo(x, y),
            _ => Skip,
        },
        88 => decode_polybezier16(payload)
            .map(PolyBezierTo16)
            .unwrap_or(Skip),
        89 => decode_polybezier16(payload)
            .map(PolyLineTo16)
            .unwrap_or(Skip),
        3 | 4 | 86 | 87 => decode_poly(payload, false, matches!(rec_type, 86 | 87))
            .map(|polys| Poly { polys, close: matches!(rec_type, 3 | 86) })
            .unwrap_or(Skip),
        7 | 8 | 90 | 91 => decode_poly(payload, true, matches!(rec_type, 90 | 91))
            .map(|polys| Poly { polys, close: matches!(rec_type, 8 | 91) })
            .unwrap_or(Skip),
        59 => BeginPath,
        60 => EndPath,
        61 => CloseFigure,
        62 => FillPath,
        63 => StrokeAndFillPath,
        64 => StrokePath,
        67 => SelectClipPath,
        68 => AbortPath,
        37 => SelectObject(u32_at(0).unwrap_or(0)),
        40 => DeleteObject(u32_at(0).unwrap_or(0)),
        39 => decode_brush(payload).unwrap_or(Skip),
        38 => decode_createpen(payload).unwrap_or(Skip),
        95 => decode_extcreatepen(payload).unwrap_or(Skip),
        82 => decode_font(payload).unwrap_or(Skip),
        84 => decode_exttextoutw(payload).unwrap_or(Skip),
        24 => u32_at(0).map_or(Skip, |c| SetTextColor(colorref(c))),
        22 => u32_at(0).map_or(Skip, SetTextAlign),
        43 => match (i32_at(0), i32_at(4), i32_at(8), i32_at(12)) {
            (Some(l), Some(t), Some(r), Some(b)) => Rectangle(l, t, r, b),
            _ => Skip,
        },
        // cbBitsSrc at 88: a BITBLT with a bitmap source is left out.
        76 if u32_at(88) == Some(0) => match (i32_at(16), i32_at(20), i32_at(24), i32_at(28)) {
            (Some(x), Some(y), Some(w), Some(h)) => PatBlt {
                x,
                y,
                w,
                h,
                rop: u32_at(32).unwrap_or(0),
            },
            _ => Skip,
        },
        _ => Skip,
    }
}

/// EMR_STRETCHDIBITS: the BITMAPINFO (`offBmiSrc`/`cbBmiSrc`) and pixel bits
/// (`offBitsSrc`/`cbBitsSrc`) are addressed from the record start, i.e. 8 bytes
/// before the payload. Returns them wrapped as a BMP file.
fn decode_stretchdibits(payload: &[u8]) -> Option<Vec<u8>> {
    decode_dib_at(payload, 40)
}

/// The offBmiSrc/cbBmiSrc/offBitsSrc/cbBitsSrc fields of a bitmap record,
/// starting at payload offset `fields`: 40 for STRETCHDIBITS, 76 for
/// EMR_BITBLT / EMR_STRETCHBLT, whose source transform and colour come first
/// ([MS-EMF] 2.3.1.2 and 2.3.1.6).
fn decode_dib_at(payload: &[u8], fields: usize) -> Option<Vec<u8>> {
    let u32_at = |off: usize| -> Option<usize> {
        payload
            .get(off..off + 4)
            .map(|s| u32::from_le_bytes(s.try_into().unwrap()) as usize)
    };
    let part = |off: usize, len: usize| -> Option<&[u8]> {
        let start = off.checked_sub(8)?;
        payload.get(start..start.checked_add(len)?)
    };
    let bmi = part(u32_at(fields)?, u32_at(fields + 4)?)?;
    let bits = part(u32_at(fields + 8)?, u32_at(fields + 12)?)?;
    if bmi.len() < 40 || bits.is_empty() {
        return None;
    }
    Some(super::wmf::bmp_from_parts(bmi, bits))
}

/// Bounds RectL, then either a point count and the points, or (`multi`) a
/// figure count, the total point count, each figure's count and the points;
/// points are POINTS (i16 pairs) when `short`, else POINTL.
fn decode_poly(payload: &[u8], multi: bool, short: bool) -> Option<Vec<Vec<(i32, i32)>>> {
    let u32_at = |o: usize| payload.get(o..o + 4).map(|s| u32::from_le_bytes(s.try_into().unwrap()) as usize);
    let (counts, mut at) = if multi {
        let n = u32_at(16)?;
        let counts = (0..n).map(|k| u32_at(24 + 4 * k)).collect::<Option<Vec<_>>>()?;
        (counts, 24 + 4 * n)
    } else {
        (vec![u32_at(16)?], 20)
    };
    let size = if short { 2 } else { 4 };
    let coord = |o: usize| -> Option<i32> {
        let b = payload.get(o..o + size)?;
        Some(if short {
            i16::from_le_bytes(b.try_into().unwrap()) as i32
        } else {
            i32::from_le_bytes(b.try_into().unwrap())
        })
    };
    let mut polys = Vec::with_capacity(counts.len());
    for c in counts {
        let pts = (0..c)
            .map(|k| Some((coord(at + 2 * size * k)?, coord(at + 2 * size * k + size)?)))
            .collect::<Option<Vec<_>>>()?;
        at += 2 * size * c;
        polys.push(pts);
    }
    Some(polys)
}

fn decode_polybezier16(payload: &[u8]) -> Option<Vec<(i16, i16)>> {
    // Payload: bounds RectL (16 bytes), count u32, then `count` POINTS (2x i16 each).
    if payload.len() < 20 {
        return None;
    }
    let count = u32::from_le_bytes(payload[16..20].try_into().unwrap()) as usize;
    let pts_start = 20;
    let need = pts_start + count * 4;
    if payload.len() < need {
        return None;
    }
    let mut pts = Vec::with_capacity(count);
    for k in 0..count {
        let off = pts_start + k * 4;
        let x = i16::from_le_bytes(payload[off..off + 2].try_into().unwrap());
        let y = i16::from_le_bytes(payload[off + 2..off + 4].try_into().unwrap());
        pts.push((x, y));
    }
    Some(pts)
}

/// EMF COLORREF packs color as 0x00BBGGRR.
fn colorref(v: u32) -> [u8; 3] {
    [
        (v & 0xFF) as u8,
        ((v >> 8) & 0xFF) as u8,
        ((v >> 16) & 0xFF) as u8,
    ]
}

fn decode_brush(payload: &[u8]) -> Option<EmfRecord> {
    // EMR_CREATEBRUSHINDIRECT: ihBrush u32, LogBrush32 { style u32, color COLORREF, hatch u32 }
    if payload.len() < 16 {
        return None;
    }
    let handle = u32::from_le_bytes(payload[0..4].try_into().unwrap());
    let style = u32::from_le_bytes(payload[4..8].try_into().unwrap());
    let color = colorref(u32::from_le_bytes(payload[8..12].try_into().unwrap()));
    Some(EmfRecord::CreateBrushIndirect {
        handle,
        color: (style != 1).then_some(color), // BS_NULL
    })
}

/// PS_NULL in the low bits of a pen style draws nothing.
fn pen_color(style: u32, color: u32) -> Option<[u8; 3]> {
    (style & 0xF != 5).then(|| colorref(color))
}

fn decode_createpen(payload: &[u8]) -> Option<EmfRecord> {
    // EMR_CREATEPEN: ihPen u32, LogPen { PenStyle u32, Width PointL (x used), Color }
    if payload.len() < 20 {
        return None;
    }
    let u32_at = |o: usize| u32::from_le_bytes(payload[o..o + 4].try_into().unwrap());
    Some(EmfRecord::CreatePen {
        handle: u32_at(0),
        width: u32_at(8) as i32,
        color: pen_color(u32_at(4), u32_at(16)),
    })
}

fn decode_font(payload: &[u8]) -> Option<EmfRecord> {
    // EMR_EXTCREATEFONTINDIRECTW: ihFont u32, LogFont { Height, Width, Escapement,
    // Orientation, Weight (i32 each), Italic u8, …, FaceName [u16; 32] at 28 }
    if payload.len() < 96 {
        return None;
    }
    let i32_at = |o: usize| i32::from_le_bytes(payload[o..o + 4].try_into().unwrap());
    let face: Vec<u16> = payload[32..96]
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .take_while(|&c| c != 0)
        .collect();
    Some(EmfRecord::CreateFont {
        handle: i32_at(0) as u32,
        height: i32_at(4),
        escapement: i32_at(12),
        bold: i32_at(20) >= 600,
        italic: payload[24] != 0,
        face: String::from_utf16_lossy(&face),
    })
}

fn decode_exttextoutw(payload: &[u8]) -> Option<EmfRecord> {
    // Bounds RectL, iGraphicsMode, exScale, eyScale (28 bytes), then EmrText:
    // Reference PointL, Chars u32, offString u32, Options u32, Rectangle RectL,
    // offDx u32. Offsets count from the record start, 8 bytes before the payload.
    let u32_at = |o: usize| -> Option<u32> {
        payload
            .get(o..o + 4)
            .map(|s| u32::from_le_bytes(s.try_into().unwrap()))
    };
    let x = u32_at(28)? as i32;
    let y = u32_at(32)? as i32;
    let chars = u32_at(36)? as usize;
    let off_string = (u32_at(40)? as usize).checked_sub(8)?;
    let options = u32_at(44)?;
    let units: Vec<u16> = payload
        .get(off_string..off_string + chars * 2)?
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    // ETO_PDY: the array holds an (x, y) pair per character.
    let stride = if options & 0x2000 != 0 { 2 } else { 1 };
    let dx = match u32_at(64).filter(|&o| o >= 8) {
        Some(off) => (0..chars)
            .map(|k| u32_at(off as usize - 8 + k * stride * 4).map(|v| v as i32))
            .collect::<Option<Vec<_>>>()
            .unwrap_or_default(),
        None => Vec::new(),
    };
    Some(EmfRecord::ExtTextOut {
        x,
        y,
        text: String::from_utf16_lossy(&units),
        dx,
    })
}

fn decode_extcreatepen(payload: &[u8]) -> Option<EmfRecord> {
    // EMR_EXTCREATEPEN: ihPen u32, offBmi u32, cbBmi u32, offBits u32, cbBits u32,
    //                   elp: { PenStyle u32, Width u32, BrushStyle u32, Color COLORREF, ... }
    // elp starts at payload offset 20.
    if payload.len() < 36 {
        return None;
    }
    let u32_at = |o: usize| u32::from_le_bytes(payload[o..o + 4].try_into().unwrap());
    Some(EmfRecord::CreatePen {
        handle: u32_at(0),
        width: u32_at(24) as i32,
        color: pen_color(u32_at(20), u32_at(32)),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn polygons_and_polypolygons_decode_their_points() {
        let le = |v: &[u32]| v.iter().flat_map(|w| w.to_le_bytes()).collect::<Vec<u8>>();
        let shorts = |v: &[i16]| v.iter().flat_map(|w| w.to_le_bytes()).collect::<Vec<u8>>();
        // EMR_POLYGON16: bounds, count 3, three POINTS.
        let mut poly16 = le(&[0, 0, 0, 0, 3]);
        poly16.extend(shorts(&[1, 2, 3, 4, 5, 6]));
        match decode(86, &poly16) {
            EmfRecord::Poly { polys, close } => {
                assert!(close);
                assert_eq!(polys, vec![vec![(1, 2), (3, 4), (5, 6)]]);
            }
            _ => panic!("POLYGON16 not decoded"),
        }
        // EMR_POLYPOLYLINE (32-bit): two figures of two points.
        let polypoly = le(&[0, 0, 0, 0, 2, 4, 2, 2, 1, 1, 2, 2, 7, 7, 8, 8]);
        match decode(7, &polypoly) {
            EmfRecord::Poly { polys, close } => {
                assert!(!close);
                assert_eq!(polys, vec![vec![(1, 1), (2, 2)], vec![(7, 7), (8, 8)]]);
            }
            _ => panic!("POLYPOLYLINE not decoded"),
        }
    }

    #[test]
    fn rejects_non_emf() {
        assert!(!is_emf(b"\x89PNG\r\n\x1a\n"));
        assert!(!is_emf(b""));
        assert!(parse_header(b"\x89PNG\r\n\x1a\n").is_none());
    }

    #[test]
    fn parses_header_from_real_emf() {
        let mut hdr = vec![0u8; 92];
        // type=1
        hdr[0..4].copy_from_slice(&1u32.to_le_bytes());
        // size = 92
        hdr[4..8].copy_from_slice(&92u32.to_le_bytes());
        // bounds: 100, 200, 300, 400
        hdr[8..12].copy_from_slice(&100i32.to_le_bytes());
        hdr[12..16].copy_from_slice(&200i32.to_le_bytes());
        hdr[16..20].copy_from_slice(&300i32.to_le_bytes());
        hdr[20..24].copy_from_slice(&400i32.to_le_bytes());
        // frame: 1000..4000
        hdr[24..28].copy_from_slice(&1000i32.to_le_bytes());
        hdr[28..32].copy_from_slice(&2000i32.to_le_bytes());
        hdr[32..36].copy_from_slice(&3000i32.to_le_bytes());
        hdr[36..40].copy_from_slice(&4000i32.to_le_bytes());
        // signature " EMF"
        hdr[40..44].copy_from_slice(&EMF_MAGIC);
        // device
        hdr[72..76].copy_from_slice(&1024i32.to_le_bytes());
        hdr[76..80].copy_from_slice(&768i32.to_le_bytes());
        hdr[80..84].copy_from_slice(&320i32.to_le_bytes());
        hdr[84..88].copy_from_slice(&240i32.to_le_bytes());

        let h = parse_header(&hdr).expect("parses");
        assert_eq!(h.bounds, (100, 200, 300, 400));
        assert_eq!(h.bounds_size(), (200, 200));
        // 1024 px over 320 mm: 0.032 px per 0.01mm.
        assert_eq!(h.frame, (32.0, 64.0, 96.0, 128.0));
    }

    #[test]
    fn decodes_clip_path_records() {
        assert!(matches!(
            decode(67, &5u32.to_le_bytes()),
            EmfRecord::SelectClipPath
        ));
        assert!(matches!(decode(68, &[]), EmfRecord::AbortPath));
    }

    #[test]
    fn bitmap_emf_becomes_bmp() {
        // Header + one STRETCHDIBITS carrying a 1x1 24-bpp DIB.
        let mut data = vec![0u8; 108];
        data[0..4].copy_from_slice(&1u32.to_le_bytes());
        data[4..8].copy_from_slice(&108u32.to_le_bytes());
        data[40..44].copy_from_slice(&EMF_MAGIC);

        let mut bih = vec![0u8; 40];
        bih[0..4].copy_from_slice(&40u32.to_le_bytes());
        bih[4..8].copy_from_slice(&1i32.to_le_bytes()); // width
        bih[8..12].copy_from_slice(&1i32.to_le_bytes()); // height
        bih[14..16].copy_from_slice(&24u16.to_le_bytes()); // bpp
        let bits = [0x10u8, 0x20, 0x30, 0x00];

        data.extend_from_slice(&81u32.to_le_bytes());
        data.extend_from_slice(&(80u32 + 40 + 4).to_le_bytes()); // record size
        data.extend_from_slice(&[0u8; 40]); // bounds, dest/src origin, src size
        for v in [80u32, 40, 120, 4] {
            data.extend_from_slice(&v.to_le_bytes()); // offBmi, cbBmi, offBits, cbBits
        }
        data.extend_from_slice(&[0u8; 16]); // usage, rop, dest size
        data.extend_from_slice(&bih);
        data.extend_from_slice(&bits);

        let bmp = emf_to_raster(&data).expect("bitmap EMF converts");
        assert_eq!(&bmp[54..], &bits);
        let (w, h, fmt, _) = super::super::images::image_dimensions(&bmp).unwrap();
        assert_eq!((w, h, fmt), (1, 1, crate::model::ImageFormat::Bmp));

        // Vector drawing (here FILLPATH) keeps the EMF on the translator path.
        let mut mixed = data.clone();
        mixed.extend_from_slice(&62u32.to_le_bytes());
        mixed.extend_from_slice(&8u32.to_le_bytes());
        assert!(emf_to_raster(&mixed).is_none());
        assert!(emf_to_raster(&data[..108]).is_none());
    }

    #[test]
    fn decodes_exttextoutw() {
        // 28 bytes of bounds/mode/scales, EmrText (40 bytes), then "Ab" and its Dx.
        let mut p = vec![0u8; 68];
        p[28..32].copy_from_slice(&3i32.to_le_bytes()); // reference x
        p[32..36].copy_from_slice(&24i32.to_le_bytes()); // reference y
        p[36..40].copy_from_slice(&2u32.to_le_bytes()); // chars
        p[40..44].copy_from_slice(&(8u32 + 68).to_le_bytes()); // offString
        p[64..68].copy_from_slice(&(8u32 + 72).to_le_bytes()); // offDx
        p.extend_from_slice(&[b'A', 0, b'b', 0]);
        p.extend_from_slice(&7i32.to_le_bytes());
        p.extend_from_slice(&8i32.to_le_bytes());
        match decode(84, &p) {
            EmfRecord::ExtTextOut { x, y, text, dx } => {
                assert_eq!((x, y, text.as_str(), dx), (3, 24, "Ab", vec![7, 8]));
            }
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn for_each_record_stops_at_eof() {
        let mut data = vec![0u8; 92];
        data[0..4].copy_from_slice(&1u32.to_le_bytes()); // header
        data[4..8].copy_from_slice(&92u32.to_le_bytes());
        data[40..44].copy_from_slice(&EMF_MAGIC);
        // EOF record: type=14, size=8 (minimum)
        data.extend_from_slice(&14u32.to_le_bytes());
        data.extend_from_slice(&8u32.to_le_bytes());

        let mut count = 0;
        for_each_record(&data, |_| {
            count += 1;
            true
        });
        assert_eq!(count, 2); // header + eof
    }
}
