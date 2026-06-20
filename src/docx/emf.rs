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
    /// Physical frame size in hundredths of a millimetre.
    pub frame: (i32, i32, i32, i32),
    /// Device size in pixels.
    pub device_px: (i32, i32),
    /// Device size in millimetres.
    pub device_mm: (i32, i32),
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
    Some(EmfHeader {
        bounds: (i32_at(8), i32_at(12), i32_at(16), i32_at(20)),
        frame: (i32_at(24), i32_at(28), i32_at(32), i32_at(36)),
        device_px: (i32_at(72), i32_at(76)),
        device_mm: (i32_at(80), i32_at(84)),
    })
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
    SetMapMode(u32),
    SetBkMode(u32),
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
    BeginPath,
    EndPath,
    CloseFigure,
    /// Fill the current path using the current brush + fill rule.
    FillPath,
    StrokePath,
    StrokeAndFillPath,
    CreateBrushIndirect { handle: u32, color: [u8; 3] },
    ExtCreatePen { handle: u32, width: i32, color: [u8; 3] },
    SelectObject(u32),
    DeleteObject(u32),
    /// Any record we don't decode — `(record_type, payload_bytes)`.
    Skip,
}

/// Walk all records after the header, calling `f` for each. Stops at EOF, on
/// a malformed record (zero/negative size), or when `f` returns `false`.
pub(crate) fn for_each_record(data: &[u8], mut f: impl FnMut(&EmfRecord) -> bool) {
    let mut i = 0usize;
    while i + 8 <= data.len() {
        let rec_type = u32::from_le_bytes(data[i..i + 4].try_into().unwrap());
        let rec_size = u32::from_le_bytes(data[i + 4..i + 8].try_into().unwrap()) as usize;
        if rec_size < 8 || i + rec_size > data.len() {
            return;
        }
        let payload = &data[i + 8..i + rec_size];
        let record = decode(rec_type, payload);
        let cont = f(&record);
        if matches!(record, EmfRecord::Eof) || !cont {
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
        17 => SetMapMode(u32_at(0).unwrap_or(0)),
        18 => SetBkMode(u32_at(0).unwrap_or(0)),
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
        88 => decode_polybezier16(payload).map(PolyBezierTo16).unwrap_or(Skip),
        89 => decode_polybezier16(payload).map(PolyLineTo16).unwrap_or(Skip),
        59 => BeginPath,
        60 => EndPath,
        61 => CloseFigure,
        62 => FillPath,
        63 => StrokeAndFillPath,
        64 => StrokePath,
        37 => SelectObject(u32_at(0).unwrap_or(0)),
        40 => DeleteObject(u32_at(0).unwrap_or(0)),
        39 => decode_brush(payload).unwrap_or(Skip),
        95 => decode_extcreatepen(payload).unwrap_or(Skip),
        _ => Skip,
    }
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
    [(v & 0xFF) as u8, ((v >> 8) & 0xFF) as u8, ((v >> 16) & 0xFF) as u8]
}

fn decode_brush(payload: &[u8]) -> Option<EmfRecord> {
    // EMR_CREATEBRUSHINDIRECT: ihBrush u32, LogBrush32 { style u32, color COLORREF, hatch u32 }
    if payload.len() < 16 {
        return None;
    }
    let handle = u32::from_le_bytes(payload[0..4].try_into().unwrap());
    let color = colorref(u32::from_le_bytes(payload[8..12].try_into().unwrap()));
    Some(EmfRecord::CreateBrushIndirect { handle, color })
}

fn decode_extcreatepen(payload: &[u8]) -> Option<EmfRecord> {
    // EMR_EXTCREATEPEN: ihPen u32, offBmi u32, cbBmi u32, offBits u32, cbBits u32,
    //                   elp: { PenStyle u32, Width u32, BrushStyle u32, Color COLORREF, ... }
    // elp starts at payload offset 20.
    if payload.len() < 36 {
        return None;
    }
    let handle = u32::from_le_bytes(payload[0..4].try_into().unwrap());
    let width = u32::from_le_bytes(payload[24..28].try_into().unwrap()) as i32;
    let color = colorref(u32::from_le_bytes(payload[32..36].try_into().unwrap()));
    Some(EmfRecord::ExtCreatePen { handle, width, color })
}

#[cfg(test)]
mod tests {
    use super::*;

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
        assert_eq!(h.device_px, (1024, 768));
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
