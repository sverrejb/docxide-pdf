use std::collections::HashMap;

/// The Unicode code points of WinAnsi (Windows-1252) bytes 0x80..=0x9F, the
/// only bytes that differ from Latin-1; 0 for the five undefined bytes.
const HIGH: [u16; 32] = [
    0x20AC, 0, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, // 0x80..=0x87
    0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0, 0x017D, 0, // 0x88..=0x8F
    0, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, // 0x90..=0x97
    0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0, 0x017E, 0x0178, // 0x98..=0x9F
];

/// Windows-1252 (WinAnsi) byte to Unicode char mapping.
/// Bytes 0x80-0x9F are remapped; all others map directly to their Unicode codepoint.
pub(super) fn winansi_to_char(byte: u8) -> char {
    match byte {
        0x80..=0x9F => char::from_u32(HIGH[(byte - 0x80) as usize] as u32)
            .filter(|&c| c != '\0')
            .unwrap_or(byte as char),
        _ => byte as char,
    }
}

/// Map a single Unicode char to its WinAnsi byte, or 0 if unmappable.
pub(super) fn char_to_winansi(c: char) -> u8 {
    match c as u32 {
        0x0020..=0x007F | 0x00A0..=0x00FF => c as u8,
        u => u16::try_from(u)
            .ok()
            .filter(|&u| u != 0)
            .and_then(|u| HIGH.iter().position(|&h| h == u))
            .map_or(0, |i| 0x80 + i as u8),
    }
}

/// Convert a UTF-8 string to WinAnsi (Windows-1252) bytes for PDF Str encoding.
pub(crate) fn to_winansi_bytes(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(s.len());
    for ch in s.chars() {
        let b = char_to_winansi(ch);
        if b != 0 {
            out.push(b);
        }
    }
    out
}

/// Encode UTF-8 text as big-endian 2-byte glyph IDs for CIDFont content streams.
pub(crate) fn encode_as_gids(text: &str, char_to_gid: &HashMap<char, u16>) -> Vec<u8> {
    encode_as_gids_noting(text, char_to_gid, &std::cell::Cell::new(false))
}

/// `encode_as_gids`, setting `drew_notdef` when a char has no glyph.
pub(super) fn encode_as_gids_noting(
    text: &str,
    char_to_gid: &HashMap<char, u16>,
    drew_notdef: &std::cell::Cell<bool>,
) -> Vec<u8> {
    let mut out = Vec::with_capacity(text.len() * 2);
    for ch in text.chars() {
        let gid = char_to_gid.get(&ch).copied().unwrap_or_else(|| {
            drew_notdef.set(true);
            0
        });
        out.extend_from_slice(&gid.to_be_bytes());
    }
    out
}

/// Approximate Helvetica widths at 1000 units/em for WinAnsi chars 32..=255.
pub(super) fn helvetica_widths() -> Vec<f32> {
    (32u8..=255u8)
        .map(|b| match b {
            32 => 278.0,                          // space
            33..=47 => 333.0,                     // punctuation
            48..=57 => 556.0,                     // digits
            58..=64 => 333.0,                     // more punctuation
            73 | 74 => 278.0,                     // I J (narrow uppercase)
            77 => 833.0,                          // M (wide)
            65..=90 => 667.0,                     // uppercase A-Z (average)
            91..=96 => 333.0,                     // brackets etc.
            102 | 105 | 106 | 108 | 116 => 278.0, // narrow lowercase: f i j l t
            109 | 119 => 833.0,                   // m w (wide)
            97..=122 => 556.0,                    // lowercase a-z (average)
            _ => 556.0,
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn high_bytes_round_trip() {
        for byte in 0x80u8..=0x9F {
            let c = winansi_to_char(byte);
            if HIGH[(byte - 0x80) as usize] == 0 {
                assert_eq!(c, byte as char);
                assert_eq!(char_to_winansi(c), 0);
            } else {
                assert_eq!(char_to_winansi(c), byte);
            }
        }
        assert_eq!(char_to_winansi('\0'), 0);
        assert_eq!(char_to_winansi('\u{2011}'), 0);
        assert_eq!(char_to_winansi('\u{20AC}'), 0x80);
        assert_eq!(char_to_winansi('\u{00E9}'), 0xE9);
    }
}
