use std::collections::HashMap;

/// Windows-1252 (WinAnsi) byte to Unicode char mapping.
/// Bytes 0x80-0x9F are remapped; all others map directly to their Unicode codepoint.
pub(super) fn winansi_to_char(byte: u8) -> char {
    match byte {
        0x80 => '\u{20AC}',
        0x82 => '\u{201A}',
        0x83 => '\u{0192}',
        0x84 => '\u{201E}',
        0x85 => '\u{2026}',
        0x86 => '\u{2020}',
        0x87 => '\u{2021}',
        0x88 => '\u{02C6}',
        0x89 => '\u{2030}',
        0x8A => '\u{0160}',
        0x8B => '\u{2039}',
        0x8C => '\u{0152}',
        0x8E => '\u{017D}',
        0x91 => '\u{2018}',
        0x92 => '\u{2019}',
        0x93 => '\u{201C}',
        0x94 => '\u{201D}',
        0x95 => '\u{2022}', // bullet
        0x96 => '\u{2013}',
        0x97 => '\u{2014}',
        0x98 => '\u{02DC}',
        0x99 => '\u{2122}',
        0x9A => '\u{0161}',
        0x9B => '\u{203A}',
        0x9C => '\u{0153}',
        0x9E => '\u{017E}',
        0x9F => '\u{0178}',
        _ => byte as char,
    }
}

/// Map a single Unicode char to its WinAnsi byte, or 0 if unmappable.
pub(super) fn char_to_winansi(c: char) -> u8 {
    match c as u32 {
        0x0020..=0x007F => c as u8,
        0x00A0..=0x00FF => c as u8,
        0x20AC => 0x80,
        0x201A => 0x82,
        0x0192 => 0x83,
        0x201E => 0x84,
        0x2026 => 0x85,
        0x2020 => 0x86,
        0x2021 => 0x87,
        0x02C6 => 0x88,
        0x2030 => 0x89,
        0x0160 => 0x8A,
        0x2039 => 0x8B,
        0x0152 => 0x8C,
        0x017D => 0x8E,
        0x2018 => 0x91,
        0x2019 => 0x92,
        0x201C => 0x93,
        0x201D => 0x94,
        0x2022 => 0x95,
        0x2013 => 0x96,
        0x2014 => 0x97,
        0x02DC => 0x98,
        0x2122 => 0x99,
        0x0161 => 0x9A,
        0x203A => 0x9B,
        0x0153 => 0x9C,
        0x017E => 0x9E,
        0x0178 => 0x9F,
        _ => 0,
    }
}

/// Convert a UTF-8 string to WinAnsi (Windows-1252) bytes for PDF Str encoding.
pub(crate) fn to_winansi_bytes(s: &str) -> Vec<u8> {
    s.chars().map(char_to_winansi).filter(|&b| b != 0).collect()
}

/// Encode UTF-8 text as big-endian 2-byte glyph IDs for CIDFont content streams.
pub(crate) fn encode_as_gids(text: &str, char_to_gid: &HashMap<char, u16>) -> Vec<u8> {
    let mut out = Vec::with_capacity(text.len() * 2);
    for ch in text.chars() {
        let gid = char_to_gid.get(&ch).copied().unwrap_or(0);
        out.push((gid >> 8) as u8);
        out.push((gid & 0xFF) as u8);
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
