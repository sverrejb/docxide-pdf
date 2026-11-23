//! A fast, deterministic hasher for the small keys layout looks up per
//! character (chars, glyph ids, font names). std's SipHash is DoS-resistant
//! but costs more than the lookup itself on these keys: char widths, glyph
//! ids and kerning pairs are read for every glyph of every line. This is the
//! Fx hash (rustc's), with the final rotation of rustc-hash 2 so the bucket
//! index (the low bits) depends on every input bit, not only the low ones.
//!
//! Nothing may depend on these maps' iteration order: std's maps already
//! iterate in a random order per process.

use std::collections::{HashMap, HashSet};
use std::hash::{BuildHasherDefault, Hasher};

pub(crate) type FastMap<K, V> = HashMap<K, V, BuildHasherDefault<FxHasher>>;
pub(crate) type FastSet<T> = HashSet<T, BuildHasherDefault<FxHasher>>;

const SEED: u64 = 0xf135_7aea_2e62_a9c5;

#[derive(Default, Clone, Copy)]
pub(crate) struct FxHasher {
    hash: u64,
}

impl FxHasher {
    #[inline]
    fn add(&mut self, word: u64) {
        self.hash = self.hash.wrapping_add(word).wrapping_mul(SEED);
    }
}

impl Hasher for FxHasher {
    #[inline]
    fn write(&mut self, bytes: &[u8]) {
        let mut chunks = bytes.chunks_exact(8);
        for chunk in &mut chunks {
            self.add(u64::from_le_bytes(chunk.try_into().unwrap()));
        }
        let rest = chunks.remainder();
        if !rest.is_empty() {
            let mut buf = [0u8; 8];
            buf[..rest.len()].copy_from_slice(rest);
            self.add(u64::from_le_bytes(buf));
        }
    }

    #[inline]
    fn write_u8(&mut self, i: u8) {
        self.add(i as u64);
    }

    #[inline]
    fn write_u16(&mut self, i: u16) {
        self.add(i as u64);
    }

    #[inline]
    fn write_u32(&mut self, i: u32) {
        self.add(i as u64);
    }

    #[inline]
    fn write_u64(&mut self, i: u64) {
        self.add(i);
    }

    #[inline]
    fn write_usize(&mut self, i: usize) {
        self.add(i as u64);
    }

    #[inline]
    fn finish(&self) -> u64 {
        self.hash.rotate_left(26)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn maps_find_what_they_hold() {
        let mut widths: FastMap<char, f32> = FastMap::default();
        for (i, ch) in ('\u{20}'..'\u{3000}').enumerate() {
            widths.insert(ch, i as f32);
        }
        assert_eq!(widths[&'A'], 33.0);
        assert_eq!(widths.get(&'\u{3000}'), None);
        let names: FastSet<&str> = ["F1", "F12", "F1_fallback"].into_iter().collect();
        assert!(names.contains("F12") && !names.contains("F2"));
    }
}
