//! Arabic contextual shaping by Unicode presentation forms.
//!
//! Each letter becomes the code point of the form its neighbours call for
//! (initial, medial, final), and lam + alef becomes its ligature. Fonts that
//! Word uses for Arabic map these code points to the same glyphs their GSUB
//! init/medi/fina/rlig features pick: Word's PDF of Arial Arabic shows
//! exactly the advances of Arial's U+FExx glyphs (arabic_rice heading: 3.211
//! 5.186 2.484 2.930 4.049 … 6.527 for lam-alef). So widths, line breaking,
//! subsetting and glyph lookup need no glyph-level shaping; text extraction
//! maps the forms back to their letters (`source_letters`).
//
// ponytail: presentation forms only. Every Office Arabic font carries them;
// the faces that don't (Nastaliq, Thuluth display fonts) draw the plain
// letters, and need real OpenType shaping (GSUB contexts, GPOS marks) anyway.

/// Letter → [isolated, final, initial, medial] presentation forms (0 = none),
/// from the Unicode decomposition data of U+FB50–FDFF and U+FE70–FEFF.
const FORMS: &[(char, [u16; 4])] = &[
    ('\u{0621}', [0xFE80, 0, 0, 0]),
    ('\u{0622}', [0xFE81, 0xFE82, 0, 0]),
    ('\u{0623}', [0xFE83, 0xFE84, 0, 0]),
    ('\u{0624}', [0xFE85, 0xFE86, 0, 0]),
    ('\u{0625}', [0xFE87, 0xFE88, 0, 0]),
    ('\u{0626}', [0xFE89, 0xFE8A, 0xFE8B, 0xFE8C]),
    ('\u{0627}', [0xFE8D, 0xFE8E, 0, 0]),
    ('\u{0628}', [0xFE8F, 0xFE90, 0xFE91, 0xFE92]),
    ('\u{0629}', [0xFE93, 0xFE94, 0, 0]),
    ('\u{062A}', [0xFE95, 0xFE96, 0xFE97, 0xFE98]),
    ('\u{062B}', [0xFE99, 0xFE9A, 0xFE9B, 0xFE9C]),
    ('\u{062C}', [0xFE9D, 0xFE9E, 0xFE9F, 0xFEA0]),
    ('\u{062D}', [0xFEA1, 0xFEA2, 0xFEA3, 0xFEA4]),
    ('\u{062E}', [0xFEA5, 0xFEA6, 0xFEA7, 0xFEA8]),
    ('\u{062F}', [0xFEA9, 0xFEAA, 0, 0]),
    ('\u{0630}', [0xFEAB, 0xFEAC, 0, 0]),
    ('\u{0631}', [0xFEAD, 0xFEAE, 0, 0]),
    ('\u{0632}', [0xFEAF, 0xFEB0, 0, 0]),
    ('\u{0633}', [0xFEB1, 0xFEB2, 0xFEB3, 0xFEB4]),
    ('\u{0634}', [0xFEB5, 0xFEB6, 0xFEB7, 0xFEB8]),
    ('\u{0635}', [0xFEB9, 0xFEBA, 0xFEBB, 0xFEBC]),
    ('\u{0636}', [0xFEBD, 0xFEBE, 0xFEBF, 0xFEC0]),
    ('\u{0637}', [0xFEC1, 0xFEC2, 0xFEC3, 0xFEC4]),
    ('\u{0638}', [0xFEC5, 0xFEC6, 0xFEC7, 0xFEC8]),
    ('\u{0639}', [0xFEC9, 0xFECA, 0xFECB, 0xFECC]),
    ('\u{063A}', [0xFECD, 0xFECE, 0xFECF, 0xFED0]),
    ('\u{0641}', [0xFED1, 0xFED2, 0xFED3, 0xFED4]),
    ('\u{0642}', [0xFED5, 0xFED6, 0xFED7, 0xFED8]),
    ('\u{0643}', [0xFED9, 0xFEDA, 0xFEDB, 0xFEDC]),
    ('\u{0644}', [0xFEDD, 0xFEDE, 0xFEDF, 0xFEE0]),
    ('\u{0645}', [0xFEE1, 0xFEE2, 0xFEE3, 0xFEE4]),
    ('\u{0646}', [0xFEE5, 0xFEE6, 0xFEE7, 0xFEE8]),
    ('\u{0647}', [0xFEE9, 0xFEEA, 0xFEEB, 0xFEEC]),
    ('\u{0648}', [0xFEED, 0xFEEE, 0, 0]),
    ('\u{0649}', [0xFEEF, 0xFEF0, 0xFBE8, 0xFBE9]),
    ('\u{064A}', [0xFEF1, 0xFEF2, 0xFEF3, 0xFEF4]),
    ('\u{0671}', [0xFB50, 0xFB51, 0, 0]),
    ('\u{0677}', [0xFBDD, 0, 0, 0]),
    ('\u{0679}', [0xFB66, 0xFB67, 0xFB68, 0xFB69]),
    ('\u{067A}', [0xFB5E, 0xFB5F, 0xFB60, 0xFB61]),
    ('\u{067B}', [0xFB52, 0xFB53, 0xFB54, 0xFB55]),
    ('\u{067E}', [0xFB56, 0xFB57, 0xFB58, 0xFB59]),
    ('\u{067F}', [0xFB62, 0xFB63, 0xFB64, 0xFB65]),
    ('\u{0680}', [0xFB5A, 0xFB5B, 0xFB5C, 0xFB5D]),
    ('\u{0683}', [0xFB76, 0xFB77, 0xFB78, 0xFB79]),
    ('\u{0684}', [0xFB72, 0xFB73, 0xFB74, 0xFB75]),
    ('\u{0686}', [0xFB7A, 0xFB7B, 0xFB7C, 0xFB7D]),
    ('\u{0687}', [0xFB7E, 0xFB7F, 0xFB80, 0xFB81]),
    ('\u{0688}', [0xFB88, 0xFB89, 0, 0]),
    ('\u{068C}', [0xFB84, 0xFB85, 0, 0]),
    ('\u{068D}', [0xFB82, 0xFB83, 0, 0]),
    ('\u{068E}', [0xFB86, 0xFB87, 0, 0]),
    ('\u{0691}', [0xFB8C, 0xFB8D, 0, 0]),
    ('\u{0698}', [0xFB8A, 0xFB8B, 0, 0]),
    ('\u{06A4}', [0xFB6A, 0xFB6B, 0xFB6C, 0xFB6D]),
    ('\u{06A6}', [0xFB6E, 0xFB6F, 0xFB70, 0xFB71]),
    ('\u{06A9}', [0xFB8E, 0xFB8F, 0xFB90, 0xFB91]),
    ('\u{06AD}', [0xFBD3, 0xFBD4, 0xFBD5, 0xFBD6]),
    ('\u{06AF}', [0xFB92, 0xFB93, 0xFB94, 0xFB95]),
    ('\u{06B1}', [0xFB9A, 0xFB9B, 0xFB9C, 0xFB9D]),
    ('\u{06B3}', [0xFB96, 0xFB97, 0xFB98, 0xFB99]),
    ('\u{06BA}', [0xFB9E, 0xFB9F, 0, 0]),
    ('\u{06BB}', [0xFBA0, 0xFBA1, 0xFBA2, 0xFBA3]),
    ('\u{06BE}', [0xFBAA, 0xFBAB, 0xFBAC, 0xFBAD]),
    ('\u{06C0}', [0xFBA4, 0xFBA5, 0, 0]),
    ('\u{06C1}', [0xFBA6, 0xFBA7, 0xFBA8, 0xFBA9]),
    ('\u{06C5}', [0xFBE0, 0xFBE1, 0, 0]),
    ('\u{06C6}', [0xFBD9, 0xFBDA, 0, 0]),
    ('\u{06C7}', [0xFBD7, 0xFBD8, 0, 0]),
    ('\u{06C8}', [0xFBDB, 0xFBDC, 0, 0]),
    ('\u{06C9}', [0xFBE2, 0xFBE3, 0, 0]),
    ('\u{06CB}', [0xFBDE, 0xFBDF, 0, 0]),
    ('\u{06CC}', [0xFBFC, 0xFBFD, 0xFBFE, 0xFBFF]),
    ('\u{06D0}', [0xFBE4, 0xFBE5, 0xFBE6, 0xFBE7]),
    ('\u{06D2}', [0xFBAE, 0xFBAF, 0, 0]),
    ('\u{06D3}', [0xFBB0, 0xFBB1, 0, 0]),
];

const LAM: char = '\u{0644}';

/// Alef variant → lam-alef ligature [isolated, final] (U+FEF5–FEFC).
const LAM_ALEF: &[(char, [char; 2])] = &[
    ('\u{0622}', ['\u{FEF5}', '\u{FEF6}']),
    ('\u{0623}', ['\u{FEF7}', '\u{FEF8}']),
    ('\u{0625}', ['\u{FEF9}', '\u{FEFA}']),
    ('\u{0627}', ['\u{FEFB}', '\u{FEFC}']),
];

#[derive(Clone, Copy, PartialEq)]
enum Joining {
    /// Joins on both sides.
    Dual,
    /// Joins only to the letter before it (alef, dal, reh, waw, …).
    Right,
    /// Tatweel and ZWJ: joins both neighbours without a form of its own.
    Causing,
    /// Combining marks: invisible to joining.
    Transparent,
    None,
}

fn forms(ch: char) -> Option<&'static [u16; 4]> {
    FORMS
        .binary_search_by_key(&ch, |&(c, _)| c)
        .ok()
        .map(|i| &FORMS[i].1)
}

/// An Arabic combining mark (harakat, Quranic marks): transparent to joining,
/// a non-spacing mark to bidi.
pub(crate) fn is_arabic_mark(ch: char) -> bool {
    matches!(ch,
        '\u{0610}'..='\u{061A}'
        | '\u{064B}'..='\u{065F}'
        | '\u{0670}'
        | '\u{06D6}'..='\u{06DC}'
        | '\u{06DF}'..='\u{06E4}'
        | '\u{06E7}'..='\u{06E8}'
        | '\u{06EA}'..='\u{06ED}'
    )
}

fn joining(ch: char) -> Joining {
    match ch {
        '\u{0640}' | '\u{200D}' => Joining::Causing,
        c if is_arabic_mark(c) => Joining::Transparent,
        _ => match forms(ch) {
            Some(f) if f[2] != 0 => Joining::Dual,
            Some(f) if f[1] != 0 => Joining::Right,
            _ => Joining::None,
        },
    }
}

/// `text` with every Arabic letter replaced by its contextual form; text
/// without Arabic comes back unchanged. Isolated letters keep their own
/// code point (the same glyph in every font that has the forms).
pub(crate) fn shape(text: &str) -> String {
    let chars: Vec<char> = text.chars().collect();
    let kinds: Vec<Joining> = chars.iter().map(|&c| joining(c)).collect();
    // The nearest non-transparent neighbour on each side.
    let neighbour = |i: usize, step: isize| {
        let mut j = i as isize + step;
        while j >= 0 && (j as usize) < chars.len() && kinds[j as usize] == Joining::Transparent {
            j += step;
        }
        (j >= 0 && (j as usize) < chars.len()).then_some(j as usize)
    };
    let mut out = String::with_capacity(text.len());
    let mut i = 0;
    while i < chars.len() {
        let ch = chars[i];
        if !matches!(kinds[i], Joining::Dual | Joining::Right) {
            out.push(ch);
            i += 1;
            continue;
        }
        let joins_before =
            neighbour(i, -1).is_some_and(|p| matches!(kinds[p], Joining::Dual | Joining::Causing));
        if ch == LAM
            && let Some(lig) = chars
                .get(i + 1)
                .and_then(|a| LAM_ALEF.iter().find(|(alef, _)| alef == a))
        {
            out.push(lig.1[usize::from(joins_before)]);
            i += 2;
            continue;
        }
        let joins_after = kinds[i] == Joining::Dual
            && neighbour(i, 1).is_some_and(|n| {
                matches!(kinds[n], Joining::Dual | Joining::Right | Joining::Causing)
            });
        let slot = match (joins_before, joins_after) {
            (false, false) => None,
            (true, false) => Some(1),
            (false, true) => Some(2),
            (true, true) => Some(3),
        };
        let form = slot
            .and_then(|s| forms(ch).map(|f| f[s]))
            .and_then(|cp| char::from_u32(u32::from(cp)))
            .filter(|&c| c != '\0');
        out.push(form.unwrap_or(ch));
        i += 1;
    }
    out
}

/// The letters a presentation form was shaped from, for text extraction:
/// one letter, or lam and alef for a ligature.
pub(crate) fn source_letters(ch: char) -> Option<(char, Option<char>)> {
    if let Some((alef, _)) = LAM_ALEF.iter().find(|(_, lig)| lig.contains(&ch)) {
        return Some((LAM, Some(*alef)));
    }
    let cp = u16::try_from(u32::from(ch)).ok()?;
    if !(0xFB50..=0xFEFC).contains(&cp) {
        return None;
    }
    FORMS
        .iter()
        .find(|(_, f)| f.contains(&cp))
        .map(|&(letter, _)| (letter, None))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn codes(s: &str) -> Vec<u32> {
        s.chars().map(u32::from).collect()
    }

    #[test]
    fn forms_follow_the_joining_neighbours() {
        // فوائد: feh initial, waw final (it never joins on), alef isolated,
        // yeh-hamza initial, dal final.
        assert_eq!(
            codes(&shape("فوائد")),
            [0xFED3, 0xFEEE, 0x0627, 0xFE8B, 0xFEAA]
        );
        // بتب: initial, medial, final.
        assert_eq!(codes(&shape("بتب")), [0xFE91, 0xFE98, 0xFE90]);
        // A space breaks the joining; Latin passes through.
        assert_eq!(codes(&shape("ب ب x")), [0x0628, 0x20, 0x0628, 0x20, 0x78]);
    }

    #[test]
    fn lam_alef_ligates_and_marks_are_transparent() {
        // الأرز: alef isolated, lam-alef-hamza isolated (alef ends the join),
        // reh and zain isolated.
        assert_eq!(codes(&shape("الأرز")), [0x0627, 0xFEF7, 0x0631, 0x0632]);
        // تلا: teh initial, lam-alef final.
        assert_eq!(codes(&shape("تلا")), [0xFE97, 0xFEFC]);
        // A fatha between beh and teh doesn't stop them joining.
        assert_eq!(codes(&shape("بَت")), [0xFE91, 0x064E, 0xFE96]);
        // Tatweel joins both sides.
        assert_eq!(codes(&shape("بـب")), [0xFE91, 0x0640, 0xFE90]);
    }

    #[test]
    fn forms_map_back_to_their_letters() {
        assert_eq!(source_letters('\u{FED3}'), Some(('\u{0641}', None)));
        assert_eq!(source_letters('\u{FEF7}'), Some((LAM, Some('\u{0623}'))));
        assert_eq!(source_letters('\u{FBFE}'), Some(('\u{06CC}', None)));
        assert_eq!(source_letters('\u{0628}'), None);
        assert_eq!(source_letters('a'), None);
    }
}
