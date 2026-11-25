//! The fonts a document asks for that no face provides, for hosts that tell
//! the end user which font to supply (`Conversion::missing_fonts`). The list
//! belongs to the converting thread; `convert` drains it.
use std::cell::RefCell;

/// A family and style the document asks for that no embedded, supplied or
/// system face provides. A missing bold or italic is listed even when the
/// regular face stands in with synthetic bold or italic: the real face would
/// change the output.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct MissingFont {
    /// The family to supply.
    pub name: String,
    /// Whether the missing face is bold.
    pub bold: bool,
    /// Whether the missing face is italic.
    pub italic: bool,
    /// The document's own name when Word draws `name` in its place: Windows
    /// maps Helvetica, Times and Courier to Arial, Times New Roman and Courier
    /// New, even over an installed or embedded face of the document's name.
    pub replaces: Option<String>,
}

thread_local! {
    static MISSING: RefCell<Vec<MissingFont>> = const { RefCell::new(Vec::new()) };
}

/// Drain this thread's missing fonts.
pub(crate) fn take_missing_fonts() -> Vec<MissingFont> {
    MISSING.with_borrow_mut(std::mem::take)
}

/// Every render pass resolves the fonts again, so each font is kept once.
pub(crate) fn record(font: &MissingFont) {
    MISSING.with_borrow_mut(|missing| {
        if !missing.iter().any(|m| {
            m.bold == font.bold
                && m.italic == font.italic
                && m.name.eq_ignore_ascii_case(&font.name)
        }) {
            missing.push(font.clone());
        }
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::fast_hash::FastSet;
    use crate::fonts::{EmbeddedFonts, FontContext, FontFiles, FontIndex, register_font};
    use crate::model::FontTable;
    use pdf_writer::{Pdf, Ref};

    #[test]
    fn only_faces_worth_supplying_are_missing() {
        let face = std::sync::Arc::new(include_bytes!("testdata/nb-hyphen-native.ttf").to_vec());
        let embedded: EmbeddedFonts =
            [(("synthetic hyphen".to_string(), false, false), face)].into();
        let ctx = FontContext {
            embedded_fonts: &embedded,
            font_table: &FontTable::new(),
            used_chars: &FastSet::default(),
            word_text: true,
            extra_glyphs: None,
            font_index: &FontIndex::default(),
            font_files: &FontFiles::default(),
        };
        let mut next = 0;
        let mut missing = |name: &str, bold: bool| {
            let mut alloc = || {
                next += 1;
                Ref::new(next)
            };
            register_font(
                &mut Pdf::new(),
                name,
                bold,
                false,
                "F1".into(),
                &mut alloc,
                &ctx,
            )
            .missing
        };
        assert_eq!(missing("Synthetic Hyphen", false), None);
        let bold = missing("Synthetic Hyphen", true).expect("no bold face");
        assert!(bold.bold && bold.name == "Synthetic Hyphen");
        assert_eq!(bold.replaces, None);
        let arial = missing("Helvetica", false).expect("no Arial");
        assert_eq!(
            (arial.name.as_str(), arial.replaces.as_deref()),
            ("Arial", Some("Helvetica"))
        );
        assert_eq!(missing("Bahnschrift", false), None);

        take_missing_fonts();
        record(&bold);
        record(&bold);
        assert_eq!(take_missing_fonts(), [bold]);
    }
}
