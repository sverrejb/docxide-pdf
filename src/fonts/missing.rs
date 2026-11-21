//! The fonts a document asks for that no face provides, for hosts that tell
//! the end user which font to supply (`Conversion::missing_fonts`). The list
//! belongs to the converting thread; `convert` drains it.
use std::cell::RefCell;

use super::resolve::{is_bahnschrift, word_substitute};
use super::{FontContext, primary_font_name};

/// A family and style the document asks for that no embedded, supplied or
/// system face provides. A missing bold or italic is listed even when the
/// regular face stands in with synthetic bold or italic: the real face would
/// change the output.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct MissingFont {
    pub name: String,
    pub bold: bool,
    pub italic: bool,
}

thread_local! {
    static MISSING: RefCell<Vec<MissingFont>> = const { RefCell::new(Vec::new()) };
}

/// Drain this thread's missing fonts.
pub(crate) fn take_missing_fonts() -> Vec<MissingFont> {
    MISSING.with_borrow_mut(std::mem::take)
}

/// Every render pass resolves the fonts again, so each font is kept once.
fn record(name: &str, bold: bool, italic: bool) {
    MISSING.with_borrow_mut(|missing| {
        if !missing
            .iter()
            .any(|m| m.bold == bold && m.italic == italic && m.name.eq_ignore_ascii_case(name))
        {
            missing.push(MissingFont {
                name: name.to_owned(),
                bold,
                italic,
            });
        }
    });
}

/// Lists the document's own font when neither an embedded nor an installed or
/// supplied face has its exact style. Only the name the document asks for:
/// the fallbacks `register_font` goes on to try are ours, not the document's.
/// The name is the one `register_font` looks up (Word draws Helvetica as
/// Arial), and never Bahnschrift, which Word cannot export, so supplying a
/// listed face always changes the output.
pub(crate) fn note_if_missing(font_name: &str, bold: bool, italic: bool, ctx: &FontContext) {
    let primary = primary_font_name(font_name);
    let name = word_substitute(primary).unwrap_or(primary);
    let key = (name.to_lowercase(), bold, italic);
    if !name.is_empty()
        && !is_bahnschrift(name)
        && !ctx.embedded_fonts.contains_key(&key)
        && ctx.font_index.exact_face(&key).is_none()
    {
        record(name, bold, italic);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::fonts::{EmbeddedFonts, FontFiles, discovery};
    use crate::model::FontTable;
    use std::collections::HashSet;

    #[test]
    fn only_faces_worth_supplying_are_missing() {
        let face = std::sync::Arc::new(include_bytes!("testdata/nb-hyphen-native.ttf").to_vec());
        let embedded: EmbeddedFonts =
            [(("synthetic hyphen".to_string(), false, false), face)].into();
        let ctx = FontContext {
            embedded_fonts: &embedded,
            font_table: &FontTable::new(),
            used_chars: &HashSet::new(),
            word_text: true,
            extra_glyphs: None,
            font_index: discovery::test_index(),
            font_files: &FontFiles::default(),
        };
        take_missing_fonts();
        note_if_missing("Synthetic Hyphen", false, false, &ctx);
        assert!(take_missing_fonts().is_empty());
        note_if_missing("Synthetic Hyphen", true, false, &ctx);
        note_if_missing("Synthetic Hyphen", true, false, &ctx);
        note_if_missing("Helvetica", false, false, &ctx);
        note_if_missing("Bahnschrift", false, false, &ctx);
        let mut missing = take_missing_fonts();
        assert!(
            missing
                .iter()
                .all(|m| m.name != "Helvetica" && m.name != "Bahnschrift")
        );
        missing.retain(|m| m.name == "Synthetic Hyphen");
        assert_eq!(missing.len(), 1);
        assert!(missing[0].bold && missing[0].name == "Synthetic Hyphen");
    }
}
