//! The fonts a document asks for that no face provides, for hosts that tell
//! the end user which font to supply (`Conversion::missing_fonts`). The list
//! belongs to the converting thread; `convert` drains it.
use std::cell::RefCell;

/// A family and style the document asks for that no embedded, supplied or
/// system face provides. A missing bold or italic is listed even when the
/// regular face stands in with synthetic bold or italic: the real face would
/// change the output.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MissingFont {
    pub name: String,
    pub bold: bool,
    pub italic: bool,
}

thread_local! {
    static MISSING: RefCell<Vec<MissingFont>> = const { RefCell::new(Vec::new()) };
}

/// Drain this thread's missing fonts.
pub(crate) fn take_font_requests() -> Vec<MissingFont> {
    MISSING.with_borrow_mut(std::mem::take)
}

/// Every render pass resolves the fonts again, so each font is kept once.
pub(super) fn record(name: &str, bold: bool, italic: bool) {
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
