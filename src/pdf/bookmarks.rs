//! Carry paragraph fragments through the same page placement and trial-layout
//! handling as links. These internal markers are removed before PDF assembly.
use std::cell::RefCell;
use std::collections::BTreeMap;

use super::LinkAnnotation;
use crate::model::Paragraph;

/// The first and last laid-out fragments of a bookmarked body/cell paragraph.
/// Pages are zero-based; coordinates are PDF points measured up from the bottom
/// of the page, including section vertical alignment. This describes the
/// paragraph carrying bookmarkStart, not a multi-paragraph bookmark range.
#[derive(Debug, Clone, PartialEq)]
pub struct BookmarkPosition {
    pub name: String,
    pub page: usize,
    pub top: f32,
    pub end_page: usize,
    pub bottom: f32,
}

thread_local! {
    static POSITIONS: RefCell<Vec<BookmarkPosition>> = const { RefCell::new(Vec::new()) };
}

/// Drain the final render pass's paragraph positions from this thread's latest
/// conversion. A new conversion clears them, including when parsing fails.
/// Headers, footers and text boxes are not part of this body-paragraph API.
pub fn take_bookmark_positions() -> Vec<BookmarkPosition> {
    POSITIONS.with_borrow_mut(std::mem::take)
}

const MARKER: &str = "\0docxide-bookmark:";

pub(super) fn mark(links: &mut Vec<LinkAnnotation>, para: &Paragraph, top: f32, bottom: f32) {
    for name in &para.bookmarks {
        links.push(LinkAnnotation {
            rect: pdf_writer::Rect::new(0.0, bottom, 0.0, top),
            url: format!("{MARKER}{name}"),
            node: None,
            text: String::new(),
        });
    }
}

pub(super) fn collect(pages: &mut [Vec<LinkAnnotation>]) -> Vec<BookmarkPosition> {
    let mut positions = BTreeMap::<String, BookmarkPosition>::new();
    for (page, links) in pages.iter_mut().enumerate() {
        links.retain(|link| {
            let Some(name) = link.url.strip_prefix(MARKER) else {
                return true;
            };
            let position = positions
                .entry(name.to_owned())
                .or_insert_with(|| BookmarkPosition {
                    name: name.to_owned(),
                    page,
                    top: link.rect.y2,
                    end_page: page,
                    bottom: link.rect.y1,
                });
            position.end_page = page;
            position.bottom = link.rect.y1;
            false
        });
    }
    positions.into_values().collect()
}

pub(super) fn store(mut positions: Vec<BookmarkPosition>, offsets: &[f32]) {
    for position in &mut positions {
        position.top -= offsets[position.page];
        position.bottom -= offsets[position.end_page];
    }
    POSITIONS.with_borrow_mut(|current| *current = positions);
}
