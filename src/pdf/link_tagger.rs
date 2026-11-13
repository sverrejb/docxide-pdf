//! Tags a paragraph's links, language spans and math zones while its lines are drawn.

use std::sync::Arc;

use pdf_writer::{Content, Rect};

use super::layout::WordChunk;

pub(crate) struct LinkAnnotation {
    pub(super) rect: Rect,
    pub(super) url: String,
    /// Link structure element the annotation belongs to (tagged body text).
    pub(super) node: Option<usize>,
    /// Link text, for the annotation's /Contents description.
    pub(super) text: String,
}

/// Tags a body paragraph's hyperlinks as Link elements with their own marked
/// content while its lines are drawn (see `tagging`).
pub(super) struct LinkTagger<'a> {
    pub(super) tags: &'a mut super::tagging::Tags,
    pub(super) page: usize,
    pub(super) para: usize,
    pub(super) link: Option<(String, usize)>,
    /// The open Span or Formula (see `inline`) and its element.
    inline: Option<(Inline, usize)>,
}

/// A stretch of chunks tagged inside the open Link or paragraph.
enum Inline {
    /// Its `/Lang`, and whether it carries `/ActualText`.
    Span { lang: Option<String>, actual: bool },
    /// An Office Math zone, by its spoken form.
    Formula(Arc<str>),
}

impl<'a> LinkTagger<'a> {
    pub(super) fn new(tags: &'a mut super::tagging::Tags, page: usize, para: usize) -> Self {
        Self {
            tags,
            page,
            para,
            link: None,
            inline: None,
        }
    }

    /// Text in another language than the document's goes in a Span with
    /// `/Lang`, so a screen reader switches voice; caps and small caps, which
    /// draw other letters than the source has, in one whose `/ActualText`
    /// carries the source. One Span per stretch of chunks that need the same,
    /// inside the open Link or paragraph: a structure element, not nested
    /// marked content, which Poppler's structure reader loses text after.
    /// An Office Math zone's text goes in a Formula instead, whose `/Alt`
    /// speaks all of it (Word: "cap T equals 2 pi"); its glyphs stay the
    /// content for text extraction. Returns true when the marked content
    /// switched (the text matrix is reset).
    fn inline(
        &mut self,
        content: &mut Content,
        formula: Option<&Arc<str>>,
        lang: Option<&str>,
        actual: Option<&str>,
        before_switch: &mut dyn FnMut(&mut Content),
    ) -> bool {
        let lang = lang.filter(|l| !self.tags.is_document_lang(l));
        let same = match (&self.inline, formula) {
            (Some((Inline::Formula(open), _)), Some(f)) => Arc::ptr_eq(open, f),
            (Some((Inline::Span { lang: l, actual: a }, span)), None) => {
                let same = l.as_deref() == lang && *a == actual.is_some();
                if let (true, Some(text)) = (same, actual) {
                    self.tags.push_actual(*span, text);
                }
                same
            }
            (None, None) => lang.is_none() && actual.is_none(),
            _ => false,
        };
        if same {
            return false;
        }
        before_switch(content);
        content.end_text();
        let parent = self.open();
        self.inline = match formula {
            Some(f) => Some((Inline::Formula(f.clone()), self.tags.add_formula(parent, f))),
            None if lang.is_some() || actual.is_some() => {
                let span = self.tags.add_span(parent, lang, actual);
                let state = Inline::Span {
                    lang: lang.map(str::to_string),
                    actual: actual.is_some(),
                };
                Some((state, span))
            }
            None => None,
        };
        let node = self.inline.as_ref().map_or(parent, |i| i.1);
        self.tags.begin(content, self.page, node);
        content.begin_text();
        true
    }

    /// Before drawing `chunk` in a text object: switch to its Link (a note's
    /// reference mark links to the note), Span or Formula, and nest its Note.
    /// Returns true when the marked content switched (the text matrix is
    /// reset), after `before_switch` wrote what belongs to the old one.
    pub(super) fn chunk(
        &mut self,
        content: &mut Content,
        chunk: &WordChunk,
        link_url: Option<&str>,
        boundary_space: bool,
        before_switch: &mut dyn FnMut(&mut Content),
    ) -> bool {
        let mut switched = false;
        // Glyph-less chunks (underline bridges over spaces) don't break a link.
        if !chunk.text.is_empty() {
            switched |= self.enter(content, link_url, before_switch);
            // A Span's /ActualText covers the boundary space the Tj carries too.
            let actual = chunk.actual_text.as_ref().map(|t| {
                if boundary_space {
                    format!("{t} ")
                } else {
                    t.clone()
                }
            });
            let formula = chunk.formula.as_ref();
            // Punctuation and digits have no language: they stay in the open
            // language Span (the commas between Arabic words aren't English).
            let neutral = formula.is_none()
                && actual.is_none()
                && !chunk.text.chars().any(char::is_alphabetic);
            let lang = match &self.inline {
                None if neutral => None,
                Some((
                    Inline::Span {
                        lang,
                        actual: false,
                    },
                    _,
                )) if neutral => lang.clone(),
                _ => chunk.lang.as_deref().map(str::to_string),
            };
            switched |= self.inline(
                content,
                formula,
                lang.as_deref(),
                actual.as_deref(),
                before_switch,
            );
        }
        // The Note goes inside the link on its reference mark (Word nests it there).
        if let Some((endnote, id)) = chunk.note() {
            let parent = self.open();
            self.tags.note(endnote, id, parent);
        }
        switched
    }

    /// Switch the open marked content to the chunk's Link (or back to the
    /// paragraph). Like Word, marked content never changes inside a text
    /// object, so it is ended and reopened around the switch; returns true
    /// when it did (the text matrix is reset).
    fn enter(
        &mut self,
        content: &mut Content,
        url: Option<&str>,
        before_switch: &mut dyn FnMut(&mut Content),
    ) -> bool {
        if self.link.as_ref().map(|(u, _)| u.as_str()) == url {
            return false;
        }
        before_switch(content);
        content.end_text();
        self.inline = None;
        let node = match url {
            Some(u) => {
                let n = self.tags.add(self.para, "Link");
                self.link = Some((u.to_string(), n));
                n
            }
            None => {
                self.link = None;
                self.para
            }
        };
        self.tags.begin(content, self.page, node);
        content.begin_text();
        true
    }

    /// The element the current text belongs to.
    fn open(&self) -> usize {
        self.link.as_ref().map_or(self.para, |&(_, n)| n)
    }

    /// Before drawing an inline picture (outside a text object): a Figure
    /// inside the open element, read where the picture sits in the text, or
    /// an artifact when the picture is decorative.
    pub(super) fn begin_picture(
        &mut self,
        content: &mut Content,
        alt: Option<&str>,
        decorative: bool,
    ) {
        if decorative {
            super::tagging::Tags::end(content);
        } else {
            let figure = self.tags.add_figure(self.open(), alt);
            self.tags.begin(content, self.page, figure);
        }
    }

    /// Back to the open element after a picture or a Span.
    pub(super) fn resume(&mut self, content: &mut Content) {
        self.inline = None;
        self.tags.begin(content, self.page, self.open());
    }

    /// Draw something that isn't the paragraph's content (outside a text
    /// object) as an artifact, then go on in the open Span, Link or paragraph.
    pub(super) fn artifact(&mut self, content: &mut Content, draw: impl FnOnce(&mut Content)) {
        super::tagging::Tags::end(content);
        draw(content);
        let node = self.inline.as_ref().map_or_else(|| self.open(), |i| i.1);
        self.tags.begin(content, self.page, node);
    }

    /// `artifact` inside a text object (a text shadow's gray copy): the
    /// object is closed around it, so the text matrix starts afresh.
    pub(super) fn text_artifact(&mut self, content: &mut Content, draw: impl FnOnce(&mut Content)) {
        content.end_text();
        self.artifact(content, |content| {
            content.begin_text();
            draw(content);
            content.end_text();
        });
        content.begin_text();
    }

    pub(super) fn finish(mut self, content: &mut Content) {
        if self.link.take().is_some() | self.inline.take().is_some() {
            self.tags.begin(content, self.page, self.para);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn caps_span_groups_chunks_and_returns_to_the_paragraph() {
        use super::super::tagging::{self, ROOT, Tags};
        let mut tags = Tags::new();
        let p = tags.add(ROOT, "P");
        let mut content = tagging::artifact_content();
        tags.begin(&mut content, 0, p);
        content.begin_text();
        let mut lt = LinkTagger::new(&mut tags, 0, p);
        assert!(
            lt.inline(&mut content, None, None, Some("Pirmasis "), &mut |_| {}),
            "opens a Span"
        );
        assert!(
            !lt.inline(&mut content, None, None, Some("skirsnis"), &mut |_| {}),
            "extends it"
        );
        assert!(
            lt.inline(&mut content, None, None, None, &mut |_| {}),
            "back to the paragraph"
        );
        content.end_text();
        lt.finish(&mut content);
        let stream = String::from_utf8_lossy(&content.finish()).into_owned();
        assert_eq!(stream.matches("/Span").count(), 1);
        assert_eq!(stream.matches("BDC").count(), 3, "P, Span, P again");

        let mut pdf = pdf_writer::Pdf::new();
        let mut next = 1;
        let mut alloc = || {
            next += 1;
            pdf_writer::Ref::new(next)
        };
        tags.write(&mut pdf, &mut alloc, &[pdf_writer::Ref::new(1)], &[]);
        let bytes = pdf.finish();
        assert!(bytes.windows(19).any(|w| w == b"(Pirmasis skirsnis)"));
    }

    #[test]
    fn another_language_gets_a_lang_span() {
        use super::super::tagging::{self, ROOT, Tags};
        let mut tags = Tags::new();
        tags.lang = "en-US".into();
        let p = tags.add(ROOT, "P");
        let mut content = tagging::artifact_content();
        tags.begin(&mut content, 0, p);
        content.begin_text();
        let mut lt = LinkTagger::new(&mut tags, 0, p);
        assert!(
            !lt.inline(&mut content, None, Some("en-GB"), None, &mut |_| {}),
            "same language, no Span"
        );
        assert!(
            lt.inline(&mut content, None, Some("fr-FR"), None, &mut |_| {}),
            "French opens one"
        );
        assert!(
            !lt.inline(&mut content, None, Some("fr-FR"), None, &mut |_| {}),
            "and keeps it"
        );
        assert!(
            lt.inline(
                &mut content,
                None,
                Some("fr-FR"),
                Some("Bonjour"),
                &mut |_| {}
            ),
            "caps need their own"
        );
        assert!(
            lt.inline(&mut content, None, None, None, &mut |_| {}),
            "back to the paragraph"
        );
        content.end_text();
        lt.finish(&mut content);

        let mut pdf = pdf_writer::Pdf::new();
        let mut next = 1;
        let mut alloc = || {
            next += 1;
            pdf_writer::Ref::new(next)
        };
        tags.write(&mut pdf, &mut alloc, &[pdf_writer::Ref::new(1)], &[]);
        let bytes = pdf.finish();
        let count = |pat: &[u8]| bytes.windows(pat.len()).filter(|w| *w == pat).count();
        assert_eq!(count(b"/Lang (fr-FR)"), 2);
        assert_eq!(count(b"/ActualText (Bonjour)"), 1);
    }
}
