//! The font lookups of a conversion that found nothing, for hosts that supply
//! fonts on demand: convert, read the requests, register a face for a name the
//! host has (`register_font`) and convert again. The log belongs to the
//! converting thread; a conversion starts it anew.
use std::cell::RefCell;

/// Family names absent from the host/system index during one resolution, in
/// lookup order, and the requested style. Includes missing exact styles even
/// when a regular face can be used with synthetic bold or italic.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontRequest {
    pub names: Vec<String>,
    pub bold: bool,
    pub italic: bool,
}

#[derive(Default)]
struct Requests {
    current: Vec<FontRequest>,
    completed: Vec<FontRequest>,
}

thread_local! {
    static REQUESTS: RefCell<Requests> = RefCell::new(Requests::default());
}

/// Drain this thread's missing-font requests from its latest conversion. A new
/// conversion clears old requests, including when parsing fails. Requests are
/// deduplicated across render passes. Register a requested face and retry to
/// replace a fallback; the list describes the index, not embedded DOCX fonts.
pub fn take_font_requests() -> Vec<FontRequest> {
    REQUESTS.with_borrow_mut(|requests| std::mem::take(&mut requests.completed))
}

pub(crate) fn reset() {
    REQUESTS.with_borrow_mut(|requests| *requests = Requests::default());
}

/// One font resolution: the misses noted while it lives form one request.
pub(super) struct RequestScope;

pub(super) fn begin(bold: bool, italic: bool) -> RequestScope {
    REQUESTS.with_borrow_mut(|requests| {
        requests.current.push(FontRequest {
            names: Vec::new(),
            bold,
            italic,
        })
    });
    RequestScope
}

fn record(requests: &mut Requests, request: FontRequest) {
    if !request.names.is_empty() && !requests.completed.contains(&request) {
        requests.completed.push(request);
    }
}

impl Drop for RequestScope {
    fn drop(&mut self) {
        REQUESTS.with_borrow_mut(|requests| {
            if let Some(request) = requests.current.pop() {
                record(requests, request);
            }
        });
    }
}

pub(super) fn note_miss(name: &str, bold: bool, italic: bool) {
    REQUESTS.with_borrow_mut(|requests| {
        if let Some(current) = requests.current.last_mut() {
            if !current.names.iter().any(|n| n.eq_ignore_ascii_case(name)) {
                current.names.push(name.to_owned());
            }
        } else {
            record(
                requests,
                FontRequest {
                    names: vec![name.to_owned()],
                    bold,
                    italic,
                },
            );
        }
    });
}
