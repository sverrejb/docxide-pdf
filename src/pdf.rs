use pdf_writer::{Content, Name, Pdf, Rect, Ref, Str};

use crate::error::Error;
use crate::model::Document;

pub fn render(doc: &Document) -> Result<Vec<u8>, Error> {
    let mut pdf = Pdf::new();

    let catalog_id = Ref::new(1);
    let pages_id = Ref::new(2);
    let page_id = Ref::new(3);
    let content_id = Ref::new(4);
    let font_id = Ref::new(5);

    pdf.catalog(catalog_id).pages(pages_id);
    pdf.pages(pages_id).kids([page_id]).count(1);

    let mut content = Content::new();

    let mut cursor_y = doc.page_height - doc.margin_top;
    let cursor_x = doc.margin_left;

    for para in &doc.paragraphs {
        cursor_y -= para.space_before;

        for run in &para.runs {
            let text_bytes = run.text.as_bytes();
            content
                .begin_text()
                .set_font(Name(b"F1"), run.font_size)
                .next_line(cursor_x, cursor_y)
                .show(Str(text_bytes))
                .end_text();
            cursor_y -= run.font_size * 1.2;
        }

        cursor_y -= para.space_after;
    }

    pdf.stream(content_id, &content.finish());

    pdf.page(page_id)
        .media_box(Rect::new(0.0, 0.0, doc.page_width, doc.page_height))
        .parent(pages_id)
        .contents(content_id)
        .resources()
        .fonts()
        .pair(Name(b"F1"), font_id);

    pdf.type1_font(font_id).base_font(Name(b"Helvetica"));

    Ok(pdf.finish())
}
