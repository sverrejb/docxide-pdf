//! Bookmark positions need no installed fonts, external PDF tools or private
//! fixtures: every paragraph names a font that is never found.
use std::io::{Cursor, Write};
use std::sync::{Arc, Barrier};

use docxide_pdf::{BookmarkPosition, Conversion, ConvertOptions, convert};

use crate::common::{convert_bytes, pdf_contains};

fn document(body: &str, section: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    zip.start_file(
        "word/document.xml",
        zip::write::SimpleFileOptions::default(),
    )
    .unwrap();
    write!(zip, r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="4000" w:h="4000"/><w:pgMar w:top="400" w:bottom="400" w:left="400" w:right="400"/>{section}</w:sectPr></w:body></w:document>"#).unwrap();
    zip.finish().unwrap().into_inner()
}

fn paragraph(name: &str, runs: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:widowControl w:val="0"/><w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:bookmarkStart w:id="1" w:name="{name}"/><w:r><w:rPr><w:rFonts w:ascii="Missing" w:hAnsi="Missing"/></w:rPr>{runs}</w:r><w:bookmarkEnd w:id="1"/></w:p>"#
    )
}

fn lines(count: usize) -> String {
    "<w:t>Line</w:t><w:br/>".repeat(count)
}

fn position(out: Conversion, name: &str) -> BookmarkPosition {
    out.bookmarks.into_iter().find(|p| p.name == name).unwrap()
}

fn names(out: &Conversion) -> Vec<&str> {
    out.bookmarks.iter().map(|p| p.name.as_str()).collect()
}

#[test]
fn bookmark_positions_cover_body_cells_nested_tables_and_page_splits() {
    let body = paragraph("body", "<w:t>Body</w:t>");
    let cell = paragraph("cell", "<w:t>Cell</w:t>");
    let nested = paragraph("nested", "<w:t>Nested</w:t>");
    let table = format!(
        "<w:tbl><w:tr><w:tc>{cell}<w:tbl><w:tr><w:tc>{nested}</w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>"
    );
    let long = paragraph("split", &lines(45));
    let out = convert_bytes(&document(&format!("{body}{table}{long}"), ""));
    assert!(!pdf_contains(&out.pdf, b"docxide-bookmark"));
    assert!(!pdf_contains(&out.pdf, b"/Subtype /Link"));
    let positions = out.bookmarks;
    for name in ["body", "cell", "nested", "split"] {
        let p = positions.iter().find(|p| p.name == name).unwrap();
        assert!(p.top.is_finite() && p.bottom.is_finite());
        assert!(p.end_page >= p.page);
        if name == "split" {
            assert!(p.end_page > p.page);
        } else {
            assert!(p.top > p.bottom);
        }
    }
    let unmarked = "<w:p><w:r><w:t>-</w:t></w:r></w:p>";
    assert!(convert_bytes(&document(unmarked, "")).bookmarks.is_empty());
}

#[test]
fn bookmark_positions_cell_split_reports_last_page() {
    let cell = paragraph("split_cell", &lines(45));
    let out = convert_bytes(&document(
        &format!("<w:tbl><w:tr><w:tc>{cell}</w:tc></w:tr></w:tbl>"),
        "",
    ));
    let p = position(out, "split_cell");
    assert!(p.end_page > p.page);
}

#[test]
fn bookmark_positions_column_split_keeps_physical_page_indices() {
    let body = paragraph("columns", &lines(20));
    let out = convert_bytes(&document(&body, r#"<w:cols w:num="2" w:space="200"/>"#));
    let p = position(out, "columns");
    assert_eq!((p.page, p.end_page), (0, 0));
    assert!(p.top > p.bottom);
}

#[test]
fn bookmark_positions_include_final_vertical_alignment() {
    let body = paragraph("body", "<w:t>Body</w:t>");
    let top = position(convert_bytes(&document(&body, "")), "body");
    let out = convert_bytes(&document(&body, r#"<w:vAlign w:val="center"/>"#));
    let centered = position(out, "body");
    assert!(top.top - centered.top > 40.0);
    assert!(((top.top - centered.top) - (top.bottom - centered.bottom)).abs() < 0.01);
    assert_eq!(centered.page, 0);
}

#[test]
fn bookmark_positions_cleared_by_failed_conversion() {
    convert_bytes(&document(&paragraph("old", "<w:t>Old</w:t>"), ""));
    assert!(convert(b"not a zip", &ConvertOptions::new()).is_err());
    let out = convert_bytes(&document(&paragraph("new", "<w:t>New</w:t>"), ""));
    assert_eq!(names(&out), ["new"]);
}

#[test]
fn bookmark_positions_belong_to_the_converting_thread() {
    let barrier = Arc::new(Barrier::new(2));
    let threads: Vec<_> = (0..2)
        .map(|i| {
            let barrier = barrier.clone();
            std::thread::spawn(move || {
                let bookmark = format!("thread_{i}");
                barrier.wait();
                let out = convert_bytes(&document(&paragraph(&bookmark, "<w:t>-</w:t>"), ""));
                barrier.wait();
                assert_eq!(names(&out), [bookmark]);
            })
        })
        .collect();
    for thread in threads {
        thread.join().unwrap();
    }
}

#[test]
fn bookmark_positions_from_layout_leave_no_links_and_clear_on_failure() {
    let cell = paragraph("cell", "<w:t>Cell</w:t>");
    let body = format!(
        "{}<w:tbl><w:tr><w:tc>{cell}</w:tc></w:tr></w:tbl>",
        paragraph("body", "<w:t>Body</w:t>")
    );
    let options = ConvertOptions::new().font_cache(false).display_list(true);
    let out = convert(&document(&body, ""), &options).unwrap();
    let layout = out.display.unwrap();
    assert!(layout.pages.iter().all(|page| page.links.is_empty()));
    assert_eq!(out.bookmarks.len(), 2);

    convert_bytes(&document(&body, ""));
    assert!(convert(b"not a zip", &options).is_err());
    let unmarked = "<w:p><w:r><w:t>-</w:t></w:r></w:p>";
    let out = convert(&document(unmarked, ""), &options).unwrap();
    assert!(out.bookmarks.is_empty() && out.display.is_some());
}
