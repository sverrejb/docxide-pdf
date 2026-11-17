//! Host APIs need no installed fonts, external PDF tools or private fixtures.
use std::io::{Cursor, Write};
use std::sync::{Arc, Barrier};

use docxide_pdf::{
    BookmarkPosition, ConvertOptions, clear_fonts, convert_docx_bytes_to_pdf_bytes_with,
    register_font, take_bookmark_positions, take_font_requests, use_system_fonts,
};

const FONT: &[u8] = include_bytes!("../../src/fonts/testdata/nb-hyphen-native.ttf");

struct Host;

impl Host {
    fn new() -> Self {
        clear_fonts();
        use_system_fonts(false);
        take_font_requests();
        take_bookmark_positions();
        Self
    }
}

impl Drop for Host {
    fn drop(&mut self) {
        clear_fonts();
        use_system_fonts(true);
    }
}

fn font_name() -> String {
    ttf_parser::Face::parse(FONT, 0)
        .unwrap()
        .names()
        .into_iter()
        .find(|n| n.name_id == ttf_parser::name_id::FAMILY && n.is_unicode())
        .unwrap()
        .to_string()
        .unwrap()
}

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

fn paragraph(name: &str, font: &str, extra: &str, runs: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:widowControl w:val="0"/><w:spacing w:line="240" w:lineRule="exact"/>{extra}</w:pPr><w:bookmarkStart w:id="1" w:name="{name}"/><w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/></w:rPr>{runs}</w:r><w:bookmarkEnd w:id="1"/></w:p>"#
    )
}

fn convert(input: &[u8]) -> Vec<u8> {
    let pdf = convert_docx_bytes_to_pdf_bytes_with(input, &ConvertOptions::new().font_cache(false))
        .unwrap();
    assert!(pdf.starts_with(b"%PDF-"));
    pdf
}

fn pdf_contains(pdf: &[u8], needle: &[u8]) -> bool {
    let contains = |bytes: &[u8]| bytes.windows(needle.len()).any(|w| w == needle);
    contains(pdf)
        || pdf
            .windows(7)
            .enumerate()
            .filter(|(_, w)| *w == b"stream\n")
            .any(|(start, _)| {
                let data = &pdf[start + 7..];
                let Some(end) = data.windows(10).position(|w| w == b"\nendstream") else {
                    return false;
                };
                miniz_oxide::inflate::decompress_to_vec_zlib(&data[..end])
                    .is_ok_and(|data| contains(&data))
            })
}

fn position(name: &str) -> BookmarkPosition {
    take_bookmark_positions()
        .into_iter()
        .find(|p| p.name == name)
        .unwrap()
}

#[test]
fn host_fonts_can_be_registered_requested_and_cleared() {
    let _host = Host::new();
    let name = font_name();
    let doc = document(&paragraph("body", &name, "", "<w:t>-</w:t>"), "");
    convert(&doc);
    assert!(take_font_requests().iter().any(|r| r.names.contains(&name)));
    assert_eq!(register_font(vec![0, 1, 2]), 0);
    assert_eq!(register_font(FONT.to_vec()), 1);
    let pdf = convert(&doc);
    assert!(!take_font_requests().iter().any(|r| r.names.contains(&name)));
    assert!(position("body").top > 0.0);
    // A real font subset, not only the standard-14 fallback, reached the PDF.
    assert!(pdf_contains(&pdf, b"/FontFile2"));
    clear_fonts();
    convert(&doc);
    assert!(take_font_requests().iter().any(|r| r.names.contains(&name)));
    assert!(take_font_requests().is_empty());
}

#[test]
fn host_requests_exact_style_even_when_regular_is_registered() {
    let _host = Host::new();
    let name = font_name();
    register_font(FONT.to_vec());
    let body = paragraph("bold", &name, "", "<w:t>-</w:t>").replace("</w:rPr>", "<w:b/></w:rPr>");
    convert(&document(&body, ""));
    assert!(
        take_font_requests()
            .iter()
            .any(|r| r.bold && r.names.contains(&name))
    );
}

#[test]
fn host_positions_cover_body_cells_nested_tables_and_page_splits() {
    let _host = Host::new();
    let body = paragraph("body", "Missing", "", "<w:t>Body</w:t>");
    let cell = paragraph("cell", "Missing", "", "<w:t>Cell</w:t>");
    let nested = paragraph("nested", "Missing", "", "<w:t>Nested</w:t>");
    let table = format!(
        "<w:tbl><w:tr><w:tc>{cell}<w:tbl><w:tr><w:tc>{nested}</w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>"
    );
    let long = paragraph("split", "Missing", "", &"<w:t>Line</w:t><w:br/>".repeat(45));
    let pdf = convert(&document(&format!("{body}{table}{long}"), ""));
    assert!(!pdf_contains(&pdf, b"docxide-bookmark"));
    assert!(!pdf_contains(&pdf, b"/Subtype /Link"));
    let positions = take_bookmark_positions();
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
    assert!(take_bookmark_positions().is_empty());
}

#[test]
fn host_cell_split_reports_last_page() {
    let _host = Host::new();
    let cell = paragraph(
        "split_cell",
        "Missing",
        "",
        &"<w:t>Line</w:t><w:br/>".repeat(45),
    );
    convert(&document(
        &format!("<w:tbl><w:tr><w:tc>{cell}</w:tc></w:tr></w:tbl>"),
        "",
    ));
    let p = position("split_cell");
    assert!(p.end_page > p.page);
}

#[test]
fn host_column_split_keeps_physical_page_indices() {
    let _host = Host::new();
    let body = paragraph(
        "columns",
        "Missing",
        "",
        &"<w:t>Line</w:t><w:br/>".repeat(20),
    );
    convert(&document(&body, r#"<w:cols w:num="2" w:space="200"/>"#));
    let p = position("columns");
    assert_eq!((p.page, p.end_page), (0, 0));
    assert!(p.top > p.bottom);
}

#[test]
fn host_positions_include_final_vertical_alignment() {
    let _host = Host::new();
    let body = paragraph("body", "Missing", "", "<w:t>Body</w:t>");
    convert(&document(&body, ""));
    let top = position("body");
    convert(&document(&body, r#"<w:vAlign w:val="center"/>"#));
    let centered = position("body");
    assert!(top.top - centered.top > 40.0);
    assert!(((top.top - centered.top) - (top.bottom - centered.bottom)).abs() < 0.01);
    assert_eq!(centered.page, 0);
}

#[test]
fn host_failed_conversion_clears_metadata() {
    let _host = Host::new();
    convert(&document(
        &paragraph("old", "Missing", "", "<w:t>Old</w:t>"),
        "",
    ));
    assert!(convert_docx_bytes_to_pdf_bytes_with(b"not a zip", &ConvertOptions::new()).is_err());
    assert!(take_font_requests().is_empty());
    assert!(take_bookmark_positions().is_empty());
}

#[test]
fn host_threads_do_not_share_fonts_requests_or_positions() {
    let barrier = Arc::new(Barrier::new(2));
    let threads: Vec<_> = (0..2)
        .map(|i| {
            let barrier = barrier.clone();
            std::thread::spawn(move || {
                let _host = Host::new();
                let name = font_name();
                if i == 0 {
                    register_font(FONT.to_vec());
                }
                barrier.wait();
                let bookmark = format!("thread_{i}");
                convert(&document(
                    &paragraph(&bookmark, &name, "", "<w:t>-</w:t>"),
                    "",
                ));
                barrier.wait();
                let missing = take_font_requests().iter().any(|r| r.names.contains(&name));
                assert_eq!(missing, i == 1);
                assert_eq!(
                    take_bookmark_positions()
                        .iter()
                        .map(|p| p.name.as_str())
                        .collect::<Vec<_>>(),
                    [bookmark]
                );
            })
        })
        .collect();
    for thread in threads {
        thread.join().unwrap();
    }
}
