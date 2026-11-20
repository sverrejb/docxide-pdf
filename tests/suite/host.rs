//! Host APIs need no installed fonts, external PDF tools or private fixtures.
use std::io::{Cursor, Write};
use std::sync::{Arc, Barrier};

use docxide_pdf::display::FontSource;
use docxide_pdf::{
    ConvertOptions, clear_fonts, convert_docx_bytes_to_pdf_bytes_with, layout_docx_bytes_with,
    register_font, take_font_requests, use_system_fonts,
};

use crate::common::{convert_bytes as convert, pdf_contains};

/// Its family name is "NB Hyphen Fixture native".
const FONT: &[u8] = include_bytes!("../../src/fonts/testdata/nb-hyphen-native.ttf");
const FONT_NAME: &str = "NB Hyphen Fixture native";

/// Only registered fonts on this thread, from a clean registry.
struct Host;

impl Host {
    fn new() -> Self {
        clear_fonts();
        use_system_fonts(false);
        take_font_requests();
        Self
    }
}

impl Drop for Host {
    fn drop(&mut self) {
        clear_fonts();
        use_system_fonts(true);
    }
}

/// One paragraph in `font`, with `rpr` added to its run properties.
fn document(font: &str, rpr: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    zip.start_file(
        "word/document.xml",
        zip::write::SimpleFileOptions::default(),
    )
    .unwrap();
    write!(zip, r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>{rpr}</w:rPr><w:t>-</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="4000" w:h="4000"/><w:pgMar w:top="400" w:bottom="400" w:left="400" w:right="400"/></w:sectPr></w:body></w:document>"#).unwrap();
    zip.finish().unwrap().into_inner()
}

fn requested(name: &str) -> bool {
    take_font_requests()
        .iter()
        .any(|r| r.names.iter().any(|n| n == name))
}

/// A real font subset, not only the standard-14 fallback, reached the PDF.
fn embeds_font(pdf: &[u8]) -> bool {
    pdf_contains(pdf, b"/FontFile2")
}

#[test]
fn host_fonts_can_be_registered_and_cleared() {
    let _host = Host::new();
    let doc = document(FONT_NAME, "");
    assert!(!embeds_font(&convert(&doc)));
    assert_eq!(register_font(vec![0, 1, 2]), 0);
    assert_eq!(register_font(FONT.to_vec()), 1);
    assert!(embeds_font(&convert(&doc)));
    clear_fonts();
    assert!(!embeds_font(&convert(&doc)));
}

#[test]
fn host_fonts_cover_a_missing_style_with_the_registered_face() {
    let _host = Host::new();
    register_font(FONT.to_vec());
    assert!(embeds_font(&convert(&document(FONT_NAME, "<w:b/>"))));
}

#[test]
fn host_fonts_belong_to_the_registering_thread() {
    let barrier = Arc::new(Barrier::new(2));
    let threads: Vec<_> = (0..2)
        .map(|i| {
            let barrier = barrier.clone();
            std::thread::spawn(move || {
                let _host = Host::new();
                if i == 0 {
                    register_font(FONT.to_vec());
                }
                barrier.wait();
                let pdf = convert(&document(FONT_NAME, ""));
                assert_eq!(embeds_font(&pdf), i == 0);
            })
        })
        .collect();
    for thread in threads {
        thread.join().unwrap();
    }
}

#[test]
fn host_fonts_reach_the_display_list_as_bytes() {
    let _host = Host::new();
    register_font(FONT.to_vec());
    let options = ConvertOptions::new().font_cache(false);
    let layout = layout_docx_bytes_with(&document(FONT_NAME, ""), &options).unwrap();
    assert!(
        layout
            .fonts
            .iter()
            .any(|f| matches!(f.source, FontSource::Data { .. }))
    );
    assert!(
        !layout
            .fonts
            .iter()
            .any(|f| matches!(f.source, FontSource::File { .. }))
    );
}

#[test]
fn host_requests_name_missing_fonts_until_registered() {
    let _host = Host::new();
    let name = FONT_NAME.to_string();
    let doc = document(&name, "");
    convert(&doc);
    assert!(requested(&name));
    register_font(FONT.to_vec());
    convert(&doc);
    assert!(!requested(&name));
    clear_fonts();
    convert(&doc);
    assert!(requested(&name));
    assert!(take_font_requests().is_empty());
}

#[test]
fn host_requests_exact_style_even_when_regular_is_registered() {
    let _host = Host::new();
    let name = FONT_NAME.to_string();
    register_font(FONT.to_vec());
    convert(&document(&name, "<w:b/>"));
    assert!(
        take_font_requests()
            .iter()
            .any(|r| r.bold && r.names.contains(&name))
    );
}

#[test]
fn host_requests_cleared_by_failed_conversion() {
    let _host = Host::new();
    convert(&document(FONT_NAME, ""));
    assert!(convert_docx_bytes_to_pdf_bytes_with(b"not a zip", &ConvertOptions::new()).is_err());
    assert!(take_font_requests().is_empty());
}

#[test]
fn host_requests_belong_to_the_converting_thread() {
    let barrier = Arc::new(Barrier::new(2));
    let threads: Vec<_> = (0..2)
        .map(|i| {
            let barrier = barrier.clone();
            std::thread::spawn(move || {
                let _host = Host::new();
                if i == 0 {
                    register_font(FONT.to_vec());
                }
                barrier.wait();
                convert(&document(FONT_NAME, ""));
                barrier.wait();
                assert_eq!(requested(FONT_NAME), i == 1);
            })
        })
        .collect();
    for thread in threads {
        thread.join().unwrap();
    }
}
