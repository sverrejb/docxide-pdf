//! Host APIs need no installed fonts, external PDF tools or private fixtures.
use std::io::{Cursor, Write};
use std::sync::{Arc, Barrier};

use docxide_pdf::display::FontSource;
use docxide_pdf::{
    ConvertOptions, clear_fonts, layout_docx_bytes_with, register_font, use_system_fonts,
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
