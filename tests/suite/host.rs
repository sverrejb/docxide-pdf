//! Host APIs need no installed fonts, external PDF tools or private fixtures.
use std::io::{Cursor, Write};
use std::sync::{Arc, Barrier};

use docxide_pdf::display::FontSource;
use docxide_pdf::{
    Conversion, ConvertOptions, clear_fonts, convert, register_font, use_system_fonts,
};

use crate::common::{convert_bytes, pdf_contains};

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

fn requested(out: &Conversion, name: &str) -> bool {
    out.missing_fonts.iter().any(|m| m.name == name)
}

fn pdf(docx: &[u8]) -> Vec<u8> {
    convert_bytes(docx).pdf
}

/// A real font subset, not only the standard-14 fallback, reached the PDF.
fn embeds_font(pdf: &[u8]) -> bool {
    pdf_contains(pdf, b"/FontFile2")
}

#[test]
fn host_fonts_can_be_registered_and_cleared() {
    let _host = Host::new();
    let doc = document(FONT_NAME, "");
    assert!(!embeds_font(&pdf(&doc)));
    assert_eq!(register_font(vec![0, 1, 2]), 0);
    assert_eq!(register_font(FONT.to_vec()), 1);
    assert!(embeds_font(&pdf(&doc)));
    clear_fonts();
    assert!(!embeds_font(&pdf(&doc)));
}

#[test]
fn host_fonts_cover_a_missing_style_with_the_registered_face() {
    let _host = Host::new();
    register_font(FONT.to_vec());
    assert!(embeds_font(&pdf(&document(FONT_NAME, "<w:b/>"))));
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
                let pdf = pdf(&document(FONT_NAME, ""));
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
    let options = ConvertOptions::new().font_cache(false).display_list(true);
    let out = convert(&document(FONT_NAME, ""), &options).unwrap();
    let layout = out.display.unwrap();
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
    assert!(requested(&convert_bytes(&doc), &name));
    register_font(FONT.to_vec());
    assert!(!requested(&convert_bytes(&doc), &name));
    clear_fonts();
    assert!(requested(&convert_bytes(&doc), &name));
}

#[test]
fn host_requests_leave_out_the_fallbacks_tried_instead() {
    let _host = Host::new();
    let out = convert_bytes(&document(FONT_NAME, "<w:i/>"));
    let missing: Vec<_> = out
        .missing_fonts
        .iter()
        .map(|m| (m.name.as_str(), m.bold, m.italic))
        .collect();
    assert_eq!(missing, [(FONT_NAME, false, true)]);
}

#[test]
fn host_requests_exact_style_even_when_regular_is_registered() {
    let _host = Host::new();
    let name = FONT_NAME.to_string();
    register_font(FONT.to_vec());
    let out = convert_bytes(&document(&name, "<w:b/>"));
    assert!(out.missing_fonts.iter().any(|m| m.bold && m.name == name));
}

#[test]
fn host_requests_cleared_by_failed_conversion() {
    let _host = Host::new();
    convert_bytes(&document(FONT_NAME, ""));
    assert!(convert(b"not a zip", &ConvertOptions::new()).is_err());
    register_font(FONT.to_vec());
    assert!(
        convert_bytes(&document(FONT_NAME, ""))
            .missing_fonts
            .is_empty()
    );
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
                let out = convert_bytes(&document(FONT_NAME, ""));
                barrier.wait();
                assert_eq!(requested(&out, FONT_NAME), i == 1);
            })
        })
        .collect();
    for thread in threads {
        thread.join().unwrap();
    }
}

#[test]
fn display_list_leaves_the_pdf_unchanged() {
    // Pictures, forms and charts: what the recorder captures besides text.
    for case in ["case78", "picture-part-reuse", "case30", "charts_ex"] {
        let docx = std::fs::read(format!("tests/fixtures/cases/{case}/input.docx")).unwrap();
        let options = ConvertOptions::new().font_cache(false);
        let plain = convert(&docx, &options).unwrap();
        let recorded = convert(&docx, &options.clone().display_list(true)).unwrap();
        assert!(plain.display.is_none());
        assert!(recorded.display.is_some_and(|d| !d.pages.is_empty()));
        assert!(
            plain.pdf == recorded.pdf,
            "{case}: the display list moved PDF bytes"
        );
    }
}
