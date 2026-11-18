//! Host APIs need no installed fonts, external PDF tools or private fixtures.
use std::io::{Cursor, Write};
use std::sync::{Arc, Barrier};

use docxide_pdf::{
    ConvertOptions, clear_fonts, convert_docx_bytes_to_pdf_bytes_with, register_font,
    take_font_requests, use_system_fonts,
};

const FONT: &[u8] = include_bytes!("../../src/fonts/testdata/nb-hyphen-native.ttf");

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
    let doc = document(&font_name(), "");
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
    assert!(embeds_font(&convert(&document(&font_name(), "<w:b/>"))));
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
                let pdf = convert(&document(&font_name(), ""));
                assert_eq!(embeds_font(&pdf), i == 0);
            })
        })
        .collect();
    for thread in threads {
        thread.join().unwrap();
    }
}

#[test]
fn host_requests_name_missing_fonts_until_registered() {
    let _host = Host::new();
    let name = font_name();
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
    let name = font_name();
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
    convert(&document(&font_name(), ""));
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
                convert(&document(&font_name(), ""));
                barrier.wait();
                assert_eq!(requested(&font_name()), i == 1);
            })
        })
        .collect();
    for thread in threads {
        thread.join().unwrap();
    }
}
