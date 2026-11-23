//! Host APIs need no installed fonts, external PDF tools or private fixtures.
use std::io::{Cursor, Write};
use std::sync::{Arc, Barrier};

use docxide_pdf::display::FontSource;
use docxide_pdf::{Conversion, ConvertOptions, convert};

use crate::common::{convert_bytes_with as run, pdf_contains};

/// Its family name is "NB Hyphen Fixture native".
const FONT: &[u8] = include_bytes!("../../src/fonts/testdata/nb-hyphen-native.ttf");
const FONT_NAME: &str = "NB Hyphen Fixture native";

/// No system fonts: only what the options supply.
fn host() -> ConvertOptions {
    ConvertOptions::new().font_cache(false).system_fonts(false)
}

fn with_font() -> ConvertOptions {
    host().font(FONT.to_vec())
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

/// A real font subset, not only the standard-14 fallback, reached the PDF.
fn embeds_font(pdf: &[u8]) -> bool {
    pdf_contains(pdf, b"/FontFile2")
}

#[test]
fn host_fonts_can_be_supplied_and_left_out() {
    let doc = document(FONT_NAME, "");
    assert!(!embeds_font(&run(&doc, &host()).pdf));
    assert!(!embeds_font(&run(&doc, &host().font(vec![0, 1, 2])).pdf));
    assert!(embeds_font(&run(&doc, &with_font()).pdf));
    assert!(!embeds_font(&run(&doc, &host()).pdf));
}

#[test]
fn host_fonts_cover_a_missing_style_with_the_supplied_face() {
    assert!(embeds_font(
        &run(&document(FONT_NAME, "<w:b/>"), &with_font()).pdf
    ));
}

#[test]
fn host_fonts_reach_other_threads_with_the_options() {
    // Built on this thread, converted on others: the fonts travel with the options.
    let options = [with_font(), host()];
    let barrier = Arc::new(Barrier::new(2));
    let threads: Vec<_> = options
        .into_iter()
        .enumerate()
        .map(|(i, options)| {
            let barrier = barrier.clone();
            std::thread::spawn(move || {
                barrier.wait();
                let pdf = run(&document(FONT_NAME, ""), &options).pdf;
                assert_eq!(embeds_font(&pdf), i == 0);
            })
        })
        .collect();
    for thread in threads {
        thread.join().unwrap();
    }
}

#[test]
fn host_fonts_match_the_same_font_in_a_directory() {
    let dir = std::env::temp_dir().join(format!("docxide-host-font-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    std::fs::write(dir.join("nb-hyphen-native.ttf"), FONT).unwrap();
    for rpr in ["", "<w:b/>"] {
        let doc = document(FONT_NAME, rpr);
        let supplied = run(&doc, &with_font()).pdf;
        let installed = run(&doc, &host().font_dirs([&dir])).pdf;
        assert!(embeds_font(&supplied));
        assert!(
            supplied == installed,
            "{rpr:?}: supplied and installed differ"
        );
    }
    std::fs::remove_dir_all(&dir).unwrap();
}

#[test]
fn host_fonts_reach_the_display_list_as_bytes() {
    let options = with_font().display_list(true);
    let out = run(&document(FONT_NAME, ""), &options);
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
fn host_requests_name_missing_fonts_until_supplied() {
    let name = FONT_NAME.to_string();
    let doc = document(&name, "");
    assert!(requested(&run(&doc, &host()), &name));
    assert!(!requested(&run(&doc, &with_font()), &name));
    assert!(requested(&run(&doc, &host()), &name));
}

#[test]
fn host_requests_leave_out_the_fallbacks_tried_instead() {
    let out = run(&document(FONT_NAME, "<w:i/>"), &host());
    let missing: Vec<_> = out
        .missing_fonts
        .iter()
        .map(|m| (m.name.as_str(), m.bold, m.italic))
        .collect();
    assert_eq!(missing, [(FONT_NAME, false, true)]);
}

#[test]
fn host_requests_exact_style_even_when_regular_is_supplied() {
    let name = FONT_NAME.to_string();
    let out = run(&document(&name, "<w:b/>"), &with_font());
    assert!(out.missing_fonts.iter().any(|m| m.bold && m.name == name));
}

#[test]
fn host_requests_cleared_by_failed_conversion() {
    run(&document(FONT_NAME, ""), &host());
    assert!(convert(b"not a zip", &ConvertOptions::new()).is_err());
    assert!(
        run(&document(FONT_NAME, ""), &with_font())
            .missing_fonts
            .is_empty()
    );
}

#[test]
fn host_requests_belong_to_the_converting_thread() {
    let options = [with_font(), host()];
    let barrier = Arc::new(Barrier::new(2));
    let threads: Vec<_> = options
        .into_iter()
        .enumerate()
        .map(|(i, options)| {
            let barrier = barrier.clone();
            std::thread::spawn(move || {
                barrier.wait();
                let out = run(&document(FONT_NAME, ""), &options);
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
        let plain = run(&docx, &options);
        let recorded = run(&docx, &options.clone().display_list(true));
        assert!(plain.display.is_none());
        assert!(recorded.display.is_some_and(|d| !d.pages.is_empty()));
        assert!(
            plain.pdf == recorded.pdf,
            "{case}: the display list moved PDF bytes"
        );
    }
}

#[test]
fn display_list_keeps_text_after_reusing_line_buffers() {
    use docxide_pdf::display::Item;

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    zip.start_file(
        "word/document.xml",
        zip::write::SimpleFileOptions::default(),
    )
    .unwrap();
    write!(zip, r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#).unwrap();
    for text in ["- --", "--- -- -"] {
        write!(zip, r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="{FONT_NAME}" w:hAnsi="{FONT_NAME}"/></w:rPr><w:t>{text}</w:t></w:r></w:p>"#).unwrap();
    }
    write!(zip, "</w:body></w:document>").unwrap();
    let input = zip.finish().unwrap().into_inner();
    let plain = run(&input, &with_font());
    let recorded = run(&input, &with_font().display_list(true));
    assert_eq!(plain.pdf, recorded.pdf);
    fn glyphs(items: &[Item]) -> usize {
        items
            .iter()
            .map(|item| match item {
                Item::Text(run) => run
                    .glyphs
                    .iter()
                    .flat_map(|glyph| glyph.text.chars())
                    .filter(|ch| !ch.is_whitespace())
                    .count(),
                Item::Group(items) => glyphs(items),
                _ => 0,
            })
            .sum()
    }
    let display = recorded.display.unwrap();
    assert_eq!(
        display
            .pages
            .iter()
            .map(|page| glyphs(&page.items))
            .sum::<usize>(),
        9
    );
}
