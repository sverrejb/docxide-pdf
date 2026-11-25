//! Host APIs need no installed fonts, external PDF tools or private fixtures.
use std::io::{Cursor, Write};
use std::sync::{Arc, Barrier};

use docxide_pdf::display::FontSource;
use docxide_pdf::{Compression, Conversion, ConvertOptions, convert};

use crate::common::{convert_bytes_with as run, pdf_contains, pdf_streams};

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

/// A DOCX of `parts`: (ZIP entry name, contents).
fn docx(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, data) in parts {
        zip.start_file(*name, zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// One paragraph in `font`, with `rpr` added to its run properties.
fn document(font: &str, rpr: &str) -> Vec<u8> {
    let xml = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>{rpr}</w:rPr><w:t>-</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="4000" w:h="4000"/><w:pgMar w:top="400" w:bottom="400" w:left="400" w:right="400"/></w:sectPr></w:body></w:document>"#
    );
    docx(&[("word/document.xml", xml.as_bytes())])
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

/// 400 paragraphs in the test font: streams big enough for the zlib level to
/// change their size.
fn long_document() -> Vec<u8> {
    let paragraphs: String = (0..400)
        .map(|i| format!(r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="{FONT_NAME}" w:hAnsi="{FONT_NAME}"/></w:rPr><w:t>- -- --- {i} - -- ---</w:t></w:r></w:p>"#))
        .collect();
    let xml = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{paragraphs}</w:body></w:document>"#
    );
    docx(&[("word/document.xml", xml.as_bytes())])
}

/// Every stream of `pdf` inflated and sorted, except the cross-reference
/// stream, whose byte offsets follow the file's size.
fn inflated_streams(pdf: &[u8]) -> Vec<Vec<u8>> {
    let mut streams: Vec<Vec<u8>> = pdf_streams(pdf)
        .filter(|&(start, _)| {
            let dictionary = &pdf[start.saturating_sub(300)..start];
            !dictionary.windows(5).any(|w| w == b"/XRef")
        })
        .map(|(_, data)| data)
        .collect();
    streams.sort();
    streams
}

#[test]
fn fast_compression_writes_the_same_pdf_once_inflated() {
    let input = long_document();
    let with = |compression| run(&input, &with_font().compression(compression)).pdf;
    let smallest = with(Compression::Smallest);
    let fast = with(Compression::Fast);
    assert!(fast.len() > smallest.len());
    let streams = inflated_streams(&smallest);
    assert!(streams.len() > 3);
    assert_eq!(inflated_streams(&fast), streams);
}

/// One inline picture: `jpeg` shown at one inch square.
fn picture_document(jpeg: &[u8]) -> Vec<u8> {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill><pic:spPr><a:xfrm><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#;
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/p.jpg"/></Relationships>"#;
    docx(&[
        ("word/document.xml", xml.as_bytes()),
        ("word/_rels/document.xml.rels", rels.as_bytes()),
        ("word/media/p.jpg", jpeg),
    ])
}

#[test]
fn fast_compression_embeds_a_jpeg_as_stored() {
    let pixels =
        image::RgbImage::from_fn(64, 64, |x, y| image::Rgb([x as u8 * 4, y as u8 * 4, 90]));
    let mut jpeg = Vec::new();
    image::codecs::jpeg::JpegEncoder::new_with_quality(&mut jpeg, 95)
        .encode_image(&pixels)
        .unwrap();
    let input = picture_document(&jpeg);
    let with = |compression| run(&input, &host().compression(compression)).pdf;
    assert!(pdf_contains(&with(Compression::Fast), &jpeg));
    // The default re-encodes at Word's quality 75.
    assert!(!pdf_contains(&with(Compression::Smallest), &jpeg));
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
