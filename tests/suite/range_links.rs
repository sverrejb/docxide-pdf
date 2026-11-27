//! Range links need no installed fonts, external PDF tools or private
//! fixtures: every run names a font that is never found.
use std::collections::BTreeMap;
use std::io::{Cursor, Write};

use docxide_pdf::{ConvertOptions, RangeLink};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// A letter page with one-inch margins.
const SECTION: &str = r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>"#;
const LONG: &str = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ";

fn document(body: &str) -> Vec<u8> {
    document_with(body, &[])
}

fn document_with(body: &str, parts: &[(&str, &[u8])]) -> Vec<u8> {
    document_in(body, SECTION, parts)
}

fn document_in(body: &str, section: &str, parts: &[(&str, &[u8])]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file("word/document.xml", options).unwrap();
    write!(
        zip,
        r#"<w:document xmlns:w="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>{body}{section}</w:body></w:document>"#
    )
    .unwrap();
    for (name, bytes) in parts {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn bookmark(id: u32, name: &str) -> String {
    format!(r#"<w:bookmarkStart w:id="{id}" w:name="{name}"/><w:bookmarkEnd w:id="{id}"/>"#)
}

fn run(text: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="Missing" w:hAnsi="Missing"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#
    )
}

fn bold(text: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="Missing" w:hAnsi="Missing"/><w:b/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#
    )
}

fn paragraph(content: &str) -> String {
    format!("<w:p>{content}</w:p>")
}

fn convert(input: &[u8], links: Vec<RangeLink>) -> Vec<u8> {
    let options = ConvertOptions::new().font_cache(false).range_links(links);
    let pdf = docxide_pdf::convert(input, &options).unwrap().pdf;
    assert!(pdf.starts_with(b"%PDF-"));
    pdf
}

/// A link annotation: its URI and rectangle (x1, y1, x2, y2), PDF space.
#[derive(Debug, Clone, PartialEq)]
struct Link {
    uri: String,
    rect: [f32; 4],
}

impl Link {
    fn width(&self) -> f32 {
        self.rect[2] - self.rect[0]
    }

    fn height(&self) -> f32 {
        self.rect[3] - self.rect[1]
    }
}

fn find(hay: &[u8], needle: &[u8]) -> Option<usize> {
    hay.windows(needle.len()).position(|w| w == needle)
}

fn number_after(text: &str, key: &str) -> Option<usize> {
    let rest = &text[text.find(key)? + key.len()..];
    rest.trim_start()
        .split(|c: char| !c.is_ascii_digit())
        .next()?
        .parse()
        .ok()
}

/// Every object of a docxide PDF (plain or in an object stream), by number.
fn objects(pdf: &[u8]) -> BTreeMap<usize, String> {
    let mut objects = BTreeMap::new();
    let mut at = 0;
    while let Some(start) = find(&pdf[at..], b" 0 obj\n").map(|i| at + i) {
        let number_start = pdf[..start]
            .iter()
            .rposition(|b| !b.is_ascii_digit())
            .map_or(0, |i| i + 1);
        let number: usize = std::str::from_utf8(&pdf[number_start..start])
            .unwrap()
            .parse()
            .unwrap();
        let body_start = start + 7;
        let end = find(&pdf[body_start..], b"endobj").unwrap() + body_start;
        let body = &pdf[body_start..end];
        let dict_end = find(body, b"stream\n").unwrap_or(body.len());
        let dict = String::from_utf8_lossy(&body[..dict_end]).into_owned();
        if dict.contains("/Type /ObjStm") {
            let data = &body[dict_end + 7..];
            let data = &data[..find(data, b"\nendstream").unwrap()];
            let data = miniz_oxide::inflate::decompress_to_vec_zlib(data).unwrap();
            let count = number_after(&dict, "/N").unwrap();
            let first = number_after(&dict, "/First").unwrap();
            let header: Vec<usize> = String::from_utf8_lossy(&data[..first])
                .split_whitespace()
                .map(|n| n.parse().unwrap())
                .collect();
            for i in 0..count {
                let offset = first + header[2 * i + 1];
                let next = header.get(2 * i + 3).map_or(data.len(), |o| first + o);
                objects.insert(
                    header[2 * i],
                    String::from_utf8_lossy(&data[offset..next]).into_owned(),
                );
            }
        } else {
            objects.insert(number, dict);
        }
        at = end;
    }
    objects
}

/// The link annotations of each page, in order.
fn links(pdf: &[u8]) -> Vec<Vec<Link>> {
    let objects = objects(pdf);
    // Pages are numbered in page order.
    objects
        .values()
        .filter(|o| o.contains("/Type /Page") && !o.contains("/Type /Pages"))
        .map(|page| {
            let Some(at) = page.find("/Annots [") else {
                return Vec::new();
            };
            let list = &page[at + 9..at + page[at..].find(']').unwrap()];
            let refs: Vec<usize> = list
                .split_whitespace()
                .step_by(3)
                .map(|n| n.parse().unwrap())
                .collect();
            refs.iter()
                .map(|r| &objects[r])
                .filter(|a| a.contains("/Subtype /Link"))
                .map(|a| {
                    let rect = &a[a.find("/Rect [").unwrap() + 7..];
                    let rect: Vec<f32> = rect[..rect.find(']').unwrap()]
                        .split_whitespace()
                        .map(|v| v.parse().unwrap())
                        .collect();
                    let uri = &a[a.find("/URI (").unwrap() + 6..];
                    Link {
                        uri: uri[..uri.find(')').unwrap()].to_string(),
                        rect: [rect[0], rect[1], rect[2], rect[3]],
                    }
                })
                .collect()
        })
        .collect()
}

fn only<'a>(page: &'a [Link], uri: &str) -> Vec<&'a Link> {
    page.iter().filter(|l| l.uri == uri).collect()
}

#[test]
fn range_links_cover_text_between_bookmarks_inside_a_line() {
    // "Fula|no|," — the inner range ends inside a word, the outer one
    // covers the whole word with its comma.
    let body = paragraph(&format!(
        "{}{}{}{}{}{}{}{}{}",
        run("Hello "),
        bookmark(1, "word"),
        run("Fula"),
        bookmark(2, "inner"),
        run("no"),
        bookmark(3, "inner_end"),
        run(","),
        bookmark(4, "word_end"),
        run(" and more text after it."),
    ));
    let pdf = convert(
        &document(&body),
        vec![
            RangeLink::new("word", "word_end", "x:word"),
            RangeLink::new("inner", "inner_end", "x:inner"),
        ],
    );
    let pages = links(&pdf);
    assert_eq!(pages.len(), 1);
    let (word, inner) = (only(&pages[0], "x:word"), only(&pages[0], "x:inner"));
    assert_eq!((word.len(), inner.len()), (1, 1));
    let (word, inner) = (word[0], inner[0]);
    // After "Hello " at the left margin, on the first line.
    assert!(word.rect[0] > 72.0 + 10.0, "{word:?}");
    assert!(word.rect[3] < 792.0 - 72.0 + 1.0 && word.rect[3] > 792.0 - 72.0 - 20.0);
    // "no" lies inside "Fulano,", on the same line, narrower than the word.
    assert!(inner.rect[0] > word.rect[0] + 5.0, "{inner:?} {word:?}");
    assert!(inner.rect[2] < word.rect[2] - 1.0, "{inner:?} {word:?}");
    assert_eq!((inner.rect[1], inner.rect[3]), (word.rect[1], word.rect[3]));
    assert!(inner.width() < word.width() / 2.0);
    // A glyph box: about the font size tall.
    assert!((word.height() - 12.0).abs() < 2.5, "{word:?}");
}

#[test]
fn range_links_span_paragraphs_pages_and_bold_headings() {
    // Bookmarks between paragraphs: a bold heading and enough text for two
    // pages.
    let body = format!(
        "{}{}{}{}{}",
        bookmark(1, "block"),
        paragraph(&bold("A bold heading")),
        paragraph(&run(&LONG.repeat(40))),
        bookmark(2, "block_end"),
        paragraph(&run("After the block.")),
    );
    let pdf = convert(
        &document(&body),
        vec![RangeLink::new("block", "block_end", "x:block").heading_uri("x:heading")],
    );
    let pages = links(&pdf);
    assert!(pages.len() >= 2, "{pages:?}");
    for page in &pages[..2] {
        let rects = only(page, "x:block");
        // First line, the lines in between, last line.
        assert_eq!(rects.len(), 3, "{page:?}");
        assert!(rects[0].rect[1] >= rects[1].rect[3] - 0.5);
        assert!(rects[1].rect[1] >= rects[2].rect[3] - 0.5);
        assert!(rects[1].height() > 100.0);
    }
    // The heading link covers the bold line only, on the first page.
    let heading = only(&pages[0], "x:heading");
    assert_eq!(heading.len(), 1);
    assert_eq!(heading[0].rect, only(&pages[0], "x:block")[0].rect);
    assert!(only(&pages[1], "x:heading").is_empty());
    // "After the block." is not part of it.
    let last_page = pages
        .iter()
        .rposition(|p| !only(p, "x:block").is_empty())
        .unwrap();
    assert_eq!(last_page, 1);
}

#[test]
fn range_links_skip_headings_that_are_not_bold() {
    let body = format!(
        "{}{}{}",
        paragraph(&format!("{}{}", bookmark(1, "s"), run("Plain first line"))),
        paragraph(&run("Second")),
        paragraph(&bookmark(2, "e")),
    );
    let pdf = convert(
        &document(&body),
        vec![RangeLink::new("s", "e", "x:block").heading_uri("x:heading")],
    );
    let pages = links(&pdf);
    assert_eq!(only(&pages[0], "x:block").len(), 2);
    assert!(only(&pages[0], "x:heading").is_empty());
}

#[test]
fn range_links_stay_inside_their_table_cell() {
    let cell = |content: &str| {
        format!(r#"<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>{content}</w:tc>"#)
    };
    let table = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid><w:tr>{}{}</w:tr></w:tbl>"#,
        cell(&paragraph(&format!(
            "{}{}{}",
            bookmark(1, "cell"),
            run("Left cell"),
            bookmark(2, "cell_end")
        ))),
        cell(&paragraph(&run("Right cell"))),
    );
    let pdf = convert(
        &document(&format!("{table}{}", paragraph(""))),
        vec![RangeLink::new("cell", "cell_end", "x:cell")],
    );
    let rects = only(&links(&pdf)[0], "x:cell")
        .into_iter()
        .cloned()
        .collect::<Vec<_>>();
    assert_eq!(rects.len(), 1);
    // The first column ends 234pt into the text area.
    assert!(
        rects[0].rect[0] >= 72.0 && rects[0].rect[2] < 72.0 + 234.0,
        "{rects:?}"
    );
}

#[test]
fn range_links_cover_a_picture_only_range() {
    let mut png = Vec::new();
    image::RgbImage::new(4, 2)
        .write_to(&mut Cursor::new(&mut png), image::ImageFormat::Png)
        .unwrap();
    // A 100x50pt inline picture.
    let picture = r#"<w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#;
    let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>"#;
    let picture = paragraph(&format!(
        "{}{picture}{}",
        bookmark(1, "print"),
        bookmark(2, "print_end")
    ));
    // In the body and in a table cell.
    let cell = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>{picture}</w:tc></w:tr></w:tbl><w:p/>"#
    );
    for body in [picture.clone(), cell] {
        let input = document_with(
            &body,
            &[
                ("word/_rels/document.xml.rels", rels),
                ("word/media/image1.png", &png),
            ],
        );
        let pdf = convert(
            &input,
            vec![RangeLink::new("print", "print_end", "x:print")],
        );
        let rects = links(&pdf)[0].clone();
        assert_eq!(rects.len(), 1, "{rects:?}");
        assert_eq!(rects[0].uri, "x:print");
        assert!((rects[0].width() - 100.0).abs() < 0.5, "{rects:?}");
        assert!((rects[0].height() - 50.0).abs() < 0.5, "{rects:?}");
        assert!((rects[0].rect[0] - 72.0).abs() < 6.0, "{rects:?}");
    }
}

#[test]
fn range_links_follow_vertical_alignment() {
    let body = paragraph(&format!(
        "{}{}{}",
        bookmark(1, "s"),
        run("Centred"),
        bookmark(2, "e")
    ));
    let link = || vec![RangeLink::new("s", "e", "x:s")];
    let top = convert(&document(&body), link());
    let centred_section =
        SECTION.replace("</w:sectPr>", r#"<w:vAlign w:val="center"/></w:sectPr>"#);
    let centred = convert(&document_in(&body, &centred_section, &[]), link());
    let (top, centred) = (&links(&top)[0][0], &links(&centred)[0][0]);
    // Moved down with the text by half the empty page.
    assert!(top.rect[3] - centred.rect[3] > 200.0, "{top:?} {centred:?}");
    assert!((top.height() - centred.height()).abs() < 0.01);
}

#[test]
fn range_links_left_out_change_nothing() {
    let body = format!(
        "{}{}",
        paragraph(&format!(
            "{}{}{}",
            bookmark(1, "s"),
            run("Text"),
            bookmark(2, "e")
        )),
        paragraph(&run(&LONG.repeat(3)))
    );
    let input = document(&body);
    let plain = convert(&input, Vec::new());
    assert!(links(&plain).iter().all(Vec::is_empty));
    // Bookmarks that are missing, or out of order, link nothing.
    let unmatched = convert(
        &input,
        vec![
            RangeLink::new("missing", "e", "x:a"),
            RangeLink::new("e", "s", "x:b"),
        ],
    );
    assert_eq!(plain, unmatched);
}

#[test]
fn range_links_in_caps_do_not_include_other_words_in_the_merged_run() {
    for property in ["<w:caps/>", "<w:smallCaps/>", "<w:caps/><w:smallCaps/>"] {
        let styled = |value: &str| run(value).replace("</w:rPr>", &format!("{property}</w:rPr>"));
        let body = paragraph(&format!(
            "{}{}{}{}{}",
            styled("before "),
            bookmark(1, "s"),
            styled("Alice Example"),
            bookmark(2, "e"),
            styled(" after")
        ));
        let pdf = convert(&document(&body), vec![RangeLink::new("s", "e", "x:field")]);
        let pages = links(&pdf);
        let range = only(&pages[0], "x:field");
        assert_eq!(range.len(), 1);
        assert!(range[0].rect[0] > 100.0, "{property}: {range:?}");
        assert!(range[0].width() < 120.0, "{property}: {range:?}");
    }
}

const PICTURE_RELS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>"#;
/// 1pt in EMU.
const EMU: i64 = 12700;

fn png() -> Vec<u8> {
    let mut png = Vec::new();
    image::RgbImage::new(4, 2)
        .write_to(&mut Cursor::new(&mut png), image::ImageFormat::Png)
        .unwrap();
    png
}

/// The only link of `body`'s first page, a range `print`..`print_end`.
fn picture_link(body: &str) -> Link {
    let png = png();
    let input = document_with(
        body,
        &[
            ("word/_rels/document.xml.rels", PICTURE_RELS),
            ("word/media/image1.png", &png),
        ],
    );
    let pdf = convert(
        &input,
        vec![RangeLink::new("print", "print_end", "x:print")],
    );
    let rects = links(&pdf)[0].clone();
    assert_eq!(rects.len(), 1, "{rects:?}");
    assert_eq!(rects[0].uri, "x:print");
    rects[0].clone()
}

/// A run with a 100x50pt picture floating `x`, `y` pt from the page's or
/// the column's and paragraph's (`relative`) origin.
fn anchored_picture(relative: (&str, &str), x: i64, y: i64) -> String {
    format!(
        r#"<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="{}"><wp:posOffset>{}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="{}"><wp:posOffset>{}</wp:posOffset></wp:positionV><wp:extent cx="1270000" cy="635000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="p"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#,
        relative.0,
        x * EMU,
        relative.1,
        y * EMU
    )
}

fn in_range(content: &str) -> String {
    format!(
        "{}{content}{}",
        bookmark(1, "print"),
        bookmark(2, "print_end")
    )
}

fn assert_box(link: &Link, x: f32, y_top: f32) {
    // PDF space: the letter page is 792pt high.
    let expected = [x, 792.0 - y_top - 50.0, x + 100.0, 792.0 - y_top];
    for (got, want) in link.rect.iter().zip(expected) {
        assert!((got - want).abs() < 0.5, "{link:?} != {expected:?}");
    }
}

#[test]
fn range_links_cover_a_floating_picture() {
    let picture = anchored_picture(("page", "page"), 200, 300);
    let link = picture_link(&format!(
        "{}{}",
        paragraph(&run("Before")),
        paragraph(&in_range(&picture))
    ));
    assert_box(&link, 200.0, 300.0);
}

#[test]
fn range_links_keep_a_floating_picture_apart_from_the_lines() {
    // Five short lines and, anchored in the third, a picture floating to the
    // right of them: merging it with the middle lines would cover the page
    // between the two.
    let lines: String = (1..=5)
        .map(|i| {
            let picture = if i == 3 {
                anchored_picture(("page", "page"), 400, 90)
            } else {
                String::new()
            };
            paragraph(&format!("{}{picture}", run(&format!("Line {i}"))))
        })
        .collect();
    let png = png();
    let input = document_with(
        &format!(
            "{}{lines}{}",
            bookmark(1, "print"),
            bookmark(2, "print_end")
        ),
        &[
            ("word/_rels/document.xml.rels", PICTURE_RELS),
            ("word/media/image1.png", &png),
        ],
    );
    let pdf = convert(
        &input,
        vec![RangeLink::new("print", "print_end", "x:print")],
    );
    let page = &links(&pdf)[0];
    let rects = only(page, "x:print");
    assert_eq!(rects.len(), 4, "{rects:?}");
    let (pictures, text): (Vec<&&Link>, Vec<&&Link>) =
        rects.iter().partition(|link| link.rect[0] > 300.0);
    assert_eq!(pictures.len(), 1, "{rects:?}");
    assert_box(pictures[0], 400.0, 90.0);
    assert!(text.iter().all(|link| link.rect[2] < 300.0), "{rects:?}");
}

#[test]
fn range_links_cover_a_floating_picture_in_a_table_cell() {
    let picture = paragraph(&in_range(&anchored_picture(("column", "paragraph"), 10, 0)));
    let body = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>{picture}</w:tc></w:tr></w:tbl><w:p/>"#
    );
    let link = picture_link(&body);
    assert!((link.width() - 100.0).abs() < 0.5, "{link:?}");
    assert!((link.height() - 50.0).abs() < 0.5, "{link:?}");
    // Inside the cell, at the top of the page's text area.
    assert!(
        link.rect[0] > 72.0 && link.rect[2] < 72.0 + 234.0 + 1.0,
        "{link:?}"
    );
    assert!(
        link.rect[3] <= 720.0 + 0.5 && link.rect[3] > 690.0,
        "{link:?}"
    );
}

#[test]
fn range_links_cover_a_floating_vml_picture() {
    let picture = r#"<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="p" style="position:absolute;margin-left:200pt;margin-top:300pt;width:100pt;height:50pt;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page"><v:imagedata r:id="rId1" o:title="" xmlns:o="urn:schemas-microsoft-com:office:office"/></v:shape></w:pict></w:r>"#;
    let link = picture_link(&format!(
        "{}{}",
        paragraph(&run("Before")),
        paragraph(&in_range(picture))
    ));
    assert_box(&link, 200.0, 300.0);
}

#[test]
fn range_links_cover_a_picture_alone_in_a_text_box() {
    let inline = r#"<w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/><wp:docPr id="2" name="p"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#;
    let story = paragraph(&in_range(inline));
    let text_box = format!(
        r#"<w:r><w:drawing><wp:anchor xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>{y}</wp:posOffset></wp:positionV><wp:extent cx="{w}" cy="{h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="Box"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>{story}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#,
        x = 200 * EMU,
        y = 300 * EMU,
        w = 200 * EMU,
        h = 100 * EMU,
    );
    let link = picture_link(&paragraph(&text_box));
    assert!((link.width() - 100.0).abs() < 0.5, "{link:?}");
    assert!((link.height() - 50.0).abs() < 0.5, "{link:?}");
    // In the box, 200pt from the page's left and 300pt from its top.
    assert!((link.rect[0] - 200.0).abs() < 2.0, "{link:?}");
    assert!(
        link.rect[3] <= 792.0 - 300.0 + 0.5 && link.rect[3] > 792.0 - 300.0 - 20.0,
        "{link:?}"
    );
}
