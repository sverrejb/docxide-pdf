//! Compiled by check-wasm-host.sh as a tiny host for the library's public API.
use docxide_pdf::{
    clear_fonts, convert_docx_bytes_to_pdf_bytes, register_font, take_font_requests,
    use_system_fonts,
};

const FONT_NAME: &str = "NB Hyphen Fixture native";

fn requested() -> bool {
    take_font_requests()
        .iter()
        .any(|r| r.names.iter().any(|n| n == FONT_NAME))
}

const INPUT: &[u8] = include_bytes!("../tests/host/input.docx");

#[unsafe(no_mangle)]
pub extern "C" fn smoke_test() -> usize {
    use_system_fonts(false);
    clear_fonts();
    let fallback = convert_docx_bytes_to_pdf_bytes(INPUT).unwrap();
    assert!(requested());
    assert_eq!(
        register_font(include_bytes!("../src/fonts/testdata/nb-hyphen-native.ttf").to_vec()),
        1
    );
    let pdf = convert_docx_bytes_to_pdf_bytes(INPUT).unwrap();
    assert!(pdf.starts_with(b"%PDF-"));
    // The registered face is embedded; the standard-14 fallback is not.
    assert!(pdf.len() > fallback.len());
    assert!(!requested());
    assert!(convert_docx_bytes_to_pdf_bytes(b"invalid").is_err());
    assert!(take_font_requests().is_empty());
    clear_fonts();
    pdf.len()
}
