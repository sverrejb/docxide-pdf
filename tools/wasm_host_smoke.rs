//! Compiled by check-wasm-host.sh as a tiny host for the library's public API.
use docxide_pdf::{
    clear_fonts, convert_docx_bytes_to_pdf_bytes, register_font, take_bookmark_positions,
    take_font_requests, use_system_fonts,
};

#[unsafe(no_mangle)]
pub extern "C" fn smoke_test() -> usize {
    use_system_fonts(false);
    clear_fonts();
    assert_eq!(
        register_font(include_bytes!("../src/fonts/testdata/nb-hyphen-native.ttf").to_vec()),
        1
    );
    let pdf = convert_docx_bytes_to_pdf_bytes(include_bytes!("../tests/host/input.docx")).unwrap();
    assert!(pdf.starts_with(b"%PDF-"));
    assert!(
        !take_font_requests()
            .iter()
            .any(|r| r.names.iter().any(|n| n == "NB Hyphen Fixture native"))
    );
    let positions = take_bookmark_positions();
    assert_eq!(positions.len(), 2);
    for name in ["wasm_body", "wasm_cell"] {
        let position = positions.iter().find(|p| p.name == name).unwrap();
        assert_eq!(position.page, 0);
        assert!(position.top > position.bottom);
    }
    assert!(convert_docx_bytes_to_pdf_bytes(b"invalid").is_err());
    assert!(take_font_requests().is_empty());
    assert!(take_bookmark_positions().is_empty());
    clear_fonts();
    pdf.len()
}
