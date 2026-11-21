//! Compiled by check-wasm-host.sh as a tiny host for the library's public API.
use docxide_pdf::{
    Conversion, ConvertOptions, clear_fonts, convert, register_font, use_system_fonts,
};

const FONT_NAME: &str = "NB Hyphen Fixture native";

fn requested(out: &Conversion) -> bool {
    out.missing_fonts.iter().any(|m| m.name == FONT_NAME)
}

const INPUT: &[u8] = include_bytes!("../tests/host/input.docx");

#[unsafe(no_mangle)]
pub extern "C" fn smoke_test() -> usize {
    use_system_fonts(false);
    clear_fonts();
    let options = ConvertOptions::new();
    let fallback = convert(INPUT, &options).unwrap();
    assert!(requested(&fallback));
    assert_eq!(
        register_font(include_bytes!("../src/fonts/testdata/nb-hyphen-native.ttf").to_vec()),
        1
    );
    let out = convert(INPUT, &options).unwrap();
    assert!(out.pdf.starts_with(b"%PDF-"));
    // The registered face is embedded; the standard-14 fallback is not.
    assert!(out.pdf.len() > fallback.pdf.len());
    assert!(!requested(&out));
    assert!(convert(b"invalid", &options).is_err());
    clear_fonts();
    out.pdf.len()
}
