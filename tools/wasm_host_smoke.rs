//! Compiled by check-wasm-host.sh as a tiny host for the library's public API.
use docxide_pdf::convert_docx_bytes_to_pdf_bytes;

#[unsafe(no_mangle)]
pub extern "C" fn smoke_test() -> usize {
    let pdf = convert_docx_bytes_to_pdf_bytes(include_bytes!("../tests/host/input.docx")).unwrap();
    assert!(pdf.starts_with(b"%PDF-"));
    assert!(convert_docx_bytes_to_pdf_bytes(b"invalid").is_err());
    pdf.len()
}
