//! The integration suites as one test binary, so the shared harness in tests/common compiles
//! once instead of once per suite (it was most of the rebuild after a src/ edit). One suite:
//! `cargo test --test suite visual_comparison` (the module name is a test-name filter).
#[path = "../common/mod.rs"]
mod common;

mod accessibility;
mod file_size;
mod font_validation;
mod image_count;
mod libreoffice_comparison;
mod page_geometry;
mod speed;
mod text_boundary;
mod visual_comparison;
