#!/usr/bin/env bash
# Requires the wasm32-unknown-unknown Rust target and Node.js.
set -euo pipefail
cd "$(dirname "$0")/.."
target_dir="${CARGO_TARGET_DIR:-target}"
wasm_dir="$target_dir/wasm32-unknown-unknown/debug"
cargo build --locked --lib --no-default-features --target wasm32-unknown-unknown
rustc --edition=2024 --crate-type=cdylib --target wasm32-unknown-unknown \
    --extern "docxide_pdf=$wasm_dir/libdocxide_pdf.rlib" \
    -L "dependency=$wasm_dir/deps" -L "dependency=$target_dir/debug/deps" \
    tools/wasm_host_smoke.rs \
    -o "$wasm_dir/host-smoke.wasm"
node tools/check-wasm-host.mjs "$wasm_dir/host-smoke.wasm"
