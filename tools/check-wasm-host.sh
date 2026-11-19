#!/usr/bin/env bash
# Requires the wasm32-unknown-unknown Rust target and Node.js.
set -euo pipefail
cd "$(dirname "$0")/.."
cargo build --locked --no-default-features --target wasm32-unknown-unknown --example wasm_host_smoke
node tools/check-wasm-host.mjs "${CARGO_TARGET_DIR:-target}/wasm32-unknown-unknown/debug/examples/wasm_host_smoke.wasm"
