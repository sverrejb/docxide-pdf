import { readFileSync } from "node:fs";

const module = await WebAssembly.compile(readFileSync(process.argv[2]));
const imports = {};
// The bare library needs no JS clock, filesystem, or randomness. Retained
// wasm-bindgen descriptors must be unreachable in the conversion path.
for (const { module: namespace, name, kind } of WebAssembly.Module.imports(module)) {
  if (kind !== "function") throw new Error(`Unexpected import: ${namespace}.${name}`);
  imports[namespace] ??= {};
  imports[namespace][name] = () => { throw new Error(`Unexpected host call: ${namespace}.${name}`); };
}
const instance = await WebAssembly.instantiate(module, imports);
const size = instance.exports.smoke_test();
if (!(size > 0)) throw new Error("Conversion did not produce a PDF");
console.log(`Wasm host smoke test passed: ${size} PDF bytes; no host imports called.`);
