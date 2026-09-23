# CP09 custom-symbol text 1

Stable non-Git checkpoint for P09 legacy custom-symbol `text` artwork using explicit verified font outlines, building on `CP09-custom-symbol-erase-1`.

Stable source SHA256: `9b37a2f47a5bf068bdacae7bd8534caa5bd19954e23d5c02d5db141e7d88a331`.

Verification receipt: `engine/target/p09-verification/2026-09-14T08-10-44-975Z/report.json`.

The complete 14-step P09 gate passed with `sourceUnchanged=true`. Custom-symbol Rust review passes 15/15. Native/WASM custom-symbol parity passes with 60 requests, 19 scenes, maximum numeric difference 0 and exact SVG parity. Owned Edge acceptance passes with seven frozen definitions, 13 artwork paths, 63 real zoom/DPR pointer cases, trusted keyboard/clipboard events and independent host validation.

Legacy text definitions keep their original `text`, `fontFamily`, `fontStyle`, `fontWeight`, `fontSize`, position and transforms. They are never rendered through browser/system `font-family`. Layout requires an exact `(fontFamily, fontStyle, fontWeight)` binding to a font already registered from real bytes. Font identity is content-addressed as SHA-256 plus face index. Rust shapes the text with rustybuzz and extracts real glyph outlines with ttf-parser before artwork enters the scene.

The final scene contains only vector path geometry and an `outlineFontId` provenance reference. Both Rust and the independent browser host require that ID to exist in the scene's verified font identity list and reject unresolved `textSource`, unknown font provenance, missing/wrong bindings or malformed text metadata. SVG/DOM output contains no `<text>`, no `font-family`, no Canvas/system-font fallback and no raw authored markup.

No production application or Office/VSTO path was switched. No user Word document was touched. P00 protection remains red because of independent external Office/VSTO/evidence work. No commit or push belongs to this checkpoint.
