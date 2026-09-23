# CP09 custom-symbol clip 1

Stable non-Git checkpoint for the P09 advanced custom-symbol artwork slice through `clipRect`.

Source SHA256: `5c3d93e13e2ec79728bef40d89a6a7032564d0ec5dce55cf0d3daeee958b8614`

Verification receipt: `engine/target/p09-verification/2026-09-14T07-21-56-250Z/report.json`

Custom-symbol Native/WASM receipt: `engine/target/custom-symbol-parity-review/2026-09-14T07-24-36-886Z/report.json`

Owned Edge receipt: `engine/target/custom-symbol-browser-review/2026-09-14T07-24-38-750Z/report.json`

This checkpoint keeps the previously frozen immutable custom-symbol foundation and adds deterministic SVG arc conversion, rounded rectangles, true SVG fill/stroke semantics under affine transforms, and `clipRect`. Imported artwork now preserves the intended three-stage transform structure: outer user transform -> clip rectangle -> inner/base matrix. The self-owned SVG view mirrors the same structure using generated SVG groups and clip paths; no authored SVG/HTML/CSS is injected.

The strict 14-step P09 gate passed with `sourceUnchanged=true`. The custom-symbol parity gate includes exact Native/WASM SVG parity with max numeric difference 0. The owned Edge gate includes the clipped imported path, 45 zoom x DPR pointer cases, trusted structural insertion and rich clipboard transfer.

Legacy `erase` boolean subtraction and `text` artwork remain intentionally unsupported at this checkpoint. Text must not fall back to browser `font-family`; later support must bind a real verified font outline. Erase must remain scoped to the same immutable custom asset and must never affect neighboring formula glyphs.

No production application or Office/VSTO path was switched. No user Word document was touched. P00 protection remains red because of independent parallel Office/VSTO/evidence work. No commit or push belongs to this checkpoint.
