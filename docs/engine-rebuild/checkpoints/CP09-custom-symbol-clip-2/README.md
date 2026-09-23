# CP09 custom-symbol clip 2

Stable non-Git checkpoint for the P09 advanced custom-symbol artwork slice through `clipRect`, with the P09 custom-symbol specification updated to match the implemented capability set.

Source SHA256: `e63b63116800c082db207ead1ba74e049177aa03e0daf066be6b7a558c6c40ce`

Verification receipt: `engine/target/p09-verification/2026-09-14T07-26-06-595Z/report.json`

Custom-symbol Native/WASM receipt: `engine/target/custom-symbol-parity-review/2026-09-14T07-28-41-349Z/report.json`

Owned Edge receipt: `engine/target/custom-symbol-browser-review/2026-09-14T07-28-43-191Z/report.json`

This checkpoint includes deterministic SVG A/a arc conversion, rounded rectangles, real SVG fill/stroke/stroke-width/line-cap/line-join semantics under affine transforms, and `clipRect`. Imported/base `transform.matrix` remains inside the clip while user translate/scale/skew/rotate remains outside it: outer user transform -> clip rectangle -> inner/base matrix. Rust SVG export and the self-owned browser view preserve that same nesting with generated SVG nodes.

The strict 14-step P09 gate passed with `sourceUnchanged=true`. Custom-symbol Native/WASM parity has max numeric difference 0 and exact SVG parity. Owned Edge acceptance includes the layered clip in the real DOM, 45 zoom x DPR pointer cases, trusted structural insertion, and rich clipboard transfer.

The P09 specification now records arcs, rounded rectangles, general affine stroke semantics and clip rectangles as supported. Remaining advanced artwork gaps are `erase`/boolean subtraction and `text`; text must bind a real verified font outline rather than use browser `font-family`, and erase must remain scoped to the same immutable custom asset rather than affecting neighboring formula glyphs.

`CP09-custom-symbol-clip-1` remains untouched as an earlier archive whose implementation was valid but whose included P09 specification still described clip/arc/stroke/rounded-rect as unsupported. This checkpoint supersedes it as the stable clip point.

No production application or Office/VSTO path was switched. No user Word document was touched. P00 protection remains red because of independent parallel Office/VSTO/evidence work. No commit or push belongs to this checkpoint.
