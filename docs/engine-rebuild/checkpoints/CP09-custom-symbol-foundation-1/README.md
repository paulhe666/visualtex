# CP09 custom-symbol foundation 1

Stable non-Git checkpoint for the P09 self-owned custom-symbol foundation.

Source SHA256: `510dd3ec3666db83e052cdff0065f5babd479f6ced32ca91cc2aaf07acf77186`

Verification receipt: `engine/target/p09-verification/2026-09-14T06-17-44-803Z/report.json`

The stable 14-step P09 gate passed with `sourceUnchanged=true`. It includes the complete P08 isolated engine/sandbox regression (including the real Windows IME gate), command inventory/state/browser, built-in macro parity/browser, matrix-family parity/browser, `\middle` parity/browser, the input bridge, trusted-browser structural clipboard, custom-symbol Native/WASM parity, and owned Edge custom-symbol acceptance.

This slice freezes exact legacy custom-symbol definitions into immutable content-addressed resources instead of looking up same-name library entries at render time. Supported vector geometry is rendered by the Rust core and the self-owned SVG scene/view, not MathLive CSS masks or MathJax. Mathematical role, authored metrics, operator limits behavior, source round-trip, structural insertion, duplicate/undo/redo, source patches, rich clipboard asset transfer, natural-size SVG export and document isolation are covered. Clipboard transfers definitions without live NodeId/session identity and rejects conflicting or tampered snapshots.

The current foundation intentionally fails closed for legacy artwork capabilities that are not yet connected to the self-owned vector compiler, including browser/system-font `text` artwork, clipping, erase/boolean subtraction, general stroked paths/polygons/ellipses, SVG arc segments and rounded rectangles. The exact legacy definition is retained and the symbol remains non-exportable rather than drawing a partial or approximate result. Therefore this checkpoint is **not** P09 complete.

No production application or Office/VSTO path was switched. No user Word document was touched. P00 protection remains red because of independent Office/VSTO/evidence changes outside this engine task. No commit or push belongs to this checkpoint.
