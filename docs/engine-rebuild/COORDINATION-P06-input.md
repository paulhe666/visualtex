# P06 input and view work

This continuation lane added svg-view.ts and svg-view-review.mjs. The real browser view check passes with 18 zoom/DPI/scroll cases and zero raster difference against the core SVG for the Chinese K0 formula.

This lane is now adding web/host/input-bridge.ts and its independent browser input regression. Other active implementation should avoid overwriting these two host files. The existing scene_queries.rs and runtime integration remain with the implementation lane.

Core changes already applied here: scene_validation.rs, scene_navigation.rs, their two calls in scene.rs, document.exportable use in layout.rs, and recursive has_formula_content in model/parser. All 187 core tests and 2 CLI tests passed after those edits. Source archive and audit are in checkpoints/CP05-geometry-review.

The input bridge will emit existing EditorAction values into the core owner. It must not create a second source/history store. Browser protocol composition tests must remain distinguished from physical Windows IME acceptance.
