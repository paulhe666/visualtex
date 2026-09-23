# P07.1 — square roots and indexed roots

The independent engine supports `\\sqrt{...}` and `\\sqrt[index]{...}` through parsing, lossless/canonical LaTeX, structural editing, navigation, Unicode-aware text children, source projection, history, clipboard duplication, font-driven layout and vector SVG. Production MathLive/MathJax entrypoints remain unchanged.

The radicand is an editable row. The optional degree is a separate editable row with its own identity; absence and an explicitly empty degree are different. Both remain navigable when empty, but missing/empty required slots cannot be exported as completed formulas. Horizontal navigation visits degree then radicand. The root is selected by the first boundary delete and removed by the next; undo restores identities. Inserting a root around a selection preserves the selected node identities.

Optional arguments use brace-protected TeX delimiter scanning. A nested indexed root must be grouped inside an outer index. Structural insertion automatically groups only the affected degree, e.g. `\\sqrt[{n\\sqrt[3]{z}}]{x}`; repeated edits/reopen do not accumulate groups. Incomplete or unprotected optional delimiters retain the source and produce diagnostics. Unknown commands remain unsupported rather than becoming an implicit fallback.

The radicand uses the ambient math style with cramped scripts. The degree uses scriptscript style. Radical rule thickness, gaps, extra ascender, index kerns and index raise percentage come from the injected OpenType MATH font. `stretch::vertical` chooses a designed glyph variant or assembles real bottom-to-top parts with bounded extender repetition and valid connector overlaps. No glyph is stretched independently on one axis. Unsupported or malformed font constructions fail explicitly. Screen and SVG share the same scene; zoom and pixel density do not change semantic geometry.

References used for implementation:
- https://learn.microsoft.com/en-us/typography/opentype/spec/math
- https://www.w3.org/TR/mathml-core/#radicals-msqrt-mroot
- Locked local ttf-parser 0.25.1 MATH API.

Tests: `crates/core/tests/radicals.rs` (10 tests, including 200 seeded expressions and 1200 source edits), `stretch.rs` (3 tests), `tools/layout-parity.mjs` (root geometry and structural editing in actual Native/WASM), `tools/radical-browser.mjs` (actual independent sandbox controls and pointer/key events). The failing nested-index-bracket regression was reproduced before correcting the implementation. Existing unsupported-layout tests now use an unsupported accent, while the old root fixture is retained as a positive test; no negative protection test was removed.

Not complete: other P07 structures, matrices, general macro expansion, automatic paired-delimiter editing, full command registry coverage, MathML/Office target compatibility, production main app, Office UI, WebView2 and macOS host acceptance. These remain explicit later gates; P07.1 is not P07 completion.
