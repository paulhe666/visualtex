# CP09 custom symbols foundation 1

This is a non-Git recovery checkpoint for the isolated VisualTeX engine workspace. It is not a P09-complete claim and does not authorize production/main-app or Office switching.

Stable verification receipt: `engine/target/p09-verification/2026-09-14T06-07-44-117Z/report.json`.

Stable source SHA256: `ad92fd7e191482b4a4f9c9e1e21e1905a0e82763f77248023f2eace3b0dbc168`.

The 14-step stable gate passed with sourceUnchanged=true. It includes the complete P08 regression gate, current command-registry audit, command input, bounded compatibility macros, matrix-family/language extensions, middle delimiters, input bridge, structural clipboard, custom-symbol Native/WASM parity and owned-Edge custom-symbol acceptance.

Custom-symbol foundation completed at this checkpoint:

- Immutable legacy custom-symbol definition snapshots and SHA-256 content identity.
- Native `CustomSymbol` nodes with role-aware spacing and operator limits behavior.
- Exact source preservation; later same-name user-library changes cannot silently redraw an existing formula.
- Bounded Rust vector compiler for the currently supported legacy path/shape subset.
- Separate artwork scene resources and explicit paint order; custom artwork never masquerades as a font glyph.
- Rust SVG export and browser SvgSceneView render the same validated vector scene.
- Rich clipboard carries identity-free frozen definitions into a session with no prior user-symbol catalog; live NodeIds/session identities are never copied.
- Conflicting same-name/different-definition resources fail atomically.
- Native/WASM parity: 44 requests, 14 scenes, maximum numeric difference 0, exact SVG parity.
- Owned Edge: three definitions, 27 zoom/DPR pointer cases, trusted keyboard/button insertion and trusted rich clipboard transfer, isolation sentinel unchanged.

Explicit remaining advanced legacy artwork for the next slice: unflattened system-font `text` artwork, clip rectangles, erase/boolean subtraction, SVG arc segments, general stroke semantics and rounded/non-miter rectangle geometry. These definitions are retained but fail closed as unsupported; no partial symbol is drawn.

No Office actions, production-app writes, commits or pushes are part of this checkpoint. P00 remains red because unrelated Office/VSTO/evidence changes exist in the shared checkout and must not be touched.