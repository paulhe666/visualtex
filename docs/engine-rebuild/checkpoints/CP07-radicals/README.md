# CP07.1 — verified autonomous radical slice

2026-09-13 14:50 +08:00. Base checkout `reference-3d207d7`, HEAD `00b48e17b736c7e57a0a72e19e271df82906f89d`. No Git commit/push or installed product change.

Implemented square roots and indexed roots across owned mathematical structure, source grammar/projection, node identity, argument navigation, copy/delete/history, OpenType MATH metric input, designed variants and bounded genuine glyph assemblies, same-scene SVG rendering, host types, accessible position descriptions and independent sandbox buttons.

A targeted negative test first reproduced missing brace protection when another indexed root is inserted inside a root index. The source delimiter scanner and serializer were corrected; the failing case remains a positive regression. Incomplete/unknown source is retained and not falsely marked export-ready. The old root sample was retained and promoted to supported; unsupported accent cases preserve the earlier failure assertions.

Verification receipt: `engine/target/verification/2026-09-13T06-49-22-750Z/report.json`.
Verified source SHA256: `4c40aed2f5a0a1fe865dba61123697b406884ea1e39c2c9bfdaf957b366b8836`.
All six verification steps passed; no inputs changed during the run:

- Build + 213 native Rust tests, 0 failed; Native/WASM protocol 2151, source-session 222, source edits 440, editing/history 473 cases.
- Layout parity: 270 requests and 158 scenes, including roots in all four math styles, nested degree roots, Chinese radicands, large assemblies, shared history and pixel-size changes. Numeric difference 0 at the fixed font resources; exact SVG parity.
- Independent sandbox full K0 editing/source/export regression.
- Actual root UI controls: construct nested indexed and square roots from blank, edit the degree by actual pointer position, undo/redo, retain an incomplete optional argument, and export a large root built from 8 genuine font pieces. Screenshot `engine/target/radical-browser/tall-root.png` was opened and visually inspected.
- Input fault injection 15/15; SVG zoom/DPR/scroll geometry 18 cases and vector/raster agreement fixture.
- Office/production protection: 12 groups, 12 artifacts, 65 original dirty working-file bytes; approved historical Git transition acknowledged.

Root-native tests comprise 10 root tests (including 200 generated nested expressions and 1200 source disturbances) and 3 glyph-construction tests. These include fresh-parse comparison, copied node disjointness, invalid delimiter preservation, font constraints and natural pixel geometry independence.

The validation runner `engine/tools/verify-current.mjs` writes only uniquely named owned logs under engine/target/verification and does not commit or change refs. Each step records exit code, duration, log hash and actual source inventory. Existing independent browser tests now compile current TypeScript rather than using stale test output directories.

Actual Windows IME was verified earlier in the separate immutable `CP06-native-ime-verified` checkpoint. This root browser test uses browser keys/pointer and does not falsely assert a new Windows IME run.

Gate: P07.1 roots may advance to P07.2 delimiters in the isolated engine. P07 as a whole is not complete. No production app/default switch, Office output/interactive/hidden-converter integration, installer or release has been accepted. Mac/Web production paths remain untouched. Mathematical MathML/Office target export is a later compatibility gate, not proven by this SVG slice.
