# CP07 colors — tested bounded paint slice

Non-Git checkpoint, 2026-09-13. Baseline branch reference-3d207d7,
HEAD 00b48e17b736c7e57a0a72e19e271df82906f89d. No engine commit/push,
production engine switch, Office document change or package installation.

## Implementation

Authored textcolor wrappers and row-local color declarations, 19 base color names,
bounded named mixtures, rgb/RGB/HTML/gray inputs, source spelling/comment retention,
scoped paint, typed safe RGB scene data and vector fill export. Visible glyphs and
rules (fractions, roots, accents and borders) inherit color. Phantom subtrees still
emit no paint. Body editing, source patching, clone/delete, undo/redo, position maps,
readable location descriptions and the existing sandbox controls all participate.
No CSS author string or URL is accepted as paint. Unsupported source remains
explicitly unexportable rather than silently changed to black.

Full finite supported/unsupported scope is in `engine/specs/P07-colors.md`.
This slice is not a claim to implement all xcolor, all P07 or full MathLive removal.

## Executed verification

Stable-source receipt:
`engine/target/verification/2026-09-13T09-47-25-653Z/report.json`
Source SHA256:
`a2a160239b9e1dc3f091f33c8cca21f569742f075d2fb0a40c8abe88f98e09a7`
`sourceUnchanged: true` throughout the run.

All eleven engine/build/browser steps PASS. Strict overall status remains FAILED
at the final original P00 production-protection step. Do not convert this to a
globally green gate or treat this checkpoint as permission for Office integration.

- 300 Rust tests passed. Fourteen new color integration tests and two paint-value
  tests include 640 seeded source disturbances (0xc0102026), nested group restoration,
  glyph/dimension identity, colored rules, hidden paint exclusion, edits/history,
  copied IDs, source ranges, cache freshness and forged/wire color rejection.
- Actual Native/WASM protocol2151, source-session222, source-edit440, editor473
  cases, with100 history states, PASS.
- Layout parity3945 requests /2689 scenes PASS: numeric difference0 at fixed
  identical font resources, exact SVG and numerical paint.
- Existing K0/root/fence/accent/operator/style/spacing browser tests PASS.
- New colors browser PASS,12 trusted input events and16 explicitly synthetic
  malformed View-paint calls; all malformed values rejected without replacing the
  valid SVG. Mathematical selections/actions use real browser controls and keys.
- Input fault injection and18 zoom/DPR/scroll geometry cases PASS. The existing
  fixed vector/raster fixture remains mean/max channel difference0; this is not
  claimed as universal pixel identity for all formulas/colors/platforms.
- Separate independent layout review519 requests /360 scenes PASS.
- Separate Office contract snapshot11 sources /3 interfaces /10 routes /3 negative
  mutations PASS, writes0.

The first full receipt at09:44:59 failed its SVG View runner because the legacy
/view.js alias omitted the new relative paint.js module. The runner now serves
current host modules at actual /host paths. No geometry or rendering assertion was
removed. That failed receipt remains in target/verification as historical evidence.
The corrected full receipt above tests the actual fixed source, not stale assets.

Browser report: `engine/target/color-browser/report.json`.
Screenshot: `engine/target/color-browser/colors.png`, actually read and visually
inspected through Devspace: dark-blue integral/border, red root/fraction, hidden
blue content absent. Individual browser build sandbox-937b0fd650967d1f00337cf4,
Edge152.0.4191.62.

Native CLI SHA256:
`5c7f2ac0720e0ea585c39d71ef0085352d6715db58758acc08b7d44f44c927ed`
Node WASM SHA256:
`e86020afa7572a2ea10a02bf4710c21e211394bb51446ffdbfd60182d538887d`
Web WASM SHA256:
`979527ba8339a99b1a9826bdb34f4415c8dc42337a46a25d48df0d4410a28222`

## Protection and next steps

P00 guard09:48:48Z reports ThisAddIn, WordFormulaFormatConversion,
WordFormulaService.CopyPaste/FormatConversion/main, target-word.json,
ui-actions.ndjson, ui-uia.json, windowsSource/wordVsto/vstoShared groups and
historical-untracked-path changes outside this engine implementation. Original
baseline/guard assertions were kept. No external file was overwritten/deleted by
this continuation, and no writer identity is inferred from these differences.

No new physical Windows IME, Word/Ribbon, PowerPoint, WebView2 installed UI,
MathML/Office compatibility or final performance acceptance is asserted. Font
binaries and generated builds are excluded from the source archive. Archive
integrity and per-input verification are recorded in audit.json.

Next: audit remaining P07 Greek/symbol/rare-integral command coverage and implement
bounded independent registry entries before P08 matrices/cases/aligned. Production
UI and Office integration remain gated; keep both this and P07.6 prior checkpoints.
