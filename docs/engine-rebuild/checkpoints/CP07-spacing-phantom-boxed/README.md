# CP07.6 — verified spacing / phantom / boxed slice

2026-09-13. Local baseline `reference-3d207d7` at
`00b48e17b736c7e57a0a72e19e271df82906f89d` remains unchanged. Non-Git
source checkpoint; no engine staging/commit/push, installation, Office edit or
production MathLive/MathJax switch was performed.

## Status

All ten engine/build/browser verification steps PASS. The strict runner remains
FAILED at the final production protection step because existing concurrent
VSTO/source/evidence changes differ from the original P00 baseline. The guard and
baseline were not relaxed or replaced. This checkpoint does not authorize Office
integration and is not a claim that all P07 or the whole engine replacement is done.

## Implemented and exercised

Fourteen explicit spacing command spellings, signed advance, mu-versus-em spacing,
non-atom glue classification, full/horizontal/vertical phantoms, real hidden-body
editing positions and paint suppression, and editable boxed display formulas with
vector borders. Parser/projection/IDs/source patching/typing/wrapping/copy/delete/
undo/redo and the existing authoritative sandbox host are connected. Hidden glyphs,
outline resources and hidden rules are excluded from SVG, not made transparent.
See `engine/specs/P07-spacing-phantom-boxed.md` for exact scope and border defaults.

Two actual defects were reproduced and corrected rather than waived:

1. The P07.5 style-source patch test originally inspected identities but not changed
   child content. Adding a direct assertion reproduced `Source changed but styled
   math still contains old x`. `source_edit::rebase_and_replace` now traverses Style,
   Phantom and Boxed. The stronger test and nested/generated source checks remain.
2. Allowing signed mathematical advance initially let a preexisting malicious scene
   mutation test pass invalid data. That old rejection assertion is unchanged.
   Scene validation now ties signed advance to root logical width and directed
   root caret displacement, accepting real negative glue but rejecting corruption.

The first new browser test used unsupported Ctrl+Home/End chords and consequently
inserted an empty phantom before unselected text. It was corrected to actual Ctrl+A
and now checks the mathematical selection before clicking the wrapping button.
No source-output assertion was removed and no scripted editor callback substituted
for keyboard or pointer input.

## Executed evidence

Stable-source receipt:
`engine/target/verification/2026-09-13T09-20-55-996Z/report.json`

Verified source SHA256:
`b2a3063165bbba49a56c376aa3ca941d5e57b3144e5f5073c1a12c37669d605f`

The receipt confirms `sourceUnchanged: true` throughout the run.

- 284 Rust tests PASS, including 14 new spacing/phantom/box integration tests,
  the strengthened style test and retained negative-scene corruption assertions.
  The new randomized source corpus includes 480 disturbances, seed 0x706e2026.
- Native/WASM protocol 2151, source-session 222, source-edit 440 and editor/history
  473 cases PASS, including 100 history states.
- Expanded real Native/WASM layout: 3467 requests, 2361 validated scenes, numeric
  difference 0 at fixed identical font bytes, exact SVG parity.
- Existing sandbox, roots, fences, accents, operators, styles, input-fault-injection
  and SVG zoom/DPR/scroll regression steps PASS in the same receipt.
- New spacing browser test PASS: actual controls and keyboard/pointer input,
  source/selection inspection, hidden-body typing, selection wrapping, negative
  glue, visible boxed sibling, undo/redo, zoom and invalid-source recovery.
- Independent `node tools/layout-parity-review.mjs`: 519 requests / 360 scenes PASS.
- Separate `node ../docs/engine-rebuild/tools/contract-snapshot.mjs --check`: PASS,
  11 sources, 3 interfaces, 10 routes, 3 negative mutations, no writes.

Browser report: `engine/target/spacing-browser/report.json`.
Screenshot: `engine/target/spacing-browser/spacing-phantom-boxed.png`, read through
Devspace and visually inspected: hidden contents absent, evaluation-height line
present and visible box borders not clipped. Browser was Edge 152.0.4191.62,
host `sandbox-b7ddfc1788a5e06193442013` in the individual verified run.

Native CLI SHA256:
`39bb8f6c4ae89122052d3fab26d3c178a9bb4403768604f4debb60a0c63019be`
Node-target WASM SHA256:
`3ea0b70e5023394c0dac6f34dfa13d0a59349c6ba40ded8e2b69142fd87ff400`
Web-target WASM SHA256:
`e5278b8130e7f1bf1d01342ff2fd51462d4cacf3c6c3705f901ffab5ea172458`

## Protection / limits

Final guard at 09:22:11Z reported preexisting differences in ThisAddIn,
WordFormulaFormatConversion, WordFormulaService.FormatConversion,
WordFormulaService, target-word.json, windowsSource/wordVsto/vstoShared groups and
historical untracked paths. No tracked Office or production file was written by
this continuation. The failure is retained, not attributed to an absent tool reply.
The current task does not identify the writer of external files.

No new physical Windows IME, actual Word/Ribbon, PowerPoint, installed WebView2,
macOS/WPS, MathML/Office conversion or package acceptance is asserted. No event-to-
paint/Office performance claim is derived from these correctness tests. Final
production UI is still unchanged; the sandbox is a test interface only.

Next isolated P07 work: colors/paint and remaining symbols/compatibility coverage,
then P08 matrices/cases/aligned. Keep all existing non-Git checkpoints and external
working changes intact. Archive integrity is recorded separately in audit.json.
