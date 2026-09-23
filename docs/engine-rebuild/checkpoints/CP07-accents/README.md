# CP07.3 — accents and annotations, verified source checkpoint

2026-09-13, approximately 15:42 +08:00. Product baseline: `reference-3d207d7`, HEAD `00b48e17b736c7e57a0a72e19e271df82906f89d`. Non-Git source archive only. No engine staging/commit/push, production switch, installer rebuild or user Office operation.

## Scope actually established

The supported P07.3 slice now has 29 accent/decoration commands and 4 annotation commands, with source preservation, real mathematical nodes, structure edits/wrapping, independent copied identities, source patching, delete/history, fixed-font layout and same-scene vector SVG. Font-driven wide glyph assemblies, over/under rules, annotation labels and brace labels are covered. Incomplete and unknown input remain preserved and non-exportable.

Two integration failures were reproduced before correction: selection description failed with `invalidPosition` for new accent/annotation subtrees, and Right arrow only exited the innermost accent. Position traversal now describes the label/base/body; Right exits a completed final accent chain without changing source/revision/history and does not skip siblings or fraction slots. Shift expansion retains normal selection. Browser host types/actions and the existing independent sandbox controls were updated, without adding a second editor owner.

## Executed checks

Primary receipt: `engine/target/verification/2026-09-13T07-40-30-846Z/report.json`.
Verified source SHA256: `693343a0a257433e3fc21857f41c454d48d5e1c6edc05fb42ceb958d28da5034`.
Seven existing runner steps PASS and source inventory unchanged during that run:

- Rebuild: 240 native Rust tests, 0 failed. Accent tests: 14; scene queries: 7. The accented corpus includes 600 seeded source edits and all four mathematical styles.
- Rebuilt Native/WASM: 2151 protocol cases, 222 source-session cases, 440 source-edit requests, 473 editor/history requests including 100 history states.
- Expanded Native/WASM layout: 1058 requests, 693 valid scenes, maximum numeric difference 0 with fixed font bytes, exact SVG parity.
- Existing sandbox, root and fence browser controls all pass. Input fault injection and SVG zoom/DPR/scroll tests pass.
- Production/Office guard: 12 groups, 12 artifacts, 65 original dirty working-file byte checks PASS, with the previously authorized Git transition acknowledged.

Additional commands were executed separately without source edits after the primary run:

- `node tools/layout-parity-review.mjs`: PASS, independent 519 requests / 360 valid scenes, relative 1e-9 numerical tolerance and exact SVG/identities.
- `node ../docs/engine-rebuild/tools/contract-snapshot.mjs --check`: PASS, 11 sources, 3 interfaces, 10 routes, 3 negative mutations; writes=0.
- `node tools/accent-browser.mjs`: PASS, actual browser controls/keys/pointer, 26 trusted browser events. Starts from blank to construct nested accents, exits with Right, wraps a range, edits the upper label by pointer, exercises undo/redo, creates a labelled underbrace, changes zoom, exports an assembled long brace and recovers an invalid source draft. Report: `engine/target/accent-browser/report.json`. Screenshot: `engine/target/accent-browser/accents-and-annotations.png`; the generated screenshot was opened and visually inspected.

The attempted addition of these extra steps to `verify-current.mjs` was blocked by the tool safety check. The runner was left unchanged. The separate commands above are explicitly not misrepresented as part of its seven-step receipt.

Native CLI SHA256: `bfd19de3e4aa22342c102e3a2577b7b98e056a6defe3be8951659a5a026291eb`.
Node WASM SHA256: `4f80b72ea041b69188ede23ea79d73146597196055ee983a73dbaf2883919000`.
Browser WASM SHA256: `3e3fb02925f1e6c59187e76b8a8da8bb7f7005684b483040eeb87965074d98b7`.
Browser host build: `sandbox-809226cd05bb6a164a2fd152`, Edge `152.0.4191.62`.

## Limits and next gate

P07 remains in progress. This checkpoint is the verified supported accent/annotation slice, not all existing command-registry coverage. See `engine/specs/P07-accents.md` for finite wide-accent variants, optional flattened-accent feature and cross-font limitations. Real Windows IME evidence remains in CP06; this run does not claim another physical IME test. MathML/Office output, actual Word/Ribbon, WebView2, installed UI, macOS/Web production integration and final release are not accepted here.

Next isolated slice: P07.4 large operators and limits/nolimits, including authored source, display/inline policy, script navigation, glyph variants and shared history. Keep production MathLive/MathJax paths and all Office contracts unchanged. Preserve this checkpoint; subsequent archives must use another directory. Font binaries and build/runtime outputs are excluded from the source archive.
