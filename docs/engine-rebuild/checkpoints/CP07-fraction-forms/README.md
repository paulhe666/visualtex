# CP07 fraction forms — verified two-argument extension

2026-09-13, after 20:58 +08:00. Independent engine continuation from immutable
CP07-style-aliases. Checkout reference-3d207d7, HEAD
00b48e17b736c7e57a0a72e19e271df82906f89d. No Git staging/commit/push,
installation, production engine switch or Office document action.

## Implemented supported slice

binom/dbinom/tbinom/nicefrac now parse and preserve authored source and have real
editable numerator/denominator slots. InsertStructure supports binomial and
skewedFraction. Selection wrapping retains old node identities, copies remain
independent, source patches update actual mathematical children, and all forms
share existing navigation, position maps, deletion and source/visual history.

Zero-bar binomials use MATH stack shifts and gap values plus real font parentheses.
Skewed forms use reduced argument boxes, MATH clearances and a real slash glyph.
No fake CSS, opaque picture, duplicate-owner region, old renderer or anisotropic
outline scaling is used. Colors, hidden phantom paint, nested expressions, signed
spacing and natural-size/layout-cache invariants remain covered. Finite policy
and unsupported general TeX forms are explicit in engine/specs/P07-fraction-forms.md.

## Verification actually executed

Base receipt: engine/target/verification/2026-09-13T12-56-01-837Z/report.json
Source SHA256: 85ecbe4b4ac4ef39fafc92c5b8efe8c4f7cf64d8eacae3344c17c13f2039b6da
sourceUnchanged: true; 141 input files. Every input was rehashed after the separate
supplemental tests and still matched, with zero changed paths.

The base runner's twelve engine/build/browser steps PASS:
- 332 native Rust tests, including 14 new fraction integration tests and one
  bounded fraction-registry test. New generated source cases: 480, seed B10F2026.
- Actual Native/WASM lifecycle 2151, source-session 222, source-edit 440 and
  editor/history 473 requests, including 100 history states.
- Existing layout 3945 requests/2689 scenes and symbols/aliases 2570 requests/
  1891 scenes pass, with unchanged layout transcript identities from the prior
  implementation for the preexisting input set.
- Old sandbox/root/fence/accent/operator/style/spacing/color/symbol browser,
  input-fault-injection and SVG zoom/DPR/scroll tests all PASS.

The two NEW test commands were run separately. Attempts to add them to the old
build/verification runner were blocked by the tool check, so they are NOT claimed
as steps inside the base receipt:
- node tools/fraction-form-parity.mjs: PASS, real native and WASM, 349 requests,
  244 scenes, maximum numeric difference 0, exact SVG. Covers all four styles,
  empty/invalid input, real slot construction, same-owner vertical navigation,
  duplicate/undo/redo, local source patching and malformed-action atomicity.
- node tools/fraction-form-browser.mjs: PASS, 33 trusted browser events through
  actual controls, keys and pointer positioning; blank construction of both forms,
  separate argument editing, up/down movement, selection wrapping, undo/redo,
  SVG natural dimensions and invalid-source recovery. Inspector callbacks are
  read-only, not editor actions. Screenshot was opened and visually inspected.
- Separate independent layout review: 519 requests/360 scenes PASS.
- Separate Office contract snapshot: 11 sources/3 interfaces/10 routes/3 negative
  mutation tests PASS, writes 0.

Supplemental reports:
engine/target/fraction-form-parity/report.json
SHA256 51557a1d6e4c95e0e40111984cec58650d580ea6ac78ce85546a39669dd52ac4
engine/target/fraction-form-browser/report.json
SHA256 a9c880374402635ee753e74d85dbb209ac2a62edaa53715f589dd7d84184d66a
Screenshot: engine/target/fraction-form-browser/fraction-forms.png.

Initial native test compilation found an unwrapped Result in the new test harness;
that was corrected. The first run then found a wrong test count for
phantom{binom}+x: both visible '+' AND 'x' must remain. It was replaced by an exact
two-visible-glyph assertion plus exclusion of every hidden subtree ID from glyphs
and rules. No production/math failure was reclassified as a passing test.

Native CLI: 987a41d65ecf013aac2b16d1c87697155488bea251fcbcf1bd2b3892d270f912
Node WASM: 2e8ce9a2f916f6d075d7675347ba1451c9a34590c439fb32ba6fac5614f7aab6
Web WASM: 58a4cf2ba64e5f992f465c36ba96b8d93ebec355e915e68a31e6dac86c29503a
Fraction transcript: 509f3717674e242a9abcc4c6dd330428a9d11d68ae560062f5788f34cd837866
Browser: Edge 152.0.4191.62; host sandbox-6ba0a633aa03a178b9029573.

## Protection and remaining work

Strict overall receipt remains FAILED at original P00 protection. The reported
paths are preexisting external WordEquationNumbering, ThisAddIn,
WordFormulaFormatConversion, WordFormulaService.CopyPaste/FormatConversion/main,
target-word.json, ui-actions.ndjson, ui-uia.json and source/path groups. No guard or
baseline was weakened, no external changes restored/deleted, and no writer is
identified by assumption. A passing component suite is not permission to integrate
with or replace production Office paths.

No new physical Windows IME, WebView2, Word/PowerPoint/Ribbon, MathML/OMML,
installer or final performance acceptance is asserted. General mathop, remaining
ornament/source compatibility and ESINT-specific artwork are still tracked gaps.
P08 Shift+Enter is still specified as single-frame contained multiline with one
owner and one undo transaction; its implementation is NOT established here.
All engine changes remain uncommitted. Archive identity and verified-input
matching are recorded in audit.json; font/build/output files are excluded.
