# CP09 builtin compatibility macros 1

This is a non-Git P09 engine checkpoint. It does not change the production app,
Office/VSTO paths, P00 protection baseline, or MathLive/MathJax production
selection. No commit or push was made for this engine work.

## Stable source

- branch: `reference-3d207d7`
- base HEAD: `00b48e17b736c7e57a0a72e19e271df82906f89d`
- stable verification: `engine/target/p09-verification/2026-09-14T01-08-09-777Z/report.json`
- stable source SHA256: `4d25a397ac10eda35e8f66c900fe6abfbabcfd4a763ae8c92aa1435dc097e890`
- verification status: PASS, source unchanged

## Capability in this slice

P09 self-owned command input from `CP09-command-input-1` remains passing. This
checkpoint additionally introduces a fixed, bounded `MacroCall` model for:

`\abs`, `\norm`, `\dd`, `\dv`, `\pdv`, `\comm`, `\acomm`, `\pb`, `\bra`,
`\ket`, `\braket`, `\expval`, `\mel`, `\ketbra`, `\vb`, `\va`, `\vu`.

The original command and argument trivia are preserved. Every argument is an
identified editable Row and participates in source reconciliation, projection,
TreeIndex, PositionMap, duplication, source patching and history. Layout is
performed by the Rust engine from the injected OpenType MATH fonts. No macro in
this slice is rendered or expanded through MathLive, MathJax, CSS margins or a
TeX executable.

Unknown or executable TeX such as `\input`, `\write18`, `\def` and arbitrary
user commands remains unsupported and is never executed.

The historical `physics-macro-p09` corpus sample was promoted from its explicit
P09-pending negative classification to a positive `\pdv{f}{x}` regression. The
surrounding unknown/executable-source negative fixtures remain unchanged and
passing.

## Evidence

`crates/core/tests/builtin_macros_review.rs`: 7 dedicated Rust tests passed for
all 17 macros, comments/trivia, parameter editing, source patching, copy with
fresh identities, undo/redo, real-font SVG geometry and negative drafts.

`tools/macro-parity-review.mjs`: PASS with actual Native and WASM execution:
381 requests, 267 scenes, 29 fresh source oracles, maximum numeric difference 0,
and exact SVG parity.

`tools/macro-browser-review.mjs`: PASS in Edge 152.0.4191.62. All 17 macro
sources reopened/exported with real argument carets. `\pdv` was edited through
trusted browser pointer/keyboard input without changing its sibling argument
identity. 45 representative zoom/DPR pointer cases passed. An independent B
formula snapshot remained unchanged.

The current 411-entry Windows command-registry audit increased the number of
materialized templates accepted by the new parser from 358 to 375, exactly the
17 macro commands added by this slice. The remaining entries are not claimed.

The stable P09 verifier first reran the entire P08 acceptance gate, then command
inventory/state/browser and macro Native/WASM/browser gates. All six steps
passed with the source hash unchanged.

## Explicitly not complete

P09 as a whole is not complete. Matrix environment aliases, additional
delimiters/templates, arbitrary bounded user macros, versioned rich clipboard,
custom-symbol assets, OCR revision targeting, and configuration compatibility
remain later P09 work. P10 production integration, P11 Office compatibility and
P15 production dependency removal are not part of this checkpoint.

P00 remains a separate red protection gate because unrelated parallel
Office/VSTO/evidence files in the checkout are dirty. No attempt was made to
clean, overwrite or normalize those external changes.
