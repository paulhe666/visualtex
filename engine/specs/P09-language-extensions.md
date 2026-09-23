# P09 language-extension closure

Status: implemented and accepted in the isolated engine/sandbox scope. This specification does not enable the main application or Office.

## Scope

This slice closes product-authored command templates that were not covered by the earlier P07/P08 grammar without adding a general TeX executor.

Supported structural additions:

- Matrix-family environments `matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`. Each remains one Grid with one row/cell/column identity model. Delimiters are font-driven presentation owned by that Grid, not CSS or a second editor.
- `\middle<delimiter>` inside the nearest `\left...\right` Fence. It is a real `MiddleDelimiter` node and uses the same vertical stretch target as the owning Fence. `\middle` outside a Fence is unsupported and losslessly preserved.
- One- or two-column `cases`. A one-column authored case list does not invent condition cells; mixed ragged one/two-column rows remain unsupported.
- Bounded `array` with explicit `l/c/r` column specifications and `|` vertical rules. Array alignment and rules are geometry, not DOM tables or source rewrites. Unsupported modifiers such as `@{...}` remain lossless/non-exportable.
- Rare integral aliases `\dashint`, `\ddashint`, `\varointctrclockwise` mapped to their exact Unicode mathematical glyph semantics (`⨍`, `⨎`, `∳`) through the existing rare-operator path. Missing font glyphs remain explicit errors; no lookalike substitution is allowed.

## Invariants

1. Every accepted environment round-trips authored LaTeX losslessly until the user edits the affected structure.
2. New structures participate in NodeId/column identity, source patching, copy/duplicate, history and scene validation.
3. Matrix-family wrappers and array vertical rules are part of core layout/SVG output and scale only through the natural font-size transform.
4. Array column insertion/deletion is refused because it would require rewriting the authored column specification; row insertion is allowed and retains that specification.
5. Unknown environments, unsupported array modifiers, ragged grids and unfinished delimiters remain recoverable source, never silently normalized to a different formula.
6. Command-registry coverage is evaluated with structurally safe placeholder materialization. Contextual templates beginning with `^`/`_` are tested with an explicit base instead of being misclassified as standalone formulas.

## Capability audit

`tools/p09-command-inventory.mjs` schema v2 records both standalone and contextual capability. The target condition is that all existing static Windows command-registry templates are covered by one of those two modes, while trigger collisions remain explicit data rather than being collapsed.

## Acceptance layers

- Rust structure/layout/source/history tests for matrix variants, one-column cases, arrays and middle delimiters.
- Native/WASM exact response and SVG parity for matrix variants, middle delimiters and the existing rare-integral catalog.
- Owned Edge pointer/keyboard/zoom/DPR evidence for matrix variants and middle delimiters.
- P08 regression remains a lower-phase gate.

No claim in this file covers the main application, Office, MathML export or installed-package acceptance.