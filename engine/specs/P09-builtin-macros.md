# P09 bounded compatibility macros

Status: implemented and independently tested in the engine/sandbox. This is not
P09 completion, a production-app switch, an Office acceptance result, or a
claim that arbitrary TeX macro definitions are supported.

## Contract

VisualTeX does not execute these commands by passing their source through
MathLive, MathJax, TeX, a shell, or an arbitrary macro evaluator. The supported
compatibility set is a finite registry in `crates/core/src/macros.rs`. Each
accepted source command is parsed into one `MacroCall` node with:

- the original command spelling;
- one fixed `BuiltinMacroKind`;
- an exact fixed arity;
- lossless trivia before every authored argument; and
- one independently identified editable Row for every argument.

The first bounded set is:

`\abs`, `\norm`, `\dd`, `\dv`, `\pdv`, `\comm`, `\acomm`, `\pb`, `\bra`,
`\ket`, `\braket`, `\expval`, `\mel`, `\ketbra`, `\vb`, `\va`, and `\vu`.

Their mathematical meanings follow the existing VisualTeX package-compatibility
layer, but that legacy expansion is reference data only. The new engine owns the
structure and layout itself.

## Source and editing

Lossless serialization reproduces the authored macro call, including whitespace
and comments between command and arguments. Canonical serialization normalizes
only the already-owned mathematical syntax; it does not replace a MacroCall with
an expanded macro body.

Arguments participate in TreeIndex, source reconciliation, projection,
PositionMap, duplication, source patching, history, scene description, pointer
selection and accessibility paths. Editing `\pdv{f}{x}` inside the first
argument must not change the macro identity or the denominator Row identity.
Duplicating a macro allocates a disjoint subtree. Undo/redo restores structural
identity, not merely equivalent LaTeX text.

Missing arguments remain drafts and cannot be exported. Unknown commands remain
opaque/unsupported and are never interpreted as a supported MacroCall.

## Layout

Layout uses the same injected OpenType MATH resources as native structures:

- absolute value, norm and Dirac outer delimiters use real stretchable glyphs;
- derivative and partial-derivative presentation uses the engine fraction MATH
  parameters plus real `d`/`∂` glyphs;
- commutator, anticommutator, Poisson bracket and Dirac multi-argument forms use
  real delimiter and separator glyphs;
- vector aliases use the engine's mathematical variant and accent mechanisms.

No CSS margins, HTML text, DOM table measurements, raster substitution or hidden
MathJax rendering is accepted as macro geometry.

## Security boundary

This capability is deliberately not a TeX macro interpreter. Commands such as
`\input`, `\write18`, `\def` and unknown user commands remain unsupported and
cannot execute files, processes, scripts or arbitrary expansion logic. Resource
and parser depth limits still apply before committing a session state.

User-defined macros and custom-symbol assets remain later P09 work. Their future
implementation must have an explicit versioned/bounded definition schema; this
slice must not be generalized into unrestricted TeX expansion.

## Acceptance

`crates/core/tests/builtin_macros_review.rs` covers all 17 commands, authored
trivia, editable arguments, source patching, fresh parse equivalence, copy with
fresh identities, exact history restoration, real-font layout and negative
unknown/missing-argument cases.

`tools/macro-parity-review.mjs` executes actual native Rust and WASM requests and
compares every response, including layout/SVG. The first accepted run covered
381 requests, 267 scenes and 29 fresh source oracles with zero numeric difference
and exact SVG parity.

`tools/macro-browser-review.mjs` uses real Edge pointer/keyboard events. It
reopens and exports all 17 macros, edits a `\pdv` argument, checks undo/redo and
an independent second formula frame, and exercises representative macro
arguments across zoom 0.5/1/2 and DPR 1/1.25/2.

The current Windows command-registry inventory increased from 358 to 375
materialized templates supported by the new parser after this slice. That count
is evidence of finite progress, not permission to claim the remaining registry
entries as supported.

## Not covered by this slice

- arbitrary/user-defined macro definitions or recursive expansion;
- matrix environment aliases such as `bmatrix`/`pmatrix`/`vmatrix`;
- remaining delimiter variants and templates;
- custom-symbol asset definitions;
- production React/CodeMirror integration (P10);
- Office compatibility/output (P11+) or the production MathLive/MathJax removal
  gate (P15).
