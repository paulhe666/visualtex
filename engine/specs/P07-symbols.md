# P07.8 — bounded symbol catalog and explicit mathematical font fallback

This document describes the implemented isolated slice, not completion of all P07
or a production MathLive/MathJax replacement. The preceding color/structure/layout
capabilities and Office compatibility gates remain in effect.

## Added scope

`symbol_catalog.rs` declares 175 additional command spellings (including aliases)
for Greek variants, letterlike/physics symbols, ordinary delimiters, relations,
sets, binary operations and arrows. The declaration generates both compiled
lookup and an iterable inventory. A named command retains its exact spelling on
the mathematical Symbol node; it is not rewritten into a different command.
Unambiguous direct Unicode input uses the corresponding mathematical class.
Bare ambiguous bars are not guessed as left/right delimiters.

`InsertSymbol` is a real structural transaction and accepts only registered atoms.
It does not parse arbitrary command strings or run product templates. In
particular `\\partial` inserts one symbol; a derivative fraction is a separate
explicit template action. Existing source projection, safe control-word spacing,
copy identity, PositionMap, deletion and shared undo/redo are reused.

`rare_operators.rs` adds 22 Unicode-defined integral command spellings. The old
product's `extendedIntegralCompatibility.ts` supplies the spelling inventory;
no old rendering function is called. These are real LargeOperator nodes with
editable bound rows and the established limits/nolimits/displaylimits policy.
Their glyph is looked up at the exact semantic codepoint, never approximated by
another integral, image or anisotropically scaled shape.

## Fonts and failure behavior

The original fixed Latin Modern Math font lacks digamma, Diamond and several
Unicode integral glyphs. Initial all-command rendering tests therefore failed.
The corrected implementation introduces explicit ordered `mathFallbacks` resource
IDs in internal LayoutOptions. It defaults to empty and is not an Office DTO or
user font setting. All IDs must be registered, unique, bounded (at most four), and
provide OpenType MATH data. Primary-font glyphs win; only an actually absent glyph
can select the next font. Invalid fonts and missing glyphs remain errors.

Glyph metrics, outlines, italic correction, accent attachment and designed large
operator variant are read from the selected font before layout. The scene records
actual font IDs for every glyph resource. Ordered fallback identity is part of
layout cache context; removing/changing a fallback cannot reuse stale geometry.
The primary mathematical layout profile and em scale remain unchanged.

Current fallback coverage is mathematical scalar symbols and glyph operators.
Do not infer general text/bidi shaping, all accent/brace assembly fallback, or
cross-font visual equivalence from this slice. Font family preference and final
product resource packaging are later gates.

Pinned local test resources:

- Existing Latin Modern Math: SHA256
  `6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f`.
- Optional test fallback STIX Two Math v2.13, face 0: SHA256
  `f2076b9f1676438439dd41e23676f5ab99056e83d6b8f8c27841591ef2ccfa72`,
  838652 bytes. Its official SIL OFL text SHA256 is
  `bec17cee6412788db45fd2644b301afbc99cc5d372ddd3345dc4a7bfdefb9f04`.

`tools/prepare-symbol-test-font.mjs` checks a cached fixture or obtains that fixed
official resource into `engine/target/test-resources/` only. It refuses a changed
cached hash rather than overwriting it. No font is installed or added to production
resources. Font files are excluded from all source archives and user deliverables.
The sandbox requests the fallback explicitly only in its symbol test configuration.

## Required evidence for this slice

Native catalog tests iterate every added symbol's source, canonical round-trip,
copy/delete/undo identities and font rendering in all four mathematical styles.
Integral tests cover all four limit modes and styles, actual lower/upper editing,
copy disjointness and mode undo. Invalid symbol actions and invalid fallback lists
are atomic. Strict missing-font rejection and cache retirement remain tested.

`tools/symbol-parity.mjs` executes the same cases in actual native Rust and WASM,
including matching resource identities, SVG, geometry, cache behavior and editing
history. `tools/symbol-browser.mjs` uses actual sandbox controls and browser
keyboard/pointer input, starting from an empty formula. It does not assert native
OS IME or Office UI acceptance. Raw counts and build hashes belong to the current
verification receipt, not to a permanently assumed number in this specification.

## Explicit remaining scope

ESINT-specific shapes/aliases (`sqint`, `sqiint`, `dotsint`, `idotsint`,
`ointclockwise` and related spellings) are not declared implemented here. Some
legacy mappings were intended for semantic conversion only and are not faithful
rendering definitions. Unknown commands remain intact and explicitly unsupported.
Full product template/macro/custom-symbol coverage, all existing style semantics,
MathML/Office compatibility and final performance remain separate requirements.
Do not present 175 new atoms as coverage of the entire product command registry.

P08 follows the separate user decision in `P08-contained-multiline.md`: default
Shift+Enter creates/inserts environment rows inside exactly one formula frame and
must never merge neighboring frames. That file currently defines requirements
and NOT_RUN tests; this symbol slice does not implement the shortcut or matrices.
