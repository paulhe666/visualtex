# P07.4 — large operators and limits

Scope: isolated mathematical core, Native/WASM, current SVG host and independent sandbox. No production UI, Office contract, document object, installer or other platform entry point is switched.

## Mathematical representation and source

`LargeOperator` retains the author's command and an ordered sequence of `OperatorPart::Script` and `OperatorPart::Limits`. Commands supported in this slice: sum, prod, coprod; bigcup, bigcap, bigsqcup, biguplus, bigvee, bigwedge, bigodot, bigoplus, bigotimes; int, iint, iiint, oint, oiint, oiiint; upright lim, limsup, liminf, min, max, inf, sup. Corresponding single-code-point glyph input is recognized as the same mathematical structure. Arbitrary new integral aliases, user-defined mathop/operatorname and macro redefinition are not claimed here.

Tail order, bound order, grouping and comments are preserved. Limits directives can appear before, between or after scripts. Effective placement uses the last directive. Unsupported commands remain opaque, duplicate bounds are diagnosed, and incomplete or deliberately empty bound arguments remain non-exportable drafts. A bare operator with no authored bound is valid. Source and tail resource budgets are explicit. No regular-expression rewrite of the whole formula is used for operator edits.

## Editing and navigation

`InsertOperator` creates requested editable lower/upper rows. A selected expression remains the following operand with original identities; it is not silently used as a bound. The isolated sandbox default navigation is lower -> upper -> following expression even when authored source spells the upper bound first. Existing product input preferences are not overwritten: final application preference integration is a later gate.

Adding a sub/superscript at the caret after an existing operator adds or re-enters that operator's corresponding bound without replacing its ID. In contrast, applying a power to an explicitly selected whole operator creates a grouped outer power; an existing upper bound is never turned into duplicate script syntax. Copy generates fresh independent subtrees; delete/undo restore exact IDs and content through existing transaction/history machinery.

`SetOperatorLimits` operates on a selected operator or the operator enclosing the current bound. Repeating the same effective mode is a no-op. Explicit changes are normal undoable source transactions. Restoring Auto removes placement directives while retaining their comments/trivia and the bound identities. This differs from a mathematical display-style change: changing the ambient display/text style changes only layout, not source, document revision or undo depth.

Vertical navigation prioritizes a sibling bound of the same nearest mathematical operator rather than the baseline of an unrelated outer row. Targets use actual layout baselines and caret geometry, not source lengths. Selection descriptions identify largeOperator/lowerLimit/upperLimit.

## Font and layout

Display glyph selection reads the fixed font's OpenType MATH displayOperatorMinHeight and designed vertical variants. The operand's height does not stretch the operator. If a font has no sufficiently large variant, only its largest designed glyph is used, with no arbitrary anisotropic enlargement or borrowed delimiter assembly. The verified fixed Latin Modern Math resource has suitable display variants for tested sums/integrals/products. General cross-font appearance remains separately subject to acceptance.

Sum-like and named limiting operators stack their bounds by default in display style and use side bounds in text/script styles. Integrals use side bounds by default. Explicit limits always stacks; nolimits always uses side bounds; displaylimits stacks only in display style. Upper/lower minimum gaps and baseline distances and side-script shifts come from MATH metrics. Glyph italic correction is applied to the appropriate bound placement. Logical advance and ink bounds stay separate. All geometry is in em with natural CSS-pixel output at the SVG boundary.

Upright named operators are shaped through the injected math font, not through arbitrary browser text layout. Their following atom spacing uses the Operator class. Bounds and glyphs belong to one layout snapshot; changing effective limits participates in the geometry cache key.

Primary technical references: Microsoft OpenType MATH table (https://learn.microsoft.com/en-us/typography/opentype/spec/math), W3C MathML Core (https://www.w3.org/TR/mathml-core/). Font data and these standards do not themselves prove implementation correctness; tests below provide project evidence.

## Verification coverage

`tests/operators.rs`: 17 tests including the 25-command corpus with directive orders/trivia, all four mathematical styles and four limit modes, fixed-font variants/gaps, bound editing and re-entry, selection wrapping, independent copies, shared undo/redo, unknown input/resource failure atomicity, cache invalidation, and 600 seeded source disturbances. Related scene query tests cover the same upper/lower target in both stacked and side layouts.

`tools/layout-parity.mjs`: same Native/WASM source/edit and geometry requests, identical fixed font bytes, explicit operator actions, vertical queries, limits changes and undo. Tests include mixed operators with accents, roots, fractions and Chinese bound text. Existing independent layout, source, history and component tests remain in place.

`tools/operator-browser.mjs`: actual owned browser controls, key and pointer input starting from blank; lower-upper-body construction; pointer bound selection; up/down navigation; explicit limits/no-limits/default controls; undo/redo; mathematical style switching without source/history mutation; side/display screenshots; same-scene SVG dimensions; invalid upper-bound draft recovery. Previously checked export previews are invalidated after source or mathematical-style changes, but not simple zoom changes. The test uses browser input, not a new physical Windows IME or Office acceptance run.

Reproduced regressions retained: whole-operator power originally emitted duplicate upper-bound syntax; vertical bound movement initially landed in the root row; sandbox style changes originally left stale export preview text. A separate test assertion was corrected after logging demonstrated it incorrectly added a glyph's left ink bearing twice; final assertions compare exact logical side-bound coordinates to glyph origin+advance+italic correction.

Phase/source/binary/archive hashes and final pass counts are recorded in immutable checkpoint receipts. This document does not assert completion of all P07 structures, MathML encoding, existing command-registry coverage, Office compatibility or full MathLive removal.
