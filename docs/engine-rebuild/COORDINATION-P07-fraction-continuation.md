# P07 fraction-family continuation

Continuation of the user's engine rebuild, after immutable CP07-style-aliases.
Actual checkout reference-3d207d7, HEAD 00b48e17b736c7e57a0a72e19e271df82906f89d.
No Git staging/commit/push, production entrypoint, installer or Office writes are
permitted in this continuation. Existing tracked Office/evidence changes are
external to this lane and remain intact. The original P00 guard remains strict.

Current owner work: bounded fraction-family specification and complete
parse/edit/layout/source/history/browser support for binom/dbinom/tbinom and
nicefrac, retaining the existing Fraction numerator/denominator ownership rather
than inventing a second document state. Read-only OpenType MATH specification and
current product command data guide this extension. Normal frac/dfrac/tfrac output
must remain unchanged. Independently named review tests may be added; do not
concurrently overwrite core/parser/layout/transactions or the existing host.

This note is not a phase acceptance. New tests must include wrong/missing input,
zero-bar stack constraints, actual-font parenthesis/slash geometry, source
roundtrips, node IDs and position maps, copy/edit/delete/history, real sandbox
keyboard/pointer controls, unchanged K0 regressions and actual Native/WASM parity.
A new immutable source checkpoint must be created only after recorded verification.

P08 user requirement stays in engine/specs/P08-contained-multiline.md:
Shift+Enter converts/splits rows inside the currently owned formula frame only;
no adjacent frame merge, no global selection migration, composition confirmation
first, one transaction for conversion, and independent two-frame regression.
No implementation or acceptance of P08 is asserted by this note.
