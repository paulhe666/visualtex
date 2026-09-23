# P07 compatibility slice — Unicode mathematical style spellings

This slice extends existing Style nodes; it does not introduce another editor,
a macro renderer or an Office integration. The source remains authoritative.

## Supported additions

Seventeen spellings from the product compatibility command registry are recognized:
symup, symit, symbf, symbfup, symbfit, symbb, symcal, symscr, symbfcal, symbfscr,
symfrak, symbffrak, symsfup, symsfit, symbfsfup, symbfsfit and symtt.

Each is mapped to its explicit MathVariant before font lookup. Authored spelling,
comments and grouping survive source round-trips. Canonical commands for all 15
MathVariant values now map back to the same variant; previously unused mappings
for bold script, bold Fraktur and sans combinations could lose style information.
No old MathLive/MathJax macro runs. Existing product symbf behavior is upright bold.

The ordinary Greek phi/kappa and their variant forms remain distinct. The encoded
Unicode Mathematical Alphanumeric ranges now include the capital theta symbol,
nabla, partial differential, epsilon/theta/kappa/phi/rho/pi variants across the
five encoded bold/italic/sans-bold Greek alphabets. Values without an encoded
requested variant retain the existing defined behavior; this is not a claim that
every character has every font style or that all bm semantics are implemented.

Primary font availability is not assumed: Latin Modern Math lacks some lowercase
script glyphs. The exact-character explicit fallback introduced by P07.8 is used
in tests. The initial lowercase-script missing-glyph test failure is not waived;
the same full alphabet sample is rerun with the declared pinned fallback. No
synthetic italic/slant, outline boldening or post-layout font substitution occurs.

## Verification scope

The style-alias regression iterates the full alias inventory, all four math
styles, source comments/canonical round-trip, glyph IDs, scope boundaries, empty
slot typing, independent copies, delete and history. A fixed seed 0x57a12026 runs
340 local source edits inside nested formula/style structures and checks against
a fresh parse. Unknown spellings and legacy declaration syntax are kept intact
and explicitly unsupported rather than guessed as braced wrappers.

Native/WASM symbol parity includes these alias sources, repeated layout and font
size changes. The actual isolated browser symbol runner additionally selects
math with the keyboard, applies sans-bold-italic and bold-script controls, types
into a blank style body, and verifies undo/redo and SVG through the existing host.
Per-run counts, source inventory and hashes must come from its verification receipt.

## Limits and following work

This does not implement arbitrary unicode-math options, variable font controls,
legacy declarations, all mathcal/mathscr product distinctions, general macros,
ESINT-specific integral artwork, matrices or multiline environments. The separate
P08 contained-multiline specification remains mandatory: Shift+Enter acts inside
one formula frame and never joins neighboring frames. P08 tests remain NOT_RUN.
Production UI, Office interfaces and installed packages are unchanged by this slice.
