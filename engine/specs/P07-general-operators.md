# P07 editable general operators

## Language and ownership

Supported new spelling: mathop with a mathematical argument, optional authored
lower/upper scripts and limits/nolimits/displaylimits directives. GeneralOperator
stores command, exact before-body trivia, one real body Row and the existing
ordered OperatorPart tail. Fixed LargeOperator nodes retain their old shape and
registry behavior. Source command spelling/order/comments are retained; unknown
or unfinished source is never silently repaired or dropped. This does not
implement arbitrary TeX primitives, DeclareMathOperator, operatorname*, or macros.

The body is mathematical content, not an opaque string or the following operand.
It can contain symbols, styled text, fractions, roots, other operators and current
supported structures. Empty bodies and explicitly unfilled bounds are editable
but not exportable. Forged commands/tails are rejected by tree validation.

## Product behavior

InsertGeneralOperator can wrap a current selection while retaining its node IDs,
or create an empty editable body. Optional lower/upper rows get fresh IDs. With
bounds present, initial focus and Right navigation run lower -> upper -> body ->
following expression, skipping absent bounds. A body-only wrapper focuses its body.
This extends the existing lower/upper ordering without treating editable body text
as an externally owned operand. Left/right selection and delete follow the common
structural rules; no source-offset or string-length cursor inference is added.

At a caret immediately following a general operator, the ordinary sub/superscript
commands add or re-enter that operator's bounds. A whole explicitly selected
operator raised to a power is instead grouped, preserving any existing upper
limit. Shared limits controls work from its body, bounds or adjacent caret.
Changing mode is one history transaction; repeated same-mode is a true no-op.
Undo/redo restore source, selection and identities, without rewinding revision or
allocator IDs. All old positions are remapped through PositionMap.

Vertical navigation uses the same operator's direct child baselines. Its editable
body is a real middle stop between lower and upper limits. Fixed glyph operators
still navigate directly between their two bounds. Neither path lands in the
outer root row or unrelated neighboring expression when a same-owner argument is
available. The accessible path identifies generalOperator/operatorBody and bounds.

## Layout

The authored body is laid out once under its inherited mathematical style, font
variant and baseline, then treated as an operator-class atom. It is not replaced
with an arbitrary registry glyph or uniformly enlarged as if it were a sum sign.
Any actual nested sum/integral retains its own explicit style and font variant.
Automatic general-operator bounds stack in display style and use side scripts in
other styles. Explicit limits/nolimits/displaylimits override that policy.

The fixed and general operators now share the existing limit-placement routine,
MATH upper/lower gaps, rise/drop and side-script calculations. Existing fixed
operator geometry must remain unchanged. Body, limits, paint, display list,
source ownership and SVG all belong to one session and scene. Cache identity
includes body structure and effective placement mode; source trivia does not
invalidate geometry. Colors and phantom suppression remain inherited/scoped.

Public behavioral references:
https://learn.microsoft.com/en-us/office/math/latex
https://www.w3.org/TR/MathML2/chapter3.html
https://learn.microsoft.com/en-us/typography/opentype/spec/math
The references guide operator class/placement and font metrics; no renderer or
TeX execution is embedded. Exact legacy/Office output acceptance is a later gate.

## Required verification

general_operators.rs checks authored bodies/tails and canonical roundtrips,
unfinished/unknown/duplicate-bound source, blank and selected-body construction,
ID-disjoint copy/edit/delete, limits/no-op/undo, later added/re-entered bounds,
whole-object powers, actual local body/bound source edits, four-style placement,
unchanged body typography, vertical navigation, nested colored/hidden output,
forged/wire rejection and 600 source disturbances, seed0x0F3A2026.

The existing layout-parity runner gains actual native/WASM general-body sources
and interactive transactions. The existing operator-browser runner preserves its
original cases and adds real controls/keys/pointer for body/bounds, wrapping,
mode changes, history and SVG; its general-operators.png screenshot must be read.
No synthetic editor callback may replace input. The old full suite and separate
fraction-form/independent layout checks must still run against the same source.

This slice is not P08: Shift+Enter still requires contained internal rows within
one formula frame, never merging adjacent frames. Production Office/default
MathLive integration, physical OS IME rerun, installers, MathML/OMML and final
performance remain unaccepted. Existing external P00 protection alerts stay red;
no baseline or safety gate is weakened by a component pass.
