# P07: binomial and bevelled two-argument forms

## Finite supported surface

Existing frac/dfrac/tfrac are unchanged. Added source commands are binom, dbinom,
tbinom and nicefrac. All share one real Fraction node with two editable row
identities, numerator and denominator; the authored command determines the
presentation through fractions::spec. Commands, braces, comments and trivia are
retained by the existing lossless source ownership. Canonical export retains the
specific authored command. This does not add genfrac, infix over/atop/choose,
arbitrary TeX execution or a fallback to a legacy engine.

## Editing and history

InsertStructure::Binomial and ::SkewedFraction are genuine tree transactions.
A selected mathematical range becomes the numerator, retaining its node IDs, and
the caret starts in the new denominator. Empty insertion starts in the numerator.
Existing fraction navigation, selection, copy/delete and undo/redo rules apply.
Both slots remain editable in empty drafts; they do not become finished exports
until required content is present. Source patches must change real argument
content, not only the source string, and cannot exchange IDs of repeated forms.
Unknown constructed command values are rejected by tree validation before source
projection. Source/visual history remains owned by EngineSession.

## Mathematical layout

A binomial's two arguments use the fraction child style; dbinom forces display
and tbinom forces text, independent of ambient style. No fraction bar is painted.
Vertical shifts and minimum gap use OpenType MATH StackTopShiftUp,
StackTopDisplayStyleShiftUp, StackBottomShiftDown,
StackBottomDisplayStyleShiftDown, StackGapMin and StackDisplayStyleGapMin. Any gap
deficit is split equally between top and bottom shifts. Distances are checked for
nonfinite/negative values at use. The font registration's existing validation was
not weakened.

Parentheses are centered on the math axis with a shared target, using real MATH
variants or assemblies. Ink overhang and signed advance are accounted for when
placing them. Delimiter glyphs belong to the Fraction owner; no duplicate same-ID
layout region or noneditable pseudo-child is introduced.

nicefrac uses reduced script-style argument boxes on opposite sides of a genuine
slash glyph. Vertical clearance uses SkewedFractionVerticalGap around the math
axis. Horizontal placement uses the argument ink/logical union, actual slash
advance and SkewedFractionHorizontalGap. This is the bounded engine's documented
bevelled-fraction policy, not exact execution of every units/nicefrac package
option. The slash uses a designed font variant/assembly; unavailable required
geometry fails explicitly rather than stretching an outline anisotropically.

All forms participate in math alphabets, color, phantom paint suppression, nested
structures, layout caching and same-scene SVG. Natural pixel dimensions stay
independent of host zoom. Repeated pixel-size changes reuse em geometry.

Primary references read for the implementation:
- Microsoft OpenType MATH, MathConstants and MathVariants:
  https://learn.microsoft.com/en-us/typography/opentype/spec/math
- W3C MathML Core, 3.3.2.2 zero-line-thickness fraction (stack constraints):
  https://www.w3.org/TR/mathml-core/#fraction-with-zero-line-thickness
These provide font/layout rules, not an imported rendering engine.

## Verification required for this slice

- fraction_forms.rs: authored/invalid/unfilled input, source identity, two real
  slots, wrap/copy/delete/history, stack/axis/gap constraints, forced styles,
  actual slash and parenthesis glyphs, colors/phantoms/signed glue, cache/undo,
  malicious constructed commands and 480 seeded source disturbances (0xB10F2026).
- fraction-form-parity.mjs: execute the actual native binary and actual WASM
  under identical fixed math/CJK bytes; compare full structures/geometry/SVG,
  error atomicity and editing history across four mathematical styles.
- fraction-form-browser.mjs: compile current host modules, click actual visible
  controls, type through browser keys, use pointer coordinates to edit each
  argument, navigate up/down, wrap selection, undo/redo, export current scene,
  retain and recover an incomplete source draft; inspect the screenshot.
- Rerun existing native, Native/WASM, host/input/SVG and production guards. A red
  external Office guard is still red and does not grant production integration.

No production main UI, Office UI, OS IME rerun, MathML/OMML target conversion,
installer, cross-platform or event-to-paint performance acceptance is established
by these tests. P08's contained Shift+Enter decision remains separately specified
and is not implemented by fraction-family support.
