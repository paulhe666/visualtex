# P11 — Office output compatibility without Session takeover

Status: implemented and verified offline on Windows. This phase does not enable the core editor in any Office Session and does not modify Word/PowerPoint document algorithms.

## Boundary

P11 adds one pure-core export operation and one host compatibility adapter:

- `exportPresentationMathMl` encodes the retained mathematical tree directly to Presentation MathML. It never calls MathJax, the DOM, Tauri, Office, or source-string reparsing.
- `coreOfficeCompatibility.ts` converts one locked core owner into the existing `OfficeExportResult` shape without changing field names or units.
- The existing `MathTypeMtefCodec`, `OfficeOlePreview`, Session DTOs, VSTO services, numbering, COM/OLE ABI and document replacement algorithms remain unchanged.

P12 is responsible for deciding when an actual interactive Office Session uses this adapter.

## Presentation MathML contract

The encoder supports the structures already accepted by the core: rows, atoms, text, fractions/binomials/bevelled fractions, radicals, fences and middle delimiters, scripts, accents/annotations, large/general operators and limit policy, styles, bounded macros, explicit spacing, phantom/boxed/color structures, matrices/cases/arrays/aligned environments and immutable custom-symbol semantic fallback.

Export is fail-closed when the source is incomplete, unsupported, contains an unfilled slot, contains opaque source, or a custom symbol has no valid `ommlFallback`.

Aligned tables carry a private `data-visualtex-grid-id` only inside the core reply so the host can bind layout `gridColumns` to the matching table. The adapter removes that private identity before returning Office MathML. It adds only the established `data-visualtex-mtef-ruler-stops` compatibility attribute.

A single top-level aligned grid is emitted directly as the `<math>` child rather than through an extra `<mrow>`. This is required by the existing MathType codec's root PILE path and was verified by real MTEF round-trip.

## Geometry and units

Core layout remains the authority for `width`, `height`, `baseline`, SVG and aligned column anchors. Office padding is requested through core layout rather than by post-scaling geometry:

- inline VisualTeX: 1 CSS px each side;
- block VisualTeX: 10 CSS px each side;
- MathType: zero horizontal padding and vertical padding equal to `0.5 × fontSizePt` CSS px, matching the current product contract;
- `fontSizePx = fontSizePt × 96 / 72`;
- MathType ruler stop: aligned right-cell anchor in CSS px × 24, yielding MTEF 1/32 pt units.

Ruler stops must be positive, strictly increasing, and no greater than `65535`. A malformed or incomplete aligned geometry mapping fails instead of silently emitting default MathType tabs.

The `wordOmml` path is MathML-only and must make zero layout/SVG calls. This preserves the existing conversion fast path.

## SVG/PNG compatibility

The ordinary core SVG output is path/rect vector geometry and is accepted unchanged by the existing `OfficeOlePreview` SVG→EMF consumer. P11 does not relax that consumer.

Custom artwork using mask/clip semantics that the current Office vector consumer cannot represent is rejected by the adapter for full vector export. Its Presentation MathML may still use the frozen custom-symbol `ommlFallback`. No raster image is disguised as a vector EMF.

PNG remains a host post-process injected into the adapter. P11 does not put browser canvas behavior into the pure core.

## Verification gates

P11 requires all of the following:

1. core Presentation MathML structural tests;
2. exact Native/WASM Presentation MathML parity;
3. TypeScript compile of the Windows product;
4. actual WASM `OfficeExportResult` adapter contract tests;
5. existing MathType MTEF consumer semantic/ruler/font-size round-trip using the current x64 VSTO DLL, without Word automation;
6. existing Office SVG→EMF consumer accepting core path SVG and validating the resulting EMF as vector-only;
7. existing MathType alignment, metadata, formula-document, Session payload and apply-shortcut regressions;
8. targeted .NET MathML/MathType consumer tests.

Real Office Session creation, Ribbon interaction, Word document mutation, VSTO service modification and hidden converter takeover are explicitly outside P11 and belong to P12/P13.
