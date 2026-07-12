# ADR 0002: Tolerant parsing with lossless opaque regions

- Status: Accepted
- Date: 2026-07-12

## Context

LaTeX is programmable and projects commonly contain custom macros, incomplete syntax during typing, template-specific environments and verbatim-like regions. A visual editor cannot safely claim to understand every construct.

## Decision

The parser and semantic layer classify nodes as `native`, `partial`, `opaque` or `unstable`.

- Native nodes can be edited structurally.
- Partial nodes expose safe fields while preserving the surrounding source.
- Opaque nodes remain source-only and are never regenerated.
- Unstable nodes represent temporarily incomplete syntax and remain editable in source mode.

Serialization is local: a visual edit replaces only the inner source range belonging to the selected semantic node. Unchanged source bytes are not pretty-printed or normalized.

## Consequences

- Unknown macros and environments survive round trips.
- Visual coverage can expand incrementally through template adapters.
- The UI must visibly distinguish read-only/opaque nodes.
- A future Tree-sitter grammar may replace the current parser implementation without changing the semantic support contract.
