# ADR 0003: Preview the real TeX build artifact

- Status: Accepted
- Date: 2026-07-12

## Context

An HTML approximation cannot guarantee line breaks, fonts, floats, references or page geometry identical to the exported paper. VisualTeX must not show a convenient preview that differs from the final PDF.

## Decision

The preview source is the PDF produced by the configured local TeX toolchain. Compilation happens in `.visualtex/build`, and successful artifacts include the source revision and SyncTeX file. The UI atomically switches only to a successful PDF.

SyncTeX is the baseline source/page mapping. A later shadow-build instrumenter may add semantic node boxes only when it can prove the instrumented build does not change layout.

## Consequences

- Users see the same artifact they export.
- Structured editing does not require compilation, but page updates do.
- Build failures preserve source edits and diagnostics instead of replacing the last good preview.
- PDF rendering and page overlays are replaceable view layers; they cannot become a second document authority.
