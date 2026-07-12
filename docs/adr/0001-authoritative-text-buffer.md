# ADR 0001: LaTeX source buffer is authoritative

- Status: Accepted
- Date: 2026-07-12

## Context

VisualTeX exposes source, structured and PDF-oriented views. Allowing each view to own a separate document model would create merge ambiguity, feedback loops and data loss, especially when an external editor changes the same files.

## Decision

The Rust `DocumentBuffer` is the only authoritative mutable representation. It stores UTF-8 text in a Rope and accepts revisioned byte-range `TextEdit` operations. Every client receives snapshots or patches and must submit edits against an explicit base revision.

ProseMirror, MathLive, PDF overlays, OCR and plugins are projections or operation producers. They never save independent paper content. VS Code is the one host-specific exception: its `TextDocument` remains authoritative, and the extension mirrors it into Core while applying visual changes through `WorkspaceEdit`.

## Consequences

- Stale operations are rejected instead of silently merged.
- Undo/redo order is shared across source and visual views.
- Unknown LaTeX remains in the authoritative text even when no visual renderer understands it.
- All selection mappings use explicit UTF-8 byte spans at the Core boundary.
- Clients need rebase/reload behavior after revision conflicts.
