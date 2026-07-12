# ADR 0004: Restricted local execution by default

- Status: Accepted
- Date: 2026-07-12

## Context

Opening a LaTeX project can trigger compiler processes, read included files and optionally execute arbitrary commands through shell escape. OCR and editor bridges add further local-process and filesystem boundaries.

## Decision

VisualTeX defaults to restricted mode:

- shell escape is disabled and rejected;
- subprocesses are launched without a shell;
- project, build and OCR paths are canonicalized and scoped;
- symlinked source files are rejected;
- the general bridge is stdio-only until an authenticated local transport is implemented;
- OCR has no downloader or network client in the base worker;
- VS Code Core launch requires a trusted workspace.

## Consequences

- Some legitimate minted, externalization and custom-build projects require a future explicit trusted mode.
- Security checks live in Core or worker managers, not only in UI validation.
- Plugin and socket designs must preserve the same capability boundaries.
