# ADR 0005: Tree-sitter incremental syntax below the tolerant semantic layer

- Status: Accepted
- Date: 2026-07-12

## Context

VisualTeX requires fast parsing while the user is typing incomplete LaTeX, but LaTeX is programmable and cannot be made fully lossless by regenerating source from a conventional AST. The existing semantic layer already defines the safe editing contract through native, partial, opaque and unstable nodes.

Project-wide authoring also needs an include graph that follows unsaved buffer edits, detects missing targets and reports cycles without making generated files authoritative.

## Decision

A new `vt-latex-syntax` crate owns a Tree-sitter LaTeX parser and one syntax tree per open LaTeX-family source file.

- The grammar is provided by `codebook-tree-sitter-latex`.
- Core text edits are first applied to a candidate revisioned Rope buffer.
- The corresponding Tree-sitter `InputEdit` is calculated from UTF-8 byte offsets and row/column points.
- Tree-sitter reparses against the edited previous tree.
- If syntax updating succeeds, Core commits the candidate buffer and refreshes the existing tolerant semantic document.
- If syntax updating fails, the authoritative buffer is not committed.
- Tree-sitter error and missing nodes are recorded as syntax issues rather than rejecting incomplete source.
- Verbatim, minted and code-like environment ranges are recorded so higher layers can avoid interpreting their contents as normal LaTeX.
- `\input`, `\include`, `\subfile` and `\subfileinclude` commands produce project-relative dependency edges.
- The graph reports resolved and unresolved targets and deterministic include cycles.
- Generated and dependency directories such as `.visualtex`, `.git`, `target` and `node_modules` are excluded from project source discovery.

The existing semantic layer remains responsible for visual support levels, stable protocol nodes and local source serialization. Tree-sitter does not become a second source of truth.

## Consequences

- Source edits, visual edits, OCR edits, undo and redo all share the same incremental syntax path.
- Syntax trees remain derived state and can be rebuilt from the authoritative source buffer.
- A malformed document can still be edited and saved.
- Include graph changes are visible before save and through JSON-RPC, Tauri IPC, the CLI and the Node Adapter SDK.
- Full macro expansion, package semantics and universal include resolution remain separate future work.
- Parser upgrades require corpus and performance regression tests because grammar changes can alter node kinds and recovery behavior.
