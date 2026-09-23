# P09 versioned structural clipboard

Status: implemented and accepted in the isolated engine/sandbox scope. It does not replace the product's persistence format or Office clipboard behavior.

## Formats

The browser bridge publishes three clipboard flavors for an accepted mathematical selection:

- `text/plain`: portable source fallback.
- `application/x-latex`: explicit LaTeX fallback.
- `application/vnd.visualtex.engine-fragment+json`: VisualTeX internal flavor.

The protocol-level `ClipboardFragment` remains schema version 1 for compatibility. Its optional `structure` member has its own independent format:

- `schema = "visualtex-engine-fragment"`
- `version = 1`
- SHA-256 binding to the exact LaTeX fallback
- identity-free semantic `shape`
- bounded `assets` list reserved for the custom-symbol slice

The custom browser flavor uses `schemaVersion: 2` only when the structure member is present. Existing custom v1 `{schemaVersion:1, latex}` remains readable.

## Trust boundary

A structural payload is never hydrated directly into a live mathematical tree. Core paste performs all of the following before committing:

1. validate schema/version and payload/source budgets;
2. validate the SHA-256 binding to the supplied LaTeX fallback;
3. reject unsupported asset references in the current slice;
4. validate JSON nesting/object/string budgets and known structural node kinds;
5. independently parse the LaTeX with the self-owned parser;
6. derive a fresh identity-free shape from that parse and require exact semantic equality with the clipboard shape;
7. execute the normal paste transaction, which allocates fresh live identities.

This prevents custom clipboard JSON from bypassing the parser, importing live NodeId/session identity, or executing HTML/JavaScript/TeX definitions.

## Identity and history

The structural shape excludes live `id`, `origin`, NodeId, document/session/handle/source identities, row trivia ownership and Grid column IDs. Grid column count remains semantic metadata. Exact source spelling is protected separately by the source digest.

Paste into the same formula allocates independent nodes; paste into a rectangular Grid selection retains destination Grid/cell/column identities while creating fresh descendants. A successful structured paste is one normal history transaction and participates in undo/redo.

## Failure and fallback

- Unsupported or incomplete copied source may still expose LaTeX/text fallback but receives no structural trust marker.
- Invalid/tampered/oversized structure is rejected atomically by core.
- Corrupt/unrelated custom MIME cannot suppress valid plain/LaTeX fallback at the browser boundary.
- HTML is ignored and never evaluated.
- Asset-bearing structural payloads are deliberately refused until the following custom-symbol slice defines an asset snapshot contract.

## Acceptance

- Rust structural clipboard contract tests cover complex macros/Array/MiddleDelimiter content, fresh identities, digest/shape/version/schema tampering, asset refusal, budget enforcement and history.
- The input-bridge contract covers v2/v1/LaTeX/plain fallback and inert HTML.
- Owned Edge uses trusted Ctrl+C/Ctrl+X/Ctrl+V on a real matrix rectangle and confirms custom v2, `application/x-latex`, plain fallback, destination identity retention, undo/redo, zoom/DPR and independent formula-frame isolation.

No claim here covers arbitrary spreadsheet formats, physical Office clipboard integration or custom-symbol assets yet.