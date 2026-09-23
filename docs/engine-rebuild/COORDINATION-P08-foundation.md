# Active P08 contained-grid foundation

Continuation after immutable CP07-general-operators (354 native tests, real
Native/WASM, browser and archive verification). The original production P00 guard
still reports external Office/evidence differences. This lane touches only engine/
and docs/engine-rebuild/, never resets those external files or changes Git refs.

P08 starts on the verified basic P07 structures. This does not falsely mark all
legacy macros, ESINT-specific artwork, prime shorthand or middle-fence commands
as implemented; those gaps remain in the command capability matrix and gate final
product compatibility. P08 and P09 remain isolated from production UI/Office.

Initial vertical slice: gathered environment in one EngineSession and one formula
frame; Grid and GridRow with ordinary editable Row cells. Grid rows and cells have
separate NodeIds; column metadata uses a typed GridColumnId and the same monotonic
allocation pool, never a source-offset or caret identity. Later matrix/cases/
aligned work extends this grid foundation rather than inventing separate editors.

Shift+Enter becomes one atomic content transaction. At a root row boundary it
creates a gathered group with left/right rows and retains the root/session ID plus
original mathematical children. Within the existing applicable gathered cell it
splits that row. Interior fraction/script/accent/root/text positions and noncaret
selections are initially rejected atomically, not cut via string operations.
Input composition/unfinished input has priority; ordinary Enter stays separate.

Acceptance includes exact lossless source, safe token-level environment scanning,
new identities and undo/redo, explicit PositionMap row split/merge, direct-grid
vertical navigation, common em geometry and SVG, true browser key/pointer tests,
independently mounted A/B formula isolation and IME/preedit no-extra-row checks.
A checkpoint must distinguish which M01-M12 items are actually run. This note is
not a claim that all P08 environments/cell selection/virtualization are complete.

## Coordination update after 22:00 +08:00

Concurrent review added Grid/GridRow semantic arms to four existing helpers while
this implementation did the same. Both writers then removed different duplicates,
leaving none. That explains the latest native build failure, not a data/model
regression. This implementation now reconciles exactly one Grid/GridRow arm in
accents.rs, operators.rs, radicals.rs and parser_contract_review.rs. Review lane:
please do not modify those four helpers or host/core files concurrently; keep
independently named regression files only. The independently added Shift+Enter
host handler is retained, with Home/End and input-port/ARIA changes layered on it.

A subsequent stable-source verification must prove the combined files are frozen
during tests and archive. Earlier mutable-source observations are not acceptance.
