# P12 Windows Office interactive core integration

Status: complete. Real Word and PowerPoint Ribbon/UI acceptance passed on dedicated test documents on 2026-09-15.

## Scope

P12 switches only the interactive Windows Office editor surface to the verified VisualTeX core editor. It does not change the Office Session wire DTO, `useOfficeSession`, Tauri Office routes/state, C# Contracts, Word/PowerPoint VSTO document algorithms, COM/OLE storage, numbering, references, or the hidden `vsto-convert` batch runtime.

The production Windows Office bundle contains the core editor/WASM/font resources. Runtime ownership remains explicit: desktop and interactive Office editor generations may select core after preflight; `vsto-convert` is not allowed to select or load the core editor.

## Ownership and stale-result rules

Interactive Office core export is exposed through `CoreEditorExportPort`. The port carries the retained core owner, core document identity/revision, and verified font-layout resolver. Office export requests are bound to the current `sessionId + formulaId + core documentId + revision + owner` tuple. Any Session switch, formula switch, owner replacement, or revision change retires an in-flight export before it can PATCH the Session.

Initial Session hydration does not mark the formula dirty and does not create a core edit transaction. Core content history remains the only content-history owner while the interactive formula is core-owned; the existing Office Undo/Redo buttons delegate to core history rather than maintaining a second legacy content history.

## Export behavior

Interactive core-owned formulas use the P11 `coreOfficeCompatibility` adapter. SVG, Presentation MathML, natural CSS-pixel geometry, baseline and optional PNG all come from the same retained core revision. MathType output keeps Times-family output metadata and P11 ruler-stop generation while the core layout itself only accepts verified mathematical fonts with an OpenType MATH table; if the requested visual family has no mathematical face, layout falls back to the verified Latin Modern Math resource instead of treating a normal text font such as Times New Roman as a math font. Word OMML keeps the MathML-only/preserved-geometry path.

The external Office WebView obtains verified system-font bytes through the existing Tauri app command `read_system_font_asset`. A dedicated remote capability is restricted to the `office-session-editor` window and `https://127.0.0.1:43127/*`, with a single custom permission for that command. No Office companion HTTP route, Session route or Office server protocol was added.

Autosave is two-stage for core-owned formulas: source/state are persisted immediately through the existing Session save path; the asynchronous export PATCH is accepted only if the exact Session/formula/owner/revision is still current. Explicit Apply waits for the current core export before sending the existing commit signal. Cancel continues to use the existing Session cancellation/close path and never asks core to set Office Session status.

Tagged formulas remain outside the P12 core gate because the current core tree does not own the Office equation-tag wrapper. This is an explicit eligibility rule, not a MathJax fallback inside a core session.

## Hidden converter boundary

`vsto-convert` remains P13 territory. It does not mount the core editor, create a core input bridge, request the core WASM/font assets, or use P11 core export in P12. Existing `generateSessionExportResult()` behavior remains the hidden converter path.

## Browser/contract verification

The production Windows Office bundle is exercised in an isolated real Edge instance with a mock companion implementing the existing Session routes. The browser gates cover interactive core selection, no MathLive field in a core-owned editor, WASM/font resource load, initial clean hydration, real keyboard edit, immediate source PATCH, same-revision SVG/MathML/PNG/baseline export, rapid-edit stale export rejection, Session-switch stale owner rejection, core Undo/Redo, theme synchronization, Apply, Cancel, and `vsto-convert` remaining legacy. P11 compatibility verification is rerun after P12 changes.

## Real Office acceptance

The dedicated Word matrix passed through the real VisualTeX Ribbon/UI. Verified cases include VisualTeX OLE inline creation and in-place edit, display VisualTeX OLE with numbering, display VisualTeX OLE without numbering, inline Word OMML, and numbered display MathType OLE. The MathType object was verified as `Equation.DSMT4`. Cancel left the document unchanged; native window X honored `autoCommitOnClose`; the same companion process and WebView were reused across Sessions; repeated edits updated the original object rather than duplicating it.

The Word test document was saved, closed, reopened and structurally re-inspected. It retained four OLE objects, one OMath, eleven fields, seven bookmarks, one VisualTeX SEQ field and five MathType numbering fields. Reopening the first VisualTeX OLE through the real Ribbon restored exact source `x^{2}+1+z+q` into the core editor.

PowerPoint was tested only through its current supported interactive format set. Real Ribbon `New Formula` produced a `crossPlatformPicture` Session whose SVG/PNG/MathML export came from core (`HasMathJax=false`), inserted a `VisualTeX_...` shape and persisted VisualTeX metadata in `AlternativeText`. Real `Edit Selected Formula` restored the formula source and updated the existing shape rather than inserting a duplicate. The user independently confirmed PowerPoint insertion works.

The acceptance record is frozen in `engine/specs/P12-real-office-acceptance.json`.

## Protection result

No Word VSTO, PowerPoint VSTO, Contracts, VstoShared, WindowsOleBridge, Office Session state/wire, numbering, references or Office server route changes are part of P12. User Word PID 63776 was kept open and untouched during the dedicated acceptance runs. Dedicated Word/PowerPoint test processes were closed after verification.

## Completion rule

P12 is complete. P13 is permitted to start, but P13 must remain a separate phase and may only switch the hidden conversion/bulk-import/redraw outlet after re-establishing the P12 frozen checkpoint. No P12 result authorizes changing WordFormulaService, numbering, COM/OLE contracts or Office Session semantics.
