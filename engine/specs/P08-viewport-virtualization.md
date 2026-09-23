# P08 long-document formula viewport virtualization

Status: implemented isolated engine/host slice; production application remains
unchanged until its later integration phase.

## Ownership rule

A formula's mathematical state belongs to one `EngineSession`, never to its DOM
view. `EngineEditorSessionOwner` holds the core handle, current editor snapshot and
event sequence. `EngineEditorHost` is a disposable view/input attachment for that
owner. Reattaching a formula must resume the same owner and handle; recreating the
formula from LaTeX is not an acceptable virtualization mechanism because it would
replace node identities, selection and undo history.

Unvisited formulas are represented only by their persisted source record and a
passive viewport shell. The first materialization creates the session. A formula
which has been materialized and then leaves the viewport is **parked**: the exact
`EngineSession` object moves out of the active runtime map while layout/scene
caches and input/SVG DOM are released. Resume moves that same object back to the
active map.

## Runtime budgets

- Active editor sessions remain capped at 64. Virtualization does not raise this
  existing protection to fit a large document.
- Parked sessions have a separate cap of 2048 and a conservative retained-weight
  budget of 512 MiB accounting units. The number is a resource budget, not a claim
  that the value equals process RSS.
- Parking removes the session's layout engine and current scene token/cache.
- Document identity remains reserved while parked. Creating another session with
  the same document ID is a conflict.
- Session handles are monotonic and never recycled by park/resume.
- Destroy works for active or parked sessions and releases document ownership and
  parked accounting.

`Stats` keeps its existing wire shape and still reports only active sessions, so
adding virtualization does not silently change existing protocol consumers. WASM
also exposes test/host getters for parked count and retained weight.

## Input and lifecycle safety

A session with an active composition cannot be parked. The browser viewport pins
any formula whose native input still has pending/composition state or whose input
owns focus. After composition finishes and focus moves to another visible formula,
the offscreen formula becomes eligible for parking.

Park/resume itself is not an edit: revision, source, tree IDs, selection, history,
redo state, event receipts and formula/document identity are unchanged. A replayed
browser/native event must remain recognized after resume. Undo/redo after resume
must restore the same internal identities recorded before parking.

## DOM virtualization

`VirtualFormulaViewport` owns passive shells for document flow. Only formulas in
or near the scroll viewport receive `EngineEditorHost`, SVG and textarea DOM.
Default maximum mounted hosts is 32 and the configurable upper bound is 48, leaving
headroom under the global 64 active-session budget for other isolated hosts.
Offscreen shells retain measured minimum height to avoid scroll collapse but contain
no formula input owner and no SVG scene.

IntersectionObserver drives normal mount/unmount. A deterministic `refresh()` hook
exists for tests and host coordination; it does not own formula state. Focused or
composing formulas may temporarily remain mounted offscreen until input ownership
moves elsewhere.

## Acceptance actually required for this slice

1. Native core: park/resume exact editor state, history, node identities, event
   receipts and handle; stale revisions and composition refusals are atomic.
2. At least 1000 simple retained formula sessions can be parked while active count
   remains bounded, then resumed and destroyed without identity reuse or leaks.
3. Native CLI/WASM protocol parity includes park/resume while legacy protocol cases
   remain unchanged.
4. Real browser: 1000 formula shells, complete viewport traversal, no accumulation
   of textarea/SVG/editor hosts, active session count equal to mounted host count.
5. Edit one formula, park it, traverse the long document, resume it and compare the
   complete snapshot byte-for-byte before testing undo/redo identity restoration.
6. Mixed Matrix/Cases/Aligned/Fraction samples in the long document remain valid;
   virtualization is structure-agnostic and does not flatten them.
7. Browser-protocol composition pins the offscreen formula until composition ends;
   this does not substitute for physical Windows IME acceptance.
8. Disposing the virtual document releases active and parked sessions and removes
   all SVG/input DOM.

## Explicit non-claims

This slice does not switch the VisualTeX React application to the new engine, does
not touch Office/VSTO/OLE/session contracts, does not claim physical Windows IME
acceptance, and does not remove MathLive/MathJax. Those remain later phase gates.
