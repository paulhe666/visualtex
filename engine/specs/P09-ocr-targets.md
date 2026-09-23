# P09 deferred OCR insertion targets

Status: implemented and reviewed in the isolated engine/sandbox. The production OCR services, image processing, credentials, model installer, no-OCR packaging and Office integration are not changed.

## Ownership and lifetime

OCR still returns a LaTeX string. Before recognition starts, `EngineEditorHost.beginOcrInsertion()` captures the original `EngineEditorSessionOwner`. The public ticket is an opaque frozen object; it contains no numeric session handle, NodeId or selection and is not part of the recognizer's payload. The owner retains at most one outstanding ticket. Another owner/runtime, a cloned object, a cancelled ticket, or an earlier superseded request cannot use it even when runtimes happen to allocate the same numeric handle.

Rust `captureOcrTarget` creates a bounded pending target with a monotonically increasing request ID. It records session handle, document ID, revision, source ID, selection epoch and full mathematical selection. Capturing does not change source, visible selection, document revision or history. `cancelOcrTarget` cancels only the matching request ID; a late cancellation cannot cancel a newer request.

The runtime checks the original handle before dispatch. `applyOcrLatex` then checks the exact pending record, original revision/source identity and current selection/epoch in the same Rust transaction as the insertion. No temporary selection jump is used. Identical event replay is acknowledged by the existing receipt machinery and never inserts twice. An accepted OCR result consumes the pending target and is one ordinary structural-paste undo record.

## Strict stale-target policy

This first implementation deliberately accepts only an unchanged target. It does not guess offsets, silently rebase a stale selection, or insert at the newest focused editor. Selection away-and-back is stale even when final coordinates match; editing followed by undo is stale even when source text matches. Beginning IME composition permanently invalidates the old request, including after composition cancellation. Parking invalidates the core target; destroying the owner drops it. Host disposal, full-source replacement and starting command-input preedit also retire the host ticket.

A view lifecycle must follow the existing viewport contract: detach the view, park/resume the owner, then attach a new view. Directly parking an owner underneath a still-live view invalidates its layout token and is not a supported view lifecycle. Both direct core park/resume target invalidation and proper browser detach/park/resume/attach are tested.

## Result handling

`applyOcrResult(ticket, latex)` returns either an accepted transaction or a rejection containing the original, unmodified LaTeX and a reason. It does not steal focus, look up a current editor, execute OCR/TeX code, fetch resources, normalize the whole source or auto-retry. Unsupported/incomplete/empty/over-budget payloads fail before document mutation. The product can display the rejected result for a new explicit insertion request; it must not discard or silently retarget it. No OCR target is captured inside a text-run interior or while the host has uncommitted input.

## Evidence and scope

- `crates/core/tests/ocr_target_review.rs`: capture/read-only state, atomic structural insertion, replay, undo/redo, source/selection staleness, request ordering/cancellation, IME, session lifecycle, malformed payloads and original range replacement.
- `tools/ocr-target-parity-review.mjs`: actual Native and WASM requests compared field-for-field, including target tampering, external expectedRevision bypass attempts and resource limits.
- `tools/ocr-target-browser-review.mjs`: owned Edge with trusted pointer/keyboard capture, delayed Promise delivery, focus in B while A receives its own result, cross-runtime tickets, no partial mutations, correct detach/resume and zoom/DPR cases. The recognizer is not invoked: the deferred payload is controlled to test the boundary deterministically.

P09's stable-source gate includes both OCR runners in addition to every previous P08/P09 gate. These tests do not claim OCR model accuracy, actual recognition API success or Word acceptance, and do not make P09 complete by themselves.
