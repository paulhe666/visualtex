# CP09 OCR target 1

Independent non-Git checkpoint following CP09-custom-symbol-text-1. P09 is not marked complete.

Stable source SHA256: `dc500cf75a1fef8bb46e4895c9bdea99789b93e1b55ecd23e844bdc0fc112e49`

Verification receipt: `engine/target/p09-verification/2026-09-14T08-44-48-242Z/report.json`

The expanded stable-source gate passed 16/16 with sourceUnchanged=true and 220 verified source inputs. It includes the previous P08/P09 gates plus actual OCR Native/WASM protocol parity and owned Edge deferred-insertion acceptance. OCR Rust review passed 10/10. OCR parity compared 124 requests exactly, including 20 error paths. Owned Edge recorded 23 trusted button events, 22 delayed deliveries (10 rejected and retained), nine zoom/DPR cases, unchanged B content and no stolen focus. Recognizer APIs/models were not invoked; delayed LaTeX results were controlled to test the insertion boundary.

Core captureOcrTarget retains one monotonically identified request per session. applyOcrLatex validates the exact pending target, session handle, document ID, original revision/source ID and selection epoch/selection, then inserts through the existing structural transaction and shared undo history. Old/superseded/cancelled, IME-interrupted, parked, destroyed and reopened targets fail without retargeting. The host uses opaque object-identity tickets so a different owner or runtime cannot consume a ticket even when handles match. Rejections return the original LaTeX; delivery never steals focus or writes the currently active formula instead.

The first owned-Edge run was FAILED because its lifecycle test parked a still-mounted owner and left a stale layout token. That failed report is retained at engine/target/ocr-target-browser-review/2026-09-14T08-42-36-919Z/report.json. The test was corrected to follow the existing detach -> park/resume -> attach contract. Direct core park invalidation remains independently tested. The later standalone run and the complete stable gate both passed.

All edits are confined to engine/ and new checkpoint archives. No production OCR/Office/VSTO path or Word document was changed. No Git mutations were made. P00 protection remains red due to pre-existing parallel Office/VSTO/evidence modifications; this checkpoint does not relabel it as green. Branch remains reference-3d207d7 at 00b48e17b736c7e57a0a72e19e271df82906f89d.
