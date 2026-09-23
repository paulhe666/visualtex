# Confirmed input bridge failure paths

The independent synthetic component suite `engine/tools/input-bridge-contract-review.mjs` returned 9 passed / 3 failed against BrowserInputBridge at 13:37:
- commitComposition returning false loses compositionId;
- commitComposition throwing loses compositionId;
- rejected plain input unconditionally clears textarea.value.

The normal actual-browser end-to-end test already passes (44 events, 41 trusted, 3 CDP compositionend untrusted). These failure-injection tests are deliberately synthetic InputPort boundaries, not OS IME or core tests.

Focused changes are now complete in input-bridge.ts. PendingInput, retryPendingInput(), discardPendingInput(), hasUncommittedInput are added; BrowserInputBridge/InputPort remain compatible. Enter retries pending input, Escape explicitly discards/cancels it. Failed composition retains its original ID and text. New composition cannot overwrite a pending failure. 15/15 synthetic component cases pass; actual browser E2E again passes all 44 events with no duplicate commit. Implementation lane may resume bridge work after reading the current file.

editor-host.ts is now adapted to BrowserInputBridge and has history(redo), pending-input guards on structure/source/export/pointer operations, and catches drawing errors separately from transaction rejection. The main sandbox now imports this host; no second controller is needed. Please preserve these verified behaviors. No production files are involved.

Actual Windows IME is not yet tested: native focus guard refused every ASCII key because the foreground user's Chrome window covered the test browser. No Unicode or CDP composition substitute was used in that runner. Further unsafe focus escalation is not authorized or necessary for input-failure testing.
