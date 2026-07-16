import assert from "node:assert/strict";
import {
  ImeCompositionGuard,
  POST_COMPOSITION_ENTER_WINDOW_MS,
} from "../src/editor/imeCompositionGuard.ts";

function keyEvent(overrides = {}) {
  return {
    key: "a",
    isComposing: false,
    keyCode: 65,
    ...overrides,
  };
}

{
  const guard = new ImeCompositionGuard();
  guard.compositionStart();
  assert.equal(guard.isComposing(), true);
  assert.equal(
    guard.keyDown(
      keyEvent({ key: "Enter", isComposing: true, keyCode: 229 }),
      10,
    ),
    "composition",
    "Enter must remain owned by the active IME transaction",
  );
  guard.compositionEnd(20);
  assert.equal(guard.isComposing(), false);
  assert.equal(
    guard.keyDown(keyEvent({ key: "Enter", keyCode: 13 }), 21),
    "post-composition-enter",
    "WebKit's Enter replay after compositionend must be swallowed",
  );
  assert.equal(
    guard.keyDown(keyEvent({ key: "Enter", keyCode: 13 }), 22),
    null,
    "the post-composition guard must be one-shot",
  );
}

{
  const guard = new ImeCompositionGuard();
  guard.compositionStart();
  guard.compositionEnd(100);
  assert.equal(
    guard.keyDown(
      keyEvent({ key: "Enter", keyCode: 13 }),
      100 + POST_COMPOSITION_ENTER_WINDOW_MS + 1,
    ),
    null,
    "a deliberate later Enter must still create a formula line",
  );
}

{
  const guard = new ImeCompositionGuard();
  assert.equal(
    guard.keyDown(keyEvent({ key: "x", keyCode: 229 }), 1),
    "composition",
    "legacy keyCode 229 must be treated as IME composition",
  );
}

{
  const guard = new ImeCompositionGuard();
  guard.compositionStart();
  assert.equal(
    guard.keyDown(keyEvent({ key: "ArrowRight", keyCode: 39 }), 30),
    null,
    "a real non-composing key must release stale local composition state",
  );
  assert.equal(guard.isComposing(), false);
}

{
  const guard = new ImeCompositionGuard();
  guard.compositionStart();
  guard.compositionEnd(50);
  assert.equal(
    guard.keyDown(keyEvent({ key: "Shift", keyCode: 16 }), 51),
    null,
  );
  assert.equal(
    guard.keyDown(keyEvent({ key: "Enter", keyCode: 13 }), 52),
    "post-composition-enter",
    "modifier-only events must not disarm the immediate Enter replay guard",
  );
}

process.stdout.write("VisualTeX IME Enter regression: PASS\n");
