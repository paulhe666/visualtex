import assert from "node:assert/strict";
import {
  createUuid,
  installBrowserCompatibility,
} from "../src/runtime/browserCompatibility.ts";
import { safeStorage } from "../src/runtime/safeStorage.ts";

function override(target, key, value) {
  Object.defineProperty(target, key, {
    configurable: true,
    writable: true,
    value,
  });
}

override(Array.prototype, "at", undefined);
override(String.prototype, "replaceAll", undefined);
override(globalThis, "queueMicrotask", undefined);
override(globalThis, "ResizeObserver", undefined);
if (globalThis.crypto) override(globalThis.crypto, "randomUUID", undefined);

installBrowserCompatibility();

assert.equal(["a", "b", "c"].at(-1), "c");
assert.equal(["a", "b", "c"].at(-4), undefined);
assert.equal("a+b+a".replaceAll("a", "x"), "x+b+x");
assert.equal("ab".replaceAll("", "-"), "-a-b-");

let microtaskRan = false;
queueMicrotask(() => {
  microtaskRan = true;
});
await Promise.resolve();
assert.equal(microtaskRan, true);

let resizeCallbacks = 0;
const observer = new ResizeObserver(() => {
  resizeCallbacks += 1;
});
observer.observe({});
await Promise.resolve();
assert.equal(resizeCallbacks, 1);
observer.disconnect();

const uuid = createUuid();
assert.match(
  uuid,
  /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/,
);

safeStorage.setItem("visualtex-startup-smoke", "ready");
assert.equal(safeStorage.getItem("visualtex-startup-smoke"), "ready");
safeStorage.removeItem("visualtex-startup-smoke");
assert.equal(safeStorage.getItem("visualtex-startup-smoke"), null);

console.log("Startup compatibility smoke test passed");
