import assert from "node:assert/strict";
import { createSafeStorage as windowsStorage } from "../src/runtime/safeStorage.ts";
import { createSafeStorage as macosStorage } from "../../macos/src/runtime/safeStorage.ts";
import { memoryStorage } from "./memory_storage.ts";

for (const factory of [windowsStorage, macosStorage]) {
  const disk = memoryStorage();
  let writeFails = false;
  const backend: Storage = { ...disk,
    get length() { return disk.length; },
    setItem(key, value) { if (writeFails) throw new Error("quota"); disk.setItem(key, value); },
    removeItem(key) { if (writeFails) throw new Error("quota"); disk.removeItem(key); },
  };
  const storage = factory(() => backend);
  storage.setItem("key", "old");
  writeFails = true;
  storage.setItem("key", "unsaved");
  assert.equal(storage.getItem("key"), "unsaved");
  assert.equal(disk.getItem("key"), "old");
  assert.throws(() => storage.setItemStrict("key", "strict"), /quota/);
  assert.equal(storage.getItem("key"), "unsaved");
  storage.removeItem("key");
  assert.equal(storage.getItem("key"), null, "failed delete must not revive old disk value");
  writeFails = false;
  storage.setItemStrict("key", "durable");
  assert.equal(factory(() => backend).getItem("key"), "durable");
  disk.setItem("key", "other window");
  assert.equal(storage.getItem("key"), "other window");
  storage.removeItemStrict("key");
  assert.equal(disk.getItem("key"), null);
  for (const getBackend of [() => null, () => { throw new Error("denied"); }]) {
    const unavailable = factory(getBackend);
    assert.throws(() => unavailable.setItemStrict("asset", "data"));
    assert.throws(() => unavailable.removeItemStrict("asset"));
    unavailable.setItem("preference", "session only");
    assert.equal(unavailable.getItem("preference"), "session only");
  }
}
console.log("Windows and macOS storage contracts passed");
