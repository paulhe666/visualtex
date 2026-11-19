// Every command the Windows app registers is allowed by one of its permissions (src-tauri/permissions/*.toml). With
// permissions of its own, the app's commands go through Tauri's access control, which refuses a command no permission
// names ("not allowed by ACL") at run time; neither the compiler nor the unit tests notice it.
import assert from "node:assert/strict";
import { readdir, readFile } from "node:fs/promises";

const lib = await readFile("src-tauri/src/lib.rs", "utf8");
const handlers = [...lib.matchAll(/generate_handler!\[([\s\S]*?)\]/g)];
assert.equal(handlers.length, 1, "one invoke handler in lib.rs");
const commands = handlers[0][1]
  .replace(/\/\/[^\n]*/g, "")
  .split(",")
  .map((entry) => entry.trim())
  .filter(Boolean)
  .map((entry) => entry.split("::").pop());
assert.ok(commands.length > 0, "the invoke handler names its commands");

const allowed = new Set();
for (const name of (await readdir("src-tauri/permissions")).filter((file) => file.endsWith(".toml"))) {
  const text = await readFile(`src-tauri/permissions/${name}`, "utf8");
  for (const list of text.matchAll(/commands\.allow\s*=\s*\[([\s\S]*?)\]/g)) {
    for (const command of list[1].matchAll(/"([^"]+)"/g)) allowed.add(command[1]);
  }
}
const missing = commands.filter((command) => !allowed.has(command));
assert.deepEqual(missing, [], `commands no permission allows: ${missing.join(", ")}`);
console.log(`command permissions: all ${commands.length} commands are allowed`);
