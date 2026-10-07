import assert from "node:assert/strict";
import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";

// Every Tauri command the desktop frontend invokes by name must be allowed by a
// permission granted to some window capability, and the main-window features
// below must be granted to main. A missing entry only fails at runtime
// ("Command ... not allowed by ACL"), e.g. export / settings save in 1.2.8.

const read = async (path) => (await readFile(path, "utf8")).replace(/\r\n?/g, "\n");

async function listSourceFiles(dir) {
  const files = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) files.push(...(await listSourceFiles(path)));
    else if (/\.(ts|tsx|js|jsx|mjs)$/.test(entry.name)) files.push(path);
  }
  return files;
}

const lib = await read("src-tauri/src/lib.rs");
const handlerBlock = lib.match(/generate_handler!\[([\s\S]*?)\]\)/);
assert.ok(handlerBlock, "lib.rs has no tauri::generate_handler! block");
const registered = new Set(
  handlerBlock[1]
    .replace(/\/\/.*$/gm, "")
    .split(",")
    .map((item) => item.trim().split("::").pop())
    .filter((name) => /^[a-z_0-9]+$/.test(name ?? "")),
);

const permissionsByIdentifier = new Map();
for (const name of await readdir("src-tauri/permissions")) {
  if (!name.endsWith(".toml")) continue;
  const text = await read(join("src-tauri/permissions", name));
  for (const block of text.split("[[permission]]").slice(1)) {
    const identifier = block.match(/identifier\s*=\s*"([^"]+)"/)?.[1];
    const allow = block.match(/commands\.allow\s*=\s*\[([\s\S]*?)\]/)?.[1] ?? "";
    if (identifier) {
      permissionsByIdentifier.set(
        identifier,
        [...allow.matchAll(/"([a-z_0-9]+)"/g)].map((match) => match[1]),
      );
    }
  }
}

const allowedBy = (capability) =>
  capability.permissions
    .filter((permission) => typeof permission === "string")
    .flatMap((permission) => permissionsByIdentifier.get(permission) ?? []);

const defaultCapability = JSON.parse(await read("src-tauri/capabilities/default.json"));
assert.deepEqual(defaultCapability.windows, ["main"]);
const allowedForMain = new Set(allowedBy(defaultCapability));
const allowedAnywhere = new Set();
for (const name of await readdir("src-tauri/capabilities")) {
  if (!name.endsWith(".json")) continue;
  const capability = JSON.parse(await read(join("src-tauri/capabilities", name)));
  for (const command of allowedBy(capability)) allowedAnywhere.add(command);
}

for (const command of [
  "write_export_file",
  "minimize_visualtex_main_window",
  "get_word_numbering_user_configuration",
  "apply_word_numbering_user_configuration",
  "extract_system_math_glyph",
]) {
  assert.ok(allowedForMain.has(command), `main window permission is missing ${command}`);
}

const used = new Set();
for (const file of await listSourceFiles("src")) {
  if (/\.test\.|__tests__/.test(file)) continue;
  const text = await read(file);
  for (const command of registered) {
    if (text.includes(`"${command}"`) || text.includes(`'${command}'`)) used.add(command);
  }
}

const missing = [...used].filter((command) => !allowedAnywhere.has(command)).sort();
assert.deepEqual(
  missing,
  [],
  `frontend invokes commands that no capability allows: ${missing.join(", ")}`,
);

console.log(
  `Windows main ACL regression passed: all ${used.size} frontend-invoked commands are allowed by a capability.`,
);
