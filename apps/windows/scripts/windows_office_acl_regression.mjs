import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const read = async (path) => (await readFile(path, "utf8")).replace(/\r\n?/g, "\n");

const defaultCapability = JSON.parse(await read("src-tauri/capabilities/default.json"));
let officeSessionCapability = null;
try {
  officeSessionCapability = JSON.parse(
    await read("src-tauri/capabilities/office-session-editor.json"),
  );
} catch {
  // The remote Office capability belongs to a separate feature stream and is
  // optional for this desktop ACL regression.
}
const permission = await read(
  "src-tauri/permissions/office-desktop-integration.toml",
);

assert.deepEqual(defaultCapability.windows, ["main"]);
assert.ok(
  defaultCapability.permissions.includes("allow-office-desktop-integration"),
  "main window must receive the Office desktop integration permission",
);

const requiredCommands = [
  "get_office_platform_status",
  "get_office_companion_status",
  "get_mathtype_double_click_edit_enabled",
  "set_mathtype_double_click_edit_enabled",
  "open_word",
  "open_powerpoint",
  "install_windows_ole_integration",
  "uninstall_windows_ole_integration",
  "repair_windows_office_integration",
  "test_windows_office_runtime",
  "open_windows_office_logs",
  "set_office_background_start",
  "start_office_companion",
  "stop_office_companion",
];
for (const command of requiredCommands) {
  assert.ok(
    permission.includes(`"${command}"`),
    `Office desktop permission is missing ${command}`,
  );
}

if (officeSessionCapability) {
  assert.ok(
    !officeSessionCapability.permissions.includes("allow-office-desktop-integration"),
    "remote Office editor WebViews must not receive desktop Office lifecycle privileges",
  );
  for (const command of ["open_word", "open_powerpoint", "get_office_platform_status"]) {
    assert.ok(
      !JSON.stringify(officeSessionCapability).includes(command),
      `remote Office editor capability unexpectedly exposes ${command}`,
    );
  }
}

console.log(
  "Windows Office ACL regression passed: main window allowed; remote capability remains isolated when present.",
);
