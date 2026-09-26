import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const read = async (path) => (await readFile(path, "utf8")).replace(/\r\n?/g, "\n");

const defaultCapability = JSON.parse(await read("src-tauri/capabilities/default.json"));
const hudCapability = JSON.parse(await read("src-tauri/capabilities/silent-ocr-hud.json"));
const desktopPermission = await read("src-tauri/permissions/ocr-desktop.toml");
const hudPermission = await read("src-tauri/permissions/ocr-silent-hud.toml");

assert.deepEqual(defaultCapability.windows, ["main"]);
assert.ok(
  defaultCapability.permissions.includes("allow-ocr-desktop"),
  "main window must receive the OCR desktop permission",
);

const requiredDesktopCommands = [
  "get_ocr_provider_configuration",
  "save_ocr_provider_configuration",
  "get_ocr_runtime_status",
  "configure_ocr_storage_location",
  "open_ocr_storage_location",
  "install_ocr_runtime",
  "get_ocr_install_status",
  "cancel_ocr_install",
  "open_ocr_install_logs",
  "recognize_formula_image",
  "cancel_ocr_recognition",
  "restart_ocr_worker",
  "warmup_ocr_model",
  "reset_ocr_runtime",
  "install_optional_ocr_model",
  "remove_optional_ocr_model",
  "get_ocr_model_catalog",
  "get_ocr_model_download_status",
  "download_ocr_model",
  "cancel_ocr_model_download",
  "capture_windows_quick_ocr",
  "write_windows_ocr_clipboard_text",
  "configure_silent_ocr",
];

for (const command of requiredDesktopCommands) {
  assert.ok(
    desktopPermission.includes(`"${command}"`),
    `OCR desktop permission is missing ${command}`,
  );
}

assert.deepEqual(hudCapability.windows, ["silent-ocr-hud"]);
assert.ok(
  hudCapability.permissions.includes("allow-ocr-silent-hud-status"),
  "silent OCR HUD must receive its bounded status permission",
);
assert.ok(
  hudPermission.includes('"get_silent_ocr_hud_status"'),
  "silent OCR HUD permission is missing get_silent_ocr_hud_status",
);
assert.ok(
  !hudCapability.permissions.includes("allow-ocr-desktop"),
  "silent OCR HUD must not receive full OCR desktop privileges",
);
for (const command of ["install_ocr_runtime", "recognize_formula_image", "configure_ocr_storage_location"]) {
  assert.ok(
    !hudPermission.includes(`"${command}"`),
    `silent OCR HUD permission unexpectedly exposes ${command}`,
  );
}

console.log(
  "Windows OCR ACL regression passed: main OCR commands allowed and silent HUD remains least-privilege.",
);
