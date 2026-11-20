import { invoke } from "@tauri-apps/api/core";
import type { OfficeHotkeyBinding } from "../shortcuts/officeHotkeys";

export async function configureOfficeHotkeys(
  bindings: readonly OfficeHotkeyBinding[],
) {
  await invoke("configure_office_hotkeys", {
    bindings: bindings.map(({ actionId, chord }) => ({
      actionId,
      chord: {
        code: chord.code,
        ctrlKey: chord.ctrlKey,
        altKey: chord.altKey,
        shiftKey: chord.shiftKey,
        metaKey: chord.metaKey,
      },
    })),
  });
}
