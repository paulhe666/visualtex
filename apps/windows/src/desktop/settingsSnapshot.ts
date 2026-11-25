import { invoke, isTauri } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";

/**
 * A copy of the app's settings in VisualTeX's settings folder (settings-snapshot.json): everything the main
 * window keeps in localStorage (theme, editor settings, hotkeys, custom tiles, colours, the document library
 * state). That storage sits in the WebView folder, which is mostly cache and which an uninstall deletes; when
 * VisualTeX starts with a saved key missing, the main window puts it back before the page loads
 * (settings_snapshot.rs). Saved when something changed, at most every few seconds, and when the window hides.
 */
const SAVE_INTERVAL_MS = 4000;

function readLocalStorage(): Record<string, string> {
  const entries: Record<string, string> = {};
  for (let index = 0; index < window.localStorage.length; index += 1) {
    const key = window.localStorage.key(index);
    if (key === null) continue;
    const value = window.localStorage.getItem(key);
    if (value !== null) entries[key] = value;
  }
  return entries;
}

/** The app's main window (the other windows share its storage; a page outside the app, as in the browser tests, has none). */
function inMainWindow(): boolean {
  return isTauri() && getCurrentWindow().label === "main";
}

/** Saves the copy now (before VisualTeX quits for an update); resolves once it is written or failed. */
export async function saveSettingsSnapshotNow(): Promise<void> {
  if (!inMainWindow()) return;
  try {
    await invoke("save_settings_snapshot", { snapshot: JSON.stringify(readLocalStorage()) });
  } catch {
    // the copy saved every few seconds stays
  }
}

export function startSettingsSnapshot() {
  if (!inMainWindow()) return;
  let saved = "";
  const save = () => {
    let snapshot: string;
    try {
      snapshot = JSON.stringify(readLocalStorage());
    } catch {
      return;
    }
    if (snapshot === saved) return;
    saved = snapshot;
    void invoke("save_settings_snapshot", { snapshot }).catch(() => {
      saved = "";
    });
  };
  window.setTimeout(save, 1500);
  window.setInterval(save, SAVE_INTERVAL_MS);
  window.addEventListener("pagehide", save);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") save();
  });
}
