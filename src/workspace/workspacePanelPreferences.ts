import { readLocalStorage, writeLocalStorage } from "../runtime/safeStorage";

type WorkspacePanelPreference = "toolbar" | "tiles" | "source";

const workspacePanelStorageKeys: Record<WorkspacePanelPreference, string> = {
  toolbar: "visualtex-web-editor-toolbar-open",
  tiles: "visualtex-web-editor-tiles-open",
  source: "visualtex-web-editor-source-open",
};

export function readWorkspacePanelOpen(panel: WorkspacePanelPreference, fallback = true) {
  const stored = readLocalStorage(workspacePanelStorageKeys[panel]);
  if (stored === "true" || stored === "1") return true;
  if (stored === "false" || stored === "0") return false;
  return fallback;
}

export function writeWorkspacePanelOpen(panel: WorkspacePanelPreference, open: boolean) {
  writeLocalStorage(workspacePanelStorageKeys[panel], String(open));
}
