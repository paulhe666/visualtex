import { readLocalStorage, writeLocalStorage } from "../runtime/safeStorage";
import type { WorkspaceMode } from "./workspaceTypes";

export type WorkspacePanelPreference = "toolbar" | "tiles" | "source";
type WorkspacePanelScope = "web" | "office";

const workspacePanelStorageKeys: Record<
  WorkspacePanelScope,
  Record<WorkspacePanelPreference, string>
> = {
  web: {
    toolbar: "visualtex-web-editor-toolbar-open",
    tiles: "visualtex-web-editor-tiles-open",
    source: "visualtex-web-editor-source-open",
  },
  office: {
    toolbar: "visualtex-office-editor-toolbar-open",
    tiles: "visualtex-office-editor-tiles-open",
    source: "visualtex-office-editor-source-open",
  },
};

function workspacePanelScope(mode: WorkspaceMode): WorkspacePanelScope {
  return mode === "web" ? "web" : "office";
}

export function workspacePanelStorageKey(
  mode: WorkspaceMode,
  panel: WorkspacePanelPreference,
) {
  return workspacePanelStorageKeys[workspacePanelScope(mode)][panel];
}

export function readWorkspacePanelOpen(
  mode: WorkspaceMode,
  panel: WorkspacePanelPreference,
  fallback = true,
) {
  const stored = readLocalStorage(workspacePanelStorageKey(mode, panel));
  if (stored === "true" || stored === "1") return true;
  if (stored === "false" || stored === "0") return false;
  return fallback;
}

export function writeWorkspacePanelOpen(
  mode: WorkspaceMode,
  panel: WorkspacePanelPreference,
  open: boolean,
) {
  writeLocalStorage(workspacePanelStorageKey(mode, panel), String(open));
}
