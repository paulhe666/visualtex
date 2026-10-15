import { safeStorage } from "./safeStorage";

const EDITOR_STORAGE_KEY = "visualtex-editor";
const CRASH_BACKUP_STORAGE_KEY = "visualtex-editor.crash-backup.v1";

const riskySessionKeys = [
  "title",
  "latex",
  "lines",
  "activeLineId",
  "history",
] as const;

/**
 * Prepare a crash-safe restart without wiping the user's WebKit profile.
 *
 * The editor persistence envelope mixes long-lived preferences with the last
 * open formula document. Before removing only the document-like fields, keep
 * the complete original envelope under a separate recovery key so no formula
 * source is silently discarded.
 */
export function prepareEditorCrashSafeRestart() {
  const raw = safeStorage.getItem(EDITOR_STORAGE_KEY);
  if (!raw) return { hadPersistedEditor: false, backupCreated: false };

  let cleaned: string | null = null;
  try {
    const parsed = JSON.parse(raw) as {
      state?: Record<string, unknown>;
      [key: string]: unknown;
    };
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
      const state =
        parsed.state &&
        typeof parsed.state === "object" &&
        !Array.isArray(parsed.state)
          ? { ...parsed.state }
          : {};
      for (const key of riskySessionKeys) delete state[key];
      cleaned = JSON.stringify({ ...parsed, state });
    }
  } catch {
    // Invalid JSON is backed up byte-for-byte below. Storage failures are not parse failures.
  }
  const existingBackup = safeStorage.getItem(CRASH_BACKUP_STORAGE_KEY);
  if (existingBackup !== null) {
    if (raw === cleaned) return { hadPersistedEditor: true, backupCreated: false };
    if (existingBackup !== raw) {
      throw new Error("已有待恢复的公式备份。请先保存或恢复备份，再清理新的公式内容。");
    }
  }
  if (existingBackup === null) safeStorage.setItemStrict(CRASH_BACKUP_STORAGE_KEY, raw);
  if (cleaned === null) safeStorage.removeItemStrict(EDITOR_STORAGE_KEY);
  else safeStorage.setItemStrict(EDITOR_STORAGE_KEY, cleaned);
  return { hadPersistedEditor: true, backupCreated: existingBackup === null };
}

export const VISUALTEX_EDITOR_CRASH_BACKUP_STORAGE_KEY =
  CRASH_BACKUP_STORAGE_KEY;
