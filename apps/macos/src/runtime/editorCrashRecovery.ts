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
 * open formula document. A malformed/pathological persisted formula can make
 * MathLive fail again on every startup. Before removing only the document-like
 * fields, keep the complete original envelope under a separate recovery key so
 * no formula source is silently discarded.
 */
export function prepareEditorCrashSafeRestart() {
  const raw = safeStorage.getItem(EDITOR_STORAGE_KEY);
  if (!raw) return { hadPersistedEditor: false, backupCreated: false };

  // Do not mutate the active editor state until a byte-for-byte backup is
  // persisted successfully.
  safeStorage.setItemStrict(CRASH_BACKUP_STORAGE_KEY, raw);

  try {
    const parsed = JSON.parse(raw) as {
      state?: Record<string, unknown>;
      [key: string]: unknown;
    };
    if (!parsed || typeof parsed !== "object") {
      safeStorage.removeItem(EDITOR_STORAGE_KEY);
      return { hadPersistedEditor: true, backupCreated: true };
    }

    const state =
      parsed.state && typeof parsed.state === "object" && !Array.isArray(parsed.state)
        ? { ...parsed.state }
        : {};
    for (const key of riskySessionKeys) delete state[key];

    safeStorage.setItemStrict(
      EDITOR_STORAGE_KEY,
      JSON.stringify({ ...parsed, state }),
    );
    return { hadPersistedEditor: true, backupCreated: true };
  } catch {
    // Corrupt JSON cannot be selectively recovered. The original bytes are
    // already preserved above, so remove only the unreadable active envelope.
    safeStorage.removeItem(EDITOR_STORAGE_KEY);
    return { hadPersistedEditor: true, backupCreated: true };
  }
}

export const VISUALTEX_EDITOR_CRASH_BACKUP_STORAGE_KEY =
  CRASH_BACKUP_STORAGE_KEY;
