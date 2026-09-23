import { safeStorage } from "./safeStorage";

interface EditorStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

interface EditorEnvelope {
  state: Record<string, unknown>;
  [key: string]: unknown;
}

const editorStorageKey = "visualtex-editor";
const officeSessionKeys = [
  "title",
  "lines",
  "activeLineId",
  "formulaAlignment",
  "latexCodeFormat",
  "history",
] as const;

export function isOfficeEditorPersistenceScope(
  location: Pick<Location, "pathname" | "search"> | null =
    typeof window === "undefined" ? null : window.location,
): boolean {
  if (!location) return false;
  const view = new URLSearchParams(location.search).get("view");
  return location.pathname.endsWith("/office-native-dialog.html") ||
    view === "office-formula" ||
    view === "office-document-import" ||
    view === "office-word-latex-redraw";
}

function parseEnvelope(raw: string | null): EditorEnvelope | null {
  if (raw === null) return null;
  try {
    const value: unknown = JSON.parse(raw);
    if (!value || typeof value !== "object" || Array.isArray(value)) return null;
    const state = (value as { state?: unknown }).state;
    if (!state || typeof state !== "object" || Array.isArray(state)) return null;
    return value as EditorEnvelope;
  } catch {
    return null;
  }
}

/**
 * macOS resident Office WebViews share the main application's localStorage.
 * They may exchange preferences (including latexFormatProfile), but their
 * per-session document and undo/history state must never replace the main
 * document. Enforce this at the persistence boundary, for every store action.
 */
export function createEditorPersistenceStorage(
  storage: EditorStorage = safeStorage,
  officeScope: () => boolean = isOfficeEditorPersistenceScope,
): EditorStorage {
  return {
    getItem(key) {
      const raw = storage.getItem(key);
      if (key !== editorStorageKey || !officeScope()) return raw;
      const envelope = parseEnvelope(raw);
      if (!envelope) return null;
      const state = { ...envelope.state };
      for (const sessionKey of officeSessionKeys) delete state[sessionKey];
      return JSON.stringify({ ...envelope, state });
    },
    setItem(key, value) {
      if (key !== editorStorageKey || !officeScope()) {
        storage.setItem(key, value);
        return;
      }
      const incoming = parseEnvelope(value);
      if (!incoming) return;
      const previousRaw = storage.getItem(key);
      const previous = parseEnvelope(previousRaw);
      // Do not overwrite unparseable existing user data with a session's
      // default state. Recovery belongs to the main application, not Office.
      if (previousRaw !== null && !previous) return;
      const state = { ...previous?.state, ...incoming.state };
      for (const sessionKey of officeSessionKeys) {
        if (previous && Object.prototype.hasOwnProperty.call(previous.state, sessionKey)) {
          state[sessionKey] = previous.state[sessionKey];
        } else {
          delete state[sessionKey];
        }
      }
      const merged = JSON.stringify({ ...previous, ...incoming, state });
      if (merged !== previousRaw) storage.setItem(key, merged);
    },
    removeItem(key) {
      if (key === editorStorageKey && officeScope()) return;
      storage.removeItem(key);
    },
  };
}

export const editorPersistenceStorage = createEditorPersistenceStorage();
