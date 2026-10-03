import type { PersistStorage, StorageValue } from "zustand/middleware";
import { safeStorage } from "./safeStorage";

interface EditorStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

type EditorEnvelope = StorageValue<Record<string, unknown>>;
const editorStorageKey = "visualtex-editor";
const officeSessionKeys = new Set([
  "title", "lines", "activeLineId", "formulaAlignment", "latexCodeFormat", "history",
]);

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

// Operate on Zustand's values before JSON encoding. Office persists preferences
// only; moving the caret in either host must not serialize a whole document.
export function createEditorPersistenceStorage(
  storage: EditorStorage = safeStorage,
  officeScope: () => boolean = isOfficeEditorPersistenceScope,
): PersistStorage<Record<string, unknown>> {
  let previousInput: EditorEnvelope | null = null;
  const persistentKeys = (state: EditorEnvelope["state"], office: boolean) =>
    Object.keys(state).filter(name => office
      ? !officeSessionKeys.has(name)
      : name !== "activeLineId");
  return {
    getItem(key) {
      const envelope = parseEnvelope(storage.getItem(key));
      if (!envelope || key !== editorStorageKey || !officeScope()) return envelope;
      const state = Object.fromEntries(Object.entries(envelope.state)
        .filter(([name]) => !officeSessionKeys.has(name)));
      return { ...envelope, state };
    },
    setItem(key, incoming) {
      const office = key === editorStorageKey && officeScope();
      const savedKeys = persistentKeys(incoming.state, office);
      if (key === editorStorageKey && previousInput &&
          incoming.version === previousInput.version &&
          savedKeys.length === persistentKeys(previousInput.state, office).length &&
          savedKeys.every(name => Object.is(incoming.state[name], previousInput!.state[name]))) return;

      if (!office) {
        storage.setItem(key, JSON.stringify(incoming));
      } else {
        const previousRaw = storage.getItem(key);
        const previous = parseEnvelope(previousRaw);
        // Keep corrupt main-window data available for recovery.
        if (previousRaw !== null && !previous) return;
        const preferences = Object.fromEntries(savedKeys.map(name => [name, incoming.state[name]]));
        const merged = JSON.stringify({ ...previous, ...incoming, state: { ...previous?.state, ...preferences } });
        if (merged !== previousRaw) storage.setItem(key, merged);
      }
      if (key === editorStorageKey) previousInput = incoming;
    },
    removeItem(key) {
      if (key === editorStorageKey && officeScope()) return;
      if (key === editorStorageKey) previousInput = null;
      storage.removeItem(key);
    },
  };
}

export const editorPersistenceStorage = createEditorPersistenceStorage();
