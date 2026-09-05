import { isLandingPreview } from "./landingPreview";

const memoryStorage = new Map<string, string>();

function browserStorage(): Storage | null {
  // Showcase reads and writes stay in this frame; never touch user documents.
  if (isLandingPreview) return null;
  try {
    return typeof window === "undefined" ? null : window.localStorage;
  } catch {
    return null;
  }
}

export const safeStorage = {
  getItem(key: string): string | null {
    const storage = browserStorage();
    if (storage) {
      try {
        const value = storage.getItem(key);
        if (value !== null) {
          memoryStorage.set(key, value);
          return value;
        }
        memoryStorage.delete(key);
        return null;
      } catch {
        // Fall back to memory for this session.
      }
    }
    return memoryStorage.get(key) ?? null;
  },

  setItem(key: string, value: string): void {
    memoryStorage.set(key, value);
    const storage = browserStorage();
    if (!storage) return;
    try {
      storage.setItem(key, value);
    } catch {
      // Keep the in-memory preference when WebView storage is unavailable.
    }
  },

  /**
   * Persist a value without swallowing a real browser storage failure. This is
   * used for user-authored assets where silently falling back to process memory
   * would make data appear saved until the application restarts.
   */
  setItemStrict(key: string, value: string): void {
    const storage = browserStorage();
    if (!storage) {
      memoryStorage.set(key, value);
      return;
    }
    try {
      storage.setItem(key, value);
      memoryStorage.set(key, value);
    } catch (error) {
      throw new Error(
        `VisualTeX could not persist ${key}: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
    }
  },

  removeItem(key: string): void {
    memoryStorage.delete(key);
    const storage = browserStorage();
    if (!storage) return;
    try {
      storage.removeItem(key);
    } catch {
      // The in-memory fallback has already been cleared.
    }
  },
};

export function readLocalStorage(key: string): string | null {
  return safeStorage.getItem(key);
}

export function writeLocalStorage(key: string, value: string): void {
  safeStorage.setItem(key, value);
}
