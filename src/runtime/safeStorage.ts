const memoryStorage = new Map<string, string>();

function browserStorage(): Storage | null {
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
        return memoryStorage.get(key) ?? null;
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
