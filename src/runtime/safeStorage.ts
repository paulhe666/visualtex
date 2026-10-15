// Only failed preference writes live here; successful storage remains authoritative.
function browserStorage(): Storage | null {
  return typeof window === "undefined" ? null : window.localStorage;
}

export function createSafeStorage(getStorage: () => Storage | null = browserStorage) {
  const pending = new Map<string, string | null>();
  const requiredStorage = () => {
    const storage = getStorage();
    if (!storage) throw new Error("Persistent storage is unavailable.");
    return storage;
  };
  return {
    getItem(key: string): string | null {
      if (pending.has(key)) return pending.get(key) ?? null;
      try { return getStorage()?.getItem(key) ?? null; }
      catch { return null; }
    },
    setItem(key: string, value: string): void {
      try {
        requiredStorage().setItem(key, value);
        pending.delete(key);
      } catch {
        pending.set(key, value);
      }
    },
    removeItem(key: string): void {
      try {
        requiredStorage().removeItem(key);
        pending.delete(key);
      } catch {
        pending.set(key, null);
      }
    },
    // User-authored content must succeed on the backing store before reporting success.
    setItemStrict(key: string, value: string): void {
      requiredStorage().setItem(key, value);
      pending.delete(key);
    },
    removeItemStrict(key: string): void {
      requiredStorage().removeItem(key);
      pending.delete(key);
    },
  };
}

export const safeStorage = createSafeStorage();

export function readLocalStorage(key: string): string | null {
  return safeStorage.getItem(key);
}
export function writeLocalStorage(key: string, value: string): void {
  safeStorage.setItem(key, value);
}
