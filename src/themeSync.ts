import type { Theme } from "./types/formula";
import {
  applyThemePalette,
  isTheme,
  subscribeCustomTheme,
} from "./themeCustomization";

const ACTIVE_THEME_STORAGE_KEY = "visualtex.active-theme";
const EDITOR_STORAGE_KEY = "visualtex-editor";
const THEME_CHANNEL_NAME = "visualtex-theme";

export function normalizeSynchronizedTheme(value: unknown): Theme {
  return isTheme(value) ? value : "light";
}

function parsePersistedEditorTheme(raw: string | null): Theme | null {
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw) as { state?: { theme?: unknown } };
    return normalizeSynchronizedTheme(parsed.state?.theme);
  } catch {
    return null;
  }
}

function readPersistedEditorTheme(): Theme | null {
  try {
    return parsePersistedEditorTheme(window.localStorage.getItem(EDITOR_STORAGE_KEY));
  } catch {
    return null;
  }
}

export function readPublishedSynchronizedTheme(): Theme {
  try {
    const activeTheme = window.localStorage.getItem(ACTIVE_THEME_STORAGE_KEY);
    if (activeTheme) return normalizeSynchronizedTheme(activeTheme);
  } catch {
    // Fall through to the persisted editor store.
  }
  return readPersistedEditorTheme() ?? "light";
}

export function readSynchronizedTheme(): Theme {
  const urlTheme = new URLSearchParams(window.location.search).get("theme");
  if (urlTheme) return normalizeSynchronizedTheme(urlTheme);
  return readPublishedSynchronizedTheme();
}

export function applyDocumentTheme(theme: Theme) {
  const normalized = normalizeSynchronizedTheme(theme);
  document.documentElement.dataset.theme = normalized;
  applyThemePalette(normalized);
}

export function publishSynchronizedTheme(theme: Theme) {
  const normalized = normalizeSynchronizedTheme(theme);
  applyDocumentTheme(normalized);
  try {
    window.localStorage.setItem(ACTIVE_THEME_STORAGE_KEY, normalized);
  } catch {
    // Theme application still succeeds when storage is unavailable.
  }
  if (typeof BroadcastChannel === "undefined") return;
  const channel = new BroadcastChannel(THEME_CHANNEL_NAME);
  channel.postMessage(normalized);
  channel.close();
}

export function subscribeSynchronizedTheme(
  listener: (theme: Theme) => void,
): () => void {
  const unsubscribeCustomization = subscribeCustomTheme();
  const handleStorage = (event: StorageEvent) => {
    if (event.key === ACTIVE_THEME_STORAGE_KEY && event.newValue) {
      listener(normalizeSynchronizedTheme(event.newValue));
      return;
    }
    if (event.key === EDITOR_STORAGE_KEY) {
      const nextTheme = parsePersistedEditorTheme(event.newValue);
      if (nextTheme) listener(nextTheme);
    }
  };
  window.addEventListener("storage", handleStorage);

  const channel =
    typeof BroadcastChannel === "undefined"
      ? null
      : new BroadcastChannel(THEME_CHANNEL_NAME);
  if (channel) {
    channel.onmessage = (event) => listener(normalizeSynchronizedTheme(event.data));
  }

  return () => {
    window.removeEventListener("storage", handleStorage);
    channel?.close();
    unsubscribeCustomization();
  };
}
