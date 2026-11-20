// The homepage showcase and the tutorial's practice editor are separate,
// ephemeral editor sessions.
const editorQuery =
  typeof window !== "undefined" && /^\/editor\/?$/.test(window.location.pathname)
    ? new URLSearchParams(window.location.search)
    : null;

export const isLandingPreview = Boolean(editorQuery?.has("landing-preview"));

/** Lesson id when /editor runs inside the tutorial page (`/editor?tutorial=<id>`). */
export const tutorialLessonId = editorQuery?.get("tutorial") ?? null;
export const tutorialLanguage = editorQuery?.get("lang") === "en" ? "en" : "cn";

const isEditorSandbox = isLandingPreview || tutorialLessonId !== null;

export const LANDING_PREVIEW_ZOOM = 0.5;

class MemoryStorage implements Storage {
  private readonly values = new Map<string, string>();
  get length() {
    return this.values.size;
  }
  clear() {
    this.values.clear();
  }
  getItem(key: string) {
    return this.values.get(String(key)) ?? null;
  }
  key(index: number) {
    return Array.from(this.values.keys())[index] ?? null;
  }
  removeItem(key: string) {
    this.values.delete(String(key));
  }
  setItem(key: string, value: string) {
    this.values.set(String(key), String(value));
  }
}

if (isEditorSandbox) {
  // The sandbox shares an origin with /editor. The synced macOS editor reads
  // localStorage and IndexedDB directly, so replace both before any editor
  // module loads; the preview must never read or overwrite user documents.
  Object.defineProperty(window, "localStorage", {
    configurable: true,
    value: new MemoryStorage(),
  });
  Object.defineProperty(window, "indexedDB", {
    configurable: true,
    value: undefined,
  });

  // The sandbox is embedded in a longer page. Scrolling or focusing inside it
  // would otherwise scroll the host page to the iframe.
  Element.prototype.scrollIntoView = function scrollIntoView() {};
  const focus = HTMLElement.prototype.focus;
  HTMLElement.prototype.focus = function focusWithoutScroll(options?: FocusOptions) {
    focus.call(this, { ...options, preventScroll: true });
  };
}
