// The homepage showcase is a separate, ephemeral editor session.
export const isLandingPreview =
  typeof window !== "undefined" &&
  /^\/editor\/?$/.test(window.location.pathname) &&
  new URLSearchParams(window.location.search).has("landing-preview");

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

if (isLandingPreview) {
  // The showcase shares an origin with /editor. The synced macOS editor reads
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

  // The preview is embedded below the fold of the landing page. Scrolling or
  // focusing inside it would otherwise scroll the host page down to the iframe.
  Element.prototype.scrollIntoView = function scrollIntoView() {};
  const focus = HTMLElement.prototype.focus;
  HTMLElement.prototype.focus = function focusWithoutScroll(options?: FocusOptions) {
    focus.call(this, { ...options, preventScroll: true });
  };
}
