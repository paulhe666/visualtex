// The homepage showcase is a separate, ephemeral editor session.
export const isLandingPreview =
  typeof window !== "undefined" &&
  /^\/editor\/?$/.test(window.location.pathname) &&
  new URLSearchParams(window.location.search).has("landing-preview");

export const LANDING_PREVIEW_ZOOM = 0.5;

// The preview is embedded below the fold of the landing page. Scrolling or focusing inside it
// would otherwise scroll the host page down to the iframe on load.
if (isLandingPreview) {
  Element.prototype.scrollIntoView = function scrollIntoView() {};
  const focus = HTMLElement.prototype.focus;
  HTMLElement.prototype.focus = function focusWithoutScroll(options?: FocusOptions) {
    focus.call(this, { ...options, preventScroll: true });
  };
}
