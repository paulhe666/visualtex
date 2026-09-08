// The homepage showcase is a separate, ephemeral editor session.
export const isLandingPreview =
  typeof window !== "undefined" &&
  /^\/editor\/?$/.test(window.location.pathname) &&
  new URLSearchParams(window.location.search).has("landing-preview");

export const LANDING_PREVIEW_ZOOM = 0.5;
