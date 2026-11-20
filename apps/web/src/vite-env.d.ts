/// <reference types="vite/client" />

interface ImportMetaEnv {
  /** browser regressions open the editor at "/" instead of the landing page */
  readonly VITE_VISUALTEX_EDITOR_AT_ROOT?: string;
}

// The Office editor's install token (the desktop apps' Office companion page); the shared OCR service reads it to
// tell the Office editor apart. Never set in a browser.
interface Window {
  __VISUALTEX_INSTALL_TOKEN__?: string;
}
