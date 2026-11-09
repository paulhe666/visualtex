/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_RELEASE_UI_PROBE?: string;
  readonly VITE_VISUALTEX_IME_DIAGNOSTICS?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
