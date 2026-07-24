/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_RELEASE_UI_PROBE?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
