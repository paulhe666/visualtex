// Environment flags the frontend reads (the apps reference vite/client in their own src/vite-env.d.ts).
interface ImportMetaEnv {
  readonly VITE_RELEASE_UI_PROBE?: string;
  readonly VITE_VISUALTEX_IME_DIAGNOSTICS?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
