import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import {
  visualTexKernelEntry,
  visualTexKernelHost,
  visualTexKernelSsrEntry,
} from "./vite.visualtexKernelHost";
import { sharedPackageDedupe } from "./vite.sharedPackages";

const tauriShim = fileURLToPath(new URL("./src/web/tauriShim.ts", import.meta.url));

// The macOS app's build (apps/macos/vite.config.ts) for the browser: the same editing kernel and shared packages,
// the Tauri packages replaced by src/web/tauriShim.ts.
export default defineConfig(({ command }) => {
  // The web version is the public edition (src/edition.ts: document mode Markdown only), whatever the environment
  // says; only a dev server (vite) can run the personal edition, for testing.
  if (command === "build") {
    if (process.env.VITE_VISUALTEX_EDITION && process.env.VITE_VISUALTEX_EDITION !== "public") {
      console.warn(`VITE_VISUALTEX_EDITION=${process.env.VITE_VISUALTEX_EDITION} ignored: the web build is public`);
    }
    process.env.VITE_VISUALTEX_EDITION = "public";
  }
  return {
    plugins: [visualTexKernelHost(), react()],
    base: "/",
    resolve: {
      dedupe: sharedPackageDedupe,
      alias: [
        // the editing kernel (shared/visualtex-core/src/kernel)
        { find: /^mathlive$/, replacement: visualTexKernelEntry },
        { find: /^mathlive\/ssr$/, replacement: visualTexKernelSsrEntry },
        {
          find: /^@tauri-apps\/(api\/core|api\/event|api\/window|api\/webview|plugin-dialog|plugin-opener)$/,
          replacement: tauriShim,
        },
      ],
    },
    preview: {
      port: 4173,
      strictPort: true,
    },
    build: {
      target: "es2020",
      minify: "esbuild",
      sourcemap: false,
      rollupOptions: {
        input: {
          main: "index.html",
          // the document's PDF: its print page (src/document/print-main.ts)
          print: "print.html",
        },
      },
    },
  };
});
