import { fileURLToPath } from "node:url";
import { defineConfig, searchForWorkspaceRoot } from "vite";
import react from "@vitejs/plugin-react";
import {
  visualTexKernelEntry,
  visualTexKernelHost,
  visualTexKernelSsrEntry,
} from "../shared/visualtex-app/vite/visualtexKernelHost";
import { sharedPackageDedupe } from "../shared/visualtex-app/vite/sharedPackages";

const tauriShim = fileURLToPath(new URL("./src/web/tauriShim.ts", import.meta.url));

// The web edition: the shared app (apps/shared) with this app's platform modules (src/platform), in a browser. The
// Tauri packages the shared code imports are src/web/tauriShim.ts here. The build goes to the repository's dist/,
// which Cloudflare serves with the OCR relay (worker/index.js; wrangler.jsonc at the repository root).
export default defineConfig({
  plugins: [visualTexKernelHost(), react()],
  base: "/",
  resolve: {
    dedupe: sharedPackageDedupe,
    alias: [
      // what shared code asks of this app (src/platform, imported as "@platform/<module>")
      { find: /^@platform\//, replacement: fileURLToPath(new URL("./src/platform/", import.meta.url)) },
      // the editing kernel (apps/shared/visualtex-core/src/kernel)
      { find: /^mathlive$/, replacement: visualTexKernelEntry },
      { find: /^mathlive\/ssr$/, replacement: visualTexKernelSsrEntry },
      {
        find: /^@tauri-apps\/(api\/core|api\/event|api\/window|api\/webview|api\/dpi|plugin-dialog|plugin-opener)$/,
        replacement: tauriShim,
      },
    ],
  },
  server: {
    fs: {
      allow: [searchForWorkspaceRoot(process.cwd()), fileURLToPath(new URL("../shared", import.meta.url))],
    },
  },
  preview: {
    port: 4173,
    strictPort: true,
  },
  build: {
    outDir: fileURLToPath(new URL("../../dist", import.meta.url)),
    emptyOutDir: true,
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
});
