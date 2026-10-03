import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import {
  mathLiveBrowserEntry,
  visualTexMathLiveKernel,
} from "./vite.mathlive";

const tauriShim = fileURLToPath(new URL("./src/web/tauriShim.ts", import.meta.url));

export default defineConfig({
  plugins: [visualTexMathLiveKernel(), react()],
  base: "/",
  resolve: {
    alias: [
      // Same owned MathLive kernel as the macOS app (vendor/mathlive).
      { find: /^mathlive$/, replacement: mathLiveBrowserEntry },
      { find: /^@tauri-apps\/(api\/core|api\/window|plugin-dialog)$/, replacement: tauriShim },
    ],
  },
  optimizeDeps: {
    exclude: ["mathlive"],
  },
  preview: {
    port: 4173,
    strictPort: true,
  },
  build: {
    target: "es2020",
    minify: "esbuild",
    sourcemap: false,
  },
});
