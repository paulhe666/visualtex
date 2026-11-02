import { fileURLToPath } from "node:url";
import { defineConfig, searchForWorkspaceRoot } from "vite";
import react from "@vitejs/plugin-react";
import {
  visualTexKernelEntry,
  visualTexKernelHost,
  visualTexKernelSsrEntry,
} from "./vite.visualtexKernelHost";
import { sharedPackageDedupe } from "./vite.sharedPackages";

const host = process.env.TAURI_DEV_HOST;

export default defineConfig({
  plugins: [visualTexKernelHost(), react()],
  resolve: {
    // Shared code in apps/shared (the custom symbol designer) uses the app's
    // own copies of these packages.
    dedupe: sharedPackageDedupe,
    alias: [
      // the editing kernel (apps/shared/visualtex-core/src/kernel)
      {
        find: /^mathlive$/,
        replacement: visualTexKernelEntry,
      },
      {
        find: /^mathlive\/ssr$/,
        replacement: visualTexKernelSsrEntry,
      },
    ],
  },
  clearScreen: false,
  server: {
    port: 1420,
    strictPort: true,
    host: host || false,
    hmr: host
      ? { protocol: "ws", host, port: 1421 }
      : undefined,
    // The formula layout engine and its fonts live in apps/shared.
    fs: {
      allow: [
        searchForWorkspaceRoot(process.cwd()),
        fileURLToPath(new URL("../shared", import.meta.url)),
      ],
    },
  },
  envPrefix: ["VITE_", "TAURI_ENV_*"],
  build: {
    target: process.env.TAURI_ENV_PLATFORM === "windows" ? "chrome105" : "safari13",
    minify: process.env.TAURI_ENV_DEBUG ? false : "esbuild",
    sourcemap: !!process.env.TAURI_ENV_DEBUG,
    rollupOptions: {
      input: {
        main: "index.html",
        // VisualTeX Anywhere's formula overlay (src/anywhere)
        anywhere: "anywhere.html",
      },
    },
  },
});
