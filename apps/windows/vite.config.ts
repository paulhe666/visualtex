import { fileURLToPath } from "node:url";
import { defineConfig, searchForWorkspaceRoot } from "vite";
import react from "@vitejs/plugin-react";
import {
  visualTexKernelEntry,
  visualTexKernelHost,
  visualTexKernelSsrEntry,
} from "./vite.visualtexKernelHost";

const host = process.env.TAURI_DEV_HOST;

export default defineConfig({
  plugins: [visualTexKernelHost(), react()],
  resolve: {
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
  },
});
