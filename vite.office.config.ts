import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const root = fileURLToPath(new URL(".", import.meta.url));

const stripRemoteFallbackMarkers = {
  name: "visualtex-office-strip-remote-fallback-markers",
  enforce: "post" as const,
  renderChunk(code: string) {
    const sanitized = code.replaceAll(
      "jsdelivr.net/",
      "visualtex-local-vendor.invalid/",
    );
    return sanitized === code ? null : { code: sanitized, map: null };
  },
};

export default defineConfig({
  plugins: [react(), stripRemoteFallbackMarkers],
  base: "/",
  publicDir: "office",
  clearScreen: false,
  esbuild: {
    legalComments: "none",
  },
  build: {
    outDir: "dist-office",
    emptyOutDir: true,
    target: "safari13",
    minify: "esbuild",
    sourcemap: false,
    rollupOptions: {
      input: {
        bridge: resolve(root, "office-bridge.html"),
        dialog: resolve(root, "office-dialog.html"),
      },
      output: {
        entryFileNames: "assets/[name]-[hash].js",
        chunkFileNames: "assets/[name]-[hash].js",
        assetFileNames: "assets/[name]-[hash][extname]",
      },
    },
  },
});
