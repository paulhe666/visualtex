import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const root = fileURLToPath(new URL(".", import.meta.url));

const rejectWindowsImports = {
  name: "visualtex-macos-office-reject-windows-imports",
  enforce: "pre" as const,
  resolveId(source: string) {
    if (
      source.includes("/office/windows-ole/") ||
      source.includes("\\office\\windows-ole\\")
    ) {
      throw new Error(`macOS Office bundle cannot import Windows OLE code: ${source}`);
    }
    return null;
  },
};

const stripRemoteFallbackMarkers = {
  name: "visualtex-office-strip-remote-fallback-markers",
  enforce: "post" as const,
  renderChunk(code: string) {
    const sanitized = code
      .replaceAll("jsdelivr.net/", "visualtex-local-vendor.invalid/")
      .replaceAll(
        "https://esm.run/@cortex-js/compute-engine",
        "visualtex-local-vendor.invalid/compute-engine",
      );
    return sanitized === code ? null : { code: sanitized, map: null };
  },
};

export default defineConfig({
  plugins: [react(), rejectWindowsImports, stripRemoteFallbackMarkers],
  base: "/",
  publicDir: "office/macos",
  clearScreen: false,
  esbuild: { legalComments: "none" },
  build: {
    outDir: "dist-office-macos",
    emptyOutDir: true,
    target: "safari13",
    minify: "esbuild",
    sourcemap: false,
    rollupOptions: {
      input: {
        bridge: resolve(root, "office-macos-bridge.html"),
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
