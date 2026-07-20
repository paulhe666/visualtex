import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const root = fileURLToPath(new URL(".", import.meta.url));

const rejectMacImports = {
  name: "visualtex-windows-office-reject-macos-imports",
  enforce: "pre" as const,
  resolveId(source: string) {
    if (source.includes("/office/macos/") || source.includes("\\office\\macos\\")) {
      throw new Error(`Windows OLE Office bundle cannot import macOS code: ${source}`);
    }
    return null;
  },
};

export default defineConfig({
  plugins: [react(), rejectMacImports],
  base: "/",
  publicDir: "office/windows/ole",
  clearScreen: false,
  esbuild: { legalComments: "none" },
  build: {
    outDir: "dist-office-windows-ole",
    emptyOutDir: true,
    target: "es2018",
    minify: "esbuild",
    sourcemap: false,
    rollupOptions: {
      input: {
        bridge: resolve(root, "office-windows-ole-bridge.html"),
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
