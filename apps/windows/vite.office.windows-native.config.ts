import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import {
  visualTexKernelEntry,
  visualTexKernelHost,
  visualTexKernelSsrEntry,
} from "./vite.visualtexKernelHost";

const root = fileURLToPath(new URL(".", import.meta.url));

const rejectLegacyOfficeJsImports = {
  name: "visualtex-windows-native-office-reject-legacy-officejs",
  enforce: "pre" as const,
  resolveId(source: string) {
    if (
      source.includes("/office/windows-ole/") ||
      source.includes("\\office\\windows-ole\\") ||
      source.includes("@microsoft/office-js")
    ) {
      throw new Error(
        `Windows native Office companion UI cannot import the retired Office.js branch: ${source}`,
      );
    }
    return null;
  },
};

export default defineConfig({
  plugins: [
    visualTexKernelHost(),
    react(),
    rejectLegacyOfficeJsImports,
  ],
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
  base: "/",
  publicDir: false,
  clearScreen: false,
  esbuild: { legalComments: "none" },
  build: {
    outDir: "dist-office-windows-native",
    emptyOutDir: true,
    target: "es2018",
    minify: "esbuild",
    sourcemap: false,
    rollupOptions: {
      input: {
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
