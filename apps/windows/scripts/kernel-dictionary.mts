// The editing kernel's command dictionary (LATEX_COMMANDS, MATH_SYMBOLS,
// ENVIRONMENTS and its macro tables) for Node scripts, bundled with this app's
// host module the way scripts/kernel-node.mts bundles the kernel.
//   import { kernelDictionary } from "./kernel-dictionary.mts";
import { build } from "esbuild";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { kernelHostModuleSource } from "../vite.visualtexKernelHost.ts";

const appDir = fileURLToPath(new URL("..", import.meta.url));
const outDir = path.join(appDir, "node_modules/.cache/visualtex-kernel-dictionary");
fs.mkdirSync(outDir, { recursive: true });
const hostFile = path.join(outDir, "host.mjs");
fs.writeFileSync(hostFile, kernelHostModuleSource());
const outFile = path.join(outDir, "definitions.mjs");
await build({
  entryPoints: [path.join(appDir, "../shared/visualtex-core/src/kernel/latex-commands/definitions.ts")],
  outfile: outFile,
  bundle: true,
  platform: "node",
  format: "esm",
  logLevel: "warning",
  loader: { ".css": "empty" },
  plugins: [{
    name: "visualtex-kernel-dictionary",
    setup(b) {
      b.onResolve({ filter: /^virtual:visualtex-kernel-host$/ }, () => ({ path: hostFile }));
      b.onResolve({ filter: /\?raw$/ }, (args) => ({
        path: path.resolve(args.resolveDir, args.path.replace(/\?raw$/, "")),
        namespace: "raw",
      }));
      b.onLoad({ filter: /.*/, namespace: "raw" }, (args) => ({
        contents: fs.readFileSync(args.path, "utf8"),
        loader: "text",
      }));
    },
  }],
});

export const kernelDictionary = await import(pathToFileURL(outFile).href);
