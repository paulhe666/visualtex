// The editing kernel for Node scripts (tests, probes):
//   tsx --import ./scripts/kernel-node.mts scripts/<script>.mts
// builds the kernel's DOM-free part (parsing, validation, markup by the
// VisualTeX layout engine) with this app's host module, and makes `mathlive`
// and `mathlive/ssr` resolve to that build. The browser element is not part
// of it.
import { build } from "esbuild";
import fs from "node:fs";
import path from "node:path";
import { register } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";
import { kernelHostModuleSource, visualTexKernelSsrEntry } from "../vite.visualtexKernelHost.ts";

const appDir = fileURLToPath(new URL("..", import.meta.url));
const outDir = path.join(appDir, "node_modules/.cache/visualtex-kernel-node-macos");
fs.mkdirSync(outDir, { recursive: true });

const hostFile = path.join(outDir, "host.mjs");
fs.writeFileSync(hostFile, kernelHostModuleSource());

// what `mathlive` means in the build: the DOM-free entry and the registration hooks
const kernelFile = path.join(outDir, "kernel.ts");
const hooksFile = path.join(path.dirname(visualTexKernelSsrEntry), "../editor-mathfield/visualtex-hooks.ts");
fs.writeFileSync(
  kernelFile,
  `export * from ${JSON.stringify(visualTexKernelSsrEntry)};\n` +
    `export { setVisualTexLayoutEngine, setVisualTexEditingCore } from ${JSON.stringify(hooksFile)};\n`,
);

// the build registers the app's layout engine; without a canvas, characters the
// TeX fonts lack are sized by a fixed estimate (0.5em wide, 0.7em/0.2em high)
const entryFile = path.join(outDir, "entry.ts");
fs.writeFileSync(
  entryFile,
  `import { createAppLayoutEngine } from ${JSON.stringify(path.join(appDir, "src/editor/visualTexLayoutRuntime.ts"))};\n` +
    `import { setVisualTexLayoutEngine } from ${JSON.stringify(kernelFile)};\n` +
    `setVisualTexLayoutEngine(createAppLayoutEngine((_family, text) => ` +
    `({ width: 0.5 * Array.from(text).length, ascent: 0.7, descent: 0.2 })));\n` +
    `export * from ${JSON.stringify(kernelFile)};\n`,
);

const outFile = path.join(outDir, "kernel.mjs");
await build({
  entryPoints: [entryFile],
  outfile: outFile,
  bundle: true,
  platform: "node",
  format: "esm",
  logLevel: "warning",
  alias: { mathlive: kernelFile },
  loader: { ".css": "empty" },
  plugins: [
    {
      name: "visualtex-kernel",
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
    },
  ],
});

register("./kernel-node-hooks.mjs", import.meta.url, {
  data: { kernelUrl: pathToFileURL(outFile).href },
});
