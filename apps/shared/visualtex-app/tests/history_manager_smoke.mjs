import { build } from "esbuild";
import { rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

const output = join(
  tmpdir(),
  `visualtex-history-smoke-${process.pid}-${Date.now()}.mjs`,
);

try {
  await build({
    entryPoints: ["scripts/history_manager_cases.ts"],
    outfile: output,
    bundle: true,
    platform: "node",
    format: "esm",
    target: "node20",
    sourcemap: "inline",
    logLevel: "silent",
    define: {
      "process.env.NODE_ENV": '"test"',
    },
  });
  await import(`${pathToFileURL(output).href}?run=${Date.now()}`);
} finally {
  await rm(output, { force: true }).catch(() => undefined);
}
