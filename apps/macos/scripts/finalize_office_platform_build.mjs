import { copyFile, mkdir, rm } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const platform = process.argv[2];
const config = {
  macos: {
    dist: "dist-office-macos",
    bridge: "office-macos-bridge.html",
  },
  "windows-ole": {
    dist: "dist-office-windows-ole",
    bridge: "office-windows-ole-bridge.html",
  },
}[platform];
if (!config) {
  throw new Error("Usage: node finalize_office_platform_build.mjs <macos|windows-ole>");
}

const dist = join(root, config.dist);
for (const [sourceName, destination] of [
  [config.bridge, join("bridge", "index.html")],
  ["office-dialog.html", join("dialog", "index.html")],
]) {
  const source = join(dist, sourceName);
  const target = join(dist, destination);
  await mkdir(dirname(target), { recursive: true });
  await copyFile(source, target);
  await rm(source);
}
console.log(`Finalized independent ${platform} Office build layout.`);
