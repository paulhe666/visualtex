import { copyFile, mkdir, rm } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const platform = process.argv[2];
if (platform !== "windows-ole") {
  throw new Error("Usage: node finalize_office_platform_build.mjs windows-ole");
}

const dist = join(root, "dist-office-windows-ole");
for (const [sourceName, destination] of [
  ["office-windows-ole-bridge.html", join("bridge", "index.html")],
  ["office-dialog.html", join("dialog", "index.html")],
]) {
  const source = join(dist, sourceName);
  const target = join(dist, destination);
  await mkdir(dirname(target), { recursive: true });
  await copyFile(source, target);
  await rm(source);
}
console.log("Finalized Windows Office compatibility build layout.");
