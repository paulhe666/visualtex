import { copyFile, mkdir, rm } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const dist = join(root, "dist-office");

for (const [sourceName, destination] of [
  ["office-bridge.html", join("bridge", "index.html")],
  ["office-dialog.html", join("dialog", "index.html")],
]) {
  const source = join(dist, sourceName);
  const target = join(dist, destination);
  await mkdir(dirname(target), { recursive: true });
  await copyFile(source, target);
  await rm(source);
}

console.log("Finalized Office build layout.");
