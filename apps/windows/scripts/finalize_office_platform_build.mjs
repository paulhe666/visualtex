import { copyFile, mkdir, rm } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const platform = process.argv[2];
if (platform !== "windows-native") {
  throw new Error("Usage: node finalize_office_platform_build.mjs windows-native");
}

const dist = join(root, "dist-office-windows-native");
const source = join(dist, "office-dialog.html");
const target = join(dist, "dialog", "index.html");
await mkdir(dirname(target), { recursive: true });
await copyFile(source, target);
await rm(source);
console.log("Finalized Windows native Office companion UI layout.");
