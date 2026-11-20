import { createHash } from "node:crypto";
import { readFile, stat, writeFile, mkdir } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import process from "node:process";

const root = process.cwd();
const distRoot = resolve(root, "dist");
const indexPath = join(distRoot, "index.html");
const manifestPath = resolve(
  root,
  "src-tauri/target/release/frontend-dist-manifest.json",
);

function referencedAssets(indexHtml) {
  const assets = new Set();
  const expression = /(?:src|href)\s*=\s*["']([^"']+)["']/giu;
  for (const match of indexHtml.matchAll(expression)) {
    const raw = match[1].split(/[?#]/u, 1)[0].trim();
    if (!raw || /^(?:[a-z]+:|data:|#)/iu.test(raw)) continue;
    const relative = raw.replace(/^\/+/, "");
    if (relative.endsWith(".js") || relative.endsWith(".css")) {
      assets.add(relative);
    }
  }
  return [...assets].sort();
}

async function fileRecord(relativePath) {
  const absolutePath = join(distRoot, relativePath);
  const bytes = await readFile(absolutePath).catch((error) => {
    throw new Error(
      `dist/index.html references a missing asset: ${relativePath} (${error.message})`,
    );
  });
  const metadata = await stat(absolutePath);
  if (!metadata.isFile() || bytes.length === 0) {
    throw new Error(`Frontend asset is empty or not a file: ${absolutePath}`);
  }
  return {
    path: relativePath.replaceAll("\\", "/"),
    bytes: bytes.length,
    sha256: createHash("sha256").update(bytes).digest("hex").toUpperCase(),
  };
}

const indexBytes = await readFile(indexPath).catch((error) => {
  throw new Error(
    `Tauri frontend entry is missing before codegen: ${indexPath} (${error.message})`,
  );
});
if (indexBytes.length === 0) {
  throw new Error(`Tauri frontend entry is empty: ${indexPath}`);
}
const indexHtml = indexBytes.toString("utf8");
const assets = referencedAssets(indexHtml);
if (!assets.some((asset) => asset.endsWith(".js"))) {
  throw new Error("dist/index.html does not reference a JavaScript bundle");
}
if (!assets.some((asset) => asset.endsWith(".css"))) {
  throw new Error("dist/index.html does not reference a CSS bundle");
}

const records = await Promise.all(assets.map(fileRecord));
const manifest = {
  schemaVersion: 1,
  generatedAt: new Date().toISOString(),
  distRoot,
  index: {
    path: "index.html",
    bytes: indexBytes.length,
    sha256: createHash("sha256")
      .update(indexBytes)
      .digest("hex")
      .toUpperCase(),
  },
  assets: records,
};
await mkdir(dirname(manifestPath), { recursive: true });
await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
console.log(
  `Verified deterministic Tauri frontend dist: index.html + ${records.length} referenced JS/CSS assets`,
);
console.log(`Frontend manifest: ${manifestPath}`);
