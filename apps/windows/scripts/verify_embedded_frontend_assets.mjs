import { readFile, readdir } from "node:fs/promises";
import { join, resolve } from "node:path";
import process from "node:process";

function option(name, fallback) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : fallback;
}

const root = process.cwd();
const distRoot = resolve(root, option("--dist", "dist"));
const executablePath = resolve(
  root,
  option("--exe", "src-tauri/target/release/visualtex.exe"),
);
const indexPath = join(distRoot, "index.html");
const assetRoot = join(distRoot, "assets");

function referencedAssets(indexHtml) {
  const assets = new Set();
  const expression = /(?:src|href)\s*=\s*["']([^"']+)["']/giu;
  for (const match of indexHtml.matchAll(expression)) {
    const relative = match[1]
      .split(/[?#]/u, 1)[0]
      .replace(/^\/+/, "")
      .trim();
    if (relative.endsWith(".js") || relative.endsWith(".css")) {
      assets.add(relative);
    }
  }
  return [...assets].sort();
}

const indexBytes = await readFile(indexPath);
const indexHtml = indexBytes.toString("utf8");
const assets = referencedAssets(indexHtml);
const executable = await readFile(executablePath).catch((error) => {
  throw new Error(`Release executable is missing: ${executablePath} (${error.message})`);
});

const formulaFontAssets = (await readdir(assetRoot))
  .filter((name) => /^KaTeX_.+\.woff2$/u.test(name))
  .sort();
const requiredFormulaFontFamilies = [
  "KaTeX_AMS-Regular-",
  "KaTeX_Main-Regular-",
  "KaTeX_Main-Italic-",
  "KaTeX_Main-Bold-",
  "KaTeX_Main-BoldItalic-",
  "KaTeX_Math-Italic-",
  "KaTeX_Math-BoldItalic-",
  "KaTeX_Size1-Regular-",
  "KaTeX_Size2-Regular-",
  "KaTeX_Size4-Regular-",
];
const missingFormulaFontFamilies = requiredFormulaFontFamilies.filter(
  (prefix) => !formulaFontAssets.some((name) => name.startsWith(prefix)),
);
if (formulaFontAssets.length !== 19 || missingFormulaFontFamilies.length > 0) {
  throw new Error(
    `Expected the 19 bundled MathLive/KaTeX WOFF2 assets; found ${formulaFontAssets.length}. Missing required families: ${missingFormulaFontFamilies.join(", ") || "none"}. Assets: ${formulaFontAssets.join(", ")}`,
  );
}
const requiredMarkers = ["index.html", ...assets, ...formulaFontAssets.map((name) => `assets/${name}`)];
const missingMarkers = requiredMarkers.filter(
  (marker) => executable.indexOf(Buffer.from(marker, "utf8")) < 0,
);
if (missingMarkers.length > 0) {
  throw new Error(
    `The release executable does not contain the generated Tauri asset table entries: ${missingMarkers.join(", ")}`,
  );
}

// Tauri's generated asset table includes the full index payload. Requiring a
// distinctive module/CSS reference from this exact build prevents an old EXE
// with coincidentally similar generic asset names from passing.
const indexReferences = assets.map((asset) => `/${asset}`);
const missingIndexReferences = indexReferences.filter(
  (reference) => executable.indexOf(Buffer.from(reference, "utf8")) < 0,
);
if (missingIndexReferences.length > 0) {
  throw new Error(
    `The release executable does not contain the current dist/index.html references: ${missingIndexReferences.join(", ")}`,
  );
}

console.log(
  `Verified embedded Tauri frontend asset table in ${executablePath}: index.html + ${assets.length} current JS/CSS assets + ${formulaFontAssets.length} bundled formula fonts`,
);
