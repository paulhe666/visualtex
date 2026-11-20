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

// Formulas are drawn by the shared layout engine in its own fonts: Computer
// Modern/AMS (data/fonts) and each font profile's (data/profiles/<name>/fonts,
// e.g. newtx for "Times New Roman"). Every face their fonts.css declare must be bundled.
const layoutData = new URL("../../shared/tex-layout/data/", import.meta.url);
const profileNames = await readdir(new URL("profiles/", layoutData)).catch(() => []);
const formulaFontFaces = [
  new URL("fonts/fonts.css", layoutData),
  ...profileNames.map((name) => new URL(`profiles/${name}/fonts/fonts.css`, layoutData)),
];
const requiredFormulaFonts = [];
for (const faces of formulaFontFaces) {
  for (const match of (await readFile(faces, "utf8")).matchAll(/url\(["']?\.\/([^"')]+)\.otf["']?\)/gu)) {
    requiredFormulaFonts.push(`${match[1]}-`);
  }
}
const formulaFontAssets = (await readdir(assetRoot))
  .filter((name) => /^vtx-.+\.otf$/u.test(name))
  .sort();
const missingFormulaFonts = requiredFormulaFonts.filter(
  (prefix) => !formulaFontAssets.some((name) => name.startsWith(prefix)),
);
if (
  requiredFormulaFonts.length === 0 ||
  formulaFontAssets.length !== requiredFormulaFonts.length ||
  missingFormulaFonts.length > 0
) {
  throw new Error(
    `Expected the ${requiredFormulaFonts.length} layout engine OTF fonts declared in apps/shared/tex-layout/data/fonts/fonts.css and data/profiles/*/fonts/fonts.css; found ${formulaFontAssets.length}. Missing: ${missingFormulaFonts.join(", ") || "none"}. Assets: ${formulaFontAssets.join(", ")}`,
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
