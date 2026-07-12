import { cp, mkdir, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const packageRoot = join(root, "node_modules", "@microsoft", "office-js");
const mathJaxRoot = join(root, "node_modules", "mathjax-full");
const source = join(packageRoot, "dist");
const target = join(root, "office", "vendor", "office-js");
const licenses = join(root, "office", "licenses");

const rootFiles = [
  "office.js",
  "word-mac-16.00.js",
  "powerpoint-mac-16.00.js",
  "o15apptofilemappingtable.js",
  "es6-promise.js",
];

function sanitizeOfficeSource(value) {
  return value
    .replaceAll(
      "https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE",
      "ES6 Promise license is included with the vendored Office.js package",
    )
    .replaceAll(
      "http://go.microsoft.com/fwlink/?LinkId=266419.",
      "Office.js offline documentation is bundled with VisualTeX.",
    );
}

async function copySanitized(relativePath) {
  const input = await readFile(join(source, relativePath), "utf8");
  const output = sanitizeOfficeSource(input);
  const destination = join(target, relativePath);
  await mkdir(dirname(destination), { recursive: true });
  await writeFile(destination, output, "utf8");
}

await rm(target, { recursive: true, force: true });
await mkdir(target, { recursive: true });
await mkdir(licenses, { recursive: true });

for (const file of rootFiles) await copySanitized(file);

const entries = await readdir(source, { withFileTypes: true });
for (const entry of entries) {
  if (!entry.isDirectory()) continue;
  const localeFile = join(source, entry.name, "office_strings.js");
  try {
    await readFile(localeFile);
  } catch {
    continue;
  }
  await copySanitized(join(entry.name, "office_strings.js"));
}

await cp(
  join(packageRoot, "LICENSE.md"),
  join(licenses, "OFFICE-JS-LICENSE.md"),
);
await cp(
  join(mathJaxRoot, "LICENSE"),
  join(licenses, "MATHJAX-LICENSE.txt"),
);

const packageJson = JSON.parse(
  await readFile(join(packageRoot, "package.json"), "utf8"),
);
const officeJs = await readFile(join(target, "office.js"));
const manifest = {
  package: packageJson.name,
  version: packageJson.version,
  officeJsSha256: createHash("sha256").update(officeJs).digest("hex"),
  scope: "macOS Word and PowerPoint plus all Office UI locale strings",
};
await writeFile(
  join(target, "VERSION.json"),
  `${JSON.stringify(manifest, null, 2)}\n`,
  "utf8",
);

console.log(
  `Prepared Office.js ${manifest.version} for offline macOS Word/PowerPoint use.`,
);
