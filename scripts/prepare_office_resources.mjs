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
const icons = join(root, "office", "icons");
const manifests = join(root, "office", "manifests");
const companionOrigin = "https://127.0.0.1:43127";
const wordAddinId = "d6fcb260-4c37-4f73-a173-cf24674f81f2";
const powerpointAddinId = "a6d13cf2-54e8-4dfa-a20c-15de864ab3c5";

const rootFiles = [
  "office.js",
  "word-mac-16.00.js",
  "powerpoint-mac-16.00.js",
  "o15apptofilemappingtable.js",
  "es6-promise.js",
];

function fourPartVersion(value) {
  const parts = String(value)
    .split(".")
    .map((part) => part.match(/^\d+/)?.[0])
    .filter(Boolean)
    .slice(0, 4);
  while (parts.length < 4) parts.push("0");
  return parts.join(".");
}

function renderManifest(template, appVersion) {
  const rendered = template
    .replaceAll("{{WORD_ADDIN_ID}}", wordAddinId)
    .replaceAll("{{POWERPOINT_ADDIN_ID}}", powerpointAddinId)
    .replaceAll("{{MANIFEST_VERSION}}", fourPartVersion(appVersion))
    .replaceAll("{{COMPANION_ORIGIN}}", companionOrigin);
  if (rendered.includes("{{") || rendered.includes("}}")) {
    throw new Error("Office manifest still contains unresolved placeholders");
  }
  const forbidden = [
    "appsforoffice.microsoft.com",
    "github.com",
    "githubusercontent.com",
    "googleapis.com",
    "gstatic.com",
    "unpkg.com",
    "jsdelivr.net",
    "cloudflare.com",
  ];
  const remote = forbidden.find((domain) => rendered.includes(domain));
  if (remote) throw new Error(`Office manifest contains forbidden domain: ${remote}`);
  return rendered;
}

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
await mkdir(icons, { recursive: true });
await mkdir(manifests, { recursive: true });

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
await cp(
  join(root, "src-tauri", "icons", "32x32.png"),
  join(icons, "icon-16.png"),
);
await cp(
  join(root, "src-tauri", "icons", "32x32.png"),
  join(icons, "icon-32.png"),
);
await cp(
  join(root, "src-tauri", "icons", "128x128.png"),
  join(icons, "icon-80.png"),
);

const appPackageJson = JSON.parse(
  await readFile(join(root, "package.json"), "utf8"),
);
for (const [templateName, outputName] of [
  ["visualtex-word.template.xml", "VisualTeX.Word.xml"],
  ["visualtex-powerpoint.template.xml", "VisualTeX.PowerPoint.xml"],
]) {
  const template = await readFile(join(manifests, templateName), "utf8");
  await writeFile(
    join(manifests, outputName),
    renderManifest(template, appPackageJson.version),
    "utf8",
  );
}

const officePackageJson = JSON.parse(
  await readFile(join(packageRoot, "package.json"), "utf8"),
);
const officeJs = await readFile(join(target, "office.js"));
const manifest = {
  package: officePackageJson.name,
  version: officePackageJson.version,
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
