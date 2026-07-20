import { cp, mkdir, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const platform = process.argv[2];
const configs = {
  "windows-ole": {
    publicRoot: join(root, "office", "windows", "ole"),
    wordId: "7c7d3b35-56b2-4c40-88d9-c9eb836d6021",
    powerpointId: "fdc8d615-7e60-4586-bff4-5a1d728f9f6c",
    rootFiles: [
      "office.js",
      "word-win32-16.00.js",
      "word-win32-16.01.js",
      "powerpoint-win32-16.00.js",
      "powerpoint-win32-16.01.js",
      "o15apptofilemappingtable.js",
      "es6-promise.js",
    ],
    scope: "Windows Word and PowerPoint Office.js OLE ribbon integration",
    // Keep the product version while allowing Office to refresh command
    // manifests independently of a full desktop release.
    manifestRevision: 3,
  },
};

const config = configs[platform];
if (!config) {
  throw new Error("Usage: node prepare_office_platform_resources.mjs windows-ole");
}

const manifestRevision = config.manifestRevision;
if (!Number.isInteger(manifestRevision) || manifestRevision < 0) {
  throw new Error(`${platform} Office manifest revision is invalid`);
}

const packageRoot = join(root, "node_modules", "@microsoft", "office-js");
const mathJaxRoot = join(root, "node_modules", "mathjax-full");
const source = join(packageRoot, "dist");
const target = join(config.publicRoot, "vendor", "office-js");
const licenses = join(config.publicRoot, "licenses");
const icons = join(config.publicRoot, "icons");
const manifests = join(config.publicRoot, "manifests");
const companionOrigin = "https://127.0.0.1:43127";

function fourPartVersion(value, revision = 0) {
  const parts = String(value)
    .split(".")
    .map((part) => part.match(/^\d+/)?.[0])
    .filter(Boolean)
    .slice(0, 4);
  while (parts.length < 4) parts.push("0");
  parts[3] = String(revision);
  return parts.join(".");
}

function renderManifest(template, appVersion) {
  const rendered = template
    .replaceAll("{{WORD_ADDIN_ID}}", config.wordId)
    .replaceAll("{{POWERPOINT_ADDIN_ID}}", config.powerpointId)
    .replaceAll(
      "{{MANIFEST_VERSION}}",
      fourPartVersion(appVersion, manifestRevision),
    )
    .replaceAll("{{COMPANION_ORIGIN}}", companionOrigin);
  if (rendered.includes("{{") || rendered.includes("}}")) {
    throw new Error(`${platform} Office manifest has unresolved placeholders`);
  }
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
  const destination = join(target, relativePath);
  await mkdir(dirname(destination), { recursive: true });
  await writeFile(destination, sanitizeOfficeSource(input), "utf8");
}

await rm(target, { recursive: true, force: true });
await rm(licenses, { recursive: true, force: true });
await rm(icons, { recursive: true, force: true });
await mkdir(target, { recursive: true });
await mkdir(licenses, { recursive: true });
await mkdir(icons, { recursive: true });
await mkdir(manifests, { recursive: true });

for (const file of config.rootFiles) await copySanitized(file);
for (const entry of await readdir(source, { withFileTypes: true })) {
  if (!entry.isDirectory()) continue;
  const relative = join(entry.name, "office_strings.js");
  try {
    await readFile(join(source, relative));
  } catch {
    continue;
  }
  await copySanitized(relative);
}

await cp(join(packageRoot, "LICENSE.md"), join(licenses, "OFFICE-JS-LICENSE.md"));
await cp(join(mathJaxRoot, "LICENSE"), join(licenses, "MATHJAX-LICENSE.txt"));
await cp(join(root, "src-tauri", "icons", "32x32.png"), join(icons, "icon-16.png"));
await cp(join(root, "src-tauri", "icons", "32x32.png"), join(icons, "icon-32.png"));
await cp(join(root, "src-tauri", "icons", "128x128.png"), join(icons, "icon-80.png"));

const packageJson = JSON.parse(await readFile(join(root, "package.json"), "utf8"));
for (const [templateName, outputName] of [
  ["visualtex-word.template.xml", "VisualTeX.Word.xml"],
  ["visualtex-powerpoint.template.xml", "VisualTeX.PowerPoint.xml"],
]) {
  const template = await readFile(join(manifests, templateName), "utf8");
  await writeFile(
    join(manifests, outputName),
    renderManifest(template, packageJson.version),
    "utf8",
  );
}

const officePackageJson = JSON.parse(
  await readFile(join(packageRoot, "package.json"), "utf8"),
);
const officeJs = await readFile(join(target, "office.js"));
await writeFile(
  join(target, "VERSION.json"),
  `${JSON.stringify(
    {
      package: officePackageJson.name,
      version: officePackageJson.version,
      officeJsSha256: createHash("sha256").update(officeJs).digest("hex"),
      platform,
      scope: config.scope,
    },
    null,
    2,
  )}\n`,
  "utf8",
);

console.log(`Prepared independent ${platform} Office resources.`);
