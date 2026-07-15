import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const offlineRoot = join(repositoryRoot, "office", "macos-offline");
const resourcesRoot = join(offlineRoot, "resources");
const packageVersion = JSON.parse(readFileSync(join(repositoryRoot, "package.json"), "utf8")).version;

function usage() {
  process.stderr.write(
    "Usage: node scripts/package_macos_offline_addins.mjs --word /path/VisualTeX.dotm --powerpoint /path/VisualTeX.ppam\n",
  );
}

function argument(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

const wordInput = argument("--word");
const powerpointInput = argument("--powerpoint");
if (!wordInput || !powerpointInput) {
  usage();
  process.exit(2);
}

function run(program, args, options = {}) {
  const encoding = Object.prototype.hasOwnProperty.call(options, "encoding")
    ? options.encoding
    : "utf8";
  return execFileSync(program, args, {
    encoding: encoding === "buffer" ? null : encoding,
    cwd: options.cwd,
    stdio: options.stdio ?? ["ignore", "pipe", "pipe"],
    maxBuffer: 32 * 1024 * 1024,
  });
}

function zipEntries(path) {
  return run("/usr/bin/unzip", ["-Z1", path])
    .split(/\r?\n/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function containsModuleName(buffer, moduleName) {
  return (
    buffer.includes(Buffer.from(moduleName, "utf8")) ||
    buffer.includes(Buffer.from(moduleName, "utf16le"))
  );
}

function validateMacroContainer(path, kind) {
  if (!existsSync(path)) throw new Error(`${kind} add-in does not exist: ${path}`);
  const entries = zipEntries(path);
  const vbaEntry = kind === "Word" ? "word/vbaProject.bin" : "ppt/vbaProject.bin";
  for (const required of ["[Content_Types].xml", "_rels/.rels", vbaEntry]) {
    if (!entries.includes(required)) {
      throw new Error(`${kind} add-in is missing ${required}: ${path}`);
    }
  }
  const vbaProject = run("/usr/bin/unzip", ["-p", path, vbaEntry], { encoding: "buffer" });
  const expectedModules = [
    "VTProtocol",
    "VTMetadata",
    "VTLauncher",
    "VTErrorHandling",
    "VTRibbonCallbacks",
    kind === "Word" ? "VTWordAdapter" : "VTPowerPointAdapter",
  ];
  const missing = expectedModules.filter((moduleName) => !containsModuleName(vbaProject, moduleName));
  if (missing.length > 0) {
    throw new Error(
      `${kind} VBA project does not expose the required module names: ${missing.join(", ")}. Import the reviewed .bas sources before packaging.`,
    );
  }
}

function injectRootRelationship(xml) {
  const relationshipType =
    "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility";
  const withoutOld = xml.replace(
    /\s*<Relationship\b[^>]*Type="http:\/\/schemas\.microsoft\.com\/office\/(?:2006\/relationships\/ui\/extensibility|2007\/relationships\/ui\/extensibility)"[^>]*\/>/gi,
    "",
  );
  const closing = "</Relationships>";
  if (!withoutOld.includes(closing)) {
    throw new Error("OOXML root relationship file is invalid");
  }
  const relationship =
    `<Relationship Id="rIdVisualTeXCustomUI14" Type="${relationshipType}" Target="customUI/customUI14.xml"/>`;
  return withoutOld.replace(closing, `${relationship}${closing}`);
}

function ensureXmlContentType(xml) {
  if (/\bExtension="xml"\s+ContentType="application\/xml"/i.test(xml)) return xml;
  const closing = "</Types>";
  if (!xml.includes(closing)) throw new Error("OOXML [Content_Types].xml is invalid");
  return xml.replace(closing, '<Default Extension="xml" ContentType="application/xml"/></Types>');
}

function packageAddin(inputPath, kind, outputName, ribbonSource) {
  validateMacroContainer(inputPath, kind);
  const temporaryRoot = mkdtempSync(join(tmpdir(), "visualtex-addin-package-"));
  try {
    const unpacked = join(temporaryRoot, "unpacked");
    mkdirSync(unpacked, { recursive: true });
    run("/usr/bin/unzip", ["-qq", inputPath, "-d", unpacked]);

    const customUiDirectory = join(unpacked, "customUI");
    mkdirSync(customUiDirectory, { recursive: true });
    copyFileSync(ribbonSource, join(customUiDirectory, "customUI14.xml"));

    const relationshipsPath = join(unpacked, "_rels", ".rels");
    writeFileSync(
      relationshipsPath,
      injectRootRelationship(readFileSync(relationshipsPath, "utf8")),
      "utf8",
    );
    const contentTypesPath = join(unpacked, "[Content_Types].xml");
    writeFileSync(
      contentTypesPath,
      ensureXmlContentType(readFileSync(contentTypesPath, "utf8")),
      "utf8",
    );

    mkdirSync(resourcesRoot, { recursive: true });
    const temporaryOutput = join(temporaryRoot, outputName);
    run("/usr/bin/zip", ["-X", "-q", "-r", temporaryOutput, "."], { cwd: unpacked });
    validateMacroContainer(temporaryOutput, kind);
    const entries = zipEntries(temporaryOutput);
    if (!entries.includes("customUI/customUI14.xml")) {
      throw new Error(`${kind} packaged add-in is missing customUI/customUI14.xml`);
    }
    const installedRibbon = run(
      "/usr/bin/unzip",
      ["-p", temporaryOutput, "customUI/customUI14.xml"],
    );
    const expectedRibbon = readFileSync(ribbonSource, "utf8");
    if (installedRibbon.trim() !== expectedRibbon.trim()) {
      throw new Error(`${kind} packaged Ribbon differs from the reviewed source XML`);
    }

    const destination = join(resourcesRoot, outputName);
    const staged = `${destination}.staged`;
    copyFileSync(temporaryOutput, staged);
    renameSync(staged, destination);
    return destination;
  } finally {
    rmSync(temporaryRoot, { recursive: true, force: true });
  }
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

try {
  const wordOutput = packageAddin(
    resolve(wordInput),
    "Word",
    "VisualTeX.dotm",
    join(offlineRoot, "word", "customUI14.xml"),
  );
  const powerpointOutput = packageAddin(
    resolve(powerpointInput),
    "PowerPoint",
    "VisualTeX.ppam",
    join(offlineRoot, "powerpoint", "customUI14.xml"),
  );
  const manifest = {
    schemaVersion: 1,
    pluginVersion: packageVersion,
    files: {
      "VisualTeX.dotm": { sha256: sha256(wordOutput) },
      "VisualTeX.ppam": { sha256: sha256(powerpointOutput) },
    },
  };
  writeFileSync(
    join(resourcesRoot, "addins.json"),
    `${JSON.stringify(manifest, null, 2)}\n`,
    "utf8",
  );
  process.stdout.write(
    `Packaged ${basename(wordOutput)} and ${basename(powerpointOutput)} with fixed filenames and reviewed Ribbon XML.\n`,
  );
} catch (error) {
  process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
  process.exitCode = 1;
}
