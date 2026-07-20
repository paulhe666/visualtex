import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  chmodSync,
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { homedir, tmpdir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const offlineRoot = join(repositoryRoot, "office", "macos-offline");
const resourcesRoot = join(offlineRoot, "resources");
const packageVersion = JSON.parse(readFileSync(join(repositoryRoot, "package.json"), "utf8")).version;

function usage() {
  process.stderr.write(
    "Usage: node scripts/package_macos_offline_addins.mjs --word /path/VisualTeX.dotm (--word-only | --powerpoint /path/VisualTeX.pptm [--powerpoint-shell /path/known-good-VisualTeX.ppam]) [--ribbon-icons-archive /path/visualtex-icons.zip] [--artifacts-dir /path/to/final/files] [--root-word-output /path/VisualTeX.dotm] [--root-powerpoint-output /path/VisualTeX.ppam] [--install-macos]\n",
  );
}

function argument(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

const wordInput = argument("--word");
const powerpointInput = argument("--powerpoint");
const powerpointShell = argument("--powerpoint-shell");
const artifactsDirectory = argument("--artifacts-dir");
const rootWordOutput = argument("--root-word-output");
const rootPowerPointOutput = argument("--root-powerpoint-output");
const ribbonIconsArchive = argument("--ribbon-icons-archive");
const installMacos = process.argv.includes("--install-macos");
const wordOnly = process.argv.includes("--word-only");
if (!wordInput || (!wordOnly && !powerpointInput) || (wordOnly && powerpointInput)) {
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

const MAIN_CONTENT_TYPES = {
  Word: {
    partName: "/word/document.xml",
    contentType: "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
  },
  PowerPoint: {
    partName: "/ppt/presentation.xml",
    contentType: "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml",
  },
};

function contentTypesXml(path) {
  return run("/usr/bin/unzip", ["-p", path, "\\[Content_Types\\].xml"]);
}

function hasExpectedMainContentType(path, kind) {
  const expected = MAIN_CONTENT_TYPES[kind];
  const xml = contentTypesXml(path);
  const escapedPart = expected.partName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const escapedType = expected.contentType.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(
    `<Override\\b(?=[^>]*\\bPartName="${escapedPart}")(?=[^>]*\\bContentType="${escapedType}")[^>]*/>`,
    "i",
  ).test(xml);
}

function validateMacroContainer(path, kind, options = {}) {
  const requireExpectedMainType = options.requireExpectedMainType ?? true;
  const requireModules = options.requireModules ?? true;
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
    "VTOfficePaths",
    "VTMetadata",
    "VTLauncher",
    "VTErrorHandling",
    "VTRibbonCallbacks",
    kind === "Word" ? "VTWordAdapter" : "VTPowerPointAdapter",
    kind === "Word" ? "VTWordEvents" : "VTPowerPointEvents",
  ];
  if (requireModules) {
    const missing = expectedModules.filter((moduleName) => !containsModuleName(vbaProject, moduleName));
    if (missing.length > 0) {
      throw new Error(
        `${kind} VBA project does not expose the required module names: ${missing.join(", ")}. Import the reviewed .bas and .cls sources before packaging.`,
      );
    }
  }
  if (requireExpectedMainType && !hasExpectedMainContentType(path, kind)) {
    const expected = MAIN_CONTENT_TYPES[kind];
    throw new Error(
      `${kind} add-in has the wrong OOXML main content type. Expected ${expected.contentType} for ${expected.partName}. Save the file as a real ${kind} add-in instead of renaming another Office format.`,
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

function ensureDefaultContentType(xml, extension, contentType) {
  const closing = "</Types>";
  if (!xml.includes(closing)) throw new Error("OOXML [Content_Types].xml is invalid");
  const escapedType = contentType.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const existing = new RegExp(
    `\\bExtension="${extension}"\\s+ContentType="${escapedType}"`,
    "i",
  );
  if (existing.test(xml)) return xml;
  return xml.replace(
    closing,
    `<Default Extension="${extension}" ContentType="${contentType}"/>${closing}`,
  );
}

function ensureRibbonContentTypes(xml, hasPngIcons) {
  let next = ensureDefaultContentType(xml, "xml", "application/xml");
  if (hasPngIcons) next = ensureDefaultContentType(next, "png", "image/png");
  return next;
}

const RIBBON_ICON_FILES = {
  Word: {
    VisualTeXIcon02: "icon_02_same_subject.png",
    VisualTeXIcon04: "icon_04_same_subject.png",
    VisualTeXIcon06: "icon_06_same_subject.png",
    VisualTeXIcon07: "icon_07_same_subject.png",
    VisualTeXIcon08: "icon_08_same_subject.png",
    VisualTeXIcon09: "icon_09_same_subject.png",
  },
  PowerPoint: {
    VisualTeXIcon05: "icon_05_same_subject.png",
    VisualTeXIcon07: "icon_07_same_subject.png",
  },
};

function referencedRibbonImages(ribbonXml) {
  return [...ribbonXml.matchAll(/\bimage="([A-Za-z0-9_.-]+)"/g)].map(
    (match) => match[1],
  );
}

function installRibbonImages(unpacked, kind, ribbonXml) {
  const referenced = referencedRibbonImages(ribbonXml);
  if (referenced.length === 0) return [];
  const mapping = RIBBON_ICON_FILES[kind];
  const unique = [...new Set(referenced)];
  for (const imageId of unique) {
    if (!mapping[imageId]) {
      throw new Error(`${kind} Ribbon references an unmapped custom image: ${imageId}`);
    }
  }
  const customUiDirectory = join(unpacked, "customUI");
  const imagesDirectory = join(customUiDirectory, "images");
  const relationshipsDirectory = join(customUiDirectory, "_rels");
  mkdirSync(imagesDirectory, { recursive: true });
  mkdirSync(relationshipsDirectory, { recursive: true });
  const archivePath =
    ribbonIconsArchive && existsSync(resolve(ribbonIconsArchive))
      ? resolve(ribbonIconsArchive)
      : undefined;
  const relationships = unique.map((imageId) => {
    const embeddedPath = join(imagesDirectory, `${imageId}.png`);
    const imageBytes = archivePath
      ? run(
          "/usr/bin/unzip",
          ["-p", archivePath, mapping[imageId]],
          { encoding: "buffer" },
        )
      : existsSync(embeddedPath)
        ? readFileSync(embeddedPath)
        : undefined;
    if (
      !imageBytes ||
      imageBytes.length < 8 ||
      !imageBytes.subarray(0, 8).equals(Buffer.from("89504e470d0a1a0a", "hex"))
    ) {
      throw new Error(
        archivePath
          ? `${mapping[imageId]} is not a valid PNG Ribbon image.`
          : `${kind} Ribbon image ${imageId} is missing from the reviewed Office container; provide --ribbon-icons-archive to replace it.`,
      );
    }
    writeFileSync(embeddedPath, imageBytes);
    return `<Relationship Id="${imageId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/${imageId}.png"/>`;
  });
  writeFileSync(
    join(relationshipsDirectory, "customUI14.xml.rels"),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join("")}</Relationships>\n`,
    "utf8",
  );
  return unique;
}

function packageAddin(inputPath, kind, outputName, ribbonSource, shellPath) {
  validateMacroContainer(inputPath, kind, {
    requireExpectedMainType: kind !== "PowerPoint" || !shellPath,
  });
  if (shellPath) {
    validateMacroContainer(shellPath, kind, { requireModules: false });
  }
  const temporaryRoot = mkdtempSync(join(tmpdir(), "visualtex-addin-package-"));
  try {
    const unpacked = join(temporaryRoot, "unpacked");
    mkdirSync(unpacked, { recursive: true });
    const packageSource = shellPath || inputPath;
    run("/usr/bin/unzip", ["-qq", packageSource, "-d", unpacked]);
    if (shellPath) {
      const vbaEntry = kind === "Word" ? "word/vbaProject.bin" : "ppt/vbaProject.bin";
      const vbaProject = run("/usr/bin/unzip", ["-p", inputPath, vbaEntry], {
        encoding: "buffer",
      });
      writeFileSync(join(unpacked, vbaEntry), vbaProject);
    }

    const customUiDirectory = join(unpacked, "customUI");
    mkdirSync(customUiDirectory, { recursive: true });
    copyFileSync(ribbonSource, join(customUiDirectory, "customUI14.xml"));
    const ribbonXml = readFileSync(ribbonSource, "utf8");
    const installedRibbonImages = installRibbonImages(unpacked, kind, ribbonXml);

    const relationshipsPath = join(unpacked, "_rels", ".rels");
    writeFileSync(
      relationshipsPath,
      injectRootRelationship(readFileSync(relationshipsPath, "utf8")),
      "utf8",
    );
    const contentTypesPath = join(unpacked, "[Content_Types].xml");
    writeFileSync(
      contentTypesPath,
      ensureRibbonContentTypes(
        readFileSync(contentTypesPath, "utf8"),
        installedRibbonImages.length > 0,
      ),
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
    for (const imageId of installedRibbonImages) {
      if (!entries.includes(`customUI/images/${imageId}.png`)) {
        throw new Error(`${kind} packaged Ribbon is missing ${imageId}.png`);
      }
    }
    if (
      installedRibbonImages.length > 0 &&
      !entries.includes("customUI/_rels/customUI14.xml.rels")
    ) {
      throw new Error(`${kind} packaged Ribbon image relationships are missing`);
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

function atomicCopy(source, destination, mode) {
  mkdirSync(dirname(destination), { recursive: true });
  const staged = `${destination}.staged`;
  copyFileSync(source, staged);
  if (mode !== undefined) chmodSync(staged, mode);
  renameSync(staged, destination);
  if (sha256(source) !== sha256(destination)) {
    throw new Error(`Installed artifact differs from its packaged source: ${destination}`);
  }
  return destination;
}

function syncArtifact(source, outputDirectory) {
  mkdirSync(outputDirectory, { recursive: true });
  const destination = join(outputDirectory, basename(source));
  return atomicCopy(source, destination);
}

function installMacosArtifacts(wordOutput, powerpointOutput) {
  if (process.platform !== "darwin") {
    throw new Error("--install-macos is available only on macOS");
  }
  const officeGroupContainer = join(
    homedir(),
    "Library",
    "Group Containers",
    "UBF8T346G9.Office",
  );
  const installed = [
    atomicCopy(
      wordOutput,
      join(
        officeGroupContainer,
        "User Content.localized",
        "Startup.localized",
        "Word",
        "VisualTeX.dotm",
      ),
      0o600,
    ),
  ];
  if (powerpointOutput) {
    installed.push(
      atomicCopy(
        powerpointOutput,
        join(
          officeGroupContainer,
          "VisualTeX",
          "OfficeAddins",
          "VisualTeX.ppam",
        ),
        0o600,
      ),
    );
  }
  return installed;
}

try {
  const wordOutput = packageAddin(
    resolve(wordInput),
    "Word",
    "VisualTeX.dotm",
    join(offlineRoot, "word", "customUI14.xml"),
  );
  let manifest;
  let powerpointOutput;
  if (wordOnly) {
    const manifestPath = join(resourcesRoot, "addins.json");
    manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
    manifest.pluginVersion = packageVersion;
    manifest.files["VisualTeX.dotm"] = { sha256: sha256(wordOutput) };
  } else {
    powerpointOutput = packageAddin(
      resolve(powerpointInput),
      "PowerPoint",
      "VisualTeX.ppam",
      join(offlineRoot, "powerpoint", "customUI14.xml"),
      powerpointShell ? resolve(powerpointShell) : undefined,
    );
    manifest = {
      schemaVersion: 1,
      pluginVersion: packageVersion,
      files: {
        "VisualTeX.dotm": { sha256: sha256(wordOutput) },
        "VisualTeX.ppam": { sha256: sha256(powerpointOutput) },
      },
    };
  }
  writeFileSync(
    join(resourcesRoot, "addins.json"),
    `${JSON.stringify(manifest, null, 2)}\n`,
    "utf8",
  );
  if (artifactsDirectory) {
    syncArtifact(wordOutput, resolve(artifactsDirectory));
    if (powerpointOutput) syncArtifact(powerpointOutput, resolve(artifactsDirectory));
  }
  if (rootWordOutput) {
    atomicCopy(wordOutput, resolve(rootWordOutput));
  }
  if (rootPowerPointOutput) {
    if (!powerpointOutput) {
      throw new Error("--root-powerpoint-output requires a PowerPoint package build.");
    }
    atomicCopy(powerpointOutput, resolve(rootPowerPointOutput));
  }
  if (installMacos) installMacosArtifacts(wordOutput, powerpointOutput);
  process.stdout.write(
    wordOnly
      ? `Packaged ${basename(wordOutput)} with the reviewed Word Ribbon XML; PowerPoint was not touched.\n`
      : "Packaged VisualTeX.dotm and VisualTeX.ppam with fixed filenames and reviewed Ribbon XML.\n",
  );
} catch (error) {
  process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
  process.exitCode = 1;
}
