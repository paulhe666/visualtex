// Builds VisualTeX's macOS input method (input-method/) into
// input-method/build/VisualTeX Input.app:
// - the Swift sources compiled with swiftc against librime (RIME, BSD
//   licence), which goes into Contents/Frameworks;
// - RIME's data in Contents/SharedSupport/rime: VisualTeX's default.yaml,
//   pinyin schema and dictionary (input-method/Rime), the dictionary tables
//   and spelling rules of 万象拼音 (wanxiang, amzxyz/rime_wanxiang, a pinned
//   release, CC-BY-4.0: input-method/Rime/NOTICE-wanxiang.txt), the prelude
//   presets (LGPL-3.0, RIME developers), and the tables rime_deployer
//   prebuilds from them (so the first keystroke is not slow);
// - the menu icon as a TIFF, Info.plist with the app's version, and a build
//   id (Contents/Resources/build-id) the app compares to decide whether to
//   update the copy in ~/Library/Input Methods (src-tauri/src/anywhere/ime.rs).
// The downloads are pinned and checked (sha256) and cached in
// input-method/build/cache. VISUALTEX_RIME_MIRROR=<base URL> fetches them from
// a mirror that serves the same file names.
// Signed with the release identity when tauri_build.mjs provides one
// (APPLE_SIGNING_IDENTITY), ad hoc otherwise. tauri.macos.conf.json bundles
// the result as a resource.
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

if (process.platform !== "darwin") {
  process.stdout.write("The macOS input method is built on macOS only.\n");
  process.exit(0);
}

const RIME_RELEASE = "https://github.com/rime/librime/releases/download/1.17.0";
const DOWNLOADS = [
  { name: "rime-mac.tar.bz2", url: `${RIME_RELEASE}/rime-33e7814-macOS-universal.tar.bz2`, sha256: "11d8dc663c6ec06d5ccb6111ba664a9e7b631b703ac6acd07cffbac664021850" },
  { name: "rime-prelude.tar.gz", url: "https://github.com/rime/rime-prelude/archive/082425ea0684bca36474415d4a0e8db9b016487e.tar.gz", sha256: "66239ba4745d54471e5ce5540b45e86ec663ff5dba004cb23c95531a3bdec00f" },
  // 万象拼音 v18.1.10 (2026-11-07), its base package: dictionary tables and spelling rules
  { name: "rime-wanxiang-base.zip", url: "https://github.com/amzxyz/rime_wanxiang/releases/download/v18.1.10/rime-wanxiang-base.zip", sha256: "ddebf0372a2a3d9d8f7a5017704b2afd7397a961c32be490a91e49ec3ee90999" },
];
// The wanxiang tables visualtex_wanxiang.dict.yaml imports (Rime/), and its spelling rules.
const WANXIANG_TABLES = ["zi", "jichu", "cuoyin", "duoyin", "diming", "mingren"];

const appDirectory = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(appDirectory, "input-method");
const cache = join(source, "build", "cache");
const bundle = join(source, "build", "VisualTeX Input.app");
const contents = join(bundle, "Contents");
const version = JSON.parse(readFileSync(join(appDirectory, "src-tauri", "tauri.conf.json"), "utf8")).version;
const sha256 = (file) => createHash("sha256").update(readFileSync(file)).digest("hex");

// Downloads, checked.
mkdirSync(join(cache, "downloads"), { recursive: true });
for (const download of DOWNLOADS) {
  const file = join(cache, "downloads", download.name);
  if (existsSync(file) && sha256(file) === download.sha256) continue;
  const url = process.env.VISUALTEX_RIME_MIRROR ? `${process.env.VISUALTEX_RIME_MIRROR.replace(/\/$/, "")}/${download.name}` : download.url;
  process.stdout.write(`Downloading ${download.name}\n`);
  execFileSync("curl", ["-fsSL", "--retry", "3", "-o", file, url], { stdio: "inherit" });
  const actual = sha256(file);
  if (actual !== download.sha256) {
    rmSync(file, { force: true });
    throw new Error(`${download.name}: sha256 ${actual}, expected ${download.sha256}`);
  }
}
const unpacked = join(cache, "unpacked");
rmSync(unpacked, { recursive: true, force: true });
mkdirSync(unpacked, { recursive: true });
execFileSync("tar", ["-xjf", join(cache, "downloads", "rime-mac.tar.bz2"), "-C", unpacked]);
execFileSync("tar", ["-xzf", join(cache, "downloads", "rime-prelude.tar.gz"), "-C", unpacked]);
const wanxiang = join(unpacked, "wanxiang");
mkdirSync(wanxiang, { recursive: true });
execFileSync("unzip", ["-q", "-o", join(cache, "downloads", "rime-wanxiang-base.zip"), "-d", wanxiang]);
const unpackedDirectory = (prefix) => {
  const entry = readdirSync(unpacked).find((name) => name.startsWith(`${prefix}-`));
  if (!entry) throw new Error(`${prefix} is missing from the downloads`);
  return join(unpacked, entry);
};
const rimeDist = join(unpacked, "dist");

// The bundle.
rmSync(bundle, { recursive: true, force: true });
for (const folder of ["MacOS", "Resources", "Frameworks", "SharedSupport/rime/dicts"]) {
  mkdirSync(join(contents, folder), { recursive: true });
}

// RIME's data, then its prebuilt tables.
const rimeData = join(contents, "SharedSupport", "rime");
for (const file of ["key_bindings.yaml", "punctuation.yaml", "symbols.yaml"]) {
  cpSync(join(unpackedDirectory("rime-prelude"), file), join(rimeData, file));
}
for (const table of WANXIANG_TABLES) {
  cpSync(join(wanxiang, "dicts", `${table}.dict.yaml`), join(rimeData, "dicts", `${table}.dict.yaml`));
}
cpSync(join(wanxiang, "wanxiang_algebra.yaml"), join(rimeData, "wanxiang_algebra.yaml"));
for (const file of readdirSync(join(source, "Rime"))) {
  cpSync(join(source, "Rime", file), join(rimeData, file));
}
const deployUser = join(cache, "deploy-user");
rmSync(deployUser, { recursive: true, force: true });
mkdirSync(deployUser, { recursive: true });
execFileSync(join(rimeDist, "bin", "rime_deployer"), ["--build", deployUser, rimeData, join(rimeData, "build")], {
  stdio: ["ignore", "ignore", "inherit"],
  env: { ...process.env, DYLD_LIBRARY_PATH: join(rimeDist, "lib") },
});
if (!existsSync(join(rimeData, "build", "visualtex_pinyin.prism.bin"))) {
  throw new Error("rime_deployer did not build the VisualTeX pinyin schema");
}

// The program, linked against the bundled librime.
cpSync(join(rimeDist, "lib", "librime.1.dylib"), join(contents, "Frameworks", "librime.1.dylib"));
const swiftSources = readdirSync(join(source, "Sources"))
  .filter((name) => name.endsWith(".swift"))
  .sort()
  .map((name) => join(source, "Sources", name));
const executable = join(contents, "MacOS", "VisualTeXInput");
execFileSync(
  "swiftc",
  [
    "-swift-version", "5", "-O", "-target", "arm64-apple-macos11.0",
    "-import-objc-header", join(source, "Sources", "RimeBridge.h"),
    "-Xcc", `-I${join(rimeDist, "include")}`,
    "-L", join(rimeDist, "lib"), "-lrime",
    "-Xlinker", "-rpath", "-Xlinker", "@executable_path/../Frameworks",
    "-o", executable,
    ...swiftSources,
    "-framework", "InputMethodKit", "-framework", "Cocoa", "-framework", "Carbon", "-framework", "Security",
  ],
  { stdio: "inherit" },
);

const plist = readFileSync(join(source, "Info.plist"), "utf8")
  .replaceAll("__VERSION__", version)
  .replaceAll("__BUILD__", version);
writeFileSync(join(contents, "Info.plist"), plist);
execFileSync("tiffutil", [
  "-cathidpicheck",
  join(source, "Icon", "VisualTeX16.png"),
  join(source, "Icon", "VisualTeX32.png"),
  "-out", join(contents, "Resources", "VisualTeX.tiff"),
], { stdio: "ignore" });

const identity = process.env.APPLE_SIGNING_IDENTITY;
const signWith = identity && identity !== "-" ? identity : "-";

// The build id: what the input method does and who signed it (VisualTeX
// links only to an input method signed like itself), not when it was built.
const digest = createHash("sha256");
const inputs = [
  ...swiftSources,
  join(source, "Sources", "RimeBridge.h"),
  join(source, "Info.plist"),
  ...readdirSync(join(source, "Rime")).sort().map((file) => join(source, "Rime", file)),
];
for (const file of inputs) digest.update(basename(file)).update(readFileSync(file));
for (const download of DOWNLOADS) digest.update(download.sha256);
digest.update(version);
digest.update(signWith);
writeFileSync(join(contents, "Resources", "build-id"), `${digest.digest("hex").slice(0, 16)}\n`);
execFileSync("codesign", ["--force", "--timestamp=none", "-s", signWith, join(contents, "Frameworks", "librime.1.dylib")], { stdio: "inherit" });
execFileSync("codesign", ["--force", "--timestamp=none", "-s", signWith, bundle], { stdio: "inherit" });
process.stdout.write(
  `Built the VisualTeX input method ${version} (${signWith === "-" ? "ad hoc" : "signed"}): ${bundle}\n`,
);
