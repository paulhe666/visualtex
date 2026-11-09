// Builds VisualTeX's Windows input method (input-method/) into input-method/build/VisualTeX Input:
// - visualtex_tip.dll, the TSF text input processor (registered by the installer, elevated), and
//   VisualTeXInput.exe, the per-user server with RIME and the candidate window (cargo, release);
// - rime.dll (librime, BSD licence) from librime's Windows release;
// - RIME's data in data\: VisualTeX's default.yaml, pinyin schema and dictionary (input-method/Rime),
//   the dictionary tables and spelling rules of 万象拼音 (wanxiang, amzxyz/rime_wanxiang, a pinned
//   release, CC-BY-4.0: input-method/Rime/NOTICE-wanxiang.txt), the prelude presets (LGPL-3.0, RIME
//   developers), and the tables rime_deployer prebuilds from them in data\build (so the first
//   keystroke is not slow).
// The downloads are pinned and checked (sha256) and cached in input-method/build/cache.
// VISUALTEX_RIME_MIRROR=<base URL> fetches them from a mirror that serves the same file names.
// librime's Windows release is a .7z: tar extracts it where libarchive has LZMA (Windows 11),
// else Python's py7zr does (installed for the user on first use).
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

if (process.platform !== "win32") {
  process.stdout.write("The Windows input method is built on Windows only.\n");
  process.exit(0);
}

const RIME_RELEASE = "https://github.com/rime/librime/releases/download/1.17.0";
const DOWNLOADS = [
  { name: "rime-windows.7z", url: `${RIME_RELEASE}/rime-33e7814-Windows-msvc-x64.7z`, sha256: "7478c7caa4ff6b37de86daba1f7ce4a994a4f5ba24872a820fb2b3a9b01fed15" },
  { name: "rime-prelude.tar.gz", url: "https://github.com/rime/rime-prelude/archive/082425ea0684bca36474415d4a0e8db9b016487e.tar.gz", sha256: "66239ba4745d54471e5ce5540b45e86ec663ff5dba004cb23c95531a3bdec00f" },
  // 万象拼音 v18.1.10 (2026-11-07), its base package: dictionary tables and spelling rules
  { name: "rime-wanxiang-base.zip", url: "https://github.com/amzxyz/rime_wanxiang/releases/download/v18.1.10/rime-wanxiang-base.zip", sha256: "ddebf0372a2a3d9d8f7a5017704b2afd7397a961c32be490a91e49ec3ee90999" },
];
// The wanxiang tables visualtex_wanxiang.dict.yaml imports (Rime/), and its spelling rules.
const WANXIANG_TABLES = ["zi", "jichu", "cuoyin", "duoyin", "diming", "mingren"];

const appDirectory = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(appDirectory, "input-method");
const cache = join(source, "build", "cache");
const output = join(source, "build", "VisualTeX Input");
const sha256 = (file) => createHash("sha256").update(readFileSync(file)).digest("hex");

// Downloads, checked.
mkdirSync(join(cache, "downloads"), { recursive: true });
for (const download of DOWNLOADS) {
  const file = join(cache, "downloads", download.name);
  if (existsSync(file) && sha256(file) === download.sha256) continue;
  const url = process.env.VISUALTEX_RIME_MIRROR ? `${process.env.VISUALTEX_RIME_MIRROR.replace(/\/$/, "")}/${download.name}` : download.url;
  process.stdout.write(`Downloading ${download.name}\n`);
  execFileSync("curl.exe", ["-fsSL", "--retry", "3", "-o", file, url], { stdio: "inherit" });
  const actual = sha256(file);
  if (actual !== download.sha256) {
    rmSync(file, { force: true });
    throw new Error(`${download.name}: sha256 ${actual}, expected ${download.sha256}`);
  }
}

function extract7z(archive, target) {
  mkdirSync(target, { recursive: true });
  try {
    execFileSync("tar", ["-xf", archive, "-C", target], { stdio: "ignore" });
    return;
  } catch {
    // This tar has no LZMA: py7zr below.
  }
  const python = ["python", "py"].find((name) => {
    try {
      execFileSync(name, ["--version"], { stdio: "ignore" });
      return true;
    } catch {
      return false;
    }
  });
  if (!python) throw new Error(`Extracting ${archive} needs a tar with LZMA (Windows 11) or Python.`);
  try {
    execFileSync(python, ["-c", "import py7zr"], { stdio: "ignore" });
  } catch {
    execFileSync(python, ["-m", "pip", "install", "--user", "--quiet", "py7zr"], { stdio: "inherit" });
  }
  execFileSync(python, ["-c", "import sys, py7zr; py7zr.SevenZipFile(sys.argv[1]).extractall(sys.argv[2])", archive, target], {
    stdio: "inherit",
  });
}

const unpacked = join(cache, "unpacked");
rmSync(unpacked, { recursive: true, force: true });
extract7z(join(cache, "downloads", "rime-windows.7z"), unpacked);
execFileSync("tar", ["-xzf", join(cache, "downloads", "rime-prelude.tar.gz"), "-C", unpacked]);
const wanxiang = join(unpacked, "wanxiang");
mkdirSync(wanxiang, { recursive: true });
execFileSync("tar", ["-xf", join(cache, "downloads", "rime-wanxiang-base.zip"), "-C", wanxiang]);
const unpackedDirectory = (prefix) => {
  const entry = readdirSync(unpacked).find((name) => name.startsWith(`${prefix}-`));
  if (!entry) throw new Error(`${prefix} is missing from the downloads`);
  return join(unpacked, entry);
};
const rimeDist = join(unpacked, "dist");

// The folder.
rmSync(output, { recursive: true, force: true });
const rimeData = join(output, "data");
mkdirSync(join(rimeData, "dicts"), { recursive: true });

// RIME's data, then its prebuilt tables.
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
// rime_deployer finds rime.dll beside it.
cpSync(join(rimeDist, "lib", "rime.dll"), join(rimeDist, "bin", "rime.dll"));
execFileSync(join(rimeDist, "bin", "rime_deployer.exe"), ["--build", deployUser, rimeData, join(rimeData, "build")], {
  stdio: ["ignore", "ignore", "inherit"],
});
if (!existsSync(join(rimeData, "build", "visualtex_pinyin.prism.bin"))) {
  throw new Error("rime_deployer did not build the VisualTeX pinyin schema");
}

// The programs.
const target = join(source, "build", "target");
execFileSync("cargo", ["build", "--release", "--manifest-path", join(source, "Cargo.toml"), "--target-dir", target], {
  stdio: "inherit",
});
cpSync(join(target, "release", "visualtex_tip.dll"), join(output, "visualtex_tip.dll"));
cpSync(join(target, "release", "VisualTeXInput.exe"), join(output, "VisualTeXInput.exe"));
cpSync(join(rimeDist, "lib", "rime.dll"), join(output, "rime.dll"));

// A build id: what the input method does (the installer and the app compare it to decide whether
// the installed copy is current).
const digest = createHash("sha256");
for (const file of ["visualtex_tip.dll", "VisualTeXInput.exe", "rime.dll"]) digest.update(readFileSync(join(output, file)));
for (const file of readdirSync(join(source, "Rime")).sort()) digest.update(readFileSync(join(source, "Rime", file)));
for (const download of DOWNLOADS) digest.update(download.sha256);
writeFileSync(join(output, "build-id"), `${digest.digest("hex").slice(0, 16)}\n`);
process.stdout.write(`Built ${output}\n`);
