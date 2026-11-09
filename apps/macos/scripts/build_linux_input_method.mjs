// Builds VisualTeX's Linux input method (an IBus engine) into input-method/build/linux:
// - bin/visualtex-ibus: the engine (input-method/linux, Rust), linked against the system librime;
// - rime: RIME's data, as the macOS input method bundles it (build_macos_input_method.mjs):
//   VisualTeX's default.yaml, pinyin schema and dictionary (input-method/Rime), the dictionary tables
//   and spelling rules of 万象拼音 (wanxiang, amzxyz/rime_wanxiang, a pinned release, CC-BY-4.0:
//   input-method/Rime/NOTICE-wanxiang.txt), the prelude presets (LGPL-3.0, RIME developers), and the
//   tables rime_deployer prebuilds from them (rime/build);
// - visualtex.xml: the IBus component, for /usr/share/ibus/component (the engine in
//   /usr/lib/visualtex, its data in /usr/share/visualtex/rime; VISUALTEX_LINUX_PREFIX for another).
// The downloads are pinned and checked (sha256) and cached in input-method/build/cache;
// VISUALTEX_RIME_MIRROR=<base URL> fetches them from a mirror serving the same file names.
// Needs cargo, librime's development files and rime_deployer (librime-dev, librime-bin).
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

if (process.platform !== "linux") {
  process.stdout.write("The Linux input method is built on Linux only.\n");
  process.exit(0);
}

const DOWNLOADS = [
  { name: "rime-prelude.tar.gz", url: "https://github.com/rime/rime-prelude/archive/082425ea0684bca36474415d4a0e8db9b016487e.tar.gz", sha256: "66239ba4745d54471e5ce5540b45e86ec663ff5dba004cb23c95531a3bdec00f" },
  { name: "rime-wanxiang-base.zip", url: "https://github.com/amzxyz/rime_wanxiang/releases/download/v18.1.10/rime-wanxiang-base.zip", sha256: "ddebf0372a2a3d9d8f7a5017704b2afd7397a961c32be490a91e49ec3ee90999" },
];
// The wanxiang tables visualtex_wanxiang.dict.yaml imports (Rime/), and its spelling rules.
const WANXIANG_TABLES = ["zi", "jichu", "cuoyin", "duoyin", "diming", "mingren"];

const appDirectory = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(appDirectory, "input-method");
const cache = join(source, "build", "cache");
const output = join(source, "build", "linux");
const version = JSON.parse(readFileSync(join(appDirectory, "src-tauri", "tauri.conf.json"), "utf8")).version;
const prefix = process.env.VISUALTEX_LINUX_PREFIX ?? "/usr";
const sha256 = (file) => createHash("sha256").update(readFileSync(file)).digest("hex");

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
const unpacked = join(cache, "unpacked-linux");
rmSync(unpacked, { recursive: true, force: true });
mkdirSync(unpacked, { recursive: true });
execFileSync("tar", ["-xzf", join(cache, "downloads", "rime-prelude.tar.gz"), "-C", unpacked]);
const wanxiang = join(unpacked, "wanxiang");
mkdirSync(wanxiang, { recursive: true });
execFileSync("unzip", ["-q", "-o", join(cache, "downloads", "rime-wanxiang-base.zip"), "-d", wanxiang]);
const prelude = readdirSync(unpacked).find((name) => name.startsWith("rime-prelude"));
if (!prelude) throw new Error("rime-prelude is missing from the downloads");

rmSync(output, { recursive: true, force: true });
const rimeData = join(output, "rime");
mkdirSync(join(rimeData, "dicts"), { recursive: true });
for (const file of ["key_bindings.yaml", "punctuation.yaml", "symbols.yaml"]) {
  cpSync(join(unpacked, prelude, file), join(rimeData, file));
}
for (const table of WANXIANG_TABLES) {
  cpSync(join(wanxiang, "dicts", `${table}.dict.yaml`), join(rimeData, "dicts", `${table}.dict.yaml`));
}
cpSync(join(wanxiang, "wanxiang_algebra.yaml"), join(rimeData, "wanxiang_algebra.yaml"));
for (const file of readdirSync(join(source, "Rime"))) {
  cpSync(join(source, "Rime", file), join(rimeData, file));
}
const deployUser = join(cache, "deploy-user-linux");
rmSync(deployUser, { recursive: true, force: true });
mkdirSync(deployUser, { recursive: true });
execFileSync("rime_deployer", ["--build", deployUser, rimeData, join(rimeData, "build")], {
  stdio: ["ignore", "ignore", "inherit"],
});
if (!existsSync(join(rimeData, "build", "visualtex_pinyin.prism.bin"))) {
  throw new Error("rime_deployer did not build the VisualTeX pinyin schema");
}

// The engine.
const target = process.env.CARGO_TARGET_DIR ?? join(source, "linux", "target");
execFileSync("cargo", ["build", "--release", "--manifest-path", join(source, "linux", "Cargo.toml")], {
  stdio: "inherit",
  env: { ...process.env, CARGO_TARGET_DIR: target },
});
mkdirSync(join(output, "bin"), { recursive: true });
cpSync(join(target, "release", "visualtex-ibus"), join(output, "bin", "visualtex-ibus"));

// The IBus component.
const component = `<?xml version="1.0" encoding="utf-8"?>
<!-- VisualTeX's input method: pinyin, English and formulas (installed with VisualTeX) -->
<component>
  <name>org.freedesktop.IBus.VisualTeX</name>
  <description>VisualTeX input method</description>
  <exec>${prefix}/lib/visualtex/visualtex-ibus --ibus</exec>
  <version>${version}</version>
  <author>paulhe666</author>
  <license>MIT</license>
  <homepage>https://github.com/paulhe666/visualtex</homepage>
  <textdomain>visualtex</textdomain>
  <engines>
    <engine>
      <name>visualtex</name>
      <language>zh_CN</language>
      <license>MIT</license>
      <author>paulhe666</author>
      <icon>${prefix}/share/icons/hicolor/128x128/apps/visualtex.png</icon>
      <layout>default</layout>
      <longname>VisualTeX</longname>
      <description>VisualTeX: pinyin and formulas</description>
      <rank>0</rank>
      <symbol>V</symbol>
    </engine>
  </engines>
</component>
`;
writeFileSync(join(output, "visualtex.xml"), component);
process.stdout.write(`VisualTeX input method ${version} in ${output}\n`);
