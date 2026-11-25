// The files of a Windows update, ready for the download server (publish_windows_update.mjs puts them there): both
// installers, each with its signature, and latest.json, the signed manifest VisualTeX's updater reads
// (src-tauri/src/updater.rs): the version, its notes and, for each installer flavour, the installer's address, size,
// SHA-256 and signature. Everything is signed with the release key (`tauri signer`); the key's password comes from
// TAURI_SIGNING_PRIVATE_KEY_PASSWORD, or from password.dpapi beside the key (the password protected with Windows'
// DPAPI for the user who made the key). At the end the folder is checked exactly as the updater checks what it
// downloads (cargo test the_release_folder_checks_out), with the update key the program is built with.
//
// From apps/windows:
//   node scripts/prepare_windows_update.mjs --full <VisualTeX_x.y.z_x64-setup.exe>
//     --no-ocr <VisualTeX_x.y.z_x64-no-ocr-setup.exe> --out <new folder>
//     [--key <release key file>] [--notes <release notes file>] [--release-page <url>] [--version <x.y.z>]
// The key is VISUALTEX_RELEASE_KEY's (the user environment variable the key's maker set) unless given; the version
// is tauri.conf.json's unless given, and the installers must carry it in their names.
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { copyFileSync, createReadStream, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { basename, dirname, join, resolve } from "node:path";

/** Where the updater looks (updater.rs MANIFEST_URL); installers must be on the same server. */
const DOWNLOADS = "https://download.visualtex.pauljianliao.com/visualtex-downloads";
const MANIFEST_FORMAT = "visualtex-update-1";
const TARGETS = { full: "windows-x86_64", noOcr: "windows-x86_64-noocr" };

function fail(message) {
  console.error(`prepare_windows_update: ${message}`);
  process.exit(1);
}

function options(argv) {
  const known = new Set(["--full", "--no-ocr", "--key", "--out", "--notes", "--release-page", "--version"]);
  const values = {};
  for (let i = 0; i < argv.length; i += 2) {
    if (!known.has(argv[i]) || argv[i + 1] === undefined) fail(`unknown or empty option ${argv[i]}`);
    values[argv[i].slice(2)] = argv[i + 1];
  }
  for (const name of ["full", "no-ocr", "out"]) if (!values[name]) fail(`--${name} is required`);
  return values;
}

/** The release key: VISUALTEX_RELEASE_KEY, from this process or from the user's settings (set after it started). */
function defaultKey() {
  if (process.env.VISUALTEX_RELEASE_KEY) return process.env.VISUALTEX_RELEASE_KEY;
  const result = spawnSync("powershell", ["-NoProfile", "-NonInteractive", "-Command", "[Environment]::GetEnvironmentVariable('VISUALTEX_RELEASE_KEY', 'User')"], {
    encoding: "utf8",
  });
  return (result.stdout ?? "").trim() || fail("no release key: give --key or set VISUALTEX_RELEASE_KEY");
}

/** The key's password: the environment's, else password.dpapi beside the key (DPAPI, this Windows user only). */
function keyPassword(keyPath) {
  if (process.env.TAURI_SIGNING_PRIVATE_KEY_PASSWORD) return process.env.TAURI_SIGNING_PRIVATE_KEY_PASSWORD;
  const protectedFile = join(dirname(keyPath), "password.dpapi");
  if (!existsSync(protectedFile)) fail(`no key password: set TAURI_SIGNING_PRIVATE_KEY_PASSWORD or put password.dpapi beside the key`);
  const script =
    "$s = (Get-Content -Raw -LiteralPath $env:VISUALTEX_KEY_PASSWORD_FILE).Trim() | ConvertTo-SecureString; " +
    "[Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s))";
  const result = spawnSync("powershell", ["-NoProfile", "-NonInteractive", "-Command", script], {
    env: { ...process.env, VISUALTEX_KEY_PASSWORD_FILE: protectedFile },
    encoding: "utf8",
  });
  const password = (result.stdout ?? "").trim();
  if (result.status !== 0 || !password) fail(`password.dpapi could not be read (only the Windows user who made it can): ${(result.stderr ?? "").trim()}`);
  return password;
}

/** `tauri signer sign`: writes FILE.sig (base64 of the minisign signature text) and returns it. */
function sign(file, keyPath, password) {
  const cli = resolve("node_modules", "@tauri-apps", "cli", "tauri.js");
  // (a key given in the environment would be taken instead of the file)
  const { TAURI_SIGNING_PRIVATE_KEY: _ignored, ...environment } = process.env;
  const result = spawnSync(process.execPath, [cli, "signer", "sign", "--private-key-path", keyPath, file], {
    env: { ...environment, TAURI_SIGNING_PRIVATE_KEY_PASSWORD: password },
    encoding: "utf8",
  });
  if (result.status !== 0 || !existsSync(`${file}.sig`)) fail(`signing ${basename(file)} failed: ${(result.stderr || result.stdout || "").trim()}`);
  return readFileSync(`${file}.sig`, "utf8").trim();
}

function sha256(file) {
  return new Promise((done, failed) => {
    const hash = createHash("sha256");
    createReadStream(file)
      .on("data", (chunk) => hash.update(chunk))
      .on("end", () => done(hash.digest("hex")))
      .on("error", failed);
  });
}

const args = options(process.argv.slice(2));
const version = args.version ?? JSON.parse(readFileSync(resolve("src-tauri", "tauri.conf.json"), "utf8")).version;
if (!/^\d+\.\d+\.\d+$/.test(version)) fail(`version ${version} is not x.y.z`);
const keyPath = resolve(args.key ?? defaultKey());
if (!existsSync(keyPath)) fail(`no key file ${keyPath}`);
if (args.notes && !existsSync(resolve(args.notes))) fail(`no notes file ${resolve(args.notes)}`);
const out = resolve(args.out);
if (existsSync(out) && readdirSync(out).length) fail(`${out} is not empty`);
mkdirSync(out, { recursive: true });

const installers = {
  full: { path: resolve(args.full), name: `VisualTeX_${version}_x64-setup.exe` },
  noOcr: { path: resolve(args["no-ocr"]), name: `VisualTeX_${version}_x64-no-ocr-setup.exe` },
};
for (const [flavour, installer] of Object.entries(installers)) {
  if (!existsSync(installer.path)) fail(`no ${flavour} installer ${installer.path}`);
  if (basename(installer.path) !== installer.name) fail(`the ${flavour} installer must be named ${installer.name} (it is ${basename(installer.path)})`);
}

const password = keyPassword(keyPath);
const platforms = {};
for (const [flavour, installer] of Object.entries(installers)) {
  const copy = join(out, installer.name);
  copyFileSync(installer.path, copy);
  platforms[TARGETS[flavour]] = {
    url: `${DOWNLOADS}/releases/v${version}/${installer.name}`,
    size: statSync(copy).size,
    sha256: await sha256(copy),
    signature: sign(copy, keyPath, password),
  };
  console.log(`${installer.name}: ${platforms[TARGETS[flavour]].size} bytes, sha256 ${platforms[TARGETS[flavour]].sha256}, signed`);
}

const manifest = {
  product: "VisualTeX",
  version,
  name: `VisualTeX ${version}`,
  notes: args.notes ? readFileSync(resolve(args.notes), "utf8").replace(/\r\n/g, "\n").trim() : "",
  pub_date: new Date().toISOString(),
  release_page: args["release-page"] ?? `https://github.com/paulhe666/visualtex/releases/tag/v${version}`,
  platforms,
};
const manifestFile = join(out, "manifest.json");
writeFileSync(manifestFile, `${JSON.stringify(manifest, null, 2)}\n`);
const envelope = {
  format: MANIFEST_FORMAT,
  manifest: readFileSync(manifestFile).toString("base64"),
  signature: sign(manifestFile, keyPath, password),
};
// (manifest.json stays as the readable copy; its signature is in latest.json)
rmSync(`${manifestFile}.sig`);
writeFileSync(join(out, "latest.json"), `${JSON.stringify(envelope, null, 2)}\n`);
console.log(`latest.json: VisualTeX ${version}, signed`);

// The updater's own check of what it would download, with the key the program is built with (through the Windows
// Rust test runner, which gives the test program the manifest it needs; its own target folder).
console.log("checking the folder as the updater does (the_release_folder_checks_out) ...");
const check = spawnSync(
  process.execPath,
  [resolve("scripts", "run_windows_rust_tests.mjs"), "the_release_folder_checks_out", "--", "--ignored", "--nocapture"],
  {
    env: {
      ...process.env,
      VISUALTEX_RELEASE_DIR: out,
      VISUALTEX_RELEASE_VERSION: version,
      CARGO_TARGET_DIR: process.env.CARGO_TARGET_DIR ?? resolve(".cache", "target-release-check"),
      NO_PROXY: "127.0.0.1,localhost",
    },
    stdio: "inherit",
  },
);
if (check.status !== 0) fail("the folder did not pass the updater's check (wrong key, or a file changed)");
console.log(`ready: ${out} (publish it with scripts/publish_windows_update.mjs)`);
