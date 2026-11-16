// Builds VisualTeX's macOS input method (input-method/) into
// input-method/build/VisualTeX Formula.app: the Swift sources compiled with swiftc, Info.plist with
// the app's version, and a build id (Contents/Resources/build-id) the app compares to decide whether
// to update the copy in ~/Library/Input Methods (src-tauri/src/anywhere/ime.rs).
// Signed with the release identity when tauri_build.mjs provides one (APPLE_SIGNING_IDENTITY), ad hoc
// otherwise. tauri.macos.conf.json bundles the result as a resource.
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

if (process.platform !== "darwin") {
  process.stdout.write("The macOS input method is built on macOS only.\n");
  process.exit(0);
}

const appDirectory = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(appDirectory, "input-method");
const bundle = join(source, "build", "VisualTeX Formula.app");
const contents = join(bundle, "Contents");
const version = JSON.parse(readFileSync(join(appDirectory, "src-tauri", "tauri.conf.json"), "utf8")).version;

// (the bundle of the earlier input method, a keyboard input source with pinyin, is no longer built)
rmSync(join(source, "build", "VisualTeX Input.app"), { recursive: true, force: true });
rmSync(bundle, { recursive: true, force: true });
for (const folder of ["MacOS", "Resources"]) {
  mkdirSync(join(contents, folder), { recursive: true });
}

const swiftSources = readdirSync(join(source, "Sources"))
  .filter((name) => name.endsWith(".swift"))
  .sort()
  .map((name) => join(source, "Sources", name));
const executable = join(contents, "MacOS", "VisualTeXFormula");
execFileSync(
  "swiftc",
  [
    "-swift-version", "5", "-O", "-target", "arm64-apple-macos11.0",
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

const identity = process.env.APPLE_SIGNING_IDENTITY;
const signWith = identity && identity !== "-" ? identity : "-";

// The build id: what the input method does and who signed it (VisualTeX
// links only to an input method signed like itself), not when it was built.
const digest = createHash("sha256");
for (const file of [...swiftSources, join(source, "Info.plist")]) digest.update(basename(file)).update(readFileSync(file));
digest.update(version);
digest.update(signWith);
writeFileSync(join(contents, "Resources", "build-id"), `${digest.digest("hex").slice(0, 16)}\n`);
execFileSync("codesign", ["--force", "--timestamp=none", "-s", signWith, bundle], { stdio: "inherit" });
process.stdout.write(
  `Built the VisualTeX input method ${version} (${signWith === "-" ? "ad hoc" : "signed"}): ${bundle}\n`,
);
