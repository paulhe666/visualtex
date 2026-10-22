import { spawnSync } from "node:child_process";
import { randomBytes } from "node:crypto";
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";

function runStatus(command, args, env = process.env) {
  const isWindowsCmd =
    process.platform === "win32" && command.toLowerCase().endsWith(".cmd");
  const executable = isWindowsCmd ? (process.env.ComSpec ?? "cmd.exe") : command;
  const executableArgs = isWindowsCmd
    ? ["/d", "/s", "/c", command, ...args]
    : args;
  const result = spawnSync(executable, executableArgs, {
    stdio: "inherit",
    shell: false,
    env,
  });
  if (result.error) throw result.error;
  return result.status ?? 1;
}

const tauri = process.platform === "win32" ? "tauri.cmd" : "tauri";

// tauri.conf.json owns the single build:bundle invocation through its
// beforeBuildCommand. Running it here as well compiled the frontend and checked
// the Office resources twice for every release build.

if (process.platform === "darwin") {
  // Tauri may reuse generated resource and bundle directories between builds.
  // Remove them after preparing the verified source bundle so stale interrupted
  // OCR archive temporaries can never survive into a new app or DMG.
  const macosBundleDirectory = "src-tauri/target/release/bundle/macos";
  try {
    for (const entry of readdirSync(macosBundleDirectory)) {
      if (entry.endsWith(".dmg")) {
        rmSync(join(macosBundleDirectory, entry), { force: true });
      }
    }
  } catch (error) {
    if (!(error instanceof Error && "code" in error && error.code === "ENOENT")) {
      throw error;
    }
  }

  for (const generatedPath of [
    "src-tauri/target/release/ocr/offline/macos-arm64",
    "src-tauri/target/release/office/macos-offline",
    "src-tauri/target/release/bundle/macos/VisualTeX.app",
    "src-tauri/target/release/bundle/dmg",
  ]) {
    rmSync(generatedPath, { recursive: true, force: true });
  }
}

const forwardedArgs = process.argv.slice(2);
const hasExplicitFeatures = forwardedArgs.some(
  (argument) => argument === "--features" || argument === "-f",
);
const releaseFeatures = hasExplicitFeatures
  ? []
  : ["--features", "tauri/custom-protocol"];
const apiOnlyOcr = process.env.VISUALTEX_API_ONLY_OCR === "1";
const macosFullOcrConfig =
  process.platform === "darwin" && !apiOnlyOcr
    ? ["--config", "src-tauri/tauri.macos.full-ocr.conf.json"]
    : [];

// macOS remembers a user's "access data from other apps" answer for the app's
// designated requirement. An ad-hoc signature's requirement is the exact build,
// so after every update macOS denied VisualTeX the Office folders without
// asking. Releases are signed with VisualTeX's own self-signed identity
// (scripts/create_macos_signing_identity.sh) instead: its requirement is the
// bundle identifier plus that certificate, the same for every release. Tauri
// imports only Apple-issued certificates itself, so the identity is put in a
// temporary keychain here and Tauri signs with it by certificate hash.
function security(args, { capture = false } = {}) {
  const result = spawnSync("security", args, {
    encoding: "utf8",
    stdio: capture ? ["ignore", "pipe", "pipe"] : ["ignore", "ignore", "pipe"],
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`security ${args[0]} failed: ${(result.stderr ?? "").trim()}`);
  }
  return result.stdout ?? "";
}

function userKeychainSearchList() {
  return security(["list-keychains", "-d", "user"], { capture: true })
    .split("\n")
    .map((line) => line.trim().replace(/^"|"$/g, ""))
    .filter(Boolean);
}

function prepareMacosSigningIdentity() {
  const directory =
    process.env.VISUALTEX_MACOS_SIGNING_DIR ??
    join(homedir(), ".config", "visualtex", "macos-signing");
  const identityFile = join(directory, "identity.p12");
  const passwordFile = join(directory, "identity.password");
  if (!existsSync(identityFile) || !existsSync(passwordFile)) {
    if (process.env.VISUALTEX_MACOS_ADHOC_SIGNING === "1") {
      process.stdout.write("Signing ad hoc (VISUALTEX_MACOS_ADHOC_SIGNING=1); not for release.\n");
      return null;
    }
    throw new Error(
      `The VisualTeX signing identity is missing from ${directory}. Restore its backup, or ` +
        "create it once with scripts/create_macos_signing_identity.sh. An ad-hoc release " +
        "loses every user's Office folder permission (VISUALTEX_MACOS_ADHOC_SIGNING=1 builds one anyway).",
    );
  }

  const workDirectory = mkdtempSync(join(tmpdir(), "visualtex-signing-"));
  const keychain = join(workDirectory, "signing.keychain-db");
  const keychainPassword = randomBytes(24).toString("hex");
  const searchList = userKeychainSearchList();
  const cleanup = () => {
    try {
      security(["list-keychains", "-d", "user", "-s", ...searchList]);
    } finally {
      try {
        security(["delete-keychain", keychain]);
      } catch {
        // Already gone; the folder is removed below either way.
      }
      rmSync(workDirectory, { recursive: true, force: true });
    }
  };
  try {
    security(["create-keychain", "-p", keychainPassword, keychain]);
    security(["set-keychain-settings", "-t", "3600", "-u", keychain]);
    security(["unlock-keychain", "-p", keychainPassword, keychain]);
    security([
      "import", identityFile, "-k", keychain, "-f", "pkcs12",
      "-P", readFileSync(passwordFile, "utf8").trim(), "-T", "/usr/bin/codesign",
    ]);
    // Lets codesign use the key without a keychain prompt.
    security([
      "set-key-partition-list", "-S", "apple-tool:,apple:,codesign:",
      "-s", "-k", keychainPassword, keychain,
    ]);
    security(["list-keychains", "-d", "user", "-s", ...searchList, keychain]);
    // Not trusted by any Mac, so only "Matching identities" lists it.
    const identity = security(["find-identity", "-p", "codesigning", keychain], { capture: true })
      .match(/\b([0-9A-F]{40})\s+"VisualTeX Code Signing"/)?.[1];
    if (!identity) throw new Error(`${identityFile} holds no "VisualTeX Code Signing" identity.`);
    process.stdout.write(`Signing with the VisualTeX identity ${identity}.\n`);
    return { identity, cleanup };
  } catch (error) {
    cleanup();
    throw error;
  }
}

const signing = process.platform === "darwin" ? prepareMacosSigningIdentity() : null;
let tauriStatus;
try {
  tauriStatus = runStatus(
    tauri,
    ["build", ...releaseFeatures, ...macosFullOcrConfig, ...forwardedArgs],
    {
      ...process.env,
      VISUALTEX_TAURI_NATIVE_PREBUILT: "1",
      ...(signing ? { APPLE_SIGNING_IDENTITY: signing.identity } : {}),
    },
  );
} finally {
  signing?.cleanup();
}
if (tauriStatus !== 0) process.exit(tauriStatus);
