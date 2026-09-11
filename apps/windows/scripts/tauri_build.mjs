import { spawnSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import process from "node:process";
import { windowsPowerShellPath } from "./windows_powershell.mjs";

const require = createRequire(import.meta.url);

function run(command, args, env = process.env) {
  const isWindowsCmd =
    process.platform === "win32" && command.toLowerCase().endsWith(".cmd");
  const executable = isWindowsCmd ? (process.env.ComSpec ?? "cmd.exe") : command;
  const executableArgs = isWindowsCmd
    ? ["/d", "/s", "/c", command, ...args]
    : args;
  console.log(`\n> ${command} ${args.join(" ")}`);
  const result = spawnSync(executable, executableArgs, {
    stdio: "inherit",
    shell: false,
    env,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}

function preparePatchedNsisTemplate({ noOcrBundle = false } = {}) {
  const nativePackage =
    process.arch === "arm64"
      ? "@tauri-apps/cli-win32-arm64-msvc"
      : "@tauri-apps/cli-win32-x64-msvc";
  const nativeCliPath = require.resolve(nativePackage);
  const binary = readFileSync(nativeCliPath);

  let template;
  for (const newline of ["\r\n", "\n"]) {
    const startMarker = Buffer.from(`Unicode true${newline}ManifestDPIAware true`);
    const endMarker = Buffer.from(
      `  !insertmacro SetLnkAppUserModelId "$DESKTOP\\\${PRODUCTNAME}.lnk"${newline}FunctionEnd${newline}`,
    );
    const start = binary.indexOf(startMarker);
    const endStart = start >= 0 ? binary.indexOf(endMarker, start) : -1;
    if (start >= 0 && endStart >= 0) {
      template = binary
        .subarray(start, endStart + endMarker.length)
        .toString("utf8")
        .replace(/\r\n/g, "\n");
      break;
    }
  }
  if (!template) {
    throw new Error(
      `Unable to extract the exact NSIS installer template from ${nativeCliPath}`,
    );
  }

  if (noOcrBundle) {
    const hookInclude = '!include "{{installer_hooks}}"';
    if (!template.includes(hookInclude)) {
      throw new Error("The Tauri NSIS installer-hooks include changed unexpectedly");
    }
    template = template.replace(
      hookInclude,
      `!define VISUALTEX_NO_OCR_BUNDLE\n${hookInclude}`,
    );
  }

  const functionStart = "Function PageReinstall\n";
  if (template.split(functionStart).length !== 2) {
    throw new Error("The Tauri NSIS PageReinstall function changed unexpectedly");
  }
  const acceptanceStart = [
    "Function PageReinstall",
    "  ; Installed-release acceptance uses a clean custom directory and must not",
    "  ; enter maintenance mode for another installed VisualTeX copy.",
    "  ${GetParameters} $R7",
    "  ClearErrors",
    '  ${GetOptions} $R7 "/VISUALTEXACCEPTANCE" $R8',
    "  ${IfNot} ${Errors}",
    "    Abort",
    "  ${EndIf}",
    "",
  ].join("\n");
  template = template.replace(functionStart, acceptanceStart);

  const oldSelection = [
    "    ; Check the first radio button if this the first time",
    "    ; we enter this page or if the second button wasn't",
    "    ; selected the last time we were on this page",
    "    ${If} $ReinstallPageCheck <> 2",
    "      SendMessage $R2 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "    ${Else}",
    "      SendMessage $R3 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "    ${EndIf}",
    "",
    "    ${NSD_SetFocus} $R2",
  ].join("\n");
  const newSelection = [
    '    ; Same-version maintenance defaults to the second option, "Uninstall',
    '    ; VisualTeX". Preserve an explicit user selection when navigating back.',
    "    ; Upgrade/downgrade pages keep Tauri's original first-option default.",
    "    ${If} $R0 = 0",
    "      ${If} $ReinstallPageCheck = 1",
    "        SendMessage $R2 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "        ${NSD_SetFocus} $R2",
    "      ${Else}",
    "        SendMessage $R3 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "        StrCpy $ReinstallPageCheck 2",
    "        ${NSD_SetFocus} $R3",
    "      ${EndIf}",
    "    ${Else}",
    "      ${If} $ReinstallPageCheck <> 2",
    "        SendMessage $R2 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "        ${NSD_SetFocus} $R2",
    "      ${Else}",
    "        SendMessage $R3 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "        ${NSD_SetFocus} $R3",
    "      ${EndIf}",
    "    ${EndIf}",
  ].join("\n");
  if (!template.includes(oldSelection)) {
    throw new Error(
      "The Tauri NSIS maintenance selection block changed unexpectedly; refusing an unverified installer build",
    );
  }
  template = template.replace(oldSelection, newSelection);

  const oldSameVersionLeave = [
    "  ${If} $R0 = 0 ; Same version, proceed",
    "    ${If} $R1 = 1              ; User chose to add/reinstall",
    "      Goto reinst_done",
    "    ${Else}                    ; User chose to uninstall",
    "      Goto reinst_uninstall",
    "    ${EndIf}",
  ].join("\n");
  const newSameVersionLeave = [
    "  ${If} $R0 = 0 ; Same version: always remove the installed payload before reinstalling.",
    "    ; Tauri's in-place same-version path can leave the previous EXE/resources",
    "    ; untouched while still running post-install hooks from the stale install.",
    "    ; Force a real uninstall so this installer writes its exact payload.",
    "    Goto reinst_uninstall",
  ].join("\n");
  if (!template.includes(oldSameVersionLeave)) {
    throw new Error(
      "The Tauri NSIS same-version PageLeaveReinstall block changed unexpectedly; refusing an installer that can retain stale payload files",
    );
  }
  template = template.replace(oldSameVersionLeave, newSameVersionLeave);

  const oldMaintenancePageAnchor = [
    "; 4. Custom page to ask user if he wants to reinstall/uninstall",
    ";    only if a previous installation was detected",
    "Var ReinstallPageCheck",
    "Page custom PageReinstall PageLeaveReinstall",
  ].join("\n");
  const newMaintenancePageAnchor = [
    "; 4. VisualTeX native Office integration choice and Office-process guard",
    ";    This must run before Tauri maintenance can uninstall an existing build.",
    "Page custom VisualTeXOfficePageCreate VisualTeXOfficePageLeave",
    "",
    "; 5. Custom page to ask user if he wants to reinstall/uninstall",
    ";    only if a previous installation was detected",
    "Var ReinstallPageCheck",
    "Page custom PageReinstall PageLeaveReinstall",
  ].join("\n");
  if (!template.includes(oldMaintenancePageAnchor)) {
    throw new Error(
      "The Tauri NSIS maintenance-page anchor changed unexpectedly; refusing a build where Office may be uninstalled before the running-Office guard",
    );
  }
  template = template.replace(oldMaintenancePageAnchor, newMaintenancePageAnchor);

  const oldInstallerPageSequence = [
    "; 5. Choose install directory page",
    "!define MUI_PAGE_CUSTOMFUNCTION_PRE SkipIfPassive",
    "!insertmacro MUI_PAGE_DIRECTORY",
    "",
    "; 6. Start menu shortcut page",
    "Var AppStartMenuFolder",
  ].join("\n");
  const newInstallerPageSequence = noOcrBundle
    ? [
        "; 6. Choose install directory page",
        "!define MUI_PAGE_CUSTOMFUNCTION_PRE SkipIfPassive",
        "!insertmacro MUI_PAGE_DIRECTORY",
        "",
        "; 7. Start menu shortcut page",
        "Var AppStartMenuFolder",
      ].join("\n")
    : [
        "; 6. Choose install directory page",
        "!define MUI_PAGE_CUSTOMFUNCTION_PRE SkipIfPassive",
        "!insertmacro MUI_PAGE_DIRECTORY",
        "",
        "; 7. VisualTeX optional offline OCR resources choice",
        "Page custom VisualTeXOcrPageCreate VisualTeXOcrPageLeave",
        "",
        "; 8. Start menu shortcut page",
        "Var AppStartMenuFolder",
      ].join("\n");
  if (!template.includes(oldInstallerPageSequence)) {
    throw new Error(
      "The Tauri NSIS installer page sequence changed unexpectedly; refusing a build where the VisualTeX OCR choice might be hidden",
    );
  }
  template = template.replace(oldInstallerPageSequence, newInstallerPageSequence);

  const oldResourceCopy = [
    "  ; Copy resources",
    "  {{#each resources_dirs}}",
    '    CreateDirectory "$INSTDIR\\\\{{this}}"',
    "  {{/each}}",
    "  {{#each resources}}",
    '    File /a "/oname={{this.[1]}}" "{{no-escape @key}}"',
    "  {{/each}}",
  ].join("\n");
  const newResourceCopy = [
    "  ; Copy resources through VisualTeX hooks so optional OCR payloads can be",
    "  ; skipped before extraction while all non-OCR Tauri resources remain exact.",
    "  {{#each resources_dirs}}",
    '    !insertmacro VisualTeXCreateBundledResourceDirectory "{{this}}"',
    "  {{/each}}",
    "  {{#each resources}}",
    '    !insertmacro VisualTeXInstallBundledResource "{{this.[1]}}" "{{no-escape @key}}"',
    "  {{/each}}",
  ].join("\n");
  if (!template.includes(oldResourceCopy)) {
    throw new Error(
      "The Tauri NSIS resource copy block changed unexpectedly; refusing an installer where optional OCR extraction cannot be verified",
    );
  }
  template = template.replace(oldResourceCopy, newResourceCopy);

  const output = resolve(
    "src-tauri",
    noOcrBundle ? "target-no-ocr" : "target",
    "nsis-template",
    "visualtex-installer.nsi",
  );
  mkdirSync(dirname(output), { recursive: true });
  writeFileSync(output, template, "utf8");
  console.log(`Prepared verified VisualTeX NSIS template: ${output}`);
}

function releaseArguments(args) {
  const result = [];
  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--debug" || argument === "-d") {
      throw new Error("VisualTeX Windows installer builds must use release mode");
    }
    if (argument === "--no-bundle" || argument === "--no-ocr-bundle") continue;
    if (argument === "--bundles" || argument === "-b") {
      index += 1;
      continue;
    }
    if (argument.startsWith("--bundles=")) continue;
    result.push(argument);
  }
  return result;
}

const npm = process.platform === "win32" ? "npm.cmd" : "npm";
const tauri = process.platform === "win32" ? "tauri.cmd" : "tauri";
const node = process.execPath;
const powershell = process.platform === "win32" ? windowsPowerShellPath() : "powershell";
const noOcrBundle =
  process.argv.includes("--no-ocr-bundle") ||
  process.env.VISUALTEX_NO_OCR_BUNDLE === "1";
if (process.argv.includes("--prepare-nsis-template-only")) {
  if (process.platform !== "win32") {
    throw new Error("VisualTeX NSIS template preparation is Windows-only");
  }
  preparePatchedNsisTemplate({ noOcrBundle });
  process.exit(0);
}
const forwarded = releaseArguments(process.argv.slice(2));

if (process.platform !== "win32") {
  if (noOcrBundle) {
    throw new Error("The no-OCR installer flavor is Windows-only");
  }
  run(npm, ["run", "build:desktop"]);
  run(node, ["scripts/verify_frontend_dist.mjs"]);
  run(tauri, ["build", ...forwarded]);
  process.exit(0);
}

const noOcrTargetRoot = resolve("src-tauri", "target-no-ocr");
const targetRoot = noOcrBundle ? noOcrTargetRoot : resolve("src-tauri", "target");
const flavorEnvironment = noOcrBundle
  ? {
      ...process.env,
      VISUALTEX_NO_OCR_BUNDLE: "1",
      CARGO_TARGET_DIR: noOcrTargetRoot,
    }
  : process.env;
const flavorConfigArguments = noOcrBundle
  ? ["--config", "src-tauri/tauri.no-ocr.conf.json"]
  : ["--config", "src-tauri/tauri.ocr-bundle.conf.json"];

// 1. Remove only stale outputs for this flavor. The lightweight build uses its
// own Cargo/Tauri target tree so it cannot overwrite the already-accepted full
// installer or reuse its generated resource table.
const cleanArguments = ["scripts/clean_windows_release_outputs.mjs"];
if (noOcrBundle) cleanArguments.push("--target-dir", "src-tauri/target-no-ocr");
run(node, cleanArguments, flavorEnvironment);

// 2. Prepare the exact Tauri-version NSIS template and native bundle inputs.
// The lightweight flavor deliberately skips private OCR Python preparation.
preparePatchedNsisTemplate({ noOcrBundle });
run(npm, ["run", "build:bundle"], flavorEnvironment);

// 3. Build the main frontend exactly once, then validate every referenced asset
// before Tauri build.rs/codegen is allowed to execute.
run(npm, ["run", "build:desktop"], flavorEnvironment);
run(node, ["scripts/verify_frontend_dist.mjs"], flavorEnvironment);

// 4. Build the Rust application without bundling. beforeBuildCommand performs
// validation only and therefore cannot rebuild or replace dist.
run(
  tauri,
  ["build", "--no-bundle", ...flavorConfigArguments, ...forwarded],
  {
    ...flavorEnvironment,
    VISUALTEX_FRONTEND_PREBUILT: "1",
  },
);
const builtExecutable = resolve(targetRoot, "release", "visualtex.exe");
run(
  node,
  ["scripts/verify_embedded_frontend_assets.mjs", "--exe", builtExecutable],
  flavorEnvironment,
);

// 5. Bundle directly with the verified custom NSIS template. This avoids
// depending on Tauri's ephemeral generated installer.nsi after makensis exits.
run(
  tauri,
  ["bundle", "--bundles", "nsis", ...flavorConfigArguments, ...forwarded],
  {
    ...flavorEnvironment,
    VISUALTEX_FRONTEND_PREBUILT: "1",
  },
);

const baseConfig = JSON.parse(readFileSync(resolve("src-tauri", "tauri.conf.json"), "utf8"));
const version = String(baseConfig.version ?? "").trim();
if (!version) throw new Error("VisualTeX version is missing from tauri.conf.json");
const generatedInstaller = resolve(
  targetRoot,
  "release",
  "bundle",
  "nsis",
  `VisualTeX_${version}_x64-setup.exe`,
);
if (!existsSync(generatedInstaller)) {
  throw new Error(`Expected NSIS installer is missing: ${generatedInstaller}`);
}

// 6. Verify the correct flavor and run the clean-directory installed-runtime
// smoke. For the lightweight flavor, publish a distinct local artifact name only
// after verification; the accepted full installer remains untouched.
if (noOcrBundle) {
  run(
    powershell,
    [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      "scripts/verify_windows_no_ocr_installer.ps1",
      "-InstallerPath",
      generatedInstaller,
    ],
    flavorEnvironment,
  );
  if (process.env.VISUALTEX_SKIP_INSTALL_SMOKE !== "1") {
    run(
      powershell,
      [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        "scripts/test_windows_installed_release.ps1",
        "-InstallerPath",
        generatedInstaller,
      ],
      flavorEnvironment,
    );
  }
  const deliveryDirectory = resolve("src-tauri", "target", "release", "bundle", "nsis");
  const deliveryInstaller = resolve(
    deliveryDirectory,
    `VisualTeX_${version}_x64-no-ocr-setup.exe`,
  );
  mkdirSync(deliveryDirectory, { recursive: true });
  copyFileSync(generatedInstaller, deliveryInstaller);
  console.log(`No-OCR installer copied without replacing the full installer: ${deliveryInstaller}`);
} else {
  run(powershell, [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    "scripts/verify_windows_release_artifacts.ps1",
  ]);
  if (process.env.VISUALTEX_SKIP_INSTALL_SMOKE !== "1") {
    run(powershell, [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      "scripts/test_windows_installed_release.ps1",
    ]);
  }
}
