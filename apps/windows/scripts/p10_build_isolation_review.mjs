import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readdirSync } from "node:fs";
import process from "node:process";

const npm = process.platform === "win32"
  ? { executable: process.env.ComSpec || "cmd.exe", prefix: ["/d", "/s", "/c", "npm"] }
  : { executable: "npm", prefix: [] };
function runCoreBuild(enabled) {
  const env = { ...process.env };
  if (enabled) env.VITE_VISUALTEX_CORE_EDITOR = "1";
  else delete env.VITE_VISUALTEX_CORE_EDITOR;
  const result = spawnSync(npm.executable, [...npm.prefix, "run", "build:desktop"], {
    cwd: process.cwd(),
    env,
    encoding: "utf8",
    timeout: 120_000,
    maxBuffer: 32 * 1024 * 1024,
  });
  if (result.status !== 0 || result.error) {
    throw new Error(`${enabled ? "core" : "legacy"} build failed\n${result.stdout ?? ""}\n${result.stderr ?? ""}\n${result.error?.message ?? ""}`);
  }
  const assets = readdirSync("dist/assets").sort();
  return {
    assets,
    core: assets.filter((name) =>
      /visualtex_engine_wasm|latinmodern-math|CoreMathEditor|coreStaticPreview/u.test(name),
    ),
  };
}

const legacy = runCoreBuild(false);
assert.deepEqual(legacy.core, [], "Default legacy build must not carry candidate core resources");
const candidate = runCoreBuild(true);
assert.ok(candidate.core.some((name) => /visualtex_engine_wasm.*\.wasm$/u.test(name)), "Candidate build is missing engine WASM");
assert.ok(candidate.core.some((name) => /latinmodern-math.*\.otf$/u.test(name)), "Candidate build is missing verified math font");
assert.ok(candidate.core.some((name) => /^CoreMathEditor-.*\.js$/u.test(name)), "Candidate build is missing lazy core editor chunk");
console.log(JSON.stringify({
  status: "PASS",
  defaultBuildCoreAssets: legacy.core,
  candidateBuildCoreAssets: candidate.core,
  defaultResourceIsolation: true,
  candidateResourcesExplicitOnly: true,
}, null, 2));
