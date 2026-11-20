// Builds visualtex-mcp.exe, the stdio MCP server agents start (docs/office-mcp/DESIGN.md §4.1),
// and stages it as the Tauri sidecar src-tauri/binaries/visualtex-mcp-<triple>.exe; the
// installer places it beside visualtex.exe.
import { spawnSync } from "node:child_process";
import { copyFileSync, mkdirSync, statSync } from "node:fs";
import { resolve } from "node:path";
import process from "node:process";

if (process.platform !== "win32") {
  throw new Error("visualtex-mcp.exe is built on Windows.");
}
const manifest = resolve("mcp-launcher", "Cargo.toml");
const result = spawnSync(
  "cargo.exe",
  ["build", "--release", "--locked", "--manifest-path", manifest],
  { stdio: "inherit", shell: false },
);
if (result.error) throw result.error;
if (result.status !== 0) throw new Error(`visualtex-mcp build failed with ${result.status}`);
const built = resolve("mcp-launcher", "target", "release", "visualtex-mcp.exe");
const binaries = resolve("src-tauri", "binaries");
mkdirSync(binaries, { recursive: true });
const staged = resolve(binaries, "visualtex-mcp-x86_64-pc-windows-msvc.exe");
copyFileSync(built, staged);
console.log(`visualtex-mcp sidecar: ${staged} (${statSync(staged).size} bytes)`);
