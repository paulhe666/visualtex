import { spawnSync } from "node:child_process";
import { resolve } from "node:path";
import process from "node:process";

if (process.platform !== "win32") {
  throw new Error("The Windows Rust test runner must be executed on Windows.");
}

// The VisualTeX library tests exercise Office companion HTTP routes that pull
// Tauri's Windows UI stack into Rust's libtest executable. Unlike the packaged
// Tauri executable, libtest has no Common Controls v6 manifest by default, so
// Windows loads comctl32 5.82 where TaskDialogIndirect does not exist and exits
// before main with STATUS_ENTRYPOINT_NOT_FOUND (0xc0000139). Embed the same v6
// dependency at link time for test artifacts only; production Tauri builds keep
// their normal manifest/resource pipeline unchanged.
const manifest = resolve("src-tauri/tests/common-controls-v6.manifest").replaceAll("\\", "/");
const manifestFlags = [
  "-Clink-arg=/MANIFEST:EMBED",
  `-Clink-arg=/MANIFESTINPUT:${manifest}`,
];
const encodedRustFlags = [
  process.env.CARGO_ENCODED_RUSTFLAGS,
  ...manifestFlags,
]
  .filter(Boolean)
  .join("\x1f");

const cargo = "cargo.exe";
const args = [
  "test",
  "--manifest-path",
  "src-tauri/Cargo.toml",
  "--lib",
  "--no-fail-fast",
  ...process.argv.slice(2),
];
const result = spawnSync(cargo, args, {
  stdio: "inherit",
  shell: false,
  env: {
    ...process.env,
    CARGO_ENCODED_RUSTFLAGS: encodedRustFlags,
  },
});

if (result.error) throw result.error;
process.exit(result.status ?? 1);
