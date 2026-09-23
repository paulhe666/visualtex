// Isolated compiler driver. Only compiler-generated outputs below engine/target
// and engine/web/pkg* are produced; production projects are never built here.
import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const ext = process.platform === 'win32' ? '.exe' : '';
const env = { ...process.env, CARGO_TARGET_DIR: resolve(root, 'target') };
function run(binary, args) {
  const child = spawnSync(binary, args, { cwd: root, env, stdio: 'inherit' });
  if (child.error) throw child.error;
  if (child.status !== 0) throw new Error(`${binary} failed with status ${child.status}`);
}
const bindgen = resolve(root, `.tools/bin/wasm-bindgen${ext}`);
if (!existsSync(bindgen)) throw new Error('Install pinned wasm-bindgen-cli 0.2.105 into engine/.tools first; see README.md');
run(process.execPath, ['tools/prepare-symbol-test-font.mjs']);
run('cargo', ['+1.96.0', 'test', '--workspace', '--locked']);
run('cargo', ['+1.96.0', 'build', '-p', 'visualtex-engine-test-cli', '--locked']);
run('cargo', ['+1.96.0', 'build', '-p', 'visualtex-engine-wasm', '--target', 'wasm32-unknown-unknown', '--locked']);
const wasm = resolve(root, 'target/wasm32-unknown-unknown/debug/visualtex_engine_wasm.wasm');
run(bindgen, [wasm, '--out-dir', 'web/pkg-node', '--target', 'nodejs']);
run(bindgen, [wasm, '--out-dir', 'web/pkg', '--target', 'web']);
run(process.execPath, ['tools/protocol-parity.mjs']);
run(process.execPath, ['tools/source-session-parity.mjs']);
run(process.execPath, ['tools/source-edit-parity.mjs']);
run(process.execPath, ['tools/editor-parity.mjs']);
run(process.execPath, ['tools/layout-parity.mjs']);
run(process.execPath, ['tools/symbol-parity.mjs']);
