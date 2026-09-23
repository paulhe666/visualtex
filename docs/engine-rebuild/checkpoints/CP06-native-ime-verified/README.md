# CP06 native IME and current-source browser checkpoint

Captured during continuation on 2026-09-13 after 14:27 +08:00. This is a non-Git source checkpoint, not an application release. Base checkout remains `reference-3d207d7` at `00b48e17b736c7e57a0a72e19e271df82906f89d`.

## Actual changes in this continuation

- `engine/tools/owned-ime-review.mjs`, `editor-browser-review.mjs`, `input-bridge-contract-review.mjs` and `svg-view-review.mjs` now compile and consume the current TypeScript host via `createSandboxAssets`, rather than trusting a possibly old `target/svg-review/modules` directory. Browser/native reports include the tested host build and module hashes where applicable.
- `engine/package.json` protection command explicitly supplies the previously authorized baseline Git transition. The failed invocation without that parameter is retained as a command error, not a source regression. No protection assertion was weakened.
- Existing core, SVG view, input bridge and authoritative editor host were reused. No competing editor implementation or production entrypoint was added.

## Verification actually executed

`node tools/build.mjs`: 200 Rust tests passed (0 failed), then real Native/WASM protocol 2151 cases, source-session 222 cases, source-edit 440 cases and editor/history 473 cases (100 history states) passed.

`node tools/sandbox-browser.mjs`: actual isolated sandbox controls build `\\frac{x_{1}+\\text{速度}}{y^{2}}` from empty input, share source/visual history, preserve invalid drafts through undo/redo, export SVG without rebuilding geometry, and preserve source during zoom/blink. PASS. This particular composition test is Chromium protocol input, not Windows IME.

`node tools/owned-ime-review.mjs`: PASS using installed Windows IME in an owned visible Edge process, ASCII Win32 SendInput keys only; no Unicode text injection, clipboard, or CDP composition. Pinyin `sudu` produces `速度`; uncommitted preedit does not change the formula; Backspace/Escape preserve committed content; Space commits once; 50% and 200% zoom keep the input proxy at the mathematical caret. IME Enter commits the IME's raw `sudu` behavior without splitting the formula or adding a second undo record. Thirty-two native/inspection steps are retained in `engine/target/owned-ime-review/report.json`. Test key mode/layout restoration was inspected. Candidate and committed screenshots are present under that same owned output directory; the candidate screenshot was actually opened and inspected.

`node tools/input-bridge-contract-review.mjs`: 15/15 synthetic fault-injection component cases pass, including failed composition/ordinary insertion retry and explicit discard. These are not native input events.

`node tools/editor-browser-review.mjs`: PASS, 44 recorded events (41 trusted browser events; 3 CDP compositionend events explicitly marked untrusted), real controls/navigation/history and accessible position text.

`node tools/svg-view-review.mjs`: PASS, 18 zoom/DPR/scroll combinations (zoom .5/1/2, DPR 1/1.25/2, scroll 0/180), stale scene/overlay rejection, non-mutating blink, invalid snapshot rejection, and one SVG/View pixel comparison with mean/max difference 0 for the fixed fixture. This is not a general cross-platform pixel identity claim.

`npm run check:protection`: PASS, 12 groups, 12 artifacts, 65 original dirty file byte checks, authorized Git transition acknowledged. No Office/source/binary hash change.

## Build provenance

Browser: `Edg/152.0.4191.62`.
Host build: `sandbox-c7109e07bd6aec2debae9872`.
Web-target WASM SHA256: `25af4f2fdfae910fe8a9749d0ffbce90b0c5462906debfdb48e1b7852a100bea`.
Native CLI SHA256: `d72a03762002f17dc65e357d6cbfc0c953ecb5ca4c9e716d5fabb73b116af0e1`.
Node-target WASM SHA256: `8eead4484c4ec7c805e01e0e2b03034044c6b859a3813e1a7f0516bedbbd4600` (bindgen target differs from browser).
Host module hashes: editor-host `3965eb53a5a9b6b5c2df95b1e796d6e3b639a87f9a49f6dda9d744166ee44194`; input-bridge `e441f303d66181ebdab531457c1bc5baaa92a82b2564a873c02a5d32f479435f`; svg-view `582fce7cb5d2d93c85cdca00f353a32d2a8b4df6e9c0f5eedd61b5a5cbc4289a`.
Fixed font hashes remain in the generated test reports. Font binaries are not included in this source archive.

## Gate and limitations

K0 autonomous editing/layout/SVG and Windows Edge native-IME closure is verified sufficiently to continue isolated P07 structure work. This is not complete MathLive replacement. Formal main application, Office window, hidden converter, installer, WebView2 and macOS/WPS paths are not switched or newly accepted. Command candidate/macro coverage, matrices and broader structures remain later-stage work. No Git add/commit/push or package installation/release was performed. Original user documents were not saved or closed. Only owned test browser instances were closed.
