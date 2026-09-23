# P14 — Cross-host, resource, performance and fault hardening

Status: complete for the verified Windows scope on 2026-09-16. `P15` may start after `engine/tools/p14-verify.mjs` passes.

## Boundary

P14 hardens the autonomous core after the P13 Office batch switch. It does not redesign the Office Session protocol, Word/PowerPoint document algorithms, numbering/reference logic, COM/OLE contracts, private MathType runtime or macOS Office integration. The protected Office/VSTO/server paths remain frozen. macOS/WKWebView is explicitly untested and unchanged in this stage.

The verified Windows host set is:

- Native Rust core tests and Native/WASM parity runners.
- Owned Microsoft Edge browser host using the same WASM/font bytes and current TypeScript host modules.
- Core-enabled Windows desktop production bundle in an actual product page browser smoke.
- Real Windows WebView2 hosted by the `--office-background` VisualTeX companion and opened from the real Word VisualTeX Ribbon.

No browser result is described as macOS, installed-package or final P16 acceptance.

## Resource identity and ABI hardening

`apps/windows/src/engine/coreEngineRuntime.ts` now verifies both the core WASM and the bundled mathematical font by SHA-256 before registration/instantiation. The fixed candidate identities are:

- `visualtex_engine_wasm_bg-DLwqUPYp.wasm`: `6ee676b2c25a5b3b6cfe16241a208808110e9b7561514c94e4eddca343954ff7`.
- `latinmodern-math-CUFlBgaJ.otf`: `6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f`.

Core-enabled desktop and Office production bundles contain byte-identical copies of both resources. `p14_core_resource_consistency.mjs` also verifies that each host has exactly one current hashed core WASM/font asset, that generated entry/chunk references point to those assets, MIME expectations are explicit and a path containing spaces and non-ASCII text can read the same bytes.

`p14_companion_static_resource_probe.mjs` starts only its own debug companion and proves the real HTTPS `/assets` route serves the production WASM as `application/wasm`, the font as a valid SFNT MIME, preserves hashed-asset cacheability, emits `nosniff`/same-origin resource policy and serves bytes identical to `dist-office-windows-native`.

`p14_core_resource_recovery_browser.mjs` uses the production Office bundle to inject a bad WASM generation and a bad font generation. Both fail closed without modifying the formula source; restoring the correct resources in the same page succeeds without deleting configuration, clearing WebView data or reloading the page.

The plain-browser interactive editor still treats non-bundled system-font preferences as unresolved because that host has no trusted access to installed font bytes. The non-interactive Office conversion compatibility layer is deliberately narrower: when a plain-browser acceptance host cannot resolve a MathType output preference such as `times`, conversion geometry uses the bundled SHA-verified math font while the Office output metadata remains `formulaLetterFont=times`. A real Tauri/WebView2 converter still resolves a verified system mathematical face first and uses the same bundled font only as the P12 fallback. This keeps P10 interactive font semantics strict without making the P13 headless compatibility verifier depend on an OS-font bridge it does not own.

## Runtime fault recovery

Business/protocol errors remain ordinary engine errors and must not invalidate the runtime. Only an actual exception escaping `EngineRuntime.dispatch()` is classified as a fatal runtime fault. The Windows runtime wrapper retires that instance and clears the shared runtime cache. The core editor preserves the latest successfully committed source in the existing product store, disposes the faulted host and rebuilds a fresh runtime/owner from that checkpoint. The recovered formula retains source and can again produce SVG and Presentation MathML. Core undo history restarts at the preserved checkpoint after a fatal instance rebuild; no false claim is made that an in-memory WASM history survives a trap.

`p14_core_runtime_recovery_browser.mjs` proves that an ordinary unsupported command does not retire the runtime, a forced instance invalidation is recognized as fatal, the old owner still exposes its last snapshot, and a new runtime recreates the same source and layout/export without user configuration deletion.

## Race and asynchronous ownership hardening

`p14_office_core_race_browser.mjs` uses the production Office bundle and proves:

- two interactive windows keep independent Session/source ownership;
- parallel edits never cross-write;
- a double Apply produces one commit and one close;
- native close while PNG export is delayed waits for the current export;
- no stale export PATCH occurs after Session close.

Deferred OCR ownership is revalidated through native core tests, Native/WASM parity and real browser input. Captured targets are one-shot and bound to the owner, revision, selection epoch and request identity. Selection movement, source edits followed by undo, a newer request, cancellation, IME start/cancel, parking/destroy/reopen, cloned tickets, cross-owner and cross-runtime delivery all prevent stale OCR insertion. Invalid OCR output fails atomically and retains the source.

## Input, clipboard, zoom and accessibility

The current host is exercised with real browser keyboard/pointer/clipboard events. Chinese composition commits exactly once; candidate Enter does not split a formula; grapheme deletion respects Chinese characters; Ctrl+Z/Y restores the same core transaction; the input proxy exposes a meaningful mathematical `aria-label` path/offset description.

`owned-ime-review.mjs` independently uses an owned visible Edge process and the installed Windows input method. It sends only ASCII pinyin keys through Win32 `SendInput`; it does not use CDP composition, Unicode injection or clipboard to claim native IME success. The Windows IME path passes at 50% and 200% zoom.

Structured matrix clipboard uses real trusted copy/cut/paste events. Identity-bearing core nodes are not copied into the external fragment; wrong-size paste fails atomically. Nine zoom/DPR cases retain hit ownership.

The 1000-formula virtual viewport mounts only the necessary editor DOM. The full scroll materializes every formula while the editable DOM remains bounded; offscreen owners park and later restore exact history/identity. After disposal the active/parked Session counts and editor DOM return to zero.

## Performance and memory evidence

`p14-browser-performance-review.mjs` measures an approximately 200-node formula in an owned real Edge host. For the 187-node fixture:

- local edit event → next animation frame p95: `15.5213 ms` (target `<=32 ms`);
- cursor event → next animation frame p95: `5.4880 ms`;
- 100 cursor moves add exactly zero layout calls and zero document revisions;
- JS heap returns below its starting measured value after host disposal, and active Sessions return to zero.

The cursor measurement includes browser event dispatch and the next frame, so it is not presented as a direct proof of the core-only `<=8 ms` target. The stronger invariant is measured directly: cursor movement performs no layout and no content revision.

P13 headless production-bundle performance is rerun after P14 changes: 100 mixed formulas complete with zero failures in 1104 ms, and 1000 OMML formulas complete with zero failures in 1935 ms while remaining MathML-only and reusing one core WASM/font load.

## Real Windows WebView2 acceptance

A fresh debug Tauri build refreshed `target/debug/office` from the current `dist-office-windows-native` through the declared Tauri resource mapping. No installed `%LOCALAPPDATA%` Office UI was overwritten.

A dedicated owned Word process opened the current native VisualTeX WebView2 editor through the real Ribbon. Test-only WebView2 remote debugging was enabled only for read-only DOM/resource inspection and deterministic input focus. The real page reported one core editor/input, zero MathLive fields and loaded the current hashed Office JS/WASM/font assets. The external HTTPS WebView CSP blocks Tauri's preferred `http://ipc.localhost` custom-protocol fetch, after which Tauri falls back to its postMessage transport; a read-only `Cambria Math` system-font command succeeded through that fallback and returned a valid MATH face. The frozen Office server/CSP is intentionally unchanged.

Source mutation and Apply used OS `SendKeys` into the exact foreground WebView2 editor HWND. A new inline VisualTeX OLE was created from source `abc`; editing the same object through the real `Edit Selected Formula` Ribbon restored the source, appended `d`, and updated the same OLE to `abcd` without inserting a duplicate. Final metadata retained 10.5 pt source/render size, inline position 0 and core-derived natural geometry/baseline.

The machine-readable record is `engine/specs/P14-windows-host-acceptance.json`.

## Historical Tao crash observation

P13 recorded three non-deterministic debug-companion failures at Windows exception `0xc000001d`, disassembled to a Tao `EventLoopRunner` `Rc` clone in native `Window::new`. P14 could not reproduce a matching crash under exact 1000 converter batching, lifecycle stress, concurrent Session work, resource failure/retry, race tests, repeated production browser tests or the final real WebView2 create/edit acceptance. The evidence does not support attributing those historical failures to formula parsing, MathML, Session ordering or the P13 core provider. It remains a diagnostic runtime observation rather than a silently claimed root-cause fix.

## P15 gate

P15 may start only after the P14 verifier passes with `sourceUnchanged=true` and the frozen Office working-tree snapshot unchanged. P15 is responsible for removing MathLive/MathJax from formal production runtime paths; P14 intentionally keeps migration-era legacy dependencies where P15 has not yet removed them.
