# CP14 cross-host hardening complete 1

Status: `complete`.

`p15Permitted=true`.

This checkpoint freezes P14: VisualTeX core resources are byte-verified across desktop/Office hosts, resource and runtime faults fail closed and recover from the last committed source checkpoint, stale asynchronous OCR/session work cannot write into a newer target, large-document browser ownership remains bounded, and the final Windows debug companion candidate passed real Word/WebView2 create/edit acceptance without modifying the frozen Office protocol/VSTO boundary.

## Final verification

Final receipt:

`engine/target/p14-verification/2026-09-16T04-03-39-149Z/report.json`

Receipt SHA256:

`822fb4fcdb4e6188594a3c6efd97bec9de7d9baa3662da532602073da61798ae`

Verified P14 source digest:

`4c11167999ec089ea4279ed8215179a598b4dab51ee569a11812c257c050790e`

The verifier passed 24/24 gates with `sourceUnchanged=true`, `realWindowsHostAcceptance=PASS`, `frozenOfficeWorkingTreeUnchanged=true` and `p15Permitted=true`.

The verified gates include Windows TypeScript, production Office UI build, desktop core build isolation, resource consistency, desktop product smoke, Office interactive core guard, headless converter browser acceptance, corrupt-resource recovery, fatal-runtime recovery, Office race isolation, debug companion build/static-resource probe, Rust resource/fault tests, OCR browser and Native/WASM parity, browser input/accessibility, structured clipboard + zoom/DPR, 1000-formula virtual viewport, real Windows IME, browser performance, P13 100/1000 headless performance and three existing Office contract regressions.

## Resource and ABI hardening

The bundled core WASM and Latin Modern Math bytes are SHA-bound before runtime initialization. The verified WASM SHA256 is:

`6ee676b2c25a5b3b6cfe16241a208808110e9b7561514c94e4eddca343954ff7`

The verified bundled math-font SHA256 is:

`6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f`

Core-enabled desktop and Office production builds use byte-identical source assets. Corrupt WASM and corrupt font tests fail without mutating the formula source and the same host can subsequently recover with valid resources without clearing user configuration.

A fatal `EngineRuntime.dispatch()` throw is distinguished from normal structured engine errors. The Windows host invalidates the faulted runtime, preserves the last committed source checkpoint, rebuilds a fresh runtime/owner and resumes from that source. Unsupported/incomplete formula errors are not misclassified as runtime crashes.

## Windows host acceptance

Frozen record:

`engine/specs/P14-windows-host-acceptance.json`

SHA256:

`85a26917c95e8c0f516c7556dee5f99a1be85874ee067405e5ce1b5cbc1776df`

The final candidate was rebuilt through the normal Tauri debug resource pipeline so `target/debug/office/dialog/index.html` referenced the final Office bundle, not an installed/AppData copy. Real WebView2 CDP showed one VisualTeX core editor, one core input and zero MathLive fields. It loaded the final Office JS/core chunks plus `visualtex_engine_wasm_bg-DLwqUPYp.wasm` and the verified bundled math font.

The Tauri system-font command was exercised from the real WebView2 page. `Cambria Math` resolved as a bounded native font asset with an OpenType MATH table. The HTTPS page's CSP blocks the preferred `http://ipc.localhost/...` transport, but Tauri's documented postMessage fallback completed the command; P14 did not modify the frozen Office server/CSP boundary merely to suppress that warning.

Real Word Ribbon/WebView2 acceptance on an owned blank Word process created one VisualTeX OLE from source `abc`; editing that same object through `编辑所选公式`, typing `d` in the real WebView2 core input and pressing product Ctrl+S left exactly one OLE with metadata source `abcd`. Font size/render font size stayed 10.5pt and inline Word position stayed 0.

## Interaction, memory and performance

Real Windows native IME acceptance used the installed OS input method with ASCII Win32 SendInput keys, not CDP composition, Unicode injection or clipboard. Chinese preedit/commit remained one core transaction and remained correctly anchored at 50% and 200% zoom.

The browser host passed actual keyboard/pointer/clipboard/accessibility tests. Deferred OCR targets reject stale revision/selection/session/runtime identities. Structured clipboard is versioned and identity-free; untrusted HTML is not executed.

The 1000-formula virtual viewport materialized all formulas while keeping live editable DOM bounded (about 11 mounted editors at peak in the accepted run), and disposal returned active/parked session ownership to zero.

For an approximately 200-node hot formula, real browser local-edit event-to-next-frame p95 stayed below the 32ms target and cursor movement stayed below the browser-frame guard; cursor movement caused zero additional layout calls and zero document revision changes. P13 headless Office conversion performance also remained passing after P14 hardening.

## Headless Office compatibility

Interactive browser sessions keep the P10/P12 rule that a non-bundled mathematical font cannot be claimed verified when the host has no trusted system-font byte source. P14 found that this strict rule made the pure-browser P13 headless MathType regression reject its `times` output preference.

The fix is isolated to the non-interactive Office compatibility provider: when a browser-only headless converter cannot resolve a native math font, it uses the SHA-verified bundled math layout as the geometry compatibility base while retaining MathType `formulaLetterFont=times` output metadata. This does not loosen the interactive CoreMathEditor font-selection contract. Both the P13 headless converter regression and the P10 desktop product smoke pass after the change.

## Frozen Office boundary

P14 did not modify the protected WordVsto, PowerPointVsto, Windows Office Contracts, FormulaOleServer/VstoShared/WindowsOleBridge protocol components, Tauri Office server/session/state/backend/pipe files, `src/office/shared`, `src/office/api`, or `src/office/dialog/useOfficeSession.ts`. The verifier confirmed this boundary was empty at start and end.

The baseline remains:

`docs/engine-rebuild/checkpoints/CP13-office-batch-complete-1/`

The branch is `reference-3d207d7`; HEAD remains `d1f0aff8db9275c90dc18c7318dc2a4d7bbbd760`.

## Repository safety

No reset, clean, stash, branch switch, pull, rebase, merge, commit, push, installer overwrite or release action was performed while finalizing P14. Existing dirty/untracked files were preserved. Only owned test Word/companion/browser processes were closed; no user Word document was saved, closed or modified.
