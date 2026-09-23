# CP10-main-app-tauri-stable-1

Status: `tauriStable-notP10Complete`.

This checkpoint advances the prior `CP10-main-app-browser-stable-1` state by completing the previously blocked real Windows Tauri/WebView2 close-and-relaunch acceptance. It does **not** mark P10 complete.

## Source baseline

The product/core source remains the exact source frozen by `CP10-main-app-browser-stable-1`:

- P10 browser receipt: `engine/target/p10-verification/2026-09-14T14-24-07-439Z/report.json`
- verified source SHA256: `f39d8f3dbb59e39042a91f8d2aed5db30619d86fe47187d4012cd1f5effa5825`
- source archive: `docs/engine-rebuild/checkpoints/CP10-main-app-browser-stable-1/visualtex-p10-main-app-browser-stable-1.tgz`
- archive SHA256: `b057aeb5097885211411117d5ff7c08288f42272e8ead53f615469928d30a21a`
- current audit after the Tauri acceptance: PASS, 267/267 verified inputs still match, `currentSourceStillMatches=true`.

The repository HEAD observed during this acceptance is `d1f0aff8db9275c90dc18c7318dc2a4d7bbbd760` on `reference-3d207d7`, not the earlier `00b48e17...` handoff HEAD. That new commit is the external parallel Office/VSTO commit `fix(office): unify inline baseline handling`; this task did not reset, amend, commit, clean, stash, merge, rebase, pull or otherwise alter it. The P10 source archive audit proves that the 267 P10-owned verified inputs remain unchanged.

## Real Tauri acceptance

Runtime receipt: `engine/target/p10-tauri-relaunch/2026-09-15T20-00-46+08-00/report.json`.

The candidate was built with `VITE_VISUALTEX_CORE_EDITOR=1` and a debug Tauri `--no-bundle` build using isolated identifier `com.visualtex.studio.p10acceptance`, so user-installed VisualTeX WebView2/localStorage/config data was not reused or overwritten. No NSIS was built.

The first real Tauri launch loaded `http://tauri.localhost/` in WebView2 and mounted `visualtex-core`; MathLive field count was 0, the core WASM loaded, and the bundled Latin Modern Math asset loaded. Real CDP keyboard input changed the persisted source from the Gaussian integral baseline to the same source plus `+p10r`. The normal product Save button produced a `.visualtex.json` whose first formula source exactly matched the persisted core store. The product `apply_app_window_configuration` path then changed the main inner window size from 1182×728 to 980×620 without changing formula source or engine selection. PID 77808 closed cleanly.

A second fresh Tauri process (PID 76860) was launched from the same candidate. It again selected `visualtex-core` from session preflight, still had 0 MathLive fields, reloaded the WASM and bundled math font, restored the same line id `4f71274c-67ba-4042-a532-d61e63641853`, restored the edited `+p10r` source exactly, and restored the main inner window size to 980×620. The second process also closed cleanly. No VisualTeX main process was left running.

## Safety boundary

No Word document was opened, saved, closed or modified. No Office/VSTO/Office Session action was performed. The existing external Office commit and Office evidence files were not edited or cleaned. The only acceptance helper created by this task is under ignored/generated `engine/target/` and is not part of the P10 source archive.

## Remaining P10 gate

P10 is still **not complete**. The only explicit remaining external gate is installer-level no-OCR size/install verification. Do not build or install a no-OCR NSIS solely to close P10 without an explicit user instruction. Until that gate is completed and reviewed, do not enter P11 `OfficeCompatibilityAdapter` and do not modify Office wire DTOs, Session ownership, C# WordFormulaService or Word document algorithms.
