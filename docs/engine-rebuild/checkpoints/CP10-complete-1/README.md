# CP10-complete-1

Status: **P10 COMPLETE**

This checkpoint closes P10 without changing the already verified P10 product/core source. It composes the frozen browser/main-app source checkpoint with the two external gates that were still outstanding.

## Frozen source

- Source receipt: `engine/target/p10-verification/2026-09-14T14-24-07-439Z/report.json`
- Source status: 15/15 PASS, `sourceUnchanged=true`
- Source SHA256: `f39d8f3dbb59e39042a91f8d2aed5db30619d86fe47187d4012cd1f5effa5825`
- Frozen archive: `docs/engine-rebuild/checkpoints/CP10-main-app-browser-stable-1/visualtex-p10-main-app-browser-stable-1.tgz`
- Archive SHA256: `b057aeb5097885211411117d5ff7c08288f42272e8ead53f615469928d30a21a`
- Archive audit after installer acceptance: PASS, 267/267 current source files still match, forbidden archive files `[]`.

The P10 spec/verifier source itself was intentionally restored to the frozen bytes after a status-only edit caused the old P09 receipt audit to reject the changed verifier/spec. P10 completion is therefore recorded in this checkpoint rather than mutating the already audited source set.

## External gate 1 — real Tauri/WebView2 close/relaunch

Receipt: `engine/target/p10-tauri-relaunch/2026-09-15T20-00-46+08-00/report.json`

PASS:

- real Tauri `http://tauri.localhost/` WebView2 launched with the core candidate;
- `data-engine="visualtex-core"` present and MathLive fields = 0;
- real keyboard edit changed the persisted formula source to include `+p10r`;
- product Save emitted a `.visualtex.json` whose LaTeX matched the core/store source exactly;
- product window configuration changed the main inner size from `1182×728` to `980×620`;
- process was closed normally and a second process was launched;
- second launch restored the same formula line/source, selected core again, loaded WASM/font assets, kept MathLive fields at 0, and restored `980×620`;
- acceptance processes were closed after verification.

## External gate 2 — no-OCR installer size/install

Built with the existing isolated `target-no-ocr` flavor.

Accepted local artifact:

- `apps/windows/src-tauri/target/release/bundle/nsis/VisualTeX_1.2.7_x64-no-ocr-setup.exe`
- size: `121188850` bytes
- SHA256: `4CDCD75BCC0B71A240840280829BF621EE2C0B8A78E344206C7F9205575CABE6`

Comparison full installer present during the acceptance:

- size: `348508502` bytes
- SHA256: `933159E6A900BB44888A1CC41B4CBBA631C39BF995142076B0414A8485D022A4`
- no-OCR reduction: `227319652` bytes

Static flavor verification passed: the no-OCR config/template does not bundle the OCR worker, private Python, wheelhouse, or OCR model catalog.

Clean-directory installed-release acceptance passed. Report: `apps/windows/build-logs/windows-installed-release.json`.

Relevant installed acceptance facts:

- `ocrResourcesOmitted=true`
- `embeddedMainAssetsVerified=true`
- `secondLaunchRecreatedMainWindow=true`
- `desktopCloseExited=true`
- `officeRuntimeVerified=false`
- `officeEditorCreatedBeforeRequest=false`
- Office payload hashes embedded in the installer matched the current build resources.

The user Word process remained running as PID `127920` after acceptance. No Word document interaction, Ribbon action, Office editor session, or Office runtime/readonly verification was performed for this P10 gate.

## Boundary / next phase

P10 is complete. The next permitted phase is P11: create an `OfficeCompatibilityAdapter` and prove output compatibility offline. P11 must not take over a real Office Session and must not modify `WordFormulaService`, numbering, COM/OLE identities, Session state machines, or user Word documents.

No commit, push, release, branch switch, reset, clean, stash, pull, rebase, or merge is authorized by this checkpoint.
