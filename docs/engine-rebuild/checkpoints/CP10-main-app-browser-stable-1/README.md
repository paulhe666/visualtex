# CP10 main app browser stable 1

Stable non-Git checkpoint for the P10 Windows desktop main-application integration state that has passed the complete browser/main-app and source-level verification set available without disturbing the user's running installed VisualTeX process.

Completion label: `browserStable-notP10Complete`

Stable source SHA256: `f39d8f3dbb59e39042a91f8d2aed5db30619d86fe47187d4012cd1f5effa5825`

Archive SHA256: `b057aeb5097885211411117d5ff7c08288f42272e8ead53f615469928d30a21a`

Verification receipt: `engine/target/p10-verification/2026-09-14T14-24-07-439Z/report.json`

Mandatory P09 lower-phase receipt: `engine/target/p09-verification/2026-09-14T14-20-44-010Z/report.json` (`20/20 PASS`, current-source audit PASS).

The P10 stable-source verifier passed 15/15 gates with `sourceUnchanged=true`. It includes a current-source audit of the full P09 stable receipt, the complete engine build/Native-WASM parity, strict core selection-token owned-Edge acceptance, default-vs-candidate build resource isolation, the real product-page core editor acceptance, Windows verified system-font byte extraction for Cambria Math and Microsoft YaHei, the no-OCR source/config dependency contract, legacy editor/input-behavior/source-editor/formula-hotkey regressions, HistoryManager, application data, and configuration compatibility.

The desktop integration now chooses the engine before mounting an editable math surface. A supported document session mounts the self-owned core; an unsupported new session starts on legacy from the beginning. Once a core owner exists, later unsupported layout/font content does not silently replace it with MathLive. A second initialization check remains fail-closed but no longer performs a mid-session legacy fallback.

Core selection ownership is separate from MathLive offsets. The token binds owner identity, document ID, revision, source ID, and structural selection. Tokens are rejected across another owner, after content identity changes, or after recreating a same-named session. CodeMirror, header Undo/Redo, visual edits, shortcut conversion and source edits share the core history while a core session is active.

Input behavior is reimplemented structurally rather than copying MathLive's event/DOM workarounds. Superscript, subscript, accent and style-scope auto-exit use core structure paths and `exitStructure`. `End` retains its local structural meaning; `ArrowRight` performs normal structural navigation. Product inline shortcuts use the bounded `ReplaceRecentMath` core transaction and merge with the same typing group, so for example `alpha -> \\alpha` is one undo item. Command suggestions use the product command data but the self-owned `CommandInputState` and structural `EditCommand` actions.

Custom-symbol definitions are frozen at session creation. Product candidate confirmation, usage personalization, toolbar insertion, source projection and user custom-symbol insertion pass on the real product page. The custom-symbol engine remains the P09 immutable resource implementation; no browser HTML/SVG/CSS execution is introduced.

Font layout is content-addressed. Latin Modern Math is verified from the bundled font bytes. Other Windows preferences are resolved by a bounded read-only Tauri system-font command that accepts only requested family names, returns font bytes + face index + SHA256, and never exposes a general filesystem path/read primitive. The frontend recomputes SHA256 before registering the bytes with the core. Real Windows acceptance resolved Cambria Math and Microsoft YaHei and rejected a nonexistent family. Layout preference updates do not rewrite source or history; unresolved changes keep the previous verified layout.

`MathPreview(staticLayout)` has a core-first noninteractive path for safe formulas. It creates no BrowserInputBridge or long-lived editor host, exports validated SVG from a short-lived owner, destroys the owner on every cache miss, and enforces an active-session leak guard. The real product acceptance observed 318 static previews rendered by core SVG with no `.ML__latex` and no `.vt-core-input`. Changing font identity retires old core preview markup; unsupported previews retain the legacy static renderer instead of creating a partially valid core image.

Save/open is covered through the real product page. Save produces the existing `.visualtex.json` schema with the exact current core LaTeX. Opening through the real file input creates a new document/editor generation and re-runs session preflight; a supported opened formula selects core without carrying the previous owner.

Default production builds remain isolated: without `VITE_VISUALTEX_CORE_EDITOR=1`, the output contains no core WASM, bundled core font, CoreMathEditor chunk or core static-preview implementation. Candidate builds include them explicitly. Core editor/runtime/preview source has no local OCR runtime/model/download dependency; the base Tauri config embeds no OCR assets and the no-OCR overlay adds none.

This checkpoint is **not P10 complete**. Two explicit P10 gates remain unexecuted: (1) close and second launch of the real Tauri application, because an installed user VisualTeX process was detected running and must not be closed or stolen by a dev single-instance test; (2) installer-level no-OCR size/install verification, because no installer is being produced in this phase. Office interaction and Office document algorithms remain intentionally untouched.

The archive contains exactly the 267 source inputs locked by the P10 receipt. Archive audit confirms every file hash matches, the current working source still matches the verification receipt, and no `target/`, `dist/`, `node_modules/`, font binary, WASM, executable, key, certificate or environment file is present.

No Git commit, push, pull, merge, rebase, reset, clean, stash or branch switch belongs to this checkpoint. No user Word document, Office/VSTO implementation, Office Session, or evidence work was touched. P00 protection remains intentionally red because of independent parallel work.
