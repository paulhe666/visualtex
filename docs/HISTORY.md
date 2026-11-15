# Where removed material went

The restructuring of November 2026 took work logs, evidence and dead code out of the tree. Nothing was rewritten
in history: every file is still in git. Read one with `git show <commit>^:<path>`, list a folder with
`git ls-tree -r --name-only <commit>^ -- <folder>`, or restore it with `git checkout <commit>^ -- <path>`.

| Removed | Commit |
|---|---|
| Separate Windows and Mac histories, each with an old copy of the other app (`apps/macos` in the Windows line, `apps/windows` in the Mac line) | `60df1878` (the merge that joined them) |
| `engine/` (Rust/WASM editor rewrite) and `docs/engine-rebuild/`; Windows `src/engine/*`, `CoreMathEditor*`, `office/compat/*`; phase scripts `p10_*`–`p16_*` | `124b95c5` |
| Windows Office.js-era frontend (`office/adapters`, `office/bridge`, `office/api/companion*`, `office/windows-ole`, Office.js manifest templates) and other files no page loaded, with the tests that only covered them | `5563ed12` |
| `docs/remediation-3d207d7/` and `docs/audit-3d207d7/` (evidence, scripts, reports), `artifacts/`, 30 dated stage/checkpoint/review documents in `docs/` | `3f393ed4` |
| Stray files (write probe, a screenshot named `-`, a `.orig` merge leftover, an empty `.sha256.txt`, a compiled smoke `.exe`) | `73147fb4` |
| 38 empty or comment-only C# files in `src-windows` | `5bf98320` |
| Windows checkpoint scripts `cp239_*`–`cp260_*` | `c3138129` |
| 188 scripts no file named (Office probes and repros, installed-app checks, retired stubs, the macOS copy of the Windows reference editor regression) | `3d1e6e4b` |
| The two targeted scenario runners (`apps/windows/scripts/targeted_editor_regression.mjs`, `apps/shared/visualtex-app/tests/targeted_editor_scenarios.mjs`), replaced by `apps/shared/visualtex-app/tests/editor-scenarios/` (one file per scenario). Not carried over, as no script ran them and they failed against the current editor (and against the builds from before the restructuring, Windows 368ef679 and macOS 22981652): from the Windows runner atomic-structure, bounded-operator-model, candidate-query-reset, classic-panel-resize, context-style, cursor-placement, direct-shortcut-placeholder, formula-chinese-font, history-core, native-space-selection, ocr-storage-ui, persistent-formatting, placeholder-visual-audit, script-template, settings, source-layout, startup, suggestions, toolbar-compact, toolbar-placeholder-overflow, toolbar-postfix; from the macOS runner accent-placeholder, candidate-query-reset, configuration, context-style, cursor-placement, font-settings, font-variant-formatting, limit-candidate, native-space-selection, native-structure-input-over, ocr-empty-environment, output-fonts, placeholder-selection, row-stability, selection-geometry, settings, source-font-size, source-layout, source-structural-draft, structural-placeholder, suggestions, text-command-fonts, toolbar-compact, vertical-structure-navigation, wrapper, wrapper-auto, wrapper-bm, wrapper-continuous, wrapper-prefix, and its own versions of delete, export, formula-formatting, formula-tiles, native-input-popover, navigation, raw-placeholder-visual and upright (the Windows versions are the ones kept) | `d8c10bcd` |
| Browser tests and fixtures nothing ran any more: macOS `auto_escape`, `editor_layout_switch`, `office_unified_toolbar`, `quick_format_toolbar` and `theme_customization` regressions (only already removed scripts named them), the macOS copies of `editor_regression_smoke.mjs` and `editor_history_regression.mjs` (no macOS script since July; the Windows ones run), `macos_ax_ime_probe.swift`, `verify_powerpoint_native_omml.py`, the `document-import-mixed-math` fixture and the macOS copy of `targeted_upright_typography_acceptance.mjs`; Windows `ocr_align_editor_geometry_acceptance.mjs` (the `.mts` is the one that runs), the Windows copy of `native_suggestion_preview_regression.mjs` (the macOS one runs) and the `windows-office-latex-compatibility` / `windows-office-special-operators` fixtures | `d8c10bcd` |

Test inputs that lived in `docs/remediation-3d207d7/evidence` and are still used moved to
`apps/windows/test-fixtures` (`mtef-redraw/`, `mathtype-create/`).
