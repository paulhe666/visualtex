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

Test inputs that lived in `docs/remediation-3d207d7/evidence` and are still used moved to
`apps/windows/test-fixtures` (`mtef-redraw/`, `mathtype-create/`).
