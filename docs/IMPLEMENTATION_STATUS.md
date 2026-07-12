# Implementation status against the architecture plan

This file records what is executable and tested in the repository as of the current revision. A phase is called complete only when its stated exit criteria are met across the intended platforms; “substantially complete” means the main product path exists but release, breadth or platform validation remains.

| Phase | Status | Implemented | Remaining before phase exit |
| --- | --- | --- | --- |
| 0. Engineering baseline | Substantially complete | Cargo/pnpm workspaces, Rust/TypeScript formatting and strict checks, protocol versioning, ADRs, multi-platform CI definitions, Tauri build, UTF-8 and restricted-mode policies | Signed release builds and reproducible release validation on macOS, Windows and Linux |
| 1. Project and text buffer | Substantially complete | Root discovery, templates, multi-file source list/open, Rope buffer, revisions, UTF-8 edits, dirty state, atomic save, recovery, filesystem watcher, external-editor reload/conflict handling, shared undo/redo, CodeMirror | Rename/move/delete project UI, three-way merge UI, broader large-file and long-running random edit campaigns |
| 2. Tolerant syntax tree | Functional baseline | `vt-latex-syntax` backed by Tree-sitter LaTeX, full and incremental parses, byte/point edit mapping, changed ranges, malformed/missing-node reporting, verbatim/minted/code-environment protection ranges, `input`/`include`/`subfile` dependency extraction, missing-target status, cycle detection, project-wide source preload and edit/undo/redo integration | Broader `import`/package/class/macro dependency edges, removal/rename lifecycle integration, larger real-world corpus, dedicated fuzz targets and performance thresholds |
| 3. Semantic model | Functional baseline | `VisualNode`, support levels, spans, deterministic IDs, common paper structures, attributes for figures/tables, typed index, incremental `VisualPatch`, local serializers | Stronger ID continuity after large structural rewrites, complete multi-file semantic graph, broader macro/package registry |
| 4. Bidirectional structured editing | Substantially complete | CodeMirror and the real compiled-page editor share one revisioned Core buffer; PDF-node selection reveals the UTF-8 source position without stealing editor focus; MathLive edits inline/display formulas; visual/source edits, attribute edits, loop prevention, Core undo/redo and Unicode handling | Rich clipboard schema, range-level cross-surface selection, larger IME/e2e matrix and broader directly editable node coverage |
| 5. Local compilation | Substantially complete | Tool detection, latexmk/Tectonic adapters, XeLaTeX/pdfLaTeX/LuaLaTeX, isolated output, diagnostics, timeout, restricted shell escape, real PDF and SyncTeX, Windows `.exe` discovery and external-tool path normalization, desktop compile coalescing that prevents concurrent `latexmk` writers and discards stale artifacts | Persistent build cache policy, Core-level cancellation/prioritization scheduler, clean/export commands and full template/toolchain matrix |
| 6. PDF preview and SyncTeX | Substantially complete | PDFium document inspection, page/tile and thumbnail rendering, fingerprinted cache, Windows-safe atomic PNG publication, pixel comparison, desktop canvas viewer, VS Code PDF viewer, forward/inverse SyncTeX, highlights and click handling | Larger-document virtualization/performance profiling and native three-platform PDFium release validation |
| 7. Shadow instrumentation | Functional baseline | Shadow source generation, instrumented compilation, parsed source-to-page SyncTeX boxes, confidence/source revision metadata, Windows PDF pixel comparison and a real Chinese fixture proving zero changed pixels; mapping failures are surfaced instead of silently downgraded | Broader document-class/template compatibility, larger quantitative zero-layout-change corpus, production-safe zero-width markers and fallback adapters |
| 8. Direct PDF-page editing | Functional baseline | The desktop “结构化编辑” tab now uses the actual compiled PDF page rather than semantic cards; high-confidence sections, paragraphs, inline/display formulas, figures and tables can be selected in place; formulas open an auto-focused MathLive overlay; commits immediately write through Core and trigger coalesced recompilation | More robust reflow relocation, complex nested content hit testing, title/author mapping through `maketitle`, drag/resize interactions and broader node coverage |
| 9. Complete paper authoring | Functional baseline | Templates, project tree, source completion, labels/citations/bibliography/macro index, search, hash-guarded replace, typed rename, figure/table attributes, diagnostics, toolchain/model panels | Bibliography manager, image asset manager, richer table UI, texlab integration, Git UI, spellcheck and comprehensive refactor catalogue |
| 10. Formula OCR | Functional product path; model optional | Offline worker manager, verified model packages, health/capabilities, bounded image normalization, candidates/confidence/model version, desktop MathLive review, VS Code review and native undoable insertion | Ship and benchmark a production model package for each release platform; establish quality thresholds and model supply-chain release process |
| 11. Full-page OCR | Functional product path; model optional | Layout/document worker method, regions, confidence, reading order, type/order/content correction, low-confidence highlighting, formula editing, ignored regions, LaTeX generation, atomic new OCR project preserving original page and structured JSON, desktop and VS Code review | Ship/benchmark production layout/text/table/formula model set; multi-page import and larger document benchmark suite |
| 12. VS Code extension | Substantially complete | `CustomTextEditorProvider`, trusted local workspace requirement, `TextDocument` authority, `WorkspaceEdit`, native undo/redo, external-save hash confirmation, Core auto-reconnect/resync, PDFium preview, bidirectional SyncTeX, formula/full-page OCR, 12 extension tests | Packaged VSIX validation, `@vscode/test-electron` host tests, marketplace metadata and native Windows/Linux/macOS installation matrix |
| 13. TeXstudio and general Bridge | Substantially complete | Authenticated loopback Socket/JSON-RPC, random port/token/discovery, protocol negotiation, persistent per-project Core, CLI wrappers, external-file refresh/conflict refusal, compilation, bidirectional SyncTeX, POSIX/PowerShell adapters, non-overwriting installers, desktop open command, versioned `visualtex://` scheme, Node Adapter SDK, multi-project/disconnect/invalid-client/Unicode tests, native Windows PowerShell install/start/compile/SyncTeX/shutdown validation with UTF-8 JSON paths | TeXstudio-version-specific macro profiles for cursor/viewer consumption, automated GUI-host validation, optional richer live selection channel and packaged adapter release artifact |
| 14. Plugins | Not implemented | Extension boundaries and security requirements documented | Manifest, permissions, signed package policy, snippets/templates registry, CLI adapter controls and WASI sandbox |
| 15. Release and stability | Partial | Offline defaults, recovery, model package verification, dark/light UI, CI checks, `visualstudio` 0.1.1 debug/release builds, unsigned Windows x64 NSIS installer with silent install, installed executable/version/window-title launch check and uninstall verification, security documentation and real fixture/e2e tests | Authenticode signing, macOS notarization, Linux/macOS installers, updater verification, accessibility/i18n audit, memory/leak profiling, crash telemetry policy and release regression matrix |

## Current end-to-end workflows

### Desktop

1. Open or create a LaTeX project.
2. Edit source or switch to “结构化编辑” to modify high-confidence content directly on the real compiled page.
3. Keep all edits in one revisioned Rust buffer and undo history.
4. Save atomically with recovery and external-change conflict protection.
5. Compile with the local TeX toolchain.
6. Render the real PDF through PDFium and use bidirectional SyncTeX.
7. Run formula or full-page OCR with explicit review before insertion.
8. Open versioned `visualtex://` project and SyncTeX actions without silently switching away from another dirty project.

### VS Code

1. Open a `.tex` document with the VisualTeX custom editor.
2. Keep the VS Code `TextDocument` authoritative.
3. Apply visual/OCR changes through `WorkspaceEdit` and native undo/redo.
4. Restart and resynchronize Core after transport failure.
5. Compile and render PDFium pages in the Webview.
6. Use source-to-PDF and PDF-to-source SyncTeX.

### TeXstudio and other editors

1. Install adapter scripts without modifying editor preferences.
2. Start or auto-start one authenticated Bridge per project.
3. Save in the external editor and safely refresh clean disk changes.
4. Compile and perform SyncTeX through the persistent Core.
5. Open the same project in the desktop application.
6. Use `packages/adapter-sdk` or the versioned URI scheme from another local editor.

## Verified tests beyond unit suites

- Real Chinese `ctexart` compilation with PDF and SyncTeX output.
- PDFium page rendering and a real Chinese authoritative-vs-shadow PDF comparison with zero changed pixels.
- TeXstudio POSIX adapter install/refusal/force/uninstall flow.
- Windows PowerShell adapter parse/install/idempotency/refusal/force flow on Windows 10 with TeXstudio 4.8.9 and TeX Live 2025 installed.
- Persistent Bridge start/status/compile/forward search/inverse search/shutdown on a path containing Chinese characters and spaces, including a native Windows clean-build round trip back to `main.tex` line 9.
- Bridge-to-disk save, external-editor reload and simultaneous-edit conflict refusal.
- CLI `visualtex://` open, forward search, inverse search and unsupported-version rejection.
- Tauri debug desktop build with registered deep-link configuration.
- Unsigned `visualstudio` 0.1.1 Windows x64 NSIS installer build plus silent install, `visualstudio.exe` version/process/window-title launch check and uninstall smoke test.

## Claims that remain deliberately unsupported

The repository must not be described as having:

- a complete, corpus-qualified LaTeX include/macro dependency graph with universal package and document-class coverage;
- universal compatibility with arbitrary LaTeX macros and document classes;
- bundled, benchmark-qualified production OCR models for every platform;
- a third-party plugin sandbox or marketplace;
- automated TeXstudio GUI macro/profile validation across supported TeXstudio versions;
- signed, notarized and auto-updating release artifacts for all three platforms.

These remain independent deliverables. The current source authority, protocol, security and adapter boundaries are designed so they can be added without replacing the implemented Core.
