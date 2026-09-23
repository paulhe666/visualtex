# P13 Windows Office hidden conversion / batch integration

Status: `complete`.

## Preconditions and frozen baseline

P13 starts only after the final P12 completion checkpoint:

`docs/engine-rebuild/checkpoints/CP12-office-interactive-complete-1/`

P12 final verifier: `engine/target/p12-verification/2026-09-15T23-58-21-717Z/report.json`, status `PASS`, `completion=complete`, `p13Permitted=true`, `sourceUnchanged=true`.

P13 baseline branch is `reference-3d207d7`; baseline HEAD reference is `d1f0aff8db9275c90dc18c7318dc2a4d7bbbd760`. The repository already contains historical dirty/untracked work from the engine project and parallel Office work. P13 must preserve it and must distinguish pre-existing working-tree state from changes made by this stage.

The following paths remain frozen: WordVsto, PowerPointVsto, Windows Office Contracts, FormulaOleServer, VstoShared, WindowsOleBridge, Tauri Office Session/server/lifecycle/state/backend/pipe code, `src/office/shared`, `src/office/api`, and `src/office/dialog/useOfficeSession.ts`. P13 does not alter Session DTO/state transitions, COM/OLE contracts, numbering/reference algorithms, Word document replacement algorithms, converter queue semantics, or Office routes.

## Integration boundary

The existing hidden conversion lifecycle is retained. Both single hidden conversion and the existing sequential converter-batch queue call `OfficeDialogApp.generateSessionExportResult()`. P13 changes only the export provider behind this function.

The P13 provider is headless: it may create an `EngineEditorSessionOwner`, parse/layout/export and reuse the cached core runtime/font resources, but it must not create `CoreMathEditor`, a DOM input view, IME host, candidate panel, focus-repair loop, or any interactive editor state.

`wordOmml` remains a MathML-only fast path and must not rasterize or request SVG layout. `mathTypeOle` requires compatible Presentation MathML and vector SVG for the existing preview consumer, but does not require PNG. `nativeOle` and `crossPlatformPicture` keep the full graphical contract expected by the existing consumers, including PNG where required.

MathType output continues to use the P11 `OfficeCompatibilityAdapter` and its core `gridColumns`-derived ruler stops. Existing private MathType native runtime/WMF handling remains untouched.

## Identity, ordering, and failure semantics

The current `vsto-convert` queue, `takeOfficeConverterBatch()`, per-session PATCH target, status transitions, failure continuation, polling/event handling and hidden-window close behavior remain unchanged. Every core owner is bound to the source Session/formula identity and is destroyed after that one export. Runtime/font/glyph resources are cached across owners; formula document state is not shared across Session ids.

Malformed or unsupported formula input must fail that Session before Office replacement. Existing batch behavior decides whether the caller continues, skips, or reports the failure; P13 does not silently substitute another formula or write a partial result.

## Equation tags

Office `equationTag` is an existing business/metadata semantic. The core parser does not own the legacy `\\tag` command, so P13 does not inject `\\tag` into core source. The hidden-conversion provider exports the formula body through core, then applies an Office-compatibility wrapper: Presentation MathML is wrapped as the existing `mtable/mlabeledtr/mtext` label structure, while graphical object modes create the parenthesized label through a second headless core owner and merge the resulting path-only SVG with the body. The merged SVG retains the Office-safe vector subset and is rasterized only when the target object contract requires PNG.

Production-bundle browser acceptance now covers tagged `wordOmml`, tagged `nativeOle`, and tagged `mathTypeOle`. The three paths use the core compatibility wrapper; tagged graphical output contains no MathJax ownership, DOM text/image/script nodes, or interactive editor surface.

## Browser acceptance completed so far

The production Office bundle has been exercised through `scripts/p13_office_core_converter_browser_smoke.mjs`. The hidden `vsto-convert` surface contains zero MathLive fields, zero core input elements and zero interactive core editor mounts. The same browser process verifies `nativeOle` as core SVG+MathML+PNG, `wordOmml` as core MathML-only with preserved geometry, `mathTypeOle` as core SVG+MathML with no PNG and P11 ruler metadata, and `crossPlatformPicture` as core SVG+MathML+PNG. The converter batch preserves requested Session order and identity, continues after an unsupported formula fails, and reuses one loaded core WASM/font resource set across the batch. Tagged OMML/native/MathType paths are also covered through the core compatibility wrapper.

## Final acceptance

The frozen real-Office acceptance record is `engine/specs/P13-real-office-acceptance.json`.

Production-bundle browser acceptance covers both sides of the P13 boundary. `scripts/p13_office_interactive_core_guard.mjs` keeps the P12 interactive Office editor on the retained core editor, including real browser input, same-revision SVG/MathML/PNG export, product Undo/Redo, Session switching/stale-export rejection and the existing Apply/close handshake. `scripts/p13_office_core_converter_browser_smoke.mjs` independently verifies that `vsto-convert` is headless: no MathLive field, no core input view and no interactive core editor are mounted. Native OLE, OMML, MathType and PowerPoint picture exports use the expected core output contracts; an unsupported formula fails only its Session, the following Session still completes, batch terminal writes preserve order/identity, and equation tags are produced by the core compatibility wrapper. `scripts/p13_office_core_converter_performance.mjs` passes 100 mixed formulas and 1000 OMML formulas with zero failures while reusing the core runtime/font resources; the OMML path remains MathML-only.

Real Word Ribbon/UI acceptance completes the six full-document conversion directions on a 100-formula corpus with zero failures. Object audits show 100 VisualTeX OLE objects after redraw, exactly 100 Word OMath objects on OMML legs, exactly 100 `Equation.DSMT4` objects on MathType legs, and stable 101-paragraph document structure. A later 146-object MathType → VisualTeX full-document run also passes while preserving two unrelated inline shapes and one unrelated OMML equation. Selected OMML → MathType and MathType → VisualTeX conversions pass 1/1; the earlier selected VisualTeX → OMML run also completed 1/1 with `wholeDocument=False`.

The 1000-formula real Word redraw is accepted independently from its later intentional Undo probe. The product log records `redraw-complete formulas=1000 unique=1000 skipped=0`, average core render 8 ms, max 49 ms, insert total 7233 ms and max 127 ms. Strict materialization validation subsequently records `redraw-omml-materialized-verified formulas=1000 elapsedMs=17759`. The later return to LaTeX is the explicit Undo acceptance captured by the `p13-core-office-1000-undo-*` evidence, not a spontaneous product rollback. The Undo probe restores the source text but does not claim a byte/format-perfect restoration; that pre-existing Word undo-format issue is not used as P13 completion evidence.

Real source restoration also passes through the actual `重绘全文 → 全文 VisualTeX OLE 转为 LaTeX 代码` Ribbon path: 146 VisualTeX OLE formulas are restored in place, while the two unrelated inline shapes and the unrelated OMML formula remain. Real bulk import passes through the visible `批量导入` window and `导入到 Word` action using `engine/fixtures/p13-office-bulk-import.md`: the preview reports 7 blocks, 1 inline formula and 2 display formulas; Word receives 3 OMath objects, zero OLE objects, 8 paragraphs and the surrounding Chinese prose remains present in source order. The sample includes a 2×2 matrix and a long fraction/sum/product/integral expression.

The malformed/unsupported formula requirement is deliberately separated from Word mutation: the production converter browser smoke proves that an unsupported formula Session receives `failed` with no partial export and that the next Session still commits. Word replacement is therefore never attempted for that failed Session.

Protected Office paths remain frozen at the P12 boundary: WordVsto, PowerPointVsto, Windows Office Contracts, FormulaOleServer, VstoShared, WindowsOleBridge, Tauri Office Session/server/lifecycle/state/backend/pipe code, `src/office/shared`, `src/office/api`, and `src/office/dialog/useOfficeSession.ts` have no P13 tracked diff. P13 changes the headless export provider and its core runtime helper, not the Office wire/document algorithms.

During exploratory real-Office work three non-deterministic debug-companion crashes were recorded at Windows exception `0xc000001d`, fault offset `0x114eace`; disassembly maps the trap to a Tao `EventLoopRunner` `Rc` clone in native `Window::new`. It could not be reproduced by exact 1000-Session converter batching, idle prewarm, concurrent first creation, 30 converter close/recreate rounds, or a final 20-round lifecycle stress on a Word-started `--office-background` companion. Final 100/146 and selected conversion acceptance passed without a new crash. This runtime observation is carried forward to P14 fault hardening; available evidence does not identify a P13 formula-core, MathML, Session identity or ordering failure.

P13 completion is gated by `engine/tools/p13-verify.mjs`; the final checkpoint may be created only from a PASS receipt with unchanged verified inputs and an unchanged frozen Office boundary.
