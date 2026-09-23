# P07.3 — accents and over/under annotations

Scope: independent core and its existing SVG/input sandbox. Production VisualTeX, Office protocols, packages and user documents remain unchanged. This does not claim arbitrary LaTeX or final MathLive replacement.

## Supported structures

`Accent` owns an editable body row, command spelling and trivia. `Annotation` owns an editable label row and base row, placement direction and original command spelling. Supported narrow marks: hat, tilde, vec, bar, dot, ddot, dddot, ddddot, breve, check, acute, grave, mathring. Stretchable marks: widehat, widetilde; over/under right/left/bidirectional arrows; over/under braces; over/under parens and arc aliases. Overline/underline use font MATH rule and gap values. Labels cover overset, underset, stackrel and stackbin, including nested over/under structures. Labelled braces retain the authored Scripts structure but stack their attached labels in layout.

Capabilities tested independently: lossless and canonical source, structural insertion/wrapping, stable identities, source edits, copy independence, deletion and unified undo/redo, glyph geometry, SVG, scene queries and browser input. Unknown command preservation, incomplete drafts and unfilled required slots remain distinct from export-ready support.

## Navigation and interactions

A normal Right arrow at the end of a completed accent chain exits the chain without a content revision or undo entry. It stops before an unvisited sibling, nonfinal script argument, fraction or annotation boundary. Shift+Right retains the original anchor and normal incremental selection. Explicit ExitStructure is a separate non-content command, exposed as a sandbox control; it does not redefine final product shortcut preferences. A new annotation starts in its label, then normal navigation enters its base. Pointer hits and accessible position descriptions use the same core scene and node IDs.

The browser still has one EngineEditorHost owner. No MathLive offsets, external math DOM measurements or second history manager are introduced. The input bridge rejects toolbar changes during uncommitted composition. P06 native Windows IME evidence is retained separately; P07 browser key tests do not claim another physical IME run.

## Layout policy and limits

Same injected font resources and OpenType MATH data as the P05/P07 root/fence baselines. Single-glyph accents use top-accent attachment data; horizontal decorations use font variants and genuine bounded assemblies. Over/underbar thickness, gaps and extra ascender/descender values are font inputs. All outline pieces are uniformly scaled; no anisotropic glyph stretching or browser math typesetting occurs. Narrow marks, wide marks and annotation labels are not interchangeable.

For a wide accent whose font provides only finite variants and no assembly, the existing explicit horizontal-accent policy may select the widest designed variant; it does not fabricate a longer outline. This can leave a mark narrower than an extremely long base. Large font constructions remain bounded and missing glyphs fail explicitly. Flattened accents via the optional OpenType flac feature and exhaustive cross-font typography are NOT yet accepted by this slice. System font files remain read-only fixtures and are excluded from source archives.

Reference: Microsoft OpenType MATH table (top-accent attachments, mathematical constants, glyph constructions); W3C MathML Core horizontal stretching algorithm. Consult the current primary specifications for later extension. Their existence is not evidence that every typesetting detail has been implemented.

## Regression evidence

`crates/core/tests/accents.rs`: registered command corpus, every mathematical style, exact source preservation, label/brace stacking, rule/gap constraints, copied identities, random source mutations, invalid-input atomicity, nested Right exit and Shift selection. `scene_queries.rs`: annotation/inner-accent selection descriptions and vertical landing. `tools/layout-parity.mjs`: real Native/WASM geometry and source/edit parity, fixed resources, same SVG and unchanged natural dimensions across pixel size changes. `tools/accent-browser.mjs`: real sandbox controls and browser key/pointer events, starting from blank, label editing, shared history, zoom, long assembled braces, export and invalid-source recovery.

Two integration failures were first reproduced and then fixed: new nodes missing from accessible position traversal (`invalidPosition`), and Right arrow stopping after only the inner accent. Old unsupported-hat fixtures are retained as positive coverage; new unknown-command cases retain negative coverage. No failing sample was simply deleted.

Pass status, tested source hashes, binary hashes and source archive hashes belong in immutable checkpoint receipts, not this specification. MathML and Office-target export remain later-stage gates.
