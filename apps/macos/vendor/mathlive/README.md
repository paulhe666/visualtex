# VisualTeX MathLive kernel

This is the macOS editor's owned browser fork of MathLive 0.109.2. Edit this
source directly. Updating the npm package does **not** update this fork.
The upstream MIT notice is in `LICENSE.txt`.

`kernel-parts/part-000.mjsfrag` through `part-029.mjsfrag` are consecutive pieces
of one JavaScript module, not independent modules. `vite.mathlive.ts` validates
the sequence and concatenates it verbatim. It does not rewrite behavior.
Preserve fragment boundaries when editing; some boundaries split functions.

The only generated module, `virtual:visualtex-mathlive-data`, supplies shared
command catalog, macro definitions and checked integral glyph data. Their
sources remain under `src/autocomplete` and `src/math`, so toolbar, validator,
export and browser kernel do not acquire separate copies of product data.

Key extension locations:

- `part-002`: command search and personalized completion data.
- `part-005`, `part-013`, `part-014`: integral glyphs, rendering and command mappings.
- `part-019`: selection styling and persistent typing styles.
- `part-020`–`part-023`: editing, structural placeholders, raw LaTeX wrappers,
  completion, script/accent navigation and input safety.
- `part-026`–`part-028`: field lifecycle, composition and custom element API.
- `types.d.ts`: the public additions consumed by the application. Application
  code must not read or write `_mathfield`.

Single-formula editing belongs here. React owns document rows, document undo,
workspace layout and host integration. In particular, script/accent auto-exit
uses model branches for keyboard and toolbar input; do not recreate it from DOM
geometry in the application.

Selection uses exclusive start offsets and inclusive end offsets. Structural
expansion, direction, placeholder order and row-space pointer gestures belong
to the kernel. React keeps one document anchor/focus pair and projects it onto
the fields. On a cross-row drag it cancels the native pointer tracker before
taking ownership; it must not repair competing trackers with delayed writes.
History restores this pair together with all field selections.
Renderers bind structural boxes to their atom IDs. Bounds include SVG viewports
and drawn frames through layout wrappers, excluding struts and clipped paths.

`visualTexInsertDocumentEdges()` joins existing document fragments without
applying the current typing style to the preserved prefix or suffix.

Validation from `apps/macos`:

```sh
npm run build:desktop
npm run test:editor-refactor
npm run test:selection
npm run test:mathlive-runtime-safety:run
npx tsx scripts/physics_kernel_compatibility_regression.mts
node scripts/targeted_editor_regression.mjs native-input-popover
node scripts/targeted_editor_regression.mjs font-variant-formatting
```

The npm `mathlive/ssr` export and static fonts remain separate dependencies.
A future upstream upgrade must port the owned behaviors and run the browser
regressions; replacing these fragments with a stock distribution loses them.
