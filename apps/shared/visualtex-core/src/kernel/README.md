# Editing kernel

The formula editor's model, parser, serializer, mode editors, keyboard and IME
input, undo and the `<math-field>` element. It is TypeScript source of the
VisualTeX editing core: MathLive 0.109.2 (MIT, see `LICENSE-MathLive.txt`)
with every VisualTeX change written into it. The caret, movement, empty slots,
selection, pointer and geometry rules it follows are the core's own
(`../structures.ts`, `../navigation.ts`, ...); formulas are laid out and drawn
by the layout engine (`apps/shared/tex-layout`).

- Entry: `mathlive.ts`. Apps import it as `mathlive` (a Vite alias) and register
  the layout engine and the editing core at startup (`setVisualTexLayoutEngine`,
  `setVisualTexEditingCore`, kept in `editor-mathfield/visualtex-hooks.ts`).
- Host module: the tables generated from an app's registries (completion
  catalog, command frequency, default macros, source-wrapper maps, uncommon
  integrals) and the app's input judgment come from `virtual:visualtex-kernel-host`
  (declared in `visualtex-kernel-host.d.ts`, provided by the app's Vite plugin).
- DOM-free entry: `public/mathlive-ssr.ts` (parsing, validation, markup by the
  registered layout engine), imported as `mathlive/ssr`. Node scripts reach it
  through the app's `scripts/kernel-node.mts`.
- Stylesheets are plain CSS in `css/`, imported as text.
- `tsconfig.json` keeps MathLive's class-field semantics (`useDefineForClassFields:
  false`); Vite compiles these files with it. The source is strict-clean: apps
  type-check it in their own program (`paths` for `mathlive` and `mathlive/ssr`).
