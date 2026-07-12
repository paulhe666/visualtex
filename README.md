# visualstudio

visualstudio is an offline-first LaTeX paper editor built around one authoritative UTF-8 source model. The Tauri desktop app, CLI, VS Code extension, TeXstudio Bridge and other adapters all use the same revisioned Rust Core.

## Implemented product path

- Safe project discovery, templates, multi-file buffers, atomic save, recovery and external-editor conflict handling.
- CodeMirror source editing and real compiled-page structured editing on one undo/redo history, including a VisualTeX 1.0.6-derived formula workbench with command candidates, selection-aware insertion and matrix construction.
- Tree-sitter incremental LaTeX syntax trees with malformed-input recovery, verbatim/minted protection ranges and project include-cycle detection.
- Tolerant LaTeX semantics with native, partial, opaque and unstable nodes.
- Project index, dependency graph, search, hash-guarded replacement and typed label/citation rename.
- Real local latexmk/Tectonic compilation with XeLaTeX, pdfLaTeX or LuaLaTeX, isolated output, diagnostics, timeout and restricted shell escape.
- PDFium page/tile rendering, Windows-safe atomic cache writes, pixel-validated shadow node-to-page mapping, glyph-level prose/formula hit testing, PDF overlays and bidirectional SyncTeX.
- Offline formula and full-page OCR review workflows with bounded image import and optional verified local model packages.
- VS Code custom editor with `TextDocument` authority, `WorkspaceEdit`, Core reconnect, PDF and OCR panels.
- Authenticated persistent loopback Bridge, TeXstudio adapters, versioned `visualtex://` actions and Node Adapter SDK.

The repository does not yet include a corpus-qualified complete LaTeX/macro dependency graph, bundled benchmark-qualified OCR models, third-party plugin sandbox, or signed/notarized release pipeline. See `docs/IMPLEMENTATION_STATUS.md` for the phase-by-phase boundary.

## Requirements

- Rust 1.96 or the pinned `rust-toolchain.toml` toolchain.
- Node.js 24 with Corepack.
- A local TeX distribution. `latexmk` plus XeLaTeX, pdfLaTeX or LuaLaTeX is recommended.
- Python 3.9 or newer only for OCR.
- Native desktop libraries required by Tauri on Linux.

## Install dependencies

```bash
corepack pnpm install
```

## Desktop application

Development:

```bash
corepack pnpm --filter @visualtex/desktop tauri dev
```

Debug application build:

```bash
corepack pnpm --filter @visualtex/desktop tauri build --debug --no-bundle
```

Unsigned Windows NSIS test installer:

```bash
corepack pnpm --filter @visualtex/desktop tauri build --bundles nsis
```

The generated installer is under `target/release/bundle/nsis`. It is suitable for local testing but is not Authenticode-signed, so Windows SmartScreen may warn.

Open a project directly:

```bash
cargo run -p visualtex-cli -- open ./paper
```

The desktop binary also accepts:

```bash
visualstudio --project ./paper
```

Compiled artifacts are stored under `.visualtex/build`, PDF cache under `.visualtex/cache/pdf`, OCR imports under `.visualtex/ocr-input`, and recovery data under `.visualtex/recovery`.

In the desktop app, compile the project and switch to **结构化编辑**. The editor keeps the real PDF page visible and combines SyncTeX with PDFium glyph bounds so ordinary prose and an inline formula on the same line can be clicked separately. Formulas open a viewport-clamped floating workbench derived from VisualTeX 1.0.6, with blue selection styling, command candidates, selection-aware fraction/root/script insertion, custom matrices, clearing and whole-node deletion. Paragraphs support adding, removing and replacing text. Figures expose controlled LaTeX editing for path, width, float placement, caption and label; the page resize handle writes a `width=…\\linewidth` value rather than pretending LaTeX floats have arbitrary pixel coordinates. Confirming an edit writes to the authoritative LaTeX buffer and coalesces automatic recompilation.

## CLI

```bash
cargo run -p visualtex-cli -- init ./paper
cargo run -p visualtex-cli -- inspect ./paper
cargo run -p visualtex-cli -- dependencies ./paper
cargo run -p visualtex-cli -- compile ./paper
cargo run -p visualtex-cli -- doctor ./paper
```

PDF and SyncTeX:

```bash
cargo run -p visualtex-cli -- pdf-diff left.pdf right.pdf
cargo run -p visualtex-cli -- pdf-text-hit paper.pdf 1 300 400
cargo run -p visualtex-cli -- layout-map ./paper .visualtex/build/main.pdf
cargo run -p visualtex-cli -- forward-search ./paper main.tex 20 1 .visualtex/build/main.pdf
cargo run -p visualtex-cli -- inverse-search ./paper .visualtex/build/main.pdf 1 300 400
```

Stdio JSON-RPC:

```bash
cargo run -p visualtex-cli -- rpc ./paper
```

Each request and response occupies one line. Call `initialize` first.

## Persistent local Bridge

```bash
visualtex bridge-serve ./paper
visualtex bridge-status ./paper
visualtex bridge-request ./paper initialize --params '{}' --result-only
visualtex bridge-compile ./paper
visualtex bridge-forward-search ./paper main.tex 20 1 .visualtex/build/main.pdf
visualtex bridge-inverse-search ./paper .visualtex/build/main.pdf 1 300 400
visualtex bridge-shutdown ./paper
```

The Bridge uses a random `127.0.0.1` port and per-session token under `.visualtex/bridge`. It preserves one persistent Core per project, reloads clean external saves and refuses simultaneous-edit conflicts.

TeXstudio installation and command setup are documented in `scripts/texstudio/README.md`.

## VS Code extension

Build and test:

```bash
corepack pnpm --filter visualtex-next-vscode typecheck
corepack pnpm --filter visualtex-next-vscode test
corepack pnpm --filter visualtex-next-vscode build
```

Set `visualtex.corePath` when the CLI is outside `PATH`. Set `visualtex.modelsRoot` to an installed offline model package directory when OCR is required.

## Adapter SDK and deep links

The Node SDK is under `packages/adapter-sdk`:

```ts
import { VisualTexAdapterClient } from "@visualtex/adapter-sdk";

const client = await VisualTexAdapterClient.connect("/path/to/paper");
const dependencies = await client.projectDependencies();
const artifact = await client.compile();
```

Supported versioned actions:

```text
visualtex://open?v=1&project=...
visualtex://forward-search?v=1&project=...&source=...&line=...&column=...&pdf=...
visualtex://inverse-search?v=1&project=...&pdf=...&page=...&x=...&y=...
```

The desktop bundle registers the scheme. The CLI can execute one directly:

```bash
visualtex open-uri 'visualtex://open?v=1&project=%2Fpath%2Fto%2Fpaper'
```

## OCR and local models

Inspect, install and activate an offline package:

```bash
visualtex model-inspect ./model-package
visualtex model-install ./model-package
visualtex model-list
visualtex model-activate formula MODEL_ID VERSION
visualtex model-activate layout MODEL_ID VERSION
```

Run OCR:

```bash
visualtex ocr-health ./paper
visualtex ocr-formula ./paper figures/formula.png
visualtex ocr-document ./paper scans/page-01.png
```

The deterministic `--mock` mode is for tests and integration development. Images are never uploaded or automatically sent to a model service.

## Quality gates

```bash
cargo fmt --all -- --check
cargo clippy --workspace --all-targets -- -D warnings
cargo test --workspace
cargo check -p visualstudio
corepack pnpm -r typecheck
corepack pnpm -r test
corepack pnpm -r build
corepack pnpm --filter @visualtex/desktop tauri build --debug --no-bundle
```

The real Chinese fixture is `fixtures/templates/basic-zh`.

## Security defaults

- Project, PDF, SyncTeX and OCR paths are canonicalized and scoped.
- Symlinked source and OCR input files are rejected.
- Build output cannot escape the project.
- `shell-escape` is disabled in restricted mode.
- TeX and adapter commands are never constructed by interpolating source into a shell.
- OCR import is size/dimension/allocation bounded and normalized before worker access.
- Recovery and external-editor refresh never silently overwrite a changed disk file or dirty Core buffer.
- Bridge authentication, protocol version, loopback endpoint and message-size limits are enforced before Core dispatch.
- Deep links cannot silently switch away from a different open project.

Detailed design and limits:

- `docs/SECURITY.md`
- `docs/PROTOCOL.md`
- `docs/IMPLEMENTATION_STATUS.md`
- `PROJECT_ARCHITECTURE_AND_IMPLEMENTATION_PLAN.md`

## License

visualstudio is licensed under the MIT License. See `LICENSE`.
