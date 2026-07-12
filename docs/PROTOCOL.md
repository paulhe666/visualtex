# VisualTeX Core and adapter protocols

VisualTeX exposes one revisioned Core model through several local transports. All transports carry the same JSON-RPC 2.0 methods and Rust serde representations unless a section explicitly says otherwise.

The current Core protocol version is `1` and is defined by `vt_protocol::PROTOCOL_VERSION`.

## 1. Transports

### 1.1 Stdio Core transport

`visualtex rpc PROJECT_ROOT` uses newline-delimited JSON-RPC 2.0 over stdin/stdout. Exactly one request or response is encoded on each line. Logs are written to stderr and never mixed into the protocol stream.

This transport is used by the VS Code extension and is suitable for a parent process that owns the Core lifetime.

### 1.2 Authenticated local Bridge transport

`visualtex bridge-serve PROJECT_ROOT` starts a persistent project Core and binds a random loopback TCP endpoint under `127.0.0.1`. Discovery metadata is written atomically to:

```text
<project>/.visualtex/bridge/session.json
```

The discovery record contains:

```json
{
  "bridgeProtocolVersion": 1,
  "projectRoot": "/absolute/project/path",
  "endpoint": "127.0.0.1:49321",
  "tokenFile": "/absolute/project/path/.visualtex/bridge/token-....txt",
  "pid": 12345,
  "startedUnixMs": 1780000000000
}
```

Every bridge request is one line with this envelope:

```json
{
  "bridgeProtocolVersion": 1,
  "token": "64 hexadecimal characters",
  "request": {
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": {}
  }
}
```

The response is also one line:

```json
{
  "bridgeProtocolVersion": 1,
  "response": {
    "jsonrpc": "2.0",
    "id": 1,
    "result": {}
  }
}
```

Bridge rules:

- the endpoint must remain loopback-only;
- the token file must remain inside the canonical project bridge directory;
- the token is required before Core dispatch;
- request lines larger than 1 MiB are rejected before JSON parsing;
- only one authenticated bridge may serve a project at a time;
- multiple different projects may run concurrently;
- Core errors remain JSON-RPC errors inside the bridge envelope.

The TypeScript client is in `packages/adapter-sdk`. TeXstudio scripts use the CLI wrappers so they never parse tokens or open sockets directly.

### 1.3 Tauri IPC

The desktop application exposes typed Tauri commands that call the same Core service. TypeScript mirrors are in `packages/protocol-ts`.

### 1.4 OCR sidecar

The Python OCR worker is a separate newline-delimited JSON-RPC process managed by `vt-ocr`. It is not directly exposed to UI code. The Rust manager validates input, model configuration, timeout, project scope and process lifetime.

## 2. Initialization and capabilities

Request:

```json
{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":1}}
```

A current response includes capabilities such as:

```json
{
  "protocolVersion": 1,
  "capabilities": {
    "incrementalEdits": true,
    "incrementalSyntaxTree": true,
    "projectDependencyGraph": true,
    "visualEdits": true,
    "compile": true,
    "toolchainDetection": true,
    "undoRedo": true,
    "pdfiumRendering": true,
    "pdfTiles": true,
    "shadowLayoutMap": true,
    "formulaOcr": false,
    "documentOcr": false,
    "offlineOnly": true
  }
}
```

OCR capability values reflect the active local model packages. Clients must stop when the negotiated protocol version is unsupported and must not infer unavailable capabilities.

## 3. Authoritative text edits

```json
{
  "jsonrpc": "2.0",
  "id": 2,
  "method": "document.applyEdit",
  "params": {
    "operationId": "5f8c0f38-6bc5-4fe2-b356-7f027d32eea3",
    "origin": "sourceEditor",
    "fileId": "6d369d2c-08d7-4c4e-88f2-a22f96772e7d",
    "baseRevision": 3,
    "startByte": 10,
    "endByte": 13,
    "replacement": "中文"
  }
}
```

Offsets are UTF-8 bytes, not UTF-16 code units. An operation is accepted only when:

- `fileId` identifies an open file;
- the byte range lies on UTF-8 boundaries;
- `baseRevision` equals the current buffer revision;
- `operationId` has not already been applied.

A successful response contains the new revision and a `VisualPatch`. A stale operation must be rebased from a fresh snapshot; clients must never fabricate a revision.

## 4. Core method catalogue

### Project and document state

| Method | Purpose |
| --- | --- |
| `project.rootSnapshot` | Read the root document snapshot |
| `project.listFiles` | List `.tex`, `.bib`, `.sty` and `.cls` files |
| `project.openFile` | Open a project-relative file |
| `document.applyEdit` | Apply a revisioned source edit |
| `document.applyVisualEdit` | Serialize a semantic-node content edit to source |
| `document.applyNodeAttributes` | Update supported figure/table attributes |
| `document.undo` / `document.redo` | Traverse the shared Core history |
| `document.save` | Atomically save one Core document |
| `document.confirmSaved` | Confirm a VS Code-owned save only when disk bytes exactly match the Core buffer |
| `project.saveAll` | Save all dirty Core buffers |

### External editors and conflicts

| Method | Purpose |
| --- | --- |
| `project.checkExternalChanges` | Auto-reload clean external changes and return unresolved dirty conflicts |
| `project.refreshFromDisk` | Bridge-safe refresh; fails when any dirty conflict requires explicit resolution |
| `project.resolveExternalConflict` | Reload disk, retain buffer, or save a conflict copy and reload |

`project.refreshFromDisk` is used before TeXstudio bridge compilation and forward SyncTeX. It never overwrites a dirty Core buffer merely to follow an external editor.

### Index, search and refactoring

| Method | Purpose |
| --- | --- |
| `project.index` | Return typed labels, references, citations, bibliography entries, macros and packages |
| `project.dependencies` | Return Tree-sitter-derived include edges, missing targets and detected cycles |
| `project.search` | Search project source including unsaved open buffers |
| `project.previewReplace` | Build a hash-guarded replacement plan |
| `project.previewSymbolRename` | Build a typed label or citation rename plan |
| `project.applyReplace` | Apply a validated multi-file plan transactionally to Core buffers |

### Compilation, PDF and layout

| Method | Purpose |
| --- | --- |
| `project.compile` | Build the real project PDF through the configured local TeX toolchain |
| `toolchain.detect` | Detect local TeX tools |
| `pdf.documentInfo` | Read page geometry and PDF fingerprint through PDFium |
| `pdf.render` | Render a page or tile to the restricted cache |
| `layout.build` | Build a shadow-instrumented source-node-to-PDF layout map |
| `synctex.forwardSearch` | Source line/column to PDF rectangles |
| `synctex.inverseSearch` | PDF page/coordinate to source location |

### OCR

| Method | Purpose |
| --- | --- |
| `ocr.health` | Report formula and document backend health |
| `ocr.recognizeFormula` | Normalize a project image and return ranked LaTeX candidates |
| `ocr.recognizeDocument` | Normalize a project page and return regions, confidences and reading order |

OCR results are proposals. Desktop and VS Code surfaces display them for user correction before source insertion.

## 5. Bridge CLI

The stable bridge-facing CLI commands are:

```bash
visualtex bridge-serve PROJECT_ROOT
visualtex bridge-status PROJECT_ROOT
visualtex bridge-request PROJECT_ROOT METHOD --params '{}'
visualtex bridge-compile PROJECT_ROOT
visualtex bridge-forward-search PROJECT_ROOT SOURCE_FILE LINE COLUMN PDF_PATH
visualtex bridge-inverse-search PROJECT_ROOT PDF_PATH PAGE X Y
visualtex bridge-shutdown PROJECT_ROOT
```

`bridge-compile` refreshes clean external editor changes before compilation and exits nonzero for Core conflicts or compilation failure.

## 6. Versioned `visualtex://` actions

The current URI protocol version is `1`, implemented in Rust by `vt-uri` and in TypeScript by `packages/adapter-sdk`.

Supported forms:

```text
visualtex://open?v=1&project=...
visualtex://forward-search?v=1&project=...&source=...&line=...&column=...&pdf=...
visualtex://inverse-search?v=1&project=...&pdf=...&page=...&x=...&y=...
```

All query values are percent-encoded and support Unicode paths. Unknown actions, missing parameters, non-finite coordinates and unsupported versions are rejected.

The desktop bundle registers the `visualtex` scheme. Deep-link events are queued until the frontend listener is ready. To protect unsaved buffers, a deep link may open a project only when no project is open or when it targets the current project; switching to a different project must go through the desktop UI.

The CLI can execute the same action with:

```bash
visualtex open-uri 'visualtex://...'
```

## 7. Error behavior

Core failures use JSON-RPC `error` objects. Client-visible failures include:

- stale revisions or duplicate operations;
- invalid UTF-8 boundaries;
- project path or symlink scope violations;
- external-save hash mismatch and dirty conflicts;
- missing or timed-out TeX tools;
- restricted shell escape;
- malformed SyncTeX or PDF requests;
- absent OCR models or invalid OCR input;
- bridge authentication, version or message-size failures.

Transport failures may be retried only according to the host policy. Business-level JSON-RPC errors must not be treated as a crashed Core. The VS Code client therefore retries explicit idempotent reads only after a transport reconnection and does not retry ordinary RPC errors.

## 8. OCR worker protocol

The OCR sidecar supports:

- `health`;
- `capabilities`;
- `document.health`;
- `recognizeFormula`;
- `recognizeDocument`;
- `shutdown`.

Before the worker receives an image, Rust verifies the file, size, dimensions, symlink status and project cache scope, decodes it with bounded allocation and writes a SHA-256-deduplicated PNG under `.visualtex/ocr-input`.
