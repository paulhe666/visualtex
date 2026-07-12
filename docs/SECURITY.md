# Security model

## Trust boundary

VisualTeX treats every opened LaTeX project, model package, image, URI and configuration file as untrusted input. The Rust Core owns authoritative source mutation. UI and editor adapters submit typed, revisioned operations and do not receive unrestricted filesystem or process execution primitives.

## Restricted compilation

Restricted mode is enabled by default.

- `shell-escape` is rejected even when project configuration asks for it.
- TeX executables are launched directly with `Command`; source text is never interpolated into a shell command.
- Build output must be a relative path without parent-directory traversal.
- TeX receives conservative `openout_any=p` and `openin_any=a` values.
- Builds have a deadline, use isolated `.visualtex/build` output and terminate with their owning task.
- The Bridge and URI layers invoke fixed Core methods; they do not evaluate shell fragments.

A future trusted mode must require an explicit, per-project decision and must not silently transfer trust to a moved or replaced directory.

## Filesystem scope

- Project-relative paths cannot be absolute or contain `..`.
- Source files are canonicalized and must remain below the canonical project root.
- Symlinked source files are rejected.
- PDF, SyncTeX, layout-map and OCR operations are constrained to the project and validated VisualTeX caches.
- Tauri asset access is limited to build, PDF cache and OCR input paths; recovery content is denied.
- Multi-file replacement plans carry expected SHA-256 hashes and are rejected when stale or truncated.

## Recovery, saves and external editors

Every dirty Core buffer has a local recovery snapshot containing its source path, saved-disk SHA-256 baseline, latest complete UTF-8 text and revision.

On reopen, automatic recovery occurs only when disk still matches the recorded baseline. If another editor changed the file, VisualTeX preserves the disk version and writes the unsaved buffer under `.visualtex/recovery/conflicts`.

External-editor integration follows the same rule:

- clean disk changes may be reloaded;
- a VS Code save is confirmed only if the current disk bytes exactly match the Core buffer;
- TeXstudio bridge refresh fails when Core and disk are both modified;
- no refresh operation silently overwrites a dirty buffer.

## OCR and model packages

OCR is offline by design.

- The worker contains no automatic model downloader or network client.
- Model packages are inspected and verified before installation and are selected explicitly by capability.
- Imported images must be regular non-symlink files no larger than 64 MiB.
- Decode limits cap dimensions at 12,000 pixels per axis and bounded allocation at 256 MiB.
- Input is normalized to SHA-256-deduplicated PNG under `.visualtex/ocr-input` before worker access.
- Worker requests have deadlines and the process can be restarted.
- OCR output is shown for review; it is not silently committed as authoritative LaTeX.

The Python worker also retains its own input limits as defense in depth.

## Authenticated local Bridge

The TeXstudio/general Bridge is reachable only on the local machine, but loopback is not treated as authentication by itself.

- It binds a random `127.0.0.1` port, never a wildcard interface.
- Each session creates a 64-character random token.
- On Unix, the token file is mode `0600`.
- Discovery and token paths are verified against the canonical project `.visualtex/bridge` directory.
- A constant-time token comparison occurs before Core dispatch.
- Bridge protocol version is checked independently of Core protocol version.
- A request line larger than 1 MiB is rejected before JSON parsing; clients also cap responses.
- A startup lock prevents concurrent creation races, and a second authenticated server for the same project is rejected.
- Stale discovery is removed only after the authenticated endpoint cannot be reached.
- Different projects receive different endpoints and tokens.
- `bridge.shutdown` requires the same authenticated envelope.

The Node Adapter SDK implements the same discovery, endpoint, token-path, size, timeout and response-ID checks. TeXstudio scripts call fixed CLI wrappers and never read the token directly.

## VS Code

- Core starts only for a trusted, local workspace.
- VS Code `TextDocument` remains authoritative for editor-visible source.
- Visual edits are written with `WorkspaceEdit`, preserving native undo/redo and compatibility with other extensions.
- Core restart recovery resynchronizes from the current `TextDocument` rather than a hidden editable copy.
- JSON-RPC business errors are not retried as transport failures.
- OCR selection accepts only files inside the trusted workspace before Rust performs normalized project-cache import.
- VS Code Remote is rejected until a separately designed remote trust and transport model exists.

## Desktop deep links

The desktop bundle registers `visualtex://` through the Tauri deep-link plugin.

- URI protocol version and action names are strict.
- Required parameters are parsed and percent-decoded by typed Rust code.
- Non-finite coordinates and unsupported actions are rejected.
- Initial links are queued until frontend listeners are ready.
- A link may open a project only if no project is open or it targets the already open project.
- A link cannot silently discard unsaved buffers by switching to another project.
- PDF and source paths still pass through normal Core project-scope checks.

## Desktop WebView

- The content security policy denies arbitrary remote content.
- The asset protocol has explicit allow and deny scopes.
- Tauri commands are typed and do not expose a generic filesystem or shell command endpoint.
- OCR and deep-link events contain structured values, not executable markup.

## Current limitations

- Release signing, notarization, updater verification and installer reputation workflows are not yet complete.
- A production plugin sandbox and permission manifest are not yet implemented; third-party arbitrary plugins are therefore not loaded.
- Windows PowerShell adapter scripts are covered by static review and CI parsing when available, but runtime behavior still requires native Windows release validation.
- Bundled production OCR model supply-chain policy and reproducible model artifacts remain release work.

## Reporting

Security reports should include the affected revision, operating system, minimal redacted project, exact command, expected behavior and observed behavior. Do not attach private papers, model files or OCR images unless they have been scrubbed of confidential content.
