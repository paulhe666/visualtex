# VisualTeX Next — TeXstudio Bridge

This adapter connects TeXstudio to the same versioned VisualTeX Core used by the desktop app and VS Code extension. It does not patch TeXstudio and does not modify TeXstudio configuration automatically.

## What the bridge provides

- one authenticated local bridge per project;
- loopback-only random TCP port (`127.0.0.1`);
- per-session token stored under `<project>/.visualtex/bridge/`;
- JSON-RPC protocol and capability negotiation;
- persistent project parsing instead of starting a new Core for every command;
- safe refresh of TeXstudio-saved source files before compilation;
- real TeX compilation and PDF output;
- forward and inverse SyncTeX;
- opening the current project in the VisualTeX desktop application;
- Chinese, Unicode, spaces, and multiple simultaneous projects.

The bridge refuses an external disk refresh when its own Core contains conflicting unsaved edits. It never silently overwrites either side.

## Prerequisites

Install the VisualTeX CLI and desktop application. For a source checkout:

```bash
cargo install --path apps/cli
```

The CLI looks for the desktop executable in this order:

1. `VISUALTEX_DESKTOP_BIN`;
2. `visualtex-desktop` beside the `visualtex` CLI;
3. the installed `VisualTeX Next` application bundle on macOS;
4. `visualtex-desktop` on `PATH` on Windows/Linux.

Set `VISUALTEX_BIN` when the CLI executable is not named `visualtex` or is outside `PATH`.

## Install on macOS or Linux

```bash
sh scripts/texstudio/install.sh
```

Default destination:

```text
~/.local/share/visualtex/texstudio/bin
```

Choose another destination with `--prefix PATH`. The installer refuses to overwrite a different existing adapter unless `--force` is supplied. It does not edit TeXstudio preferences.

Uninstall only the marked VisualTeX adapter files:

```bash
sh scripts/texstudio/uninstall.sh
```

## Install on Windows

Run PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\texstudio\windows\Install-VisualTeXTexstudio.ps1
```

Default destination:

```text
%LOCALAPPDATA%\VisualTeX\texstudio\bin
```

Use `-Prefix <path>` to select another directory and `-Force` only to replace an older VisualTeX adapter. The installer does not edit TeXstudio preferences.

## TeXstudio user commands

Open **Options → Configure TeXstudio → Build → User Commands** and add the commands you need. TeXstudio placeholder names vary by release and local configuration, so substitute your version's placeholders for the capitalized arguments below. Keep every path argument quoted.

### macOS / Linux

```text
Open VisualTeX
sh "/INSTALL/bin/visualtex-open.sh" "PROJECT_ROOT"

Compile through VisualTeX
sh "/INSTALL/bin/visualtex-compile.sh" "PROJECT_ROOT"

Start bridge
sh "/INSTALL/bin/visualtex-bridge-start.sh" "PROJECT_ROOT"

Bridge status
sh "/INSTALL/bin/visualtex-bridge-status.sh" "PROJECT_ROOT"

Stop bridge
sh "/INSTALL/bin/visualtex-bridge-stop.sh" "PROJECT_ROOT"

Forward SyncTeX
sh "/INSTALL/bin/visualtex-forward-search.sh" "PROJECT_ROOT" "CURRENT_SOURCE_FILE" LINE COLUMN "PDF_PATH" "OUTPUT_JSON"

Inverse SyncTeX
sh "/INSTALL/bin/visualtex-inverse-search.sh" "PROJECT_ROOT" "PDF_PATH" PAGE X Y "OUTPUT_JSON"
```

### Windows

```text
Open VisualTeX
powershell -ExecutionPolicy Bypass -File "C:\INSTALL\VisualTeX-Open.ps1" -ProjectRoot "PROJECT_ROOT"

Compile through VisualTeX
powershell -ExecutionPolicy Bypass -File "C:\INSTALL\VisualTeX-Compile.ps1" -ProjectRoot "PROJECT_ROOT"

Start bridge
powershell -ExecutionPolicy Bypass -File "C:\INSTALL\VisualTeX-BridgeStart.ps1" -ProjectRoot "PROJECT_ROOT"

Bridge status
powershell -ExecutionPolicy Bypass -File "C:\INSTALL\VisualTeX-BridgeStatus.ps1" -ProjectRoot "PROJECT_ROOT"

Stop bridge
powershell -ExecutionPolicy Bypass -File "C:\INSTALL\VisualTeX-BridgeStop.ps1" -ProjectRoot "PROJECT_ROOT"

Forward SyncTeX
powershell -ExecutionPolicy Bypass -File "C:\INSTALL\VisualTeX-ForwardSearch.ps1" -ProjectRoot "PROJECT_ROOT" -SourceFile "CURRENT_SOURCE_FILE" -Line LINE -Column COLUMN -PdfPath "PDF_PATH" -OutputJson "OUTPUT_JSON"

Inverse SyncTeX
powershell -ExecutionPolicy Bypass -File "C:\INSTALL\VisualTeX-InverseSearch.ps1" -ProjectRoot "PROJECT_ROOT" -PdfPath "PDF_PATH" -Page PAGE -X X -Y Y -OutputJson "OUTPUT_JSON"
```

The PowerShell adapters force UTF-8 when decoding native CLI output and replace `OUTPUT_JSON` only after a successful command, so Chinese and space-containing paths remain valid under Windows PowerShell 5. TeXstudio or a companion macro can read the generated JSON and move the cursor/viewer using the returned source location or PDF rectangles.

The adapter runtime has been exercised on Windows 10 with PowerShell 5.1, TeXstudio 4.8.9 and TeX Live 2025 using a project path containing Chinese characters and spaces. This validates the scripts, Bridge, compilation and bidirectional SyncTeX; TeXstudio-version-specific placeholder and GUI macro profiles still require separate host-level validation.

## Save and synchronization behavior

TeXstudio remains authoritative for its unsaved editor buffer. Save the document before invoking a bridge command. `visualtex-compile` and forward SyncTeX call `project.refreshFromDisk` before continuing:

- clean externally changed files are reloaded into the bridge Core;
- conflicting dirty Core buffers cause a visible non-zero error;
- no source file is overwritten merely to perform a refresh.

The VisualTeX desktop app independently watches `.tex`, `.bib`, `.sty`, and `.cls` files. Once TeXstudio saves, the desktop app reloads clean buffers or opens its conflict-resolution dialog. Once VisualTeX saves, TeXstudio's normal external-file-change handling sees the update.

## Direct CLI usage

```bash
visualtex bridge-serve PROJECT_ROOT
visualtex bridge-status PROJECT_ROOT
visualtex bridge-request PROJECT_ROOT initialize --params '{}' --result-only
visualtex bridge-compile PROJECT_ROOT
visualtex bridge-forward-search PROJECT_ROOT SOURCE_FILE LINE COLUMN PDF_PATH
visualtex bridge-inverse-search PROJECT_ROOT PDF_PATH PAGE X Y
visualtex bridge-shutdown PROJECT_ROOT
visualtex open PROJECT_ROOT
```

## Security model

- The server binds only to loopback and uses a random port.
- A 64-character random token is required before Core RPC dispatch.
- On Unix, the token file is mode `0600`.
- Discovery and token files are scoped to the current project's `.visualtex/bridge` directory.
- A request line is rejected before JSON parsing when it exceeds 1 MiB.
- A second bridge for the same active project is rejected.
- Core path checks continue to confine source, PDF, SyncTeX, and OCR access to the project and its validated caches.
- The adapter never evaluates shell text received through JSON-RPC.

## Troubleshooting

Bridge logs are written to:

```text
<project>/.visualtex/bridge/texstudio-bridge.log
```

Check the session:

```bash
visualtex bridge-status PROJECT_ROOT
```

A stale session is cleaned automatically only after the authenticated endpoint is confirmed unreachable. To stop a healthy session, use `bridge-shutdown` rather than deleting metadata files manually.
