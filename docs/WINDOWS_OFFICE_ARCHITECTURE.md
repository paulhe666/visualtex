# VisualTeX Office integration architecture

VisualTeX keeps the formula editor, OCR, formula metadata/cache and Session API in one cross-platform core, while each Office integration owns its platform-specific installation and native Office control code.

## Runtime boundaries

### Shared core

- React formula editor and Office editor window: `src/office/shared/`
- Stable Session/OCR/cache service: `https://127.0.0.1:43127`
- Formula UUID and compressed metadata schema: `visualtex-formula` v1
- Rust companion HTTP server and session store: `src-tauri/src/office/`

The shared layer contains no AppleScript, COM, VSTO, LaunchAgent, Keychain or Windows registry installation code.

### macOS Office.js integration

- TypeScript bridge: `src/office/macos/`
- Manifests: `office/macos/manifests/`
- Build output: `dist-office-macos/`
- Native control: AppleScript and the existing macOS PowerPoint adapter
- Installation: Word/PowerPoint `wef` directories, login Keychain and LaunchAgent

The macOS build does not compile or bundle C#, COM, VSTO or Windows registry code.

### Windows OLE integration

- Office.js bridge: `src/office/windows-ole/`
- Manifests: `office/windows/ole/manifests/`
- Build output: `dist-office-windows-ole/`
- Rust transport: `src-tauri/src/office/windows_pipe.rs`
- C# sidecar: `src-windows/VisualTeX.WindowsOleBridge/`

Runtime path:

```text
Windows Office.js Ribbon
  -> VisualTeX HTTPS companion
  -> \\.\pipe\VisualTeX.OfficeBridge.<CurrentUserSid>
  -> C# OLE Bridge STA message loop
  -> Word / PowerPoint COM
```

The pipe ACL permits only the current user. Every sidecar launch uses a new 256-bit token and requires a token handshake. PNG files are materialized only under `%LOCALAPPDATA%\VisualTeX\office\temp`; image Base64 is never sent over the pipe.

All Office COM calls run on one STA thread with a Windows message loop. Document, Presentation, Slide, Shape, Range and InlineShape proxies are acquired per operation and released in `finally` blocks. A 30-second request watchdog terminates a stuck sidecar so the Rust backend can restart it on the next request.

### Windows native Word/PowerPoint integration

- Shared contracts: `src-windows/VisualTeX.WindowsOffice.Contracts/`
- Word add-in: `src-windows/VisualTeX.WordVsto/`
- PowerPoint add-in: `src-windows/VisualTeX.PowerPointVsto/`
- Per-user MSI: `src-windows/VisualTeX.WindowsOffice.Installer/`

The native add-ins own Ribbon callbacks, Office events, selection tracking and double-click editing. They do not implement OCR or formula rendering. They create a VisualTeX Session, open the local editor, wait for the Session result and then insert or replace the target Office object on the Office STA thread.

## Persistent formula identity and replacement

Every formula receives one UUID v4. The UUID is retained for the full lifetime of the Office object.

PowerPoint stores:

- object name: `VisualTeX_<formulaId>`
- `Shape.Tags["VisualTeXFormulaId"]`
- compressed metadata in `Shape.Tags["VisualTeXMetadata"]`
- the same compressed metadata in `AlternativeText`

Word stores compressed metadata in both `InlineShape.Title` and `InlineShape.AlternativeText`. Word formulas are not wrapped in Content Controls.

Replacement is transactional:

1. Locate the target by persistent formula UUID.
2. Record position, size, rotation and z-order where applicable.
3. Insert and fully configure the replacement image.
4. Write the object name and metadata.
5. Delete the old object only after every preceding step succeeds.
6. Delete the candidate and keep the old object on any failure.

Image fitting preserves the PNG's natural aspect ratio. PowerPoint centers the replacement inside the old bounding box. Word display formulas use an independent centered paragraph; inline formulas remain InlineShapes.

## Windows integration mode

The runtime retains `Automatic`, `OLE` and `VSTO` compatibility for development and migration, but the production Windows installer ships and enables only OLE. VSTO sources and their self-hosted acceptance path remain available for future work.

- `Automatic`: use healthy Word and PowerPoint native add-ins when both are enabled; otherwise fall back to OLE.
- `OLE`: register the current-user Trusted Catalog and disable native add-in `LoadBehavior`.
- `VSTO`: remove the OLE Trusted Catalog before enabling the native add-ins.

Both backends share the same current-user HTTPS certificate and Session protocol. They never load two VisualTeX ribbons at the same time.

## Build commands

```powershell
npm ci
npm run build:office:windows-ole
./scripts/build_windows_ole_bridge.ps1
```

The normal Tauri/NSIS build uses the Office.js bundle and self-contained OLE sidecar only. To exercise the deferred VSTO/MSI development path on a machine with Office installed:

```powershell
./scripts/build_windows_office.ps1
```

`build_windows_office.ps1` performs the .NET tests, publishes the self-contained OLE sidecar and builds the Word/PowerPoint add-ins plus WiX MSI. The sidecar copied for Tauri is:

```text
src-tauri/binaries/visualtex-windows-office-bridge-x86_64-pc-windows-msvc.exe
```

macOS remains independent:

```bash
npm ci
npm run build:office:macos
npm run tauri:build
```

## Installation scripts

- `scripts/install_windows_ole.ps1`
- `scripts/uninstall_windows_ole.ps1`
- `scripts/install_windows_vsto.ps1`
- `scripts/uninstall_windows_vsto.ps1`
- `scripts/ensure_windows_office_certificate.ps1`
- `scripts/remove_windows_office_certificate.ps1`

OLE manifests are copied to `%LOCALAPPDATA%\VisualTeX\OfficeCatalog` and registered as a current-user Trusted Catalog. The production NSIS installer offers VisualTeX only or VisualTeX + OLE, with OLE selected by default. Legacy VisualTeX VSTO products are removed when possible and always disabled before the OLE manifest is registered.

## Automated verification

Platform-independent and build checks:

```bash
npm run test:office-platform-boundaries
npm run test:windows-office-architecture
npm run verify:office-manifest
npm run test:editor:run
cargo test --manifest-path src-tauri/Cargo.toml --lib
```

Windows unit tests cover pipe authentication, the single STA dispatcher, COM release, UUID metadata, temp-path enforcement, failed replacement rollback and double-click deduplication.

Real Office acceptance requires a disposable interactive Windows session with licensed desktop Word and PowerPoint:

```powershell
./scripts/run_windows_office_acceptance.ps1 -FormulaCount 20
```

The acceptance harness inserts at least 20 formulas per host, randomly edits formulas without changing the UUID set or object count, exercises rapid repeated operations, source document/slide switching, deletion, undo/redo, read-only documents, PowerPoint slide show mode, multiple windows, Office-not-running behavior and forced sidecar restart. The manual workflow is `.github/workflows/windows-office-acceptance.yml` and requires a self-hosted runner labeled `visualtex-office`.
