# Windows Native Office + OLE implementation log

Date: 2026-07-15

Branch: `feat/windows-native-vsto-ole`

Base and `origin/dev`: `0498e461a719c9ed8f0cfea671a1378caea28427`

## Scope completed

### Protocol and metadata

- Added explicit `nativeOle` and `crossPlatformPicture` object modes across TypeScript, Rust, and C#.
- Added render dimensions and baseline to formula metadata, session storage, and cache records.
- Added validated JSON serialization for OLE structured storage.
- Preserved the existing Session Store instead of introducing a second formula data model.

### Real ATL OLE LocalServer

New projects:

- `src-windows/VisualTeX.FormulaOleServer/`
- `src-windows/VisualTeX.FormulaOleServer.Tests/`

Implemented:

- ATL EXE LocalServer with fixed ProgID `VisualTeX.Formula.1`.
- Fixed CLSID, IID, AppID, and TypeLib identities documented in the architecture specification.
- `IOleObject`, `IDataObject`, `IPersistStorage`, `IViewObject2`, and a fixed-IID dual Automation `IVisualTeXFormulaObject` with DispIds 1–3.
- Cross-process custom-interface marshaling through an embedded Automation-compatible type library and the universal Automation marshaler.
- Standard compound-document registration with `InprocHandler32=Ole32.dll`, `ServerExecutable`, `AuxUserType`, and `DataFormats`.
- Startup grace locking so the ATL EXE class factory remains available while Office's default handler prepares object creation.
- Safe placeholder JSON/vector-EMF/PNG persistence for PowerPoint's create-save-reload sequence before formula initialization.
- Immediate transactional `IOleClientSite::SaveObject` after initialize/update, with in-memory rollback if the container save fails.
- Compound-storage streams `VisualTeX.Formula.json`, `VisualTeX.Preview.emf`, and `VisualTeX.Preview.png`.
- Transactional initialization/update, view/data advise notifications, cached drawing while VisualTeX is closed, and external-editor verbs.
- Per-user COM registration and clean per-user unregistration.

### True vector EMF

Files:

- `src-windows/VisualTeX.WindowsOffice.VstoShared/OfficeOlePreview.cs`
- `src-windows/VisualTeX.WindowsOffice.Tests/OfficeOlePreviewTests.cs`
- `src-windows/VisualTeX.WindowsOffice.Tests/VisualTeXSessionVectorExportTests.cs`

Implemented:

- Self-contained MathJax SVG materialization from the existing Session result.
- Controlled path-only SVG renderer for `defs`, `path`, `use`, `g`, `rect`, `line`, polygons, circles, ellipses, and supported transforms.
- True vector EMF output.
- Fail-closed rejection of raster images, external resources, scripts, text not converted to paths, unsupported visible nodes, unsupported transparency, and raster EMF/EMF+ records.
- No PNG-inside-EMF fallback in native OLE mode.

### Word and PowerPoint VSTO

Files include:

- `src-windows/VisualTeX.WordVsto/ThisAddIn.cs`
- `src-windows/VisualTeX.WordVsto/WordFormulaService.cs`
- `src-windows/VisualTeX.PowerPointVsto/ThisAddIn.cs`
- `src-windows/VisualTeX.PowerPointVsto/PowerPointFormulaService.cs`
- `src-windows/VisualTeX.WindowsOffice.VstoShared/OlePngPreviewExtractor.cs`
- `src-windows/VisualTeX.WindowsOffice.VstoShared/WordEquationNumbering.cs`
- `src-windows/VisualTeX.WindowsOffice.VstoShared/WordFormulaMetadataReader.cs`
- `src-windows/VisualTeX.WindowsOffice.VstoShared/WordOleObjectAccessor.cs`
- `src-windows/VisualTeX.WindowsOffice.Contracts/FormulaOleInterop.cs`

Implemented:

- Word `InlineShapes.AddOLEObject` and PowerPoint `Shapes.AddOLEObject` insertion using `VisualTeX.Formula.1`.
- HRESULT-checked custom-interface initialization and in-place update using JSON + EMF + PNG.
- Embedded OLE JSON is the only native formula metadata source; Word Title/AlternativeText and PowerPoint Tags remain picture-mode compatibility metadata.
- Persisted Word objects are made running with the non-editor `SHOW` verb before the Automation interface is requested.
- Transactional migration of a selected legacy picture formula to native OLE: the old picture is deleted only after the candidate OLE object is initialized and configured.
- Native OLE double-click in Word is no longer intercepted by the legacy-picture event path; Office can invoke `IOleObject::DoVerb` normally.
- Native Ribbon commands for create, edit, selected conversion, selected deletion, selected OLE-to-picture export, and opening VisualTeX.
- Word equation-number reconciliation using `SEQ VisualTeXEquation`, persistent UUID bookmarks, centered/right tab stops, orphan cleanup, and vertical alignment.
- Transactional OLE-to-picture export by reading the OLE object's built-in PNG `IDataObject` cache without changing the system clipboard.
- Legacy picture methods remain available for explicit cross-platform output and compatibility.

Not implemented:

- Document-wide “convert all legacy pictures” command. Selected-object migration is complete. Safe bulk conversion still needs a queued/headless vector-render workflow so multiple formulas do not open overlapping editor sessions.

### x86/x64 build and installer

Files include:

- `src-windows/VisualTeX.WindowsOffice.Installer/Package.wxs`
- `src-windows/VisualTeX.WindowsOffice.Installer/VisualTeX.WindowsOffice.Installer.wixproj`
- `scripts/build_windows_office.ps1`
- `scripts/install_windows_vsto.ps1`
- `scripts/uninstall_windows_vsto.ps1`
- `src-tauri/tauri.windows.conf.json`
- `src-tauri/windows/hooks.nsh`

Implemented:

- x86 and x64 VSTO builds.
- Win32 and x64 ATL LocalServer builds.
- Separate x86 and x64 MSI packages with independent ProductCodes and a shared UpgradeCode.
- Correct `win32`/`win64` TypeLib registration and explicit Registry32/Registry64 verification.
- Office-bitness detection during installation.
- SHA-256 verification for the MSI, Word add-in, PowerPoint add-in, and LocalServer.
- Native NSIS installation path that does not start Word or PowerPoint and does not use UI Automation, cursor movement, keystroke simulation, Office add-in dialogs, or Ribbon-cache polling.
- Existing Office.js trusted-catalog scripts retained as compatibility/cleanup code; they are not called by the native production installation path.

Final package hashes from the last successful build:

- x64 MSI: `4E7B901312DB21604F10365BF39FEC8DD05CDE41B7459913651D8D35E3CDCC47`
- x86 MSI: `AEC83F23B6C7B84A6F1943FFB753B27531F1F7679213D8387D85A664350C47F4`
- x64 Word add-in: `DE98B0D7E2A973D4D7C3FFD9F70E3CCF83A072949F0366B0BD1F18366A2E87BD`
- x64 PowerPoint add-in: `F77325DB7A24071EB7CE121EEC40DE57D0A1C23FDCE54C2202E7BDC80AB9A576`
- x64 LocalServer: `5D1E5D4217F9B244F0E6CD6242BAF580BBAEE64B865E7ABEA1AE0BB389A8F9F3`
- x86 Word add-in: `BDEC9D99BD21096AF98D3389C69E458A077C1D27B5CDBA2933D28E2671DD0581`
- x86 PowerPoint add-in: `D6F069CB26E69D6A2DB187DFBEB630C925F4351714E08C4C8BCBB84B5BCA4F54`
- Win32 LocalServer: `09ED46E13D45196260B79EC0481A101EA90BBE5D9883DFAE43D5E29387590F85`

## Commands and evidence

### TypeScript and Node

```powershell
npx tsc --noEmit
npm run test:office-metadata
npm run test:windows-office-architecture
npm run build:office:windows-ole
```

Results:

- TypeScript check passed.
- Office metadata smoke test passed.
- Windows Office architecture smoke test passed.
- Independent Windows Office frontend production build passed. Vite emitted only its existing large-chunk advisory.

### C# tests

```powershell
C:\Users\pojian_liao\AppData\Local\Microsoft\dotnet\dotnet.exe test src-windows\VisualTeX.WindowsOffice.Tests\VisualTeX.WindowsOffice.Tests.csproj --configuration Release
```

Result: 61 passed, 0 failed, 0 skipped.

Coverage includes protocol security, STA dispatch, transactional replacement, metadata, persistent UUIDs, COM release, double-click deduplication, PowerPoint sizing, fixed OLE ABI, vector SVG-to-EMF validation/fail-closed behavior, Session SVG preservation, and Word VSTO numbering rules.

### Independent real COM/OLE tests

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\test_windows_formula_ole_server.ps1 -Configuration Release
```

Result:

- x64 LocalServer and x64 test container passed.
- Win32 LocalServer and Win32 test container passed.
- Verified per-user registration/unregistration, `CoCreateInstance`, cross-process custom-interface marshaling, transactional update failure, structured-storage save/load, all three streams, `IDataObject`, and `IViewObject2` drawing.
- Verified standard `OleCreate` with both `OLERENDER_NONE` and `OLERENDER_DRAW`.
- Verified placeholder save/load followed by real initialization.
- Latest saved log: `src-windows/artifacts/test-logs/formula-ole-server-20260715-182656.log`.

### Full dual-platform package build

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\build_windows_office.ps1 -Configuration Release -SkipTests
```

Result:

- C# and x86/x64 real COM test gates passed separately immediately before packaging.
- x86/x64 VSTO builds passed.
- Win32/x64 LocalServer builds passed.
- x86/x64 WiX MSI builds passed.
- Tauri/NSIS Office resources and SHA manifests were refreshed.
- Remaining build messages are the known Office Interop embed warning and nonfatal Visual Studio workload-resolver text; no build error occurred.

### Rust

```powershell
cargo check --manifest-path src-tauri\Cargo.toml --tests
```

Result: passed with ten existing unused/dead-code warnings in non-macOS platform stubs.

```powershell
cargo test --manifest-path src-tauri\Cargo.toml office:: --lib
```

Result: Rust source and test executable compiled, but the test process failed before entering the harness with Windows status `0xc0000139` (`STATUS_ENTRYPOINT_NOT_FOUND`). This is recorded as a local dynamic-loader environment blocker; Rust tests are not claimed as passed.

## Real Office native OLE acceptance

Environment:

- Microsoft Office ProPlus 2024 Volume, x64.
- Office version `16.0.17932.20162`.
- Real desktop Word and PowerPoint COM automation; no mocked Office object model.

Acceptance project:

- `src-windows/VisualTeX.NativeOfficeOleAcceptance/`

Final result: passed.

Verified in seven stages:

1. Per-user ATL LocalServer registration.
2. Real Word `InlineShapes.AddOLEObject`, dual-interface initialization, DOCX save, and object shutdown.
3. Real PowerPoint `Shapes.AddOLEObject`, placeholder save/reload, real initialization, and PPTX save.
4. LocalServer unregistration followed by offline Word/PPT reopen with valid cached previews.
5. Re-registration followed by persisted Word object activation through the non-editor SHOW verb, JSON/EMF/PNG update, save, and close.
6. Persisted PowerPoint JSON/EMF/PNG update and save.
7. Final unregistered offline reopen and cleanup verification.

Final artifact directory:

- `src-windows/artifacts/test-logs/native-office-ole-20260715-183541`

The same seven-stage acceptance also passed in the immediately preceding wrapper run:

- `src-windows/artifacts/test-logs/native-office-ole-20260715-182729`

The reusable entry point is `scripts/test_windows_native_office_ole.ps1`; it detects Office bitness, refuses to run over an active Word/PowerPoint session, rebuilds the matched LocalServer and acceptance executable, enforces a timeout, and verifies registration/process cleanup.

## Real MSI and VSTO load acceptance

Final x64 MSI installation was executed through `scripts/install_windows_vsto.ps1` using the generated SHA manifest.

Verified:

- MSI install exit code `0`.
- MSI, Word add-in, PowerPoint add-in, and LocalServer hashes all matched the generated manifest.
- Installed LocalServer and COM registrations pointed to the per-user install directory.
- Word reported `VisualTeX.WordVsto` with `COMAddIn.Connect=True`.
- PowerPoint reported `VisualTeX.PowerPointVsto` with `COMAddIn.Connect=True`.
- `scripts/uninstall_windows_vsto.ps1` removed the MSI, install directory, Word registration, PowerPoint registration, OLE CLSID/ProgID/interface/type-library registration, and LocalServer process.
- The post-uninstall check found no MSI product, Office/OLE process, diagnostic marker, or native registration residue.

Final MSI log:

- `%LOCALAPPDATA%\VisualTeX\office\install-logs\vsto-install-20260715-183200.log`

## Remaining release-qualification gates

The implementation is validated on this x64 Office 2024 machine. The following broader matrix items remain outside this single-workspace run:

- Real x86 Office installation and in-process VSTO loading on an x86 Office machine.
- MSI repair/major-upgrade testing from an older released product version.
- Ten cold Office restart cycles and manual Ribbon visual inspection.
- Chinese and English Office UI localization verification.
- Large-document/manual interaction matrix covering read-only files, many windows, slide-show editing attempts, undo/redo, forced LocalServer crash/restart, and hundreds of formulas.
- Manual native double-click editor launch with a fully installed VisualTeX desktop protocol handler.

These are release-matrix tasks rather than known implementation failures.
