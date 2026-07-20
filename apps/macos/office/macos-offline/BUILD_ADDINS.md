# Building the real VisualTeX DOTM and PPAM

The repository keeps VBA and Ribbon XML as reviewable source. Do not commit an empty OOXML shell and do not rename versioned add-ins. A valid build must contain a real VBA project created by Microsoft Office for Mac.

## Word: VisualTeX.dotm

1. Open Microsoft Word for Mac and create a blank macro-enabled template.
2. Open the Visual Basic Editor and import these modules without changing their module names:
   - `shared/VTProtocol.bas`
   - `shared/VTOfficePaths.bas`
   - `shared/VTMetadata.bas`
   - `shared/VTLauncher.bas`
   - `shared/VTErrorHandling.bas`
   - `word/VTWordAdapter.bas`
   - `word/VTWordEvents.cls`
   - `word/VTRibbonCallbacks.bas`
3. Run **Debug → Compile VBAProject**. Confirm `VTWordEvents` compiles and `AutoExec` initializes the application event sink.
4. Save the template as exactly `VisualTeX.dotm`.
5. Quit Word before packaging so Office has flushed `vbaProject.bin`.

## PowerPoint: VisualTeX.ppam

1. Open Microsoft PowerPoint for Mac and create a blank macro-enabled presentation.
2. Import these modules without changing their module names:
   - `shared/VTProtocol.bas`
   - `shared/VTOfficePaths.bas`
   - `shared/VTMetadata.bas`
   - `shared/VTLauncher.bas`
   - `shared/VTErrorHandling.bas`
   - `powerpoint/VTPowerPointAdapter.bas`
   - `powerpoint/VTPowerPointEvents.cls`
   - `powerpoint/VTRibbonCallbacks.bas`
3. Run **Debug → Compile VBAProject**. Confirm `VTPowerPointEvents` compiles and `Auto_Open` initializes the application event sink.
4. Save as a PowerPoint Add-In with the exact filename `VisualTeX.ppam`.
5. Quit PowerPoint before packaging.

## Inject and verify Ribbon XML

Run:

```bash
node scripts/package_macos_offline_addins.mjs \
  --word /absolute/path/VisualTeX.dotm \
  --powerpoint /absolute/path/VisualTeX.ppam
```

The packager performs all non-UI packaging work. It verifies the macro project and expected module names, injects the reviewed `customUI14.xml` relationship, validates the result, copies the fixed files to `resources/`, and writes `addins.json` with SHA-256 hashes.

## Required validation

```bash
npm run test:macos-offline-office
cargo test --manifest-path src-tauri/Cargo.toml --lib
npm run build:desktop
```

Then install through VisualTeX Settings. Word is not launched by the installer: launch Word manually and wait for `OfficePluginStatus/word.json` before treating Word installation as healthy. PowerPoint requires one manual registration through **Tools → PowerPoint Add-Ins**; future updates overwrite the same fixed `VisualTeX.ppam` path.
