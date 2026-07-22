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

## PowerPoint: compile a PPTM, then package VisualTeX.ppam

PowerPoint for Mac loads `.ppam` files as add-ins and does not provide a reliable direct **Save As PPAM** workflow. Compile the VBA project in a temporary macro-enabled presentation, then inject that compiled VBA project into the reviewed PPAM shell with the repository packager.

1. Open Microsoft PowerPoint for Mac and create a blank macro-enabled presentation (`.pptm`).
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
4. Save the editable build input as a `.pptm` file and quit PowerPoint so Office flushes `ppt/vbaProject.bin`.
5. Keep a known-good `VisualTeX.ppam` as the package shell. Do not attempt to open or edit the PPAM directly on macOS.

## Inject and verify Ribbon XML

Run:

```bash
node scripts/package_macos_offline_addins.mjs \
  --word /absolute/path/VisualTeX.dotm \
  --powerpoint /absolute/path/VisualTeX-build.pptm \
  --powerpoint-shell /absolute/path/known-good-VisualTeX.ppam
```

The packager performs all non-UI packaging work. For PowerPoint it extracts the compiled `ppt/vbaProject.bin` from the PPTM, injects it into the reviewed PPAM shell, restores the reviewed `customUI14.xml` relationship and images, validates the result, copies the fixed files to `resources/`, and writes `addins.json` with SHA-256 hashes.

## Required validation

```bash
npm run test:macos-offline-office
cargo test --manifest-path src-tauri/Cargo.toml --lib
npm run build:desktop
```

Then install through VisualTeX Settings. Word is not launched by the installer: launch Word manually and wait for `OfficePluginStatus/word.json` before treating Word installation as healthy. PowerPoint requires one manual registration through **Tools → PowerPoint Add-Ins**; future updates overwrite the same fixed `VisualTeX.ppam` path.
