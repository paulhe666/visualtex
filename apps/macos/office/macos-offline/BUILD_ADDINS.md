# Building the real VisualTeX DOTM and PPAM

The repository keeps VBA and Ribbon XML as reviewable source. Do not commit an empty OOXML shell and do not rename versioned add-ins. A valid build must contain a real VBA project created by Microsoft Office for Mac.

The builder, packager, and artifact smoke test inspect the VBA source embedded in
the compiled Word template with `oletools`. Install the build dependency first:

```bash
python3 -m pip install -r scripts/requirements-office-build.txt
```

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
3. Run **Debug → Compile VBAProject**. Confirm `VTWordEvents` compiles, `VTWordRibbonOnLoad` initializes the application event sink, the template does not expose an `AutoExec` startup macro, and the compiled project contains `word-structured-document-import-20260730-r61`, `VTWordRibbonDocumentImport`, `VisualTeX_InsertLatexMarkdownDocument`, `VTHandleWordBeforeDoubleClick`, `App_WindowSelectionChange`, `VTWordRibbonApplyImageFontSizePreset`, and `VTRefreshNumberedImageFormulaFontLayout`.
4. Save the template as exactly `VisualTeX.dotm`.
5. Quit Word before packaging so Office has flushed `vbaProject.bin`.

### Word diagnostic builds

`word/VTWordAdapter.bas` contains production code. The acceptance macros and
their private helpers live in `tests/office/word/VTWordAdapterDiagnostics.bas.inc`
(relative to `apps/macos`), outside the application's bundled resources.

VBA cannot split one module across files. The diagnostic builder combines the
test fragment with the adapter and enables its compile-time diagnostic hooks,
so tests retain access to private functions without making them public:

```bash
node scripts/rebuild_macos_word_addin.mjs --diagnostics \
  --preserve-word --keep-startup-files --modules VTWordAdapter \
  --output /absolute/path/VisualTeXWordBuildDiagnostics.dotm
```

The preserve mode requires Word's existing VBA project access setting to be
enabled. It verifies that the user's open document snapshot is unchanged.
For manual VBE import, generate the diagnostic module instead:

```bash
node scripts/word_vba_source.mjs --diagnostics --output /tmp/VTWordAdapter.bas
```

Use diagnostic templates only in the owned Word acceptance environment. Existing
acceptance scripts that call regression macros require this build; the scripts
that qualify macros with `VisualTeX.dotm!` also require that name in the isolated
acceptance setup. The production packager rejects diagnostic templates. Do not
install a diagnostic template into a user's production Startup folder.

Validate the build profile before running acceptance tests:

```bash
python3 scripts/verify_word_vba_source.py /path/to/diagnostic.dotm --diagnostics
python3 scripts/verify_word_vba_source.py /path/to/production.dotm
```

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
3. Run **Debug → Compile VBAProject**. Confirm `VTPowerPointEvents` compiles, `Auto_Open` initializes the application event sink, and the compiled project contains `powerpoint-office-performance-20260801-r4`, `App_WindowSelectionChange`, and `VTPowerPointRibbonApplyFormulaFontSizePreset`.
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
