# macOS native offline Office acceptance

Record the VisualTeX version, Microsoft Office version, macOS version, machine architecture, test account, and artifact SHA-256 values before running this checklist. Save console output under ignored `build-logs/macos-offline/`.

## Automated source and runtime checks

```bash
npm run test:macos-offline-office
cargo test --manifest-path src-tauri/Cargo.toml --lib
npm run build:desktop
```

Expected: all commands pass; the smoke log reports successful AppleScript compilation; the desktop build contains `office-native-dialog.html`; existing Office.js and Windows architecture tests remain unchanged.

## Word: VisualTeX.dotm

1. Install the native offline add-ins from VisualTeX Settings.
2. Confirm `VisualTeX.dotm` is in every discovered Word Startup/Word directory.
3. Start Word manually. Installation is healthy only after `~/Library/Application Support/VisualTeX/OfficePluginStatus/word.json` reports `loaded: true` and the expected plugin version.
4. Restart Word ten times. The VisualTeX Ribbon must appear every time without opening an add-in store or requesting a network connection.
5. Disconnect all network interfaces. Create and edit an inline formula; verify baseline alignment and that cancelling removes only the pending placeholder.
6. Create and edit a display formula. Toggle numbering in the editor, update equation numbers, save, close, and reopen the document.
7. Move the caret and change the selection while the VisualTeX editor is open. The formula must replace the original pending/bookmarked object, not the new caret position.
8. Switch to another Word document before saving. VisualTeX must reject the write and must not modify the other document.
9. Replace `VisualTeX.dotm` at the same Startup path with a newer build, restart Word, and confirm no re-registration is required.
10. Verify uninstall removes only VisualTeX files and leaves documents with cached formula images intact.

## PowerPoint: VisualTeX.ppam

1. Confirm the installed path is exactly `~/Library/Application Support/VisualTeX/OfficeAddins/VisualTeX.ppam`.
2. Use VisualTeX Settings to reveal the file and follow the tutorial to register it once through **Tools → PowerPoint Add-Ins**. Do not use UI automation.
3. Restart PowerPoint ten times. The Ribbon must appear every time after the one manual registration.
4. Disconnect all network interfaces. Create, edit, and delete formulas.
5. Confirm new formulas start at the current slide center and use `VisualTeX_<formulaId>` names plus `VisualTeXFormulaId`, `VisualTeXSessionId`, and `VisualTeXPending` tags.
6. Rotate a formula, move it between other shapes, and edit it to a much longer formula. Confirm the center, rotation, z-order, and visual font size are retained; width may grow and text must not be compressed into the old box.
7. Switch presentations before saving. VisualTeX must reject the write and leave both presentations unchanged.
8. Cancel a create Session and confirm only the matching pending shape is deleted.
9. Replace the PPAM at the same fixed path, restart PowerPoint, and confirm no new registration is required.
10. Uninstall the add-in and verify existing presentations still display cached formula images.

## Offline and boundary checks

- No Office.js script, HTTPS listener, local certificate, manifest, WebView task pane, Trusted Catalog, mouse simulation, keyboard simulation, language-dependent menu automation, double-click hook, polling loop, or SelectionChange background handler is used by the native Mac plug-ins.
- AppleScriptTask accepts only a canonical UUID v4 and launches only the fixed `visualtex://office/open?session=<uuid>` URL through `/usr/bin/open` with `quoted form of`.
- VisualTeX reuses one running desktop process and opens one editor window per Session.
- Existing Office.js compatibility installation remains available until this checklist passes at equal or higher strength.
