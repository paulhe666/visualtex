# macOS native offline Office acceptance

Record the VisualTeX version, Microsoft Office version, macOS version, machine architecture, test account, and artifact SHA-256 values before running this checklist. Save console output under ignored `build-logs/macos-offline/`.

## Automated source and runtime checks

```bash
npm run test:macos-offline-office
cargo test --manifest-path src-tauri/Cargo.toml --lib
npm run build:desktop
```

Expected: all commands pass; the smoke log reports successful AppleScript compilation; and the desktop build contains `office-native-dialog.html` without an Office.js bundle. After compiling each real Office artifact, run the VBA function `VTProtocolSelfTest` once in that host and require `True`; this exercises 1,000 UUIDs plus a UTF-8 round trip containing Chinese, Greek, and a supplementary Unicode character. Installation must be transactional: inject a controlled failure after at least one destination is replaced and confirm every previous VisualTeX file is restored. A same-version health record is valid only when the exact host, current plugin version, and a bounded timestamp are present.

## Word: VisualTeX.dotm

1. Install the native offline add-ins from VisualTeX Settings.
2. Confirm `VisualTeX.dotm` is in every discovered Word Startup/Word directory.
3. Start Word manually. Installation is healthy only after `~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficePluginStatus/word.json` reports `loaded: true` and the expected plugin version.
4. Restart Word ten times. The VisualTeX Ribbon must appear every time without opening an add-in store or requesting a network connection.
5. Disconnect all network interfaces. Create and edit an inline formula in a document whose filename contains Chinese characters; verify baseline alignment and UTF-8 Session import. Closing a non-empty native editor must commit and close the window; closing an empty editor must cancel, remove the transparent pending target, and leave no black square.
6. After starting a create Session, return to Word and press Enter. The one-pixel transparent pending target must remain at the original insertion location because the caret was moved after it. Complete the Session and verify the formula replaces that original target.
7. Double-click a committed inline formula. The Word application event must suppress the default picture action and open exactly one VisualTeX edit Session without requiring the Ribbon edit button.
8. Create and edit a display formula. Toggle numbering in the editor, update equation numbers, save, close, and reopen the document. A numbered display formula must stay centered while its number remains right-aligned on the same line.
9. Move the caret and change the selection while the VisualTeX editor is open. The formula must replace the original pending/bookmarked object, not the new caret position. Repeat the same commit after forcing the Session completion write to fail: Word must recognize the already committed metadata/Title pair and must not insert a duplicate formula.
10. Switch to another Word document before saving. VisualTeX must reject the write and must not modify the other document.
11. Replace `VisualTeX.dotm` at the same Startup path with a newer build, restart Word, and confirm no re-registration is required.
12. Verify uninstall removes only VisualTeX files and leaves documents with cached formula images intact.

## PowerPoint: VisualTeX.ppam

1. Confirm the installed path is exactly `~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeAddins/VisualTeX.ppam`.
2. Use VisualTeX Settings to reveal the file and follow the tutorial to register it once through **Tools → PowerPoint Add-Ins**. Do not use UI automation.
3. Restart PowerPoint ten times. The Ribbon must appear every time after the one manual registration.
4. Disconnect all network interfaces. Create, edit, and delete formulas in a presentation whose filename contains Chinese characters. Closing a non-empty native editor must commit and close the window; closing an empty editor must cancel and remove the pending shape.
5. Double-click a committed formula shape. The PowerPoint application event must suppress the default shape action and open exactly one VisualTeX edit Session without requiring the Ribbon edit button.
6. Confirm new formulas start at the current slide center and use `VisualTeX_<formulaId>` names plus `VisualTeXFormulaId`, `VisualTeXSessionId`, and `VisualTeXPending` tags.
7. Rotate a formula, move it between other shapes, and edit it to a much longer formula. Confirm the center, rotation, z-order, and visual font size are retained; width may grow and text must not be compressed into the old box.
8. Switch presentations before saving. VisualTeX must reject the write and leave both presentations unchanged. Repeat a commit after forcing the Session completion write to fail: PowerPoint must recognize the final SessionId/metadata/geometry object and must not create a second shape.
9. Cancel a create Session and confirm only the matching pending shape is deleted.
10. Replace the PPAM at the same fixed path, restart PowerPoint, and confirm no new registration is required.
11. Uninstall the add-in and verify existing presentations still display cached formula images.

## Offline and boundary checks

- No Office.js script, XML manifest, trusted-certificate installation, WebView task pane, Trusted Catalog, mouse simulation, keyboard simulation, language-dependent menu automation, polling loop, or SelectionChange background handler is used by the native Mac plug-ins. The private loopback TLS companion is limited to Session/OCR APIs. Double-click editing uses only the host-provided `Application.WindowBeforeDoubleClick` event in a persistent VBA class module.
- AppleScriptTask accepts only a canonical UUID v4 and launches only the fixed `visualtex://office/open?session=<uuid>` URL through `/usr/bin/open` with `quoted form of`.
- VisualTeX reuses one running desktop process and opens one editor window per Session. Duplicate or concurrent deliveries of the same URL reuse the imported Session.
- Completed and cancelled Sessions remove only the known request, dispatch, and rendered-PNG artifacts; unknown files prevent directory removal and are never recursively deleted.
- The retired Office.js compatibility installation is absent from the macOS application; only the native DOTM/PPAM route is supported.
