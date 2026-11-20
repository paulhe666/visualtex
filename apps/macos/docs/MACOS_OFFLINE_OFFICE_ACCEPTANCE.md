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
9. Run `VisualTeX_ProbeImageFormulaFontSize`. Record whether this Mac Word build persists `InlineShape.Range.Font.Size`; both outcomes are valid, but the VisualTeX Ribbon point-size drop-down must remain available on every build.
10. With an inline SVG formula selected, enter `10.5`, `12`, `14`, and `18` in Word's Home font-size box. When the probe passes, width, height, and inline baseline must update proportionally within one selection event or one watcher cycle. Then use the VisualTeX Ribbon drop-down to choose `五号 (10.5 pt)`, `小四 (12 pt)`, `四号 (14 pt)`, `小二 (18 pt)`, `48 pt`, `72 pt`, and `96 pt`; each choice must apply immediately without typing. All Chinese labels must render correctly rather than as mojibake. Select formulas with different sizes together and confirm the drop-down reports `当前: 混合字号` before applying one preset to all of them.
11. Create an SVG formula while the insertion selection is `12 pt`. Word must import it from the generated `formula-svg.docx` without showing the generic “cannot open formula.svg” dialog; inspect the saved DOCX package and confirm the formula drawing retains an SVG relationship plus its PNG preview. Edit and rerender it, convert it to OMML, change the OMML to `18 pt`, then convert it back to SVG. Every representation must retain the current Word point size. For a numbered display SVG formula, change `12 → 24 → 48 → 12 pt`: the visible number must use exactly the same point size and `Cambria Math` Western font after every change. Its `Font.Position` must be recomputed from the exported SVG mathematical baseline as `round(-referenceBaselinePt × imageHeight / referenceHeightPt)`; only formulas without baseline metadata may use the legacy outer-box-centre fallback. The formula must remain on the centre tab and the number on the right tab. Repeat after updating all equation numbers and after directly resizing the image; the baseline-to-height ratio, `VT_F_`/`VT_N_` Bookmarks, SEQ fields, and body cross-references must remain intact.
12. Move the caret and change the selection while the VisualTeX editor is open. The formula must replace the original pending/bookmarked object, not the new caret position. Repeat the same commit after forcing the Session completion write to fail: Word must recognize the already committed metadata/Title pair and must not insert a duplicate formula.
13. Switch to another Word document before saving. VisualTeX must reject the write and must not modify the other document.
14. Replace `VisualTeX.dotm` at the same Startup path with a newer build, restart Word, and confirm no re-registration is required.
15. Verify uninstall removes only VisualTeX files and leaves documents with cached formula images intact.

## PowerPoint: VisualTeX.ppam

1. Confirm the installed path is exactly `~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeAddins/VisualTeX.ppam`.
2. Use VisualTeX Settings to reveal the file and follow the tutorial to register it once through **Tools → PowerPoint Add-Ins**. Do not use UI automation.
3. Restart PowerPoint ten times. The Ribbon must appear every time after the one manual registration.
4. Disconnect all network interfaces. Create, edit, and delete formulas in a presentation whose filename contains Chinese characters. Closing a non-empty native editor must commit and close the window; closing an empty editor must cancel and remove the pending shape.
5. Double-click a committed formula shape. The PowerPoint application event must suppress the default shape action and open exactly one VisualTeX edit Session without requiring the Ribbon edit button.
6. Confirm new formulas start at the current slide center and use `VisualTeX_<formulaId>` names plus `VisualTeXFormulaId`, `VisualTeXSessionId`, `VisualTeXPending`, `VisualTeXFontSizePt`, `VisualTeXReferenceWidthPt`, and `VisualTeXReferenceHeightPt` tags.
7. Select a text run set to 24 pt, insert a new formula, and confirm the SVG formula inherits 24 pt. With no text selection, confirm a new formula defaults to 18 pt.
8. Select one SVG formula and use the VisualTeX Ribbon point-size drop-down to apply `小二 (18 pt)`, `小一 (24 pt)`, `一号 (26 pt)`, `小初 (36 pt)`, `48 pt`, `72 pt`, and `96 pt`. Each choice must apply immediately without typing; width and height must scale proportionally around the same center while rotation and z-order remain unchanged. All Chinese labels must render correctly rather than as mojibake. Select SVG formulas with different sizes together and confirm the drop-down reports `当前: 混合字号` before applying one preset to all of them.
9. Manually resize a formula proportionally, change the selection, then reopen or rerender it. Confirm the physical SVG height is interpreted as the current point size and the formula does not jump back to stale metadata.
10. Rotate a formula, move it between other shapes, and edit it to a much longer or taller formula. Confirm the center, rotation, z-order, and point size are retained; natural width and height may grow and the SVG must not be compressed into the old box.
11. Switch presentations before saving. VisualTeX must reject the write and leave both presentations unchanged. Repeat a commit after forcing the Session completion write to fail: PowerPoint must recognize the final SessionId/metadata/geometry object and must not create a second shape.
12. Cancel a create Session and confirm only the matching pending shape is deleted.
13. Replace the PPAM at the same fixed path, restart PowerPoint, and confirm no new registration is required.
14. Uninstall the add-in and verify existing presentations still display cached formula images.

## Offline and boundary checks

- No Office.js script, XML manifest, trusted-certificate installation, WebView task pane, Trusted Catalog, mouse simulation, keyboard simulation, language-dependent menu automation, or document-wide polling loop is used by the native Mac plug-ins. The private loopback TLS companion is limited to Session/OCR APIs. Double-click editing uses `Application.WindowBeforeDoubleClick`; narrow selection-change handlers only refresh the selected formula's size state and Ribbon control.
- AppleScriptTask accepts only a canonical UUID v4 and launches only the fixed `visualtex://office/open?session=<uuid>` URL through `/usr/bin/open` with `quoted form of`.
- VisualTeX reuses one running desktop process and opens one editor window per Session. Duplicate or concurrent deliveries of the same URL reuse the imported Session.
- Completed and cancelled Sessions remove only the known request, dispatch, and rendered-PNG artifacts; unknown files prevent directory removal and are never recursively deleted.
- The retired Office.js compatibility installation is absent from the macOS application; only the native DOTM/PPAM route is supported.
