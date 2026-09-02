# VisualTeX macOS offline Office integration

Status: supported macOS architecture on `main`. Runtime behavior must be verified against the current source and installed Office artifacts rather than an obsolete feature branch.

This document defines the native, completely offline macOS Word and PowerPoint integration. The retired Office.js route has been removed from the macOS application after the native DOTM/PPAM workflow became the supported implementation.

## Non-negotiable boundaries

- Word loads `VisualTeX.dotm` as a global template from Word's actual Startup path.
- PowerPoint loads the fixed file `~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeAddins/VisualTeX.ppam` after a one-time manual registration in PowerPoint.
- Ribbon callbacks are VBA. They call `AppleScriptTask`, which receives only a UUID session id.
- The AppleScriptTask scripts accept no shell command, file path, LaTeX, image data, or arbitrary URL. They validate the UUID and open the fixed URL `visualtex://office/open?session=<uuid>` with `/usr/bin/open` and `quoted form of`.
- Formula request and result payloads are local files under the Office application-group container at `~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeSessions/<session-id>/`, which is accessible to both sandboxed Office hosts and the desktop app.
- The Tauri application imports the request into the existing `SessionStore`; formula editing and export reuse the existing Office editor and metadata schema.
- macOS stores Word and PowerPoint picture formulas as SVG-backed `InlineShape`/`Shape` objects. PowerPoint imports the SVG directly. Word imports a minimal DOCX containing the SVG blip because Word for Mac VBA rejects raw `.svg` files through `InlineShapes.AddPicture`; PNG is retained only as the SVG compatibility preview and emergency fallback. It never registers or pretends to provide the Windows `VisualTeX.Formula.1` COM/OLE class.
- No Office.js, XML manifest, trusted-certificate installation, WebView task pane, external network access, menu-language matching, coordinate clicking, mouse simulation, or keyboard simulation is part of this route. A private loopback TLS companion remains for Session/OCR APIs only. Double-click editing is handled by the host's native `Application.WindowBeforeDoubleClick` VBA event.

## Runtime sequence

### Word create

1. VBA creates UUID v4 values for `sessionId` and `formulaId`.
2. VBA inserts a one-pixel transparent pending `InlineShape` at the current Range, writes the exact marker `visualtex:pending:v1:<sessionId>:<formulaId>` to `Title` and `AlternativeText`, and moves the Word caret after that object so pressing Enter cannot relocate the transaction target.
3. VBA atomically writes `request.json` and calls `AppleScriptTask("VisualTeXWord.scpt", "OpenVisualTeXSession", sessionId)`.
4. VisualTeX receives the custom URL, validates the UUID, imports the request into `SessionStore`, and opens a local Tauri editor window.
5. On commit, VisualTeX writes an authenticated-by-location, strictly validated `dispatch.txt`, materializes the validated SVG, its PNG compatibility preview, and `formula-svg.docx`, a minimal OOXML package that embeds both image parts.
6. The macro opens `formula-svg.docx` invisibly, transfers Word's parsed SVG `InlineShape` through `FormattedText`, fully configures the replacement, and deletes the old object only after success. It uses the PNG only if that vector-document import fails.
7. On cancel, the same fixed macro removes only the matching pending object. Closing a non-empty native editor performs the same commit transaction; closing an empty editor performs this cancel transaction before the Tauri window closes.

### Word edit

The selected object must be exactly one `InlineShape`. The VBA envelope validates that its metadata uses the VisualTeX prefix. VisualTeX then inflates and validates the complete schema, `formulaId`, and non-empty `lines` array before creating an edit Session. The original compressed marker is retained as the durable lookup key, so moving the caret while the editor is open cannot redirect the commit. `VTWordEvents` keeps a `WithEvents Word.Application` reference alive and routes a double-click on a valid formula directly to this same edit path.

### PowerPoint create/edit

VBA creates a centered pending shape for create, or captures one selected formula shape for edit. The request records the presentation identity, slide id/index, shape name, geometry, rotation, and z-order. The staged offline bridge materializes the existing Session export locally and invokes one fixed VBA transaction: it creates and fully decorates the replacement picture, verifies geometry and metadata, places it immediately above the original z-order, and deletes the original only as the final mutation. `VTPowerPointEvents` keeps a `WithEvents PowerPoint.Application` reference alive and routes a double-click on a valid VisualTeX shape to this same edit transaction. The legacy global double-click monitor yields whenever a current native plug-in health record is present, preventing duplicate edit windows.

## Persistent files

```text
~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/
  OfficeAddins/
    VisualTeX.ppam
    resources/placeholder.png
  OfficeSessions/<session-id>/
    request.json
    dispatch.txt
    formula.svg
    formula-svg.docx  # Word SVG staging package
    formula.png       # SVG preview and emergency fallback
  OfficePluginStatus/
    word.json
    powerpoint.json
```

Installed AppleScriptTask files:

```text
~/Library/Application Scripts/com.microsoft.Word/VisualTeXWord.scpt
~/Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXPowerPoint.scpt
```

The add-in filenames never contain a version. Updates replace the same path atomically.

## Trust and validation model

- Every externally supplied identifier is parsed as a canonical UUID.
- Session paths are constructed from the validated UUID; callers never provide a filesystem path.
- Request JSON has a size cap and a fixed field set.
- Image and dispatch paths are generated by VisualTeX inside the validated Session directory.
- VBA rejects control characters and path traversal before opening a result file.
- AppleScript uses only fixed commands and fixed URL prefixes.
- A commit is transactional: the previous Office object is retained until the candidate has valid geometry and metadata.
- Cancellation never modifies an existing formula; it removes only a create placeholder.

## Add-in build boundary

VBA source, Ribbon XML, AppleScript source, installer logic, and automated source-level tests are versioned. Macro-enabled Office binaries contain a compiled `vbaProject.bin`; they must be produced on a controlled Office build machine from the reviewed source and then copied to the fixed resource names:

```text
office/macos-offline/resources/VisualTeX.dotm
office/macos-offline/resources/VisualTeX.ppam
```

The build and installer fail closed when either compiled artifact is absent or has an unexpected checksum. A blank or renamed OOXML file is never accepted as a native add-in.

## Removed legacy route

The retired macOS Office.js manifests, trusted-certificate installer, web bridge, and Windows native sources are not part of this application. The supported Office route is the native DOTM/PPAM integration above. The local companion remains only for shared Session/OCR services used by the native workflow; it does not publish an Office.js task pane or install XML manifests.
