# VisualTeX Windows Native Office and OLE Architecture

Status: implementation branch `feat/windows-native-vsto-ole`; native packaging and the current x64 Office 2024 qualification run passed, while the broader release matrix remains pending.

This document defines the Windows-native Office architecture and its current implementation state. The Windows bundle now packages the native VSTO + ATL OLE path, while the existing Office.js trusted-catalog scripts remain available only as an explicit compatibility and cleanup path. Source/build completion does not by itself qualify a public release; the real-Office acceptance gates below still apply.

## 1. Platform boundary

- Windows uses the existing `VisualTeX.WordVsto` and `VisualTeX.PowerPointVsto` COM add-ins for Ribbon commands and Office events.
- Windows professional mode stores formulas as real OLE objects created with `VisualTeX.Formula.1`.
- Windows cross-platform mode continues to store PNG/SVG pictures with VisualTeX metadata.
- macOS never registers the Windows CLSID and never claims that a PNG/SVG Shape is an OLE object.
- The macOS implementation is maintained independently under `apps/macos` and is not included in the Windows application. Windows retains only its own explicit Office.js trusted-catalog compatibility/cleanup path alongside the native VSTO + ATL OLE route.

## 2. Permanent COM identities

These identities are release ABI. They must not be regenerated after publication.

| Identity | Value |
| --- | --- |
| ProgID | `VisualTeX.Formula.1` |
| VersionIndependentProgID | `VisualTeX.Formula` |
| Formula object CLSID | `{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}` |
| `IVisualTeXFormulaObject` IID | `{6C672AF0-7321-4D21-B325-868CB34592C2}` |
| LocalServer AppID | `{3C72FF7F-B04A-4FD0-AA7D-61D110D8B3C1}` |

The server is an out-of-process ATL EXE. Office bitness does not change these identities. Because the VSTO add-ins are in-process and the embedded type-library registration is view-specific, release packaging produces matched x86 and x64 MSI variants. Each variant installs the corresponding VSTO assemblies and LocalServer and writes the same CLSID/IID/ProgID into the correct per-user Registry32 or Registry64 view.

## 3. Structured storage contract

The root OLE storage contains these streams:

- `VisualTeX.Formula.json`
- `VisualTeX.Preview.emf`
- `VisualTeX.Preview.png`

`VisualTeX.Formula.json` uses UTF-8 without BOM and contains:

- `schemaVersion`
- `formulaId`
- `title`
- `latex`
- `lines`
- `codeFormat`
- `displayMode`
- `numbered`
- `renderWidthPx`
- `renderHeightPx`
- `baseline`
- `createdWithVersion`
- `updatedWithVersion`
- `createdAt`
- `updatedAt`

Unknown JSON fields must be preserved by the desktop editor when possible. Readers must reject an invalid UUID, unsupported schema version, empty `lines`, invalid dimensions, or a numbered non-Word-display formula.

The EMF stream is the primary Windows Office presentation cache. The PNG stream is the fallback and remains useful to non-Windows readers and recovery tooling. The object must draw its cached preview without VisualTeX running.

## 4. Native initialization interface

After `InlineShapes.AddOLEObject` or `Shapes.AddOLEObject` creates the object, the VSTO add-in queries the dual Automation interface `IVisualTeXFormulaObject` from `OLEFormat.Object` and calls:

```text
InitializeFromFiles(metadataJson, emfPath, pngPath)
```

The LocalServer validates that each preview path is a regular file under the current user's VisualTeX Office temporary directory, copies the bytes into object-owned memory, marks the object dirty, sends view/data advise notifications, and requests an immediate container save through `IOleClientSite::SaveObject`. The in-memory update is rolled back if the container save fails. Office persists the object through `IPersistStorage`.

Edits use:

```text
UpdateFromFiles(metadataJson, emfPath, pngPath)
```

An update is transactional inside the server: all inputs are validated and loaded before the current metadata or preview is replaced. Cancellation never calls the update method.

The custom interface is intentionally small, uses fixed DispIds 1–3 through an embedded dual Automation type library, and is versioned through the JSON schema. It is not a substitute for `IOleObject`, `IDataObject`, `IPersistStorage`, or `IViewObject2`. Native OLE formula metadata is read from the embedded JSON stream; Word Title/AlternativeText and PowerPoint Tags remain picture-mode compatibility metadata, not a second native-OLE source of truth.

## 5. OLE behavior

The first release supports external editing only.

- `OLEIVERB_PRIMARY` and `OLEIVERB_OPEN` activate VisualTeX Desktop.
- `OLEIVERB_SHOW` starts or reconnects a persisted object without launching the editor, allowing Word to retrieve the Automation interface after reopening a document.
- No in-place activation, Office menu merging, or embedded WebView is implemented.
- `IViewObject2::Draw` draws `DVASPECT_CONTENT` from the EMF cache, then PNG, then a non-destructive placeholder.
- `IDataObject` exposes `CF_ENHMETAFILE` and the registered `PNG` clipboard format when available.
- `IPersistStorage` implements `InitNew`, `Load`, `Save`, `SaveCompleted`, `HandsOffStorage`, and `IsDirty`. A new blank object first persists a safe placeholder JSON/EMF/PNG set because PowerPoint may save and reload the object before VSTO supplies the real formula.
- `IOleObject` and `IDataObject` advise holders notify Office after a committed update.
- A server crash must leave the last successfully saved storage intact.

## 6. VSTO responsibilities

The existing Word and PowerPoint add-ins remain the only native Ribbon implementations.

Word commands:

- New inline formula
- New display formula
- Edit selected formula
- Update equation numbers
- Delete selected formula
- Convert selected legacy picture to OLE through the transactional edit session
- Export selected OLE formula as a cross-platform picture
- Open VisualTeX

PowerPoint commands:

- New formula
- Edit selected formula
- Delete selected formula
- Convert selected legacy picture to OLE through the transactional edit session
- Export selected OLE formula as a cross-platform picture
- Open VisualTeX

The add-ins use true Office OLE insertion APIs in professional mode. The existing picture services remain available only for cross-platform mode and transactional migration rollback.

Word picture double-click handling is retained only for legacy picture formulas. Real OLE objects use Office's native double-click activation and must not be intercepted by a global mouse hook.

## 7. Session protocol

Protocol version 2 extends the existing local Session model; it does not create a second editor data model.

A native add-in creates a Session with:

- immutable `sessionId`, `formulaId`, host, mode, source document identity, and source object identity;
- existing formula metadata for edits;
- `objectMode` equal to `nativeOle` or `crossPlatformPicture`;
- `autoCommitOnClose` enabled unless the caller explicitly disables it.

The editor writes SVG, PNG, dimensions, and baseline to the existing atomic Session Store. Windows VSTO materializes the self-contained MathJax SVG and converts its supported path-only subset to a true vector EMF. The converter rejects raster images, external references, unsupported visible nodes, unsupported alpha, and EMF/EMF+ raster records. Native OLE mode therefore fails closed rather than putting a raster image inside an EMF container and calling it vector output.

A committed edit updates the OLE object's storage first, then marks the Session completed. A failed object update marks the Session failed and keeps the previous object unchanged. An explicit cancel never mutates Office.

## 8. Installation

The native installer is independent of the Office.js trusted catalog.

- MSI/WiX installs both VSTO add-ins and the ATL LocalServer per user.
- Add-ins use `LoadBehavior=3` under HKCU.
- The LocalServer uses HKCU `Software\\Classes` registration, including `LocalServer32`, an unquoted `ServerExecutable`, `InprocHandler32=Ole32.dll`, `Insertable`, `AuxUserType`, `DataFormats`, `DefaultIcon`, `Verb`, `ProgID`, `VersionIndependentProgID`, TypeLib, interface proxy metadata, and AppID.
- Install, repair, update, and uninstall do not start Word or PowerPoint.
- No UI Automation, cursor positioning, keyboard simulation, Office store dialog, or Ribbon cache polling is part of the native path.
- Existing certificate, UTF-8 manifest, long-path, hidden-console, first-run tutorial, and Office cleanup code remain available to the compatibility path and are not removed by this phase.

## 9. Transactional picture migration

For each legacy formula:

1. Decode and validate AlternativeText/Tags metadata.
2. Create and fully initialize a new OLE object at the same location.
3. Restore dimensions, center, rotation, z-order, inline baseline, display paragraph, and numbering.
4. Verify the OLE object exposes the expected formula UUID and can be persisted.
5. Delete the old picture only after verification.
6. On any failure, delete the candidate OLE object and retain the old picture.

Bulk conversion records per-object results and never treats partial completion as an all-or-nothing document failure.

## 10. Rollout gates

The native MSI/NSIS path is implemented and packaged, but it is not release-qualified until all of these are evidenced in saved logs on clean Office installations:

1. Word and PowerPoint Ribbon survives ten restarts.
2. Installation never starts Office and performs no UI automation.
3. Fully offline create/edit works in both hosts.
4. Both hosts create objects whose class is `VisualTeX.Formula.1`.
5. Office-native double-click invokes `IOleObject::DoVerb`.
6. Save, close, reopen, and edit preserves all JSON and previews.
7. 32-bit and 64-bit Office pass.
8. Chinese and English Office pass with identical registration logic.
9. Uninstall leaves document cache previews visible.
10. Existing picture formulas remain intact and transactional migration passes.

Passing source-level tests or compiling the LocalServer is not sufficient to switch the default.

## 11. Implementation stages and evidence

Each stage records changed files, exact commands, results, and environmental blockers in `docs/implementation-logs/`.

1. Architecture and protocol.
2. Existing VSTO Ribbon/service refactor.
3. Native installer path with no Office startup or UI automation.
4. Minimal ATL LocalServer and independent OLE container test.
5. Word `AddOLEObject` integration.
6. PowerPoint `AddOLEObject` integration.
7. Session-driven preview update.
8. Legacy picture migration.
9. MSI/NSIS integration and full acceptance.
10. Public-release qualification only after all gates pass.

Bulk conversion of every legacy formula in a document is intentionally not exposed yet. Selected-object conversion is complete and transactional; safe bulk conversion requires a headless or queued vector-render workflow so each legacy formula can obtain a validated SVG/EMF without opening overlapping editor sessions.
