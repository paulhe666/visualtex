# VisualTeX macOS offline Office protocol v1

## URL

Only this form is accepted:

```text
visualtex://office/open?session=<canonical-uuid>
```

There must be exactly one `session` query parameter. Fragments, user information, ports, additional query parameters, non-canonical UUID spellings, and paths other than `/open` are rejected.

## Request file

Location:

```text
~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeSessions/<sessionId>/request.json
```

Maximum encoded size: 256 KiB.

```json
{
  "protocolVersion": 1,
  "sessionId": "uuid",
  "host": "word | powerpoint",
  "mode": "create | edit",
  "formulaId": "uuid or null",
  "displayMode": "inline | block",
  "numbered": false,
  "nativeEquation": false,
  "sourceDocumentId": "bounded text or null",
  "sourceObjectId": "bounded text or null",
  "encodedMetadata": "visualtex:v1:deflate:... or null",
  "pendingMarker": "visualtex:pending:v1:<sessionId>:<formulaId> or null",
  "fontSizePt": 12.0,
  "referenceWidthPt": 96.0,
  "referenceHeightPt": 24.0,
  "powerPoint": {
    "presentationIdentity": "bounded text",
    "slideIndex": 1,
    "slideId": 256,
    "shapeName": "VisualTeX_<formulaId>",
    "left": 100.0,
    "top": 100.0,
    "width": 96.0,
    "height": 32.0,
    "rotation": 0.0,
    "zOrder": 3,
    "fontSizePt": 24.0,
    "referenceWidthPt": 120.0,
    "referenceHeightPt": 30.0
  }
}
```

Rules:

- `sessionId` must equal the directory name and URL value.
- Managed SVG formulas retain their existing `formulaId`/`encodedMetadata` identity rules. Pure native PowerPoint OMML requests use `nativeEquation: true` with `formulaId`, `encodedMetadata`, and `pendingMarker` all null, for both create and edit. The desktop session may allocate an internal temporary identifier, but none is written to the presentation.
- `encodedMetadata` is inflated by VisualTeX and must validate against `visualtex-formula` schema version 1, including at least one line.
- `numbered` is valid only for Word block formulas.
- `nativeEquation` selects native Office Math for Word or PowerPoint. PowerPoint native editing resolves the current `TextFrame2.TextRange.MathZones(index, 1)` at the caret/selection, not a named or tagged formula shape. Existing SVG requests continue to use `false`.
- Word `fontSizePt` is a finite value from 1 to 512. New formulas inherit the Word selection size; edits and image/OMML conversions carry the source formula size.
- Word top-level `referenceWidthPt` and `referenceHeightPt` are optional positive image dimensions at the stable 14 pt reference size.
- `powerPoint` is required only for PowerPoint and is rejected for Word. SVG requests retain optional `fontSizePt`, `referenceWidthPt`, and `referenceHeightPt` reference geometry. Native requests include `fontSizePt` and, for an existing text position, `nativeTextTarget: {rangeStart, rangeLength, leadingParagraph, trailingParagraph}`. A native edit requires a nonempty current mathematical range and both paragraph flags false. A standalone native create has zero shape references and no placeholder object.
- A new PowerPoint formula inherits a selected text/formula size when available and otherwise uses 18 pt. Editing or rerendering preserves the point size while allowing the new SVG width and height to follow the formula's natural aspect ratio.
- All strings reject NUL and control characters, and every numeric geometry value must be finite and bounded.

## Document import request

Word can open a dedicated LaTeX/Markdown document importer by adding these fields to the normal request envelope:

```json
{
  "protocolVersion": 1,
  "sessionId": "uuid",
  "host": "word",
  "mode": "create",
  "operation": "documentImport",
  "formulaId": null,
  "displayMode": "inline",
  "numbered": false,
  "nativeEquation": false,
  "sourceDocumentId": "bounded Word document identity",
  "sourceObjectId": null,
  "encodedMetadata": null,
  "pendingMarker": null,
  "fontSizePt": null,
  "referenceWidthPt": null,
  "referenceHeightPt": null,
  "powerPoint": null,
  "documentImport": {
    "bookmarkName": "VT_D_<bounded insertion bookmark>",
    "defaultFontSizePt": 12.0
  }
}
```

The importer parses a structured stream of Word paragraphs and formula blocks. Ordinary prose remains ordinary Word text and is normalized so it does not inherit italic, bold, underline, or raised-baseline formatting from the insertion caret. LaTeX `section`/`subsection`/`subsubsection`, Markdown headings, `itemize`, `enumerate`, Markdown lists, quote environments, code blocks, and center/right alignment are mapped to native Word paragraph styles, list formatting, and alignment. Blank source lines define paragraph boundaries but are not inserted as empty Word paragraphs around display formulas.

Each formula receives its own UUID, LaTeX source, display mode, optional equation-number flag, point size, OMML payload, native staging document, and VisualTeX metadata. A batch chooses one representation for all formula blocks: native Word OMML or VisualTeX SVG image formulas. Formula identity and sizing remain independent after insertion. SVG blocks may include a PNG compatibility preview; if WebView rasterization is unavailable, the backend supplies a validated fallback so the SVG transaction does not fail.

The app materializes a line-oriented `document-import.txt` manifest in the Session directory and invokes the normal Word callback with `action=documentCommit`. Text is UTF-8 Base64URL encoded. Structured paragraph items carry `paragraphId`, `paragraphStyle`, `paragraphAlignment`, `listKind`, `listLevel`, `paragraphStart`, and `paragraphEnd`. Formula entries reference only validated files under the Word VisualTeX runtime. Word performs the insertion as one rollback-capable transaction and reuses the normal inline baseline, unnumbered display centering, numbered display three-column true-centering, metadata, equation-number, and cross-reference routines. `action=documentCancel` removes only the zero-width insertion bookmark.

Limits:

- 1–2048 total blocks.
- At most 512 formula blocks.
- At most 4 MiB of decoded text.
- Each formula has an independent font size from 1 to 512 pt.
- Numbering is allowed only for display formulas.

## Dispatch file

Location:

```text
~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeSessions/<sessionId>/dispatch.txt
```

This file is written atomically by VisualTeX and read by a fixed VBA callback. It is line-oriented so Office VBA does not need a JSON parser. Keys are unique and values may not contain CR or LF.

```text
protocolVersion=1
sessionId=<uuid>
action=commit|cancel
host=word|powerpoint
mode=create|edit
formulaId=<uuid>
displayMode=inline|block
numbered=0|1
imagePath=<absolute validated SVG path generated by VisualTeX>
vectorDocumentPath=<Word-only absolute formula-svg.docx staging path>
fallbackImagePath=<absolute PNG SVG-preview and compatibility path>
metadata=visualtex:v1:deflate:<base64url>
pendingMarker=visualtex:pending:v1:<sessionId>:<formulaId>
sourceMarker=<original exact metadata marker or empty>
widthPoints=<positive target width in Word points>
heightPoints=<positive target height in Word points>
baseline=<finite target baseline in Word points>
fontSizePt=<Word SVG/OMML or PowerPoint SVG formula point size from 1 to 512>
referenceWidthPt=<positive SVG width at 14 pt>
referenceHeightPt=<positive SVG height at 14 pt>
referenceBaselinePt=<Word image baseline at 14 pt, from -256 to 0>
shapeName=VisualTeX_<formulaId>
zOrder=<positive integer>
```

The app writes a separate host pointer file immediately before invoking the fixed macro:

```text
~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeSessions/word-active-session.txt
~/Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeSessions/powerpoint-active-session.txt
```

Each pointer contains only one canonical UUID and is removed after the callback. A per-host process mutex prevents two dispatches from using the same pointer concurrently.

## Health files

Word explicit host initialization / Ribbon loading and PowerPoint `Auto_Open` atomically write:

```json
{
  "loaded": true,
  "pluginVersion": "1.2.8",
  "host": "word",
  "timestamp": "2026-07-15T00:00:00Z"
}
```

The installer treats a copied file as only `filesInstalled`. `loaded=true` with the expected version and a recent host-written timestamp is the runtime health signal.

## Result states

- `completed`: Office callback succeeded, metadata cache updated, and the Session is immutable.
- `cancelled`: create placeholder removed; existing edit target untouched.
- `failed`: old object retained; Session records a diagnostic; retry remains explicit.

## Editing Word equation numbering

The number checkbox is editable for existing block equations, for both SVG images
and native OMML. Draft saves persist the last checkbox state. Apply, or closing
with automatic apply enabled, waits for queued drafts and the current exports;
closing with automatic apply disabled cancels the edit. Exported numbered OMML
uses the current checkbox, including its numbered staging DOCX shell.

A Word commit reconciles the affected representation's following SEQ fields and
chapter/section restart before acknowledging completion. Image and native chains
remain independent. `numberingOnly=1` is a backend-derived image optimization:
it requires unchanged source, code format, font family and font size. VBA also
checks the prior LaTeX/OMML payload before retaining the existing image. Numbering
changes retain image edit metadata; native OMML is read from the live Word tree.
A failed transaction restores the original number, paragraph and content.

Run `npm run test:office-numbering-edit` for delayed-draft/rapid-toggle/apply/close
regressions. With the matching compiled Word add-in installed, run
`npm run test:word-numbering-edit` for real Word SEQ, chapter restart, rollback,
save/reopen and dense-document timing acceptance.

## Native Word equation ownership

Native equations use ordinary `m:oMath` and Word SEQ/REF fields. Completed native
transactions remove formula-specific VT bookmarks, document-variable payloads,
signatures and caret characters. Existing cross-references are preserved using
ordinary Word `_Ref` bookmarks. Editing uses the actual selected OMath tree and a
short-lived edit-session anchor removed on apply/cancel; original Word equations
and equations inserted by VisualTeX share the same selection and double-click path.

Word joins directly adjacent expressions into one native math zone. Inline
insertion therefore combines the exact old subtree and new subtree in a staged
Word package, backs up the old math, and replaces the complete zone atomically.
The source Document is captured before hidden staging is opened and passed
explicitly through the transaction. It must never be inferred from the active
hidden document. Existing runs retain their own font and colour.

Mac Word may clip `Range.OMaths(1).Range` even when its nesting level is zero.
Equation-level operations resolve `ParentOMath` to the complete root before
reading, replacing or positioning it; argument ranges remain scoped. Proven
legacy caret characters are retired with their old ownership. When such a
character separates two existing math zones, an ordinary space preserves that
separation.

Native numbering retains the source paragraph style instead of imposing Caption.
The owned number argument is resolved through OMath EqArray/Delim structure, so
bold mathematical Unicode characters do not break number-slot lookup. Old managed
numbered equations with the plugin-imposed Caption style are migrated during edit.
Number placeholders are verified by concatenating their XML text runs. Format
migration snapshots the complete native tree and restores it if staging or field
replacement fails, so a split placeholder never leaks into the document.

Native Word edit snapshots contain the live `w:document` part, including its
namespace declarations and complete OMath/SEQ tree. Flat OPC styles, themes and
font tables are excluded from this edit-source transport. Small snapshots use the
optional `nativeEditXml` field in the same atomic sandbox-inbox request as the
edit target (at most 64 KiB including JSON framing). This field is accepted only
for an explicit native Word edit session. Larger snapshots and older templates
use `native-edit-original.omml`; that file write creates the session directory.
Opening an equation does not issue a separate directory bridge call or create
a rollback document before Apply. Shared UTF-8 decoding
and Base64 encoding fill preallocated buffers instead of growing strings.

Numbered image paragraphs use independently positioned formula objects. Their
tabs, line/paragraph marks and visible number share normal font positioning.
Historical `VisualTeX Numbered Equation` style lowering is repaired in place;
neither the paragraph style assignment nor its base/next style is rewritten.
This preserves Mac Word's direct center/right tab stops and the user's Chinese
text alignment. Typography synchronization repairs the same scaffold as insertion
and numbering changes.
