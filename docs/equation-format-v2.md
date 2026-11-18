# Word equation format v2 (Windows and macOS)

Status: adopted 2026-10-16. Both platforms write this format; each platform's
「更新公式格式 / Update Equation Format」 command upgrades that platform's own older
forms to it.

## Goal

For the same formula source and settings, VisualTeX on Windows and on macOS
writes **the same Word equation**: the same OMML body, the same number, the same
paragraph. A paper can move between the two machines at any time, and either
machine can renumber, restyle or edit what the other one wrote.

The equation is pure Word: no VisualTeX bookmark, document variable, custom XML
or hidden field stays in the document.

## 1. One writer

| Step | Owner | Code |
|---|---|---|
| Editor lines → render LaTeX | each editor | Windows `OfficeDialogApp.tsx` `serializeOfficeRenderLatex`; macOS `src/office/wordMathMl/wordEquationMathMl.ts` `wordEquationRenderLatex` (the same rule on the Mac editor's document model) |
| LaTeX → MathML | shared TypeScript | `apps/shared/visualtex-app/src/export/runtime.ts` `latexToMathMl`, used by both editors |
| MathML → OMML | C# reference, Rust port | `WordOmmlConverter.TransformMathMlToOmml` ⇄ `src-tauri/src/office/omml_write` |
| Number host | C# reference, Rust port | `WordOmmlEquationNumbers.BuildWordNativeNumberedOmml` ⇄ `omml_write::numbered` |
| Math font transfer | C# reference, Rust port | `ApplyExplicitTransferMathFont(ExtractSingleOMath(x), document math font)` ⇄ `omml_write::insertion_omml` |

The Rust port (`apps/macos/src-tauri/src/office/omml_write/`) is held to the C#
output by the forward reference corpus (`ForwardDump` beside `CorpusDump` in
the Windows isolated tests): every MathML pass, the projection, every OMML
pass, the numbered host in five styles and the math font transfer are
byte-identical for all 3176 corpus rows. `omml_write/fixtures/windows_forward.json`
keeps a 236-row subset that every `cargo test` checks; the whole corpus runs
with `VT_OMML_FORWARD_CORPUS=<forward.json> cargo test --lib -- --ignored full_windows_forward`.

## 2. The equation in the document

Insertion (both platforms):

1. Import the bare `m:oMath` (never a VisualTeX wrapper) into the target paragraph.
2. Display formulas: `OMath.Type = wdOMathDisplay` (Word stores `m:oMathPara`).
3. Typography: `Range.Font.Position = 0`, `Font.Size = size`, `Font.SizeBi = size`
   on the equation range (`WordFormulaHostLayout.ApplyOmmlLocalTypography`).
   `size` follows the text: a new equation takes the size Word reports where it
   goes (`WordFormulaService.ReadCurrentTypingFontSize`; macOS add-in
   `VTV2CurrentTypingFontSize`: the insertion point, else the character before
   or after it, else its paragraph), an edited equation keeps its own size and
   LaTeX redraw keeps each formula's text size. A size picked in the editor or
   the import list applies to that formula only; no size is remembered between
   formulas.
4. Display paragraph: centred, space before/after 0, left/right/first-line indent 0
   (`WordFormulaHostLayout.ConfigureDisplayParagraph`). An inline equation that
   Word shows as display (alone in its paragraph) gets the same layout. An
   equation imported into an empty paragraph keeps that paragraph mark's own
   character formatting (Windows copies it into the imported paragraph,
   `WordOmmlInsertionXml.ParagraphMarkRunProperties`; macOS keeps the original
   mark when it merges the imported paragraph).
5. Math font: the document's math font (`Document.OMathFontName`) on every
   mathematical run and control property (`ApplyExplicitTransferMathFont`);
   `m:nor` text keeps the text font. No per-equation font is applied.
6. After a new display equation the caret moves to the start of the next
   paragraph, which keeps its style and format. Only when the equation ends the
   document is a typing paragraph created (`InsertParagraphAfter`), with the
   style and paragraph format of the body paragraph before the equation, else
   its style's own format (`WordFormulaService.ApplyBodyFormatToTypingParagraph`).

Windows imports with `Range.InsertXML` of the target paragraph. Word for Mac
rejects OMML through `InsertXML` (error 6145), so macOS imports a .docx built by
the app (`omml_write::word_equation_docx`) with `Range.InsertFile` (or opens it
hidden and transfers `FormattedText`) and merges the imported paragraph into
the prepared target. Both give Word the same OMML; the identity check (§6)
compares the saved `word/document.xml`. Font sizes are Word half points
(`FormulaFontSize.NormalizeWordOmmlSize`). Word normalizes imported OMML the
same way on both platforms (InsertXML, InsertFile and FormattedText transfer
give identical runs), so no platform-specific pre-normalization is needed.

Paragraph-level `InsertXML` (Windows) can drop the replaced paragraph's own
mark: with a Normal style that has no paragraph properties (English templates,
Pandoc/python-docx documents) Word joins the paragraph to the next one. Every
paragraph-local import therefore keeps the original mark in a temporary token
paragraph (`WordOmmlConverter` paragraph-mark guard) and restores it afterwards.

On macOS the add-in decides a number's context and asks the app for the
finished numbered equation through the synchronous OMML batch bridge
(`ConvertOmmlBatch`, request `sourceKind=nativeNumber`; one package can hold
many equations, which restyling uses as staging).

## 3. Equation number

```
m:oMath
└─ m:eqArr  (eqArrPr: maxDist=1, ctrlPr: <w:i/>)
   └─ m:e
      ├─ … equation body …
      ├─ m:r "#"
      └─ m:d  (dPr/ctrlPr: <w:i/>)
         └─ m:e
            ├─ prefix            (only for chapter/section styles)
            │    automatic heading: {STYLEREF L \w \* MERGEFORMAT}
            │    otherwise:         m:r[m:nor] "2" / "2.3" / "0" / "0.0"
            ├─ separator         m:r[m:nor] "." or "-"   (ASCII hyphen; m:nor keeps it text, not minus)
            └─ {SEQ <label>[ \r N][ \s L] \* ARABIC \* MERGEFORMAT}
```

* **Styles:** continuous; chapter-dot `(2.1)`; chapter-dash `(2-1)`;
  section-dot `(2.3.1)`; section-dash `(2.3-1)`. L = 1 for chapter, 2 for section.
* **Prefix:** a heading with Word list numbering is read live with
  `STYLEREF L \w`; a heading without list numbering, or no heading yet
  (`0`, `0.0`), is plain text.
* **Label:** the SEQ name the document's numbered equations already use (the
  series of the nearest numbered equation); for a document's first numbered
  equation, Word's own localized Equation caption label
  (`CaptionLabels(wdCaptionEquation).Name`: `Equation`, `公式`, …). One counter per
  document; Word's Insert ▸ Cross-reference lists the numbers under that label.
  Consequence: the first number of a brand-new document follows the language
  of the Word that writes it (a Chinese Word writes `公式`, an English Word
  `Equation`); every later number on either platform keeps the document's label.
* **`\s L`:** restarts the counter at every heading of level L.
* **`\r N`:** written only where the counter must be pinned:
  restyle/upgrade puts `\r 1` on the first equation of every chapter/section
  scope; inserting a number before the first equation of a series moves that
  series' `\r` to the new equation.
* **Field result runs:** boundary runs italic (`<w:i/>`), result run `<w:noProof/>`,
  exactly as `BuildWordNativeNumberedOmml` writes them; Word updates STYLEREF
  first, then SEQ.
* **Style is not stored.** The numbering style of a document is read from its
  numbers (first numbered equation), else the user's default.
* **References:** Word's own `REF _Ref… \h` to a bookmark on the number.

## 4. Editing

Editing a numbered equation replaces only the body before `#`; the number host
is kept verbatim (`ReplaceWordNativeEquationNumberHostBody`). The whole equation,
number included, then takes the formula's size (§2 step 3): the number is part of
the equation, and Word resizes it with the formula when a whole equation is
resized. Adding a number to an existing equation is an insertion into the series
(§3 label/restart rules).

## 5. 更新公式格式 / Update Equation Format

One ribbon button on each platform (group 转换与编号). Flow:

1. Scan (read-only) and show what will change, by kind. Nothing found → "already up to date".
2. Convert everything in one Undo step (Windows: two when retired numbered
   layouts are migrated first, one record each); renumber once; refresh
   references once.
3. Re-scan to verify: no older form left and the same number of equations and
   formula pictures (`OMaths.Count | InlineShapes.Count`). Any failure → the
   update's own Undo records are undone.

Idempotent. Each formula keeps its kind: OMML stays OMML (in v2), VisualTeX
image/OLE formulas stay VisualTeX formulas in their current layout. Legacy forms
per platform are listed in the plan; each platform upgrades only its own.

### 5.1 Fonts

Every VisualTeX formula records the fonts it was drawn in: the math font
(`formulaLetterFont`, `katex` when missing), the Chinese font
(`formulaChineseFont`) and the letter / Greek / digit faces (`formulaFaces`;
missing = none, so formulas from before the record count as drawn without
faces). The scan compares them with the fonts the main app draws in now (the
Chinese font only when the LaTeX has CJK text; names compared trimmed and
case-insensitively). Formulas that differ are listed in the same confirmation
("N VisualTeX formulas drawn in other fonts, to be drawn again in …") and are
drawn again from their stored LaTeX in the current fonts:

* Word / WPS Writer: in place; the formula ID, number, references, font size
  and inline/display form are kept (the number and the references live
  outside the picture). Windows writes them in an Undo record of their own
  after the format records; macOS in the same Undo record as the rest.
* PowerPoint / WPS Presentation (button in the conversion group): each
  formula keeps its form (picture or object), centre, displayed size,
  rotation and layer. PowerPoint for Windows undoes it in one step, plus one
  for the delayed size fix of OLE objects (as any OLE edit); PowerPoint for
  Mac in one step; WPS Presentation keeps every shape change apart. The displayed size comes from the
  shape against its recorded render size; a box whose width and height scale
  differ (the OLE container glitch) falls back to the formula's own font size.

A formula changed or removed while the update runs is left as it is and
reported. When VisualTeX cannot be reached, fonts are not checked and the
confirmation says so.

## 6. Identity check

The same LaTeX cases (inline, display, numbered in all five styles, matrices,
cases, aligned) are inserted with VisualTeX on Windows and on macOS, saved, and
the equation paragraphs of `word/document.xml` are compared after removing
Word's per-save noise (`w:rsid*`, `w14:paraId/textId`, `w:lastRenderedPageBreak`,
and `w:proofErr`, which appears only where Word's proofing ran).
The diff must be empty; anything Word itself writes differently per platform is
reported, not hidden:

* The label of a brand-new document's first number follows the Word UI
  language (§3); Windows Word also gives a CJK label its own field-code run.
* Test documents must state their paragraph defaults (`w:pPrDefault`): Word for
  Mac reads a missing `w:pPrDefault` as 8 pt after / 1.16 lines, Windows Word as
  0 / single.
* Word for Mac splits the first run of the paragraph that follows a paragraph
  whose style changes (any style assignment, explicit or through a merge).
  Inserting a display equation into an empty paragraph with a non-Normal style
  therefore leaves the next paragraph's first character in a run of its own on
  macOS (same text and formatting). VisualTeX for Mac assigns no paragraph style
  where Windows changes none.

## 7. Converters by path

Checked 2026-11-14. LaTeX reaches the OMML writers as MathML from the shared
`latexToMathMl` (§1) on both platforms.

| Path | Windows | macOS |
|---|---|---|
| Insert or edit a Word equation | C# `WordOmmlConverter.TransformMathMlToOmml`, `BuildWordNativeNumberedOmml` | Rust `omml_write` (`macos_offline/word.rs`: `word_equation_artifacts`, `insertion_omml`, `keep_equation_number`) |
| PowerPoint native equation | C# `PowerPointNativeOmmlKernel` (the shared writer's OMML) | Rust `omml_write::transform_mathml_to_omml` (`macos_offline/powerpoint.rs`; table cells through `prepare_mathml_for_table_cell`) |
| WPS Writer and Presentation | C# in the same add-ins (`WpsXmlInsertion`, `WpsNativeEquations`) | Rust `omml_write::wps_equation_docx`, `wps_field_instructions` (`wps/bridge.rs`, `macos_offline/word.rs`) |
| Document import | C# (the add-in writes each formula) | Rust `omml_write::word_equation_artifacts` / `wps_equation_docx` (`macos_offline/document_import.rs`) |
| Format conversion to OMML | C# (`WordFormulaService.FormatConversionCore`, `PowerPointFormulaService.FormatConversion`) | Rust `omml_batch` (`word_equation_docx`, `build_word_native_numbered_omml`) |
| LaTeX redraw | C# (`*.LatexRedraw*`) | Rust (`macos_offline/presentation_redraw.rs`; Word through the editor path) |
| Reading a Word equation (edit, conversion to VisualTeX, restore, agent reads) | C# `TransformOmmlToMathMl` + `MathMlToLatexConverter.ConvertWordEquation` | Rust `omml_latex` (`macos_offline::word_omml_to_mathml` + `mathml_to_word_equation_latex`; `omml_batch`, `automation/hosts.rs`) |
| Reading a PowerPoint native equation | C# `PowerPointNativeOmmlKernel.ReadNativeMathMl` + `ConvertWordEquation` | Rust `powerpoint_omml_to_latex` (run colours into Word run properties, then `omml_latex`) |

The Rust ports share one copy of the .NET behaviour they reproduce
(`src-tauri/src/office/dotnet`: the XLinq tree, .NET string semantics, Unicode
tables); each keeps its own XML reader and writer, as the C# code uses
different reader settings and writers.

A third converter, the TypeScript `apps/macos/src/office/omml/latexToOmml.ts`
(MathJax LaTeX → OMML), is still computed by the macOS Office editor for
PowerPoint (native equations and redraw) and sent as `ommlBase64` /
`ommlDocxBase64`. The app uses it only for an export without MathML
(fallbacks in `powerpoint.rs`, `word.rs`, `document_import.rs`), and the editor
never sends one: it always sends `wordMathMl` when it asks for an equation and
fails instead. Only `test:word-omml` exercises it.

Proposal, not done (it changes what runs): remove `latexToOmml.ts`, the `omml`
artifact and the fallbacks, so each platform has one writer and macOS's is held
byte-identical to Windows' by the forward corpus (§1).
