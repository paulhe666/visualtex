# macOS Office performance budget

This document defines the release gate for the resident macOS VisualTeX Office editor.
It measures the complete user-visible path, not an isolated helper or an average that can
hide a slow interaction.

## Hard budgets

| Operation | Start marker | Completion marker | Hard maximum |
| --- | --- | --- | ---: |
| Open an existing formula | Physical double-click event or Edit command dispatch | `editor-ready.json.epochMs` after the editor is hydrated, visible and focused | 300 ms |
| Apply a newly created formula | Apply button/keyboard activation timestamp | `apply-backend-complete.epochMs` after the Office callback has committed the result | 700 ms |
| Apply an edited formula | Apply button/keyboard activation timestamp | `apply-backend-complete.epochMs` after the Office callback has committed the result | 700 ms |

The app must already be resident because the installed Word and PowerPoint add-ins prewarm
VisualTeX during host startup. A missing resident app is a startup-health failure and must not
be mixed into the interaction-latency result.

## Required scenario matrix

Release acceptance covers these scenarios independently:

- `word-image`: inline and display SVG image formulas.
- `word-omml`: inline and display native Word equations.
- `powerpoint-svg`: PowerPoint SVG formulas.

Every representation must pass in both density profiles:

- `sparse`: the host contains only the formula needed for the measured interaction.
- `dense-50`: the Word document or PowerPoint presentation already contains at least 50
  committed VisualTeX formulas before any measured create or edit operation begins.

The dense fixture must remain open and intact throughout its measured series. All 50 formulas
must have been created through the installed VisualTeX add-in's normal **New Formula → Apply**
workflow. Tests must not populate the fixture by writing raw OMML, inserting arbitrary SVG/PNG
artwork, copying prebuilt OOXML parts, or fabricating VisualTeX metadata/bookmarks/tags.

Before timing begins, the harness must prove that all 50 formulas are genuine editable
VisualTeX formulas by reopening each formula through the same public edit path used by a user,
checking that the editor request contains the expected formula identity and persisted LaTeX,
and closing the editor without mutation. A fixture with fewer than 50 successfully reopened
formulas is invalid and produces no performance result.

For `dense-50`, edit opening and edit Apply are measured independently at these target
positions:

- `first`: formula 1.
- `middle`: formula 25.
- `last`: formula 50.

Create Apply is measured while all 50 existing formulas remain in the host. The newly created
formula becomes formula 51 and must not invalidate or rewrite the prior 50 formulas.

Every scenario records all three operations: `edit-open`, `create-apply`, and `edit-apply`.
A scenario may add more detailed variants, but it may not merge unlike representations,
density profiles, or target positions into one timing sample.

## Sampling rules

1. Restart the Office host before a suite, load the installed add-in, and confirm that the
   resident VisualTeX process is healthy.
2. Run exactly one unreported warm-up interaction for each scenario and operation.
3. Record at least 10 consecutive measured samples for each representation, density profile,
   target position, and operation.
4. Do not trim outliers, retry failed samples, or replace a slow sample with a later run.
5. Every measured sample must complete successfully and remain within the hard maximum.
6. Report count, minimum, median, p95, and maximum. The maximum determines PASS or FAIL.
7. Keep the Office document/presentation valid after every Apply and verify the persisted
   LaTeX by reopening the committed formula.
8. Record the Git revision, app version, Office version, macOS version, host, formula kind,
   and whether the editor window was reused.

## Result format

The verifier consumes one JSON document with this shape:

```json
{
  "schema": "visualtex-office-performance-budget-v1",
  "revision": "<git revision>",
  "samples": [
    {
      "scenario": "powerpoint-svg",
      "variant": "display",
      "densityProfile": "dense-50",
      "existingFormulaCount": 50,
      "editableFormulaCount": 50,
      "fixtureSource": "visualtex-create-apply",
      "targetPosition": "middle",
      "operation": "edit-open",
      "durationMs": 248,
      "success": true,
      "warmup": false
    }
  ]
}
```

`operation` is one of `edit-open`, `create-apply`, or `edit-apply`. `densityProfile` is
`sparse` or `dense-50`. Every dense record must include
`fixtureSource: "visualtex-create-apply"`, `existingFormulaCount >= 50`, and
`editableFormulaCount >= 50`. Dense edit records require `targetPosition` to be `first`,
`middle`, or `last`; dense create records use `targetPosition: "new"`. Warm-up records may be
included for auditability but are excluded from the measured count. The strict verifier
requires the complete scenario matrix by default; `--scope powerpoint-svg` is available only
for focused optimization work and is not a release acceptance substitute.

## Diagnostic stage timings

Stage timings such as request serialization, AppleScriptTask write/launch, URL receipt,
frontend hydration, SVG/OMML materialization, dispatch-file reads, VBA insertion, and focus
restoration should be recorded whenever available. They explain regressions, but only the
end-to-end timestamps above decide whether the user-facing budget passed.
