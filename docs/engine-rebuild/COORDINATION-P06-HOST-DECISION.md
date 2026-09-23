# P06 single-host decision and next checks

Implementation lane decision: reuse `web/host/editor-host.ts` / EngineEditorHost as the sole owner bridge. Do NOT create a competing editor-controller.ts. The current adapter already matches BrowserInputBridge/InputPort, and a fresh strict tsc check passed. The previous MathInputBridge export mismatch was caught, not ignored.

Review lane may continue owning editor-host.ts and its independent browser runners. Implementation lane retains input-bridge.ts, core clipboard/scene queries, and will supply sandbox assets/server plus native-IME testing. Existing svg-view.ts is reused unchanged unless an independently reproduced renderer bug requires a targeted edit.

Available core queries now include navigation (selection-aware, including node selection), boundary (row/text/document Home/End), exportSvg (read-only on the current token), in addition to hitTest/vertical/selection. Runtime copySelection returns source plus revision/selectionEpoch, not live node IDs.

Specific host issues to test and harden before P06 acceptance:
1. If rendering fails after an accepted core transaction, the source UI still must receive the new snapshot. Mark old visual geometry stale/disable pointer edits; do not display the old formula as the current revision. Initial unsupported source should remain editable in the source pane rather than destroy the entire session.
2. Private clipboard flavor is LaTeX. Prefer pasteLatex for that flavor even when the current caret is in Text; otherwise copying a text fragment and pasting it into Text inserts literal `\\text{...}` markup. Plain text in Text mode remains literal.
3. Shift-Up/Down must preserve the range anchor. The navigation query resolves the source of node selections in the core. Use boundary query for Home/End.
4. Export should use queryScene/exportSvg rather than creating a replacement token solely to export unchanged geometry.
5. Keep viewport/caret-only operations out of layout; blink must not trigger layout. Bound diagnostic logs and dispose only owned listeners/views/sessions.

No commits/pushes, production application edits or Office operations. Real Windows IME remains separate from Chromium protocol composition. Source checkpoints record only actually executed results.
