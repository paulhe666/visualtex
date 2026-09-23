# CP12 Office interactive browser stable 1

Status: `browserStable-realOfficeBlocked`.

This checkpoint freezes the P12 Windows Office interactive-editor integration after all non-Office-document gates passed. It is not a P12 completion checkpoint and does not permit P13.

Verified source receipt: `engine/target/p12-verification/2026-09-15T14-09-41-765Z/report.json`.

Verified source SHA256: `fb4a555f7f94da861da5ed65ae5a7362187083c2fe0489f18874f77248952e4e`.

Source archive: `visualtex-p12-office-interactive-browser-stable-1.tgz`.

Archive SHA256: `6c947e380bfaddf78e54921db7543084cb17bc53d1361e60cac937c6b983cfc6`.

The archive contains exactly the 11 P12 verified source inputs from the receipt. The read-only archive audit matched every file, found no forbidden generated/binary/font/private resource, and confirmed the working copies still match the receipt.

## What is proven

The production Windows Office companion bundle contains the verified VisualTeX core WASM, Latin Modern Math resource and CoreMathEditor chunk. The interactive Office runtime can select the core editor while `vsto-convert` stays on the legacy path and does not load the core editor/WASM.

The isolated production-bundle browser acceptance proves clean Session hydration, real browser keyboard input through the core input bridge, immediate source persistence through the existing PATCH route, same-revision SVG/Presentation-MathML/PNG/baseline output, stale export rejection across rapid edits and Session switches, core Undo/Redo through the existing Office controls, theme synchronization, explicit Apply through the unchanged Session commit handshake, and Cancel through the unchanged cancellation path.

P11 output compatibility was re-run from P12 and remains PASS, including the existing MathType MTEF and Office EMF consumers. Existing Office formula-document, Session payload-safety and Ctrl+S shortcut regressions also remain PASS.

No tracked differences were introduced under the frozen Word/PowerPoint VSTO, Office Contracts/VstoShared/OLE bridge, Tauri Office Session/API/lifecycle paths, `src/office/shared`, `src/office/api`, or `useOfficeSession.ts`. P12 therefore changes the editor surface/export provider only, not Office document algorithms or wire/state contracts.

## Why P12 is not complete

The project plan requires the final P12 gate to use real Word Ribbon/UI on dedicated test documents and cover VisualTeX OLE, Word OMML and MathType, inline/display, numbered/unnumbered cases, plus native close/window reuse, IME behavior, and save/reopen structure inspection.

At freeze time the companion health chain itself is valid (the shared configuration has been restored to the trusted formal certificate and recent VSTO logs show `health-passed`), but another real VisualTeX Office editor window is active for a separate Word task. Reusing or closing that window would violate the Office safety boundary, so the verifier records `BLOCKED_CONCURRENT_OFFICE_EDITOR` and `p13Permitted=false`.

Do not start P13 until the concurrent Office editor is gone and the full real Ribbon/UI P12 matrix passes on an explicitly owned test Word process/document. Do not close or save the user's existing Word documents to make the gate available.

## Boundaries kept

No reset/clean/stash/branch switch/pull/rebase/merge was used. No commit, push, installer, registry mutation, certificate installation, release action, or user-document save/close was performed for this checkpoint.
