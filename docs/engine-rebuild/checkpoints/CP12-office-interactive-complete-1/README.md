# CP12 Office interactive complete 1

Status: `complete`.

`p13Permitted=true`.

This checkpoint is the final P12 freeze for the Windows Office interactive editor. It supersedes `CP12-office-interactive-browser-stable-1` as the P12 completion checkpoint; the older checkpoint is intentionally retained unchanged as the earlier browser-stable / real-Office-blocked milestone.

## Final verification receipts

Final P12 receipt:

`engine/target/p12-verification/2026-09-15T23-58-21-717Z/report.json`

Receipt SHA256:

`08196cdd1ac004a7ab9afb3b6489c1581280612463471ab8ff32232965a89b76`

Verified P12 source digest:

`04310a3cf1369418473be0ff3d2bb42428dac027d414a02f78873a1912381b15`

The P12 verifier completed 9/9 gates with `sourceUnchanged=true`: Windows TypeScript, native system-font asset acceptance, production Office bundle, core browser review, core browser smoke, the existing Office formula-document contract, existing Office Session safety contract, existing Office apply-shortcut contract, and P11 output compatibility revalidation.

A separate final P11 verifier run also passed 11/11 with `sourceUnchanged=true`:

`engine/target/p11-verification/2026-09-16T00-00-03-717Z/report.json`

P11 receipt SHA256:

`c51ddd13a86d89008e68a1bd134fd6c2abab09e1dff61b162b77f0b034ab1e84`

P11 source digest:

`b91c7559a757c2e3167800506a4cf16d79af53ab14c6dbbfb22de109e100560d`

The frozen P11 checkpoint remains `docs/engine-rebuild/checkpoints/CP11-office-compat-stable-1/`.

## Source archive

Archive:

`visualtex-p12-office-interactive-complete-1.tgz`

Archive SHA256:

`ff3cb3c96f5c8c715a2a6fe8b89a5d3168153a5e09e1e8332e5c649bc022d04d`

Archive size: 177330 bytes.

The archive contains exactly the 19 source inputs recorded by the final P12 receipt. Read-only audit matched every archived file against the receipt and the current working copy, and found no forbidden generated/binary/font/WASM/certificate/private resource. Per-file SHA256 values are recorded in `audit.json` and in the final P12 receipt.

## Real Office acceptance

The frozen real Office acceptance record is:

`engine/specs/P12-real-office-acceptance.json`

SHA256:

`be1074cac6d22109af0a79a2c7ab571833ef6780d64982217799007f47a81442`

The dedicated Word document is `engine/target/p12-real-office/P12-Office-Core.docx`. Real Word Ribbon/UI acceptance passed for inline VisualTeX OLE creation and in-place edit, numbered display VisualTeX OLE, unnumbered display VisualTeX OLE, inline Word OMML, and numbered display MathType OLE. The MathType object was verified as `Equation.DSMT4`. Cancel, native X close auto-commit, companion/WebView reuse, save/reopen, and exact source recovery were exercised through the real Office path.

After save/reopen, the dedicated Word document retained 4 OLE objects, 1 OMath, 11 fields, 7 bookmarks, 1 VisualTeX `SEQ VisualTeXEquation` field, and 5 MathType numbering fields. Reopening the first VisualTeX formula through the real Ribbon restored exact source `x^{2}+1+z+q`.

PowerPoint acceptance used the currently supported `crossPlatformPicture` path. Real Ribbon insertion produced core SVG/PNG/MathML output with no MathJax ownership, inserted a `VisualTeX_...` shape with VisualTeX metadata in `AlternativeText`, and real `Edit Selected Formula` restored and updated the original shape without inserting a duplicate.

## System math font / external Office WebView fix

The final P12 source set includes the real Office font-access fix rather than the discarded temporary Office HTTP route. The external `https://127.0.0.1:43127/*` Office WebView is granted only the custom `read_system_font_asset` command through `office-session-editor.json` and `office-system-font.toml`.

The system-font asset reports `hasMathTable`; the core runtime accepts a system face as the mathematical layout font only when an OpenType MATH table is present. Times-family preference resolves through mathematical candidates such as XITS Math, STIX Two Math or Cambria Math, with verified Latin Modern Math as the safe fallback. Ordinary text/Chinese font use remains separate.

The native `system_font_asset_acceptance` gate passed in the final P12 verifier.

## Frozen Office boundary

At final verifier start and end there were zero tracked working-tree differences under the protected WordVsto, PowerPointVsto, Contracts, FormulaOleServer, VstoShared, WindowsOleBridge, Office Session/lifecycle/state/backend/pipe paths, `src/office/shared`, `src/office/api`, and `useOfficeSession.ts` boundaries. The frozen-path snapshot remained unchanged throughout verification.

`apps/windows/src-tauri/src/office/server.rs` has no P12 route change; its final source SHA256 is:

`7ba9bbba31486b8d33d1d9577976144a4d411867638b198ea911de1ede3c6dd4`

P12 therefore freezes an editor-surface/export-provider change, the minimal external-WebView font capability, and verified mathematical-font selection. It does not freeze any new Word document algorithm, numbering/reference algorithm, Office Session wire/state contract, COM/OLE contract, or Office server route.

The hidden `vsto-convert` path remains legacy at the P12 boundary.

## P13 boundary

P13 is now permitted, but has not been started in this checkpoint. Before making a P13 change, read the P13 section of `docs/engine-rebuild/VISUALTEX_ENGINE_PROJECT_PLAN.md` and this checkpoint, then establish a fresh protection baseline. P13 may switch only the hidden `vsto-convert` / batch import / redraw / automatic-conversion outlet to a non-interactive core instance. It must not use P12 completion as permission to refactor Session semantics, `WordFormulaService`, Word/PowerPoint VSTO algorithms, COM/OLE storage, numbering or references.

## Repository safety

Checkpoint HEAD reference: `d1f0aff8db9275c90dc18c7318dc2a4d7bbbd760` on `reference-3d207d7`.

No reset, clean, stash, branch switch, pull, rebase, merge, commit, push, installer overwrite, registry/certificate mutation, or release action was performed while finalizing this checkpoint. Existing dirty/untracked files were preserved. No user Word process or document was closed, saved, or manipulated by the P12 finalization work.
