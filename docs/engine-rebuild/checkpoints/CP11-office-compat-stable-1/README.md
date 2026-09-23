# CP11 — Office compatibility stable 1

This checkpoint freezes P11 only. It proves that VisualTeX's independent core can produce the data required by the existing Office consumers without taking over a real Office Session.

## Frozen source

- Verification receipt: `engine/target/p11-verification/2026-09-15T13-18-57-359Z/report.json`
- Source digest: `b91c7559a757c2e3167800506a4cf16d79af53ab14c6dbbfb22de109e100560d`
- Source archive: `visualtex-p11-office-compat-stable-1.tgz`
- Archive SHA256: `5c29950423d7189727444cc15426b5b0fe76c53024d844c5a77e804150f91e87`
- Verified source files: 13

The archive was audited against the receipt and the live working files after verification. It contains no target/node_modules/dist/pkg binaries, fonts, DLL/EXE/MSI/WASM, credentials or private runtime resources.

## What passed

The independent core now owns Presentation MathML generation for its supported retained tree. Native and WASM generated identical output across the P11 corpus. Incomplete/unsupported formulas fail explicitly instead of falling back to MathJax.

The host `coreOfficeCompatibility` adapter preserves the existing `OfficeExportResult` names and units, retains the OMML MathML-only fast path, derives SVG width/height/baseline from the same core layout snapshot, and derives MathType ruler stops directly from core grid-column geometry at 24 MTEF units per CSS pixel.

The current x64 `VisualTeX.WordVsto.dll` was used only as an offline consumer, not through Word automation. Its existing `MathTypeMtefCodec` accepted the core aligned MathML, preserved semantic signature, returned ruler stops `625,2094`, and kept the requested 10.5 pt MathType Full Size. Its existing `OfficeOlePreview` accepted path-only core SVG and produced an EMF that passed the existing vector-only validator. The DLL used for that acceptance had SHA256 `35EDB78232272963D22FDFAFD4BF21A63DDF749F6216A0CC527CE91D446083F8`.

The P11 verifier also passed TypeScript compilation, existing MathType alignment geometry, Office metadata/formula-document/session-safety/apply-shortcut regressions, and 51/51 targeted .NET MathML/MathType consumer tests.

## Deliberately not done in P11

No Office Session was created or patched by the new core adapter. No Ribbon/UI action was performed. No Word document was changed. `WordFormulaService`, VSTO numbering, COM/OLE identities, Session DTO/state transitions and hidden converter behavior were not modified.

Custom artwork that requires SVG mask/clip semantics not supported by the current Office vector consumer remains fail-closed for full vector export; its immutable `ommlFallback` remains available for semantic MathML where defined. The Office consumer was not relaxed to hide this limitation.

## Next gate

P12 may now integrate the core only into the interactive Office editor runtime. `vsto-convert`/hidden batch conversion must remain on its existing path until P13. Session load must not mark dirty, every async result must be tied to the active Session/formula/revision, and Apply/Cancel/X/auto-commit must continue to use the current Session mechanism. Real Office acceptance in P12 must use a dedicated test document and real Ribbon/UI, never the user's existing document.
