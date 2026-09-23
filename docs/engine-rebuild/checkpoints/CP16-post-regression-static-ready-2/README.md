# CP16-post-regression-static-ready-2

Status: `STATIC_PASS`; installed Office acceptance is still required.

This checkpoint supersedes the earlier P16 package-static candidate after the user installed that earlier package and reported visible regressions in the autonomous formula editor. It does not rewrite or invalidate the historical P15 checkpoint; instead it records that the current P16 package contains explicit post-P15 regression fixes and therefore must be accepted as new package bytes.

The repaired candidate restores a product-level command suggestion surface while retaining core-owned command state and insertion transactions. Typing a command such as `\\frac` now presents the command text, label and a VisualTeX-core mathematical SVG preview. Confirming the fraction produces two visible editable structural slots instead of blank geometry. The production interactive surface remains core-only (`math-field` count zero); MathLive was not restored as a runtime editor or renderer.

Formula-font semantics now separate structural mathematical metrics from product-facing visible font faces. Times New Roman, Palatino Linotype and Arial use their real Windows Regular, Italic, Bold and Bold Italic font files for common Latin/Greek variables, digits, ordinary symbols and fixed-size delimiters/radicals/marks whenever the selected family contains an adequate glyph. Structures that genuinely require OpenType MATH constants, designed variants or assemblies continue to use a verified mathematical face. Semantic mathematical alphabets such as `\\mathcal`, `\\mathbb` and `\\mathfrak` retain their dedicated mathematical-alphabet semantics. STIX is strict: when real STIX Two Math is not installed, the settings UI reports it unavailable rather than silently substituting Cambria or another family.

Machine-readable regression evidence:

- formula-font truth: `engine/target/p16-formula-font-truth-review/2026-09-16T08-41-30-424Z/report.json`, SHA256 `39FFA9169F092C0E7C61DD0F0C33BFC96CADC1B7D6D6E5E10A046A49F4234614`;
- command/structure visual regression: `engine/target/p16-core-visual-regression-browser/2026-09-16T08-30-19-600Z/report.json`, SHA256 `98437678A23F57CC6518175AED30F081C5FD4C4F026FDE90388939FDCC78E8D3`;
- P16 aggregate static verification: `engine/target/p16-static-verification/2026-09-16T08-42-50-578Z/report.json`, SHA256 `35ED664C35257BEE7333095CF8AB1E60E3F1F4CDDDBB8EE95CBB0C65DE06998D`, 5/5 PASS.

Current core browser WASM is `visualtex_engine_wasm_bg-BtL3PE-e.wasm`, 7,242,419 bytes, SHA256 `1D106B40B8696AF80F47C128C2D3811496D4ECCDE9D364C139F0FED2397743EA`. Desktop and Office production bundles contain byte-identical WASM.

Current installers:

- full: `apps/windows/src-tauri/target/release/bundle/nsis/VisualTeX_1.2.7_x64-setup.exe`, 348,762,103 bytes, SHA256 `A014BF2BC4FF37726E69920E17F84F920DE83F7F72769977B34AD91E069C2F49`;
- No-OCR: `apps/windows/src-tauri/target/release/bundle/nsis/VisualTeX_1.2.7_x64-no-ocr-setup.exe`, 121,454,489 bytes, SHA256 `5026CF3917F0BFB8EA36334A1823C8D56E20CD64DBCCAC6E214A2FF5675F12BC`.

The fixed Office x64/x86 MSI, Word/PowerPoint/OLE/contracts binaries, bridge, VSTO runtime and private MathType runtime remain the accepted P16 fixed inputs. Protected Office source paths remained unchanged throughout aggregate static verification.

No installer from this rebuilt candidate has been executed. No installed VisualTeX application was overwritten, no user Word or PowerPoint process/document was touched, and no commit, push or release publication was performed. `P00→P16` must not be marked complete until this exact rebuilt installer passes real installed Word/PowerPoint Ribbon/UI acceptance.
