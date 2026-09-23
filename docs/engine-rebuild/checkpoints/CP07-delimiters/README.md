# CP07.2 — editable font-driven fences

2026-09-13 15:05 +08:00. Independent engine only, based on `reference-3d207d7` HEAD `00b48e17b736c7e57a0a72e19e271df82906f89d`. No Git staging, commit or push; no original user document or installed application was changed.

Added a real Fence node with independently editable body and exact left/right delimiter spellings. Parsing covers paired, nested, one-sided and invisible `\\left`/`\\right`, including aliases and deliberately mixed interval delimiters. Incomplete/unknown/missing delimiter source is retained with diagnostics rather than repaired and committed. Source editor input `(x)` stays literal source; automatic pairing is an explicit mathematical edit command.

`TypeDelimiter` creates pairs for a typed (, [ or {. At a completed structure's end, the matching closer exits the nearest pair without adding another closer, revision or undo item. A mismatched inner closer remains literal and cannot jump to a matching outer fence. Exiting can pass through completed final scripts. Range wrapping retains selected identities; copy/delete/undo use the existing transaction and position-map rules. Text-mode bracket typing remains literal.

Layout reuses real OpenType MATH variants and assemblies and centers each visible delimiter on the math axis. A hidden dot delimiter adds no glyph or width. The independent sandbox has pair/absolute-value/one-sided controls and the existing authoritative EngineEditorHost handles typed bracket events. No second editor owner, old-engine fallback or production integration was added.

Verification receipt: `engine/target/verification/2026-09-13T07-03-43-563Z/report.json`.
Input SHA256: `2c4996e261d11a70099669e7c9a699b4fe6cb0a3f4a5c088166fe1fecaa394d5`.
All seven steps passed, source inventory unchanged during validation:

- 223 native Rust tests; 10 new delimiter tests include pairing, exit, aliases, comments, unknown source, nested roots, selection/identity, source patching and font geometry.
- Native/WASM lifecycle 2151, source-session 222, source editing 440, edit/history 473 requests.
- Native/WASM layout 354 requests, 209 scenes: zero numeric difference for fixed resources and exact SVG parity.
- Existing K0 and root sandbox UI tests passed.
- New actual browser key/control test verifies typed bracket pairs/exit, mismatch retention, selected-expression wrapping, text-mode literals, a tall single-sided fence, SVG export and invalid source undo.
- Input fault injection 15/15 and 18 zoom/DPR/scroll geometry combinations passed.
- Office/production protection 12 groups, 12 artifacts, 65 original dirty files all pass with authorized Git transition acknowledged.

The owned browser screenshot `engine/target/fence-browser/single-sided-root.png` was opened and visually inspected. It shows a single left brace around an indexed root with nested fractions. Browser-input tests do not claim native Windows IME; that is separately recorded in immutable CP06-native-ime-verified. No new Word/Ribbon/WebView2/macOS acceptance is implied.

Gate: continue P07.3 accent/annotation structures. P07 remains incomplete until the other required structures and product behavior are covered. Large delimiters unavailable in a given font fail explicitly rather than being anisotropically stretched or silently replaced. General macros, matrices, MathML/Office output compatibility, production main/Office/hidden-converter integration and final package acceptance remain later stages.
