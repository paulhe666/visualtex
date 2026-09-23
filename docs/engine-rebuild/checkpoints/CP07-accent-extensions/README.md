# CP07 decoration extension checkpoint

Captured 2026-09-13 after 21:10 +08:00. Checkout reference-3d207d7,
HEAD 00b48e17b736c7e57a0a72e19e271df82906f89d. Uncommitted source archive only.
No production entrypoint, installed package, Office document or Git ref changed.

Added supported source/edit/render forms: overgroup, undergroup, overleftharpoon,
overrightharpoon, overlinesegment, underlinesegment. Group and harpoon marks use
real specified Unicode glyphs and horizontal MATH assemblies. Line segments are
three authored vector rules with font-derived thickness and gaps, not a substitute
square-bracket glyph. Explicit missing-glyph fallback is now honored for accent
fonts. No old renderer was invoked; only public command/character data was read.
Finite metrics/policies and exact scope are in engine/specs/P07-accent-extensions.md.

Stable-source receipt:
engine/target/verification/2026-09-13T13-08-21-181Z/report.json
Source SHA256: 3f7961d503aa04b3a918743a500be15176c86c079da86fce30fb9b1025ddc1a8
143 source inputs; sourceUnchanged true. All143 were rehashed after supplemental
tests and remained identical. Twelve engine/build/browser steps PASS. Strict total
is still FAILED at the unchanged original P00 protection step for external Office
and evidence differences; no guard assertion was relaxed or bypassed.

Executed:
- 340 Rust tests PASS, eight new extended-decoration tests,360 generated source
  edits seed0xACCE2026. Source/IDs/copy/delete/history/placement/genuine long marks,
  endpoint thickness and orientation,color/phantom/cache/error atomicity covered.
- Actual Native/WASM lifecycle2151,source-session222,source-edit440,editor473 and
  100 history states PASS.
- Expanded layout4323 requests/2917 scenes: numeric difference0 at fixed font bytes,
  exact SVG. Includes new short/long/nested/colored marks and native/WASM editor
  actions for each of the six spellings, independent copies and source edits.
- Symbol/style parity2570 requests/1891 scenes PASS.
- Existing root,fence,operator,style,spacing,color,symbol,input and SVG regressions
  PASS. Expanded actual accent browser now records80 trusted browser events,
  preserving its previous cases and adding all six new controls,body typing,
  pointer edits,completed-chain exit and undo/redo.
- Separate fraction-form parity349 requests/244 scenes and fraction-form browser
  (33 trusted events) PASS; separate independent layout519/360 PASS.
- Separate Office contract check11 sources/3 interfaces/10 routes/3 negative
  mutations PASS,writes0.

The extension screenshot engine/target/accent-browser/accent-extensions.png was
opened and visually inspected. Font group shapes,endpoint-ticked lower segment,
left harpoon and lower group are visible without clipping. This is sandbox UI,
not actual Word/WebView2 or a new native OS IME test.
Native SHA: 4a3d8fc572ef7683395c547232cd211b22d5b6695c7c6a0fa720bd74ea538f18
Node WASM: 86e7e7be327254d256e520ef8995eec3b5b3b79991cd229489f7f4d662b82a6a
Web WASM: f9cc057353ab4f47b5c563d38ed68301fccd9124e1c9f74f42557439302563bc
Host: sandbox-1d0ec478313104bf7b352b1a, Edge152.0.4191.62.

The first added test compilation identified a wrong test helper field name and a
temporary snapshot borrow; both were corrected in the harness. No assertion was
removed to pass geometry. Every tested source input is included in the archive;
font/binary/generated target outputs are excluded. Full P07 compatibility,
general mathop, ESINT-specific artwork and P08 Shift+Enter remain distinct work.
P08 still means one formula frame with internal rows,never adjacent frame merging.
No engine staging/commit/push or production engine switch occurred.
