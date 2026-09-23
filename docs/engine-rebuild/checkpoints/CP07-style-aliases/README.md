# CP07 style-alias compatibility checkpoint

2026-09-13 after 20:11 +08:00. Base checkout reference-3d207d7, HEAD
00b48e17b736c7e57a0a72e19e271df82906f89d. Non-Git source archive; no staging,
commit, push, installed UI change or Office document operation.

## Added since the immutable CP07-symbols checkpoint

Seventeen authored Unicode math style spellings now use real Style nodes:
symup/symit/symbf/symbfup/symbfit/symbb/symcal/symscr/symbfcal/symbfscr/
symfrak/symbffrak/symsfup/symsfit/symbfsfup/symbfsfit/symtt.

Their source, grouping and comments remain unchanged unless explicitly edited.
Canonical command selection preserves all fifteen MathVariant values, including
sans combinations and bold script/Fraktur that previously had lossy mappings.
Greek theta/nabla/partial/kappa/phi variant mappings were completed without
collapsing ordinary phi/kappa into their different variant symbols.

Actual font tests initially found missing lowercase script z in Latin Modern Math.
The full sample was retained and passed with the previously implemented explicit
pinned STIX fallback; it was not replaced with another symbol or fake glyph.
Six new tests cover all aliases, source/glyph/history/copy behavior, and 340 seeded
local source edits (0x57a12026). The sandbox's existing style host handles the new
controls. Production toolbar and fonts remain unchanged.

The capability audit tool gained an explicit --symbol-fallback option. It still
imports only product command data and executes the new core, not the legacy engine.

## Executed stable-source verification

Receipt: engine/target/verification/2026-09-13T12-09-40-488Z/report.json
Source SHA256: 8e5b1c82b3766f897ad54df0676ef47239f9f2c0a14543edd9f165fa56e10d21
sourceUnchanged: true; 135 input files.

- 317 native Rust tests PASS.
- Native/WASM protocol 2151, source-session 222, source-edit 440, editor/history
  473 requests including 100 history states PASS.
- Existing full layout 3945 requests / 2689 scenes PASS.
- Expanded symbol/alias layout 2570 requests / 1891 scenes PASS, 175 added atom
  spellings, 22 Unicode integral spellings and 17 style aliases. Maximum numeric
  difference 0 under identical fixed resources; SVG exact.
- Twelve engine/build/browser steps all PASS, including old structure, input and
  SVG zoom/DPR/scroll regressions and current symbol/alias controls.
- Actual symbol/alias browser runner: 21 trusted events, selection wrapping into
  sans-bold-italic, blank bold-script body typing, undo/redo, source recovery and
  SVG, through actual controls/key/pointer input. The current screenshot was read
  through Devspace and visually inspected.
- Separate Office contract snapshot: 11 sources, 3 interfaces, 10 routes,
  3 negative mutation tests PASS; writes 0.

The strict overall receipt is FAILED only at the last original P00 protection
step. It reports preexisting external Word numbering/VSTO and evidence drift;
see its protection.log for exact paths. No baseline assertion was weakened and no
external content was overwritten by this continuation. Do not call this a globally
green production gate or identify an external writer without evidence.

Native CLI: 98529279a9c810033489e0f0adc4db728f15cd682ab4e8941d3240044b783db8
Node WASM: f0f7990667c54e09993ed3d12e5e7591c99f1787762c83e8c3afc5b0a0e52d2f
Web WASM: 0200183e57b889515482c8ac43f8b28e6448a57416efe2cb1311a7ff849d99f1
Symbol transcript: 0115d26e4f69bbb6d9b27e2d55590fd871dc30f6851a8c37644940bf4bbe6d92
Browser: Edge 152.0.4191.62, sandbox-b6324f9cceac19e332dde79c.
Report and screenshot: engine/target/symbol-browser/report.json and symbols.png.

## Updated actual coverage, not a replacement-complete claim

Audit receipt: engine/target/capability-audit/2026-09-13T12-11-40-592Z/report.json.
Of 411 static registry records (405 unique IDs), 363 preview samples parse/layout,
346 export SVG; 359 templates parse/layout, 179 export SVG. Many templates are
intentionally unfilled slots. These are fixed-font source/layout checks, not
411 command-editor acceptance tests, nor a percentage completion of the project.

Remaining diagnostics include environments/row breaks, middle delimiters,
binom/mathop and other structures/decoration, general physics/package macros and
ESINT-specific integral artwork. Legacy semantic Unicode placeholders that render
do not prove that their corresponding authored ESINT commands are implemented.

## User's P08 decision

The master plan P08 section and engine/specs/P08-contained-multiline.md require
Shift+Enter to create/insert internal environment rows inside one formula frame,
never join adjacent frames, with one owner and single-transaction undo. Two-frame
isolation/IME/copy/save-reopen tests are specified. P08 implementation and these
acceptance tests remain NOT_RUN; no matrix or Shift+Enter completion is asserted.

All engine changes remain uncommitted. Keep both CP07-symbols and this checkpoint.
No fonts, generated binaries or logs belong in the source archive. Full main-app,
Office/Ribbon, WebView2, macOS, MathML conversion, installer, performance and
MathLive removal acceptance remain later stages.
