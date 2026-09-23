# CP07 editable general operator checkpoint

2026-09-13 after21:28 +08:00. Base checkout reference-3d207d7,
HEAD00b48e17b736c7e57a0a72e19e271df82906f89d. Non-Git source checkpoint only;
no engine staging/commit/push, production entrypoint, Office or installer changes.

New finite support: authored mathop bodies plus ordered lower/upper and
limits/nolimits/displaylimits tails. GeneralOperator owns a real editable body
Row and existing OperatorPart rows; fixed LargeOperator remains unchanged.
Insertion/wrapping preserves source and selected IDs, body and bounds copy
independently, the shared controls change placement, and whole selected operator
powers are grouped rather than becoming duplicate upper limits. Vertical movement
uses same-owner body/bound baselines, with a real editable body middle stop.
Details and explicit limitations: engine/specs/P07-general-operators.md.

Stable-source receipt:
engine/target/verification/2026-09-13T13-25-10-721Z/report.json
Source SHA25696ff06af0b2e5026170941af436163f2dc490feb461542b93accb5e5101955cf
145 inputs,sourceUnchanged true; all inputs rehashed after supplemental checks,
zero differences. Twelve engine/build/browser steps PASS. Original P00 production
protection step remains FAILED for previously observed external Office/evidence
changes. The strict total is not reported as PASS and the guard was not weakened.

Executed verification:
-354 Rust tests PASS,14 new general-operator tests. Authored source,malformed/empty
 input,wrap/copy/delete/history,body and bound patches,mode no-ops,powers,real
 font/style geometry,nested colors/phantoms,vertical navigation,forged and malformed
 wire input,600 source disturbances (seed0x0F3A2026) all covered.
-Actual Native/WASM lifecycle2151,source-session222,source-edit440,editor473 with
 100history states PASS.
-Expanded layout4546 requests/3063 scenes PASS,numeric difference0 at fixed font
 resources,exact SVG. New general-operator edit/describe/copy/mode/history tests
 run in both targets,not solely parser or static preview checks.
-Symbol/style parity2570 requests/1891 scenes PASS.
-Existing root,fence,accent,style,spacing,color,symbol,input and SVG/browser tests
 PASS. The expanded operator browser records52 trusted events,including new blank
 construction,body-pointer editing,same-owner up/down,limits controls,selection
 wrapping and exact undo/redo. Previous fixed-operator browser assertions remain.
-Separate fraction-form parity349/244 and browser33 trusted events PASS.
-Separate independent layout519/360 PASS;Office contract11sources/3interfaces/
 10routes/3negative mutations PASS,writes0.

engine/target/operator-browser/general-operators.png was read and visually
inspected: upright rank body with lower label,multiple sums sharing bounds and a
binomial operator body render as actual vectors. This is isolated Edge sandbox,
not Word,PowerPoint,installed WebView2 or a physical Windows IME rerun.
BrowserEdge152.0.4191.62;host sandbox-52ebe1dc29786568b4d8cdb5.
Native SHA56830c23b5d394add967f0721f556dc1b0b60cc5e3e585fd57b7f50833778ea6
Node WASMe9ed9801a46378ad1389fe717e4f3cfe543396125a266132d4eb4cd2e32179e1
Web WASM281d60d01d97f659057f778614f6ffe6fb08ff25214c2452fbb07c3bf29d43c0

Remaining registry/macros,exact ESINT-specific artwork,prime shorthand and middle
fences remain tracked language coverage,not falsely called supported. This
checkpoint does not declare the entire replacement or every product command
complete. P08 remains unimplemented; its contained Shift+Enter specification
forbids merging adjacent formula frames and requires one-owner/one-undo conversion.
Font/build outputs are excluded from the source archive;archive integrity and
per-input verification are in audit.json. Existing checkpoints stay immutable.
