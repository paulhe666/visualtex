# CP05 — font/layout and double-target geometry checkpoint

Date: 2026-09-13. Product baseline `00b48e17b736c7e57a0a72e19e271df82906f89d`. Uncommitted source archive only; no Git commit/push, release, installation, production editor switch or Office document operation.

## Implemented isolated core

Actual injected font byte identity, face index/default variation coordinates, bounded outline cache, ttf-parser MATH constants/outlines, rustybuzz horizontal text shaping, independent K0 box layout, display/text/script/scriptscript sizes, fraction and script placement, math spacing, layout cache, immutable scene, carets/hit regions, structural fraction navigation, and path/rect-only SVG. Output geometry carries document/revision/source/font identities. Natural dimensions are CSS pixels, scene coordinates are em. Missing glyphs and unsupported/draft source fail explicitly.

The scene consumes the parser's authoritative exportability rule. Empty grouping does not fill required arguments. Review regressions concerning blank formulas, forced fraction style, vertical fraction navigation and malformed scene data were retained and rerun. Concurrent review corrections are included in this combined snapshot; failed exact-text edits did not overwrite those changes.

## Actual commands and results

From `engine/`:

- `node tools/build.mjs`: PASS, full Rust suite and rebuilt native/WASM bindings; preexisting protocol 2151, source-session 222, source-edit 440, editor 473 request/case comparisons all passed.
- `cargo test --workspace --locked --quiet`: 189 tests passed, 0 failed; 23 runner result blocks including empty/doc-test targets.
- `node tools/layout-parity.mjs`: PASS, 183 requests, 106 validated scenes, actual native and actual WASM. Maximum numeric difference 0; SVG identical. Includes Chinese formula `\frac{x_1+\text{速度}}{y^2}`, all four styles, scaling, changed-source cache invalidation, undo, composition-state commit, and error atomicity.
- `node tools/layout-parity-review.mjs`: PASS, independent 519 requests, 360 valid scenes; numeric tolerance relative 1e-9; SVG and identities exact.
- P00 protection guard with the recorded Git transition: PASS (12 groups, 12 artifacts, 65 original dirty files), checked 2026-09-13T05:25:57Z.
- Office contract snapshot: PASS (11 sources, 3 interfaces, 10 routes, 3 negative mutation tests; 0 writes).

Native CLI SHA256: `4965e99a6a247800605b0c63b189b3d07b3add544f5d8688cddb3fa075b2197d`.
WASM SHA256: `31ff40846ea961a412046b892da85a407e274b91c715a3052f5bc43df0b106d4`.
Implementation layout transcript SHA256: `76f2aea94f9cf6949895aec7bacdaf6f33e7e90dc83b67ca3cc2610ea7e7859d`.
Review layout transcript SHA256: `1e93874c7267ae2ffd7932bb7fda73d7348eba4316170cf3030bcb517cfaa580`.

## Fixed local resources

Latin Modern Math: existing `apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf`, SHA256 `6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f`, face 0, 1000 units/em.

Noto Sans SC: installed `C:\Windows\Fonts\NotoSansSC-VF.ttf`, SHA256 `763146584cf0710223441356b4395e279021b0806c196614377a7a0174ae074a`, face 0, 1000 units/em, default wght=100. This is a read-only test resource, not a final product typography decision. Font bytes are not included in this source archive. Both targets consumed precisely these same bytes; no font-family-name substitution occurred.

## Remaining acceptance and scope

This checkpoint proves headless K0 output, not a finished MathLive replacement. Browser rendering/screenshots, actual keyboard/pointer/OS IME behavior, multi-window UI, complete font policy, larger mathematical structures and Office compatibility are not established by these tests. RTL/bidirectional text and custom variation settings are explicitly outside current K0 support. The glyph/layout cache avoids repeating unchanged layout work, but tree fingerprinting and scene projection still traverse the current formula; no constant-time edit-performance claim is made.

Next: verify vector output visually and mount an isolated SVG/input sandbox; then execute genuine keyboard/pointer operations and separate native Chinese IME checks. Keep production MathLive/MathJax paths unchanged until their plan-stage gates are met. Do not use this archive as permission to install, publish or commit.
