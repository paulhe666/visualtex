# P09 immutable custom-symbol foundation

Status: implemented and accepted in the isolated P09 engine/sandbox. Legacy vector, clip, erase and explicitly-bound text artwork are covered below; production application wiring remains a later P10 concern.

## Ownership

A custom symbol inserted into a formula is bound to an immutable `SymbolAsset` snapshot. The snapshot contains the exact legacy definition JSON, a content SHA-256, math role, limits behavior, metrics and compiled self-owned vector geometry. Rendering never looks up a later same-name library record. A later library change therefore cannot silently alter an existing formula.

The document/session owns a bounded `SymbolCatalog`. Node identity and asset identity are separate. Copy/duplicate allocates fresh mathematical NodeIds but retains the same immutable asset content identity. Rich clipboard may carry the exact frozen definition; the destination validates and freezes it before inserting, and conflicting same-name/different-content resources fail atomically.

## Language and layout

Registered commands are parsed as native `CustomSymbol` nodes. Unknown commands remain unknown; the parser never invents a symbol from a command name alone. Reserved core/VisualTeX commands cannot be shadowed.

Role participates in normal atom spacing. Operator symbols use the existing ordered script/limits tail and may use `auto`, `limits` or `nolimits` as their frozen default. Explicit formula-local limits directives override the default without modifying the asset.

Metrics are authored layout inputs. Artwork is a separate scene resource and never masquerades as a font glyph. Font-size changes scale the natural formula output without rebuilding a different asset.

## Vector safety

Supported vector artwork is parsed and validated by Rust. Raw authored SVG/HTML/CSS is never inserted into the DOM or output SVG. The host independently validates artwork node ownership, SHA format, finite geometry, path bounds and paint order before replacing a visible scene.

Current supported legacy shapes include path geometry (M/L/H/V/C/S/Q/T/A/Z, including relative forms), line, circle, ellipse, rectangle, rounded rectangle and polygon. Fill/stroke/stroke-width/line-cap/line-join semantics stay as real SVG vector paint under affine transforms instead of being pre-expanded into approximate outlines. Elliptical arcs are deterministically converted to bounded cubic geometry. Color inheritance remains explicit scene paint.

`clipRect` is supported without flattening it to a final bounding box. Imported/base `transform.matrix` stays inside the clip, while user translate/scale/skew/rotate stays outside it: outer user transform -> clip rectangle -> inner/base matrix. Rust SVG export and the self-owned DOM view preserve the same nesting with generated `<g>`/`<clipPath>` elements; authored SVG/HTML/CSS is still never injected.

Legacy `operation:"erase"` is supported as ordered, asset-local vector compositing. The engine keeps each CustomSymbol node's artwork as one contiguous ownership group and generates a luminance mask only for that immutable asset instance: paint contributes white, erase contributes black, and later paint can restore local ink. Glyphs/rules outside that CustomSymbol never enter the mask. Rust and the independent host validator reject cross-glyph re-entry, non-monotonic path indices, or mismatched instance transform/color, so an erase layer cannot subtract neighboring formula content. This preserves ordered erase semantics without destructively flattening paths into approximate boolean geometry.

Legacy `text` artwork is supported only through an exact explicit runtime font binding. The persistent symbol definition keeps its original `fontFamily/fontStyle/fontWeight/fontSize` fields unchanged, but layout must map that exact family/style/weight tuple to a font already registered from real bytes. Registered font identity is content-addressed (`SHA-256:faceIndex`); Rust shapes the text with rustybuzz and extracts each glyph outline with ttf-parser before artwork reaches the scene. The resolved scene retains only vector paths plus `outlineFontId` provenance, and both Rust and the browser host require that ID to be present in the scene's verified font identities. Missing/wrong bindings, missing glyphs, malformed text metadata, unknown outline IDs, or unresolved `textSource` fail closed. The browser never uses `font-family`, `<text>`, Canvas measurement, or an implicit OS fallback.

## Clipboard

P09 structural clipboard carries only identity-free structure plus referenced immutable definitions. It does not carry live NodeIds, session handles or document identities. Pasting a custom symbol into a session that does not yet own the asset freezes the included definition first. Same command/id with different asset content is a conflict.

## Acceptance evidence

`custom_symbols_review.rs` covers resource freezing, validation, insertion, source edits, copy/duplicate/history, operator limits, vector scene validation, rich clipboard, ordered paint/erase/repaint and exact custom-text font binding. It verifies missing/wrong bindings fail, a correct registered-font binding becomes outline-only geometry, and forged unknown outline provenance is rejected.

`custom-symbol-parity-review.mjs` compares Native and actual WASM responses field-by-field for three representative symbol roles plus affine-stroke, layered-clip, ordered-erase and explicitly-bound text definitions, four math styles, two natural sizes, editing/history, rich clipboard and invalid resources. The clip gate requires exact transform nesting; the erase gate requires exact asset-local mask order; the text gate requires missing/wrong bindings to fail and the exact registered font SHA/face identity to produce vector-only scene geometry.

`custom-symbol-browser-review.mjs` uses an owned Edge instance with real pointer/keyboard/clipboard events, three formula hosts, 63 zoom/DPR artwork-pointer cases, structural insertion and cross-session frozen-asset transfer. It inspects the generated DOM for outer transform -> clip group -> inner matrix, verifies asset-local erase, and requires the resolved text artwork to reference a verified scene font while SVG/DOM remain path-only. It also forges cross-glyph erase re-entry, mismatched instances, unknown outline-font provenance and unresolved text metadata and requires the independent host validator to reject them.

No Office, production application, MathLive or MathJax path is exercised or enabled by this checkpoint.