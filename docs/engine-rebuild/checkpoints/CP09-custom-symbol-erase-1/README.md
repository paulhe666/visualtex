# CP09 custom-symbol erase 1

Stable non-Git checkpoint for P09 custom-symbol ordered erase support, building on `CP09-custom-symbol-clip-2`.

Stable source SHA256: `9183bed5169375d5ff895621edbb182cb9eb045812e860fcece84d471feaf68e`.

Verification receipt: `engine/target/p09-verification/2026-09-14T07-51-39-543Z/report.json`.

The complete 14-step P09 gate passed with `sourceUnchanged=true`. The custom-symbol Rust review passes 15/15. Native/WASM custom-symbol parity passes with 57 requests, 18 scenes, maximum numeric difference 0 and exact SVG parity. Owned Edge acceptance passes with six frozen definitions, 12 artwork paths, 54 real zoom/DPR pointer cases, trusted keyboard/clipboard events and independent host validation.

Legacy `operation:"erase"` is implemented as ordered, asset-local vector compositing. Each CustomSymbol node owns one contiguous artwork group. Paint layers contribute white and erase layers black to an engine-generated luminance mask; later paint can restore local ink. The visible vector paint is then masked only inside that CustomSymbol instance. Neighbor formula glyphs and rules never enter the mask.

Rust scene validation and the independent browser host reject an artwork group that re-enters after a glyph/rule, non-monotonic path indices, or mismatched x/y/scale/color inside one immutable asset instance. Dedicated tests forge cross-glyph paint order and mismatched erase instances and require rejection. This prevents erase geometry from affecting neighboring formula content without flattening authored paths into approximate destructive boolean geometry.

The remaining advanced custom-artwork gap is `text`: browser/system-font fallback remains prohibited. Future text support must bind a real validated font outline resource.

No production application or Office/VSTO path was switched. No user Word document was touched. P00 protection remains red because of independent external Office/VSTO/evidence work. No commit or push belongs to this checkpoint.
