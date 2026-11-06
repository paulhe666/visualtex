# Rare integral glyph source

`rareIntegralGlyphs.generated.ts` contains only the text- and display-style
outlines and OpenType MATH metrics needed for VisualTeX's rare integral
operators. They are extracted from STIX Two Math 2.13 b171 and converted to a
VisualTeX data registry; the original font file is not bundled or read at
runtime.

The generator maps STIX's 0.258em math axis to VisualTeX's 0.25em axis and
uniformly scales each outline and all horizontal/italic metrics. Text-style
glyphs use a 1.111em vertical advance and display-style glyphs use 2.222em,
matching the TeX Size1/Size2 integral geometry used by the existing renderer.

- Source: `STIXTwoMath-Regular.otf` from the upstream `v2.13b171` tag
- Upstream: https://github.com/stipub/stixfonts
- SHA-256: `3a5f3f26f40d5698b3c62dd085d48d6663696a3f80825aab8b553d5097518e8c`
- Copyright: Copyright 2001-2021 The STIX Fonts Project Authors
- License: SIL Open Font License 1.1; see `STIXTwoMath-OFL-1.1.txt`

Regenerate or verify the registry with:

```sh
python3 -m pip install fonttools==4.59.1
python3 scripts/generate_rare_integral_glyph_registry.py
python3 scripts/generate_rare_integral_glyph_registry.py --check
```

The generator downloads the pinned upstream file and rejects any source whose
SHA-256 differs. `--font /path/to/STIXTwoMath-Regular.otf` supports an offline
copy of that exact revision. The committed TypeScript output means application
rendering never depends on fonts installed on a user's system.
