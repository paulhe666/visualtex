#!/usr/bin/env python3
"""Generate the bundled rare-integral SVG glyph registry from STIX Two Math.

The generated TypeScript is a runtime asset: VisualTeX never reads an
installed system font.  This generator pins and verifies the exact OFL-1.1
font revision so a regeneration is deterministic on every platform.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import urllib.request
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any

try:
    from fontTools.misc.roundTools import otRound
    from fontTools.pens.boundsPen import BoundsPen
    from fontTools.pens.svgPathPen import SVGPathPen
    from fontTools.pens.transformPen import TransformPen
    from fontTools.ttLib import TTFont
except ImportError as error:  # pragma: no cover - exercised by humans/CI setup
    raise SystemExit(
        "fontTools is required: python3 -m pip install fonttools==4.59.1"
    ) from error


STIX_VERSION = "2.13 b171"
STIX_TAG = "v2.13b171"
STIX_FONT_URL = (
    "https://raw.githubusercontent.com/stipub/stixfonts/"
    f"{STIX_TAG}/fonts/static_otf/STIXTwoMath-Regular.otf"
)
STIX_FONT_SHA256 = "3a5f3f26f40d5698b3c62dd085d48d6663696a3f80825aab8b553d5097518e8c"
STIX_SOURCE_AXIS_HEIGHT = 258
TARGET_AXIS_HEIGHT = 250
TARGET_VERTICAL_ADVANCES = {"small": 1111, "large": 2222}


@dataclass(frozen=True)
class IntegralTarget:
    command: str
    code_point: int
    aliases: tuple[str, ...] = ()


TARGETS = (
    IntegralTarget("intclockwise", 0x2231),
    IntegralTarget("varointclockwise", 0x2232),
    IntegralTarget("ointctrclockwise", 0x2233),
    IntegralTarget("sumint", 0x2A0B),
    IntegralTarget("iiiint", 0x2A0C),
    IntegralTarget("intbar", 0x2A0D),
    IntegralTarget("intBar", 0x2A0E),
    IntegralTarget("fint", 0x2A0F),
    IntegralTarget("cirfnint", 0x2A10),
    IntegralTarget("awint", 0x2A11, ("intctrclockwise",)),
    IntegralTarget("rppolint", 0x2A12),
    IntegralTarget("scpolint", 0x2A13),
    IntegralTarget("npolint", 0x2A14),
    IntegralTarget("pointint", 0x2A15),
    IntegralTarget("quatint", 0x2A16),
    IntegralTarget("intlarhk", 0x2A17),
    IntegralTarget("intx", 0x2A18),
    IntegralTarget("intcap", 0x2A19),
    IntegralTarget("intcup", 0x2A1A),
    IntegralTarget("upint", 0x2A1B),
    IntegralTarget("lowint", 0x2A1C),
)


def read_source(font_path: Path | None) -> bytes:
    if font_path is not None:
        data = font_path.read_bytes()
    else:
        request = urllib.request.Request(
            STIX_FONT_URL, headers={"User-Agent": "VisualTeX-glyph-generator"}
        )
        with urllib.request.urlopen(request, timeout=30) as response:
            data = response.read()
    digest = hashlib.sha256(data).hexdigest()
    if digest != STIX_FONT_SHA256:
        raise SystemExit(
            "STIX Two Math source checksum mismatch: "
            f"expected {STIX_FONT_SHA256}, received {digest}"
        )
    return data


def name_record(font: TTFont, name_id: int) -> str:
    value = font["name"].getDebugName(name_id)
    if value is None:
        raise ValueError(f"STIX font has no name record {name_id}")
    return value


def integer_string(value: float) -> str:
    return str(otRound(value))


def extract_variant(
    font: TTFont,
    glyph_name: str,
    source_vertical_advance: int,
    target_vertical_advance: int,
    source_axis_height: int,
    italic_corrections: dict[str, int],
) -> dict[str, Any]:
    glyph_set = font.getGlyphSet()
    glyph = glyph_set[glyph_name]
    scale = target_vertical_advance / source_vertical_advance
    transform = (
        scale,
        0,
        0,
        scale,
        0,
        TARGET_AXIS_HEIGHT - source_axis_height * scale,
    )

    bounds_pen = BoundsPen(glyph_set)
    glyph.draw(TransformPen(bounds_pen, transform))
    if bounds_pen.bounds is None:
        raise ValueError(f"Glyph {glyph_name} has no outline bounds")
    x_min, y_min, x_max, y_max = (
        otRound(value) for value in bounds_pen.bounds
    )

    svg_pen = SVGPathPen(glyph_set, ntos=integer_string)
    glyph.draw(TransformPen(svg_pen, transform))
    path = svg_pen.getCommands()
    if not path:
        raise ValueError(f"Glyph {glyph_name} has an empty SVG path")
    if not path.startswith("M") or not path.endswith("Z"):
        raise ValueError(f"Glyph {glyph_name} has a non-canonical SVG path")

    advance_width, left_side_bearing = font["hmtx"][glyph_name]
    return {
        "glyphName": glyph_name,
        "path": path,
        # MathJax wraps SVG font data as `M${p}Z`.
        "mathJaxPath": path[1:-1],
        "advanceWidth": otRound(advance_width * scale),
        "leftSideBearing": otRound(left_side_bearing * scale),
        "italicCorrection": otRound(
            italic_corrections.get(glyph_name, 0) * scale
        ),
        "height": max(0, y_max),
        "depth": max(0, -y_min),
        "verticalAdvance": target_vertical_advance,
        "sourceVerticalAdvance": source_vertical_advance,
        "scale": scale,
        "bounds": {
            "xMin": x_min,
            "xMax": x_max,
            "yMin": y_min,
            "yMax": y_max,
        },
    }


def extract_registry(data: bytes) -> tuple[int, list[dict[str, Any]]]:
    font = TTFont(BytesIO(data), recalcBBoxes=False, recalcTimestamp=False)
    if name_record(font, 1) != "STIX Two Math":
        raise ValueError("Pinned font family is not STIX Two Math")
    if STIX_VERSION not in name_record(font, 5):
        raise ValueError(f"Pinned font is not STIX Two Math {STIX_VERSION}")

    units_per_em = font["head"].unitsPerEm
    if units_per_em != 1000:
        raise ValueError(f"Unexpected STIX unitsPerEm: {units_per_em}")

    cmap = font.getBestCmap()
    math_table = font["MATH"].table
    source_axis_height = math_table.MathConstants.AxisHeight.Value
    if source_axis_height != STIX_SOURCE_AXIS_HEIGHT:
        raise ValueError(f"Unexpected STIX math axis: {source_axis_height}")
    variants = math_table.MathVariants
    if len(variants.VertGlyphCoverage.glyphs) != len(
        variants.VertGlyphConstruction
    ):
        raise ValueError("STIX MATH vertical variant coverage is inconsistent")
    vertical_constructions = dict(
        zip(variants.VertGlyphCoverage.glyphs, variants.VertGlyphConstruction)
    )
    italics_info = math_table.MathGlyphInfo.MathItalicsCorrectionInfo
    if len(italics_info.Coverage.glyphs) != len(italics_info.ItalicsCorrection):
        raise ValueError("STIX MATH italic-correction coverage is inconsistent")
    italic_corrections = dict(
        zip(
            italics_info.Coverage.glyphs,
            (record.Value for record in italics_info.ItalicsCorrection),
        )
    )

    result: list[dict[str, Any]] = []
    for target in TARGETS:
        glyph_name = cmap.get(target.code_point)
        if glyph_name is None:
            raise ValueError(f"STIX cmap is missing U+{target.code_point:04X}")
        construction = vertical_constructions.get(glyph_name)
        if construction is None:
            raise ValueError(f"STIX MATH table has no variants for {glyph_name}")
        records = construction.MathGlyphVariantRecord
        if len(records) != 2:
            raise ValueError(
                f"Expected exactly text/display variants for {glyph_name}, got {len(records)}"
            )
        if records[0].VariantGlyph != glyph_name:
            raise ValueError(f"First MATH variant for {glyph_name} is not the base glyph")
        if records[1].VariantGlyph != f"{glyph_name}.dsp":
            raise ValueError(f"Display MATH variant for {glyph_name} is not .dsp")

        result.append(
            {
                "command": target.command,
                "aliases": list(target.aliases),
                "character": chr(target.code_point),
                "codePoint": target.code_point,
                "small": extract_variant(
                    font,
                    records[0].VariantGlyph,
                    records[0].AdvanceMeasurement,
                    TARGET_VERTICAL_ADVANCES["small"],
                    source_axis_height,
                    italic_corrections,
                ),
                "large": extract_variant(
                    font,
                    records[1].VariantGlyph,
                    records[1].AdvanceMeasurement,
                    TARGET_VERTICAL_ADVANCES["large"],
                    source_axis_height,
                    italic_corrections,
                ),
            }
        )
    return units_per_em, result


def ts_string(value: str) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def render_variant(variant: dict[str, Any], indent: str) -> list[str]:
    bounds = variant["bounds"]
    return [
        f"{indent}{{",
        f"{indent}  glyphName: {ts_string(variant['glyphName'])},",
        f"{indent}  path: {ts_string(variant['path'])},",
        f"{indent}  mathJaxPath: {ts_string(variant['mathJaxPath'])},",
        f"{indent}  advanceWidth: {variant['advanceWidth']},",
        f"{indent}  leftSideBearing: {variant['leftSideBearing']},",
        f"{indent}  italicCorrection: {variant['italicCorrection']},",
        f"{indent}  height: {variant['height']},",
        f"{indent}  depth: {variant['depth']},",
        f"{indent}  verticalAdvance: {variant['verticalAdvance']},",
        f"{indent}  sourceVerticalAdvance: {variant['sourceVerticalAdvance']},",
        f"{indent}  scale: {variant['scale']:.12f},",
        (
            f"{indent}  bounds: {{ xMin: {bounds['xMin']}, xMax: {bounds['xMax']}, "
            f"yMin: {bounds['yMin']}, yMax: {bounds['yMax']} }},"
        ),
        f"{indent}}}",
    ]


def render_typescript(units_per_em: int, registry: list[dict[str, Any]]) -> str:
    lines = [
        "/* This file is generated by scripts/generate_rare_integral_glyph_registry.py.",
        " * Do not hand-edit it. Outlines are derived from STIX Two Math 2.13 b171",
        " * under the SIL Open Font License 1.1; see STIXTwoMath-OFL-1.1.txt.",
        " * Geometry is normalized to the TeX Size1/Size2 advances around the",
        " * VisualTeX 0.25em math axis.",
        " */",
        "",
        "export interface RareIntegralGlyphBounds {",
        "  readonly xMin: number;",
        "  readonly xMax: number;",
        "  readonly yMin: number;",
        "  readonly yMax: number;",
        "}",
        "",
        "export interface RareIntegralGlyphVariant {",
        "  /** SVG path in the font's y-up coordinate system. */",
        "  readonly path: string;",
        "  /** MathJax SVG font payload (the outer M/Z are supplied by MathJax). */",
        "  readonly mathJaxPath: string;",
        "  readonly glyphName: string;",
        "  readonly advanceWidth: number;",
        "  readonly leftSideBearing: number;",
        "  readonly italicCorrection: number;",
        "  readonly height: number;",
        "  readonly depth: number;",
        "  readonly verticalAdvance: number;",
        "  readonly sourceVerticalAdvance: number;",
        "  readonly scale: number;",
        "  readonly bounds: RareIntegralGlyphBounds;",
        "}",
        "",
        "export interface RareIntegralGlyphDefinition {",
        "  readonly command: string;",
        "  readonly aliases: readonly string[];",
        "  readonly character: string;",
        "  readonly codePoint: number;",
        "  readonly small: RareIntegralGlyphVariant;",
        "  readonly large: RareIntegralGlyphVariant;",
        "}",
        "",
        f"export const RARE_INTEGRAL_GLYPH_UNITS_PER_EM = {units_per_em};",
        f"export const RARE_INTEGRAL_GLYPH_AXIS_HEIGHT = {TARGET_AXIS_HEIGHT};",
        "export const RARE_INTEGRAL_GLYPH_SOURCE = Object.freeze({",
        f"  family: {ts_string('STIX Two Math')},",
        f"  version: {ts_string(STIX_VERSION)},",
        f"  tag: {ts_string(STIX_TAG)},",
        f"  sha256: {ts_string(STIX_FONT_SHA256)},",
        f"  url: {ts_string(STIX_FONT_URL)},",
        "  license: \"SIL Open Font License 1.1\",",
        f"  sourceAxisHeight: {STIX_SOURCE_AXIS_HEIGHT},",
        f"  targetAxisHeight: {TARGET_AXIS_HEIGHT},",
        "});",
        "",
        "export const RARE_INTEGRAL_GLYPHS = [",
    ]
    for entry in registry:
        lines.extend(
            [
                "  {",
                f"    command: {ts_string(entry['command'])},",
                f"    aliases: {json.dumps(entry['aliases'], ensure_ascii=False)},",
                f"    character: {ts_string(entry['character'])},",
                f"    codePoint: 0x{entry['codePoint']:04x},",
                "    small:",
            ]
        )
        small = render_variant(entry["small"], "      ")
        small[-1] += ","
        lines.extend(small)
        lines.append("    large:")
        large = render_variant(entry["large"], "      ")
        large[-1] += ","
        lines.extend(large)
        lines.append("  },")
    lines.extend(
        [
            "] as const satisfies readonly RareIntegralGlyphDefinition[];",
            "",
            "export const RARE_INTEGRAL_GLYPHS_BY_COMMAND: Readonly<",
            "  Record<string, RareIntegralGlyphDefinition>",
            "> = Object.freeze(",
            "  Object.fromEntries(",
            "    RARE_INTEGRAL_GLYPHS.flatMap((glyph) =>",
            "      [glyph.command, ...glyph.aliases].map((command) => [command, glyph]),",
            "    ),",
            "  ),",
            ");",
            "",
        ]
    )
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    repo_default = Path(__file__).resolve().parents[1] / "src/math/rareIntegralGlyphs.generated.ts"
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--font",
        type=Path,
        help="Use a local copy of the pinned STIXTwoMath-Regular.otf",
    )
    parser.add_argument("--output", type=Path, default=repo_default)
    parser.add_argument(
        "--check",
        action="store_true",
        help="Fail if the committed registry differs; do not write",
    )
    parser.add_argument(
        "--stdout", action="store_true", help="Print generated TypeScript; do not write"
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    units_per_em, registry = extract_registry(read_source(args.font))
    generated = render_typescript(units_per_em, registry)
    if args.stdout:
        sys.stdout.write(generated)
        return 0
    if args.check:
        try:
            existing = args.output.read_text(encoding="utf-8")
        except FileNotFoundError:
            print(f"Missing generated registry: {args.output}", file=sys.stderr)
            return 1
        if existing != generated:
            print(
                f"Generated registry is stale: run {Path(__file__).name}",
                file=sys.stderr,
            )
            return 1
        print(f"Rare-integral registry is current: {args.output}")
        return 0
    args.output.write_text(generated, encoding="utf-8")
    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
