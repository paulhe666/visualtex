#!/usr/bin/env python3
"""Generate deterministic VisualTeX integral outlines from the official esint10 Type 1 font.

The checked-in runtime payload contains only normalized SVG path data and font
metrics. The source PFB/TFM files are never bundled. By default the generator
locates TeX Live's official esint-type1/esint files through kpsewhich and writes
one TypeScript module to stdout so callers can review or redirect it explicitly.
"""

from __future__ import annotations

import argparse
import base64
import gzip
import hashlib
import json
import math
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

PFB_SHA256 = "3c2c4b9f98b9b741cf7e05155372c53b063fd96596205d5adfe2295ca9c9035e"
TFM_SHA256 = "fc941cd26d2b483f6cc9648d03d28dcc56e1c6621b6d9b6e11435a8cf2de7666"
UNITS_PER_EM = 1000


@dataclass(frozen=True)
class Target:
    command: str
    small_slot: int
    large_slot: int
    character: str
    aliases: tuple[str, ...] = ()


# Slot pairs follow esint.dtx 1.2d. Slots are written in octal in the package.
TARGETS = (
    Target("iiiint", 0o007, 0o010, "⨌"),
    Target("idotsint", 0o011, 0o012, "∫", ("dotsint",)),
    Target("sqint", 0o017, 0o020, "⨖"),
    Target("sqiint", 0o021, 0o022, "⨖"),
    Target("ointctrclockwise", 0o027, 0o030, "∳"),
    Target("ointclockwise", 0o031, 0o032, "∱"),
    Target("varointclockwise", 0o033, 0o034, "∲"),
    Target("varointctrclockwise", 0o035, 0o036, "∳"),
    Target("fint", 0o037, 0o040, "⨏"),
    Target("varoiint", 0o041, 0o042, "∯"),
    Target("landupint", 0o043, 0o044, "⨛"),
    Target("landdownint", 0o045, 0o046, "⨜"),
)

NUMBER = re.compile(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)")
ENCODING_ENTRY = re.compile(r"^dup\s+(\d+)/([^\s]+)\s+put$")
CHARSTRING_START = re.compile(r"^/([^\s]+)\s+\{$")
PL_CHARACTER = re.compile(
    r"\(CHARACTER\s+O\s+([0-7]+)(.*?)(?=\n\(CHARACTER|\Z)", re.DOTALL
)
PL_METRIC = re.compile(r"\((CHARWD|CHARHT|CHARDP|CHARIC)\s+R\s+([-+.\d]+)\)")


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def command_output(arguments: list[str]) -> str:
    completed = subprocess.run(
        arguments,
        check=True,
        capture_output=True,
        text=True,
    )
    return completed.stdout


def kpsewhich(name: str) -> Path:
    result = command_output(["kpsewhich", name]).strip()
    if not result:
        raise SystemExit(f"kpsewhich could not locate {name}")
    return Path(result)


def verify_source(path: Path, expected: str, description: str) -> None:
    actual = digest(path)
    if actual != expected:
        raise SystemExit(
            f"{description} checksum mismatch: expected {expected}, received {actual}"
        )


def fmt(value: float) -> str:
    if abs(value) < 1e-9:
        return "0"
    rounded = round(value)
    if abs(value - rounded) < 1e-9:
        return str(rounded)
    return f"{value:.6f}".rstrip("0").rstrip(".")


class Outline:
    def __init__(self) -> None:
        self.x = 0.0
        self.y = 0.0
        self.start_x = 0.0
        self.start_y = 0.0
        self.advance_width = 0.0
        self.left_side_bearing = 0.0
        self.commands: list[str] = []
        self.x_min = math.inf
        self.y_min = math.inf
        self.x_max = -math.inf
        self.y_max = -math.inf

    def include(self, x: float, y: float) -> None:
        self.x_min = min(self.x_min, x)
        self.y_min = min(self.y_min, y)
        self.x_max = max(self.x_max, x)
        self.y_max = max(self.y_max, y)

    def move_to(self, x: float, y: float) -> None:
        self.x = x
        self.y = y
        self.start_x = x
        self.start_y = y
        self.commands.append(f"M{fmt(x)} {fmt(y)}")

    def line_to(self, x: float, y: float) -> None:
        self.include(self.x, self.y)
        self.include(x, y)
        self.x = x
        self.y = y
        self.commands.append(f"L{fmt(x)} {fmt(y)}")

    @staticmethod
    def cubic_value(p0: float, p1: float, p2: float, p3: float, t: float) -> float:
        one_minus = 1 - t
        return (
            one_minus**3 * p0
            + 3 * one_minus**2 * t * p1
            + 3 * one_minus * t**2 * p2
            + t**3 * p3
        )

    @staticmethod
    def extrema(p0: float, p1: float, p2: float, p3: float) -> Iterable[float]:
        a = -p0 + 3 * p1 - 3 * p2 + p3
        b = 2 * (p0 - 2 * p1 + p2)
        c = p1 - p0
        if abs(a) < 1e-12:
            if abs(b) > 1e-12:
                t = -c / b
                if 0 < t < 1:
                    yield t
            return
        discriminant = b * b - 4 * a * c
        if discriminant < 0:
            return
        root = math.sqrt(discriminant)
        for t in ((-b + root) / (2 * a), (-b - root) / (2 * a)):
            if 0 < t < 1:
                yield t

    def curve_to(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        x3: float,
        y3: float,
    ) -> None:
        x0, y0 = self.x, self.y
        self.include(x0, y0)
        self.include(x3, y3)
        for t in self.extrema(x0, x1, x2, x3):
            self.include(self.cubic_value(x0, x1, x2, x3, t), self.cubic_value(y0, y1, y2, y3, t))
        for t in self.extrema(y0, y1, y2, y3):
            self.include(self.cubic_value(x0, x1, x2, x3, t), self.cubic_value(y0, y1, y2, y3, t))
        self.x = x3
        self.y = y3
        self.commands.append(
            f"C{fmt(x1)} {fmt(y1)} {fmt(x2)} {fmt(y2)} {fmt(x3)} {fmt(y3)}"
        )

    def close(self) -> None:
        self.include(self.x, self.y)
        self.include(self.start_x, self.start_y)
        # Type 1 CharString closepath differs from PostScript/SVG closepath:
        # it closes the current subpath without moving the CharString current
        # point. A following relative moveto is therefore measured from the
        # point that was current before closepath, not from the subpath start.
        # Keep the parser current point unchanged and emit an SVG Z only for
        # geometry; subsequent Type 1 relative moves are converted to absolute
        # SVG M coordinates by move_to().
        self.commands.append("Z")

    def payload(self, italic_correction: float) -> dict[str, object]:
        if not self.commands or not math.isfinite(self.x_min):
            raise ValueError("glyph outline is empty")
        return {
            "path": "".join(self.commands),
            "mathJaxPath": "".join(self.commands)[1:-1],
            "advanceWidth": round(self.advance_width),
            "leftSideBearing": round(self.left_side_bearing),
            "italicCorrection": round(italic_correction),
            "height": max(0, math.ceil(self.y_max)),
            "depth": max(0, math.ceil(-self.y_min)),
            "bounds": {
                "xMin": math.floor(self.x_min),
                "xMax": math.ceil(self.x_max),
                "yMin": math.floor(self.y_min),
                "yMax": math.ceil(self.y_max),
            },
        }


def consume(stack: list[float], count: int, operator: str) -> list[float]:
    if len(stack) != count:
        raise ValueError(f"{operator} expected {count} operands, received {stack}")
    values = stack[:]
    stack.clear()
    return values


def consume_multiple(stack: list[float], stride: int, operator: str) -> list[list[float]]:
    if not stack or len(stack) % stride:
        raise ValueError(f"{operator} expected a multiple of {stride} operands, received {stack}")
    values = [stack[index : index + stride] for index in range(0, len(stack), stride)]
    stack.clear()
    return values


def parse_charstring(body: list[str]) -> Outline:
    outline = Outline()
    stack: list[float] = []
    ignored = {"hstem", "vstem", "hstem3", "vstem3", "dotsection"}
    flex_start: tuple[float, float] | None = None
    flex_points: list[tuple[float, float]] = []

    for raw_line in body:
        line = raw_line.split("%", 1)[0].strip()
        if not line:
            continue
        for token in line.split():
            if NUMBER.fullmatch(token):
                stack.append(float(token))
                continue
            if token in {"{", "}", "ND"}:
                continue
            if token == "div":
                if len(stack) < 2:
                    raise ValueError("div requires two operands")
                divisor = stack.pop()
                dividend = stack.pop()
                stack.append(dividend / divisor)
                continue
            if token in ignored:
                stack.clear()
                continue
            if token == "hsbw":
                side_bearing, width = consume(stack, 2, token)
                outline.left_side_bearing = side_bearing
                outline.advance_width = width
                outline.x = side_bearing
                outline.y = 0
                continue
            if token == "sbw":
                side_x, side_y, width, _width_y = consume(stack, 4, token)
                outline.left_side_bearing = side_x
                outline.advance_width = width
                outline.x = side_x
                outline.y = side_y
                continue
            if token == "rmoveto":
                dx, dy = consume(stack, 2, token)
                if flex_start is not None:
                    outline.x += dx
                    outline.y += dy
                else:
                    outline.move_to(outline.x + dx, outline.y + dy)
                continue
            if token == "hmoveto":
                (dx,) = consume(stack, 1, token)
                outline.move_to(outline.x + dx, outline.y)
                continue
            if token == "vmoveto":
                (dy,) = consume(stack, 1, token)
                outline.move_to(outline.x, outline.y + dy)
                continue
            if token == "rlineto":
                for dx, dy in consume_multiple(stack, 2, token):
                    outline.line_to(outline.x + dx, outline.y + dy)
                continue
            if token in {"hlineto", "vlineto"}:
                values = stack[:]
                stack.clear()
                horizontal = token == "hlineto"
                for delta in values:
                    if horizontal:
                        outline.line_to(outline.x + delta, outline.y)
                    else:
                        outline.line_to(outline.x, outline.y + delta)
                    horizontal = not horizontal
                continue
            if token == "rrcurveto":
                for dx1, dy1, dx2, dy2, dx3, dy3 in consume_multiple(stack, 6, token):
                    x0, y0 = outline.x, outline.y
                    outline.curve_to(
                        x0 + dx1,
                        y0 + dy1,
                        x0 + dx1 + dx2,
                        y0 + dy1 + dy2,
                        x0 + dx1 + dx2 + dx3,
                        y0 + dy1 + dy2 + dy3,
                    )
                continue
            if token in {"hvcurveto", "vhcurveto"}:
                groups = consume_multiple(stack, 4, token)
                horizontal_first = token == "hvcurveto"
                for first, second, third, fourth in groups:
                    x0, y0 = outline.x, outline.y
                    if horizontal_first:
                        outline.curve_to(
                            x0 + first,
                            y0,
                            x0 + first + second,
                            y0 + third,
                            x0 + first + second,
                            y0 + third + fourth,
                        )
                    else:
                        outline.curve_to(
                            x0,
                            y0 + first,
                            x0 + second,
                            y0 + first + third,
                            x0 + second + fourth,
                            y0 + first + third,
                        )
                    horizontal_first = not horizontal_first
                continue
            if token == "closepath":
                if stack:
                    raise ValueError(f"closepath received unexpected operands {stack}")
                outline.close()
                continue
            if token == "callsubr":
                if not stack:
                    raise ValueError("callsubr requires a subroutine index")
                subroutine = round(stack.pop())
                if subroutine == 1:
                    # Adobe Type 1 flex protocol: subr 1 begins collecting the
                    # reference point plus six Bézier points.
                    if stack:
                        raise ValueError(f"flex start received unexpected operands {stack}")
                    flex_start = (outline.x, outline.y)
                    flex_points = []
                    continue
                if subroutine == 2:
                    if flex_start is None:
                        raise ValueError("flex point appeared outside a flex sequence")
                    if stack:
                        raise ValueError(f"flex point received unexpected operands {stack}")
                    flex_points.append((outline.x, outline.y))
                    continue
                if subroutine == 0:
                    if flex_start is None or len(flex_points) != 7:
                        raise ValueError(
                            f"flex end expected seven recorded points, received {flex_points}"
                        )
                    # The first recorded point is the flex reference point used
                    # by rasterizers to decide between curves and straight
                    # segments. The remaining six points are the two cubics.
                    stack.clear()
                    outline.x, outline.y = flex_start
                    outline.curve_to(*flex_points[1], *flex_points[2], *flex_points[3])
                    outline.curve_to(*flex_points[4], *flex_points[5], *flex_points[6])
                    flex_start = None
                    flex_points = []
                    continue
                if subroutine == 3:
                    stack.clear()
                    continue
                raise ValueError(f"unsupported Type 1 subroutine {subroutine}")
            if token == "setcurrentpoint":
                x, y = consume(stack, 2, token)
                outline.x = x
                outline.y = y
                continue
            if token == "endchar":
                stack.clear()
                continue
            if token in {"return", "readonly", "def"}:
                continue
            raise ValueError(f"unsupported Type 1 charstring operator {token!r}")

    return outline


def parse_type1(disassembly: str) -> tuple[dict[int, str], dict[str, list[str]], str]:
    encoding: dict[int, str] = {}
    charstrings: dict[str, list[str]] = {}
    version_match = re.search(r"^/version\s+\(([^)]+)\)", disassembly, re.MULTILINE)
    version = version_match.group(1) if version_match else "unknown"
    lines = disassembly.splitlines()
    index = 0
    while index < len(lines):
        stripped = lines[index].strip()
        encoding_match = ENCODING_ENTRY.match(stripped)
        if encoding_match:
            encoding[int(encoding_match.group(1))] = encoding_match.group(2)
            index += 1
            continue
        start = CHARSTRING_START.match(stripped)
        if not start:
            index += 1
            continue
        name = start.group(1)
        body: list[str] = []
        index += 1
        while index < len(lines) and lines[index].strip() != "} ND":
            body.append(lines[index])
            index += 1
        if index >= len(lines):
            raise ValueError(f"unterminated Type 1 charstring {name}")
        charstrings[name] = body
        index += 1
    return encoding, charstrings, version


def parse_tfm(pl: str) -> dict[int, dict[str, float]]:
    result: dict[int, dict[str, float]] = {}
    for match in PL_CHARACTER.finditer(pl):
        slot = int(match.group(1), 8)
        metrics = {
            key: float(value) * UNITS_PER_EM
            for key, value in PL_METRIC.findall(match.group(2))
        }
        result[slot] = metrics
    return result


def glyph_payload(
    slot: int,
    encoding: dict[int, str],
    charstrings: dict[str, list[str]],
    tfm: dict[int, dict[str, float]],
) -> dict[str, object]:
    glyph_name = encoding.get(slot)
    if not glyph_name:
        raise ValueError(f"Type 1 encoding has no slot {slot} (octal {slot:o})")
    body = charstrings.get(glyph_name)
    if body is None:
        raise ValueError(f"Type 1 font has no charstring /{glyph_name}")
    metrics = tfm.get(slot)
    if metrics is None:
        raise ValueError(f"TFM has no slot {slot} (octal {slot:o})")
    outline = parse_charstring(body)
    expected_width = metrics.get("CHARWD", 0)
    if abs(outline.advance_width - expected_width) > 1.1:
        raise ValueError(
            f"slot {slot:o} width mismatch: Type1={outline.advance_width}, TFM={expected_width}"
        )
    result = outline.payload(metrics.get("CHARIC", 0))
    result["glyphName"] = glyph_name
    result["slot"] = slot
    result["tfmDepth"] = round(metrics.get("CHARDP", 0))
    result["tfmHeight"] = round(metrics.get("CHARHT", 0))
    return result


def generate(pfb_path: Path, tfm_path: Path) -> dict[str, object]:
    verify_source(pfb_path, PFB_SHA256, "esint10.pfb")
    verify_source(tfm_path, TFM_SHA256, "esint10.tfm")
    disassembly = command_output(["t1disasm", str(pfb_path)])
    pl = command_output(["tftopl", str(tfm_path)])
    encoding, charstrings, version = parse_type1(disassembly)
    tfm = parse_tfm(pl)
    glyphs = []
    for target in TARGETS:
        glyphs.append(
            {
                "command": target.command,
                "aliases": list(target.aliases),
                "character": target.character,
                "small": glyph_payload(target.small_slot, encoding, charstrings, tfm),
                "large": glyph_payload(target.large_slot, encoding, charstrings, tfm),
            }
        )
    return {
        "source": {
            "family": "esint10",
            "package": "esint-type1",
            "version": version,
            "pfbSha256": PFB_SHA256,
            "tfmSha256": TFM_SHA256,
            "license": "Public Domain",
        },
        "unitsPerEm": UNITS_PER_EM,
        "glyphs": glyphs,
    }


def typescript(payload: dict[str, object]) -> str:
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode()
    compressed = gzip.compress(raw, compresslevel=9, mtime=0)
    encoded = base64.b64encode(compressed).decode("ascii")
    checksum = hashlib.sha256(raw).hexdigest()
    return "\n".join(
        [
            "/**",
            " * Generated by scripts/generate_esint_glyph_registry.py.",
            " * Contains normalized SVG outlines derived from the public-domain esint10 Type 1 font.",
            " * Do not hand-edit and do not bundle the source font file.",
            " */",
            f'export const ESINT_GLYPHS_JSON_SHA256 = "{checksum}";',
            f'export const ESINT_GLYPHS_GZIP_BASE64 = "{encoded}";',
            "",
        ]
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pfb", type=Path)
    parser.add_argument("--tfm", type=Path)
    parser.add_argument("--json", action="store_true")
    arguments = parser.parse_args()
    pfb = arguments.pfb or kpsewhich("esint10.pfb")
    tfm = arguments.tfm or kpsewhich("esint10.tfm")
    payload = generate(pfb, tfm)
    if arguments.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True))
    else:
        print(typescript(payload), end="")


if __name__ == "__main__":
    main()
