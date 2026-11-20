#!/usr/bin/env python3
"""Verify that a rendered macOS app icon remains visible at Dock-sized resolution."""

from __future__ import annotations

import argparse
import struct
import sys
import zlib
from pathlib import Path

PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"


def paeth(left: int, up: int, upper_left: int) -> int:
    estimate = left + up - upper_left
    left_distance = abs(estimate - left)
    up_distance = abs(estimate - up)
    upper_left_distance = abs(estimate - upper_left)
    if left_distance <= up_distance and left_distance <= upper_left_distance:
        return left
    if up_distance <= upper_left_distance:
        return up
    return upper_left


def decode_png(path: Path) -> tuple[int, int, int, bytes]:
    data = path.read_bytes()
    if not data.startswith(PNG_SIGNATURE):
        raise ValueError(f"{path} is not a PNG file")

    offset = len(PNG_SIGNATURE)
    width = height = bit_depth = color_type = interlace = None
    compressed = bytearray()
    while offset + 12 <= len(data):
        length = struct.unpack(">I", data[offset : offset + 4])[0]
        chunk_type = data[offset + 4 : offset + 8]
        payload_start = offset + 8
        payload_end = payload_start + length
        if payload_end + 4 > len(data):
            raise ValueError(f"{path} contains a truncated PNG chunk")
        payload = data[payload_start:payload_end]
        offset = payload_end + 4
        if chunk_type == b"IHDR":
            width, height, bit_depth, color_type, _, _, interlace = struct.unpack(
                ">IIBBBBB", payload
            )
        elif chunk_type == b"IDAT":
            compressed.extend(payload)
        elif chunk_type == b"IEND":
            break

    if None in (width, height, bit_depth, color_type, interlace):
        raise ValueError(f"{path} is missing a PNG header")
    if bit_depth != 8 or interlace != 0:
        raise ValueError("icon verification supports only non-interlaced 8-bit PNG files")

    channels = {0: 1, 2: 3, 4: 2, 6: 4}.get(color_type)
    if channels is None:
        raise ValueError(f"unsupported PNG color type {color_type}")
    stride = width * channels
    raw = zlib.decompress(bytes(compressed))
    expected = height * (stride + 1)
    if len(raw) != expected:
        raise ValueError(
            f"decoded PNG length is {len(raw)}, expected {expected} for {width}x{height}"
        )

    decoded = bytearray(height * stride)
    source_offset = 0
    for row_index in range(height):
        filter_type = raw[source_offset]
        source_offset += 1
        row = bytearray(raw[source_offset : source_offset + stride])
        source_offset += stride
        previous_offset = (row_index - 1) * stride
        for index in range(stride):
            left = row[index - channels] if index >= channels else 0
            up = decoded[previous_offset + index] if row_index > 0 else 0
            upper_left = (
                decoded[previous_offset + index - channels]
                if row_index > 0 and index >= channels
                else 0
            )
            if filter_type == 1:
                row[index] = (row[index] + left) & 0xFF
            elif filter_type == 2:
                row[index] = (row[index] + up) & 0xFF
            elif filter_type == 3:
                row[index] = (row[index] + ((left + up) // 2)) & 0xFF
            elif filter_type == 4:
                row[index] = (row[index] + paeth(left, up, upper_left)) & 0xFF
            elif filter_type != 0:
                raise ValueError(f"unsupported PNG filter {filter_type}")
        destination = row_index * stride
        decoded[destination : destination + stride] = row

    return width, height, color_type, bytes(decoded)


def alpha_values(width: int, height: int, color_type: int, pixels: bytes) -> list[int]:
    channels = {0: 1, 2: 3, 4: 2, 6: 4}[color_type]
    if color_type in (0, 2):
        return [255] * (width * height)
    alpha_index = channels - 1
    return [pixels[index + alpha_index] for index in range(0, len(pixels), channels)]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("png", type=Path)
    parser.add_argument("--minimum-visible-ratio", type=float, default=0.55)
    parser.add_argument("--alpha-threshold", type=int, default=32)
    parser.add_argument("--minimum-center-alpha", type=int, default=128)
    parser.add_argument(
        "--expected-subject-rgb",
        help="Required exact subject RGB as R,G,B, for example 31,99,142",
    )
    parser.add_argument("--minimum-subject-pixels", type=int, default=0)
    parser.add_argument(
        "--subject-rgb-tolerance",
        type=int,
        default=0,
        help="Maximum per-channel RGB difference allowed for anti-aliased subject pixels",
    )
    parser.add_argument("--minimum-white-ratio", type=float, default=0.0)
    args = parser.parse_args()

    width, height, color_type, pixels = decode_png(args.png)
    alphas = alpha_values(width, height, color_type, pixels)
    visible = sum(value >= args.alpha_threshold for value in alphas)
    visible_ratio = visible / len(alphas)
    center_index = (height // 2) * width + (width // 2)
    center_alpha = alphas[center_index]

    print(
        f"macOS icon {width}x{height}: visible_ratio={visible_ratio:.3f}, "
        f"center_alpha={center_alpha}"
    )
    if width < 32 or height < 32:
        print("Rendered app icon is smaller than 32x32.", file=sys.stderr)
        return 1
    if visible_ratio < args.minimum_visible_ratio:
        print(
            "Rendered app icon has too little visible area and may collapse to a tiny Dock mark.",
            file=sys.stderr,
        )
        return 1
    if center_alpha < args.minimum_center_alpha:
        print(
            "Rendered app icon is transparent at its center and may appear visually empty in the Dock.",
            file=sys.stderr,
        )
        return 1

    channels = {0: 1, 2: 3, 4: 2, 6: 4}[color_type]
    opaque_rgbs: list[tuple[int, int, int]] = []
    for index in range(0, len(pixels), channels):
        alpha = pixels[index + channels - 1] if color_type in (4, 6) else 255
        if alpha < args.alpha_threshold:
            continue
        if color_type in (2, 6):
            opaque_rgbs.append((pixels[index], pixels[index + 1], pixels[index + 2]))
        else:
            opaque_rgbs.append((pixels[index], pixels[index], pixels[index]))

    white_ratio = opaque_rgbs.count((255, 255, 255)) / len(opaque_rgbs)
    if white_ratio < args.minimum_white_ratio:
        print(
            f"Rendered app icon white ratio is {white_ratio:.3f}, below "
            f"{args.minimum_white_ratio:.3f}.",
            file=sys.stderr,
        )
        return 1

    if args.expected_subject_rgb:
        try:
            expected_subject = tuple(
                int(component) for component in args.expected_subject_rgb.split(",")
            )
        except ValueError:
            print("Expected subject RGB must use R,G,B integers.", file=sys.stderr)
            return 1
        if len(expected_subject) != 3 or any(
            component < 0 or component > 255 for component in expected_subject
        ):
            print("Expected subject RGB must contain three values from 0 to 255.", file=sys.stderr)
            return 1
        subject_pixels = sum(
            1
            for rgb in opaque_rgbs
            if max(
                abs(rgb[channel] - expected_subject[channel])
                for channel in range(3)
            )
            <= args.subject_rgb_tolerance
        )
        print(
            f"macOS icon colors: white_ratio={white_ratio:.3f}, "
            f"subject_rgb={expected_subject}, tolerance={args.subject_rgb_tolerance}, "
            f"subject_pixels={subject_pixels}"
        )
        if subject_pixels < args.minimum_subject_pixels:
            print(
                "Rendered app icon does not preserve enough pixels of the approved subject color.",
                file=sys.stderr,
            )
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
