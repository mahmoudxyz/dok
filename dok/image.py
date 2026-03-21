"""
dok.image
~~~~~~~~~
Minimal image dimension reading for PNG and JPEG.
No external dependencies.
"""

from __future__ import annotations
from pathlib import Path


def image_dimensions(path: str | Path) -> tuple[int, int]:
    """
    Read (width, height) in pixels from a PNG or JPEG file.
    Returns (0, 0) if the format is unrecognized.
    """
    with open(path, "rb") as f:
        header = f.read(32)

    if not header:
        return (0, 0)

    # PNG: signature + IHDR chunk
    if header[:8] == b"\x89PNG\r\n\x1a\n":
        w = int.from_bytes(header[16:20], "big")
        h = int.from_bytes(header[20:24], "big")
        return (w, h)

    # JPEG: scan for SOF0/SOF2 marker
    if header[:2] == b"\xff\xd8":
        return _jpeg_dimensions(path)

    return (0, 0)


def _jpeg_dimensions(path: str | Path) -> tuple[int, int]:
    with open(path, "rb") as f:
        f.read(2)  # skip SOI
        while True:
            b = f.read(1)
            if not b:
                break
            if b != b"\xff":
                continue
            marker = f.read(1)
            if not marker:
                break
            m = marker[0]
            # SOF0 or SOF2 (baseline / progressive)
            if m in (0xC0, 0xC2):
                f.read(3)  # length(2) + precision(1)
                h = int.from_bytes(f.read(2), "big")
                w = int.from_bytes(f.read(2), "big")
                return (w, h)
            # Skip other markers
            if 0xC0 <= m <= 0xFE and m not in (0xD8, 0xD9, 0x00):
                length = int.from_bytes(f.read(2), "big")
                f.read(length - 2)
    return (0, 0)


def image_content_type(filename: str) -> str:
    """Return the MIME content type for an image filename."""
    ext = Path(filename).suffix.lower()
    return {
        ".png":  "image/png",
        ".jpg":  "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif":  "image/gif",
        ".bmp":  "image/bmp",
        ".tiff": "image/tiff",
        ".tif":  "image/tiff",
    }.get(ext, "image/png")
