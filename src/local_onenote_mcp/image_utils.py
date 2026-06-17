"""Small image helpers for OneNote page updates."""

from __future__ import annotations

from pathlib import Path


class ImageDimensionError(ValueError):
    """Raised when image dimensions cannot be read from a local file."""


def image_dimensions(path: str | Path) -> tuple[int, int]:
    """Return native image dimensions for common local image formats."""

    image_path = Path(path)
    with image_path.open("rb") as handle:
        header = handle.read(32)
        if header.startswith(b"\x89PNG\r\n\x1a\n"):
            return _png_dimensions(header)
        if header.startswith(b"\xff\xd8"):
            return _jpeg_dimensions(handle)
        if header.startswith((b"GIF87a", b"GIF89a")):
            return _gif_dimensions(header)
        if header.startswith(b"BM"):
            return _bmp_dimensions(header)
    raise ImageDimensionError(f"Unsupported image format for dimension inference: {image_path}")


def proportional_dimensions(
    path: str | Path,
    width: float | None,
    height: float | None,
) -> tuple[float | None, float | None]:
    """Resolve a partial requested size using the image's native aspect ratio."""

    if width is None and height is None:
        return None, None
    if width is not None and height is not None:
        return float(width), float(height)

    native_width, native_height = image_dimensions(path)
    if native_width <= 0 or native_height <= 0:
        raise ImageDimensionError(f"Invalid image dimensions: {path}")

    if width is not None:
        resolved_width = float(width)
        return resolved_width, resolved_width * native_height / native_width

    resolved_height = float(height)
    return resolved_height * native_width / native_height, resolved_height


def _png_dimensions(header: bytes) -> tuple[int, int]:
    if len(header) < 24 or header[12:16] != b"IHDR":
        raise ImageDimensionError("Invalid PNG header.")
    return int.from_bytes(header[16:20], "big"), int.from_bytes(header[20:24], "big")


def _gif_dimensions(header: bytes) -> tuple[int, int]:
    if len(header) < 10:
        raise ImageDimensionError("Invalid GIF header.")
    return int.from_bytes(header[6:8], "little"), int.from_bytes(header[8:10], "little")


def _bmp_dimensions(header: bytes) -> tuple[int, int]:
    if len(header) < 26:
        raise ImageDimensionError("Invalid BMP header.")
    width = int.from_bytes(header[18:22], "little", signed=True)
    height = int.from_bytes(header[22:26], "little", signed=True)
    return abs(width), abs(height)


def _jpeg_dimensions(handle) -> tuple[int, int]:
    sof_markers = {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}
    no_length_markers = set(range(0xD0, 0xD8)) | {0x01}

    while True:
        marker_prefix = handle.read(1)
        if marker_prefix == b"":
            break
        if marker_prefix != b"\xff":
            continue

        marker_byte = handle.read(1)
        while marker_byte == b"\xff":
            marker_byte = handle.read(1)
        if marker_byte == b"":
            break

        marker = marker_byte[0]
        if marker == 0xD9:
            break
        if marker in no_length_markers:
            continue

        segment_length_data = handle.read(2)
        if len(segment_length_data) != 2:
            break
        segment_length = int.from_bytes(segment_length_data, "big")
        if segment_length < 2:
            raise ImageDimensionError("Invalid JPEG segment length.")

        if marker in sof_markers:
            segment = handle.read(5)
            if len(segment) != 5:
                break
            height = int.from_bytes(segment[1:3], "big")
            width = int.from_bytes(segment[3:5], "big")
            return width, height

        handle.seek(segment_length - 2, 1)

    raise ImageDimensionError("Could not find JPEG dimensions.")
