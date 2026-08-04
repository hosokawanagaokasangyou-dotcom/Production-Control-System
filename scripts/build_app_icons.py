#!/usr/bin/env python3
"""Generate PM-AI desktop and RDP launcher window icons (PNG sizes + ICO for jpackage)."""

from __future__ import annotations

import struct
import zlib
from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:  # pragma: no cover - Pillow optional at import, required for text icons
    Image = ImageDraw = ImageFont = None  # type: ignore[misc, assignment]

ROOT = Path(__file__).resolve().parents[1]
IMG_DIR = ROOT / "code_java" / "src" / "main" / "resources" / "jp" / "co" / "pm" / "ai" / "desktop" / "images"
BRAND_DIR = ROOT / "code_java" / "branding"

SIZES = (16, 32, 48, 64, 128, 256)

_FONT_CANDIDATES = (
    Path("C:/Windows/Fonts/YuGothB.ttc"),
    Path("C:/Windows/Fonts/YuGothM.ttc"),
    Path("C:/Windows/Fonts/meiryob.ttc"),
    Path("C:/Windows/Fonts/msgothic.ttc"),
    Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"),
    Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc"),
)


def _clamp(v: int) -> int:
    return max(0, min(255, v))


def _blend(bg: tuple[int, int, int, int], fg: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    br, bgc, bb, ba = bg
    fr, fgc, fb, fa = fg
    if fa <= 0:
        return bg
    if fa >= 255:
        return fg
    t = fa / 255.0
    return (
        _clamp(int(br * (1 - t) + fr * t)),
        _clamp(int(bgc * (1 - t) + fgc * t)),
        _clamp(int(bb * (1 - t) + fb * t)),
        _clamp(int(ba * (1 - t) + fa * t)),
    )


def _hex_rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)


def _rounded_rect_mask(size: int, radius: float) -> list[list[float]]:
    cx = cy = (size - 1) / 2.0
    half = size / 2.0
    r = radius
    mask: list[list[float]] = []
    for y in range(size):
        row: list[float] = []
        for x in range(size):
            dx = max(abs(x - cx) - half + r, 0.0)
            dy = max(abs(y - cy) - half + r, 0.0)
            dist = (dx * dx + dy * dy) ** 0.5
            if dist <= r:
                row.append(1.0)
            elif dist >= r + 1.0:
                row.append(0.0)
            else:
                row.append(1.0 - (dist - r))
        mask.append(row)
    return mask


def _fill_rounded_square(
    pixels: list[list[tuple[int, int, int, int]]],
    size: int,
    color: tuple[int, int, int, int],
    inset: float,
    radius: float,
) -> None:
    mask = _rounded_rect_mask(size, radius)
    for y in range(size):
        for x in range(size):
            if x < inset or y < inset or x >= size - inset or y >= size - inset:
                continue
            alpha = mask[y][x]
            if alpha <= 0:
                continue
            c = (color[0], color[1], color[2], _clamp(int(color[3] * alpha)))
            pixels[y][x] = _blend(pixels[y][x], c)


def _stroke_rounded_rect(
    pixels: list[list[tuple[int, int, int, int]]],
    size: int,
    color: tuple[int, int, int, int],
    inset: float,
    radius: float,
    width: float,
) -> None:
    outer = _rounded_rect_mask(size, radius)
    inner = _rounded_rect_mask(size, max(radius - width, 0.5))
    for y in range(size):
        for x in range(size):
            if x < inset or y < inset or x >= size - inset or y >= size - inset:
                continue
            edge = max(0.0, outer[y][x] - inner[y][x])
            if edge <= 0:
                continue
            c = (color[0], color[1], color[2], _clamp(int(color[3] * edge)))
            pixels[y][x] = _blend(pixels[y][x], c)


def _fill_rect(
    pixels: list[list[tuple[int, int, int, int]]],
    x0: float,
    y0: float,
    x1: float,
    y1: float,
    color: tuple[int, int, int, int],
    radius: float = 0.0,
) -> None:
    size = len(pixels)
    rx0, ry0, rx1, ry1 = int(x0), int(y0), int(x1), int(y1)
    for y in range(max(0, ry0), min(size, ry1)):
        for x in range(max(0, rx0), min(size, rx1)):
            if radius > 0:
                cx = x + 0.5
                cy = y + 0.5
                near_left = cx - x0
                near_right = x1 - cx
                near_top = cy - y0
                near_bottom = y1 - cy
                if near_left < radius and near_top < radius:
                    if (near_left - radius) ** 2 + (near_top - radius) ** 2 > radius * radius:
                        continue
                if near_right < radius and near_top < radius:
                    if (near_right - radius) ** 2 + (near_top - radius) ** 2 > radius * radius:
                        continue
                if near_left < radius and near_bottom < radius:
                    if (near_left - radius) ** 2 + (near_bottom - radius) ** 2 > radius * radius:
                        continue
                if near_right < radius and near_bottom < radius:
                    if (near_right - radius) ** 2 + (near_bottom - radius) ** 2 > radius * radius:
                        continue
            pixels[y][x] = _blend(pixels[y][x], color)


def _resolve_japanese_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    if ImageFont is None:
        raise RuntimeError("Pillow is required to render Japanese text icons (pip install Pillow)")
    for path in _FONT_CANDIDATES:
        if path.is_file():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def _blit_rgba(
    pixels: list[list[tuple[int, int, int, int]]],
    overlay: Image.Image,
    ox: int,
    oy: int,
) -> None:
    canvas_size = len(pixels)
    overlay = overlay.convert("RGBA")
    for y in range(overlay.height):
        dest_y = oy + y
        if dest_y < 0 or dest_y >= canvas_size:
            continue
        for x in range(overlay.width):
            dest_x = ox + x
            if dest_x < 0 or dest_x >= canvas_size:
                continue
            r, g, b, a = overlay.getpixel((x, y))
            if a <= 0:
                continue
            pixels[dest_y][dest_x] = _blend(pixels[dest_y][dest_x], (r, g, b, a))


def _draw_centered_text(
    pixels: list[list[tuple[int, int, int, int]]],
    size: int,
    text: str,
    color: tuple[int, int, int, int],
    *,
    center_y: float | None = None,
    font_scale: float = 0.38,
    shadow: tuple[int, int, int, int] | None = None,
) -> None:
    if Image is None or ImageDraw is None:
        raise RuntimeError("Pillow is required to render Japanese text icons (pip install Pillow)")
    font_px = max(8, int(size * font_scale))
    font = _resolve_japanese_font(font_px)
    scratch = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(scratch)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    cy = size / 2 if center_y is None else center_y
    tx = (size - tw) // 2 - bbox[0]
    ty = int(cy - th / 2) - bbox[1]
    if shadow is not None and size >= 24:
        offset = max(1, int(size * 0.008))
        draw.text((tx + offset, ty + offset), text, font=font, fill=shadow)
    draw.text((tx, ty), text, font=font, fill=color)
    _blit_rgba(pixels, scratch, 0, 0)


def _draw_corner_brackets(
    pixels: list[list[tuple[int, int, int, int]]],
    size: int,
    inset: float,
    arm: float,
    width: float,
    color: tuple[int, int, int, int],
) -> None:
    right = size - inset
    bottom = size - inset
    # 左上
    _fill_rect(pixels, inset, inset, inset + arm, inset + width, color, width * 0.35)
    _fill_rect(pixels, inset, inset, inset + width, inset + arm, color, width * 0.35)
    # 右上
    _fill_rect(pixels, right - arm, inset, right, inset + width, color, width * 0.35)
    _fill_rect(pixels, right - width, inset, right, inset + arm, color, width * 0.35)
    # 左下
    _fill_rect(pixels, inset, bottom - width, inset + arm, bottom, color, width * 0.35)
    _fill_rect(pixels, inset, bottom - arm, inset + width, bottom, color, width * 0.35)
    # 右下
    _fill_rect(pixels, right - arm, bottom - width, right, bottom, color, width * 0.35)
    _fill_rect(pixels, right - width, bottom - arm, right, bottom, color, width * 0.35)


def _draw_desktop_icon(size: int) -> list[list[tuple[int, int, int, int]]]:
    pixels = [[(0, 0, 0, 0) for _ in range(size)] for _ in range(size)]

    # 余白を最小化して描画領域を限界まで広げる
    outer_inset = max(size * 0.008, 0.4)
    outer_radius = size * 0.13
    border_w = max(size * 0.018, 1.0)

    white = (*_hex_rgb("#ffffff"), 255)
    sky_bg = (*_hex_rgb("#5cafff"), 255)
    sky_deep = (*_hex_rgb("#3b8eef"), 255)
    panel = (*_hex_rgb("#eef6ff"), 255)
    amber = (*_hex_rgb("#fbbf24"), 255)
    text_main = (*_hex_rgb("#1e40af"), 255)
    text_shadow = (*_hex_rgb("#93c5fd"), 120)

    _fill_rounded_square(pixels, size, white, outer_inset, outer_radius)

    inner_inset = outer_inset + border_w
    inner_radius = max(outer_radius - border_w * 0.5, size * 0.10)
    _fill_rounded_square(pixels, size, sky_bg, inner_inset, inner_radius)
    _fill_rounded_square(
        pixels,
        size,
        sky_deep,
        inner_inset + size * 0.03,
        max(inner_radius - size * 0.02, size * 0.08),
    )

    panel_inset_x = inner_inset + size * 0.015
    panel_inset_y = inner_inset + (size * 0.13 if size >= 32 else size * 0.08)
    panel_right = size - panel_inset_x
    panel_bottom = size - inner_inset - size * 0.015
    panel_radius = size * 0.05
    frame_w = max(size * 0.007, 0.7)
    _fill_rect(
        pixels,
        panel_inset_x - frame_w,
        panel_inset_y - frame_w,
        panel_right + frame_w,
        panel_bottom + frame_w,
        amber,
        panel_radius,
    )
    _fill_rect(
        pixels,
        panel_inset_x,
        panel_inset_y,
        panel_right,
        panel_bottom,
        panel,
        panel_radius,
    )

    if size >= 24:
        bar_h = max(size * 0.055, 1.2)
        bar_y = panel_inset_y + size * 0.03
        bar_left = panel_inset_x + size * 0.04
        bar_right = panel_right - size * 0.04
        bar_specs = [
            (0.88, _hex_rgb("#22c55e")),
            (0.72, _hex_rgb("#2563eb")),
            (0.58, _hex_rgb("#f97316")),
        ]
        gap = max(size * 0.012, 0.6)
        for i, (width_ratio, rgb) in enumerate(bar_specs):
            y0 = bar_y + i * (bar_h + gap)
            if y0 + bar_h > panel_inset_y + panel_bottom * 0.35:
                break
            w = (bar_right - bar_left) * width_ratio
            _fill_rect(
                pixels,
                bar_left,
                y0,
                bar_left + w,
                y0 + bar_h,
                (*rgb, 255),
                bar_h * 0.35,
            )

    if size >= 32:
        _draw_corner_brackets(
            pixels,
            size,
            inner_inset + size * 0.018,
            size * 0.085,
            max(size * 0.011, 0.9),
            (*_hex_rgb("#f59e0b"), 230),
        )

    text_y = (panel_inset_y + panel_bottom) / 2 + (size * 0.04 if size >= 32 else 0)
    if size >= 24:
        _draw_centered_text(
            pixels,
            size,
            "計画",
            text_main,
            center_y=text_y,
            font_scale=0.54 if size >= 48 else 0.50,
            shadow=text_shadow,
        )
    else:
        _draw_centered_text(
            pixels,
            size,
            "計",
            text_main,
            center_y=size * 0.52,
            font_scale=0.60,
        )
    return pixels


def _draw_monitor(
    pixels: list[list[tuple[int, int, int, int]]],
    size: int,
    cx: float,
    cy: float,
    w: float,
    h: float,
    bezel: tuple[int, int, int, int],
    screen: tuple[int, int, int, int],
) -> None:
    x0 = cx - w / 2
    y0 = cy - h / 2
    _fill_rect(pixels, x0, y0, x0 + w, y0 + h, bezel, w * 0.08)
    pad = w * 0.10
    _fill_rect(pixels, x0 + pad, y0 + pad, x0 + w - pad, y0 + h - pad, screen, w * 0.04)
    stand_w = w * 0.34
    stand_h = h * 0.12
    _fill_rect(
        pixels,
        cx - stand_w / 2,
        y0 + h,
        cx + stand_w / 2,
        y0 + h + stand_h,
        bezel,
        stand_h * 0.2,
    )


def _draw_rdp_icon(size: int) -> list[list[tuple[int, int, int, int]]]:
    pixels = [[(0, 0, 0, 0) for _ in range(size)] for _ in range(size)]
    inset = size * 0.06
    radius = size * 0.22
    bg = (*_hex_rgb("#0f1729"), 255)
    border = (*_hex_rgb("#22d3ee"), 190)
    _fill_rounded_square(pixels, size, bg, inset, radius)
    _stroke_rounded_rect(pixels, size, border, inset, radius, max(size * 0.018, 1.0))

    bezel = (*_hex_rgb("#334155"), 255)
    screen = (*_hex_rgb("#0ea5e9"), 255)
    _draw_monitor(pixels, size, size * 0.36, size * 0.50, size * 0.30, size * 0.22, bezel, screen)
    _draw_monitor(pixels, size, size * 0.66, size * 0.44, size * 0.22, size * 0.16, bezel, (*_hex_rgb("#38bdf8"), 255))

    link_color = (*_hex_rgb("#22d3ee"), 230)
    x0, y0 = size * 0.50, size * 0.46
    x1, y1 = size * 0.56, size * 0.42
    width = max(size * 0.03, 1.5)
    for t in range(101):
        t /= 100.0
        lx = x0 + (x1 - x0) * t
        ly = y0 + (y1 - y0) * t
        _fill_rect(pixels, lx - width, ly - width, lx + width, ly + width, link_color, width)

    node_r = size * 0.035
    nx, ny = size * 0.53, size * 0.44
    for y in range(size):
        for x in range(size):
            if (x + 0.5 - nx) ** 2 + (y + 0.5 - ny) ** 2 <= node_r * node_r:
                pixels[y][x] = _blend(pixels[y][x], (*_hex_rgb("#f8fafc"), 255))
    return pixels


def _png_chunk(tag: bytes, data: bytes) -> bytes:
    crc = zlib.crc32(tag + data) & 0xFFFFFFFF
    return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", crc)


def _write_png(path: Path, pixels: list[list[tuple[int, int, int, int]]]) -> None:
    height = len(pixels)
    width = len(pixels[0])
    raw = bytearray()
    for row in pixels:
        raw.append(0)
        for r, g, b, a in row:
            raw.extend((r, g, b, a))
    compressed = zlib.compress(bytes(raw), 9)
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    png = b"\x89PNG\r\n\x1a\n" + _png_chunk(b"IHDR", ihdr) + _png_chunk(b"IDAT", compressed) + _png_chunk(b"IEND", b"")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(png)


def _write_ico(path: Path, images: dict[int, list[list[tuple[int, int, int, int]]]]) -> None:
  entries = []
  image_data_parts = []
  offset = 6 + 16 * len(images)
  for size in sorted(images.keys()):
      pixels = images[size]
      png_path = path.with_suffix(f".{size}.tmp.png")
      _write_png(png_path, pixels)
      png_bytes = png_path.read_bytes()
      png_path.unlink(missing_ok=True)
      entries.append((size, len(png_bytes)))
      image_data_parts.append(png_bytes)

  out = bytearray()
  out.extend(struct.pack("<HHH", 0, 1, len(entries)))
  data_offset = offset
  for (size, length), data in zip(entries, image_data_parts):
      dim = 0 if size >= 256 else size
      out.extend(struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, 32, length, data_offset))
      data_offset += length
  for data in image_data_parts:
      out.extend(data)
  path.parent.mkdir(parents=True, exist_ok=True)
  path.write_bytes(out)


def _render_set(draw_fn, base_name: str) -> None:
    rendered: dict[int, list[list[tuple[int, int, int, int]]]] = {}
    for size in SIZES:
        rendered[size] = draw_fn(size)
        _write_png(IMG_DIR / f"{base_name}-{size}.png", rendered[size])
    _write_ico(BRAND_DIR / f"{base_name}.ico", rendered)
    _write_png(BRAND_DIR / f"{base_name}-256.png", rendered[256])


def main() -> None:
    _render_set(_draw_desktop_icon, "app-icon")
    _render_set(_draw_rdp_icon, "rdp-launcher-icon")
    print(f"Wrote icons under {IMG_DIR} and {BRAND_DIR}")


if __name__ == "__main__":
    main()
