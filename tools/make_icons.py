#!/usr/bin/env python3
"""
Generate SlideKick task-pane icons (16/32/80 px) with the standard library only:
an amber lightning bolt on a dark rounded square, with alpha transparency.

Usage: python3 tools/make_icons.py [output_dir]
"""

import struct
import sys
import zlib
from pathlib import Path

BG = (15, 23, 42)        # slate-900
BOLT = (245, 158, 11)    # amber-500

# Lightning bolt polygon, normalized to a 0..1 box.
BOLT_POLY = [
    (0.50, 0.06), (0.16, 0.56), (0.40, 0.56),
    (0.30, 0.94), (0.84, 0.42), (0.54, 0.42), (0.66, 0.06),
]


def inside_rounded(px, py, size, radius):
    """True if (px,py) is inside the rounded square of edge `size`."""
    x0, y0, x1, y1 = 0, 0, size, size
    if px < x0 or px > x1 or py < y0 or py > y1:
        return False
    # corner centers
    for cx, cy in ((x0 + radius, y0 + radius), (x1 - radius, y0 + radius),
                   (x0 + radius, y1 - radius), (x1 - radius, y1 - radius)):
        in_corner_zone = (
            (px < x0 + radius and py < y0 + radius) or
            (px > x1 - radius and py < y0 + radius) or
            (px < x0 + radius and py > y1 - radius) or
            (px > x1 - radius and py > y1 - radius)
        )
        if in_corner_zone:
            if (px - cx) ** 2 + (py - cy) ** 2 > radius * radius:
                return False
            break
    return True


def inside_polygon(px, py, poly):
    """Even-odd scanline point-in-polygon."""
    inside = False
    n = len(poly)
    j = n - 1
    for i in range(n):
        xi, yi = poly[i]
        xj, yj = poly[j]
        if (yi > py) != (yj > py):
            xint = xi + (py - yi) * (xj - xi) / (yj - yi)
            if px < xint:
                inside = not inside
        j = i
    return inside


def render(size, ss=4):
    """Render at supersample factor ss, return 2D array of (r,g,b,a)."""
    S = size * ss
    scale = S / 1.0
    margin = S * 0.04
    box_size = S - 2 * margin
    radius = box_size * 0.20
    poly = [(margin + x * box_size, margin + y * box_size) for x, y in BOLT_POLY]

    acc = [[[0, 0, 0, 0] for _ in range(size)] for _ in range(size)]
    grid = [[0] * size for _ in range(size)]

    for y in range(S):
        sy = (y + 0.5) / ss
        for x in range(S):
            sx = (x + 0.5) / ss
            if inside_rounded(sx, sy, S, radius):
                if inside_polygon(sx, sy, poly):
                    color = BOLT + (255,)
                else:
                    color = BG + (255,)
                acc[int(sy)][int(sx)][0] += color[0]
                acc[int(sy)][int(sx)][1] += color[1]
                acc[int(sy)][int(sx)][2] += color[2]
                acc[int(sy)][int(sx)][3] += color[3]
                grid[int(sy)][int(sx)] += 1

    out = []
    for y in range(size):
        row = []
        for x in range(size):
            cell = grid[y][x]
            if cell == 0:
                row.append((0, 0, 0, 0))
            else:
                c = acc[y][x]
                row.append((
                    round(c[0] / cell), round(c[1] / cell),
                    round(c[2] / cell), round(c[3] / cell),
                ))
        out.append(row)
    return out


def write_png(path, pixels):
    h = len(pixels)
    w = len(pixels[0])
    raw = b""
    for row in pixels:
        raw += b"\x00" + b"".join(
            struct.pack("4B", *px) for px in row
        )
    def chunk(tag, data):
        c = struct.pack(">I", len(data)) + tag + data
        return c + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(raw, 9))
    png += chunk(b"IEND", b"")
    Path(path).write_bytes(png)


def main():
    outdir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "src" / "static" / "assets"
    outdir.mkdir(parents=True, exist_ok=True)
    for size in (16, 32, 80):
        pixels = render(size, ss=4)
        target = outdir / ("icon%d.png" % size)
        write_png(target, pixels)
        print("wrote", target)


if __name__ == "__main__":
    main()