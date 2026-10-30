#!/usr/bin/env python3
"""
Generate SlideKick task-pane icons (16/32/80 px).

Primary source: src/static/assets/logo_corelan.png (the Corelan diamond
badge, 200x200 RGBA). It is resampled with macOS `sips`, which preserves the
alpha channel.

Fallback (no logo file / non-macOS): regenerates the original amber-bolt
badge with the standard library only.

Usage: python3 tools/make_icons.py [output_dir]
"""

import struct
import subprocess
import sys
import zlib
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SIZES = (16, 32, 80)
LOGO = REPO / "src" / "static" / "assets" / "logo_corelan.png"

BG = (15, 23, 42)        # slate-900 (bolt fallback background)
BOLT = (245, 158, 11)    # amber-500 (bolt fallback foreground)

# Lightning bolt polygon, normalized to a 0..1 box (bolt fallback only).
BOLT_POLY = [
    (0.50, 0.06), (0.16, 0.56), (0.40, 0.56),
    (0.30, 0.94), (0.84, 0.42), (0.54, 0.42), (0.66, 0.06),
]


def build_from_logo(outdir):
    for size in SIZES:
        out = outdir / ("icon%d.png" % size)
        r = subprocess.run(
            ["sips", "-z", str(size), str(size), str(LOGO), "--out", str(out)],
            capture_output=True, text=True,
        )
        if r.returncode != 0:
            print("sips failed for %d px: %s" % (size, r.stderr.strip()))
            return False
        raw = out.read_bytes()
        w, h = struct.unpack(">II", raw[16:24])
        if (w, h) != (size, size) or raw[25] != 6:  # color type 6 = RGBA
            print("unexpected output for %d px" % size)
            return False
        print("wrote", out, "(%dx%d, %d bytes)" % (w, h, len(raw)))
    return True


def inside_rounded(px, py, size, radius):
    if px < 0 or py < 0 or px > size or py > size:
        return False
    for cx, cy in ((radius, radius), (size - radius, radius),
                   (radius, size - radius), (size - radius, size - radius)):
        in_corner_zone = (
            (px < radius and py < radius) or
            (px > size - radius and py < radius) or
            (px < radius and py > size - radius) or
            (px > size - radius and py > size - radius)
        )
        if in_corner_zone:
            if (px - cx) ** 2 + (py - cy) ** 2 > radius * radius:
                return False
            break
    return True


def inside_polygon(px, py, poly):
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


def render_bolt(size, ss=4):
    S = size * ss
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
                color = BOLT if inside_polygon(sx, sy, poly) else BG
                acc[int(sy)][int(sx)][0] += color[0]
                acc[int(sy)][int(sx)][1] += color[1]
                acc[int(sy)][int(sx)][2] += color[2]
                acc[int(sy)][int(sx)][3] += 255
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
                row.append((round(c[0] / cell), round(c[1] / cell),
                            round(c[2] / cell), round(c[3] / cell)))
        out.append(row)
    return out


def write_png(path, pixels):
    h = len(pixels)
    w = len(pixels[0])
    raw = b""
    for row in pixels:
        raw += b"\x00" + b"".join(struct.pack("4B", *px) for px in row)

    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data +
                struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    png = (b"\x89PNG\r\n\x1a\n"
           + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(raw, 9))
           + chunk(b"IEND", b""))
    Path(path).write_bytes(png)


def build_fallback(outdir):
    for size in SIZES:
        path = outdir / ("icon%d.png" % size)
        write_png(path, render_bolt(size, ss=4))
        print("wrote", path)


def main():
    outdir = Path(sys.argv[1]) if len(sys.argv) > 1 else REPO / "src" / "static" / "assets"
    outdir.mkdir(parents=True, exist_ok=True)
    if LOGO.exists() and sys.platform == "darwin":
        if build_from_logo(outdir):
            return
        print("logo build failed - falling back to bolt badge")
    build_fallback(outdir)


if __name__ == "__main__":
    main()