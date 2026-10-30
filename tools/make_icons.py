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
LOGO = REPO / "slidekick" / "static" / "assets" / "logo_corelan.png"

BG = (15, 23, 42)        # slate-900 (bolt fallback background)
BOLT = (245, 158, 11)    # amber-500 (bolt fallback foreground)

# Lightning bolt polygon, normalized to a 0..1 box (bolt fallback only).
BOLT_POLY = [
    (0.50, 0.06), (0.16, 0.56), (0.40, 0.56),
    (0.30, 0.94), (0.84, 0.42), (0.54, 0.42), (0.66, 0.06),
]


def _alpha_bbox(png_path):
    """Exact bbox of non-transparent pixels via a minimal PNG RGBA8 decoder."""
    import zlib as _z
    raw = png_path.read_bytes()
    pos = 8
    w = h = 0
    idat = b""
    while pos < len(raw):
        ln = struct.unpack(">I", raw[pos:pos + 4])[0]
        tag = raw[pos + 4:pos + 8]
        data = raw[pos + 8:pos + 8 + ln]
        if tag == b"IHDR":
            w, h = struct.unpack(">II", data[0:8])
        elif tag == b"IDAT":
            idat += data
        elif tag == b"IEND":
            break
        pos += 12 + ln
    data = _z.decompress(idat)
    stride = w * 4 + 1
    rows = []
    prev = bytearray(w * 4)
    for i in range(h):
        f = data[i * stride]
        line = bytearray(data[i * stride + 1:(i + 1) * stride])
        if f == 0:
            pass
        elif f == 1:
            for x in range(4, len(line)):
                line[x] = (line[x] + line[x - 4]) & 0xFF
        elif f == 2:
            for x in range(len(line)):
                line[x] = (line[x] + prev[x]) & 0xFF
        elif f == 3:
            for x in range(len(line)):
                a = line[x - 4] if x >= 4 else 0
                line[x] = (line[x] + ((a + prev[x]) >> 1)) & 0xFF
        elif f == 4:
            for x in range(len(line)):
                a = line[x - 4] if x >= 4 else 0
                b = prev[x]
                c = prev[x - 4] if x >= 4 else 0
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                pr = a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)
                line[x] = (line[x] + pr) & 0xFF
        prev = line
        rows.append(line)
    minx, miny, maxx, maxy = w, h, -1, -1
    for y in range(h):
        row = rows[y]
        for x in range(w):
            if row[x * 4 + 3] > 50:
                if x < minx: minx = x
                if x > maxx: maxx = x
                if y < miny: miny = y
                if y > maxy: maxy = y
    if maxx < 0:
        return None
    return (minx, miny, maxx, maxy)


def build_from_logo(outdir):
    # App-icon tile style: crop tight to the badge (drop transparent margin and
    # the micro-text ring), scale to ~80% of the tile, then pad onto a slate
    # tile. The bright diamond carries the contrast, so the tile reads on both
    # light and dark Add-ins menus (a transparent-bg logo reads as a black
    # blob on dark Office surfaces).
    TILE = (34, 39, 56)  # slate tile (matches the logo's dark ring family)
    for size in SIZES:
        out = outdir / ("icon%d.png" % size)
        inner = max(12, int(size * 0.80))
        bbox = _alpha_bbox(LOGO)
        if bbox:
            x0, y0, x1, y1 = bbox
            bw, bh = x1 - x0 + 1, y1 - y0 + 1
            side = max(bw, bh)  # square around the diamond
            cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
            off_x = max(0, min(200 - side, cx - side // 2))
            off_y = max(0, min(200 - side, cy - side // 2))
        else:
            side, off_x, off_y = 140, 30, 30
        tmp = outdir / "_crop_tmp.png"
        r1 = subprocess.run(["sips", "-c", str(side), str(side),
                             "--cropOffset", str(off_y), str(off_x), str(LOGO),
                             "--out", str(tmp)], capture_output=True, text=True)
        r2 = subprocess.run(["sips", "-z", str(inner), str(inner), str(tmp),
                             "--out", str(tmp)], capture_output=True, text=True)
        r3 = subprocess.run(["sips", "--padToHeightWidth", str(size), str(size),
                             "--padColor", "%02X%02X%02X" % TILE, str(tmp),
                             "--out", str(out)], capture_output=True, text=True)
        for r in (r1, r2, r3):
            if r.returncode != 0:
                print("sips failed: %s" % r.stderr.strip())
                return False
        tmp.unlink()
        _verify(out, size)
    return True


def _verify(out, size):
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