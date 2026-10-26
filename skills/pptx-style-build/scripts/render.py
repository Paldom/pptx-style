#!/usr/bin/env python3
"""Render a deck to PNGs plus one contact sheet for visual review (LibreOffice + poppler).

    python3 render.py DECK.pptx [--out DIR] [--contract contract.json] [--require]

Writes DIR/<deck>.pdf, DIR/slide-NN.png and DIR/contact-sheet.png (stdlib PNG tiling).
The render is labelled APPROXIMATE when a theme font is not installed: LibreOffice then
substitutes another font and line breaks differ from PowerPoint.
Exit codes: 0 rendered (or skipped: LibreOffice absent), 1 render failed,
2 usage error, 3 skipped while --require was given.
"""

from __future__ import annotations

import argparse
import json
import struct
import sys
import zlib
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fit
import pptx_core as core


def read_png(path: Path):
    """(width, height, rows of RGB bytes) for 8-bit RGB/RGBA non-interlaced PNGs."""
    data = path.read_bytes()
    pos, idat = 8, b""
    w = h = ctype = 0
    while pos < len(data):
        n = struct.unpack(">I", data[pos : pos + 4])[0]
        kind = data[pos + 4 : pos + 8]
        body = data[pos + 8 : pos + 8 + n]
        if kind == b"IHDR":
            w, h, depth, ctype, _c, _f, interlace = struct.unpack(">IIBBBBB", body)
            if depth != 8 or ctype not in (2, 6) or interlace:
                raise ValueError(f"{path.name}: unsupported PNG format")
        elif kind == b"IDAT":
            idat += body
        pos += 12 + n
    bpp = 3 if ctype == 2 else 4
    raw = zlib.decompress(idat)
    stride = w * bpp
    rows, prev = [], bytearray(stride)
    i = 0
    for _ in range(h):
        f = raw[i]
        line = bytearray(raw[i + 1 : i + 1 + stride])
        i += 1 + stride
        for x in range(stride):
            a = line[x - bpp] if x >= bpp else 0
            b = prev[x]
            c = prev[x - bpp] if x >= bpp else 0
            if f == 1:
                line[x] = (line[x] + a) & 255
            elif f == 2:
                line[x] = (line[x] + b) & 255
            elif f == 3:
                line[x] = (line[x] + (a + b) // 2) & 255
            elif f == 4:
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                line[x] = (line[x] + (a if pa <= pb and pa <= pc else b if pb <= pc else c)) & 255
        prev = line
        rows.append(bytes(line) if bpp == 3 else bytes(v for k, v in enumerate(line) if k % 4 != 3))
    return w, h, rows


def write_png(path: Path, w: int, h: int, rows) -> None:
    raw = b"".join(b"\x00" + r for r in rows)

    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d))

    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw, 6))
        + chunk(b"IEND", b"")
    )


def contact_sheet(pngs: list[Path], out: Path, cols: int = 4, tile_w: int = 320) -> None:
    """Tile slide PNGs (nearest-neighbour scaled) on a grey sheet with 8 px gutters."""
    tiles = []
    for p in pngs:
        w, h, rows = read_png(p)
        th = max(1, round(h * tile_w / w))
        scaled = []
        for y in range(th):
            src = rows[min(h - 1, y * h // th)]
            scaled.append(
                b"".join(
                    src[(x * w // tile_w) * 3 : (x * w // tile_w) * 3 + 3] for x in range(tile_w)
                )
            )
        tiles.append(scaled)
    gut = 8
    th = max(len(t) for t in tiles)
    nrows = (len(tiles) + cols - 1) // cols
    W = cols * tile_w + (cols + 1) * gut
    H = nrows * th + (nrows + 1) * gut
    grey = b"\x9a\x9a\x9a"
    sheet = [bytearray(grey * W) for _ in range(H)]
    for i, t in enumerate(tiles):
        r, c = divmod(i, cols)
        x0, y0 = gut + c * (tile_w + gut), gut + r * (th + gut)
        for y, line in enumerate(t):
            sheet[y0 + y][x0 * 3 : (x0 + tile_w) * 3] = line
    write_png(out, W, H, [bytes(r) for r in sheet])


def font_status(contract: dict | None) -> tuple[bool, list[str]]:
    if not contract:
        return False, []
    fonts = {f for f in contract["theme"]["fonts"].values() if f}
    missing = [f for f in sorted(fonts) if fit.find_font(f)[1] != "exact"]
    return bool(missing), missing


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("deck", nargs="?")
    ap.add_argument("--out", help="output folder (default <deck>-render/)")
    ap.add_argument("--contract", help="contract.json, to check theme-font availability")
    ap.add_argument(
        "--require", action="store_true", help="exit 3 instead of skipping without LibreOffice"
    )
    ap.add_argument("--self-test", action="store_true")
    a = ap.parse_args(argv)
    if a.self_test:
        import tempfile

        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "a.png"
            write_png(p, 3, 2, [b"\xff\x00\x00" * 3, b"\x00\x00\xff" * 3])
            w, h, rows = read_png(p)
            assert (w, h) == (3, 2) and rows[1][:3] == b"\x00\x00\xff"
            contact_sheet([p, p], Path(d) / "s.png", cols=2, tile_w=6)
            assert read_png(Path(d) / "s.png")[0] == 2 * 6 + 3 * 8
        print("self-test OK")
        return 0
    if not a.deck:
        ap.print_usage()
        return 2
    deck = Path(a.deck)
    out = Path(a.out) if a.out else deck.with_name(deck.stem + "-render")
    contract = core.load_contract(a.contract) if a.contract else None
    approx, missing = font_status(contract)
    r = core.render(deck, out)
    if r["status"] == "skipped":
        print(
            f"SKIPPED: {r['reason']} - render not done; lint results stand, visual review is still owed"
        )
        return 3 if a.require else 0
    if r["status"] != "ok":
        print(f"render failed: {r['reason']}", file=sys.stderr)
        return 1
    pngs = [Path(p) for p in r["pngs"]]
    sheet = None
    if pngs:
        sheet = out / "contact-sheet.png"
        try:
            contact_sheet(pngs, sheet)
        except ValueError as exc:
            print(f"contact sheet skipped: {exc}")
            sheet = None
    label = (
        f"APPROXIMATE (theme font not installed: {', '.join(missing)})" if approx else "fonts OK"
    )
    meta = {
        "pdf": r["pdf"],
        "pngs": r["pngs"],
        "contact_sheet": str(sheet) if sheet else None,
        "label": label,
    }
    (out / "render.json").write_text(json.dumps(meta, indent=1), encoding="utf-8")
    print(f"rendered {len(pngs)} slides -> {out} [{label}]")
    if sheet:
        print(f"contact sheet: {sheet}")
    if r.get("reason"):
        print(r["reason"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
