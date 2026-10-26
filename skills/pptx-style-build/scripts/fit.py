#!/usr/bin/env python3
"""Text-fit measurement with real font metrics (stdlib TrueType/OpenType reader).

    python3 fit.py --font "Lato" --size 18 --box-pt 300x100 "some text"   # ad-hoc check
    python3 fit.py --self-test

Widths come from the font's hmtx advance widths (no kerning/shaping, so a 3% safety
margin is added). Fonts are found by family name in the OS font folders plus
$PPTX_STYLE_FONT_DIRS. If the theme font is missing, a metric-compatible twin is used
(Carlito for Calibri, Liberation for Arial/Times/Courier); failing that, an average
0.55 em glyph width - and the result is labelled approximate.
"""

from __future__ import annotations

import argparse
import os
import struct
import sys
from functools import lru_cache
from pathlib import Path

TWINS = {
    "calibri": "carlito",
    "cambria": "caladea",
    "arial": "liberation sans",
    "helvetica": "liberation sans",
    "times new roman": "liberation serif",
    "courier new": "liberation mono",
    "georgia": "gelasio",
    "arial narrow": "liberation sans narrow",
}
SAFETY = 1.03
AVG_EM = 0.55
EMU_PER_PT = 12700


def font_dirs() -> list[Path]:
    home = Path.home()
    dirs = [Path(p) for p in os.environ.get("PPTX_STYLE_FONT_DIRS", "").split(os.pathsep) if p]
    dirs += [
        Path("/System/Library/Fonts"),
        Path("/Library/Fonts"),
        home / "Library/Fonts",
        Path("/usr/share/fonts"),
        Path("/usr/local/share/fonts"),
        home / ".fonts",
        home / ".local/share/fonts",
        Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts",
    ]
    return [d for d in dirs if d.is_dir()]


class Font:
    """Advance widths of one TrueType/OpenType face (first face of a collection)."""

    def __init__(self, path: Path):
        self.path = path
        data = path.read_bytes()
        off = 0
        if data[:4] == b"ttcf":
            off = struct.unpack(">I", data[12:16])[0]
        n = struct.unpack(">H", data[off + 4 : off + 6])[0]
        self.tables = {}
        for i in range(n):
            rec = data[off + 12 + 16 * i : off + 28 + 16 * i]
            tag, _cs, toff, tlen = struct.unpack(">4sIII", rec)
            self.tables[tag.decode("latin-1")] = (toff, tlen)
        self.data = data
        self.upem = struct.unpack(">H", self._t("head")[18:20])[0]
        self.n_hm = struct.unpack(">H", self._t("hhea")[34:36])[0]
        self._cmap = None
        self._hmtx = None

    def _t(self, tag) -> bytes:
        o, n = self.tables[tag]
        return self.data[o : o + n]

    def names(self) -> tuple[set[str], str]:
        """({family names}, subfamily) from the name table (IDs 1/16 and 2/17)."""
        t = self._t("name")
        count, so = struct.unpack(">HH", t[2:6])
        fam, sub = set(), ""
        for i in range(count):
            pid, _eid, _lid, nid, ln, o = struct.unpack(">HHHHHH", t[6 + 12 * i : 18 + 12 * i])
            raw = t[so + o : so + o + ln]
            s = raw.decode("utf-16-be", "ignore") if pid in (0, 3) else raw.decode("latin-1")
            if nid in (1, 16):
                fam.add(s.strip().lower())
            elif nid in (2, 17) and (not sub or nid == 17):
                sub = s.strip().lower()
        return fam, sub

    def cmap(self) -> dict[int, int]:
        if self._cmap is not None:
            return self._cmap
        t = self._t("cmap")
        n = struct.unpack(">H", t[2:4])[0]
        best = None
        for i in range(n):
            pid, eid, o = struct.unpack(">HHI", t[4 + 8 * i : 12 + 8 * i])
            fmt = struct.unpack(">H", t[o : o + 2])[0]
            if fmt == 12 and (pid, eid) in ((3, 10), (0, 4), (0, 6)):
                best = (fmt, o)
                break
            if fmt == 4 and (pid, eid) in ((3, 1), (0, 3), (0, 1)) and best is None:
                best = (fmt, o)
        m: dict[int, int] = {}
        if best and best[0] == 4:
            o = best[1]
            segx2 = struct.unpack(">H", t[o + 6 : o + 8])[0]
            ends = struct.unpack(f">{segx2 // 2}H", t[o + 14 : o + 14 + segx2])
            base = o + 16 + segx2
            starts = struct.unpack(f">{segx2 // 2}H", t[base : base + segx2])
            deltas = struct.unpack(f">{segx2 // 2}h", t[base + segx2 : base + 2 * segx2])
            ro_base = base + 2 * segx2
            ros = struct.unpack(f">{segx2 // 2}H", t[ro_base : ro_base + segx2])
            for k in range(segx2 // 2):
                for c in range(starts[k], ends[k] + 1):
                    if c == 0xFFFF:
                        continue
                    if ros[k] == 0:
                        g = (c + deltas[k]) & 0xFFFF
                    else:
                        a = ro_base + 2 * k + ros[k] + 2 * (c - starts[k])
                        g = struct.unpack(">H", t[a : a + 2])[0]
                        g = (g + deltas[k]) & 0xFFFF if g else 0
                    m[c] = g
        elif best:
            o = best[1]
            ng = struct.unpack(">I", t[o + 12 : o + 16])[0]
            for k in range(ng):
                s, e, g = struct.unpack(">III", t[o + 16 + 12 * k : o + 28 + 12 * k])
                if e - s > 0x30000:
                    continue
                for c in range(s, e + 1):
                    m[c] = g + c - s
        self._cmap = m
        return m

    def advances(self) -> list[int]:
        if self._hmtx is None:
            t = self._t("hmtx")
            self._hmtx = [struct.unpack(">H", t[4 * i : 4 * i + 2])[0] for i in range(self.n_hm)]
        return self._hmtx

    def width_em(self, text: str) -> float:
        cm, adv = self.cmap(), self.advances()
        total = 0
        for ch in text:
            g = cm.get(ord(ch), 0)
            total += adv[min(g, len(adv) - 1)]
        return total / self.upem


@lru_cache(maxsize=1)
def font_index() -> dict[str, list[tuple[str, Path]]]:
    idx: dict[str, list[tuple[str, Path]]] = {}
    for d in font_dirs():
        for p in d.rglob("*"):
            if p.suffix.lower() not in (".ttf", ".otf", ".ttc"):
                continue
            try:
                fam, sub = Font(p).names()
            except (OSError, struct.error, KeyError, IndexError):
                continue
            for f in fam:
                idx.setdefault(f, []).append((sub, p))
    return idx


@lru_cache(maxsize=64)
def find_font(family: str, bold: bool = False):
    """(Font | None, how) where how is 'exact', 'twin:<name>' or 'missing'."""
    idx = font_index()
    for name, how in ((family.lower(), "exact"), (TWINS.get(family.lower(), ""), "twin")):
        faces = idx.get(name)
        if not name or not faces:
            continue
        want = "bold" if bold else "regular"
        faces = sorted(faces, key=lambda f: (f[0] != want, "italic" in f[0], len(f[0])))
        try:
            return Font(faces[0][1]), how if how == "exact" else f"twin:{name}"
        except (OSError, struct.error, KeyError):
            continue
    return None, "missing"


def measure(text: str, size_pt: float, family: str, bold: bool = False) -> tuple[float, str]:
    font, how = find_font(family, bold)
    if font is None:
        return len(text) * AVG_EM * size_pt * (1.05 if bold else 1.0), how
    return font.width_em(text) * size_pt * SAFETY, how


def wrap(text: str, width_pt: float, size_pt: float, family: str, bold: bool = False) -> int:
    """Lines needed for one paragraph with greedy word wrap (long words break)."""
    if not text:
        return 1
    lines, cur = 1, ""
    for word in text.split(" "):
        trial = (cur + " " + word) if cur else word
        if measure(trial, size_pt, family, bold)[0] <= width_pt:
            cur = trial
            continue
        if cur:
            lines += 1
        cur = word
        while measure(cur, size_pt, family, bold)[0] > width_pt and len(cur) > 1:
            # hard-break an over-long word
            k = len(cur)
            while k > 1 and measure(cur[:k], size_pt, family, bold)[0] > width_pt:
                k -= 1
            cur = cur[k:]
            lines += 1
    return lines


def check(
    paras,
    box_emu,
    size_pt,
    family,
    bold=False,
    ins=(91440, 45720, 91440, 45720),
    indent_emu=0,
    line_spacing=1.2,
    para_space_pt=0.0,
):
    """Fit verdict for paragraphs in a box: dict(fits, needed, available, how, chars_per_line)."""
    li, ti, ri, bi = ins
    width = max(box_emu[2] - li - ri - indent_emu, 1) / EMU_PER_PT
    height = max(box_emu[3] - ti - bi, 1) / EMU_PER_PT
    height -= para_space_pt * max(len(paras) - 1, 0)  # space before each later paragraph
    line_h = size_pt * line_spacing
    needed = sum(wrap(p, width, size_pt, family, bold) for p in paras) if paras else 0
    available = max(int(height // line_h), 1)
    how = find_font(family, bold)[1]
    avg = measure("abcdefghijklmnopqrstuvwxyz ABCDE", size_pt, family, bold)[0] / 32
    return {
        "fits": needed <= available,
        "needed": needed,
        "available": available,
        "how": how,
        "chars_per_line": int(width // max(avg, 0.1)),
    }


def self_test() -> int:
    here = Path(__file__).resolve()
    fonts = (
        list((here.parents[3] / "tests" / "fonts").glob("*.ttf")) if len(here.parents) > 3 else []
    )
    if fonts:
        os.environ["PPTX_STYLE_FONT_DIRS"] = str(fonts[0].parent)
        font_index.cache_clear()
        find_font.cache_clear()
        f = Font(fonts[0])
        fam, _sub = f.names()
        name = sorted(fam)[0]
        assert f.width_em("iiii") < f.width_em("WWWW")
        assert find_font(name)[1] == "exact"
    _, how = find_font("No Such Font 123")
    assert how == "missing"
    w, how = measure("abc", 10, "No Such Font 123")
    assert abs(w - 3 * AVG_EM * 10) < 1e-6 and how == "missing"
    assert wrap("one two three four five six", 40, 10, "No Such Font 123") > 1
    r = check(["x " * 400], [0, 0, 12700 * 200, 12700 * 40], 12, "No Such Font 123")
    assert not r["fits"] and r["needed"] > r["available"]
    print("self-test OK" + ("" if fonts else " (no bundled font found: metric path not exercised)"))
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("text", nargs="*")
    ap.add_argument("--font", default="Calibri")
    ap.add_argument("--size", type=float, default=18)
    ap.add_argument("--box-pt", default="400x200", help="WIDTHxHEIGHT in points")
    ap.add_argument("--self-test", action="store_true")
    a = ap.parse_args(argv)
    if a.self_test:
        return self_test()
    w, h = (float(v) for v in a.box_pt.split("x"))
    r = check(a.text or [""], [0, 0, int(w * EMU_PER_PT), int(h * EMU_PER_PT)], a.size, a.font)
    print(r)
    return 0 if r["fits"] else 1


if __name__ == "__main__":
    sys.exit(main())
