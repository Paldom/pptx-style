"""Shared OOXML core for pptx-style-profile and pptx-style-build (stdlib only).

This file is copied byte-identically into both skills' scripts/ folders; a repo test
keeps the copies in sync. It holds the package model (zip parts, relationships,
content types), namespace-preserving XML I/O, template extraction into a style
contract, and the contract diff. Nothing here writes outside the paths it is given.
"""

from __future__ import annotations

import copy
import hashlib
import io
import json
import os
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

VERSION = "0.1.0"
SCHEMA = "pptx-style-contract/1"

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT_PRES = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
CT_TMPL = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
CT_NOTES = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
SECTION_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
XML_HEADER = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
EMU_PER_PT = 12700
SHAPE_TAGS = {"sp", "grpSp", "graphicFrame", "pic", "cxnSp", "contentPart", "AlternateContent"}
FOOTER_PH = {"dt", "ftr", "sldNum"}
TITLE_PH = {"title", "ctrTitle"}
PROMPT_CONTENT_RE = re.compile(r"^\s*[\d.\-/:]{1,6}\s*$")
LEFTOVER_RE = re.compile(r"lorem|ipsum|\bxxx+\b|\bTODO\b|\[insert|click to (add|edit)", re.I)


def q(tag: str) -> str:
    """'p:sp' -> '{uri}sp' (Clark notation)."""
    pfx, local = tag.split(":")
    return f"{{{NS[pfx]}}}{local}"


def local(el) -> str:
    return el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""


# --- namespace-preserving XML I/O --------------------------------------------
# ElementTree renames prefixes (ns0:) and drops declarations it does not need, which
# breaks mc:Ignorable="p14 a16" in PowerPoint. We record every declaration on parse,
# register the same prefixes before serializing, and re-inject any declaration the
# serializer dropped onto the root element.


def parse_xml(data: bytes):
    decls: list[tuple[str, str]] = []
    for _ev, (pfx, uri) in ET.iterparse(io.BytesIO(data), events=("start-ns",)):  # noqa: S314
        if (pfx, uri) not in decls:
            decls.append((pfx, uri))
    return ET.fromstring(data), decls  # noqa: S314


def dump_xml(root, decls) -> bytes:
    default = None
    for pfx, uri in decls:
        if pfx:
            ET.register_namespace(pfx, uri)
        else:
            default = uri
    if default:  # ET's default_namespace= rejects unqualified attributes, so use a stand-in prefix
        ET.register_namespace("dflt0", default)
    text = ET.tostring(root, encoding="unicode")
    if default:
        text = re.sub(r"(</?)dflt0:", r"\1", text).replace("xmlns:dflt0=", "xmlns=")
    end = text.index(">")
    if text[end - 1] == "/":
        end -= 1
    head = text[:end]
    extra = "".join(
        f' xmlns:{p}="{u}"' for p, u in decls if p and f"xmlns:{p}=" not in head and p != "xml"
    )
    return XML_HEADER + (head + extra + text[end:]).encode("utf-8")


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# --- package ------------------------------------------------------------------


class Package:
    """A .pptx/.potx as an ordered dict of parts; XML parts parsed lazily."""

    def __init__(self, path):
        self.path = Path(path)
        self.parts: dict[str, bytes] = {}
        if not zipfile.is_zipfile(self.path):
            raise ValueError(
                f"{self.path.name}: not an OOXML package (binary .ppt, password-protected or corrupt);"
                " save it as .pptx/.potx first"
            )
        with zipfile.ZipFile(self.path) as z:
            for info in z.infolist():
                if not info.is_dir():
                    self.parts[info.filename] = z.read(info)
        self._xml: dict[str, tuple] = {}
        self.dirty: set[str] = set()

    # raw/xml access
    def xml(self, name: str):
        if name not in self._xml:
            self._xml[name] = parse_xml(self.parts[name])
        return self._xml[name][0]

    def touch(self, name: str) -> None:
        self.dirty.add(name)

    def put_xml(self, name: str, root, decls=None) -> None:
        if decls is None:
            decls = [(p, u) for p, u in NS.items() if p in ("a", "p", "r")]
        self._xml[name] = (root, decls)
        self.parts[name] = b""
        self.dirty.add(name)

    def put_bytes(self, name: str, data: bytes) -> None:
        self._xml.pop(name, None)
        self.dirty.discard(name)
        self.parts[name] = data

    def delete(self, name: str) -> None:
        for n in (name, rels_name(name)):
            self.parts.pop(n, None)
            self._xml.pop(n, None)
            self.dirty.discard(n)
        ct = self.xml("[Content_Types].xml")
        for o in ct.findall(q("ct:Override")):
            if o.get("PartName") == "/" + name:
                ct.remove(o)
                self.touch("[Content_Types].xml")

    def flush(self) -> None:
        for name in sorted(self.dirty):
            root, decls = self._xml[name]
            self.parts[name] = dump_xml(root, decls)
        self.dirty.clear()

    def save(self, path) -> None:
        self.flush()
        order = ["[Content_Types].xml"] + [n for n in self.parts if n != "[Content_Types].xml"]
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in order:
                z.writestr(n, self.parts[n])

    # content types
    def content_type(self, name: str) -> str | None:
        ct = self.xml("[Content_Types].xml")
        for o in ct.findall(q("ct:Override")):
            if o.get("PartName") == "/" + name:
                return o.get("ContentType")
        ext = name.rsplit(".", 1)[-1].lower()
        for d in ct.findall(q("ct:Default")):
            if d.get("Extension", "").lower() == ext:
                return d.get("ContentType")
        return None

    def set_override(self, name: str, ctype: str) -> None:
        ct = self.xml("[Content_Types].xml")
        for o in ct.findall(q("ct:Override")):
            if o.get("PartName") == "/" + name:
                o.set("ContentType", ctype)
                break
        else:
            ET.SubElement(ct, q("ct:Override"), PartName="/" + name, ContentType=ctype)
        self.touch("[Content_Types].xml")

    def ensure_default(self, ext: str, ctype: str) -> None:
        ct = self.xml("[Content_Types].xml")
        if any(d.get("Extension", "").lower() == ext.lower() for d in ct.findall(q("ct:Default"))):
            return
        el = ET.Element(q("ct:Default"), Extension=ext, ContentType=ctype)
        ct.insert(0, el)
        self.touch("[Content_Types].xml")

    # relationships
    def rels_root(self, name: str, create: bool = False):
        rn = rels_name(name)
        if rn not in self.parts:
            if not create:
                return None
            self.put_xml(rn, ET.Element(q("pr:Relationships")), [("", NS["pr"])])
        return self.xml(rn)

    def rels(self, name: str) -> list[dict]:
        root = self.rels_root(name)
        if root is None:
            return []
        out = []
        for r in root.findall(q("pr:Relationship")):
            d = dict(r.attrib)
            if d.get("TargetMode") != "External":
                d["part"] = resolve(name, d["Target"])
            out.append(d)
        return out

    def related(self, name: str, rtype: str) -> list[str]:
        return [r["part"] for r in self.rels(name) if r["Type"] == RT + rtype and "part" in r]

    def rel_part(self, name: str, rid: str) -> str | None:
        for r in self.rels(name):
            if r["Id"] == rid:
                return r.get("part")
        return None

    def add_rel(self, name: str, rtype: str, target_part: str) -> str:
        root = self.rels_root(name, create=True)
        ids = {r.get("Id") for r in root}
        n = 1
        while f"rId{n}" in ids:
            n += 1
        rid = f"rId{n}"
        ET.SubElement(
            root,
            q("pr:Relationship"),
            Id=rid,
            Type=RT + rtype if "://" not in rtype else rtype,
            Target=relative(name, target_part),
        )
        self.touch(rels_name(name))
        return rid

    def drop_rel(self, name: str, rid: str) -> None:
        root = self.rels_root(name)
        for r in list(root) if root is not None else []:
            if r.get("Id") == rid:
                root.remove(r)
                self.touch(rels_name(name))

    def new_name(self, pattern: str) -> str:
        """pattern like 'ppt/slides/slide{}.xml' -> first unused number."""
        n = 1
        while pattern.format(n) in self.parts:
            n += 1
        return pattern.format(n)

    # presentation structure
    @property
    def main(self) -> str:
        for r in self.rels(""):
            if r["Type"].endswith("/officeDocument"):
                return r["part"]
        raise ValueError("no officeDocument relationship: not a PresentationML package")

    def slide_ids(self) -> list[tuple]:
        """[(sldId element, rId, slide part)] in presentation order."""
        pres = self.xml(self.main)
        lst = pres.find(q("p:sldIdLst"))
        out = []
        for s in [] if lst is None else list(lst):
            rid = s.get(q("r:id"))
            out.append((s, rid, self.rel_part(self.main, rid)))
        return out

    def slides(self) -> list[str]:
        return [p for _s, _r, p in self.slide_ids()]

    def masters(self) -> list[str]:
        return self.related(self.main, "slideMaster")

    def layouts(self, master: str) -> list[str]:
        root = self.xml(master)
        lst = root.find(q("p:sldLayoutIdLst"))
        return [self.rel_part(master, e.get(q("r:id"))) for e in ([] if lst is None else lst)]

    def reachable(self) -> set[str]:
        seen, todo = set(), [""]
        while todo:
            cur = todo.pop()
            for r in self.rels(cur):
                p = r.get("part")
                if p and p not in seen and p in self.parts:
                    seen.add(p)
                    todo.append(p)
        return seen

    def gc(self) -> list[str]:
        """Delete parts no relationship chain reaches (orphans). Returns deleted names."""
        keep = self.reachable()
        dead = [
            n
            for n in self.parts
            if n != "[Content_Types].xml"
            and not n.endswith(".rels")
            and n not in keep
            and not n.startswith("customXml/")
        ]
        for n in dead:
            self.delete(n)
        # rels files whose source part is gone
        for n in [n for n in self.parts if n.endswith(".rels") and n != "_rels/.rels"]:
            src = rels_source(n)
            if src not in self.parts:
                self.parts.pop(n)
                self._xml.pop(n, None)
                self.dirty.discard(n)
        return dead


def rels_name(name: str) -> str:
    if name == "":
        return "_rels/.rels"
    d, b = posixpath.split(name)
    return posixpath.join(d, "_rels", b + ".rels")


def rels_source(rels: str) -> str:
    d, b = posixpath.split(rels)
    return posixpath.join(posixpath.dirname(d), b[: -len(".rels")])


def resolve(source: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def relative(source: str, target_part: str) -> str:
    return posixpath.relpath(target_part, posixpath.dirname(source) or ".")


# --- geometry -------------------------------------------------------------------


def _xfrm(el):
    name = local(el)
    if name == "graphicFrame":
        return el.find(q("p:xfrm"))
    if name == "grpSp":
        return el.find(f"{q('p:grpSpPr')}/{q('a:xfrm')}")
    return el.find(f"{q('p:spPr')}/{q('a:xfrm')}")


def alt_choice(el):
    """First shape inside mc:Choice of an mc:AlternateContent, else None."""
    ch = el.find(q("mc:Choice"))
    if ch is None:
        return None
    return next((c for c in ch if local(c) in SHAPE_TAGS), None)


def own_box(el):
    if local(el) == "AlternateContent":
        inner = alt_choice(el)
        return own_box(inner) if inner is not None else None
    x = _xfrm(el)
    if x is None or x.find(q("a:off")) is None or x.find(q("a:ext")) is None:
        return None
    o, e = x.find(q("a:off")), x.find(q("a:ext"))
    return [int(o.get("x", 0)), int(o.get("y", 0)), int(e.get("cx", 0)), int(e.get("cy", 0))]


def group_map(grp):
    """Function mapping child-space boxes of a group to its parent space."""
    x = _xfrm(grp)
    if x is None or x.find(q("a:chOff")) is None:
        return lambda b: b
    o, e = x.find(q("a:off")), x.find(q("a:ext"))
    co, ce = x.find(q("a:chOff")), x.find(q("a:chExt"))
    ox, oy, cx, cy = int(o.get("x")), int(o.get("y")), int(e.get("cx")), int(e.get("cy"))
    cox, coy = int(co.get("x")), int(co.get("y"))
    sx = cx / int(ce.get("cx")) if int(ce.get("cx") or 0) else 1.0
    sy = cy / int(ce.get("cy")) if int(ce.get("cy") or 0) else 1.0

    def f(b):
        return [
            round(ox + (b[0] - cox) * sx),
            round(oy + (b[1] - coy) * sy),
            round(b[2] * sx),
            round(b[3] * sy),
        ]

    return f


def intersects(a, b) -> bool:
    return a[0] < b[0] + b[2] and b[0] < a[0] + a[2] and a[1] < b[1] + b[3] and b[1] < a[1] + a[3]


def union(boxes):
    boxes = [b for b in boxes if b]
    if not boxes:
        return None
    x0, y0 = min(b[0] for b in boxes), min(b[1] for b in boxes)
    x1, y1 = max(b[0] + b[2] for b in boxes), max(b[1] + b[3] for b in boxes)
    return [x0, y0, x1 - x0, y1 - y0]


# --- shapes ---------------------------------------------------------------------


def iter_shapes(tree, path=(), mapper=None):
    """Yield (path, element, absolute_box) for every shape, depth-first, groups included."""
    mapper = mapper or (lambda b: b)
    i = 0
    for el in tree:
        if local(el) not in SHAPE_TAGS:
            continue
        i += 1
        p = (*path, i)
        b = own_box(el)
        yield p, el, mapper(b) if b else None
        if local(el) == "grpSp":
            gm = group_map(el)
            yield from iter_shapes(el, p, lambda bb, gm=gm, m=mapper: m(gm(bb)))


def find_path(tree, path):
    """Element at an index path (1-based, counting shape elements only)."""
    el = tree
    for idx in path:
        kids = [c for c in el if local(c) in SHAPE_TAGS]
        if idx > len(kids):
            return None
        el = kids[idx - 1]
    return el


def path_str(path) -> str:
    return "/".join(map(str, path))


def parse_path(s: str):
    return tuple(int(x) for x in s.split("/"))


def cnv(el):
    for nv in el:
        c = nv.find(q("p:cNvPr"))
        if c is not None:
            return c
    return None


def ph_of(el):
    for nv in el:
        nvpr = nv.find(q("p:nvPr"))
        if nvpr is not None:
            return nvpr.find(q("p:ph"))
    return None


def ph_type(ph) -> str:
    return ph.get("type", "obj")


def paragraphs(txbody) -> list[str]:
    if txbody is None:
        return []
    out = []
    for p in txbody.findall(q("a:p")):
        out.append("".join(t.text or "" for t in p.iter(q("a:t"))))
    return out


def txbody(el):
    if local(el) == "sp":
        return el.find(q("p:txBody"))
    if local(el) == "AlternateContent":
        inner = alt_choice(el)
        return txbody(inner) if inner is not None else None
    return None


def all_txbodies(el) -> list:
    """Every text body a slot write must touch (both branches of mc:AlternateContent)."""
    if local(el) == "AlternateContent":
        return list(el.iter(q("p:txBody")))
    tb = txbody(el)
    return [tb] if tb is not None else []


def has_field(el, ftype: str | None = None) -> bool:
    return any(ftype is None or f.get("type") == ftype for f in el.iter(q("a:fld")))


def kind(el) -> str:
    name = local(el)
    if name == "graphicFrame":
        gd = el.find(f".//{q('a:graphicData')}")
        uri = gd.get("uri", "") if gd is not None else ""
        if uri.endswith("/table"):
            return "table"
        if uri.endswith("/chart"):
            return "chart"
        if "diagram" in uri:
            return "smartart"
        return "object"
    if name == "sp":
        if el.find(f"{q('p:nvSpPr')}/{q('p:cNvSpPr')}") is not None and (
            el.find(f"{q('p:nvSpPr')}/{q('p:cNvSpPr')}").get("txBox") == "1"
        ):
            return "textbox"
        return "shape"
    return {"grpSp": "group", "pic": "picture", "cxnSp": "connector"}.get(name, name)


def prst(el) -> str | None:
    g = el.find(f"{q('p:spPr')}/{q('a:prstGeom')}")
    return g.get("prst") if g is not None else None


def colors_in(el) -> list[str]:
    return [c.get("val", "").upper() for c in el.iter(q("a:srgbClr"))]


def fonts_in(el) -> list[str]:
    out = []
    for tag in ("a:latin",):
        for f in el.iter(q(tag)):
            tf = f.get("typeface")
            if tf:
                out.append(tf)
    return out


def sizes_in(el) -> list[float]:
    out = []
    for tag in ("a:rPr", "a:defRPr", "a:endParaRPr"):
        for r in el.iter(q(tag)):
            if r.get("sz"):
                out.append(int(r.get("sz")) / 100)
    return out


# --- theme ----------------------------------------------------------------------

COLOR_SLOTS = [
    "dk1",
    "lt1",
    "dk2",
    "lt2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
]


def theme_of(pkg: Package, master: str) -> dict:
    tparts = pkg.related(master, "theme")
    out = {"part": tparts[0] if tparts else None, "colors": {}, "fonts": {}}
    if not tparts:
        return out
    t = pkg.xml(tparts[0])
    cs = t.find(f".//{q('a:clrScheme')}")
    for slot in COLOR_SLOTS:
        el = cs.find(q(f"a:{slot}")) if cs is not None else None
        if el is None or not len(el):
            continue
        c = el[0]
        hexv = c.get("lastClr") if local(c) == "sysClr" else c.get("val")
        out["colors"][slot] = (hexv or "").upper()
    fs = t.find(f".//{q('a:fontScheme')}")
    for kind_, tag in (("major", "a:majorFont"), ("minor", "a:minorFont")):
        f = fs.find(f"{q(tag)}/{q('a:latin')}") if fs is not None else None
        out["fonts"][kind_] = f.get("typeface") if f is not None else None
    return out


def clr_map(pkg: Package, master: str) -> dict:
    cm = pkg.xml(master).find(q("p:clrMap"))
    return dict(cm.attrib) if cm is not None else {}


def resolve_color(el, theme: dict, cmap: dict) -> str | None:
    """Hex for an a:solidFill-like parent (srgbClr / schemeClr / sysClr)."""
    for c in el:
        n = local(c)
        if n == "srgbClr":
            return c.get("val", "").upper()
        if n == "sysClr":
            return (c.get("lastClr") or "").upper() or None
        if n == "schemeClr":
            v = c.get("val")
            v = cmap.get(v, v)
            return theme["colors"].get(v)
    return None


def luminance(hexv: str) -> float:
    def ch(v):
        v = v / 255
        return v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4

    r, g, b = (int(hexv[i : i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast(h1: str, h2: str) -> float:
    a, b = sorted((luminance(h1), luminance(h2)), reverse=True)
    return round((a + 0.05) / (b + 0.05), 2)


def background(pkg: Package, part: str, theme, cmap):
    """('solid', hex) | ('image', None) | ('none', None) for a slide/layout/master bg."""
    bg = pkg.xml(part).find(f"{q('p:cSld')}/{q('p:bg')}")
    if bg is None:
        return ("none", None)
    pr = bg.find(q("p:bgPr"))
    if pr is not None:
        if pr.find(q("a:blipFill")) is not None:
            return ("image", None)
        sf = pr.find(q("a:solidFill"))
        if sf is not None:
            return ("solid", resolve_color(sf, theme, cmap))
        return ("other", None)
    ref = bg.find(q("p:bgRef"))
    if ref is not None:
        return ("solid", resolve_color(ref, theme, cmap))
    return ("none", None)


# --- text style resolution --------------------------------------------------------


def _lvl_prop(lst, attr: str, lvl: int = 1):
    if lst is None:
        return None
    pp = lst.find(q(f"a:lvl{lvl}pPr"))
    if pp is None:
        return None
    d = pp.find(q("a:defRPr"))
    if d is None:
        return None
    if attr == "latin":
        lat = d.find(q("a:latin"))
        return lat.get("typeface") if lat is not None else None
    return d.get(attr)


def master_style(pkg: Package, master: str, which: str):
    ts = pkg.xml(master).find(q("p:txStyles"))
    return ts.find(q(f"p:{which}")) if ts is not None else None


def style_chain(pkg, layout, master, ph):
    """lstStyle chain (most specific first) for a placeholder on a slide/layout."""
    chain = []
    if ph is not None:
        t = ph_type(ph)
        for part in (layout, master):
            if part is None:
                continue
            m = match_ph(pkg.xml(part), ph, master_level=(part == master))
            if m is not None:
                tb = txbody(m)
                if tb is not None:
                    chain.append(tb.find(q("a:lstStyle")))
        which = (
            "titleStyle"
            if t in TITLE_PH
            else ("bodyStyle" if t in ("body", "obj", "subTitle") else "otherStyle")
        )
    else:
        which = "otherStyle"
    chain.append(master_style(pkg, master, which))
    return chain


def match_ph(tree_root, ph, master_level=False):
    """Find the placeholder in a layout/master that a slide placeholder inherits from."""
    t, idx = ph_type(ph), ph.get("idx", "0")
    cands = []
    for _p, el, _b in iter_shapes(tree_root.find(f"{q('p:cSld')}/{q('p:spTree')}")):
        o = ph_of(el)
        if o is not None:
            cands.append((o, el))
    if not master_level and idx != "0":
        for o, el in cands:
            if o.get("idx", "0") == idx:
                return el
    norm = {"ctrTitle": "title", "subTitle": "body", "obj": "body"}
    for o, el in cands:
        if ph_type(o) == t:
            return el
    for o, el in cands:
        if norm.get(ph_type(o), ph_type(o)) == norm.get(t, t):
            return el
    return None


def effective_box(pkg, el, box, layout, master):
    if box is not None:
        return box
    ph = ph_of(el)
    if ph is None:
        return None
    for part in (layout, master):
        if part is None:
            continue
        m = match_ph(pkg.xml(part), ph, master_level=(part == master))
        if m is not None and own_box(m):
            return own_box(m)
    return None


def text_props(pkg, el, layout, master, theme):
    """(size_pt, typeface, bold) of the first run, resolved through inheritance."""
    tb = txbody(el)
    size = face = None
    bold = False
    if tb is not None:
        for r in tb.iter(q("a:rPr")):
            size = size or r.get("sz")
            bold = bold or r.get("b") == "1"
            lat = r.find(q("a:latin"))
            face = face or (lat.get("typeface") if lat is not None else None)
            break
        chain = [tb.find(q("a:lstStyle")), *style_chain(pkg, layout, master, ph_of(el))]
    else:
        chain = style_chain(pkg, layout, master, ph_of(el))
    for lst in chain:
        size = size or _lvl_prop(lst, "sz")
        face = face or _lvl_prop(lst, "latin")
    size = int(size) / 100 if size else 18.0
    face = face or "+mn-lt"
    if face.startswith("+mj"):
        face = theme["fonts"].get("major") or face
    elif face.startswith("+mn"):
        face = theme["fonts"].get("minor") or face
    return size, face, bold


def para_spacing(pkg, el, layout, master) -> tuple[float, float]:
    """(space before each paragraph in pt, line spacing factor) from the text or its styles."""
    tb = txbody(el)
    chain = []
    if tb is not None:
        p0 = tb.find(q("a:p"))
        if p0 is not None:
            chain.append(p0.find(q("a:pPr")))
        lst = tb.find(q("a:lstStyle"))
        chain.append(lst.find(q("a:lvl1pPr")) if lst is not None else None)
    for lst in style_chain(pkg, layout, master, ph_of(el)):
        chain.append(lst.find(q("a:lvl1pPr")) if lst is not None else None)
    before = line = None
    for ppr in chain:
        if ppr is None:
            continue
        if before is None and ppr.find(f"{q('a:spcBef')}/{q('a:spcPts')}") is not None:
            before = int(ppr.find(f"{q('a:spcBef')}/{q('a:spcPts')}").get("val", "0")) / 100
        if line is None and ppr.find(f"{q('a:lnSpc')}/{q('a:spcPct')}") is not None:
            line = int(ppr.find(f"{q('a:lnSpc')}/{q('a:spcPct')}").get("val", "100000")) / 100000
    return before or 0.0, line or 1.0


def insets(el):
    tb = txbody(el)
    bp = tb.find(q("a:bodyPr")) if tb is not None else None
    g = (lambda k, d: int(bp.get(k, d))) if bp is not None else (lambda k, d: d)
    return g("lIns", 91440), g("tIns", 45720), g("rIns", 91440), g("bIns", 45720)


def capacity(box, size_pt, el=None):
    """Rough (lines, chars_per_line) estimate at 0.5em average glyph width."""
    if not box:
        return None
    li, ti, ri, bi = insets(el) if el is not None else (91440, 45720, 91440, 45720)
    w = max(box[2] - li - ri, 0) / EMU_PER_PT
    h = max(box[3] - ti - bi, 0) / EMU_PER_PT
    return [max(int(h // (size_pt * 1.2)), 1), max(int(w // (size_pt * 0.5)), 1)]


# --- extraction: template -> contract --------------------------------------------


def slug(s: str, n: int = 28) -> str:
    s = re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")
    return s[:n].strip("-") or "slide"


def ph_keys(phs) -> dict:
    """Stable human keys for placeholders: title, subtitle, body, body2, pic, ..."""
    names = {
        "title": "title",
        "ctrTitle": "title",
        "subTitle": "subtitle",
        "body": "body",
        "obj": "body",
        "pic": "pic",
        "tbl": "table",
        "chart": "chart",
        "dgm": "diagram",
        "media": "media",
        "clipArt": "pic",
        "dt": "date",
        "ftr": "footer",
        "sldNum": "slide_number",
    }
    seen: dict[str, int] = {}
    out = {}
    for ph in sorted(phs, key=lambda d: (d["box"] or [0, 0])[1::-1] if d["box"] else (0, 0)):
        base = names.get(ph["type"], ph["type"])
        seen[base] = seen.get(base, 0) + 1
        out[id(ph)] = base if seen[base] == 1 else f"{base}{seen[base]}"
    return out


def describe_shape(pkg, part, el, box, layout, master, theme):
    tb = txbody(el)
    ph = ph_of(el)
    d = {
        "name": (cnv(el).get("name") if cnv(el) is not None else ""),
        "kind": kind(el),
        "box": effective_box(pkg, el, box, layout, master),
    }
    if ph is not None:
        d["ph"] = {"type": ph_type(ph), "idx": int(ph.get("idx", "0"))}
    if tb is not None:
        d["text"] = paragraphs(tb)
    return d


def _chrome(pkg, part, layout, master, theme, where):
    out = []
    tree = pkg.xml(part).find(f"{q('p:cSld')}/{q('p:spTree')}")
    for p, el, box in iter_shapes(tree):
        if len(p) > 1:
            continue
        ph = ph_of(el)
        if ph is not None and ph_type(ph) not in FOOTER_PH:
            continue
        k = kind(el)
        role = "decoration"
        if has_field(el, "slidenum") or (ph is not None and ph_type(ph) == "sldNum"):
            role = "slide_number"
        elif ph is not None and ph_type(ph) == "ftr":
            role = "footer"
        elif ph is not None and ph_type(ph) == "dt":
            role = "date"
        elif k == "picture":
            role = "picture"
        elif k in ("textbox", "shape") and any(paragraphs(txbody(el))):
            role = "text"
        out.append(
            {
                "role": role,
                "kind": k,
                "name": cnv(el).get("name", "") if cnv(el) is not None else "",
                "box": effective_box(pkg, el, box, layout, master),
                "where": where,
                "placeholder": ph is not None,
            }
        )
    return out


def profile(path) -> dict:
    """Extract the style contract from a template file. Pure function of the file."""
    pkg = Package(path)
    data = Path(path).read_bytes()
    pres = pkg.xml(pkg.main)
    sz = pres.find(q("p:sldSz"))
    slide_w, slide_h = int(sz.get("cx")), int(sz.get("cy"))
    masters = pkg.masters()
    if not masters:
        raise ValueError("template has no slide master")
    if "macroEnabled" in (pkg.content_type(pkg.main) or ""):
        raise ValueError(
            "macro-enabled files (.pptm/.potm) are not supported; save a .pptx/.potx copy"
            " (macros are never run or copied)"
        )
    master0 = masters[0]
    theme = theme_of(pkg, master0)
    cmap = clr_map(pkg, master0)
    kind_ = "potx" if pkg.content_type(pkg.main) == CT_TMPL else "pptx"

    hashed = [pkg.main, *masters]
    if theme["part"]:
        hashed.append(theme["part"])
    contract: dict = {
        "schema": SCHEMA,
        "tool": f"pptx-style {VERSION}",
        "source": {"file": Path(path).name, "sha256": sha256(data), "kind": kind_, "parts": {}},
        "slide_size": [slide_w, slide_h],
        "theme": {"colors": theme["colors"], "fonts": theme["fonts"], "text_styles": {}},
        "masters": [],
        "layouts": [],
        "patterns": [],
        "derived": {},
        "ai_readiness": [],
    }
    for which in ("titleStyle", "bodyStyle", "otherStyle"):
        st = master_style(pkg, master0, which)
        lv = {}
        for n in range(1, 4):
            s = _lvl_prop(st, "sz", n)
            if s:
                lv[f"lvl{n}"] = int(s) / 100
        contract["theme"]["text_styles"][which] = lv

    findings = contract["ai_readiness"]

    def finding(code, severity, where, message, fix):
        findings.append(
            {"code": code, "severity": severity, "where": where, "message": message, "fix": fix}
        )

    # masters + layouts
    layout_by_part = {}
    for m in masters:
        mbg = background(pkg, m, theme, cmap)
        contract["masters"].append(
            {
                "part": m,
                "background": mbg[0],
                "chrome": _chrome(pkg, m, None, m, theme, "master"),
            }
        )
        for lp in pkg.layouts(m):
            hashed.append(lp)
            lroot = pkg.xml(lp)
            name = lroot.find(q("p:cSld")).get("name", "")
            tree = lroot.find(f"{q('p:cSld')}/{q('p:spTree')}")
            phs = []
            for _p, el, box in iter_shapes(tree):
                ph = ph_of(el)
                if ph is None or ph_type(ph) in FOOTER_PH:
                    continue
                b = effective_box(pkg, el, box, None, m)
                size, face, _bold = text_props(pkg, el, None, m, theme)
                prompt = paragraphs(txbody(el))
                prompt_text = "\n".join(t for t in prompt if t)
                phs.append(
                    {
                        "type": ph_type(ph),
                        "idx": int(ph.get("idx", "0")),
                        "box": b,
                        "prompt": prompt_text,
                        "custom_prompt": ph.get("hasCustomPrompt") == "1",
                        "prompt_looks_like_content": bool(
                            prompt_text and PROMPT_CONTENT_RE.match(prompt_text)
                        ),
                        "font_size": size,
                        "font": face,
                        "capacity": capacity(b, size, el),
                        "spacing": list(para_spacing(pkg, el, None, m)),
                    }
                )
            keys = ph_keys(phs)
            for ph in phs:
                ph["key"] = keys[id(ph)]
                if ph["prompt_looks_like_content"]:
                    finding(
                        "prompt-text-as-content",
                        "warning",
                        f"layout {name!r} placeholder {ph['key']}",
                        f"prompt text {ph['prompt']!r} looks like content but only shows in edit view, never in the show or PDF",
                        "type the value into the slide placeholder (or move it into the layout as real text)",
                    )
            phs.sort(
                key=lambda d: (d["key"] != "title", (d["box"] or [0, 0])[1], (d["box"] or [0])[0])
            )
            bg = background(pkg, lp, theme, cmap)
            pics = sum(1 for _p, el, _b in iter_shapes(tree) if kind(el) == "picture")
            ltype = lroot.get("type")
            show_master = lroot.get("showMasterSp", "1") != "0"
            chrome = _chrome(pkg, lp, lp, m, theme, "layout")
            sig = {}
            for ph in phs:
                k = re.sub(r"\d+$", "", ph["key"])
                sig[k] = sig.get(k, 0) + 1
            L = {
                "name": name,
                "part": lp,
                "type": ltype,
                "master": m,
                "show_master_shapes": show_master,
                "signature": dict(sorted(sig.items())),
                "placeholders": phs,
                "chrome": chrome,
                "art": {"background": bg[0], "pictures": pics},
            }
            contract["layouts"].append(L)
            layout_by_part[lp] = L
            if any(ord(ch) > 127 for ch in name):
                finding(
                    "localized-layout-name",
                    "info",
                    f"layout {name!r}",
                    "layout name is not ASCII/English; agents and users may not recognise it",
                    "rename the layout to a descriptive English name (keep an alias in the contract)",
                )
            for _p, el, _b in iter_shapes(tree):
                if ph_of(el) is None and kind(el) == "textbox" and not has_field(el):
                    txt = " ".join(paragraphs(txbody(el))).strip()
                    if txt:
                        finding(
                            "textbox-instead-of-placeholder",
                            "warning",
                            f"layout {name!r} shape {cnv(el).get('name')!r}",
                            f"free text box {txt[:40]!r} on a layout: AI tools cannot fill it",
                            "replace it with a real placeholder, or keep it as fixed chrome on purpose",
                        )
    # signature collisions
    groups: dict[str, list[str]] = {}
    for L in contract["layouts"]:
        key = json.dumps(L["signature"], sort_keys=True)
        groups.setdefault(key, []).append(L["name"])
    for L in contract["layouts"]:
        same = groups[json.dumps(L["signature"], sort_keys=True)]
        L["collision_group"] = sorted(same) if len(same) > 1 else []
    names = [L["name"] for L in contract["layouts"]]
    for n in {n for n in names if names.count(n) > 1}:
        finding(
            "duplicate-layout-name",
            "error",
            f"layout {n!r}",
            "two layouts share a name; name-based selection is ambiguous",
            "rename one of the layouts",
        )

    # patterns = sample slides
    used_layouts = set()
    pattern_ids: set[str] = set()
    all_colors: dict[str, list[int]] = {}
    sizes: dict[str, list[float]] = {"title": [], "body": [], "other": []}
    fonts_used: set[str] = set()
    chrome_on_slides = False
    all_sizes: list[float] = []
    for n, sp in enumerate(pkg.slides(), 1):
        sroot = pkg.xml(sp)
        lp = pkg.related(sp, "slideLayout")[0]
        L = layout_by_part.get(lp)
        used_layouts.add(lp)
        m = L["master"] if L else master0
        tree = sroot.find(f"{q('p:cSld')}/{q('p:spTree')}")
        for c in colors_in(sroot):
            all_colors.setdefault(c, []).append(n)
        all_sizes.extend(sizes_in(sroot))
        fonts_used.update(fonts_in(sroot))
        p = _pattern(pkg, sp, n, tree, lp, m, theme, L)
        base = slug(p["title"] or (L["name"] if L else "slide"))
        pid, k = base, 2
        while pid in pattern_ids:
            pid, k = f"{base}-{k}", k + 1
        pattern_ids.add(pid)
        p["id"] = pid
        contract["patterns"].append(p)
        for _p, el, _b in iter_shapes(tree):
            ph = ph_of(el)
            if ph is not None and ph_type(ph) in FOOTER_PH:
                chrome_on_slides = True
            s, _f, _bold = (
                text_props(pkg, el, lp, m, theme) if txbody(el) is not None else (None, 0, 0)
            )
            if s and any(paragraphs(txbody(el))):
                role = (
                    "title"
                    if ph is not None and ph_type(ph) in TITLE_PH
                    else ("body" if ph is not None else "other")
                )
                sizes[role].append(s)
        dup = p.pop("_dupes")
        if dup:
            finding(
                "duplicate-shape-names",
                "warning",
                f"slide {n} ({pid})",
                f"shape names repeat: {dup}; name-based addressing would hit the wrong shape",
                "rename shapes uniquely in the Selection Pane (the contract addresses by index path)",
            )
        odd = sorted({nm for nm in p.pop("_names") if re.search(r"Google Shape;|;p\d+$", nm)})
        if odd:
            finding(
                "export-artifact-names",
                "info",
                f"slide {n} ({pid})",
                f"shape names look like converter artefacts: {odd[:3]}",
                "rename shapes to describe their role",
            )
        if not p["placeholders_used"]:
            finding(
                "decorative-sample-slide",
                "info",
                f"slide {n} ({pid})",
                "sample slide has no placeholders; only clone-and-retext can reuse it",
                "rebuild the slide on placeholders if it should be fillable by AI tools",
            )

    if len(masters) > 1:
        finding(
            "multiple-masters",
            "warning",
            f"{len(masters)} slide masters",
            "theme, fonts and text styles are profiled from the first master only",
            "keep one master per template file, or give every layout a unique name",
        )
    if not contract["patterns"]:
        finding(
            "no-sample-slides",
            "warning",
            "template",
            "the file has no slides, so there are no patterns to clone; builds can only fill bare layouts",
            "add sample slides (timeline, comparison, cards, numbers) that show how content should look",
        )
    for L in contract["layouts"]:
        L["has_sample_slide"] = L["part"] in used_layouts
        if not L["has_sample_slide"]:
            finding(
                "layout-without-sample",
                "info",
                f"layout {L['name']!r}",
                "no sample slide uses this layout; tools fall back to the bare layout",
                "add a sample slide showing intended content density",
            )

    theme_hex = set(theme["colors"].values())
    off = {c: ev for c, ev in all_colors.items() if c not in theme_hex}
    for c, ev in sorted(off.items()):
        finding(
            "hardcoded-colour",
            "warning",
            f"slides {sorted(set(ev))}",
            f"colour #{c} is not a theme colour",
            "map it to a theme colour slot so recolouring the theme updates it",
        )
    master_chrome = [c for M in contract["masters"] for c in M["chrome"]]
    footer = [c["box"] for c in master_chrome if c["box"] and c["box"][1] > slide_h * 0.8] + [
        c["box"]
        for L in contract["layouts"]
        for c in L["chrome"]
        if c["box"] and c["box"][1] > slide_h * 0.8
    ]
    content = [
        ph["box"]
        for L in contract["layouts"]
        for ph in L["placeholders"]
        if ph["box"] and ph["type"] in ("title", "body", "obj", "ctrTitle", "subTitle")
    ]
    contract["derived"] = {
        "note": "Derived values are inferred from the template; evidence lists slide numbers.",
        "colors_used": {c: sorted(set(ev)) for c, ev in sorted(all_colors.items())},
        "off_theme_colors": sorted(off),
        "fonts_used": sorted(fonts_used),
        "type_scale": {k: sorted(set(v)) for k, v in sizes.items() if v},
        "min_font_size": min([*all_sizes, *(s for v in sizes.values() for s in v)], default=None),
        "safe_area": union(content),
        "footer_band": union(footer),
        "footer_placeholders_on_slides": chrome_on_slides,
        "contrast": _contrast_table(theme),
    }
    weak = sorted(
        {
            f"{r['fg']} on {r['on']} {r['ratio']}:1"
            for r in contract["derived"]["contrast"]
            if r["on"] in ("lt1", "dk1") and r["fg"].startswith("accent") and not r["aa_text"]
        }
    )
    if weak:
        finding(
            "accent-fails-aa-as-text",
            "info",
            "theme colours",
            f"below 4.5:1 for normal text: {', '.join(weak)}",
            "use these for fills or large text (>=18 pt, or >=14 pt bold) only; small text uses dk1/dk2",
        )
    for part in hashed:
        contract["source"]["parts"][part] = sha256(pkg.parts[part])
    return contract


def _contrast_table(theme) -> list[dict]:
    cols = theme["colors"]
    rows = []
    for bgk in ("lt1", "lt2", "dk1", "dk2"):
        if bgk not in cols:
            continue
        for fgk in ("dk1", "dk2", "lt1", "lt2", *[f"accent{i}" for i in range(1, 7)]):
            if fgk == bgk or fgk not in cols:
                continue
            r = contrast(cols[fgk], cols[bgk])
            rows.append(
                {"fg": fgk, "on": bgk, "ratio": r, "aa_text": r >= 4.5, "aa_large": r >= 3.0}
            )
    return rows


def _pattern(pkg, sp, n, tree, lp, master, theme, L) -> dict:
    """Describe a sample slide as a reusable pattern with tagged slots."""
    names: list[str] = []
    slots: dict[str, dict] = {}
    repeats: dict[str, dict] = {}
    counters: dict[str, int] = {}
    tables = []
    phs_used = False
    title = ""

    def tag(prefix):
        counters[prefix] = counters.get(prefix, 0) + 1
        return f"{prefix}{counters[prefix]}"

    shapes = list(iter_shapes(tree))
    for _p, el, _b in shapes:
        c = cnv(el)
        names.append(c.get("name", "") if c is not None else "")
        ph = ph_of(el)
        if ph is not None:
            phs_used = True
            if ph_type(ph) in TITLE_PH and not title:
                title = " ".join(paragraphs(txbody(el))).strip()
    if not title:
        for _p, el, _b in sorted(shapes, key=lambda t: (t[2] or [0, 0])[1]):
            txt = " ".join(paragraphs(txbody(el))).strip()
            if txt:
                title = txt
                break

    # repeat groups: sibling groups with identical structure
    def struct(el):
        return (local(el), tuple(kind(c) for c in el if local(c) in SHAPE_TAGS))

    by_parent: dict[tuple, list] = {}
    for p, el, b in shapes:
        if local(el) == "grpSp":
            by_parent.setdefault((p[:-1], struct(el)), []).append((p, el, b))
    in_repeat: set[tuple] = set()
    for (_parent, _st), members in sorted(by_parent.items(), key=lambda kv: kv[1][0][0]):
        if len(members) < 2:
            continue
        sizes = {
            (round((b or [0, 0, 0, 0])[2], -4), round((b or [0, 0, 0, 0])[3], -4))
            for _p, _e, b in members
        }
        if len(sizes) > 2:
            continue
        members.sort(key=lambda t: ((t[2] or [0, 0])[1] // 100000, (t[2] or [0])[0]))
        inner = []
        sub_counters: dict[str, int] = {}
        for sp_path, el2, b2 in iter_shapes(members[0][1], ()):
            if txbody(el2) is not None and any(paragraphs(txbody(el2))) and not has_field(el2):
                sub_counters["t"] = sub_counters.get("t", 0) + 1
                inner.append(
                    {
                        "tag": f"t{sub_counters['t']}",
                        "path": path_str(sp_path),
                        "text": paragraphs(txbody(el2)),
                        "box": b2,
                    }
                )
        if not inner:
            continue
        g = tag("g")
        repeats[g] = {
            "kind": "repeat",
            "members": [path_str(p) for p, _e, _b in members],
            "slots": inner,
            "count": len(members),
        }
        for p, _e, _b in members:
            in_repeat.add(p)

    def inside_repeat(p):
        return any(p[: len(r)] == r for r in in_repeat)

    ordered = sorted(shapes, key=lambda t: (((t[2] or [0, 0])[1]) // 50000, (t[2] or [0])[0]))
    for p, el, box in ordered:
        if inside_repeat(p):
            continue
        k = kind(el)
        ph = ph_of(el)
        b = effective_box(pkg, el, box, lp, master)
        if k == "table":
            tbl = el.find(f".//{q('a:tbl')}")
            rows = [
                [" ".join(paragraphs(tc.find(q("a:txBody")))) for tc in tr.findall(q("a:tc"))]
                for tr in tbl.findall(q("a:tr"))
            ]
            t = tag("tbl")
            slots[t] = {
                "kind": "table",
                "path": path_str(p),
                "rows": len(rows),
                "cols": len(rows[0]) if rows else 0,
                "text": rows,
                "box": b,
            }
            tables.append((t, b))
        elif k == "chart":
            slots[tag("chart")] = {"kind": "chart", "path": path_str(p), "box": b}
        elif k == "picture" and len(p) >= 1:
            slots[tag("img")] = {
                "kind": "image",
                "path": path_str(p),
                "box": b,
                "name": names[0]
                if False
                else (cnv(el).get("name", "") if cnv(el) is not None else ""),
            }
        elif txbody(el) is not None and not has_field(el):
            paras = paragraphs(txbody(el))
            if not any(paras) and ph is None:
                continue
            size, face, _b = text_props(pkg, el, lp, master, theme)
            entry = {
                "kind": "bullets" if len([x for x in paras if x]) > 1 else "text",
                "path": path_str(p),
                "text": paras,
                "box": b,
                "font_size": size,
                "font": face,
                "capacity": capacity(b, size, el),
                "spacing": list(para_spacing(pkg, el, lp, master)),
            }
            key = None
            if ph is not None:
                entry["ph"] = {"type": ph_type(ph), "idx": int(ph.get("idx", "0"))}
                for lph in (L or {}).get("placeholders", []):
                    same_idx = lph["idx"] == entry["ph"]["idx"] and entry["ph"]["idx"] != 0
                    same_title = (
                        entry["ph"]["idx"] == 0
                        and lph["type"] in TITLE_PH
                        and ph_type(ph) in TITLE_PH
                    )
                    if same_idx or same_title:
                        key = lph["key"]
                        break
            slots[key if key and key not in slots else tag("t")] = entry
    # bars / marks: text-less autoshapes overlapping a table, at top level
    for p, el, box in shapes:
        if len(p) != 1 or local(el) != "sp" or not box or inside_repeat(p):
            continue
        if any(paragraphs(txbody(el))):
            continue
        for t, tb in tables:
            if tb and intersects(box, tb):
                kind_ = "bar" if box[2] > 2 * box[3] else "mark"
                slots[tag(kind_)] = {"kind": kind_, "path": path_str(p), "table": t, "box": box}
                break
    deps = {"images": 0, "charts": 0, "embeds": 0, "other": 0}
    for r in pkg.rels(sp):
        t = r["Type"].rsplit("/", 1)[-1]
        if t == "image":
            deps["images"] += 1
        elif t == "chart":
            deps["charts"] += 1
        elif t in ("package", "oleObject"):
            deps["embeds"] += 1
        elif t not in ("slideLayout", "notesSlide", "hyperlink", "slide"):
            deps["other"] += 1
    counts: dict[str, int] = {}
    for nm in names:
        counts[nm] = counts.get(nm, 0) + 1
    notes = pkg.related(sp, "notesSlide")
    note_text = ""
    if notes:
        for _p, el, _b in iter_shapes(pkg.xml(notes[0]).find(f"{q('p:cSld')}/{q('p:spTree')}")):
            ph = ph_of(el)
            if ph is not None and ph_type(ph) == "body":
                note_text = "\n".join(paragraphs(txbody(el))).strip()
    return {
        "slide": n,
        "part": sp,
        "layout": L["name"] if L else None,
        "title": title[:80],
        "slots": slots,
        "repeats": repeats,
        "dependencies": deps,
        "placeholders_used": phs_used,
        "hidden": pkg.xml(sp).get("show") == "0",
        "asset_bank": bool(slots)
        and all(v["kind"] == "image" for v in slots.values())
        and not repeats,
        "notes": note_text[:500],
        "fingerprint": sha256(pkg.parts[sp])[:16],
        "shape_boxes": [b for p, _e, b in shapes if len(p) == 1 and b],
        "_dupes": {k: v for k, v in counts.items() if v > 1 and k},
        "_names": names,
    }


# --- diff --------------------------------------------------------------------------


def diff_contracts(old: dict, new: dict, tol_ratio: float = 0.0025) -> dict:
    """Classify differences: breaking (specs built on old may fail), changed, additive."""
    out: dict[str, list[str]] = {"breaking": [], "changed": [], "additive": []}
    tol = max(EMU_PER_PT, int(old["slide_size"][0] * tol_ratio))
    if old["slide_size"] != new["slide_size"]:
        out["breaking"].append(f"slide size {old['slide_size']} -> {new['slide_size']}")
    ol = {L["name"]: L for L in old["layouts"]}
    nl = {L["name"]: L for L in new["layouts"]}

    def geom(L):
        return sorted((p["key"], tuple(p["box"] or ())) for p in L["placeholders"])

    for name, L in ol.items():
        if name in nl:
            continue
        twin = [n for n, M in nl.items() if n not in ol and geom(M) == geom(L)]
        if twin:
            out["breaking"].append(
                f"layout {name!r} renamed to {twin[0]!r} (add an alias or update specs)"
            )
        else:
            out["breaking"].append(f"layout {name!r} removed")
    renamed_to = {
        n
        for n, M in nl.items()
        if n not in ol and any(geom(M) == geom(L) for x, L in ol.items() if x not in nl)
    }
    for name in nl:
        if name not in ol and name not in renamed_to:
            out["additive"].append(f"layout {name!r} added")
    for name in ol.keys() & nl.keys():
        op = {p["key"]: p for p in ol[name]["placeholders"]}
        np_ = {p["key"]: p for p in nl[name]["placeholders"]}
        for k, p in op.items():
            if k not in np_:
                out["breaking"].append(f"layout {name!r}: placeholder {k} removed")
                continue
            q_ = np_[k]
            if p["type"] != q_["type"]:
                out["breaking"].append(
                    f"layout {name!r}: placeholder {k} type {p['type']} -> {q_['type']}"
                )
            if p["box"] and q_["box"]:
                delta = max(abs(a - b) for a, b in zip(p["box"], q_["box"], strict=False))
                if delta > tol:
                    out["breaking"].append(
                        f"layout {name!r}: placeholder {k} moved/resized by {delta / EMU_PER_PT:.1f} pt (tolerance {tol / EMU_PER_PT:.1f} pt): {p['box']} -> {q_['box']}"
                    )
        for k in np_.keys() - op.keys():
            out["additive"].append(f"layout {name!r}: placeholder {k} added")
        if [c["box"] for c in ol[name]["chrome"]] != [c["box"] for c in nl[name]["chrome"]]:
            out["changed"].append(f"layout {name!r}: chrome shapes changed")
    opat = {p["id"]: p for p in old["patterns"]}
    npat = {p["id"]: p for p in new["patterns"]}

    def shape_sig(p):
        return json.dumps(
            [
                sorted((k, v["kind"]) for k, v in p["slots"].items()),
                sorted((g, r["count"]) for g, r in p["repeats"].items()),
            ]
        )

    added = {pid for pid in npat if pid not in opat}
    for pid, p in opat.items():
        if pid not in npat:
            twin = [x for x in sorted(added) if shape_sig(npat[x]) == shape_sig(p)]
            if twin:
                added.discard(twin[0])
                out["breaking"].append(
                    f"pattern {pid!r} renamed to {twin[0]!r} (the sample slide's title changed; update specs)"
                )
            else:
                out["breaking"].append(f"pattern {pid!r} removed")
            continue
        q_ = npat[pid]
        for s_ in p["slots"].keys() - q_["slots"].keys():
            out["breaking"].append(f"pattern {pid!r}: slot {s_} removed")
        for s_ in p["slots"].keys() & q_["slots"].keys():
            if p["slots"][s_]["kind"] != q_["slots"][s_]["kind"]:
                out["breaking"].append(
                    f"pattern {pid!r}: slot {s_} changed {p['slots'][s_]['kind']} -> {q_['slots'][s_]['kind']}"
                )
        for g in p["repeats"]:
            if g not in q_["repeats"]:
                out["breaking"].append(f"pattern {pid!r}: repeat group {g} removed")
            elif q_["repeats"][g]["count"] < p["repeats"][g]["count"]:
                out["breaking"].append(
                    f"pattern {pid!r}: repeat group {g} holds {q_['repeats'][g]['count']} items, was {p['repeats'][g]['count']}"
                )
        if p["layout"] != q_["layout"]:
            out["changed"].append(f"pattern {pid!r}: layout {p['layout']!r} -> {q_['layout']!r}")
    for pid in sorted(added):
        out["additive"].append(f"pattern {pid!r} added")
    for slot, v in old["theme"]["colors"].items():
        nv = new["theme"]["colors"].get(slot)
        if nv != v:
            out["changed"].append(f"theme colour {slot}: #{v} -> #{nv}")
    for k, v in old["theme"]["fonts"].items():
        if new["theme"]["fonts"].get(k) != v:
            out["changed"].append(f"theme {k} font: {v!r} -> {new['theme']['fonts'].get(k)!r}")
    return out


def load_contract(path) -> dict:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if data.get("schema") != SCHEMA:
        raise ValueError(f"{path}: not a {SCHEMA} contract (schema={data.get('schema')!r})")
    return data


def contract_or_profile(path) -> dict:
    """A contract.json, or profile a .pptx/.potx on the fly."""
    p = Path(path)
    if p.suffix.lower() == ".json":
        return load_contract(p)
    return profile(p)


def new_el(tag: str, **attrs):
    return ET.Element(q(tag), {k: str(v) for k, v in attrs.items()})


def sub(parent, tag: str, **attrs):
    return ET.SubElement(parent, q(tag), {k: str(v) for k, v in attrs.items()})


def deepcopy(el):
    return copy.deepcopy(el)


# --- rendering (optional; LibreOffice + poppler) -----------------------------------


def find_soffice() -> str | None:
    """Path to a working soffice. A stale shim (app removed, wrapper left) counts as absent."""
    import shutil
    import subprocess

    cands = [shutil.which("soffice"), shutil.which("libreoffice")]
    cands.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    for c in cands:
        if not c or not Path(c).exists():
            continue
        try:
            r = subprocess.run([c, "--version"], capture_output=True, timeout=60, check=False)
        except (OSError, subprocess.TimeoutExpired):
            continue
        if r.returncode == 0:
            return c
    return None


def installed_fonts() -> set[str] | None:
    """Lower-cased family names known to fontconfig, or None if fc-list is absent."""
    import shutil
    import subprocess

    if not shutil.which("fc-list"):
        return None
    out = subprocess.run(
        ["fc-list", ":", "family"], capture_output=True, text=True, timeout=60, check=False
    ).stdout
    return {f.strip().lower() for line in out.splitlines() for f in line.split(",") if f.strip()}


def render(deck, outdir, timeout: int = 240, dpi: int = 60) -> dict:
    """deck -> outdir/deck.pdf + slide-N.png. Returns {status, pdf, pngs, reason}."""
    import shutil
    import subprocess
    import tempfile

    outdir = Path(outdir)
    soffice = find_soffice()
    if not soffice:
        return {"status": "skipped", "reason": "LibreOffice (soffice) not found"}
    outdir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as prof:
        # fonts in $PPTX_STYLE_FONT_DIRS load into the private profile, so a template
        # font file renders without being installed system-wide (fit.py reads the same dirs)
        fonts = Path(prof) / "lo" / "user" / "fonts"
        for d in filter(None, os.environ.get("PPTX_STYLE_FONT_DIRS", "").split(os.pathsep)):
            for f in Path(d).glob("*"):
                if f.suffix.lower() in (".ttf", ".otf", ".ttc"):
                    fonts.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(f, fonts / f.name)
        src = Path(deck)
        if src.suffix.lower() == ".potx":  # LibreOffice opens templates as new untitled docs
            tmp = Path(prof) / (src.stem + ".pptx")
            pkg = Package(src)
            pkg.set_override(pkg.main, CT_PRES)
            pkg.save(tmp)
            src = tmp
        cmd = [
            soffice,
            f"-env:UserInstallation={Path(prof).as_uri()}/lo",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(outdir),
            str(src),
        ]
        try:
            subprocess.run(cmd, capture_output=True, timeout=timeout, check=False)
        except subprocess.TimeoutExpired:
            return {"status": "failed", "reason": f"soffice timed out after {timeout}s"}
    pdf = outdir / (Path(deck).stem + ".pdf")
    if not pdf.exists():
        return {"status": "failed", "reason": "soffice produced no PDF"}
    pngs: list[str] = []
    if shutil.which("pdftoppm"):
        for old in outdir.glob("slide-*.png"):
            old.unlink()
        subprocess.run(
            ["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(outdir / "slide")],
            capture_output=True,
            timeout=timeout,
            check=False,
        )
        pngs = sorted(str(p) for p in outdir.glob("slide-*.png"))
    return {
        "status": "ok",
        "pdf": str(pdf),
        "pngs": pngs,
        "reason": "" if pngs else "pdftoppm not found: PDF only",
    }
