#!/usr/bin/env python3
"""Build or edit a native .pptx strictly inside a style contract.

    python3 build.py SPEC.json [--out deck.pptx] [--contract C.json] [--slides 1,3]
    python3 build.py --smoke-spec C.json > smoke.json   # one slide per layout and pattern
    python3 build.py --self-test

SPEC (JSON):
  {"contract": ".pptx-style/<slug>/contract.json", "output": "deck.pptx",
   "base": "existing.pptx",            # optional: edit mode, see below
   "require_notes": true,
   "slides": [
     {"layout": "Title and Content", "fill": {"title": "...", "body": ["a", "b"]}, "notes": "..."},
     {"pattern": "timeline", "fill": {"title": "...", "tbl1": [["", "Q1"], ["Build", ""]],
                                       "bar1": {"row": 2, "from": 2, "to": 3}, "bar2": null}},
     {"pattern": "three-ways-we-help", "fill": {"t1": "...", "g1": [{"t1": "..", "t2": ".."}]}},
     {"keep": 3, "fill": {"title": "New title"}, "replace": {"old text": "new text"}}
   ]}

Fill values: text "a\\nb" | list of paragraphs (str or {"text", "level"}) | null (delete
the shape) | table rows | {"categories": [...], "series": [{"name", "values"}]} for charts |
{"image": "path.png"} | bar {"row", "from", "to"} (1-based table cells) | true/false marks.
Formatting keys (colour, font, size, position, text boxes) are rejected: styling comes
from the template. Edit mode ("base"): slides listed as {"keep": n} pass through
byte-identical unless they carry fill/replace/notes; base slides not listed are removed.

Exit codes: 0 built, lint clean | 1 built with lint/fit errors | 2 spec or contract error
(nothing written) | 3 template drifted with breaking changes (nothing written).
"""

from __future__ import annotations

import argparse
import difflib
import io
import json
import re
import struct
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fit
import pptx_core as core
from pptx_core import q

SLIDE_KEYS = {
    "layout",
    "pattern",
    "keep",
    "fill",
    "notes",
    "allow_split",
    "replace",
    "comment",
    "id",
}
SPEC_KEYS = {
    "contract",
    "template",
    "base",
    "remove",
    "output",
    "require_notes",
    "slides",
    "comment",
}
STYLE_WORDS = {
    "color",
    "colour",
    "font",
    "font_size",
    "size",
    "bold",
    "italic",
    "x",
    "y",
    "left",
    "top",
    "width",
    "height",
    "position",
    "textbox",
    "text_box",
    "shape",
    "style",
    "fill_color",
    "background",
}
ILLEGAL_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")
PLACEMENT_WORDS = {
    "x",
    "y",
    "left",
    "top",
    "width",
    "height",
    "position",
    "textbox",
    "text_box",
    "shape",
}
IMAGE_CT = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif"}
MEDIA_TYPES = {"image", "audio", "video", "media", "hdphoto"}


class SpecError(Exception):
    pass


def did_you_mean(word, options) -> str:
    m = difflib.get_close_matches(str(word), [str(o) for o in options], n=3, cutoff=0.4)
    return (
        f" (did you mean {', '.join(repr(x) for x in m)}?)"
        if m
        else f" (valid: {', '.join(sorted(map(str, options))[:12])})"
    )


def off_contract(key: str) -> str | None:
    k = key.lower()
    if k in PLACEMENT_WORDS:
        return (
            f"off-contract: {key!r} would place a free shape. Content goes into the template's "
            "placeholders and pattern slots only (notes for extra text, or ask the template owner for a layout)"
        )
    if k in STYLE_WORDS or any(w in k for w in ("color", "colour", "font")):
        return (
            f"off-contract: {key!r} would set formatting directly. Styling comes from the template: "
            "pick a layout or pattern whose styling fits, use a theme role, or ask the template owner"
        )
    return None


# --- text ------------------------------------------------------------------------------


def items(value) -> list[tuple[str, int | None]]:
    if isinstance(value, str):
        return [(t, None) for t in value.split("\n")]
    out = []
    for v in value:
        if isinstance(v, dict):
            out.append((str(v.get("text", "")), v.get("level")))
        else:
            out.append((str(v), None))
    return out


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(q("a:t")))


def set_para(p, text: str, level) -> None:
    """Replace a paragraph's text keeping pPr and the first run's rPr."""
    rpr = None
    for child in list(p):
        n = core.local(child)
        if n in ("r", "fld") and rpr is None:
            rpr = child.find(q("a:rPr"))
        if n in ("r", "fld", "br"):
            p.remove(child)
    end = p.find(q("a:endParaRPr"))
    if rpr is None and end is not None:
        rpr = core.deepcopy(end)
        rpr.tag = q("a:rPr")
    r = core.new_el("a:r")
    r.append(
        core.deepcopy(rpr) if rpr is not None else core.new_el("a:rPr", lang="en-US", dirty="0")
    )
    t = core.sub(r, "a:t")
    t.text = ILLEGAL_XML.sub("", text)
    if end is not None:
        p.insert(list(p).index(end), r)
    else:
        p.append(r)
    if level is not None:
        ppr = p.find(q("a:pPr"))
        if ppr is None:
            ppr = core.new_el("a:pPr")
            p.insert(0, ppr)
        ppr.set("lvl", str(int(level)))


def set_text(txbody, value) -> None:
    its = items(value)
    paras = txbody.findall(q("a:p"))
    if not paras:
        paras = [core.sub(txbody, "a:p")]
    proto = paras[-1]
    for i, (text, lvl) in enumerate(its):
        if i < len(paras):
            p = paras[i]
        else:
            p = core.deepcopy(proto)
            txbody.append(p)
        set_para(p, text, lvl)
    for p in paras[len(its) :]:
        txbody.remove(p)


def new_txbody(value):
    tb = core.new_el("p:txBody")
    core.sub(tb, "a:bodyPr")
    core.sub(tb, "a:lstStyle")
    p = core.sub(tb, "a:p")
    core.sub(p, "a:endParaRPr", lang="en-US", dirty="0")
    set_text(tb, value)
    return tb


def replace_runs(root, mapping: dict) -> int:
    """Run-level find/replace; every key must occur inside a single run (formatting kept)."""
    n = 0
    for old, new in mapping.items():
        hit = False
        for t in root.iter(q("a:t")):
            if t.text and old in t.text:
                t.text = t.text.replace(old, ILLEGAL_XML.sub("", str(new)))
                n += 1
                hit = True
        if not hit:
            spans = any(old in para_text(p) for p in root.iter(q("a:p")))
            raise SpecError(
                f"replace {old!r}: "
                + (
                    "the text spans differently formatted runs; replace the whole placeholder via fill instead"
                    if spans
                    else "not found on the slide"
                )
            )
    return n


# --- package surgery ----------------------------------------------------------------


def number_pattern(part: str) -> str:
    d, b = part.rsplit("/", 1)
    m = re.match(r"^(.*?)(\d*)(\.[^.]+)$", b)
    return f"{d}/{m.group(1)}{{}}{m.group(3)}"


def copy_part(src: core.Package, part: str, dst: core.Package, memo: dict) -> str:
    """Deep-copy a part and its relationship subtree; media is shared (same package) or deduped."""
    if (id(src), part) in memo:
        return memo[(id(src), part)]
    ctype = src.content_type(part)
    new = dst.new_name(number_pattern(part))
    dst.put_bytes(new, src.parts[part])
    memo[(id(src), part)] = new
    ext = part.rsplit(".", 1)[-1]
    if ctype and dst.content_type(new) != ctype:
        dst.set_override(new, ctype)
    elif ctype is None:
        dst.ensure_default(ext, "application/octet-stream")
    for r in src.rels(part):
        if r.get("TargetMode") == "External":
            rid = _add_rel_raw(dst, new, r)
            continue
        kind = r["Type"].rsplit("/", 1)[-1]
        if kind in MEDIA_TYPES:
            tgt = share_media(src, r["part"], dst)
        else:
            tgt = copy_part(src, r["part"], dst, memo)
        rid = _add_rel_raw(dst, new, {**r, "Target": core.relative(new, tgt)})
        assert rid == r["Id"]
    return new


def _add_rel_raw(pkg: core.Package, part: str, r: dict) -> str:
    root = pkg.rels_root(part, create=True)
    attrs = {"Id": r["Id"], "Type": r["Type"], "Target": r["Target"]}
    if r.get("TargetMode"):
        attrs["TargetMode"] = r["TargetMode"]
    core.ET.SubElement(root, q("pr:Relationship"), attrs)
    pkg.touch(core.rels_name(part))
    return r["Id"]


def share_media(src, part, dst) -> str:
    if src is dst:
        return part
    data = src.parts[part]
    h = core.sha256(data)
    for n, d in dst.parts.items():
        if n.startswith("ppt/media/") and core.sha256(d) == h:
            return n
    ext = part.rsplit(".", 1)[-1]
    new = dst.new_name("ppt/media/image{}." + ext)
    dst.put_bytes(new, data)
    ct = src.content_type(part)
    if ct:
        dst.ensure_default(ext, ct)
    return new


def import_slide(src: core.Package, spart: str, dst: core.Package, layout_part: str) -> str:
    """Clone a slide (same or other package): share media, deep-copy charts/embeds, drop notes."""
    new = dst.new_name("ppt/slides/slide{}.xml")
    dst.put_bytes(new, src.parts[spart])
    dst.set_override(new, core.CT_SLIDE)
    dropped = set()
    memo: dict = {}
    for r in src.rels(spart):
        kind = r["Type"].rsplit("/", 1)[-1]
        if r.get("TargetMode") == "External":
            _add_rel_raw(dst, new, r)
        elif kind == "slideLayout":
            _add_rel_raw(dst, new, {**r, "Target": core.relative(new, layout_part)})
        elif kind in ("notesSlide", "slide", "comments"):
            dropped.add(r["Id"])
        elif kind in MEDIA_TYPES:
            _add_rel_raw(
                dst, new, {**r, "Target": core.relative(new, share_media(src, r["part"], dst))}
            )
        else:
            _add_rel_raw(
                dst, new, {**r, "Target": core.relative(new, copy_part(src, r["part"], dst, memo))}
            )
    if dropped:  # remove links to slides that will not exist in the new deck
        root = dst.xml(new)
        for parent in list(root.iter()):
            for child in list(parent):
                if (
                    core.local(child) in ("hlinkClick", "hlinkHover")
                    and child.get(q("r:id")) in dropped
                ):
                    parent.remove(child)
        dst.touch(new)
    return new


def ensure_notes_master(pkg: core.Package) -> str:
    nm = pkg.related(pkg.main, "notesMaster")
    if nm:
        return nm[0]
    master = pkg.masters()[0]
    theme_src = pkg.related(master, "theme")[0]
    theme = pkg.new_name("ppt/theme/theme{}.xml")
    pkg.put_bytes(theme, pkg.parts[theme_src])
    pkg.set_override(theme, "application/vnd.openxmlformats-officedocument.theme+xml")
    part = pkg.new_name("ppt/notesMasters/notesMaster{}.xml")
    root = core.new_el("p:notesMaster")
    csld = core.sub(root, "p:cSld")
    tree = _sptree(csld)
    _ph_sp(
        tree, 2, "Slide Image Placeholder 1", "sldImg", None, [1143000, 685800, 4572000, 3429000]
    )
    _ph_sp(tree, 3, "Notes Placeholder 2", "body", 1, [685800, 4400550, 5486400, 3600450])
    core.sub(
        root,
        "p:clrMap",
        bg1="lt1",
        tx1="dk1",
        bg2="lt2",
        tx2="dk2",
        accent1="accent1",
        accent2="accent2",
        accent3="accent3",
        accent4="accent4",
        accent5="accent5",
        accent6="accent6",
        hlink="hlink",
        folHlink="folHlink",
    )
    pkg.put_xml(part, root)
    pkg.set_override(
        part, "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
    )
    pkg.add_rel(part, "theme", theme)
    rid = pkg.add_rel(pkg.main, "notesMaster", part)
    pres = pkg.xml(pkg.main)
    lst = core.new_el("p:notesMasterIdLst")
    core.sub(lst, "p:notesMasterId", **{q("r:id"): rid})
    sm = pres.find(q("p:sldMasterIdLst"))
    pres.insert(list(pres).index(sm) + 1, lst)
    pkg.touch(pkg.main)
    return part


def _sptree(csld):
    tree = core.sub(csld, "p:spTree")
    nv = core.sub(tree, "p:nvGrpSpPr")
    core.sub(nv, "p:cNvPr", id="1", name="")
    core.sub(nv, "p:cNvGrpSpPr")
    core.sub(nv, "p:nvPr")
    gp = core.sub(tree, "p:grpSpPr")
    x = core.sub(gp, "a:xfrm")
    core.sub(x, "a:off", x="0", y="0")
    core.sub(x, "a:ext", cx="0", cy="0")
    core.sub(x, "a:chOff", x="0", y="0")
    core.sub(x, "a:chExt", cx="0", cy="0")
    return tree


def _ph_sp(tree, sid, name, ptype, idx, box=None, value=None):
    sp = core.sub(tree, "p:sp")
    nv = core.sub(sp, "p:nvSpPr")
    core.sub(nv, "p:cNvPr", id=str(sid), name=name)
    c = core.sub(nv, "p:cNvSpPr")
    core.sub(c, "a:spLocks", noGrp="1")
    nvpr = core.sub(nv, "p:nvPr")
    attrs = {"type": ptype}
    if idx is not None:
        attrs["idx"] = str(idx)
    core.sub(nvpr, "p:ph", **attrs)
    sppr = core.sub(sp, "p:spPr")
    if box:
        x = core.sub(sppr, "a:xfrm")
        core.sub(x, "a:off", x=str(box[0]), y=str(box[1]))
        core.sub(x, "a:ext", cx=str(box[2]), cy=str(box[3]))
    sp.append(new_txbody(value if value is not None else ""))
    return sp


def set_notes(pkg: core.Package, slide: str, text) -> None:
    value = "\n".join(text) if isinstance(text, list) else str(text)
    existing = pkg.related(slide, "notesSlide")
    if existing:
        root = pkg.xml(existing[0])
        for _p, el, _b in core.iter_shapes(root.find(f"{q('p:cSld')}/{q('p:spTree')}")):
            ph = core.ph_of(el)
            if ph is not None and core.ph_type(ph) == "body":
                tb = core.txbody(el)
                if tb is None:
                    el.append(new_txbody(value))
                else:
                    set_text(tb, value)
                pkg.touch(existing[0])
                return
    nm = ensure_notes_master(pkg)
    part = pkg.new_name("ppt/notesSlides/notesSlide{}.xml")
    root = core.new_el("p:notes")
    tree = _sptree(core.sub(root, "p:cSld"))
    _ph_sp(tree, 2, "Slide Image Placeholder 1", "sldImg", None)
    _ph_sp(tree, 3, "Notes Placeholder 2", "body", 1, value=value)
    ov = core.sub(root, "p:clrMapOvr")
    core.sub(ov, "a:masterClrMapping")
    pkg.put_xml(part, root)
    pkg.set_override(part, core.CT_NOTES)
    pkg.add_rel(part, "notesMaster", nm)
    pkg.add_rel(part, "slide", slide)
    pkg.add_rel(slide, "notesSlide", part)


# --- layout slides -------------------------------------------------------------------


def layout_parts_by_name(pkg: core.Package) -> dict[str, str]:
    """Layout name -> part; names that occur twice (multi-master files) map to None."""
    out: dict[str, str | None] = {}
    for m in pkg.masters():
        for lp in pkg.layouts(m):
            name = pkg.xml(lp).find(q("p:cSld")).get("name", "")
            out[name] = None if name in out else lp
    return out


def add_layout_slide(pkg, layout_part, L, fill, footers: bool) -> str:
    part = pkg.new_name("ppt/slides/slide{}.xml")
    root = core.new_el("p:sld")
    tree = _sptree(core.sub(root, "p:cSld"))
    ov = core.sub(root, "p:clrMapOvr")
    core.sub(ov, "a:masterClrMapping")
    pkg.put_xml(part, root)
    pkg.set_override(part, core.CT_SLIDE)
    pkg.add_rel(part, "slideLayout", layout_part)
    ltree = pkg.xml(layout_part).find(f"{q('p:cSld')}/{q('p:spTree')}")
    lph = {}
    for _p, el, _b in core.iter_shapes(ltree):
        ph = core.ph_of(el)
        if ph is not None:
            lph[(core.ph_type(ph), ph.get("idx", "0"))] = el
    sid = 2
    for ph in L["placeholders"]:
        if ph["key"] not in fill or fill[ph["key"]] is None:
            continue
        val = fill[ph["key"]]
        src = next(
            (
                el
                for (t, i), el in lph.items()
                if (str(ph["idx"]) == i and ph["idx"]) or (not ph["idx"] and t == ph["type"])
            ),
            None,
        )
        name = core.cnv(src).get("name", ph["key"]) if src is not None else ph["key"]
        if isinstance(val, dict) and "image" in val:
            el = _pic_placeholder(pkg, part, sid, name, ph, val["image"])
            crop_to_box(el.find(q("p:blipFill")), Path(val["image"]), ph["box"])
            core.cnv(el).set("descr", val.get("alt") or Path(val["image"]).stem)
            tree.append(el)
        else:
            sp = _ph_sp(tree, sid, name, ph["type"], ph["idx"] if ph["idx"] else None, value=val)
            if ph["type"] == "obj":
                sp.find(f".//{q('p:ph')}").attrib.pop("type")
        sid += 1
    if footers:
        for (t, _i), el in lph.items():
            if t in core.FOOTER_PH:
                cp = core.deepcopy(el)
                core.cnv(cp).set("id", str(sid))
                sid += 1
                tree.append(cp)
    return part


def _pic_placeholder(pkg, part, sid, name, ph, path):
    rid = add_image(pkg, part, path)
    pic = core.new_el("p:pic")
    nv = core.sub(pic, "p:nvPicPr")
    core.sub(nv, "p:cNvPr", id=str(sid), name=name, descr=Path(path).stem)
    c = core.sub(nv, "p:cNvPicPr")
    core.sub(c, "a:picLocks", noGrp="1", noChangeAspect="1")
    nvpr = core.sub(nv, "p:nvPr")
    core.sub(nvpr, "p:ph", type="pic", idx=str(ph["idx"]))
    bf = core.sub(pic, "p:blipFill")
    core.sub(bf, "a:blip", **{q("r:embed"): rid})
    st = core.sub(bf, "a:stretch")
    core.sub(st, "a:fillRect")
    core.sub(pic, "p:spPr")
    return pic


def image_size(path: Path) -> tuple[int, int] | None:
    """(width, height) of a PNG, GIF or JPEG from its header; None if unknown."""
    data = path.read_bytes()[:65536]
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return struct.unpack(">II", data[16:24])
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return struct.unpack("<HH", data[6:10])
    if data[:2] == b"\xff\xd8":
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker, seg = data[i + 1], struct.unpack(">H", data[i + 2 : i + 4])[0]
            if marker in (
                0xC0,
                0xC1,
                0xC2,
                0xC3,
                0xC5,
                0xC6,
                0xC7,
                0xC9,
                0xCA,
                0xCB,
                0xCD,
                0xCE,
                0xCF,
            ):
                h, w = struct.unpack(">HH", data[i + 5 : i + 9])
                return w, h
            i += 2 + seg
    return None


def crop_to_box(blip_fill, path: Path, box) -> None:
    """Centre-crop the picture to the slot's aspect ratio (a:srcRect) instead of stretching it."""
    size = image_size(path)
    if not size or not box or not box[2] or not box[3] or not size[0] or not size[1]:
        return
    ia, ba = size[0] / size[1], box[2] / box[3]
    for old in blip_fill.findall(q("a:srcRect")):
        blip_fill.remove(old)
    attrs = {}
    if abs(ia - ba) > 0.01:
        if ia > ba:
            cut = round((1 - ba / ia) / 2 * 100000)
            attrs = {"l": str(cut), "r": str(cut)}
        else:
            cut = round((1 - ia / ba) / 2 * 100000)
            attrs = {"t": str(cut), "b": str(cut)}
    rect = core.new_el("a:srcRect", **attrs)
    blip = blip_fill.find(q("a:blip"))
    blip_fill.insert(list(blip_fill).index(blip) + 1, rect)


def add_image(pkg, part, path) -> str:
    p = Path(path)
    ext = p.suffix.lower().lstrip(".")
    if ext not in IMAGE_CT:
        raise SpecError(
            f"image {path}: only {sorted(IMAGE_CT)} are supported (convert SVG/EMF to PNG first)"
        )
    if not p.is_file():
        raise SpecError(f"image {path} not found")
    media = pkg.new_name("ppt/media/image{}." + ext)
    pkg.put_bytes(media, p.read_bytes())
    pkg.ensure_default(ext, IMAGE_CT[ext])
    return pkg.add_rel(part, "image", media)


# --- patterns --------------------------------------------------------------------------


def apply_pattern(pkg, part, P, fill, warnings) -> None:
    root = pkg.xml(part)
    tree = root.find(f"{q('p:cSld')}/{q('p:spTree')}")
    resolved = {
        tag: core.find_path(tree, core.parse_path(s["path"])) for tag, s in P["slots"].items()
    }
    members = {
        g: [core.find_path(tree, core.parse_path(m)) for m in R["members"]]
        for g, R in P["repeats"].items()
    }
    deletions = []
    for tag, val in fill.items():
        if tag in P["repeats"]:
            R = P["repeats"][tag]
            for i, el in enumerate(members[tag]):
                if i >= len(val):
                    deletions.append(el)
                    continue
                for s in R["slots"]:
                    sel = core.find_path(el, core.parse_path(s["path"]))
                    v = val[i].get(s["tag"])
                    if v is None:
                        deletions.append(sel)
                    else:
                        set_text(core.txbody(sel), v)
            continue
        S, el = P["slots"][tag], resolved[tag]
        if val is None or val is False:
            deletions.append(el)
        elif S["kind"] in ("text", "bullets"):
            tbs = core.all_txbodies(el)
            if not tbs:
                el.append(new_txbody(val))
            for tb in tbs:
                set_text(tb, val)
        elif S["kind"] == "table":
            fill_table(el, val)
        elif S["kind"] == "image":
            path = val["image"] if isinstance(val, dict) else val
            blip = el.find(f".//{q('a:blip')}")
            blip.set(q("r:embed"), add_image(pkg, part, path))
            crop_to_box(el.find(f".//{q('p:blipFill')}"), Path(path), core.own_box(el))
            alt = val.get("alt") if isinstance(val, dict) else None
            core.cnv(el).set("descr", alt or Path(path).stem)
        elif S["kind"] == "chart":
            fill_chart(pkg, part, el, val, warnings)
    for tag, val in fill.items():  # bars after tables so the grid is final
        if tag in P["slots"] and P["slots"][tag]["kind"] == "bar" and isinstance(val, dict):
            place_bar(resolved[tag], resolved[P["slots"][tag]["table"]], val)
    for el in deletions:
        for parent in tree.iter():
            if el in list(parent):
                parent.remove(el)
                break
    timing = root.find(q("p:timing"))
    if deletions and timing is not None:  # animations may target deleted shapes: repair prompt
        root.remove(timing)
        warnings.append(f"{part}: removed slide animations because shapes were deleted")
    pkg.touch(part)


def fill_table(frame, rows) -> None:
    tbl = frame.find(f".//{q('a:tbl')}")
    trs = tbl.findall(q("a:tr"))
    for tr in trs[len(rows) :]:
        tbl.remove(tr)
    while len(tbl.findall(q("a:tr"))) < len(rows):
        tbl.append(core.deepcopy(tbl.findall(q("a:tr"))[-1]))
    for tr, row in zip(tbl.findall(q("a:tr")), rows, strict=True):
        tcs = tr.findall(q("a:tc"))
        for k, tc in enumerate(tcs):
            tb = tc.find(q("a:txBody"))
            if tb is not None and not tc.get("hMerge") and not tc.get("vMerge"):
                set_text(tb, str(row[k]) if k < len(row) and row[k] is not None else "")
    ext = frame.find(f"{q('p:xfrm')}/{q('a:ext')}")
    ext.set("cy", str(sum(int(tr.get("h", 0)) for tr in tbl.findall(q("a:tr")))))


def place_bar(bar, frame, spec) -> None:
    tbl = frame.find(f".//{q('a:tbl')}")
    cols = [int(g.get("w")) for g in tbl.find(q("a:tblGrid")).findall(q("a:gridCol"))]
    rows = [int(tr.get("h")) for tr in tbl.findall(q("a:tr"))]
    off = frame.find(f"{q('p:xfrm')}/{q('a:off')}")
    tx, ty = int(off.get("x")), int(off.get("y"))
    r, c0, c1 = int(spec["row"]), int(spec["from"]), int(spec["to"])
    if not (1 <= r <= len(rows) and 1 <= c0 <= c1 <= len(cols)):
        raise SpecError(f"bar {spec} outside the {len(rows)}x{len(cols)} table")
    x = core._xfrm(bar)
    bo, be = x.find(q("a:off")), x.find(q("a:ext"))
    h = int(be.get("cy"))
    bo.set("x", str(tx + sum(cols[: c0 - 1])))
    be.set("cx", str(sum(cols[c0 - 1 : c1])))
    bo.set("y", str(ty + sum(rows[: r - 1]) + (rows[r - 1] - h) // 2))


def col_letter(i: int) -> str:
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def xlsx_bytes(cats, series) -> bytes:
    hdr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"

    def esc(v):
        return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    head = "".join(
        f'<c r="{col_letter(k + 1)}1" t="inlineStr"><is><t>{esc(s["name"])}</t></is></c>'
        for k, s in enumerate(series)
    )
    rows = f'<row r="1">{head}</row>'
    for i, cat in enumerate(cats):
        cells = f'<c r="A{i + 2}" t="inlineStr"><is><t>{esc(cat)}</t></is></c>'
        for k, s in enumerate(series):
            v = s["values"][i] if i < len(s["values"]) else None
            if v is not None:
                cells += f'<c r="{col_letter(k + 1)}{i + 2}"><v>{float(v)!r}</v></c>'
        rows += f'<row r="{i + 2}">{cells}</row>'
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            hdr
            + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        )
        z.writestr(
            "_rels/.rels",
            hdr
            + f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}officeDocument" Target="xl/workbook.xml"/></Relationships>',
        )
        z.writestr(
            "xl/workbook.xml",
            hdr
            + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            hdr
            + f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            hdr
            + f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{rows}</sheetData></worksheet>',
        )
    return buf.getvalue()


def fill_chart(pkg, slide, frame, data, warnings) -> None:
    rid = frame.find(f".//{q('c:chart')}").get(q("r:id"))
    cpart = pkg.rel_part(slide, rid)
    root = pkg.xml(cpart)
    cats, series = [str(c) for c in data["categories"]], data["series"]
    sers = list(root.iter(q("c:ser")))
    if len(series) > len(sers):
        raise SpecError(
            f"chart has {len(sers)} series; spec gives {len(series)} (charts cannot gain series)"
        )
    n = len(cats)

    def cache(parent, values, numeric):
        tag = "c:numCache" if numeric else "c:strCache"
        for old in list(parent):
            if core.local(old) in ("numCache", "strCache"):
                parent.remove(old)
        c = core.sub(parent, tag)
        if numeric:
            core.sub(c, "c:formatCode").text = "General"
        core.sub(c, "c:ptCount", val=str(len(values)))
        for i, v in enumerate(values):
            if v is None:
                continue
            pt = core.sub(c, "c:pt", idx=str(i))
            core.sub(pt, "c:v").text = str(v)

    for k, (ser, sd) in enumerate(zip(sers, series, strict=False)):
        col = col_letter(k + 1)
        tx = ser.find(f"{q('c:tx')}/{q('c:strRef')}")
        if tx is not None:
            tx.find(q("c:f")).text = f"Sheet1!${col}$1"
            cache(tx, [sd["name"]], False)
        for tag, vals, numeric, letter in (
            ("c:cat", cats, False, "A"),
            ("c:val", sd["values"], True, col),
        ):
            box = ser.find(q(tag))
            if box is None:
                continue
            ref = (
                box.find(q("c:numRef"))
                if box.find(q("c:numRef")) is not None
                else box.find(q("c:strRef"))
            )
            if ref is None:
                continue
            ref.find(q("c:f")).text = f"Sheet1!${letter}$2:${letter}${n + 1}"
            cache(ref, vals, numeric and core.local(ref) == "numRef")
    for ser in sers[len(series) :]:
        for parent in root.iter():
            if ser in list(parent):
                parent.remove(ser)
                break
    pkg.touch(cpart)
    emb = pkg.related(cpart, "package")
    if emb:
        pkg.put_bytes(emb[0], xlsx_bytes(cats, series))
    else:
        warnings.append(f"chart {cpart} has no embedded workbook; cached values updated only")


# --- spec validation -------------------------------------------------------------------


def validate(spec, contract, base_slides: int | None) -> list[str]:
    errs: list[str] = []
    for k in spec:
        if k not in SPEC_KEYS:
            errs.append(off_contract(k) or f"spec: unknown key {k!r}{did_you_mean(k, SPEC_KEYS)}")
    layouts = {L["name"]: L for L in contract["layouts"]}
    patterns = {P["id"]: P for P in contract["patterns"]}
    slides = spec.get("slides")
    if not isinstance(slides, list) or not slides:
        return [*errs, "spec: 'slides' must be a non-empty list"]
    for n, s in enumerate(slides, 1):
        where = f"slide {n}"
        if not isinstance(s, dict):
            errs.append(f"{where}: must be an object")
            continue
        for k in s:
            if k not in SLIDE_KEYS:
                errs.append(
                    f"{where}: "
                    + (off_contract(k) or f"unknown key {k!r}{did_you_mean(k, SLIDE_KEYS)}")
                )
        modes = [m for m in ("layout", "pattern", "keep") if m in s]
        if len(modes) != 1:
            errs.append(f"{where}: needs exactly one of layout / pattern / keep")
            continue
        fill = s.get("fill", {})
        if not isinstance(fill, dict):
            errs.append(f"{where}: fill must be an object")
            continue
        for k in fill:
            oc = off_contract(k)
            if oc:
                errs.append(f"{where}: {oc}")
        if "layout" in s:
            L = layouts.get(s["layout"])
            if L is None:
                errs.append(
                    f"{where}: layout {s['layout']!r} not in contract{did_you_mean(s['layout'], layouts)}"
                )
                continue
            keys = {p["key"]: p for p in L["placeholders"]}
            for k in fill:
                if k not in keys and not off_contract(k):
                    errs.append(
                        f"{where}: layout {L['name']!r} has no placeholder {k!r}{did_you_mean(k, keys)}"
                    )
            for p in L["placeholders"]:
                if p["prompt_looks_like_content"] and p["key"] not in fill:
                    errs.append(
                        f"{where}: placeholder {p['key']} shows prompt {p['prompt']!r} only in edit view - fill it explicitly"
                    )
            if L["placeholders"] and not any(p["type"] in core.TITLE_PH for p in L["placeholders"]):
                pass
            elif not fill.get("title") and any(p["key"] == "title" for p in L["placeholders"]):
                errs.append(f"{where}: every slide needs a title (accessibility and outline view)")
        elif "pattern" in s:
            P = patterns.get(s["pattern"])
            if P is None:
                errs.append(
                    f"{where}: pattern {s['pattern']!r} not in contract{did_you_mean(s['pattern'], patterns)}"
                )
                continue
            for k, v in fill.items():
                if k in P["repeats"]:
                    R = P["repeats"][k]
                    if not isinstance(v, list) or len(v) > R["count"]:
                        errs.append(
                            f"{where}: repeat group {k} takes a list of at most {R['count']} items (pattern has {R['count']})"
                        )
                        continue
                    inner = {x["tag"] for x in R["slots"]}
                    for item in v:
                        if not isinstance(item, dict):
                            errs.append(f"{where}: {k} items are objects with keys {sorted(inner)}")
                            break
                        for ik in item:
                            if ik not in inner:
                                errs.append(
                                    f"{where}: {k} item key {ik!r} unknown{did_you_mean(ik, inner)}"
                                )
                        for need in inner - set(item):
                            errs.append(
                                f"{where}: {k} item misses {need!r} (give text, or null to delete it)"
                            )
                    continue
                S = P["slots"].get(k)
                if S is None:
                    if not off_contract(k):
                        errs.append(
                            f"{where}: pattern {P['id']!r} has no slot {k!r}{did_you_mean(k, [*P['slots'], *P['repeats']])}"
                        )
                    continue
                if S["kind"] == "table" and v is not None:
                    if not isinstance(v, list) or not all(isinstance(r, list) for r in v):
                        errs.append(f"{where}: {k} is a table: give a list of rows")
                    elif any(len(r) > S["cols"] for r in v):
                        errs.append(f"{where}: {k} has {S['cols']} columns; a row is longer")
                if (
                    S["kind"] == "bar"
                    and v not in (None, False)
                    and not (isinstance(v, dict) and {"row", "from", "to"} <= set(v))
                ):
                    errs.append(
                        f"{where}: {k} is a bar: give {{'row', 'from', 'to'}} (1-based) or null"
                    )
                if (
                    S["kind"] == "chart"
                    and v is not None
                    and not (isinstance(v, dict) and "categories" in v and "series" in v)
                ):
                    errs.append(
                        f"{where}: {k} is a chart: give {{'categories': [...], 'series': [{{'name', 'values'}}]}}"
                    )
            for tag, S in P["slots"].items():
                if tag in fill or S["kind"] not in ("text", "bullets", "table"):
                    continue
                orig = S["text"]
                flat = [c for r in orig for c in r] if S["kind"] == "table" else orig
                if any(str(t).strip() for t in flat):
                    errs.append(
                        f"{where}: slot {tag} still holds sample text {str(flat[0])[:30]!r} - fill it or set null"
                    )
            lay = layouts.get(P["layout"]) or {"placeholders": []}
            for lp_ in lay["placeholders"]:
                if (
                    lp_["prompt_looks_like_content"]
                    and lp_["key"] in P["slots"]
                    and lp_["key"] not in fill
                ):
                    errs.append(
                        f"{where}: slot {lp_['key']} shows prompt {lp_['prompt']!r} only in edit view - fill it explicitly"
                    )
            for g, R in P["repeats"].items():
                if g not in fill:
                    errs.append(
                        f"{where}: repeat group {g} not filled - give 1..{R['count']} items"
                    )
        else:
            if base_slides is None:
                errs.append(f"{where}: 'keep' needs a 'base' deck (edit mode)")
            elif not isinstance(s["keep"], int) or not 1 <= s["keep"] <= base_slides:
                errs.append(
                    f"{where}: keep {s['keep']!r} is not a slide of the base deck (1..{base_slides})"
                )
    if "remove" in spec and base_slides is None:
        errs.append("spec: 'remove' needs a 'base' deck (edit mode)")
    if base_slides is not None:
        keeps = [
            s["keep"] for s in slides if isinstance(s, dict) and isinstance(s.get("keep"), int)
        ]
        remove = spec.get("remove", [])
        if not isinstance(remove, list) or not all(
            isinstance(x, int) and 1 <= x <= base_slides for x in remove
        ):
            errs.append(f"spec: 'remove' must list base slide numbers (1..{base_slides})")
            remove = []
        for x in sorted({k for k in keeps if keeps.count(k) > 1}):
            errs.append(f"spec: base slide {x} is kept twice; a slide can appear once")
        for x in sorted(set(keeps) & set(remove)):
            errs.append(f"spec: base slide {x} is both kept and removed")
        missing = sorted(set(range(1, base_slides + 1)) - set(keeps) - set(remove))
        if missing:
            errs.append(
                f"spec: base slides {missing} are neither kept nor removed - list each as "
                '{"keep": n} where it belongs, or in "remove" to delete it (nothing is dropped implicitly)'
            )
    return errs


def slot_geometry(contract, s, key):
    """(box, size, font, kind) for a fill key, for fit checks."""
    if "layout" in s:
        L = next(L for L in contract["layouts"] if L["name"] == s["layout"])
        p = next((p for p in L["placeholders"] if p["key"] == key), None)
        return (p["box"], p["font_size"], p["font"], p["type"], p.get("spacing")) if p else None
    if "pattern" in s:
        P = next(P for P in contract["patterns"] if P["id"] == s["pattern"])
        S = P["slots"].get(key)
        if S and S["kind"] in ("text", "bullets"):
            return (S["box"], S["font_size"], S["font"], S["kind"], S.get("spacing"))
    return None


def fit_errors(contract, slides) -> list[dict]:
    out = []
    for n, s in enumerate(slides, 1):
        for key, val in (s.get("fill") or {}).items():
            if (
                val is None
                or not isinstance(val, (str, list))
                or (isinstance(val, list) and val and isinstance(val[0], list))
            ):
                continue
            g = slot_geometry(contract, s, key)
            if not g or not g[0]:
                continue
            paras = [t for t, _lvl in items(val)]
            indent = 342900 if g[3] in ("body", "obj") and isinstance(val, list) else 0
            before, line = g[4] or (0.0, 1.0)
            r = fit.check(
                paras,
                g[0],
                g[1],
                g[2],
                indent_emu=indent,
                line_spacing=1.2 * line,
                para_space_pt=before,
            )
            if not r["fits"]:
                out.append({"slide": n, "slot": key, **r, "size": g[1], "font": g[2]})
    return out


def split_overflow(contract, slides):
    """allow_split: move list items that do not fit into continuation slides."""
    out = []
    for s in slides:
        if not s.get("allow_split"):
            out.append(s)
            continue
        cur = s
        for _guard in range(20):
            errs = fit_errors(contract, [cur])
            lists = [
                e
                for e in errs
                if isinstance(cur["fill"].get(e["slot"]), list) and len(cur["fill"][e["slot"]]) > 1
            ]
            if not lists:
                break
            key = lists[0]["slot"]
            vals = cur["fill"][key]
            k = len(vals) - 1
            while k > 1 and fit_errors(contract, [{**cur, "fill": {**cur["fill"], key: vals[:k]}}]):
                k -= 1
            head = {**cur, "fill": {**cur["fill"], key: vals[:k]}}
            out.append(head)
            title = cur["fill"].get("title", "")
            cont_title = title if title.endswith("(cont.)") else f"{title} (cont.)"
            cur = {**cur, "fill": {**cur["fill"], key: vals[k:], "title": cont_title}}
        out.append(cur)
    return out


# --- driver -------------------------------------------------------------------------------


def smoke_spec(contract) -> dict:
    """One slide per layout and pattern with dummy text: proves a contract is buildable."""
    slides = []
    for L in contract["layouts"]:
        fill = {}
        for p in L["placeholders"]:
            if p["type"] in ("pic", "tbl", "chart", "dgm", "media", "clipArt"):
                continue
            lines = (p.get("capacity") or [1, 10])[0]
            if p["prompt_looks_like_content"]:
                fill[p["key"]] = p["prompt"]
            elif p["key"] == "title" or lines < 2:
                fill[p["key"]] = "Sample"
            else:
                fill[p["key"]] = ["Point one", "Point two"]
        slides.append(
            {"layout": L["name"], "fill": fill, "notes": f"Smoke test of layout {L['name']}."}
        )
    for P in contract["patterns"]:
        fill = {}
        for tag, S in P["slots"].items():
            if S["kind"] in ("text", "bullets"):
                fill[tag] = "Sample" if (S.get("capacity") or [1, 9])[1] > 8 else "01"
            elif S["kind"] == "table":
                fill[tag] = [[f"r{r}c{c}" for c in range(S["cols"])] for r in range(S["rows"])]
        for g, R in P["repeats"].items():
            fill[g] = [{x["tag"]: f"Item {i + 1}" for x in R["slots"]} for i in range(R["count"])]
        slides.append(
            {"pattern": P["id"], "fill": fill, "notes": f"Smoke test of pattern {P['id']}."}
        )
    return {"contract": "CONTRACT", "output": "smoke.pptx", "slides": slides}


def rebuild_slide_list(pkg, order, keep_ids, create_mode):
    pres = pkg.xml(pkg.main)
    lst = pres.find(q("p:sldIdLst"))
    if lst is None:
        lst = core.new_el("p:sldIdLst")
        anchor = pres.find(q("p:notesMasterIdLst"))
        if anchor is None:
            anchor = pres.find(q("p:sldMasterIdLst"))
        pres.insert(list(pres).index(anchor) + 1, lst)
    old = {p: (int(s.get("id")), rid) for s, rid, p in pkg.slide_ids()}
    next_id = max([255, *[i for i, _r in old.values()]]) + 1
    for el in list(lst):
        lst.remove(el)
    new_ids = {}
    for part in order:
        if part in old and part in keep_ids:
            sid, rid = old[part]
        else:
            sid, rid = next_id, pkg.add_rel(pkg.main, "slide", part)
            next_id += 1
        core.sub(lst, "p:sldId", id=str(sid), **{q("r:id"): rid})
        new_ids[part] = sid
    for part, (_sid, rid) in old.items():
        if part not in order:
            pkg.drop_rel(pkg.main, rid)
    shows = pres.find(q("p:custShowLst"))
    if shows is not None:  # custom shows reference slides through presentation rels
        live = {r["Id"] for r in pkg.rels(pkg.main)}
        for show in list(shows):
            for sld in list(show.iter(q("p:sld"))):
                if sld.get(q("r:id")) not in live:
                    for parent in show.iter():
                        if sld in list(parent):
                            parent.remove(sld)
            if not list(show.iter(q("p:sld"))):
                shows.remove(show)
        if not len(shows):
            pres.remove(shows)
    pkg.touch(pkg.main)
    _sections(pkg, pres, order, new_ids, create_mode)


def _sections(pkg, pres, order, new_ids, create_mode):
    ext_lst = pres.find(q("p:extLst"))
    if ext_lst is None:
        return
    for ext in list(ext_lst):
        if ext.get("uri") != core.SECTION_URI:
            continue
        if create_mode:
            ext_lst.remove(ext)
            continue
        secs = list(ext.iter(q("p14:section")))
        live = set(new_ids.values())
        owner = {}
        for sec in secs:
            lst = sec.find(q("p14:sldIdLst"))
            for s in list(lst):
                if int(s.get("id")) not in live:
                    lst.remove(s)
                else:
                    owner[int(s.get("id"))] = sec
        prev = None
        for part in order:
            sid = new_ids[part]
            if sid not in owner:
                sec = owner.get(prev, secs[0] if secs else None)
                if sec is None:
                    continue
                lst = sec.find(q("p14:sldIdLst"))
                kids = list(lst)
                pos = next(
                    (i + 1 for i, k in enumerate(kids) if int(k.get("id")) == prev), len(kids)
                )
                el = core.new_el("p14:sldId", id=str(sid))
                lst.insert(pos, el)
                owner[sid] = sec
            prev = sid
        # PowerPoint requires each slide in exactly one section and sections in slide order
        flat = [int(s.get("id")) for sec in secs for s in sec.find(q("p14:sldIdLst"))]
        if flat != [new_ids[p] for p in order]:
            ext_lst.remove(ext)
    if not len(ext_lst):
        pres.remove(ext_lst)


def build(
    spec_path: Path, out: Path | None, contract_path: str | None, only: list[int] | None
) -> int:
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    base_dir = spec_path.parent
    cpath = Path(contract_path or spec.get("contract") or "")
    if not cpath.is_absolute():
        cpath = (Path.cwd() / cpath) if contract_path else (base_dir / cpath)
    if not cpath.is_file():
        print(
            f"error: contract not found: {cpath} (run the pptx-style-profile skill first)",
            file=sys.stderr,
        )
        return 2
    contract = core.load_contract(cpath)
    tpath = Path(spec.get("template") or contract["source"].get("path") or "")
    if spec.get("template") and not tpath.is_absolute():
        tpath = base_dir / tpath
    if not tpath.is_file():
        print(f"error: template not found: {tpath} (set 'template' in the spec)", file=sys.stderr)
        return 2
    warnings: list[str] = []
    if core.sha256(tpath.read_bytes()) != contract["source"]["sha256"]:
        fresh = core.profile(tpath)
        d = core.diff_contracts(contract, fresh)
        if d["breaking"]:
            print("error: the template changed since it was profiled - BREAKING:", file=sys.stderr)
            for line in d["breaking"]:
                print(f"  - {line}", file=sys.stderr)
            print(
                "re-profile the template (pptx-style-profile) and update the spec", file=sys.stderr
            )
            return 3
        warnings += [
            f"template drift (non-breaking): {x}" for x in d["changed"] + d["additive"]
        ] or ["template bytes changed (re-saved) but the contract is unchanged"]
        contract = {**fresh, "source": {**fresh["source"], "path": str(tpath)}}
    base = None
    if spec.get("base"):
        base = Path(spec["base"])
        if not base.is_absolute():
            base = base_dir / base
        if not base.is_file():
            print(f"error: base deck not found: {base}", file=sys.stderr)
            return 2
    if out is None:
        out = Path(spec.get("output") or "deck.pptx")
        out = out if out.is_absolute() else base_dir / out
    if only:
        out = out.with_name(out.stem + ".preview" + out.suffix)
    if base and out.resolve() == base.resolve():
        print(
            f"error: refusing to overwrite the base deck {base}; set a different 'output'",
            file=sys.stderr,
        )
        return 2
    target = core.Package(base or tpath)
    src = core.Package(tpath) if base else target
    if base:  # the base deck must be built on the contract's template
        btheme = core.theme_of(target, target.masters()[0])
        if (
            btheme["colors"] != contract["theme"]["colors"]
            or btheme["fonts"] != contract["theme"]["fonts"]
        ):
            print(
                f"error: {base.name} does not use this contract's theme - it was not built on this template;"
                " profile its real template first",
                file=sys.stderr,
            )
            return 2
    base_slides = target.slides() if base else None
    errs = validate(spec, contract, len(base_slides) if base_slides is not None else None)
    if errs:
        print(f"spec rejected ({len(errs)} problems, nothing written):", file=sys.stderr)
        for e in errs:
            print(f"  - {e}", file=sys.stderr)
        return 2
    patterns_by_id = {P["id"]: P for P in contract["patterns"]}
    slides = spec["slides"]
    if only:
        slides = [s for i, s in enumerate(slides, 1) if i in only]
    slides = split_overflow(contract, slides)
    lay_target = layout_parts_by_name(target)
    used_names = {s["layout"] for s in slides if "layout" in s} | {
        patterns_by_id[s["pattern"]]["layout"] for s in slides if "pattern" in s
    }
    ambiguous = sorted(n for n in used_names if n in lay_target and lay_target[n] is None)
    if ambiguous:
        print(
            f"error: layout names {ambiguous} exist on several masters; rename them in the template",
            file=sys.stderr,
        )
        return 2
    if base:
        missing = [L["name"] for L in contract["layouts"] if L["name"] not in lay_target]
        used = {s.get("layout") for s in slides} | {
            next(P["layout"] for P in contract["patterns"] if P["id"] == s["pattern"])
            for s in slides
            if "pattern" in s
        }
        bad = sorted(set(missing) & used)
        if bad:
            print(
                f"error: base deck lacks layouts {bad} - it is not built on this template",
                file=sys.stderr,
            )
            return 2
    layouts = {L["name"]: L for L in contract["layouts"]}
    patterns = {P["id"]: P for P in contract["patterns"]}
    footers = bool(contract["derived"].get("footer_placeholders_on_slides"))
    order, report, keep_ids = [], [], set()
    before = {p: core.sha256(target.parts[p]) for p in (base_slides or [])}
    base_src = {}
    if base and base.with_suffix(".build.json").exists():  # inherit provenance of kept slides
        base_src = {
            r["part"]: r["source"]
            for r in json.loads(base.with_suffix(".build.json").read_text())["slides"]
        }
    for n, s in enumerate(slides, 1):
        fill = s.get("fill") or {}
        try:
            if "keep" in s:
                part = base_slides[s["keep"] - 1]
                keep_ids.add(part)
                status = "unchanged"
                if fill or s.get("replace"):
                    L = _layout_of(target, part, layouts)
                    _patch_keep(target, part, L, fill, s.get("replace") or {})
                    status = "patched"
                src_desc = base_src.get(part) or f"keep:{s['keep']}"
            elif "layout" in s:
                part = add_layout_slide(
                    target, lay_target[s["layout"]], layouts[s["layout"]], fill, footers
                )
                status, src_desc = "added", f"layout:{s['layout']}"
            else:
                P = patterns[s["pattern"]]
                part = import_slide(src, P["part"], target, lay_target[P["layout"]])
                apply_pattern(target, part, P, fill, warnings)
                status, src_desc = "added", f"pattern:{P['id']}"
            if s.get("notes"):
                set_notes(target, part, s["notes"])
                if status == "unchanged":
                    status = "patched"
        except SpecError as exc:
            print(f"error: slide {n}: {exc}", file=sys.stderr)
            return 2
        order.append(part)
        report.append({"slide": n, "source": src_desc, "status": status, "part": part})
    rebuild_slide_list(target, order, keep_ids, create_mode=base is None)
    target.gc()
    if out.suffix.lower() == ".pptx" and target.content_type(target.main) == core.CT_TMPL:
        target.set_override(target.main, core.CT_PRES)
    out.parent.mkdir(parents=True, exist_ok=True)
    target.save(out)
    # verification: re-open, kept slides byte-identical, lint
    check = core.Package(out)
    identical = all(
        core.sha256(check.parts[r["part"]]) == before[r["part"]]
        for r in report
        if r["status"] == "unchanged"
    )
    counts = {
        k: sum(1 for r in report if r["status"] == k) for k in ("unchanged", "patched", "added")
    }
    counts["removed"] = len([p for p in (base_slides or []) if p not in order])
    import lint as linter

    findings = linter.lint(
        out, contract, require_notes=spec.get("require_notes", True), build_report=report
    )
    errors = [f for f in findings if f["severity"] == "error"]
    rep = {
        "output": str(out),
        "template": str(tpath),
        "contract": str(cpath),
        "preview": bool(only),
        "slides": report,
        "counts": counts,
        "unchanged_slides_byte_identical": identical,
        "warnings": warnings,
        "lint": findings,
    }
    out.with_suffix(".build.json").write_text(json.dumps(rep, indent=1), encoding="utf-8")
    print(f"built {out} ({len(order)} slides{', PREVIEW' if only else ''}): {counts}")
    if base:
        print(f"unchanged slides byte-identical: {identical}")
    for w in warnings:
        print(f"warning: {w}")
    print(linter.format_findings(findings))
    return 1 if errors or not identical else 0


def _layout_of(pkg, part, layouts):
    lp = pkg.related(part, "slideLayout")[0]
    name = pkg.xml(lp).find(q("p:cSld")).get("name", "")
    return layouts.get(name)


def _patch_keep(pkg, part, L, fill, replace) -> None:
    root = pkg.xml(part)
    tree = root.find(f"{q('p:cSld')}/{q('p:spTree')}")
    keys = {p["key"]: p for p in (L or {}).get("placeholders", [])}
    for key, val in fill.items():
        p = keys.get(key)
        if p is None:
            raise SpecError(
                f"keep-slide fill key {key!r} is not a placeholder of its layout{did_you_mean(key, keys)}; use 'replace' for pattern slides"
            )
        hit = None
        for _p, el, _b in core.iter_shapes(tree):
            ph = core.ph_of(el)
            if ph is None:
                continue
            same = (p["idx"] and int(ph.get("idx", "0")) == p["idx"]) or (
                not p["idx"] and core.ph_type(ph) in core.TITLE_PH and p["type"] in core.TITLE_PH
            )
            if same:
                hit = el
                break
        if hit is None:
            raise SpecError(f"placeholder {key!r} is not on this slide")
        tb = core.txbody(hit)
        if tb is None:
            hit.append(new_txbody(val))
        else:
            set_text(tb, val)
    if replace and not replace_runs(root, replace):
        raise SpecError(f"replace: none of {list(replace)} found on the slide")
    pkg.touch(part)


def self_test() -> int:
    tb = new_txbody(["one", {"text": "two", "level": 1}])
    assert [para_text(p) for p in tb.findall(q("a:p"))] == ["one", "two"]
    p = tb.findall(q("a:p"))[0]
    r = p.find(q("a:r"))
    r.find(q("a:rPr")).set("b", "1")
    set_text(tb, "uno")
    assert len(tb.findall(q("a:p"))) == 1 and tb.find(f".//{q('a:rPr')}").get("b") == "1"
    assert replace_runs(tb, {"uno": "eins"}) == 1 and para_text(tb.find(q("a:p"))) == "eins"
    assert off_contract("font_color") and off_contract("x") and not off_contract("title")
    assert col_letter(0) == "A" and col_letter(26) == "AA"
    assert (
        zipfile.ZipFile(io.BytesIO(xlsx_bytes(["a"], [{"name": "s", "values": [1]}]))).testzip()
        is None
    )
    assert number_pattern("ppt/charts/chart12.xml") == "ppt/charts/chart{}.xml"
    print("self-test OK")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("spec", nargs="?")
    ap.add_argument("--out")
    ap.add_argument("--contract")
    ap.add_argument("--slides", help="preview: build only these 1-based spec slides, e.g. 1,4")
    ap.add_argument(
        "--smoke-spec", metavar="CONTRACT", help="print a spec exercising every layout and pattern"
    )
    ap.add_argument("--self-test", action="store_true")
    a = ap.parse_args(argv)
    if a.self_test:
        return self_test()
    if a.smoke_spec:
        sp = smoke_spec(core.load_contract(a.smoke_spec))
        sp["contract"] = str(Path(a.smoke_spec).resolve())
        print(json.dumps(sp, indent=1))
        return 0
    if not a.spec:
        ap.print_usage()
        return 2
    only = [int(x) for x in a.slides.split(",")] if a.slides else None
    try:
        return build(Path(a.spec), Path(a.out) if a.out else None, a.contract, only)
    except (ValueError, KeyError, json.JSONDecodeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
