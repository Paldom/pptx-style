#!/usr/bin/env python3
"""Generate brand-free PowerPoint test templates from scratch (stdlib only).

    python3 tests/fixtures/make_fixtures.py OUTDIR

Writes four templates, every colour, font and word invented here:
  clean.pptx   16:9, custom theme (Lato), 7 layouts, logo + slide-number chrome on the
               master, sample slides: cards (groups), timeline (table + bars),
               comparison (table + rating marks), KPI row, bullets.
  messy.potx   real-world traits: duplicate and localized/export shape names, art as a
               layout <p:bg> image and a layout panel picture, agenda numbers that exist
               only as placeholder prompt text, a decorative slide with no placeholders,
               page number as a slidenum field in a text box, an off-theme colour, an
               icon-bank appendix, a non-English layout name, sections, no notes master.
  drift.potx   messy v2: a renamed layout, a body placeholder moved +40 px, a deleted
               layout and a changed accent2.
  charts.pptx  4:3, different theme with a font that is not installed anywhere (exercises
               the APPROXIMATE path), a native chart sample with an embedded workbook.
"""

from __future__ import annotations

import io
import struct
import sys
import zipfile
import zlib
from pathlib import Path
from xml.sax.saxutils import escape

NSDECL = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
)
HDR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT = "application/vnd.openxmlformats-officedocument."
EMU_PX = 9525


def png(rgb: tuple[int, int, int], w: int = 8, h: int = 8) -> bytes:
    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))

    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d))

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


# --- shape builders ---------------------------------------------------------


def ph_xml(ph):
    if not ph:
        return "<p:nvPr/>"
    t, idx, custom = (*ph, None, None)[:3]
    a = f' type="{t}"' if t and t != "obj" else ""
    a += f' idx="{idx}"' if idx is not None else ""
    a += ' hasCustomPrompt="1"' if custom else ""
    return f"<p:nvPr><p:ph{a}/></p:nvPr>"


def runs(text, sz=None, color=None, font=None, bold=False):
    attrs = ' lang="en-US"' + (f' sz="{sz}"' if sz else "") + (' b="1"' if bold else "")
    inner = ""
    if color:
        inner += (
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            if len(color) == 6 and color.upper() == color
            else f'<a:solidFill><a:schemeClr val="{color}"/></a:solidFill>'
        )
    if font:
        inner += f'<a:latin typeface="{font}"/>'
    rpr = f'<a:rPr{attrs} dirty="0">{inner}</a:rPr>' if inner else f'<a:rPr{attrs} dirty="0"/>'
    return f"<a:r>{rpr}<a:t>{escape(text)}</a:t></a:r>"


def body(
    paras=None, *, sz=None, color=None, font=None, bold=False, anchor=None, fld=None, nobullet=False
):
    bp = f'<a:bodyPr anchor="{anchor}"/>' if anchor else "<a:bodyPr/>"
    ps = ""
    for t in paras or []:
        ps += (
            f"<a:p>{runs(t, sz, color, font, bold)}</a:p>"
            if t
            else '<a:p><a:endParaRPr lang="en-US"/></a:p>'
        )
    if fld:
        ps += (
            '<a:p><a:pPr algn="r"/><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" '
            f'type="slidenum"><a:rPr lang="en-US" sz="{sz or 1000}"/><a:t>\u2039#\u203a</a:t></a:fld></a:p>'
        )
    if not ps:
        ps = '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    lst = (
        '<a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:buNone/></a:lvl1pPr></a:lstStyle>'
        if nobullet
        else "<a:lstStyle/>"
    )
    return f"<p:txBody>{bp}{lst}{ps}</p:txBody>"


def xfrm(box):
    x, y, w, h = box
    return f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'


def sp(
    i, name, box=None, *, ph=None, text=None, txbox=False, geom="rect", fill=None, line=None, **tk
):
    cnv = (
        '<p:cNvSpPr txBox="1"/>'
        if txbox
        else ('<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' if ph else "<p:cNvSpPr/>")
    )
    sppr = ""
    if box:
        sppr += xfrm(box) + f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>'
    if fill:
        sppr += (
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            if fill.upper() == fill and len(fill) == 6
            else f'<a:solidFill><a:schemeClr val="{fill}"/></a:solidFill>'
        )
    elif txbox:
        sppr += "<a:noFill/>"
    if line == "none":
        sppr += "<a:ln><a:noFill/></a:ln>"
    tb = body(text, **tk) if (text is not None or ph or txbox or tk.get("fld")) else ""
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="{escape(name)}"/>{cnv}{ph_xml(ph)}</p:nvSpPr>'
        f"<p:spPr>{sppr}</p:spPr>{tb}</p:sp>"
    )


def pic(i, name, box, rid, ph=None):
    return (
        f'<p:pic><p:nvPicPr><p:cNvPr id="{i}" name="{escape(name)}" descr="decorative"/>'
        f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>{ph_xml(ph)}</p:nvPicPr>'
        f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        f'<p:spPr>{xfrm(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )


def grp(i, name, box, children):
    x, y, w, h = box
    return (
        f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{i}" name="{escape(name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        f'<p:grpSpPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/>'
        f'<a:chOff x="{x}" y="{y}"/><a:chExt cx="{w}" cy="{h}"/></a:xfrm></p:grpSpPr>{"".join(children)}</p:grpSp>'
    )


def table(i, name, box, rows, sz=1200):
    x, y, w, h = box
    ncol = len(rows[0])
    cw = w // ncol
    rh = h // len(rows)
    grid = "".join(f'<a:gridCol w="{cw}"/>' for _ in range(ncol))
    trs = ""
    for r, row in enumerate(rows):
        tcs = "".join(
            f"<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>{runs(c, sz, bold=(r == 0))}</a:p></a:txBody><a:tcPr/></a:tc>"
            for c in row
        )
        trs += f'<a:tr h="{rh}">{tcs}</a:tr>'
    return (
        f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{i}" name="{escape(name)}"/>'
        '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
        f'<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cw * ncol}" cy="{rh * len(rows)}"/></p:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        f'<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>{grid}</a:tblGrid>{trs}</a:tbl>'
        "</a:graphicData></a:graphic></p:graphicFrame>"
    )


def chart_frame(i, name, box, rid):
    x, y, w, h = box
    return (
        f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{i}" name="{escape(name)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
        f'<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="{rid}"/>'
        "</a:graphicData></a:graphic></p:graphicFrame>"
    )


def sptree(shapes):
    return (
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/>'
        f'<a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>{"".join(shapes)}</p:spTree>'
    )


def bg_xml(bg):
    if not bg:
        return ""
    if bg.startswith("rId"):
        return f'<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="{bg}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>'
    return f'<p:bg><p:bgPr><a:solidFill><a:schemeClr val="{bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'


# --- theme ---------------------------------------------------------------------


def theme(name, colors, major, minor):
    cs = "".join(
        f'<a:{k}><a:srgbClr val="{v}"/></a:{k}>'
        for k, v in zip(
            [
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
            ],
            colors,
            strict=True,
        )
    )
    fill3 = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    ln3 = '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' * 3
    eff3 = "<a:effectStyle><a:effectLst/></a:effectStyle>" * 3
    return (
        HDR
        + f'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="{name}"><a:themeElements>'
        f'<a:clrScheme name="{name}">{cs}</a:clrScheme>'
        f'<a:fontScheme name="{name}"><a:majorFont><a:latin typeface="{major}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        f'<a:minorFont><a:latin typeface="{minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
        f'<a:fmtScheme name="{name}"><a:fillStyleLst>{fill3}</a:fillStyleLst><a:lnStyleLst>{ln3}</a:lnStyleLst>'
        f"<a:effectStyleLst>{eff3}</a:effectStyleLst><a:bgFillStyleLst>{fill3}</a:bgFillStyleLst></a:fmtScheme>"
        "</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"
    )


TXSTYLES = (
    "<p:txStyles>"
    '<p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="3200" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
    '<a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
    '<p:bodyStyle><a:lvl1pPr marL="228600" indent="-228600"><a:spcBef><a:spcPts val="600"/></a:spcBef>'
    '<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/><a:defRPr sz="2000"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
    '<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr>'
    '<a:lvl2pPr marL="685800" indent="-228600"><a:buFont typeface="Arial"/><a:buChar char="&#8211;"/>'
    '<a:defRPr sz="1800"><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl2pPr></p:bodyStyle>'
    '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:otherStyle>'
    "</p:txStyles>"
)


# --- package assembly -------------------------------------------------------


class Deck:
    def __init__(self, w, h, kind="pptx"):
        self.w, self.h, self.kind = w, h, kind
        self.parts: dict[str, bytes] = {}
        self.rels: dict[str, list[tuple[str, str, str]]] = {}
        self.overrides: dict[str, str] = {}
        self.layouts: list[str] = []
        self.slides: list[str] = []
        self.sections: list[tuple[str, list[int]]] = []
        self.notes_master = False

    def rel(self, src, rtype, target):
        lst = self.rels.setdefault(src, [])
        rid = f"rId{len(lst) + 1}"
        lst.append((rid, RT + rtype if "://" not in rtype else rtype, target))
        return rid

    def media(self, src, rgb):
        name = f"ppt/media/image{sum(1 for p in self.parts if p.startswith('ppt/media/')) + 1}.png"
        self.parts[name] = png(rgb)
        return self.rel(src, "image", "../media/" + name.rsplit("/", 1)[1])

    def add(self, name, xml, ctype):
        self.parts[name] = xml.encode("utf-8") if isinstance(xml, str) else xml
        if ctype:
            self.overrides[name] = ctype

    def save(self, path):
        main = CT + (
            "presentationml.template.main+xml"
            if self.kind == "potx"
            else "presentationml.presentation.main+xml"
        )
        self.overrides["ppt/presentation.xml"] = main
        ct = (
            HDR + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
            '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
            + "".join(
                f'<Override PartName="/{n}" ContentType="{c}"/>' for n, c in self.overrides.items()
            )
            + "</Types>"
        )
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            for n, data in self.parts.items():
                z.writestr(n, data)
            for src, lst in self.rels.items():
                d, b = src.rsplit("/", 1) if "/" in src else ("", src)
                rn = f"{d}/_rels/{b}.rels" if src else "_rels/.rels"
                body_ = "".join(
                    f'<Relationship Id="{i}" Type="{t}" Target="{tg}"/>' for i, t, tg in lst
                )
                z.writestr(
                    rn,
                    HDR
                    + f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body_}</Relationships>',
                )


def build(spec) -> Deck:
    d = Deck(spec["w"], spec["h"], spec.get("kind", "pptx"))
    W, H = d.w, d.h
    d.rel("", "officeDocument", "ppt/presentation.xml")
    d.rel(
        "",
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        "docProps/core.xml",
    )
    d.rel("", "extended-properties", "docProps/app.xml")
    d.add(
        "docProps/core.xml",
        HDR
        + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Fixture</dc:title><dc:creator>pptx-style tests</dc:creator></cp:coreProperties>',
        "application/vnd.openxmlformats-package.core-properties+xml",
    )
    d.add(
        "docProps/app.xml",
        HDR
        + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>pptx-style fixtures</Application></Properties>',
        CT + "extended-properties+xml",
    )
    d.add("ppt/theme/theme1.xml", theme(*spec["theme"]), CT + "theme+xml")
    # master
    M = "ppt/slideMasters/slideMaster1.xml"
    m_shapes = [
        sp(2, "Title Placeholder 1", [W // 16, H // 16, W * 7 // 8, H // 8], ph=("title",)),
        sp(
            3,
            "Text Placeholder 2",
            [W // 16, H // 4, W * 7 // 8, H * 9 // 16],
            ph=("body", 1),
            text=["Master text"],
        ),
    ]
    for kind_, extra in spec["chrome"]:
        if kind_ == "logo":
            rid = d.media(M, (40, 40, 40))
            m_shapes.append(
                pic(
                    10,
                    extra.get("name", "Logo"),
                    [W // 16, H * 15 // 16 - 60000, W // 12, H // 32],
                    rid,
                )
            )
        elif kind_ == "sldnum_ph":
            m_shapes.append(
                sp(
                    11,
                    "Slide Number Placeholder 5",
                    [W * 7 // 8, H * 15 // 16 - 60000, W // 16, H // 32],
                    ph=("sldNum", 12),
                    fld=True,
                    sz=1000,
                )
            )
        elif kind_ == "sldnum_box":
            m_shapes.append(
                sp(
                    12,
                    extra.get("name", "TextBox 7"),
                    [W * 7 // 8, H * 15 // 16 - 60000, W // 16, H // 32],
                    txbox=True,
                    fld=True,
                    sz=1000,
                )
            )
        elif kind_ == "rule":
            m_shapes.append(
                sp(
                    13,
                    "Footer Rule",
                    [W // 16, H * 29 // 32, W * 7 // 8, 12700],
                    fill="accent1",
                    line="none",
                )
            )
    d.rel(M, "theme", "../theme/theme1.xml")
    lids = []
    for n, L in enumerate(spec["layouts"], 1):
        lp = f"ppt/slideLayouts/slideLayout{n}.xml"
        d.rel(lp, "slideMaster", "../slideMasters/slideMaster1.xml")
        shapes = []
        bg = L.get("bg")
        if bg == "image":
            bg = d.media(lp, (210, 220, 235))
        for k, (pt, idx, box, prompt, custom) in enumerate(L["phs"], 2):
            plain = pt == "subTitle" or (pt == "body" and (L.get("type") == "secHead" or custom))
            shapes.append(
                sp(
                    k,
                    f"{pt} {k}",
                    box,
                    ph=(pt, idx, custom),
                    text=[prompt] if prompt else None,
                    nobullet=plain,
                )
            )
        for k, (nm, box) in enumerate(L.get("pics", []), 40):
            shapes.append(pic(k, nm, box, d.media(lp, (190, 205, 225))))
        for k, (nm, box, txt) in enumerate(L.get("textboxes", []), 60):
            shapes.append(sp(k, nm, box, txbox=True, text=[txt], sz=1200))
        lt = f' type="{L["type"]}"' if L.get("type") else ""
        d.add(
            lp,
            HDR
            + f'<p:sldLayout {NSDECL}{lt} preserve="1"><p:cSld name="{escape(L["name"])}">{bg_xml(bg)}{sptree(shapes)}</p:cSld>'
            "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>",
            CT + "presentationml.slideLayout+xml",
        )
        lids.append(d.rel(M, "slideLayout", f"../slideLayouts/slideLayout{n}.xml"))
        d.layouts.append(lp)
    lid_xml = "".join(
        f'<p:sldLayoutId id="{2147483649 + i}" r:id="{r}"/>' for i, r in enumerate(lids)
    )
    d.add(
        M,
        HDR + f"<p:sldMaster {NSDECL}><p:cSld>{bg_xml('bg1')}{sptree(m_shapes)}</p:cSld>"
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
        'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        f"<p:sldLayoutIdLst>{lid_xml}</p:sldLayoutIdLst>{TXSTYLES}</p:sldMaster>",
        CT + "presentationml.slideMaster+xml",
    )
    # notes master (optional)
    P = "ppt/presentation.xml"
    pres_rels = [d.rel(P, "slideMaster", "slideMasters/slideMaster1.xml")]
    nm_xml = ""
    if spec.get("notes_master"):
        d.add("ppt/theme/theme2.xml", theme(*spec["theme"]), CT + "theme+xml")
        NMP = "ppt/notesMasters/notesMaster1.xml"
        d.rel(NMP, "theme", "../theme/theme2.xml")
        d.add(
            NMP,
            HDR
            + f"<p:notesMaster {NSDECL}><p:cSld>{sptree([sp(2, 'Slide Image Placeholder 1', [381000, 685800, 6096000, 3429000], ph=('sldImg', None)), sp(3, 'Notes Placeholder 2', [685800, 4343400, 5486400, 4114800], ph=('body', 1))])}</p:cSld>"
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
            'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:notesMaster>',
            CT + "presentationml.notesMaster+xml",
        )
        nm_xml = f'<p:notesMasterIdLst><p:notesMasterId r:id="{d.rel(P, "notesMaster", "notesMasters/notesMaster1.xml")}"/></p:notesMasterIdLst>'
    # slides
    sld_ids = []
    for n, S in enumerate(spec["slides"], 1):
        spart = f"ppt/slides/slide{n}.xml"
        d.rel(spart, "slideLayout", f"../slideLayouts/slideLayout{S['layout']}.xml")
        shapes = S["shapes"](d, spart, W, H) if callable(S["shapes"]) else S["shapes"]
        d.add(
            spart,
            HDR + f"<p:sld {NSDECL}{S.get('root_extra', '')}><p:cSld>{sptree(shapes)}</p:cSld>"
            f"<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>{S.get('tail', '')}</p:sld>",
            CT + "presentationml.slide+xml",
        )
        sld_ids.append((255 + n, d.rel(P, "slide", f"slides/slide{n}.xml")))
        d.slides.append(spart)
        if S.get("notes") and spec.get("notes_master"):
            npart = f"ppt/notesSlides/notesSlide{n}.xml"
            d.rel(npart, "notesMaster", "../notesMasters/notesMaster1.xml")
            d.rel(npart, "slide", f"../slides/slide{n}.xml")
            d.rel(spart, "notesSlide", f"../notesSlides/notesSlide{n}.xml")
            d.add(
                npart,
                HDR
                + f"<p:notes {NSDECL}><p:cSld>{sptree([sp(2, 'Slide Image Placeholder 1', None, ph=('sldImg', None)), sp(3, 'Notes Placeholder 2', None, ph=('body', 1), text=[S['notes']])])}</p:cSld>"
                "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>",
                CT + "presentationml.notesSlide+xml",
            )
    for f in ("presProps", "viewProps", "tableStyles"):
        root = {
            "presProps": "p:presentationPr",
            "viewProps": "p:viewPr",
            "tableStyles": "a:tblStyleLst",
        }[f]
        extra = ' def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"' if f == "tableStyles" else ""
        d.add(
            f"ppt/{f}.xml",
            HDR + f"<{root} {NSDECL}{extra}/>",
            CT + f"presentationml.{f}+xml"
            if f != "tableStyles"
            else CT + "presentationml.tableStyles+xml",
        )
        pres_rels.append(d.rel(P, f, f"{f}.xml"))
    d.rel(P, "theme", "theme/theme1.xml")
    sec = ""
    if spec.get("sections"):
        secs = ""
        for k, (title, idxs) in enumerate(spec["sections"]):
            ids = "".join(f'<p14:sldId id="{sld_ids[i - 1][0]}"/>' for i in idxs)
            secs += f'<p14:section name="{title}" id="{{0000000{k}-1111-2222-3333-444444444444}}"><p14:sldIdLst>{ids}</p14:sldIdLst></p14:section>'
        sec = (
            '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">'
            f'<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">{secs}</p14:sectionLst></p:ext></p:extLst>'
        )
    sids = "".join(f'<p:sldId id="{i}" r:id="{r}"/>' for i, r in sld_ids)
    show = ""
    if spec.get("custom_show"):
        refs = "".join(f'<p:sld r:id="{sld_ids[i - 1][1]}"/>' for i in spec["custom_show"])
        show = f'<p:custShowLst><p:custShow name="Short version" id="0"><p:sldLst>{refs}</p:sldLst></p:custShow></p:custShowLst>'
    d.add(
        P,
        HDR
        + f'<p:presentation {NSDECL} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="{pres_rels[0]}"/></p:sldMasterIdLst>'
        f'{nm_xml}<p:sldIdLst>{sids}</p:sldIdLst><p:sldSz cx="{W}" cy="{H}"/><p:notesSz cx="6858000" cy="9144000"/>{show}{sec}</p:presentation>',
        None,
    )
    for extra in spec.get("extras", []):
        extra(d)
    return d


# --- concrete fixtures ----------------------------------------------------------

W169, H169 = 12192000, 6858000


def std_layouts(W, H, renamed=None, moved=0, drop=(), localized=False, agenda=False, art=False):
    T = [W // 16, H // 16, W * 7 // 8, H // 8]
    B = [W // 16, H // 4, W * 7 // 8, H * 9 // 16]
    L = [
        {
            "name": "Címdia" if localized else "Title Slide",
            "type": "title",
            "bg": "image" if art else "accent1",
            "phs": [
                ("ctrTitle", None, [W // 10, H // 3, W * 4 // 5, H // 6], "", False),
                ("subTitle", 1, [W // 10, H // 2 + 100000, W * 4 // 5, H // 10], "", False),
            ],
        },
        {
            "name": "Title and Content",
            "type": "obj",
            "phs": [
                ("title", None, T, "", False),
                ("obj", 1, [B[0], B[1] + moved, B[2], B[3]], "", False),
            ],
        },
        {
            "name": "Section Header",
            "type": "secHead",
            "bg": "image" if art else "dk2",
            "phs": [
                ("title", None, [W // 10, H * 2 // 5, W * 4 // 5, H // 6], "", False),
                ("body", 1, [W // 10, H * 3 // 5, W * 4 // 5, H // 10], "", False),
            ],
        },
        {
            "name": "Two Content",
            "type": "twoObj",
            "phs": [
                ("title", None, T, "", False),
                ("obj", 1, [B[0], B[1], B[2] // 2 - 100000, B[3]], "", False),
                ("obj", 2, [B[0] + B[2] // 2 + 100000, B[1], B[2] // 2 - 100000, B[3]], "", False),
            ],
        },
        {"name": "Title Only", "type": "titleOnly", "phs": [("title", None, T, "", False)]},
        {
            "name": "Picture with Caption",
            "type": "picTx",
            "phs": [
                ("title", None, [W // 16, H // 16, W * 3 // 8, H // 8], "", False),
                ("pic", 1, [W // 2, H // 8, W * 7 // 16, H * 3 // 4], "", False),
                ("body", 2, [W // 16, H // 4, W * 3 // 8, H // 2], "", False),
            ],
        },
        {"name": "Blank", "type": "blank", "phs": []},
    ]
    if agenda:
        phs = [("title", None, T, "", False)]
        for i in range(5):
            y = H // 4 + i * H // 9
            phs.append(("body", 10 + i, [W // 16, y, W // 12, H // 12], f"0{i + 1}", True))
            phs.append(
                ("body", 20 + i, [W // 16 + W // 10, y, W // 2, H // 12], "Agenda item", True)
            )
        L.append({"name": "Agenda", "phs": phs, "pics": [("Kép 3", [W * 2 // 3, 0, W // 3, H])]})
    out = []
    for x in L:
        if x["name"] in drop:
            continue
        if renamed and x["name"] in renamed:
            x = {**x, "name": renamed[x["name"]]}
        out.append(x)
    return out


def layout_index(layouts, name):
    return next(i for i, L in enumerate(layouts, 1) if L["name"] == name)


def sample_slides(W, H, layouts, messy=False, chart=False):
    li = lambda n: layout_index(layouts, n)  # noqa: E731
    T = [W // 16, H // 16, W * 7 // 8, H // 8]

    def title_ph(text):
        return sp(2, "Title 1", None, ph=("title",), text=[text])

    def cards(d, s, W, H):
        out = [
            title_ph("Three ways we help")
            if not messy
            else sp(
                2,
                "Google Shape;3265;p106",
                T,
                txbox=True,
                text=["Three ways we help"],
                sz=3200,
                bold=True,
            )
        ]
        cw = W * 7 // 8 // 3 - 150000
        for k in range(3):
            x = W // 16 + k * (cw + 225000)
            box = [x, H // 4, cw, H // 2]
            nm = (
                (lambda j, k=k: "Rectangle 5")
                if messy
                else (lambda j, k=k: f"Card {k + 1} part {j}")
            )
            out.append(
                grp(
                    10 + k * 4,
                    f"Card {k + 1}" if not messy else "Group 9",
                    box,
                    [
                        sp(11 + k * 4, nm(1), box, geom="roundRect", fill="lt2", line="none"),
                        sp(
                            12 + k * 4,
                            nm(2),
                            [x + 100000, H // 4 + 100000, cw - 200000, H // 10],
                            txbox=True,
                            text=[f"Card heading {k + 1}"],
                            sz=2000,
                            bold=True,
                        ),
                        sp(
                            13 + k * 4,
                            nm(3),
                            [x + 100000, H // 4 + 100000 + H // 10, cw - 200000, H // 3],
                            txbox=True,
                            text=["Short supporting sentence for this card."],
                            sz=1400,
                        ),
                    ],
                )
            )
        return out

    def timeline(d, s, W, H):
        box = [W // 16, H // 4, W * 7 // 8, H // 2]
        rows = [
            ["Workstream", "Q1", "Q2", "Q3", "Q4"],
            ["Discovery", "", "", "", ""],
            ["Build", "", "", "", ""],
            ["Rollout", "", "", "", ""],
        ]
        out = [title_ph("Timeline"), table(5, "Table 4", box, rows)]
        cw, rh = box[2] // 5, box[3] // 4
        for r, (c0, c1) in enumerate([(1, 2), (2, 3), (3, 4)], 1):
            out.append(
                sp(
                    20 + r,
                    "Ellipszis 58" if messy else f"Bar {r}",
                    [
                        box[0] + c0 * cw,
                        box[1] + r * rh + rh // 4,
                        (c1 - c0 + 1) * cw - 50000,
                        rh // 2,
                    ],
                    geom="roundRect",
                    fill="accent2",
                    line="none",
                )
            )
        return out

    def comparison(d, s, W, H):
        box = [W // 16, H // 4, W * 7 // 8, H // 2]
        rows = [
            ["Criterion", "Option A", "Option B"],
            ["Speed", "", ""],
            ["Cost", "", ""],
            ["Risk", "", ""],
        ]
        out = [title_ph("Option comparison"), table(5, "Table 4", box, rows)]
        cw, rh = box[2] // 3, box[3] // 4
        k = 30
        for r in range(1, 4):
            for c in (1, 2):
                for dot in range(3):
                    k += 1
                    out.append(
                        sp(
                            k,
                            f"Oval {k}",
                            [
                                box[0] + c * cw + 150000 + dot * 250000,
                                box[1] + r * rh + rh // 3,
                                180000,
                                180000,
                            ],
                            geom="ellipse",
                            fill="7F3FBF" if messy and k == 31 else "accent1",
                            line="none",
                        )
                    )
        return out

    def kpis(d, s, W, H):
        out = [title_ph("Key numbers")]
        cw = W * 7 // 8 // 3 - 150000
        for k in range(3):
            x = W // 16 + k * (cw + 225000)
            out.append(
                grp(
                    10 + k * 3,
                    f"KPI {k + 1}",
                    [x, H // 3, cw, H // 3],
                    [
                        sp(
                            11 + k * 3,
                            f"KPI value {k + 1}",
                            [x, H // 3, cw, H // 6],
                            txbox=True,
                            text=["42%"],
                            sz=4000,
                            bold=True,
                            color="accent1",
                        ),
                        sp(
                            12 + k * 3,
                            f"KPI label {k + 1}",
                            [x, H // 2, cw, H // 8],
                            txbox=True,
                            text=["metric label"],
                            sz=1400,
                        ),
                    ],
                )
            )
        return out

    def bullets(d, s, W, H):
        return [
            title_ph("Context"),
            sp(
                3,
                "Content Placeholder 2",
                None,
                ph=("obj", 1),
                text=["First point", "Second point", "Third point"],
            ),
        ]

    slides = [
        {
            "layout": li("Címdia" if messy else "Title Slide"),
            "shapes": [
                sp(2, "Title 1", None, ph=("ctrTitle",), text=["Presentation title"]),
                sp(3, "Subtitle 2", None, ph=("subTitle", 1), text=["Subtitle"]),
            ],
            "notes": "Cover. Do not modify the layout.",
        },
        {"layout": li("Title and Content"), "shapes": bullets, "notes": "Plain bullet slide."},
        {
            "layout": li("Title Only"),
            "shapes": cards,
            "notes": "Three cards. Delete a whole card if you need two.",
        },
        {"layout": li("Title Only"), "shapes": timeline},
        {"layout": li("Title Only"), "shapes": comparison},
        {"layout": li("Title Only"), "shapes": kpis},
    ]
    if messy:
        # PowerPoint-style markup: an Ignorable prefix declared but unused by any element
        # (ElementTree drops such declarations), an AlternateContent shape, and an
        # animation that targets a card shape (must go when that card is deleted).
        slides[5]["root_extra"] = (
            ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            ' xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
            ' xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" mc:Ignorable="p14 a14"'
        )
        alt = (
            '<mc:AlternateContent><mc:Choice Requires="p14">'
            + sp(
                60,
                "Footnote (new)",
                [W // 16, H * 3 // 4, W // 2, H // 16],
                txbox=True,
                text=["Source: fixture data"],
                sz=1400,
            )
            + "</mc:Choice><mc:Fallback>"
            + sp(
                60,
                "Footnote (old)",
                [W // 16, H * 3 // 4, W // 2, H // 16],
                txbox=True,
                text=["Source: fixture data"],
                sz=1400,
            )
            + "</mc:Fallback></mc:AlternateContent>"
        )
        slides[5]["shapes"] = lambda d, s, W, H, f=kpis: [*f(d, s, W, H), alt]
        slides[2]["tail"] = (
            '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:par><p:cTn id="2" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            '<p:set><p:cBhvr><p:cTn id="3" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="19"/></p:tgtEl>'
            "<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>"
            '<p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par>'
            "</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"
        )
        slides.append(
            {
                "layout": li("Agenda"),
                "shapes": [title_ph("Agenda")]
                + [
                    sp(
                        30 + i,
                        f"Text Placeholder {i}",
                        None,
                        ph=("body", 20 + i),
                        text=[f"Topic {i + 1}"],
                    )
                    for i in range(5)
                ]
                + [
                    sp(40 + i, f"Number {i}", None, ph=("body", 10 + i), text=[""])
                    for i in range(5)
                ],
            }
        )
        for b in range(3):

            def icons(d, s, W, H, b=b):
                return [
                    pic(
                        10 + i,
                        f"Kép {i}",
                        [W // 16 + (i % 6) * W // 7, H // 5 + (i // 6) * H // 3, W // 10, W // 10],
                        d.media(s, (20 * i % 255, 120, 200 - b * 40)),
                    )
                    for i in range(12)
                ]

            slides.append({"layout": li("Blank"), "shapes": icons})
    if chart:

        def chart_slide(d, s, W, H):
            rid = d.rel(s, "chart", "../charts/chart1.xml")
            return [
                title_ph("Revenue by quarter"),
                chart_frame(5, "Chart 4", [W // 16, H // 4, W * 7 // 8, H * 5 // 8], rid),
            ]

        slides.append({"layout": li("Title Only"), "shapes": chart_slide, "notes": "Native chart."})
    return slides


def xlsx() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr(
            "[Content_Types].xml",
            HDR
            + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        )
        z.writestr(
            "_rels/.rels",
            HDR
            + f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{RT}officeDocument" Target="xl/workbook.xml"/></Relationships>',
        )
        z.writestr(
            "xl/workbook.xml",
            HDR
            + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            HDR
            + f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{RT}worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        )
        rows = "".join(
            f'<row r="{i + 2}"><c r="A{i + 2}" t="inlineStr"><is><t>Q{i + 1}</t></is></c><c r="B{i + 2}"><v>{v}</v></c></row>'
            for i, v in enumerate([3, 5, 4, 7])
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            HDR
            + f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row>{rows}</sheetData></worksheet>',
        )
    return buf.getvalue()


def add_chart(d: Deck):
    cats = "".join(f'<c:pt idx="{i}"><c:v>Q{i + 1}</c:v></c:pt>' for i in range(4))
    vals = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate([3, 5, 4, 7]))
    xml = (
        HDR
        + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
        '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>'
        f'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/>{cats}</c:strCache></c:strRef></c:cat>'
        f'<c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/>{vals}</c:numCache></c:numRef></c:val></c:ser>'
        '<c:gapWidth val="150"/><c:axId val="111"/><c:axId val="222"/></c:barChart>'
        '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>'
        '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>'
        '</c:plotArea><c:plotVisOnly val="1"/></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>'
    )
    d.add("ppt/charts/chart1.xml", xml, CT + "drawingml.chart+xml")
    d.add("ppt/embeddings/Microsoft_Excel_Worksheet.xlsx", xlsx(), None)
    d.rel("ppt/charts/chart1.xml", "package", "../embeddings/Microsoft_Excel_Worksheet.xlsx")


CLEAN_THEME = (
    "Fixture Clean",
    [
        "1F2A36",
        "FFFFFF",
        "2E4057",
        "EEF1F4",
        "1B998B",
        "ED6A5A",
        "F4C95D",
        "6C8EAD",
        "8E7DBE",
        "3D5A80",
        "1B6CA8",
        "6C5B7B",
    ],
    "Lato",
    "Lato",
)
MESSY_THEME = (
    "Fixture Messy",
    [
        "111111",
        "FFFFFF",
        "3B3B58",
        "F2F2F2",
        "00A6A6",
        "F49D37",
        "D72638",
        "3F88C5",
        "140F2D",
        "8AAA79",
        "3F88C5",
        "6D6875",
    ],
    "Lato",
    "Lato",
)
CHART_THEME = (
    "Fixture Charts",
    [
        "202020",
        "FFFFFF",
        "404040",
        "F0F0F0",
        "0F7173",
        "E7ECEF",
        "F05D5E",
        "272932",
        "D8A47F",
        "8E8E8E",
        "0F7173",
        "404040",
    ],
    "Nonexistent Grotesk",
    "Nonexistent Grotesk",
)


def clean():
    lay = std_layouts(W169, H169)
    return build(
        {
            "w": W169,
            "h": H169,
            "theme": CLEAN_THEME,
            "layouts": lay,
            "notes_master": True,
            "chrome": [("logo", {}), ("sldnum_ph", {}), ("rule", {})],
            "slides": sample_slides(W169, H169, lay),
        }
    )


def messy(v2=False):
    kw = {"localized": True, "agenda": True, "art": True}
    if v2:
        kw.update(
            renamed={"Two Content": "Two Columns"},
            moved=40 * EMU_PX,
            drop=("Picture with Caption",),
        )
    lay = std_layouts(W169, H169, **kw)
    th = MESSY_THEME
    if v2:
        th = (th[0], [*th[1][:5], "C05746", *th[1][6:]], th[2], th[3])
    slides = sample_slides(W169, H169, lay, messy=True)
    return build(
        {
            "w": W169,
            "h": H169,
            "kind": "potx",
            "theme": th,
            "layouts": lay,
            "chrome": [("logo", {"name": "Kép 1"}), ("sldnum_box", {"name": "TextBox 7"})],
            "slides": slides,
            "sections": [("Samples", [1, 2, 3, 4, 5, 6, 7]), ("Icon bank", [8, 9, 10])],
            "custom_show": [1, 3],
        }
    )


def charts():
    W, H = 9144000, 6858000
    lay = std_layouts(W, H)
    return build(
        {
            "w": W,
            "h": H,
            "theme": CHART_THEME,
            "layouts": lay,
            "notes_master": True,
            "chrome": [("sldnum_ph", {})],
            "slides": sample_slides(W, H, lay, chart=True),
            "extras": [add_chart],
        }
    )


def main(argv) -> int:
    out = Path(argv[1] if len(argv) > 1 else "fixtures-out")
    out.mkdir(parents=True, exist_ok=True)
    clean().save(out / "clean.pptx")
    messy().save(out / "messy.potx")
    messy(v2=True).save(out / "drift.potx")
    charts().save(out / "charts.pptx")
    print(f"wrote clean.pptx messy.potx drift.potx charts.pptx to {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
