#!/usr/bin/env python3
"""Lint a built .pptx against its style contract.

    python3 lint.py DECK.pptx --contract contract.json [--json] [--no-require-notes]

Errors: broken package (content types, dangling rels, orphan parts, slide ids, sections),
layouts outside the contract, off-palette colours, non-theme fonts, text below the
template's smallest size, shapes off the slide, free shapes over master/layout chrome,
leftover sample or prompt text, missing titles, missing notes on content slides, text
overflow measured with real font metrics. Warnings: approximate fits (font missing),
empty placeholders, low contrast.
Exit codes: 0 no errors, 1 errors, 2 usage/read error.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fit
import pptx_core as core
from pptx_core import q

REL_ATTRS = (q("r:id"), q("r:embed"), q("r:link"), q("r:pict"))


def distinctive(text: str) -> bool:
    """Sample text worth flagging if it survives: 3+ words or 16+ chars ('Agenda' is fine)."""
    t = text.strip()
    return len(t.split()) >= 3 or len(t) >= 16


def lint(deck, contract: dict, require_notes: bool = True, build_report=None) -> list[dict]:
    pkg = core.Package(deck)
    out: list[dict] = []

    def add(sev, code, slide, msg):
        out.append({"severity": sev, "code": code, "slide": slide, "message": msg})

    _package_checks(pkg, Path(deck), add)
    layouts = {L["name"]: L for L in contract["layouts"]}
    theme_hex = set(contract["theme"]["colors"].values())
    allowed_colors = theme_hex | set(contract["derived"].get("colors_used", {}))
    tf = contract["theme"]["fonts"]
    allowed_fonts = {f for f in (tf.get("major"), tf.get("minor")) if f} | set(
        contract["derived"].get("fonts_used", [])
    )
    allowed_fonts |= {"+mj-lt", "+mn-lt", "+mj-ea", "+mn-ea", "+mj-cs", "+mn-cs"}
    min_size = contract["derived"].get("min_font_size") or 10
    sample_text = set()
    allowed_boxes = set()
    for P in contract["patterns"]:
        allowed_boxes |= {tuple(b) for b in P.get("shape_boxes", [])}
        for S in P["slots"].values():
            for t in S.get("text") or []:
                for x in t if isinstance(t, list) else [t]:
                    if distinctive(str(x)):
                        sample_text.add(str(x).strip())
        for R in P["repeats"].values():
            for S in R["slots"]:
                sample_text |= {t.strip() for t in S["text"] if distinctive(t)}
    sw, sh = contract["slide_size"]
    src = {r["part"]: r["source"] for r in (build_report or [])}
    master_chrome = {}
    for M in contract["masters"]:
        master_chrome[M["part"]] = [c for c in M["chrome"] if c["box"]]
    for n, part in enumerate(pkg.slides(), 1):
        root = pkg.xml(part)
        tag = f"{n}" + (f" ({src[part]})" if part in src else "")
        lp = pkg.related(part, "slideLayout")
        lp = lp[0] if lp else None
        lname = pkg.xml(lp).find(q("p:cSld")).get("name", "") if lp else ""
        L = layouts.get(lname)
        if L is None:
            add("error", "off-contract-layout", tag, f"layout {lname!r} is not in the contract")
        master = pkg.related(lp, "slideMaster")[0] if lp else pkg.masters()[0]
        theme = core.theme_of(pkg, master)
        cmap = core.clr_map(pkg, master)
        tree = root.find(f"{q('p:cSld')}/{q('p:spTree')}")
        shapes = list(core.iter_shapes(tree))
        # titles
        titles = [
            el
            for _p, el, _b in shapes
            if (ph := core.ph_of(el)) is not None and core.ph_type(ph) in core.TITLE_PH
        ]
        has_title_ph = L is not None and any(p["type"] in core.TITLE_PH for p in L["placeholders"])
        if titles and not any(" ".join(core.paragraphs(core.txbody(t))).strip() for t in titles):
            add("error", "missing-title", tag, "title placeholder is empty")
        elif not titles and has_title_ph:
            s_src = src.get(part, "")
            P = next((x for x in contract["patterns"] if f"pattern:{x['id']}" == s_src), None)
            if P is not None and not any(
                S.get("ph", {}).get("type") in core.TITLE_PH for S in P["slots"].values()
            ):
                add(
                    "warning",
                    "no-title-placeholder",
                    tag,
                    "the template's sample slide keeps its title in a text box",
                )
            else:
                add(
                    "error",
                    "missing-title",
                    tag,
                    "slide has no title (screen readers and outline view rely on it)",
                )
        elif not titles:
            add("warning", "no-title-placeholder", tag, "slide has no title placeholder")
        # notes
        notes = pkg.related(part, "notesSlide")
        note_text = ""
        if notes:
            for _p, el, _b in core.iter_shapes(
                pkg.xml(notes[0]).find(f"{q('p:cSld')}/{q('p:spTree')}")
            ):
                ph = core.ph_of(el)
                if ph is not None and core.ph_type(ph) == "body":
                    note_text += " ".join(core.paragraphs(core.txbody(el)))
        ltype = (L or {}).get("type") or ""
        content_slide = n > 1 and ltype not in ("title", "secHead")
        if content_slide and not note_text.strip():
            add(
                "error" if require_notes else "warning",
                "missing-notes",
                tag,
                "content slide has no speaker notes",
            )
        # colours / fonts / sizes
        for c in sorted(set(core.colors_in(root))):
            if c not in allowed_colors:
                add(
                    "error",
                    "off-palette-colour",
                    tag,
                    f"#{c} is neither a theme colour nor used by the template",
                )
            elif c not in theme_hex:
                add(
                    "warning",
                    "template-hardcoded-colour",
                    tag,
                    f"#{c} is off-theme but comes from the template's sample slides",
                )
        for f in sorted(set(core.fonts_in(root))):
            if f not in allowed_fonts:
                add(
                    "error",
                    "off-theme-font",
                    tag,
                    f"font {f!r} is not a theme font (theme: {tf.get('major')!r}/{tf.get('minor')!r})",
                )
        for s in sorted(set(core.sizes_in(root))):
            if s + 0.01 < min_size:
                add(
                    "error",
                    "text-too-small",
                    tag,
                    f"{s} pt is below the template's smallest text ({min_size} pt)",
                )
        # geometry, chrome, text
        chrome = (
            list(master_chrome.get(master, [])) if (L or {}).get("show_master_shapes", True) else []
        )
        chrome += [c for c in (L or {}).get("chrome", []) if c["box"]]
        chrome = [c for c in chrome if c["role"] != "decoration" or c["box"][1] > sh * 0.8]
        ids: dict[str, int] = {}
        for p, el, box in shapes:
            c = core.cnv(el)
            if c is not None:
                ids[c.get("id")] = ids.get(c.get("id"), 0) + 1
            ph = core.ph_of(el)
            ebox = core.effective_box(pkg, el, box, lp, master)
            name = c.get("name", "") if c is not None else ""
            if ebox and len(p) == 1:
                tol = core.EMU_PER_PT * 2
                if (
                    ebox[0] < -tol
                    or ebox[1] < -tol
                    or ebox[0] + ebox[2] > sw + tol
                    or ebox[1] + ebox[3] > sh + tol
                ):
                    add("error", "off-slide", tag, f"shape {name!r} extends beyond the slide")
                band = contract["derived"].get("footer_band")
                if (
                    core.kind(el) == "table"
                    and band
                    and tuple(ebox) not in allowed_boxes
                    and ebox[1] + ebox[3] > band[1]
                ):
                    add(
                        "error",
                        "covers-chrome",
                        tag,
                        f"table {name!r} grows into the footer band; cut rows or split",
                    )
                if ph is None and tuple(ebox) not in allowed_boxes and core.kind(el) != "table":
                    for ch in chrome:
                        if core.intersects(ebox, ch["box"]):
                            add(
                                "error",
                                "covers-chrome",
                                tag,
                                f"free shape {name!r} overlaps the {ch['role']} ({ch['where']} chrome)",
                            )
                            break
            tb = core.txbody(el)
            paras = core.paragraphs(tb)
            text = " ".join(paras).strip()
            for para in paras:
                if para.strip() in sample_text:
                    add(
                        "error",
                        "leftover-sample-text",
                        tag,
                        f"{name!r} still says {para.strip()[:40]!r} from the template's sample slide",
                    )
                elif core.LEFTOVER_RE.search(para):
                    add(
                        "error",
                        "placeholder-text",
                        tag,
                        f"{name!r} contains placeholder text {para.strip()[:40]!r}",
                    )
            if ph is not None and not text and core.ph_type(ph) not in core.FOOTER_PH:
                lph = next(
                    (
                        x
                        for x in (L or {}).get("placeholders", [])
                        if (x["idx"] and x["idx"] == int(ph.get("idx", "0")))
                        or (not x["idx"] and x["type"] == core.ph_type(ph))
                    ),
                    None,
                )
                if lph and lph["prompt_looks_like_content"]:
                    add(
                        "error",
                        "prompt-only-content",
                        tag,
                        f"placeholder {lph['key']} is empty; its prompt {lph['prompt']!r} never shows in the show or PDF",
                    )
                elif core.kind(el) != "picture":
                    add(
                        "warning",
                        "empty-placeholder",
                        tag,
                        f"empty placeholder {name!r} shows prompt text in edit view; delete or fill it",
                    )
            if text and ebox and not core.has_field(el):
                _fit(pkg, el, ebox, paras, lp, master, theme, tag, name, add)
                _contrast(pkg, el, part, lp, master, theme, cmap, tag, name, add)
        for sid, cnt in ids.items():
            if cnt > 1:
                add(
                    "error",
                    "duplicate-shape-id",
                    tag,
                    f"shape id {sid} used {cnt} times (PowerPoint repairs the file)",
                )
    return out


def _fit(pkg, el, box, paras, lp, master, theme, tag, name, add):
    tb = core.txbody(el)
    bp = tb.find(q("a:bodyPr"))
    if bp is not None and bp.find(q("a:spAutoFit")) is not None:
        return
    size, face, bold = core.text_props(pkg, el, lp, master, theme)
    norm = bp.find(q("a:normAutofit")) if bp is not None else None
    if norm is not None and norm.get("fontScale"):
        size *= int(norm.get("fontScale")) / 100000
    ph = core.ph_of(el)
    bullets = ph is not None and core.ph_type(ph) in ("body", "obj") and len(paras) > 1
    before, line = core.para_spacing(pkg, el, lp, master)
    r = fit.check(
        paras, box, size, face, bold, core.insets(el), 342900 if bullets else 0, 1.2 * line, before
    )
    if r["fits"]:
        return
    approx = r["how"] == "missing"
    msg = (
        f"{name!r} overflows: needs {r['needed']} lines, fits {r['available']} "
        f"(~{r['chars_per_line']} chars/line at {size:g} pt {face})"
        + (" - approximate, font not installed" if approx else "")
        + ". Shorten, move detail to notes, or split; do not shrink below the template sizes"
    )
    severe = r["needed"] > r["available"] * 1.2
    add("error" if (not approx or severe) else "warning", "text-overflow", tag, msg)


def _contrast(pkg, el, part, lp, master, theme, cmap, tag, name, add):
    if el.find(f"{q('p:spPr')}/{q('a:solidFill')}") is not None:
        return  # text sits on its own fill; judged by the template designer
    bg = None
    for p in (part, lp, master):
        kind, hexv = core.background(pkg, p, theme, cmap) if p else ("none", None)
        if kind == "image":
            return
        if kind == "solid":
            bg = hexv
            break
    bg = bg or theme["colors"].get(cmap.get("bg1", "lt1"))
    tb = core.txbody(el)
    size, _face, bold = core.text_props(pkg, el, lp, master, theme)
    for rpr in tb.iter(q("a:rPr")):
        sf = rpr.find(q("a:solidFill"))
        if sf is None:
            continue
        fg = core.resolve_color(sf, theme, cmap)
        if not fg or not bg or not re.fullmatch(r"[0-9A-F]{6}", fg + bg):
            continue
        sz = int(rpr.get("sz")) / 100 if rpr.get("sz") else size
        large = sz >= 18 or (sz >= 14 and (rpr.get("b") == "1" or bold))
        ratio = core.contrast(fg, bg)
        if ratio < (3.0 if large else 4.5):
            add(
                "warning",
                "low-contrast",
                tag,
                f"{name!r}: #{fg} on #{bg} is {ratio}:1 (AA needs {'3' if large else '4.5'}:1)",
            )
            return


def _package_checks(pkg: core.Package, path: Path, add):
    names = set(pkg.parts)
    for n in names:
        if n.endswith(".rels") or n == "[Content_Types].xml":
            continue
        if pkg.content_type(n) is None:
            add("error", "pkg-content-type", "-", f"part {n} has no content type")
    for n in sorted(names):
        if not n.endswith(".rels"):
            continue
        src = "" if n == "_rels/.rels" else core.rels_source(n)
        for r in pkg.rels(src):
            if "part" in r and r["part"] not in names:
                add(
                    "error", "pkg-dangling-rel", "-", f"{n} {r['Id']} points at missing {r['part']}"
                )
    reach = pkg.reachable()
    for n in sorted(names):
        if n.endswith(".rels") or n == "[Content_Types].xml" or n.startswith("customXml/"):
            continue
        if n not in reach:
            add(
                "error",
                "pkg-orphan-part",
                "-",
                f"{n} is not referenced by any relationship (bloat)",
            )
    for n in sorted(names):
        if not (n.endswith(".xml") and n.startswith("ppt/") and "/_rels/" not in n):
            continue
        root = pkg.xml(n)
        rids = {r["Id"] for r in pkg.rels(n)}
        for el in root.iter():
            for a in REL_ATTRS:
                v = el.get(a)
                if v and v not in rids:
                    add(
                        "error",
                        "pkg-missing-rel-id",
                        "-",
                        f"{n} references {v} which has no relationship",
                    )
    ids = [int(s.get("id")) for s, _r, _p in pkg.slide_ids()]
    if len(ids) != len(set(ids)) or any(not 256 <= i < 2**31 for i in ids):
        add("error", "pkg-slide-ids", "-", "slide ids must be unique and within 256..2^31-1")
    if any(p is None for _s, _r, p in pkg.slide_ids()):
        add("error", "pkg-slide-ids", "-", "a slide id points at no slide part")
    pres = pkg.xml(pkg.main)
    seen = []
    for s in pres.iter(q("p14:sldId")):
        seen.append(int(s.get("id")))
    if len(seen) != len(set(seen)):
        add(
            "error",
            "pkg-sections",
            "-",
            "a slide is listed in two sections (PowerPoint cannot open the file)",
        )
    if set(seen) - set(ids):
        add("warning", "pkg-sections", "-", "sections list slide ids that no longer exist")
    if seen and set(ids) - set(seen):
        add("warning", "pkg-sections", "-", "some slides are in no section")
    ct = pkg.content_type(pkg.main)
    if path.suffix.lower() == ".pptx" and ct == core.CT_TMPL:
        add(
            "error",
            "pkg-template-type",
            "-",
            ".pptx file still has the template (.potx) main content type",
        )


def format_findings(findings) -> str:
    e = [f for f in findings if f["severity"] == "error"]
    w = [f for f in findings if f["severity"] == "warning"]
    lines = [f"lint: {len(e)} errors, {len(w)} warnings"]
    for f in e + w:
        lines.append(
            f"  {f['severity'].upper():7} slide {f['slide']}: [{f['code']}] {f['message']}"
        )
    return "\n".join(lines)


def self_test() -> int:
    assert format_findings([]) == "lint: 0 errors, 0 warnings"
    out = format_findings([{"severity": "error", "code": "x", "slide": "1", "message": "m"}])
    assert "ERROR" in out and "[x]" in out
    print("self-test OK")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("deck", nargs="?")
    ap.add_argument("--contract")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--no-require-notes", action="store_true")
    ap.add_argument("--self-test", action="store_true")
    a = ap.parse_args(argv)
    if a.self_test:
        return self_test()
    if not a.deck or not a.contract:
        ap.print_usage()
        return 2
    try:
        contract = core.load_contract(a.contract)
        report_path = Path(a.deck).with_suffix(".build.json")
        report = json.loads(report_path.read_text())["slides"] if report_path.exists() else None
        findings = lint(a.deck, contract, not a.no_require_notes, report)
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(findings, indent=1) if a.json else format_findings(findings))
    return 1 if any(f["severity"] == "error" for f in findings) else 0


if __name__ == "__main__":
    sys.exit(main())
