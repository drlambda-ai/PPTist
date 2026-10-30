"""PPTX template engine: read a template's text slots, clone slides, fill text in place, render previews.

Everything here is deterministic. The LLM only produces a spec (see agent.py); this file turns
spec + template into a native .pptx that keeps the template's own fonts, colours and pictures.
"""
import copy
import math
import os
import subprocess
import unicodedata
import zipfile
from xml.sax.saxutils import escape

from lxml import etree

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

EMU_PER_PT = 12700
# Trend colours are off by default: ▲/▼ carry the direction in the template's own palette (and red/green read
# opposite ways in China vs the West). TREND_COLORS="1E9E4A,D62828" turns colouring on.
UP, DOWN = (os.environ.get("TREND_COLORS", "").split(",") + ["", ""])[:2]
CJK_FONT = os.environ.get("CJK_FONT", "Microsoft YaHei")
PREVIEW_CJK_FONT = os.environ.get("PREVIEW_CJK_FONT", "Hiragino Sans GB")  # installed on macOS


def _walk(shapes):
    for sh in shapes:
        if sh.shape_type == 6:  # group
            yield from _walk(sh.shapes)
        else:
            yield sh


def _font_pt(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            if r.font.size:
                return r.font.size.pt
    return 12.0  # ponytail: inherited sizes (layout/master) not resolved; 12pt is the safe guess


def _lines(h_emu, pt):
    return max(1, int(h_emu / EMU_PER_PT / (pt * 1.2)))


def _capacity(w_emu, h_emu, pt):
    """Rough characters-that-fit, counting a CJK char as 1 and a Latin char as ~0.6."""
    w_pt = max(w_emu / EMU_PER_PT - 14, 10)
    return int(w_pt / pt * _lines(h_emu, pt))


def slots_of(slide, sw, sh_):
    """Every visible text-bearing shape (and table cell) on a slide, keyed by a stable local id."""
    out = []
    for sh in _walk(slide.shapes):
        if sh.left is None or sh.left >= sw or sh.top >= sh_ or sh.left + sh.width <= 0 or sh.top + sh.height <= 0:
            continue  # off-canvas notes left in the template
        box = dict(x=round(sh.left / sw * 100), y=round(sh.top / sh_ * 100),
                   w=round(sh.width / sw * 100), h=round(sh.height / sh_ * 100))
        if sh.has_text_frame:
            text = sh.text_frame.text.strip()
            if not text and sh.shape_type != 17 and not sh.is_placeholder:
                continue  # empty decorative autoshape
            pt = _font_pt(sh.text_frame)
            out.append(dict(id=str(sh.shape_id), text=text, pt=pt, lines=_lines(sh.height, pt),
                            max_chars=_capacity(sh.width, sh.height, pt), accent=_has_accent(sh.text_frame), **box))
        elif getattr(sh, "has_table", False) and sh.has_table:
            tbl = sh.table
            for r, row in enumerate(tbl.rows):
                for c, cell in enumerate(row.cells):
                    w = tbl.columns[c].width
                    pt = _font_pt(cell.text_frame)
                    out.append(dict(id=f"{sh.shape_id}:r{r}c{c}", text=cell.text_frame.text.strip(), pt=pt,
                                    lines=_lines(row.height, pt), max_chars=_capacity(w, row.height, pt), **box))
    return out


def _style(rPr):
    """What a reader sees of a run style: colour, weight, face (serialised XML differs by namespace noise)."""
    if rPr is None:
        return ()
    fill = rPr.find(qn("a:solidFill"))
    clr = fill[0] if fill is not None and len(fill) else None
    latin = rPr.find(qn("a:latin"))
    return (clr.tag if clr is not None else None, clr.get("val") if clr is not None else None, rPr.get("b"),
            latin.get("typeface") if latin is not None else None)


def _has_accent(tf):
    """True when the template mixes a highlight style into this box's text (e.g. blue bold key words)."""
    colours = {_style(r.find(qn("a:rPr")))[:2] for r in tf._txBody.iter(qn("a:r")) if (r.findtext(qn("a:t")) or "").strip()}
    return len(colours) > 1  # a highlight is a colour change; a smaller unit ("x" + "次") is not one


def _chrome_shapes(prs):
    """Text the masters/layouts used by the deck print on every slide (footers like "Team Name • 2024"), keyed
    by layout/master name + shape id. Placeholders (slides fill those) and fields (slide numbers) are skipped."""
    used = list({id(s.slide_layout._element): s.slide_layout for s in prs.slides}.values())
    masters = list({id(l.slide_master._element): l.slide_master for l in used}.values())
    out = {}
    for part in masters + used:
        for sh in part.shapes:
            if sh.has_text_frame and not sh.is_placeholder and sh.text_frame.text.strip() \
                    and next(sh._element.iter(qn("a:fld")), None) is None:
                out.setdefault(f"{part.name or 'master'}:{sh.shape_id}", sh)
    return out


def chrome(path):
    return [dict(id=k, text=sh.text_frame.text.strip()) for k, sh in _chrome_shapes(Presentation(path)).items()]


def analyze(path):
    prs = Presentation(path)
    sw, sh = prs.slide_width, prs.slide_height
    return [dict(n=i, slots=slots_of(s, sw, sh)) for i, s in enumerate(prs.slides, 1)]


# ---------- slide cloning ----------

def clone_slide(prs, src):
    new = prs.slides.add_slide(src.slide_layout)
    rid_map = {}
    for rId, rel in src.part.rels.items():
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        target = rel.target_ref if rel.is_external else rel.target_part
        rid_map[rId] = new.part.relate_to(target, rel.reltype, is_external=rel.is_external)
    r_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

    def dup(el):
        c = copy.deepcopy(el)
        for e in c.iter():
            for k, v in e.attrib.items():
                if k.startswith(r_ns) and v in rid_map:
                    e.set(k, rid_map[v])
        return c

    # python-pptx caches the spTree object, so refill the existing elements instead of swapping them
    new_cSld, src_cSld = new._element.cSld, src._element.cSld
    new_tree, src_tree = new_cSld.spTree, src_cSld.spTree
    for c in list(new_tree):
        new_tree.remove(c)
    for c in src_tree:
        new_tree.append(dup(c))
    bg = src_cSld.find(qn("p:bg"))
    if bg is not None:
        new_cSld.insert(0, dup(bg))
    for child in src._element:  # clrMapOvr, transition, timing
        if child.tag != qn("p:cSld"):
            old = new._element.find(child.tag)
            if old is not None:
                new._element.remove(old)
            new._element.append(dup(child))
    return new


def _drop_slide(prs, sldId):
    prs.part.drop_rel(sldId.get(qn("r:id")))
    prs.slides._sldIdLst.remove(sldId)


# ---------- text fill ----------

def _is_wide(ch):
    return unicodedata.east_asian_width(ch) in "WF" or ch in "▲▼"


_metric = None


def _metric_font():
    """Bundled Inter Bold at 100px for measuring Latin glyph widths (bold is the wider, safe side), or None."""
    global _metric
    if _metric is None:
        try:
            from PIL import ImageFont
            _metric = ImageFont.truetype(os.path.join(os.path.dirname(__file__), "fonts", "Inter-Bold.otf"), 100)
        except Exception:  # no Pillow: fall back to the unit estimate
            _metric = False
    return _metric or None


def _units(line):
    """Line width in ems. CJK glyphs (and full-width punctuation) are a uniform 1em; Latin is measured."""
    m = _metric_font()
    return sum(1.0 if _is_wide(ch) else (m.getlength(ch) / 100 if m else 0.62) for ch in line)


def _fit_pt(text, w_emu, h_emu, pt, max_lines=None, ls=1.0):
    """Shrink font (down to 60%) until wrapped text fits the box height (or max_lines). 15% slack."""
    # 12% width margin: the renderer draws glyphs a few % wider than measured, and auto-growing boxes push a
    # wrapped line into the shape below, so a near-fit must count as a miss
    w_pt, h_pt = max(w_emu / EMU_PER_PT - 14, 10) * 0.88, h_emu / EMU_PER_PT
    if max_lines:
        h_pt = min(h_pt, max_lines * pt * 1.2 * ls)
    size = pt
    while size > pt * 0.6:
        lines = sum(max(1, math.ceil(_units(l) * size / w_pt)) for l in text.split("\n"))
        if lines * size * 1.2 * ls <= h_pt * 1.15 or lines == 1 and _units(text) * size <= w_pt * 1.15:
            return size
        size -= 0.5
    return size


def _line_spacing(pPr):
    """Paragraph line spacing as a multiple (a:lnSpc/a:spcPct 150000 -> 1.5); 1.0 when unset."""
    pct = pPr.find(qn("a:lnSpc") + "/" + qn("a:spcPct")) if pPr is not None else None
    return int(pct.get("val")) / 100000 if pct is not None else 1.0


def _set_fill(rPr, hex_):
    for tag in ("a:solidFill", "a:gradFill", "a:noFill"):
        for el in rPr.findall(qn(tag)):
            rPr.remove(el)
    fill = rPr.makeelement(qn("a:solidFill"), {})
    clr = fill.makeelement(qn("a:srgbClr"), {"val": hex_})
    fill.append(clr)
    ln = rPr.find(qn("a:ln"))
    rPr.insert(list(rPr).index(ln) + 1 if ln is not None else 0, fill)


def _inherited_face(sh):
    """Latin typeface a placeholder inherits from its layout (titles often set no font on the run)."""
    if not sh.is_placeholder:
        return None
    try:
        lp = sh.part.slide_layout.placeholders.get(idx=sh.placeholder_format.idx)
    except (AttributeError, KeyError, ValueError):
        return None
    latin = next(lp._element.iter(qn("a:latin")), None) if lp is not None else None
    return latin.get("typeface") if latin is not None else None


def _cjk_font(rPr, inherited=None, weight=True):
    # Google-Slides exports set the East Asian font to a Latin face (e.g. Inter), which has no
    # CJK glyphs. The ea slot only applies to CJK characters, so this never touches Latin text.
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        latin = rPr.find(qn("a:latin"))
        rPr.insert(list(rPr).index(latin) + 1 if latin is not None else len(rPr), ea)
    ea.set("typeface", CJK_FONT)
    # Weight is often carried in the Latin family name ("Inter ExtraBold") with no b flag; CJK glyphs come
    # from the ea font at regular weight, so headings turn thin. Carry the weight over as bold.
    latin = rPr.find(qn("a:latin"))
    face = (latin.get("typeface", "") if latin is not None else inherited or "").lower()
    if weight and rPr.get("b") is None and face.endswith(("bold", "black", "heavy")):
        rPr.set("b", "1")


def _accent_rPr(runs, body):
    """The template's highlight style in this box (e.g. blue bold key words inside grey text); else bold body."""
    for r in runs:
        p = r.find(qn("a:rPr"))
        if p is not None and _style(p)[:2] != _style(body)[:2]:
            return copy.deepcopy(p)  # keeps its own size: card key lines are larger than the body
    a = copy.deepcopy(body)
    a.set("b", "1")
    return a


def _size(rPr):
    return int(rPr.get("sz") or 0) if rPr is not None else 0


def _dominant(runs):
    """The style carrying most characters (the first run is often a coloured accent word); ties (number "x" vs
    unit "次 ") go to the larger font: the number's style."""
    chars = {}
    for r in runs:
        k = (_style(r.find(qn("a:rPr"))), _size(r.find(qn("a:rPr"))))
        chars[k] = chars.get(k, 0) + len((r.findtext(qn("a:t")) or "").strip())
    best = max(runs, key=lambda r: (chars[(_style(r.find(qn("a:rPr"))), _size(r.find(qn("a:rPr"))))],
                                    _size(r.find(qn("a:rPr")))), default=None)
    return best.find(qn("a:rPr")) if best is not None else None


def _own(rPr, txBody):
    rPr = copy.deepcopy(rPr) if rPr is not None else txBody.makeelement(qn("a:rPr"), {})
    rPr.tag = qn("a:rPr")
    return rPr


def set_text(tf, text, emphasis=None, shape_filled=False, box=None):
    txBody = tf._txBody
    paras = txBody.findall(qn("a:p"))
    has_text = lambda r: (r.findtext(qn("a:t")) or "").strip()
    runs = [r for r in txBody.iter(qn("a:r")) if has_text(r)] or list(txBody.iter(qn("a:r")))
    # One style per template paragraph: new line i takes paragraph i's look (label line / value line,
    # lead / key / detail line); extra lines reuse the last paragraph's.
    styles = [(p.find(qn("a:pPr")), _own(_dominant(pr), txBody))
              for p in paras if (pr := [r for r in p.iter(qn("a:r")) if has_text(r)])]
    if not styles:
        end = paras[0].find(qn("a:endParaRPr")) if paras else None
        styles = [(paras[0].find(qn("a:pPr")) if paras else None, _own(_dominant(runs) or end, txBody))]
    base = _own(_dominant(runs), txBody) if runs else styles[0][1]
    accent = _accent_rPr(runs, base)
    if any(_is_wide(ch) for ch in text.replace("**", "")):
        for st in [st for _, st in styles] + [accent]:
            _cjk_font(st, (box or {}).get("face"))

    lines = text.split("\n")
    if emphasis in ("up", "down", "key"):
        # the number's line carries the emphasis (a label line above it stays as the template has it)
        num = [i for i, l in enumerate(lines) if any(c.isdigit() for c in l)]  # the comparison line carries the arrow
        j = next((i for i in num if "vs" in lines[i].lower()), num[0] if num else 0)
        st = copy.deepcopy(styles[min(j, len(styles) - 1)][1])
        if emphasis != "key" and not lines[j].startswith(("▲", "▼", "↑", "↓", "+", "-", "−")):
            lines[j] = ("▲ " if emphasis == "up" else "▼ ") + lines[j]
        if emphasis != "key" and not shape_filled and UP:  # coloured text on a coloured box is unreadable
            _set_fill(st, UP if emphasis == "up" else DOWN)
        st.set("b", "1")
        styles = [(pp, st if i == min(j, len(styles) - 1) else s) for i, (pp, s) in enumerate(styles)]

    if box and text:
        pt = max(_size(s) for _, s in styles) / 100 or box["pt"]
        ls = max(_line_spacing(pp) for pp, _ in styles)
        new_pt = _fit_pt("\n".join(lines).replace("**", ""), box["w"], box["h"], pt, box.get("lines"), ls)
        if new_pt < pt:
            for st in [s for _, s in styles] + [accent]:
                if st.get("sz"):
                    st.set("sz", str(int(_size(st) * new_pt / pt)))

    for p in paras:
        txBody.remove(p)
    for i, line in enumerate(lines):
        pPr, rPr = styles[min(i, len(styles) - 1)]
        p = txBody.makeelement(qn("a:p"), {})
        if pPr is not None:
            p.append(copy.deepcopy(pPr))
        if line.replace("**", ""):
            trend = emphasis is None and line.startswith(("▲", "▼")) and not shape_filled and UP
            for k, seg in enumerate(line.split("**")):  # **key phrase** -> the template's own accent style
                if not seg:
                    continue
                r = p.makeelement(qn("a:r"), {})
                lr = copy.deepcopy(accent if k % 2 else rPr)
                if trend:
                    _set_fill(lr, UP if line[0] == "▲" else DOWN)  # per-line trend inside a multi-item box
                r.append(lr)
                t = r.makeelement(qn("a:t"), {})
                t.text = seg
                r.append(t)
                p.append(r)
        else:
            end = copy.deepcopy(rPr)
            end.tag = qn("a:endParaRPr")
            p.append(end)
        txBody.append(p)


def _shape_filled(sh):
    spPr = sh._element.find(qn("p:spPr"))
    return spPr is not None and spPr.find(qn("a:solidFill")) is not None


def fill_slide(slide, fills, emphasis=None, remove=(), lines=None, sw=None):
    emphasis, lines = emphasis or {}, lines or {}
    by_id = {str(sh.shape_id): sh for sh in _walk(slide.shapes)}
    for sid in remove:
        base, _, cell = str(sid).partition(":")
        sh = by_id.get(base)
        if sh is None:
            continue
        if cell:  # a table cell: blank it; removing the shape would delete the whole table
            fills = {**fills, sid: ""}
        else:
            sh._element.getparent().remove(sh._element)
    for sid, text in fills.items():
        base, _, cell = str(sid).partition(":")
        sh = by_id.get(base)
        if sh is None or (sid in remove and not cell):
            continue
        text = str(text)
        if cell:
            r, c = (int(x) for x in cell[1:].split("c"))
            tc = sh.table.cell(r, c)
            box = dict(w=sh.table.columns[c].width, h=sh.table.rows[r].height, pt=_font_pt(tc.text_frame), lines=lines.get(sid))
            set_text(tc.text_frame, text, emphasis.get(sid), True, box)
        elif sh.has_text_frame:
            if sw and sh.width < sw and (sh.left < 0 or sh.left + sh.width > sw):
                sh.left = min(max(sh.left, 0), sw - sh.width)  # a box hanging off the slide clips real text: pull it in
            box = dict(w=sh.width, h=sh.height, pt=_font_pt(sh.text_frame), lines=lines.get(sid), face=_inherited_face(sh))
            set_text(sh.text_frame, text, emphasis.get(sid), _shape_filled(sh), box)


def build(template_path, spec, out_path):
    """spec = {"slides": [{"template": n, "fills": {id: text}, "emphasis": {id: up|down|key}, "remove": [id]}]}"""
    prs = Presentation(template_path)
    originals = list(prs.slides)
    old_ids = list(prs.slides._sldIdLst)
    for item in spec["slides"]:
        src = originals[int(item["template"]) - 1]
        s = clone_slide(prs, src)
        fill_slide(s, item.get("fills", {}), item.get("emphasis"), item.get("remove", []), item.get("lines"), prs.slide_width)
    for sldId in old_ids:
        _drop_slide(prs, sldId)
    shapes = _chrome_shapes(prs)
    for cid, text in spec.get("chrome", {}).items():
        if cid in shapes:
            sh = shapes[cid]
            set_text(sh.text_frame, text, box=dict(w=sh.width, h=sh.height, pt=_font_pt(sh.text_frame), lines=1))
    prs.save(out_path)
    return out_path


def _preview_copy(pptx_path, out_dir):
    """Preview-only copy where every run holding CJK text names the CJK font. Templates often set the East Asian
    font to a Latin face (Inter); PowerPoint falls back to a CJK font on screen, LibreOffice sometimes drops the
    text instead, so the preview (and the template profile built from it) would miss it."""
    prs = Presentation(pptx_path)
    parts = [s.part for s in prs.slides] + [l.part for m in prs.slide_masters for l in m.slide_layouts]
    for part in parts:
        for r in part._element.iter(qn("a:r")):
            if any(_is_wide(ch) for ch in r.findtext(qn("a:t")) or ""):
                rPr = r.find(qn("a:rPr"))
                if rPr is None:
                    rPr = r.makeelement(qn("a:rPr"), {})
                    r.insert(0, rPr)
                _cjk_font(rPr, weight=False)  # as PowerPoint shows it: the font fallback, no added weight
    out = os.path.join(out_dir, os.path.basename(pptx_path))
    prs.save(out)
    return out


def _preview_substitutes(pptx_path):
    """{CJK typeface as the deck spells it: installed CJK sans}. Previews only.
    ponytail: Latin faces must be installed for exact previews (LibreOffice's fallback for a missing face is
    nondeterministic, sans one run, serif the next); mapping them by name broke CJK rendering, so don't."""
    faces = {CJK_FONT}
    with zipfile.ZipFile(pptx_path) as z:
        for name in z.namelist():
            if name.endswith(".xml") and name.startswith("ppt/"):
                faces |= {el.get("typeface") or "" for el in etree.fromstring(z.read(name)).iter(qn("a:ea"))}
    return {f: PREVIEW_CJK_FONT for f in faces if f.lower() == CJK_FONT.lower()}


# ---------- rendering ----------

def render(pptx_path, out_dir, dpi=80):
    """pptx -> PNG per slide via LibreOffice + pdftoppm. Returns sorted PNG paths."""
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):  # a version number reused after undo must not keep the old deck's pages
        if f.startswith("s-") and f.endswith(".png"):
            os.remove(os.path.join(out_dir, f))
    profile = "file://" + os.path.join(out_dir, ".lo")  # own profile: parallel soffice calls otherwise collide
    # Preview only: the deck's CJK font (Windows) is missing here and LibreOffice would fall back to a serif,
    # which misrepresents the design. Map it to an installed CJK sans; the .pptx itself keeps CJK_FONT.
    pptx_path = _preview_copy(pptx_path, out_dir)
    pairs = "".join(
        f'<node oor:name="_{i}" oor:op="replace"><prop oor:name="Always" oor:op="fuse"><value>true</value></prop>'
        f'<prop oor:name="OnScreenOnly" oor:op="fuse"><value>false</value></prop>'
        f'<prop oor:name="ReplaceFont" oor:op="fuse"><value>{escape(a)}</value></prop>'
        f'<prop oor:name="SubstituteFont" oor:op="fuse"><value>{escape(b)}</value></prop></node>'
        for i, (a, b) in enumerate(_preview_substitutes(pptx_path).items()))
    os.makedirs(os.path.join(out_dir, ".lo", "user"), exist_ok=True)
    with open(os.path.join(out_dir, ".lo", "user", "registrymodifications.xcu"), "w", encoding="utf-8") as f:
        f.write(f'''<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema">
<item oor:path="/org.openoffice.Office.Common/Font/Substitution"><prop oor:name="Replacement" oor:op="fuse"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Font/Substitution/FontPairs">{pairs}</item>
</oor:items>''')
    subprocess.run(["soffice", f"-env:UserInstallation={profile}", "--headless", "--convert-to", "pdf",
                    "--outdir", out_dir, pptx_path], check=True, capture_output=True, timeout=180)
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", pdf, os.path.join(out_dir, "s")], check=True, timeout=120)
    return sorted(os.path.join(out_dir, f) for f in os.listdir(out_dir) if f.startswith("s-") and f.endswith(".png"))


if __name__ == "__main__":
    # self-check: clone slide 1 twice + fill, confirm text landed and originals are gone
    import sys, tempfile
    tpl = sys.argv[1]
    info = analyze(tpl)
    first = info[0]["slots"][0]["id"]
    with tempfile.TemporaryDirectory() as d:
        out = os.path.join(d, "t.pptx")
        build(tpl, {"slides": [{"template": 1, "fills": {first: "HELLO 测试"}}, {"template": len(info)}]}, out)
        got = analyze(out)
    assert len(got) == 2, len(got)
    assert any(s["text"] == "HELLO 测试" for s in got[0]["slots"]), got[0]["slots"]
    tf = Presentation(tpl).slides[0].shapes[0].text_frame if Presentation(tpl).slides[0].shapes[0].has_text_frame else None
    if tf is not None:
        set_text(tf, "普通 **重点** 普通")
        assert [r.text for r in tf.paragraphs[0].runs] == ["普通 ", "重点", " 普通"], [r.text for r in tf.paragraphs[0].runs]
    print("ok")
