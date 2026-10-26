"""Write values + images into the original .xlsm by editing OOXML directly (keeps vbaProject, drawings, protection, validation, merges)."""
import zipfile, re, io, posixpath
from lxml import etree
from PIL import Image
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
      "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
M = "{%s}" % NS["m"]; R_DRAW = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
R_IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
AFTER_DRAWING = ["legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]

def col2n(s):
    n = 0
    for ch in s: n = n*26 + ord(ch)-64
    return n
def n2col(n):
    s = ""
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s
def split(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref); return m.group(1), int(m.group(2))
def rels_path(part): d, f = posixpath.split(part); return posixpath.join(d, "_rels", f + ".rels")
def resolve(base, target): return posixpath.normpath(posixpath.join(posixpath.dirname(base), target)) if not target.startswith("/") else target[1:]

def sheet_parts(z):
    wb = etree.fromstring(z.read("xl/workbook.xml")); rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rmap = {r.get("Id"): resolve("xl/workbook.xml", r.get("Target")) for r in rels}
    return {s.get("name"): rmap[s.get("{%s}id" % NS["r"])] for s in wb.find(M + "sheets")}

def set_cell(sheet_xml, ref, value):
    sd = sheet_xml.find(M + "sheetData"); col, rown = split(ref)
    row = None
    for rw in sd.findall(M + "row"):
        rn = int(rw.get("r"))
        if rn == rown: row = rw; break
        if rn > rown:
            row = etree.Element(M + "row", r=str(rown)); rw.addprevious(row); break
    if row is None: row = etree.SubElement(sd, M + "row", r=str(rown))
    c = None
    for cc in row.findall(M + "c"):
        cr = cc.get("r")
        if cr == ref: c = cc; break
        if col2n(split(cr)[0]) > col2n(col):
            c = etree.Element(M + "c", r=ref); cc.addprevious(c); break
    if c is None: c = etree.SubElement(row, M + "c", r=ref)
    if c.find(M + "f") is not None: return False  # never overwrite formulas
    for ch in list(c): c.remove(ch)
    if "t" in c.attrib: del c.attrib["t"]
    if value is None or value == "": return True
    try:
        num = float(str(value).replace(",", "")) if re.fullmatch(r"-?[\d,]*\.?\d+(e-?\d+)?", str(value).strip(), re.I) else None
    except ValueError: num = None
    if num is not None:
        etree.SubElement(c, M + "v").text = str(value).replace(",", "").strip()
    else:
        c.set("t", "inlineStr"); is_ = etree.SubElement(c, M + "is"); t = etree.SubElement(is_, M + "t"); t.text = str(value)
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return True

EMU = 9525
class Geometry:
    """Column widths / row heights of a worksheet in pixels (Excel defaults, 96 dpi)."""
    def __init__(self, x):
        fp = x.find(M + "sheetFormatPr"); self.dw = 8.43; self.dh = 15.0
        if fp is not None:
            if fp.get("defaultColWidth"): self.dw = float(fp.get("defaultColWidth"))
            elif fp.get("baseColWidth"): self.dw = float(fp.get("baseColWidth")) + 0.71
            if fp.get("defaultRowHeight"): self.dh = float(fp.get("defaultRowHeight"))
        self.cw = {}; self.rh = {}
        cols = x.find(M + "cols")
        if cols is not None:
            for c in cols:
                for i in range(int(c.get("min")), int(c.get("max")) + 1): self.cw[i] = 0 if c.get("hidden") == "1" else float(c.get("width", self.dw))
        sd = x.find(M + "sheetData")
        for r in sd.findall(M + "row"):
            if r.get("ht"): self.rh[int(r.get("r"))] = 0 if r.get("hidden") == "1" else float(r.get("ht"))
    def cpx(self, i): w = self.cw.get(i, self.dw); return 0 if w == 0 else int(w * 7 + 5)
    def rpx(self, i): return self.rh.get(i, self.dh) * 96 / 72
    def x0(self, col): return sum(self.cpx(i) for i in range(1, col))
    def y0(self, row): return sum(self.rpx(i) for i in range(1, row))
    def col_at(self, x):
        i, acc = 1, 0
        while acc + self.cpx(i) <= x and i < 16384: acc += self.cpx(i); i += 1
        return i - 1, x - acc
    def row_at(self, y):
        i, acc = 1, 0.0
        while acc + self.rpx(i) <= y and i < 1048576: acc += self.rpx(i); i += 1
        return i - 1, y - acc

def layout(geo, area, sizes, pad=6):
    """Fit N images into a cell range (grid, aspect preserved). Returns [(x, y, w, h)] in px."""
    a, b = area.split(":") if ":" in area else (area, area)
    (c1, r1), (c2, r2) = split(a), split(b); c1, c2 = col2n(c1), col2n(c2)
    X, Y = geo.x0(c1), geo.y0(r1); W = geo.x0(c2 + 1) - X; H = geo.y0(r2 + 1) - Y
    n = len(sizes); best = None
    for nc in range(1, n + 1):
        nr = -(-n // nc); cw, ch = W / nc, H / nr
        sc = min(min((cw - 2*pad) / w, (ch - 2*pad) / h, 1.0) for w, h in sizes)
        if best is None or sc > best[0]: best = (sc, nc, cw, ch)
    _, nc, cw, ch = best; out = []
    for i, (w, h) in enumerate(sizes):
        s = min((cw - 2*pad) / w, (ch - 2*pad) / h, 1.0); pw, ph = max(8, w*s), max(8, h*s)
        out.append((X + (i % nc)*cw + pad, Y + (i // nc)*ch + pad, pw, ph))
    return out

def export(template, out, values, images):
    """values: {sheet: {cell: value}}; images: {sheet: [(area 'B12:R18', image_bytes, name)]} – images are fitted into the area."""
    zin = zipfile.ZipFile(template); files = {i.filename: zin.read(i.filename) for i in zin.infolist()}; order = [i.filename for i in zin.infolist()]
    parts = sheet_parts(zin)
    ct = etree.fromstring(files["[Content_Types].xml"])
    def ensure_default(ext, mime):
        if not any(d.get("Extension", "").lower() == ext for d in ct.findall("{%s}Default" % NS["ct"])):
            etree.SubElement(ct, "{%s}Default" % NS["ct"], Extension=ext, ContentType=mime)
    report = {"written": 0, "skipped_formula": [], "images": 0}
    for sheet, cells in values.items():
        p = parts.get(sheet)
        if not p: continue
        x = etree.fromstring(files[p])
        for ref, v in cells.items():
            if set_cell(x, ref, v): report["written"] += 1
            else: report["skipped_formula"].append(f"{sheet}!{ref}")
        files[p] = etree.tostring(x, xml_declaration=True, encoding="UTF-8", standalone=True)
    media_n = 100; drawing_n = 100
    for sheet, imgs in images.items():
        if not imgs or sheet not in parts: continue
        p = parts[sheet]; x = etree.fromstring(files[p]); rp = rels_path(p)
        geo = Geometry(x)
        srels = etree.fromstring(files[rp]) if rp in files else etree.Element("{%s}Relationships" % NS["pr"], nsmap={None: NS["pr"]})
        drel = next((r for r in srels if r.get("Type") == R_DRAW), None)
        if drel is not None:
            dpath = resolve(p, drel.get("Target")); dx = etree.fromstring(files[dpath])
        else:
            while f"xl/drawings/drawing{drawing_n}.xml" in files: drawing_n += 1
            dpath = f"xl/drawings/drawing{drawing_n}.xml"; drawing_n += 1
            dx = etree.Element("{%s}wsDr" % NS["xdr"], nsmap={"xdr": NS["xdr"], "a": NS["a"], "r": NS["r"]})
            rid = "rId%d" % (1 + max([int(re.sub(r"\D", "", r.get("Id")) or 0) for r in srels] + [0]))
            etree.SubElement(srels, "{%s}Relationship" % NS["pr"], Id=rid, Type=R_DRAW, Target="../drawings/" + posixpath.basename(dpath))
            de = etree.Element(M + "drawing"); de.set("{%s}id" % NS["r"], rid)
            nxt = next((ch for ch in x if etree.QName(ch).localname in AFTER_DRAWING), None)
            (nxt.addprevious(de) if nxt is not None else x.append(de))
            etree.SubElement(ct, "{%s}Override" % NS["ct"], PartName="/" + dpath, ContentType="application/vnd.openxmlformats-officedocument.drawing+xml")
            files[p] = etree.tostring(x, xml_declaration=True, encoding="UTF-8", standalone=True)
        files[rp] = etree.tostring(srels, xml_declaration=True, encoding="UTF-8", standalone=True)
        drp = rels_path(dpath)
        drels = etree.fromstring(files[drp]) if drp in files else etree.Element("{%s}Relationships" % NS["pr"], nsmap={None: NS["pr"]})
        ids = [int(e.get("id")) for e in dx.iter("{%s}cNvPr" % NS["xdr"]) if (e.get("id") or "").isdigit()] + [1]
        nid = max(ids) + 1
        groups = {}
        for area, data, name in imgs: groups.setdefault(area, []).append((data, name))
        for area, items in groups.items():
            pics = []
            for data, name in items:
                im = Image.open(io.BytesIO(data)); im.load()
                if im.mode in ("P", "LA") or (im.mode == "RGBA" and im.getextrema()[3][0] < 255): im = im.convert("RGBA"); fmt = "png"
                else: im = im.convert("RGB"); fmt = "jpeg"
                pics.append((im, fmt, name))
            for (im, fmt, name), (px, py, pw, ph) in zip(pics, layout(geo, area, [p[0].size for p in pics])):
                lim = 1600; s = min(1.0, lim / max(im.size)); im2 = im.resize((max(1, int(im.width*s)), max(1, int(im.height*s)))) if s < 1 else im
                buf = io.BytesIO(); im2.save(buf, "PNG" if fmt == "png" else "JPEG", quality=85)
                ext = "png" if fmt == "png" else "jpeg"
                while f"xl/media/image{media_n}.{ext}" in files: media_n += 1
                mpath = f"xl/media/image{media_n}.{ext}"; files[mpath] = buf.getvalue(); order.append(mpath); media_n += 1
                rid = "rId%d" % (1 + max([int(re.sub(r"\D", "", r.get("Id")) or 0) for r in drels] + [0]))
                etree.SubElement(drels, "{%s}Relationship" % NS["pr"], Id=rid, Type=R_IMG, Target="../media/" + posixpath.basename(mpath))
                c0, co = geo.col_at(px); r0, ro = geo.row_at(py); cx, cy = int(pw * EMU), int(ph * EMU)
                A = lambda tag: "{%s}%s" % (NS["xdr"], tag)
                an = etree.SubElement(dx, A("oneCellAnchor"))
                fr = etree.SubElement(an, A("from"))
                for t, v in (("col", c0), ("colOff", int(co*EMU)), ("row", r0), ("rowOff", int(ro*EMU))): etree.SubElement(fr, A(t)).text = str(v)
                etree.SubElement(an, A("ext"), cx=str(cx), cy=str(cy))
                pic = etree.SubElement(an, A("pic")); nv = etree.SubElement(pic, A("nvPicPr"))
                etree.SubElement(nv, A("cNvPr"), id=str(nid), name=f"Picture {nid}", descr=name[:100]); nid += 1
                cn = etree.SubElement(nv, A("cNvPicPr")); etree.SubElement(cn, "{%s}picLocks" % NS["a"], noChangeAspect="1")
                bf = etree.SubElement(pic, A("blipFill")); b = etree.SubElement(bf, "{%s}blip" % NS["a"]); b.set("{%s}embed" % NS["r"], rid)
                st = etree.SubElement(bf, "{%s}stretch" % NS["a"]); etree.SubElement(st, "{%s}fillRect" % NS["a"])
                sp = etree.SubElement(pic, A("spPr")); xf = etree.SubElement(sp, "{%s}xfrm" % NS["a"])
                etree.SubElement(xf, "{%s}off" % NS["a"], x="0", y="0"); etree.SubElement(xf, "{%s}ext" % NS["a"], cx=str(cx), cy=str(cy))
                pg = etree.SubElement(sp, "{%s}prstGeom" % NS["a"], prst="rect"); etree.SubElement(pg, "{%s}avLst" % NS["a"])
                etree.SubElement(an, A("clientData"))
                report["images"] += 1
        ensure_default("png", "image/png"); ensure_default("jpeg", "image/jpeg")
        for pth, el in ((dpath, dx), (drp, drels)):
            if pth not in files: order.append(pth)
            files[pth] = etree.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)
        if rp not in order: order.append(rp)
    wbx = etree.fromstring(files["xl/workbook.xml"]); cp = wbx.find(M + "calcPr")
    if cp is None: cp = etree.SubElement(wbx, M + "calcPr")
    cp.set("fullCalcOnLoad", "1"); files["xl/workbook.xml"] = etree.tostring(wbx, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        zo.writestr("[Content_Types].xml", files["[Content_Types].xml"])
        for n in dict.fromkeys(order):
            if n != "[Content_Types].xml": zo.writestr(n, files[n])
    return report
