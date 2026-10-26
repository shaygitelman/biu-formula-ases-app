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

def export(template, out, values, images):
    """values: {sheet: {cell: value}}; images: {sheet: [(anchor_cell, png_bytes, name)]}"""
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
        nid = max(ids) + 1; stack = {}
        for anchor, data, name in imgs:
            im = Image.open(io.BytesIO(data)); im = im.convert("RGB") if im.mode not in ("RGB", "RGBA") else im
            w, h = im.size; scale = min(1.0, 640 / w, 480 / h); w, h = int(w*scale), int(h*scale)
            buf = io.BytesIO(); im.resize((w, h)).save(buf, "PNG")
            while f"xl/media/image{media_n}.png" in files: media_n += 1
            mpath = f"xl/media/image{media_n}.png"; files[mpath] = buf.getvalue(); order.append(mpath); media_n += 1
            rid = "rId%d" % (1 + max([int(re.sub(r"\D", "", r.get("Id")) or 0) for r in drels] + [0]))
            etree.SubElement(drels, "{%s}Relationship" % NS["pr"], Id=rid, Type=R_IMG, Target="../media/" + posixpath.basename(mpath))
            col, row = split(anchor); off = stack.get(anchor, 0)
            cx, cy = w * 9525, h * 9525
            A = lambda tag: "{%s}%s" % (NS["xdr"], tag)
            an = etree.SubElement(dx, A("oneCellAnchor"))
            fr = etree.SubElement(an, A("from"))
            for t, v in (("col", col2n(col)-1), ("colOff", 0), ("row", row-1 + off), ("rowOff", 0)): etree.SubElement(fr, A(t)).text = str(v)
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
            stack[anchor] = off + int(cy / 190500) + 1  # next picture below (default row ~15pt)
            report["images"] += 1
        ensure_default("png", "image/png")
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
