"""Rebuild the field map from an ASES template (data-driven, no hard-coded cells).
Usage: python3 build_fieldmap.py [template.xlsm] [out.json]
Heuristics: inputs = unlocked cells (sheet protection), label = nearest locked text to the left in the row,
column header = nearest locked text above, dropdowns = inline-list data validations, picture slots = cells whose
text says 'paste image/insert pictures/...'. Per-cell corrections live in data/overrides.json."""
import openpyxl, json, re, sys, os
from openpyxl.utils import get_column_letter as L, range_boundaries
HERE = os.path.dirname(os.path.abspath(__file__))
INPUT_SHEETS = {
 "Cover Sheet":"דף שער ופרטי צוות","Accumulator Pics":"תמונות ומידות המצבר","EV5.5.5 Alt. Matl - 3pt Bending":"חומר חלופי – כפיפה 3 נקודות",
 "EV5.5.5 Alt. Matl - Shear":"חומר חלופי – גזירה","Datasheets Fiber":"דאטהשיטים – סיבים","Datasheets Other Matl":"דאטהשיטים – חומרים אחרים",
 "Datasheets Glue":"דאטהשיטים – דבקים","EV5.5.10-12 Cell Fixation":"קיבוע תאים","EV5.5.9 Segment":"Segment","EV5.5.9 TSAC":"TSAC",
 "EV5.5.9 TSAC attachments":"חיבורי TSAC לשלדה","EV5.5.13 Fasteners":"מחברים (Fasteners)","EV5.5.14-17 Openings&Separation":"פתחים והפרדות",
 "Additional Info":"מידע נוסף","Version History":"היסטוריית גרסאות","Rules Clarifications":"הבהרות חוקים"}
IMG_RE = re.compile(r"(paste image|paste picture|insert picture|attach picture|show \(cad\) picture|insert .*pictures|any pictures|picture of measured|paste images)", re.I)
INSTR_RE = re.compile(r"(paste|insert|please|show|add |guidance|- )", re.I)
BADLAB = re.compile(r"^(guidance|please|insert|paste|yes|no|n/a|x|y|z|tab|separation shown|compliance shown)\b|BACK to", re.I)
UNIT_RE = [re.compile(r"\n\s*([A-Za-z]+)\s*$"), re.compile(r"\(([A-Za-zµ²³%°/^0-9]+)\)\s*$"), re.compile(r"\[([^\]]+)\]\s*$"),
           re.compile(r"/\s*([A-Za-zµ²³%°]+[A-Za-z0-9²³/^]*)\s*\)?\s*$")]
def unit_of(lbl):
    for r in UNIT_RE:
        m = r.search(lbl or "")
        if m and m.group(1).strip() not in ("s", "es"): return m.group(1).strip()
    return None
def build(path, overrides):
    wb = openpyxl.load_workbook(path, keep_vba=True); wv = openpyxl.load_workbook(path, data_only=True)
    out = {"template": os.path.basename(path), "sheets": [], "guidance": {}}
    for ws in wv.worksheets:
        if "Guidance" in ws.title or ws.title == "Read_Me":
            out["guidance"][ws.title] = "\n".join(str(c.value) for r in ws.iter_rows() for c in r if isinstance(c.value, str))
    order = [n for n in INPUT_SHEETS if n in wb.sheetnames] + [n for n in wb.sheetnames if n not in INPUT_SHEETS and wb[n].sheet_state == "visible" and "Guidance" not in n and n != "Read_Me"]
    for name in order:
        ws = wb[name]; vs = wv[name]
        top, mrg = {}, {}
        for m in ws.merged_cells.ranges:
            mrg[(m.min_row, m.min_col)] = m.coord
            for r in range(m.min_row, m.max_row+1):
                for c in range(m.min_col, m.max_col+1):
                    if (r, c) != (m.min_row, m.min_col): top[(r, c)] = True
        dv = {}
        for d in ws.data_validations.dataValidation:
            f1 = d.formula1 or ""
            if d.type != "list" or not f1.startswith('"'): continue
            opts = [o.strip() for o in f1.strip('"').split(",") if o.strip()]
            for rg in str(d.sqref).split():
                a, b, c2, e = range_boundaries(rg)
                for r in range(b, e+1):
                    for c in range(a, c2+1): dv[(r, c)] = opts
        maxc = ws.max_column
        def hdr_of(r, c, lab, near=20):
            for rr in range(r-1, max(0, r-near), -1):
                hc = ws.cell(rr, c); hv = vs.cell(rr, c).value
                if hc.protection.locked is False and (rr, c) not in top:
                    if hc.value is not None: return None
                    continue
                if isinstance(hv, str) and hv.strip() and not hv.startswith("#") and hv.strip() != lab and len(hv) < 60 and not IMG_RE.search(hv) and "BACK to" not in hv:
                    return hv.strip()
            return None
        fields, images = [], []
        for r in range(1, ws.max_row+1):
            label, lcol, nfield, clean = None, 0, 0, True
            for c in range(1, maxc+1):
                if (r, c) in top: continue
                cell = ws.cell(r, c); v = cell.value; unl = cell.protection.locked is False
                txt = vs.cell(r, c).value
                if isinstance(txt, str) and IMG_RE.search(txt.strip()):
                    images.append({"id": cell.coordinate, "anchor": f"{L(c)}{r+1}", "label": txt.strip()[:200]})
                if not unl:
                    if v is None: continue
                    if isinstance(v, str) and not v.startswith("=") and v.strip() and not BADLAB.search(v.strip()) and not IMG_RE.search(v):
                        label, lcol, nfield, clean = v.strip(), c, 0, True
                    else: clean = False
                    continue
                if isinstance(v, str) and (v.startswith("=") or (len(v) > 25 and INSTR_RE.search(v))): continue
                key = (r, c)
                if (label is None or c - lcol > 30) and key not in dv: continue
                hdr = hdr_of(r, c, label, 20 if key in mrg else 9)
                if label is None: label, lcol, nfield, clean = (hdr or "Material type"), c, 0, False
                if not ((nfield == 0 and clean) or hdr or key in dv): continue
                unit = unit_of(label) or unit_of(hdr or "")
                num = unit is not None or isinstance(v, (int, float))
                fields.append({"cell": f"{L(c)}{r}", "range": mrg.get(key, f"{L(c)}{r}"), "label": label, "col_header": hdr,
                               "type": "select" if key in dv else ("number" if num else "text"), "options": dv.get(key), "unit": unit, "default": v})
                nfield += 1
        if not fields:  # table-style sheet
            hr = next((r for r in range(1, 15) if any(isinstance(ws.cell(r, c).value, str) for c in range(2, 6)) and ws.cell(r+1, 1).protection.locked is False), None)
            if hr:
                cols = [c for c in range(1, maxc+1) if isinstance(ws.cell(hr, c).value, str)]
                for i in range(10):
                    for c in cols:
                        fields.append({"cell": f"{L(c)}{hr+1+i}", "range": f"{L(c)}{hr+1+i}", "label": f"#{i+1}", "col_header": ws.cell(hr, c).value,
                                       "type": "text" if c == cols[0] else "textarea", "options": None, "unit": None, "default": None})
        ov = overrides.get(name, {})
        for f in fields: f.update(ov.get("fields", {}).get(f["cell"], {}))
        fields = [f for f in fields if not f.get("skip")] + ov.get("extra_fields", [])
        images += ov.get("extra_images", [])
        if not images and ov.get("canvas"): images.append({"id": ov["canvas"], "anchor": ov["canvas"], "label": "Pictures / CAD / calculations / explanations"})
        # repeated blocks (e.g. 8 material columns, 12 fastener groups): number them by column position
        from collections import defaultdict
        from openpyxl.utils import column_index_from_string as CI
        bylab = defaultdict(set)
        for f in fields: bylab[f["label"]].add(CI(re.match(r"[A-Z]+", f["cell"]).group()))
        starts = sorted({min(v) for v in []})
        for f in fields:
            cols = sorted(bylab[f["label"]]); col = CI(re.match(r"[A-Z]+", f["cell"]).group())
            spread = len(cols) > 1 and min(b - a for a, b in zip(cols, cols[1:])) >= 8
            f["block"] = cols.index(col) + 1 if spread and f["label"] != "Material" else 0
        colhdr = {}
        for f in fields:
            if f["col_header"]: colhdr.setdefault(re.match(r"[A-Z]+", f["cell"]).group(), f["col_header"])
        for f in fields:
            if not f["col_header"]: f["col_header"] = colhdr.get(re.match(r"[A-Z]+", f["cell"]).group())
        for f in fields: f["default"] = None if f["default"] is None else str(f["default"])
        if not fields and not images: continue
        out["sheets"].append({"name": name, "title_he": INPUT_SHEETS.get(name, name), "fields": fields, "images": images, "notes": ov.get("notes_he")})
    return out
if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "templates/ases/FS26_ASES_v1.0.xlsm")
    dst = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(src), "fieldmap.json")
    ovp = os.path.join(os.path.dirname(src), "overrides.json")
    fm = build(src, json.load(open(ovp)) if os.path.exists(ovp) else {})
    json.dump(fm, open(dst, "w"), ensure_ascii=False, indent=1)
    for s in fm["sheets"]: print(f'{s["name"]:34} fields={len(s["fields"]):3} images={len(s["images"])}')
