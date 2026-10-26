"""Fill every field/table/picture slot via the API, export, read back with openpyxl, compare. Then backup->restore and xlsm import round trips."""
import os, sys, io, json, tempfile, zipfile
tmp = tempfile.mkdtemp(); os.environ["DATA_DIR"] = tmp; os.environ.pop("ACCESS_CODE", None)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import server, openpyxl
from PIL import Image, ImageDraw
c = server.app.test_client(); fm = c.get("/api/ases/fieldmap").json; exp = {}
def img(label, w=800, h=500, color=(200, 30, 30)):
    im = Image.new("RGB", (w, h), color); d = ImageDraw.Draw(im); d.rectangle([5, 5, w-6, h-6], outline=(255, 255, 255), width=8)
    d.text((20, 20), label, fill=(255, 255, 255)); b = io.BytesIO(); im.save(b, "PNG"); return b.getvalue()
nimg = 0
for s in fm["sheets"]:
    n = s["name"]; e = exp.setdefault(n, {})
    for i, f in enumerate(s["fields"]):
        if f["type"] == "select": v = f["options"][-1] if f["options"][-1] != f.get("default") else f["options"][0]
        elif f["type"] == "number": v = str(i + 1.5)
        elif f["type"] == "email": v = f"a{i}@biu.ac.il"
        else: v = f"T-{f['cell']}"
        if n == "Additional Info": v = "line one\nline two"
        r = c.post("/api/ases/value", json={"sheet": n, "cell": f["cell"], "value": v}); assert r.status_code == 200, (n, f, r.json)
        e[f["cell"]] = float(v) if f["type"] == "number" else v
    for t in s.get("tables", []):
        rows = [[str(10*k + j + 0.25) for j in range(len(t["cols"]))] for k in range(3)]
        assert c.post("/api/ases/value", json={"sheet": n, "cell": "T:" + t["id"], "value": json.dumps(rows)}).status_code == 200
        col, row = server.exporter.split(t["start"])
        for k in range(3):
            for j in range(len(t["cols"])): e[server.col_shift(server.row_shift(t["start"], k), j)] = 10*k + j + 0.25
    for k, im in enumerate(s["images"]):
        files = [(io.BytesIO(img(f"{n} / {im['id']} #{q}", *( (600, 900) if q else (900, 500)), color=(180, 20+k*10 % 200, 40))), f"p{q}.png") for q in range(3 if k == 0 else 1)]
        r = c.post("/api/ases/image", data={"sheet": n, "slot": im["id"], "file": files}, content_type="multipart/form-data"); assert r.status_code == 200, r.json
        nimg += len(files)
# Additional Info: text lines + notes
c.post("/api/ases/notes", json={"sheet": "Cover Sheet", "text": "N/A because X"})
exp["Additional Info"] = {"B9": "line one", "B10": "line two", "B11": "[Cover Sheet] N/A because X"}
# invalid requests
assert c.post("/api/ases/value", json={"sheet": "Cover Sheet", "cell": "ZZ999", "value": "x"}).status_code == 400
ov = c.get("/api/ases/overview").json; assert all(o["pct"] == 100 for o in ov), [o for o in ov if o["pct"] != 100]
r = c.get("/api/ases/export"); assert r.status_code == 200; data = r.data
out = os.path.join(tmp, "out.xlsm"); open(out, "wb").write(data)
wb = openpyxl.load_workbook(out, keep_vba=True); bad = 0; tot = 0
for n, cells in exp.items():
    for cell, v in cells.items():
        tot += 1; got = wb[n][cell].value
        if not (got == v or (isinstance(v, float) and isinstance(got, (int, float)) and abs(got - v) < 1e-9)): bad += 1; print("MISMATCH", n, cell, repr(v), repr(got))
z = zipfile.ZipFile(out); media = [x for x in z.namelist() if x.startswith("xl/media/image") and int(''.join(ch for ch in x.split('/')[-1] if ch.isdigit()) or 0) >= 100]
print(f"cells checked {tot}, mismatches {bad}; images uploaded {nimg}, media in export {len(media)}; vba kept: {'xl/vbaProject.bin' in z.namelist()}")
assert bad == 0 and len(media) == nimg
open("/workspace/fsg/roundtrip_export.xlsm", "wb").write(data)
# backup -> wipe via restore of empty? restore same backup and compare overview
bk = c.get("/api/ases/backup").data; ov1 = c.get("/api/ases/overview").json
r = c.post("/api/ases/restore", data={"file": (io.BytesIO(bk), "b.json")}, content_type="multipart/form-data"); print("restore", r.json); assert r.json["ok"] and r.json["skipped"] == 0
assert c.get("/api/ases/overview").json == ov1
d2 = c.get("/api/ases/export").data; wb2 = openpyxl.load_workbook(io.BytesIO(d2))
assert all(wb2[n][cell].value == wb[n][cell].value for n, cells in exp.items() for cell in cells)
# import the exported xlsm into a fresh db
os.environ["DB_PATH"] = os.path.join(tmp, "fresh.db"); server.DB = os.environ["DB_PATH"]
r = c.post("/api/ases/import", data={"file": (io.BytesIO(data), "x.xlsm")}, content_type="multipart/form-data"); print("import", r.json)
d3 = c.get("/api/ases/export").data; wb3 = openpyxl.load_workbook(io.BytesIO(d3)); bad = 0
for n, cells in exp.items():
    for cell in cells:
        if wb3[n][cell].value != wb[n][cell].value: bad += 1; print("IMPORT MISMATCH", n, cell, wb[n][cell].value, wb3[n][cell].value)
print("import mismatches", bad); assert bad == 0
print("ALL OK")
