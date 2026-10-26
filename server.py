"""Multi-template FSG document filler. Templates live in templates/<id>/ (template file, fieldmap.json, overrides.json, rule_map.json); registry in templates/registry.json."""
import os, json, sqlite3, time, re
from flask import Flask, request, jsonify, send_file, send_from_directory, g, abort
import exporter, checks
HERE = os.path.dirname(os.path.abspath(__file__)); TD = os.path.join(HERE, "templates")
REG = json.load(open(os.path.join(TD, "registry.json")))
RULES = json.load(open(os.path.join(HERE, "data/rules_2027.json")))
T = {}
for tid, meta in REG.items():
    fmp = os.path.join(TD, tid, "fieldmap.json")
    if not meta.get("enabled") or not os.path.exists(fmp): continue
    fm = json.load(open(fmp)); rmp = os.path.join(TD, tid, "rule_map.json")
    T[tid] = {"meta": meta, "fm": fm, "sheets": {s["name"]: s for s in fm["sheets"]}, "file": os.path.join(TD, tid, meta["file"]),
              "rule_map": json.load(open(rmp)) if os.path.exists(rmp) else {}}
GUIDE_MAP = {"EV5.5.10-12 Cell Fixation": ["EV5.5.10-11 Guidance Notes", "EV5.5.12 Guidance Notes"],
             "EV5.5.9 Segment": ["EV5.5.9 Guidance Notes"], "EV5.5.9 TSAC": ["EV5.5.9 Guidance Notes"],
             "EV5.5.9 TSAC attachments": ["EV5.5 TSAC Attachment Guidance", "T3.5.9 Guidance Notes", "EV5.5.9 Guidance Notes"],
             "EV5.5.5 Alt. Matl - 3pt Bending": ["EV5.5.3 Guidance Notes"], "EV5.5.5 Alt. Matl - Shear": ["EV5.5.3 Guidance Notes"], "Accumulator Pics": ["EV5.5.3 Guidance Notes"]}
DD = os.environ.get("DATA_DIR", os.path.join(HERE, "data")); os.makedirs(DD, exist_ok=True)
DB = os.path.join(DD, "app.db"); UP = os.path.join(DD, "uploads") if os.environ.get("DATA_DIR") else os.path.join(HERE, "uploads"); os.makedirs(UP, exist_ok=True)
app = Flask(__name__, static_folder="static")
ACCESS = os.environ.get("ACCESS_CODE")
@app.before_request
def gate():
    if not ACCESS: return
    a = request.authorization
    if not a or a.password != ACCESS:
        from flask import Response
        return Response("Access code required", 401, {"WWW-Authenticate": 'Basic realm="BIU ASES"'})
def tpl(tid):
    if tid not in T: abort(404)
    return T[tid]
def db():
    if "db" not in g:
        g.db = sqlite3.connect(DB); g.db.row_factory = sqlite3.Row
        g.db.executescript("""create table if not exists vals(tpl text, sheet text, cell text, value text, user text, ts real, primary key(tpl,sheet,cell));
          create table if not exists imgs(id integer primary key, tpl text, sheet text, slot text, anchor text, file text, name text, user text, ts real);
          create table if not exists notes(tpl text, sheet text, text text, user text, ts real, primary key(tpl,sheet));
          create table if not exists log(ts real, tpl text, user text, sheet text, cell text, value text);""")
    return g.db
@app.teardown_appcontext
def close(e):
    d = g.pop("db", None)
    if d: d.commit(); d.close()
def state(tid, name):
    d = db()
    vals = {r["cell"]: r["value"] for r in d.execute("select cell,value from vals where tpl=? and sheet=?", (tid, name))}
    imgs = [dict(r) for r in d.execute("select id,slot,name,user,ts from imgs where tpl=? and sheet=?", (tid, name))]
    n = d.execute("select text from notes where tpl=? and sheet=?", (tid, name)).fetchone()
    return vals, imgs, (n["text"] if n else "")
def progress(tid, name):
    s = tpl(tid)["sheets"][name]; vals, imgs, _ = state(tid, name)
    tot = len(s["fields"]) + (1 if s["images"] else 0)
    done = sum(1 for f in s["fields"] if vals.get(f["cell"]) not in (None, "")) + (1 if s["images"] and imgs else 0)
    return {"done": done, "total": tot, "pct": round(100*done/tot) if tot else 100}
def findings(tid, name):
    t = tpl(tid); vals, imgs, notes = state(tid, name); cnt = {}
    for i in imgs: cnt[i["slot"]] = cnt.get(i["slot"], 0) + 1
    return checks.review(t["sheets"][name], vals, cnt, notes, None) + checks.rule_findings(name, t["rule_map"], RULES), vals
def rules_ctx(tid, name):
    rm = tpl(tid)["rule_map"].get(name) or {}
    return "\n".join(f"{r}: {RULES[r]}" for r in rm.get("rules_2027", []) if r in RULES) + (f"\nNUMBERING NOTE: {rm['mismatch']}" if rm.get("mismatch") else "")
@app.get("/")
def index(): return send_from_directory(HERE + "/static", "index.html")
@app.get("/api/templates")
def templates(): return jsonify({k: {**v, "loaded": k in T} for k, v in REG.items()})
@app.get("/api/rules/<rid>")
def rule(rid): r = RULES.get(rid.replace(" ", "")); return jsonify({"id": rid, "text": r}) if r else (jsonify({"error": "not found"}), 404)
@app.get("/api/<tid>/fieldmap")
def fieldmap(tid):
    t = tpl(tid); return jsonify({"template": t["fm"]["template"], "title": t["meta"]["title"],
        "sheets": [{**{k: s[k] for k in ("name", "title_he", "fields", "images")}, "rules": t["rule_map"].get(s["name"], {})} for s in t["fm"]["sheets"]]})
@app.get("/api/<tid>/progress")
def all_progress(tid): return jsonify({n: progress(tid, n) for n in tpl(tid)["sheets"]})
@app.get("/api/<tid>/sheet/<path:name>")
def get_sheet(tid, name):
    vals, imgs, notes = state(tid, name); return jsonify({"values": vals, "images": imgs, "notes": notes, "progress": progress(tid, name)})
@app.post("/api/<tid>/value")
def set_value(tid):
    j = request.json; d = db(); user = j.get("user") or "anon"; S = tpl(tid)["sheets"]
    if j["sheet"] not in S or j["cell"] not in {f["cell"] for f in S[j["sheet"]]["fields"]}: return jsonify({"error": "unknown field"}), 400
    d.execute("insert or replace into vals values(?,?,?,?,?,?)", (tid, j["sheet"], j["cell"], j.get("value"), user, time.time()))
    d.execute("insert into log values(?,?,?,?,?,?)", (time.time(), tid, user, j["sheet"], j["cell"], j.get("value")))
    return jsonify({"ok": True, "progress": progress(tid, j["sheet"])})
@app.post("/api/<tid>/notes")
def set_notes(tid):
    j = request.json; db().execute("insert or replace into notes values(?,?,?,?,?)", (tid, j["sheet"], j.get("text", ""), j.get("user"), time.time())); return jsonify({"ok": True})
@app.post("/api/<tid>/image")
def up_image(tid):
    sheet, slot = request.form["sheet"], request.form["slot"]; f = request.files["file"]
    s = tpl(tid)["sheets"].get(sheet); sl = next((i for i in (s or {}).get("images", []) if i["id"] == slot), None)
    if sl is None: return jsonify({"error": "unknown slot"}), 400
    if not re.search(r"\.(png|jpe?g|gif|bmp|webp)$", f.filename, re.I): return jsonify({"error": "רק קבצי תמונה (png/jpg)"}), 400
    fn = f"{int(time.time()*1000)}_{re.sub(r'[^A-Za-z0-9._-]', '_', f.filename)}"; f.save(os.path.join(UP, fn))
    db().execute("insert into imgs(tpl,sheet,slot,anchor,file,name,user,ts) values(?,?,?,?,?,?,?,?)", (tid, sheet, slot, sl["anchor"], fn, f.filename, request.form.get("user"), time.time()))
    return jsonify({"ok": True})
@app.delete("/api/<tid>/image/<int:iid>")
def del_image(tid, iid): db().execute("delete from imgs where id=? and tpl=?", (iid, tid)); return jsonify({"ok": True})
@app.get("/api/<tid>/image/<int:iid>")
def get_image(tid, iid):
    r = db().execute("select file from imgs where id=? and tpl=?", (iid, tid)).fetchone(); return send_from_directory(UP, r["file"]) if r else abort(404)
@app.get("/api/<tid>/check/<path:name>")
def check(tid, name): return jsonify(findings(tid, name)[0])
@app.post("/api/<tid>/chat")
def chat(tid):
    j = request.json; name = j["sheet"]; f, vals = findings(tid, name)
    guid = "\n".join(tpl(tid)["fm"]["guidance"].get(k, "") for k in GUIDE_MAP.get(name, []) + ["Read_Me"])
    return jsonify({"answer": checks.assistant(tpl(tid)["sheets"][name], j["message"], vals, [x for x in f if x["level"] != "rule"], guid, j.get("history"), rules_ctx(tid, name)), "llm": bool(checks.llm_config())})
def build_export(tid, path):
    d = db(); values, images = {}, {}; S = tpl(tid)["sheets"]
    for r in d.execute("select sheet,cell,value from vals where tpl=?", (tid,)): values.setdefault(r["sheet"], {})[r["cell"]] = r["value"]
    for r in d.execute("select sheet,anchor,file,name from imgs where tpl=? order by ts", (tid,)):
        images.setdefault(r["sheet"], []).append((r["anchor"], open(os.path.join(UP, r["file"]), "rb").read(), r["name"]))
    notes = [dict(r) for r in d.execute("select sheet,text from notes where tpl=? and text<>''", (tid,))]
    if notes and "Additional Info" in S:  # N/A justifications -> Additional Info tab
        ai = values.setdefault("Additional Info", {})
        for i, n in enumerate(notes): ai[f"B{6+i}"] = f"[{n['sheet']}] {n['text']}"
        images["Additional Info"] = [(f"B{9+len(notes)}" if a == "B8" else a, b, c) for a, b, c in images.get("Additional Info", [])]
    return exporter.export(tpl(tid)["file"], path, values, images)
@app.get("/api/<tid>/export")
def export(tid):
    ext = os.path.splitext(tpl(tid)["file"])[1]; path = os.path.join(DD, f"export_{tid}{ext}"); build_export(tid, path)
    return send_file(path, as_attachment=True, download_name=f"BIU_{tid.upper()}_filled{ext}")
if __name__ == "__main__":
    app.run(host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", 8077)), debug=False)
