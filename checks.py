"""Rule-based section review + pluggable LLM assistant."""
import re, os, json, urllib.request
TEAM_GUIDE = {  # from the team's internal PDF guide (FSG 2027 prep) – ASES chapter
 "general": ["Use only the official ASES template of the competition year; do not delete or rename tabs.",
             "Use SI units everywhere.", "Put CAD, pictures, datasheets, calculations and test results in the designated places.",
             "Document any official Rules Clarification you rely on (Rules Clarifications tab).",
             "Internal deadlines: first full version T-21 days, engineering review T-14, upload-ready T-7.",
             "Every TSAC must be openable for inspection; each segment removable in < 15 min without moving/tilting the TSAC.",
             "What is submitted must match the car actually built (and SE3D)."],
 "EV5.5.10-12 Cell Fixation": ["Mechanical stress on a cell must not exceed 0.05 MPa shear and 1.0 MPa compression (or the datasheet limit if stricter).",
             "Friction-based fixation requires physical friction testing (setup pictures, procedure, raw data, results).",
             "Pouch cells: fixed on the large surface(s) only, ≥80 % of each used surface; tabs must not carry loads."],
 "EV5.5.13 Fasteners": ["Document type/size, location, use, positive-locking method and CAD/pictures for every fastener group.",
             "Non-standard attachments that need physical tests: document the test results in SES/ASES."],
 "EV5.5.14-17 Openings&Separation": ["Document every opening/cutout, wiring penetrations, cooling/ventilation ducts, sealing, internal walls.",
             "AIRs and main fuse must be separated from the cell segments.", "Show the effect of openings on the structure."],
 "Accumulator Pics": ["Provide CAD/pictures with dimensions: floor, walls, cover thicknesses; position of the accumulator in the chassis."],
 "EV5.5.5 Alt. Matl - 3pt Bending": ["Additively manufactured structural parts need physical tests and anisotropic properties."],
 "Datasheets Fiber": ["Attach datasheets with mechanical properties (also at required temperatures) and fire-retardant evidence."],
 "Datasheets Other Matl": ["Attach datasheets with mechanical properties incl. at 60°C and fire rating evidence."],
 "Datasheets Glue": ["Attach adhesive datasheets incl. shear strength at 60°C."],
}
SEGMENT_LIMITS = {"V": 120.0, "MJ": 6.0, "kg": 12.0}  # FS Rules 2027 EV5.3.4 (per mechanical SECTION in 2027)

def num(v):
    try: return float(str(v).replace(",", "."))
    except (TypeError, ValueError): return None

def rule_findings(name, rule_map, rules):
    out = []; rm = rule_map.get(name) or {}
    if rm.get("mismatch"): out.append({"level": "warn", "msg": "⚠ אי-התאמת מספור (תבנית 2026 מול FS Rules 2027): " + rm["mismatch"]})
    for r in rm.get("rules_2027", []):
        if r in rules: out.append({"level": "rule", "rule": r, "msg": f"{r} (2027): {rules[r][:600]}"})
    return out

def review(sheet, vals, imgs, notes, all_vals):
    """sheet: fieldmap sheet dict; vals: {cell: value}; imgs: {slot_id: count}; returns list of {level, msg}"""
    out = []; add = lambda lvl, m: out.append({"level": lvl, "msg": m})
    name = sheet["name"]; F = {f["cell"]: f for f in sheet["fields"]}
    g = lambda c: (vals.get(c) if vals.get(c) not in (None, "") else F.get(c, {}).get("default"))
    empty = [f for f in sheet["fields"] if vals.get(f["cell"]) in (None, "") and f["type"] != "select"]
    if name not in ("Version History", "Rules Clarifications", "Datasheets Fiber", "Datasheets Other Matl", "Datasheets Glue", "EV5.5.13 Fasteners", "EV5.5.14-17 Openings&Separation", "EV5.5.5 Alt. Matl - 3pt Bending", "EV5.5.5 Alt. Matl - Shear") and empty:
        add("warn", f"{len(empty)} שדות ריקים: " + ", ".join(f'{f["label"][:30]} ({f["cell"]})' for f in empty[:8]))
    for f in sheet["fields"]:
        v = vals.get(f["cell"])
        if f["type"] == "number" and v not in (None, "") and num(v) is None: add("error", f'"{f["label"][:40]}" ({f["cell"]}) חייב להיות מספר ביחידות {f.get("unit") or "SI"}')
        if v == "N/A" and not (notes or "").strip(): add("warn", f'נבחר N/A ב-"{f["label"][:40]}" ({f["cell"]}) – חובה להסביר בקצרה למה (שדה "הצדקות / הערות")')
    if sheet["images"] and sum(imgs.values()) == 0 and name not in ("Additional Info",):
        add("warn", "לא הועלו תמונות/CAD בלשונית הזו – ה-ASES דורש ראיות ויזואליות (CAD/תמונות עם מידות).")
    if name == "Cover Sheet":
        for c in ("N16", "N17"):
            if vals.get(c) and not re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+$", vals[c]): add("error", f"כתובת מייל לא תקינה ב-{c}")
        nos = [f for f in sheet["fields"] if f["type"] == "select" and (g(f["cell"]) or "No") == "No"]
        if nos: add("warn", f'{len(nos)} שורות בטבלת הציות עדיין "No" (Compliance shown) – לפני הגשה כל שורה רלוונטית צריכה להיות Yes, ו-N/A מוסבר.')
    if name == "EV5.5.10-12 Cell Fixation":
        m, h, w, n, Fc, mu = (num(g(c)) for c in ("M15", "M16", "M17", "M20", "M22", "M25"))
        typ, fix = g("M12") or "Pouch", g("M13") or "Friction based"
        if None not in (m, h, w, n) and h * w > 0:
            area = h * w * 0.8; shear = m/1000*40*9.81*n/area
            add("ok" if shear <= 0.05 else ("warn" if n <= 6 else "error"), f"Shear stress מחושב = {shear:.4f} MPa (גבול 0.05 MPa – EV5.5.14 ב-2027){'' if shear<=0.05 else ' – חורג! ' + ('n≤6: ייתכן EXCEPTION לפי התבנית, נדרש נימוק' if n<=6 else 'FAIL')}")
            if Fc is not None:
                comp = Fc/area; add("ok" if comp <= 1.0 else "error", f"Compression stress = {comp:.3f} MPa (גבול 1.0 MPa – EV5.5.14)")
            if fix == "Friction based":
                if mu is None: add("error", "קיבוע מבוסס חיכוך: חסר מקדם חיכוך מניסוי (M25) – נדרש physical friction testing (EV5.5.15 ב-2027).")
                elif Fc is not None:
                    need = m*n/1000*40*9.81; ff = Fc*mu
                    add("ok" if ff > need else "error", f"כוח חיכוך {ff:.1f} N מול נדרש {need:.1f} N (40g) → {'PASS' if ff>need else 'FAIL'}")
        else: add("warn", "חסרים מסת תא / מידות / n – אי אפשר לחשב את מאמצי הגזירה והלחיצה.")
        if g("M28") == "Yes": add("warn", "נבחר Yes לחומרים דחיסים (PCB) – יש להסביר איך EV4.5.12 מתקיים.")
        if typ == "Pouch" and g("M28") == "N/A": add("error", "לתאי Pouch אסור לבחור N/A בשאלת החומרים הדחיסים.")
        if fix == "Other" and not g("M14"): add("error", "נבחר Fixation=Other – חובה לפרט (M14).")
    if name == "EV5.5.14-17 Openings&Separation":
        for f in sheet["fields"]:
            v = num(vals.get(f["cell"])); u = (f.get("unit") or "")
            if v is not None and u in SEGMENT_LIMITS and v > SEGMENT_LIMITS[u]:
                add("error", f'{f["label"]}: {v} {u} חורג מגבול ({SEGMENT_LIMITS[u]} {u}) – FS Rules 2027 EV5.3.4. שימו לב: ב-2027 הגבול חל על section מכני (EV5.1.4) ולא על segment.')
        nos = [f for f in sheet["fields"] if f["type"] == "select" and f["cell"].startswith("R") and (g(f["cell"]) or "No") == "No"]
        if nos: add("warn", f"{len(nos)} דרישות פתחים/הפרדה עדיין No – הוסיפו תמונות ותיעוד והפכו ל-Yes.")
    if name == "EV5.5.13 Fasteners":
        blocks = {}
        for f in sheet["fields"]: blocks.setdefault((f["cell"][:2].rstrip("0123456789"), int(re.sub(r"\D", "", f["cell"])) // 24), []).append(f)
        used = [b for b in blocks.values() if any(vals.get(f["cell"]) for f in b)]
        if not used: add("warn", "לא מולאה אף קבוצת fasteners.")
        for b in used:
            for f in b:
                if not vals.get(f["cell"]): add("warn", f'קבוצת fasteners חלקית: חסר "{f["label"]}" ({f["cell"]})')
    if name.startswith("Datasheets"):
        for f in sheet["fields"]:
            if f["type"] == "select" and vals.get(f["cell"]) == "NO": add("warn", f'"{f["label"]}" = NO ב-{f["cell"]} – צרפו datasheet או ניסוי עצמי.')
    if name == "Version History" and not any(vals.values()): add("info", "אם זו הגשה חוזרת – חובה לתעד את השינויים כאן.")
    for tip in TEAM_GUIDE.get(name, []): add("tip", tip)
    if not [o for o in out if o["level"] in ("error", "warn")]: add("ok", "לא נמצאו בעיות בבדיקה האוטומטית.")
    return out

def llm_config():
    if os.environ.get("XAI_API_KEY"): return ("https://api.x.ai/v1/chat/completions", os.environ["XAI_API_KEY"], os.environ.get("LLM_MODEL", "grok-4"))
    if os.environ.get("OPENAI_API_KEY"): return (os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1") + "/chat/completions", os.environ["OPENAI_API_KEY"], os.environ.get("LLM_MODEL", "gpt-4o-mini"))
    return None

def assistant(sheet, question, vals, findings, guidance, history=None, rules_ctx=""):
    ctx = "\n".join(TEAM_GUIDE["general"] + TEAM_GUIDE.get(sheet["name"], [])) + "\nFS Rules 2027 (authoritative; template numbering is 2026):\n" + rules_ctx
    filled = {f'{f["label"]} [{f.get("col_header") or ""}] ({f["cell"]})': vals.get(f["cell"]) for f in sheet["fields"] if vals.get(f["cell"])}
    cfg = llm_config()
    if cfg:
        url, key, model = cfg
        sysmsg = ("You help a Formula Student EV team fill the FSG ASES (Accumulator Structural Equivalency Spreadsheet) so it is not rejected. "
                  "Answer in Hebrew, keep rule numbers/technical terms in English. Be concrete: what is missing, what evidence to add, which rule. "
                  f"Current tab: {sheet['name']}.\nOfficial guidance text from the template:\n{guidance[:6000]}\nTeam checklist:\n{ctx}\n"
                  f"Automatic check findings: {json.dumps(findings, ensure_ascii=False)}\nEntered data: {json.dumps(filled, ensure_ascii=False)[:4000]}")
        msgs = [{"role": "system", "content": sysmsg}] + (history or [])[-8:] + [{"role": "user", "content": question}]
        req = urllib.request.Request(url, data=json.dumps({"model": model, "messages": msgs, "temperature": 0.2}).encode(),
                                     headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=60) as r: return json.load(r)["choices"][0]["message"]["content"]
        except Exception as e: return f"שגיאת LLM: {e}"
    # offline fallback: keyword search in guidance + findings
    rq = [r.replace(" ", "") for r in re.findall(r"\b(?:EV|T|EV ?)\s?\d+(?:\.\d+)+", question, re.I)]
    hits = [l for l in rules_ctx.split("\n") if any(l.upper().startswith(q.upper() + ":") for q in rq)]
    words = [w.lower() for w in re.findall(r"[A-Za-z][A-Za-z0-9.]{2,}", question)]
    lines = [l for l in (guidance + "\n" + ctx).split("\n") if l.strip() and any(w in l.lower() for w in words)][:8]
    msg = "🔌 מצב לא-מקוון (לא הוגדר XAI_API_KEY / OPENAI_API_KEY) – תשובה מבוססת חוקים וטקסט ההנחיות:\n"
    msg += "\n".join("• " + f["msg"] for f in findings if f["level"] in ("error", "warn"))[:2000] or "• אין ממצאים פתוחים."
    if hits: msg += "\n\nכלל מהתקנון 2027:\n" + "\n".join(hits)
    if lines: msg += "\n\nקטעים רלוונטיים מההנחיות:\n" + "\n".join("– " + l.strip()[:300] for l in lines)
    return msg
