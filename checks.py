"""Rule-based section review (English) + pluggable LLM assistant."""
import re, os, json, urllib.request
TEAM_GUIDE = {  # team's internal FSG-2027 preparation guide – ASES chapter
 "general": ["Use only the official ASES template of the competition year; do not delete or rename tabs.",
             "Use SI units everywhere.", "Put CAD, pictures, datasheets, calculations and test results in the designated places.",
             "Document every official Rules Clarification you rely on (Rules Clarifications tab).",
             "Internal deadlines: first full version T-21 days, engineering review T-14, upload-ready T-7.",
             "Every TSAC must be openable for inspection; each segment removable in < 15 min without moving/tilting the TSAC (EV5.3.3).",
             "What is submitted must match the car actually built."],
 "EV5.5.10-12 Cell Fixation": ["Cell stress must not exceed 0.05 MPa shear and 1.0 MPa compression, or the datasheet limit if lower (2027: EV5.5.14).",
             "Friction-based fixation requires physical friction testing – setup pictures, procedure, raw data, results (2027: EV5.5.15).",
             "Pouch cells: fixed on the large surface(s) only, ≥ 80 % of each used surface; tabs must not carry loads (2027: EV5.5.12)."],
 "EV5.5.13 Fasteners": ["Document type/size, location, use, positive-locking method and CAD/pictures for every fastener group (2027: EV5.5.16, T10).",
             "Structural fasteners must follow T10 completely (grade 8.8, ≥ M4, positive locking)."],
 "EV5.5.14-17 Openings&Separation": ["Document every opening/cutout, wiring penetration, cooling/ventilation duct, sealing and internal walls.",
             "AIRs and main fuse must be separated from the cell segments (2027: EV5.5.17)."],
 "Accumulator Pics": ["Provide CAD/pictures with dimensions: floor, walls and cover thicknesses; position of the accumulator in the chassis."],
 "EV5.5.5 Alt. Matl - 3pt Bending": ["Additively manufactured structural parts need physical tests (2027: EV5.5.10)."],
 "Datasheets Fiber": ["Attach datasheets showing mechanical properties and fire-retardancy evidence."],
 "Datasheets Other Matl": ["Attach datasheets with mechanical properties incl. at 60 °C and fire-rating evidence."],
 "Datasheets Glue": ["Attach adhesive datasheets incl. shear strength at 60 °C."],
}
SECTION_LIMITS = {"V": 120.0, "MJ": 6.0, "kg": 12.0}  # FS Rules 2027 EV5.3.4 – per mechanical SECTION
EMAIL = re.compile(r"[^@\s]+@[^@\s]+\.[^@\s]+$")

def num(v):
    try: return float(str(v).replace(",", "."))
    except (TypeError, ValueError): return None

def rule_findings(name, rule_map, rules):
    out = []; rm = rule_map.get(name) or {}
    if rm.get("mismatch"): out.append({"level": "warn", "msg": "Rule numbering: this template uses 2026 numbers, FS Rules 2027 apply. " + rm["mismatch"]})
    for r in rm.get("rules_2027", []):
        if r in rules: out.append({"level": "rule", "rule": r, "msg": f"{r} (2027): {rules[r][:600]}"})
    return out

def field_hint(f, v):
    """Inline validation for one field value -> message or None."""
    if v in (None, ""): return None
    if f["type"] == "number" and num(v) is None: return f"Must be a number{(' in ' + f['unit']) if f.get('unit') else ''}."
    if f["type"] == "email" and not EMAIL.match(str(v).strip()): return "Not a valid email address."
    if f["type"] == "select" and f.get("options") and v not in f["options"]: return "Value not in the template's list."
    u = f.get("unit")
    if f["type"] == "number" and u in SECTION_LIMITS and num(v) > SECTION_LIMITS[u]:
        return f"Exceeds {SECTION_LIMITS[u]:g} {u} limit (FS Rules 2027 EV5.3.4)."
    if f["type"] == "number" and num(v) < 0: return "Negative value."
    return None

def review(sheet, vals, imgs, notes, skipped=()):
    """sheet: fieldmap sheet; vals: {cell|T:id: value}; imgs: {slot: count}; skipped: section ids marked 'not used'."""
    out = []; add = lambda lvl, m, cell=None: out.append({"level": lvl, "msg": m, **({"cell": cell} if cell else {})})
    name = sheet["name"]; F = {f["cell"]: f for f in sheet["fields"]}
    g = lambda c: vals.get(c) if vals.get(c) not in (None, "") else None
    active = active_sections(sheet, vals, imgs, skipped)
    for f in sheet["fields"]:
        if f["section"] not in active: continue
        v = g(f["cell"]); h = field_hint(f, v)
        if h: add("error", f'{f["label"]} ({f["cell"]}): {h}', f["cell"])
        elif v is None and f.get("required"):
            if f["type"] == "select" and f.get("default"): add("warn", f'Confirm: {f["label"]} – the template shows "{f["default"]}"; pick the value that is true for our car ({f["cell"]})', f["cell"])
            else: add("warn", f'Missing: {f["label"]} ({f["cell"]})', f["cell"])
        if v == "N/A" and not (notes or "").strip(): add("warn", f'"{f["label"]}" is N/A – add a short justification in the notes box.', f["cell"])
    for i in sheet["images"]:
        if i["section"] in active and not imgs.get(i["id"]): add("warn", f'No picture uploaded: {i["label"][:90]}')
    for t in sheet.get("tables", []):
        if t["section"] in active and sheet["name"] not in ("Version History", "Rules Clarifications") and not g("T:" + t["id"]):
            add("warn", f'No data pasted: {t["label"]}')
    if name == "Cover Sheet":
        nos = [f for f in sheet["fields"] if f["type"] == "select" and g(f["cell"]) in (None, "No")]
        if nos: add("warn", f'{len(nos)} compliance-matrix entries are still "No"/unset. Before submission every relevant row should be "Yes" (N/A explained).')
    if name == "EV5.5.10-12 Cell Fixation":
        m, h, w, n, Fc, mu = (num(g(c)) for c in ("M15", "M16", "M17", "M20", "M22", "M25"))
        typ, fix = g("M12") or F["M12"]["default"], g("M13") or F["M13"]["default"]
        if None not in (m, h, w, n) and h * w > 0:
            area = h * w * 0.8; shear = m/1000*40*9.81*n/area
            add("ok" if shear <= 0.05 else ("warn" if n <= 6 else "error"),
                f"Shear stress on most critical cell = {shear:.4f} MPa (limit 0.05 MPa, 2027 EV5.5.14)" + ("" if shear <= 0.05 else (" – EXCEPTION per template (n ≤ 6), justification required" if n <= 6 else " – FAIL")))
            if Fc is not None:
                comp = Fc/area; add("ok" if comp <= 1.0 else "error", f"Compression stress = {comp:.3f} MPa (limit 1.0 MPa, 2027 EV5.5.14)")
            if fix == "Friction based":
                if mu is None: add("error", "Friction-based fixation: friction coefficient from testing (M25) is missing – physical friction test required (2027 EV5.5.15).")
                elif Fc is not None:
                    need = m*n/1000*40*9.81; ff = Fc*mu
                    add("ok" if ff > need else "error", f"Friction force {ff:.1f} N vs. required {need:.1f} N (40 g) → {'PASS' if ff > need else 'FAIL'}")
        else: add("warn", "Cell mass, dimensions or n missing – shear/compression stresses cannot be calculated.")
        if g("M28") == "Yes": add("warn", "Compressible materials in the tab stack-up: explain how EV4.5.12 is fulfilled.")
        if typ == "Pouch" and g("M28") == "N/A": add("error", "For pouch cells the compressible-materials answer cannot be N/A.")
        if fix == "Other" and not g("M14"): add("error", "Fixation method 'Other' selected – specify it (M14).")
    if name == "EV5.5.14-17 Openings&Separation":
        nos = [f for f in sheet["fields"] if f["cell"].startswith("R") and g(f["cell"]) in (None, "No")]
        if nos: add("warn", f"{len(nos)} opening/separation requirements are still 'No' – document them with pictures and set to Yes.")
    if name == "EV5.5.13 Fasteners":
        groups = {}
        for f in sheet["fields"]: groups.setdefault(f["section"], []).append(f)
        used = [b for b in groups.values() if any(g(f["cell"]) for f in b)]
        if not used: add("warn", "No fastener group filled in.")
        for b in used:
            for f in b:
                if not g(f["cell"]): add("warn", f'Fastener group incomplete: {f["label"]} ({f["cell"]})', f["cell"])
            if g(b[0]["cell"]) == "Yes": add("info", f'{b[0]["label"].split(" – ")[0]}: structural → must follow T10 completely (2027 T10.1, T10.2).')
    if name.startswith("Datasheets"):
        for f in sheet["fields"]:
            if f["type"] == "select" and g(f["cell"]) == "NO": add("warn", f'"{f["label"]}" = NO – attach the datasheet or perform your own test.', f["cell"])
    if name == "Version History" and not g("T:A5"): add("info", "If this is a resubmission, document the changes here.")
    miss = [o for o in out if o["level"] == "warn" and o["msg"].startswith(("Missing: ", "Confirm: ", "No picture uploaded: ", "No data pasted: "))]
    if len(miss) > 6:
        out = [o for o in out if o not in miss]
        out.insert(0, {"level": "warn", "msg": f"{len(miss)} items still missing: " + "; ".join(o["msg"].split(": ", 1)[1][:60] for o in miss[:12]) + (" …" if len(miss) > 12 else ""), **({"cell": miss[0]["cell"]} if miss[0].get("cell") else {})})
    for tip in TEAM_GUIDE.get(name, []): add("tip", tip)
    if not [o for o in out if o["level"] in ("error", "warn")]: add("ok", "No issues found by the automatic check.")
    return out

def active_sections(sheet, vals, imgs, skipped=()):
    act = set()
    for sec in sheet["sections"]:
        if sec["id"] in skipped: continue
        cells = [f for f in sheet["fields"] if f["section"] == sec["id"]]
        touched = any(vals.get(f["cell"]) not in (None, "") for f in cells) or any(imgs.get(i["id"]) for i in sheet["images"] if i["section"] == sec["id"]) \
            or any(vals.get("T:" + t["id"]) for t in sheet.get("tables", []) if t["section"] == sec["id"])
        if touched or any(f.get("required") for f in cells) or (not cells and sheet["name"] not in ("Version History", "Rules Clarifications") and sec is sheet["sections"][0]):
            act.add(sec["id"])
    return act

def llm_config():
    if os.environ.get("XAI_API_KEY"): return ("https://api.x.ai/v1/chat/completions", os.environ["XAI_API_KEY"], os.environ.get("LLM_MODEL", "grok-4"))
    if os.environ.get("OPENAI_API_KEY"): return (os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1") + "/chat/completions", os.environ["OPENAI_API_KEY"], os.environ.get("LLM_MODEL", "gpt-4o-mini"))
    return None

def assistant(sheet, question, vals, findings, guidance, history=None, rules_ctx=""):
    ctx = "\n".join(TEAM_GUIDE["general"] + TEAM_GUIDE.get(sheet["name"], [])) + "\nFS Rules 2027 (authoritative; template numbering is 2026):\n" + rules_ctx
    filled = {f'{f["label"]} ({f["cell"]})': vals.get(f["cell"]) for f in sheet["fields"] if vals.get(f["cell"])}
    cfg = llm_config()
    if cfg:
        url, key, model = cfg
        sysmsg = ("You help a Formula Student electric team fill the FSG ASES (Accumulator Structural Equivalency Spreadsheet) so it is not rejected. "
                  "Answer in English, concisely and concretely: what is missing, what evidence to add, which rule (cite FS Rules 2027 numbers). "
                  f"Current tab: {sheet['name']}.\nOfficial guidance text from the template:\n{guidance[:6000]}\nTeam checklist and rules:\n{ctx}\n"
                  f"Automatic check findings: {json.dumps(findings, ensure_ascii=False)}\nEntered data: {json.dumps(filled, ensure_ascii=False)[:4000]}")
        msgs = [{"role": "system", "content": sysmsg}] + (history or [])[-8:] + [{"role": "user", "content": question}]
        req = urllib.request.Request(url, data=json.dumps({"model": model, "messages": msgs, "temperature": 0.2}).encode(),
                                     headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=60) as r: return json.load(r)["choices"][0]["message"]["content"]
        except Exception as e: return f"LLM error: {e}"
    rq = [r.replace(" ", "") for r in re.findall(r"\b(?:EV|T|A|IN)\s?\d+(?:\.\d+)+", question, re.I)]
    hits = [l for l in rules_ctx.split("\n") if any(l.upper().startswith(q.upper() + ":") for q in rq)]
    words = [w.lower() for w in re.findall(r"[A-Za-z][A-Za-z0-9.]{3,}", question)]
    lines = [l for l in (guidance + "\n" + ctx).split("\n") if l.strip() and any(w in l.lower() for w in words)][:8]
    msg = "Offline mode (no AI key configured) – answer based on the automatic checks, the template guidance and FS Rules 2027:\n"
    msg += "\n".join("• " + f["msg"] for f in findings if f["level"] in ("error", "warn"))[:2000] or "• No open findings."
    if hits: msg += "\n\nFS Rules 2027:\n" + "\n".join(hits)
    if lines: msg += "\n\nRelevant guidance:\n" + "\n".join("– " + l.strip()[:300] for l in lines)
    return msg
