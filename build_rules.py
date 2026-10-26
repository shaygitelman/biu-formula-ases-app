"""Index FS Rules PDF text by rule number -> data/rules.json"""
import re, json, sys
txt = open(sys.argv[1] if len(sys.argv) > 1 else "../rules.txt").read().split("\n")
rules, cur = {}, None
pat = re.compile(r"^([A-Z]{1,3}) (\d+(?:\.\d+){0,3})\s{2,}(\S.*)$")
for ln in txt[300:]:
    if re.match(r"^Formula Student Rules 2027|^\s*\d+ / \d+\s*$", ln.strip()): continue
    m = pat.match(ln)
    if m and not re.match(r"^\d\.\d\s{2,}", m.group(3)):
        cur = f"{m.group(1)}{m.group(2)}"; rules[cur] = m.group(3).strip(); continue
    if cur and ln.startswith("    ") and ln.strip(): rules[cur] += " " + ln.strip()
json.dump(rules, open("data/rules_2027.json", "w"), ensure_ascii=False, indent=0)
print(len(rules), {k: rules[k][:80] for k in ["EV5.3.4", "EV5.5.10", "EV5.5.13", "T10.1"] if k in rules})
