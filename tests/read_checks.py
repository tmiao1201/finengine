"""读 Checks 面板各断言的实时计算状态。"""
import sys, json, re, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import formulas

BASE = os.path.join(os.path.dirname(__file__), "..")
sol = formulas.ExcelModel().loads(os.path.join(BASE, "dist/finengine.xlsx")).finish().calculate()
snap = json.load(open(os.path.join(BASE, "layout_snapshot.json")))
rows = {v: k.split("|")[1] for k, v in snap["rows"].items() if k.startswith("Checks|")}
ok = fail = other = 0
pat = re.compile(r"\[finengine\.xlsx\]CHECKS'!I(\d+)$", re.IGNORECASE)
for k, v in sol.items():
    m = pat.search(k)
    if not m:
        continue
    try:
        s = str(v.value[0, 0])
    except Exception:
        s = str(v.value)
    r = int(m.group(1))
    if "PASS" in s:
        ok += 1
    elif "FAIL" in s:
        fail += 1
        print("FAIL", rows.get(r, r), s[:30])
    else:
        other += 1
print(f"Checks: {ok} PASS / {fail} FAIL / {other} 其他")
