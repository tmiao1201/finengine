"""roundtrip + Checks 面板验证（合并）：数据与逻辑分离 + 勾稽总灯。"""
import sys, os, re
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import engine, meta
import formulas

BASE = os.path.join(os.path.dirname(__file__), "..")
PATH = os.path.join(BASE, "dist", "coffee-model.xlsx")

# 1) roundtrip
base = engine.build()
rows = meta.load_meta_from_xlsx(PATH)
meta.bind(rows)
again = engine.build()
bad = sum(1 for stmt in ["IS", "BS", "CF"] for k in base[stmt]
          for i in range(5)
          if abs(float(base[stmt][k][i]) - float(again[stmt][k][i])) > 1e-6)
print(f"roundtrip：从 xlsx 读回 {len(rows)} 行 Meta → 引擎重算 → {'✅ 全一致' if bad == 0 else f'❌ {bad} 格'}")

# 2) Checks 面板
sol = formulas.ExcelModel().loads(PATH).finish().calculate()
ok = fail = 0
pat = re.compile(r"\[coffee-model\.xlsx\]CHECKS'!I(\d+)$", re.IGNORECASE)
for k, v in sol.items():
    mm = pat.search(k)
    if not mm:
        continue
    try:
        s = str(v.value[0, 0])
    except Exception:
        s = str(v.value)
    if "PASS" in s:
        ok += 1
    elif "FAIL" in s:
        fail += 1
        print("FAIL:", mm.group(1))
print(f"Checks 面板：{ok} PASS / {fail} FAIL")
sys.exit(0 if bad == 0 and fail == 0 else 1)
