"""重算验证器：formulas 引擎重算整个 workbook → 与影子期望值逐格对比。
这是『完成=有证据』的落地点。输出 PASS/FAIL 报告。"""
import sys, os, json
sys.path.insert(0, os.path.dirname(__file__))
import formulas

BASE = os.path.join(os.path.dirname(__file__), "..")

def main(path=None):
    path = path or os.path.join(BASE, "dist", "finengine.xlsx")
    snap = json.load(open(os.path.join(BASE, "layout_snapshot.json")))
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()

    cache = {}
    def read(sheet, cell):
        key = f"{sheet}|{cell}"
        if key in cache: return cache[key]
        up = f"'[{os.path.basename(path).upper()}]{sheet.upper()}'!{cell}"
        v = None
        for k, val in sol.items():
            if k.upper().endswith(f"]{sheet.upper()}'!{cell}[0,0]") or k.upper() == up:
                try: v = val.value[0, 0]
                except Exception: v = val.value
                break
        cache[key] = v
        return v

    exp = snap["expected"]
    fails, n = [], 0
    for k, want in exp.items():
        parts = k.split("|")
        sheet, key, i = parts[0], parts[1], int(parts[2])
        row = snap["rows"].get(f"{sheet}|{key}")
        if row is None: continue
        cell = f"{chr(67+i)}{row}"
        got = read(sheet, cell)
        n += 1
        if isinstance(want, str):
            if str(got).strip() != want.strip():
                fails.append((sheet, key, i, want, got))
            continue
        if got is None or isinstance(got, str):
            fails.append((sheet, key, i, want, got)); continue
        try:
            if abs(float(got) - float(want)) > max(1e-6, abs(float(want)) * 1e-9):
                fails.append((sheet, key, i, want, got))
        except (TypeError, ValueError):
            fails.append((sheet, key, i, want, got))

    print(f"重算对比：{n} 格 vs 影子值 → {'✅ 全部一致' if not fails else f'❌ {len(fails)} 格不一致'}")
    for s, k, i, w, g in fails[:25]:
        print(f"  ✗ {s}!{k}[{i}] 期望={w!r} 实得={g!r}")
    # 勾稽差异行（chk_*）专门报告
    for k, row in snap["rows"].items():
        sheet, key = k.split("|")
        if key.startswith("chk_"):
            diffs = [read(sheet, f"{chr(67+i)}{row}") for i in range(5)]
            bad = [d for d in diffs if d is None or (isinstance(d,(int,float)) and abs(d) > 1e-4)]
            tag = "✅" if not bad else "❌"
            print(f"  {tag} {sheet} · {key}: {[round(float(d),4) if d is not None else None for d in diffs]}")
    return 0 if not fails else 1

if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else None))
