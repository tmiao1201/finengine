"""一键重建：Meta 数据层 → 全部 sheet → dist/coffee-model.xlsx + 坐标快照。"""
import sys, os, json
sys.path.insert(0, os.path.dirname(__file__))
from openpyxl import Workbook
import engine
from layout import Layout
from gen_meta import gen as gen_meta
from gen_sheets import gen

OUT = os.path.join(os.path.dirname(__file__), "..", "dist", "coffee-model.xlsx")

def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    R = engine.build()
    wb = Workbook(); wb.remove(wb.active)
    L = Layout()
    gen_meta(wb, L, R)
    gen(wb, L, R)
    wb.save(OUT)
    snap = {"rows": {f"{s}|{k}": r for (s, k), r in L.row_of.items()},
            "expected": {f"{s}|{k}|{i}": v for (s, k, i), v in L.expected.items()}}
    with open(os.path.join(os.path.dirname(__file__), "..", "layout_snapshot.json"), "w") as f:
        json.dump(snap, f, ensure_ascii=False)
    print(f"✅ build → {os.path.abspath(OUT)}（{len(wb.sheetnames)} sheets）")

if __name__ == "__main__":
    main()
