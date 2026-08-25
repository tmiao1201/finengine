"""主装配：按依赖顺序生成全部 sheet → dist/finengine.xlsx。
用法：python build.py [stage]   stage: assump | holdco | all"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from openpyxl import Workbook
import engine
from layout import Layout
from gen_assumptions import gen as gen_assum, gen_capex

OUT = os.path.join(os.path.dirname(__file__), "..", "dist", "finengine.xlsx")

def main(stage="all"):
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    eng = engine.build_all()
    wb = Workbook(); wb.remove(wb.active)
    L = Layout()
    gen_assum(wb, L); gen_capex(wb, L, eng)
    if stage in ("holdco", "all"):
        from gen_subs import gen as gen_subs
        from gen_holdco import gen as gen_hold
        gen_subs(wb, L, eng)   # 子公司先（HoldCo T1 公式引用其 rev_ext）
        gen_hold(wb, L, eng)
    if stage == "all":
        for mod, fn in [("gen_subs", "gen"), ("gen_ic", "gen"),
                        ("gen_consol", "gen"), ("gen_analysis", "gen"), ("gen_checks", "gen")]:
            try:
                m = __import__(mod); getattr(m, fn)(wb, L, eng)
            except ImportError:
                print(f"  (skip {mod} — 未实现)")
    wb.save(OUT)
    # 坐标与期望值快照，供 recompute 使用
    import json
    snap = {"rows": {f"{s}|{k}": r for (s, k), r in L.row_of.items()},
            "expected": {f"{s}|{k}|{i}": v for (s, k, i), v in L.expected.items()}}
    with open(os.path.join(os.path.dirname(__file__), "..", "layout_snapshot.json"), "w") as f:
        json.dump(snap, f, ensure_ascii=False)
    print(f"✅ build[{stage}] → {os.path.abspath(OUT)}（{len(wb.sheetnames)} sheets）")

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "all")
