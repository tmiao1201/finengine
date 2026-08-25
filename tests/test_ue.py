"""UE 自测：构建全链 workbook（assumptions→capex→subs→holdco→ic→consol→UE）存 /tmp/test_ue.xlsx，
formulas 引擎重算，断言所有 L.reg expected 格与重算值一致（容差：相对 1e-6）。
用法：.venv/bin/python tests/test_ue.py"""
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "src")
sys.path.insert(0, SRC)

from openpyxl import Workbook
import formulas
import engine
from layout import Layout
from gen_assumptions import gen as gen_assumptions, gen_capex
from gen_subs import gen as gen_subs
from gen_holdco import gen as gen_holdco
from gen_ic import gen as gen_ic
from gen_consol import gen as gen_consol
from gen_ue import gen as gen_ue

OUT = "/tmp/test_ue.xlsx"

def build_workbook():
    eng = engine.build_all()
    wb = Workbook(); wb.remove(wb.active)
    L = Layout()
    gen_assumptions(wb, L); gen_capex(wb, L, eng)
    gen_subs(wb, L, eng); gen_holdco(wb, L, eng)
    gen_ic(wb, L, eng); gen_consol(wb, L, eng)
    gen_ue(wb, L, eng)
    wb.save(OUT)
    return L, eng

def main():
    L, eng = build_workbook()
    print(f"workbook → {OUT}（重算中，约需数十秒）")
    sol = formulas.ExcelModel().loads(OUT).finish().calculate()

    base = os.path.basename(OUT).upper()
    cache = {}
    def read(sheet, cell):
        key = f"{sheet}|{cell}"
        if key in cache: return cache[key]
        up = f"'[{base}]{sheet.upper()}'!{cell}"
        v = None
        for k, val in sol.items():
            if k.upper().endswith(f"]{sheet.upper()}'!{cell}[0,0]") or k.upper() == up:
                try: v = val.value[0, 0]
                except Exception: v = val.value
                break
        cache[key] = v
        return v

    fails, n_ue, n_all = [], 0, 0
    for (sheet, key, i), want in sorted(L.expected.items()):
        row = L.row_of[(sheet, key)]
        got = read(sheet, f"{chr(67+i)}{row}")
        n_all += 1
        if sheet == "UE": n_ue += 1
        ok = False
        if got is not None and not isinstance(got, str):
            try:
                # 相对 1e-6；近零影子值（D() 舍入尘埃）按绝对 1e-6，与 recompute.py 口径一致
                ok = abs(float(got) - float(want)) <= max(1e-6, abs(float(want)) * 1e-6)
            except (TypeError, ValueError):
                ok = False
        if not ok:
            fails.append((sheet, key, i, want, got))

    print(f"重算对比：全簿 {n_all} 格（其中 UE {n_ue} 格）→ "
          f"{'全部一致 ✅' if not fails else f'{len(fails)} 格不一致 ❌'}")
    for s, k, i, w, g in fails[:25]:
        print(f"  ✗ {s}!{k}[{i}] 期望={w!r} 实得={g!r}")

    # 证据：UE 关键影子值摘要
    gpm = [L.val("UE", "t_gpm", i) for i in range(5)]
    print("\nUE 关键值摘要（影子值）：")
    print("  单token毛利率 :", [f"{L.val('UE','m_gm',i):.1%}" for i in range(5)])
    print("  剪刀差(成本−价):", [f"{L.val('UE','sc_d',i):+.1%}" for i in range(5)])
    print("  LTV/CAC (倍)  :", [f"{L.val('UE','c_ltvcac',i):.2f}x" for i in range(5)])
    print("  CAC回收期(月) :", [f"{L.val('UE','c_pb',i):.1f}" for i in range(5)])
    print("  C端月毛利(元) :", [f"{gpm[i]:.2f}" for i in range(5)])
    print("  C端回本(月)   :", [f"{L.val('UE','t_pb',i):.1f}" for i in range(5)])

    if fails:
        print("\nFAIL")
        return 1
    print("\nPASS")
    return 0

if __name__ == "__main__":
    sys.exit(main())
