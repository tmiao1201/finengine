"""Dashboard 自测：构建全 workbook → formulas 重算 → 断言。
(a) Dashboard 反查差异行 5 列全 0（FactTable SUMIFS 反查 = Consol_IS!rev）
(b) 默认选择（四个维度全 "*"）下收入格重算值 = 影子合并收入（5 年合计）
跑法：.venv/bin/python tests/test_dashboard.py → 输出 PASS/FAIL 明细
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from openpyxl import Workbook
import formulas
import engine
from layout import Layout
from gen_assumptions import gen as gen_assum, gen_capex
from gen_subs import gen as gen_subs
from gen_holdco import gen as gen_holdco
from gen_ic import gen as gen_ic
from gen_consol import gen as gen_consol
from gen_dashboard import gen as gen_dashboard

OUT = "/tmp/test_dash.xlsx"


def build():
    eng = engine.build_all()
    wb = Workbook()
    wb.remove(wb.active)
    L = Layout()
    gen_assum(wb, L)
    gen_capex(wb, L, eng)
    gen_subs(wb, L, eng)      # 子公司先（HoldCo 引用其 rev_ext）
    gen_holdco(wb, L, eng)
    gen_ic(wb, L, eng)
    gen_consol(wb, L, eng)
    gen_dashboard(wb, L, eng)
    wb.save(OUT)
    return L, eng


def main():
    L, eng = build()
    print(f"workbook → {OUT}（含 Dashboard/FactTable）")
    sol = formulas.ExcelModel().loads(OUT).finish().calculate()

    def read(sheet, cell):
        for k, v in sol.items():
            if k.upper().endswith(f"]{sheet.upper()}'!{cell}[0,0]") or \
               k.upper().endswith(f"]{sheet.upper()}'!{cell}"):
                try:
                    return v.value[0, 0]
                except Exception:
                    return v.value
        return None

    fails = []

    # (a) 反查差异行：5 列全 0
    chk = L.row_of[("Dashboard", "chk_dash_rev")]
    print(f"\n(a) 反查差异行 Dashboard!C{chk}:G{chk}（SUMIFS 合并全量收入 − Consol_IS!rev）")
    for i in range(5):
        cell = f"{chr(67 + i)}{chk}"
        v = read("Dashboard", cell)
        ok = v is not None and not isinstance(v, str) and abs(float(v)) < 1e-4
        print(f"    {cell}: {v!r:>22}  {'✅' if ok else '❌'}  ({A_LABELS[i]})")
        if not ok:
            fails.append(f"(a) {cell}={v!r}")

    # (b) 默认选择（全 "*"）收入格 = 影子合并收入（期间=全部 → 5 年合计）
    rev_row = L.row_of[("Dashboard", "dash_rev")]
    got = read("Dashboard", f"C{rev_row}")
    want = sum(eng["C"]["IS"]["rev"])
    ok = got is not None and not isinstance(got, str) and \
        abs(float(got) - want) <= max(1e-4, abs(want) * 1e-9)
    print(f"\n(b) 默认收入格 Dashboard!C{rev_row}（主体/业务线/地区/期间 全部=“*”）")
    print(f"    重算值 = {got!r}")
    print(f"    影子合并收入 5 年合计 = {want:,.2f}（逐年 {[round(x) for x in eng['C']['IS']['rev']]}）")
    print(f"    {'✅' if ok else '❌'}")
    if not ok:
        fails.append(f"(b) 收入格 {got!r} != {want!r}")

    # 附：逐年反查明细（诊断用，不计入 FAIL）
    ft_row = L.row_of[("Dashboard", "ft_rev_all")]
    per_year = [read("Dashboard", f"{chr(67 + i)}{ft_row}") for i in range(5)]
    print(f"\n附 · 反查 SUMIFS 逐年收入: {[round(float(x)) if x is not None else None for x in per_year]}")
    print(f"附 · 影子 Consol_IS rev   : {[round(x) for x in eng['C']['IS']['rev']]}")

    print("\n" + "=" * 50)
    if fails:
        print("FAIL ❌")
        for f in fails:
            print("  ✗", f)
        return 1
    print("PASS ✅  (a) 反查 5 列全 0；(b) 默认收入格=影子合并收入")
    return 0


A_LABELS = ["2023A", "2024A", "2025A", "2026E", "2027E"]

if __name__ == "__main__":
    sys.exit(main())
