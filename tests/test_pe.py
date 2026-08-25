"""ProjectEval 全链自测：整包 workbook 构建（assumptions→capex→subs→holdco→ic→consol→projecteval）
→ formulas 重算 → ProjectEval 全部注册格逐格 vs 影子值；NPV/IRR 另用独立手算口径复核
（不复用 gen_projecteval 的影子函数）。跑法：.venv/bin/python tests/test_pe.py"""
import os, sys, math
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

from openpyxl import Workbook
import formulas
import numpy_financial as npf

import engine
from layout import Layout
from gen_assumptions import gen as gen_assum, gen_capex
from gen_subs import gen as gen_subs
from gen_holdco import gen as gen_hold
from gen_ic import gen as gen_ic
from gen_consol import gen as gen_cons
from gen_projecteval import gen as gen_pe

OUT = "/tmp/test_pe.xlsx"

# ── 1. 全链构建 ─────────────────────────────────────────
eng = engine.build_all()
wb = Workbook(); wb.remove(wb.active)
L = Layout()
gen_assum(wb, L); gen_capex(wb, L, eng)
gen_subs(wb, L, eng); gen_hold(wb, L, eng)
gen_ic(wb, L, eng); gen_cons(wb, L, eng)
gen_pe(wb, L, eng)
wb.save(OUT)
print(f"workbook 构建：{len(wb.sheetnames)} sheets → {OUT}")

# ── 2. formulas 重算 ───────────────────────────────────
sol = formulas.ExcelModel().loads(OUT).finish().calculate()

def read(sheet, cell):
    up = f"'[{os.path.basename(OUT).upper()}]{sheet.upper()}'!{cell}"
    for k, v in sol.items():
        if k.upper().endswith(f"]{sheet.upper()}'!{cell}[0,0]") or k.upper() == up:
            try: return v.value[0, 0]
            except Exception: return v.value
    return None

def close(got, want):
    return (got is not None and not isinstance(got, str)
            and abs(float(got) - float(want)) <= max(1e-6, abs(float(want)) * 1e-6))

# ── 3. 注册影子值逐格断言（仅 ProjectEval；inputs+现金流+指标+敏感性）──
fails, n = [], 0
for (sheet, key, i), want in sorted(L.expected.items()):
    if sheet != "ProjectEval":
        continue
    row = L.row_of[(sheet, key)]
    cell = f"{chr(67 + i)}{row}"
    got = read(sheet, cell)
    n += 1
    ok = (got == want) if isinstance(want, str) else close(got, want)
    if not ok:
        fails.append((key, cell, want, got))
print(f"影子值对比：{n} 格 → {'✅ 全部一致' if not fails else f'❌ {len(fails)} 格不一致'}")
for k, c, w, g in fails[:20]:
    print(f"  ✗ {k}!{c} 期望={w!r} 实得={g!r}")

# ── 4. NPV / IRR 独立手算复核（口径硬编码，与 sheet 输入一致）──
W = 0.10
def cf_of(contract, dyears, ratios, maint, mcr):
    cf = [0.0]
    for j in range(1, dyears + 1):
        cf.append(contract / dyears - contract * ratios[j - 1])
    cf += [maint * (1 - mcr)] * (5 - dyears)
    return cf

CASES = {  # (合同额, 交付年数, 各年成本比例, 运维收入, 运维成本率)
    "P1": (1800, 2, [0.55, 0.25], 180, 0.40),
    "P2": (1200, 1, [0.65], 120, 0.35),
    "P3": (900, 1, [0.60], 150, 0.30),
}
hand_fails = []
for pid, (c, d, rt, m, cr) in CASES.items():
    cf = cf_of(c, d, rt, m, cr)
    npv_hand = sum(cf[i] / (1 + W) ** i for i in range(1, 6))   # NPV 语义：Y1 起按第 1 期折现
    irr_hand = npf.irr(cf)                                      # IRR 语义：Y0-5 全列（首值 t=0）
    npv_cell = f"C{L.row_of[('ProjectEval', f'{pid}_npv')]}"
    irr_cell = f"D{L.row_of[('ProjectEval', f'{pid}_irr')]}"
    got_npv, got_irr = read("ProjectEval", npv_cell), read("ProjectEval", irr_cell)
    cf_row = L.row_of[("ProjectEval", f"{pid}_cf")]
    got_cf = [read("ProjectEval", f"{chr(67 + i)}{cf_row}") for i in range(6)]
    irr_fin = irr_hand is not None and math.isfinite(float(irr_hand))
    line_ok = close(got_npv, npv_hand) and all(close(g, w) for g, w in zip(got_cf, cf))
    irr_ok = close(got_irr, irr_hand) if irr_fin else got_irr == "—"
    tag = "✅" if (line_ok and irr_ok) else "❌"
    print(f"{tag} {pid}: NPV 手算={npv_hand:.4f} / 表={got_npv!r}；"
          f"IRR 手算={f'{float(irr_hand):.6f}' if irr_fin else '无有限解'}"
          f" / 表={got_irr!r}；净现金流={[round(x, 4) for x in cf]}")
    if not (line_ok and irr_ok):
        hand_fails.append(pid)

# ── 5. 敏感性矩阵角点独立复核（最差档 / 基准邻档 / 最好档）──
c, d, r1, r2, m, cr = 1800, 2, 0.55, 0.25, 180, 0.40
def sens_hand(mm, kk):
    fl = [c * mm / d - c * r1 * kk, c * mm / d - c * r2 * kk] + [m * (1 - cr)] * 3
    return sum(v / (1 + W) ** (i + 1) for i, v in enumerate(fl))
sens_checks = []   # (行档 m, 列档 k, 期望)
for mm, kk in [(0.80, 1.20), (0.90, 1.10), (1.10, 0.90), (1.20, 0.80)]:
    i, j = [0.80, 0.90, 1.10, 1.20].index(mm), [0.80, 0.90, 1.10, 1.20].index(kk)
    cell = f"{chr(67 + j)}{L.row_of[('ProjectEval', f'P1_sens_{i + 1}')]}"
    got = read("ProjectEval", cell)
    want = sens_hand(mm, kk)
    ok = close(got, want)
    print(f"{'✅' if ok else '❌'} 敏感性[{mm}x][{kk}x] {cell}: 手算={want:.4f} / 表={got!r}")
    if not ok:
        hand_fails.append(f"sens{mm}-{kk}")

ok_all = not fails and not hand_fails
print("\n结论：", "PASS ✅" if ok_all else "FAIL ❌")
sys.exit(0 if ok_all else 1)
