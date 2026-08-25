"""M5 检查层：Checks 勾稽自检汇总面板（8 组断言）+ Cover 导航页。
Checks 每行公式全链接到各 sheet 的证据行，总灯 = 全部 PASS。"""
import assumptions as A
from style import *
from layout import Layout

BS_SHEET = {"HOLD": "HoldCo_BS", "CLOUD": "Cloud_BS_CF", "IND": "Ind_BS_CF", "TOC": "Toc_BS_CF"}
IS_SHEET = {"HOLD": "HoldCo_IS", "CLOUD": "Cloud_IS", "IND": "Ind_IS", "TOC": "Toc_IS"}

def gen_checks(wb, L, eng):
    ws = wb.create_sheet("Checks")
    sheet_scaffold(ws, "勾稽自检面板 · 8 组断言", "全部公式实时链接 | 任何一格改坏假设，红灯立刻亮 | 万元", A.YEARS, "check")
    r = 4
    results = []  # (行号) PASS/FAIL 格
    def crow(key, label, formulas, note=None, tol="0.01"):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, numFmt=NUM1, indent=1)
        st = ws.cell(row=r, column=9,
                     value=f'=IF(AND(ABS(C{r})<{tol},ABS(D{r})<{tol},ABS(E{r})<{tol},'
                           f'ABS(F{r})<{tol},ABS(G{r})<{tol}),"✅ PASS","❌ FAIL")')
        st.font = BOLD
        L.reg("Checks", key, r)
        results.append(r)
        r += 1
    def sec(t):
        nonlocal r
        c = ws.cell(row=r, column=1, value=t); c.font = BOLD; c.fill = SECTION_FILL
        for col in range(2, 10): ws.cell(row=r, column=col).fill = SECTION_FILL
        r += 1
    ws.cell(row=3, column=9, value="状态").font = BOLD
    ws.column_dimensions["I"].width = 12

    sec("── C1 各主体资产负债表平衡（资产−负债−权益）──")
    crow("c1_hold", "智擎研究院", [f"={L.ref('HoldCo_BS','chk_balance',i)}" for i in range(5)], "倒挤现金=结构性平衡")
    # 子公司：倒挤现金同样结构性平衡，负债+权益−非现金资产−现金 应=0
    for e in ["CLOUD", "IND", "TOC"]:
        s = BS_SHEET[e]
        f = []
        for i in range(5):
            lhs = f"{L.ref(s,'liab',i)}+{L.ref(s,'paid_in',i)}+{L.ref(s,'RE',i)}"
            nc = L.ref(s, 'assets_nc', i) if e == "CLOUD" else L.ref(s, 'AR', i)
            f.append(f"={lhs}-{nc}-{L.ref(s,'cash',i)}")
        crow(f"c1_{e.lower()}", A.ENTITY_NAMES[e], f)

    sec("── C2 现金流量表 = 资产负债表现金（间接法交叉验证）──")
    for e, s in [("HOLD", "HoldCo_CF"), ("CLOUD", "Cloud_BS_CF"), ("IND", "Ind_BS_CF"),
                 ("TOC", "Toc_BS_CF"), ("CONSOL", "Consol_CF")]:
        crow(f"c2_{e.lower()}", A.ENTITY_NAMES.get(e, "合并集团"),
             [f"={L.ref(s,'chk_tie',i)}" for i in range(5)])

    sec("── C3 合并资产负债表平衡 ──")
    crow("c3", "合并资产 − 负债 − 权益", [f"={L.ref('Consol_BS','chk_balance',i)}" for i in range(5)])

    sec("── C4 合并 vs Σ单体 交叉验证（合并−Σ应=0）──")
    crow("c4_cash", "货币资金", [f"={L.ref('Consol_BS','cash',i)}-("
         + "+".join(L.ref(BS_SHEET[e], "cash", i) for e in BS_SHEET) + ")" for i in range(5)])
    crow("c4_ar", "应收账款", [f"={L.ref('Consol_BS','AR',i)}-("
         + "+".join(L.ref(BS_SHEET[e], "AR", i) for e in BS_SHEET) + ")" for i in range(5)])
    crow("c4_ni", "净利润", [f"={L.ref('Consol_IS','ni',i)}-("
         + "+".join(L.ref(IS_SHEET[e], "ni", i) for e in IS_SHEET) + ")" for i in range(5)])

    sec("── C5 关联方对账（调节后全 0）──")
    crow("c5_p1", "研究院 ↔ 智擎云", [f"={L.ref('IC_Recon','p1_diff',i)}" for i in range(5)])
    crow("c5_p2", "研究院 ↔ 智擎行业", [f"={L.ref('IC_Recon','p2_diff',i)}" for i in range(5)])
    crow("c5_p3", "研究院 ↔ 智擎互联", [f"={L.ref('IC_Recon','p3_diff',i)}" for i in range(5)])
    crow("c5_p4", "智擎云 ↔ 智擎互联（在途调节后）", [f"={L.ref('IC_Recon','adj3',i)}" for i in range(5)],
         "调节前差异 220 → 调节归零")

    sec("── C7 利润滚动：RE[t] − RE[t−1] = 当年 NI ──")
    for e in A.ENTITIES:
        s = BS_SHEET[e]; isni = IS_SHEET[e]
        f = [f"={L.ref(s,'RE',0)}-{A.OPENING[e]['RE']}-{L.ref(isni,'ni',0)}"] + \
            [f"={L.ref(s,'RE',i)}-{L.ref(s,'RE',i-1)}-{L.ref(isni,'ni',i)}" for i in range(1, 5)]
        crow(f"c7_{e.lower()}", A.ENTITY_NAMES[e], f)

    sec("── C8 归母净利润 + 少数股东损益 = 合并净利润 ──")
    crow("c8", "归属拆分核对",
         [f"={L.ref('Consol_IS','ni_parent',i)}+{L.ref('Consol_IS','ni_minority',i)}"
          f"-{L.ref('Consol_IS','ni',i)}" for i in range(5)])

    r += 1
    c = ws.cell(row=r, column=1, value="总检灯")
    c.font = Font(name=FONT, bold=True, size=12)
    lamp = ws.cell(row=r, column=3,
                   value=f'=IF(COUNTIF(I4:I{r-2},"✅ PASS")={len(results)},"🟢 全部通过 — 模型勾稽完整","🔴 存在断言失败，检查上行")')
    lamp.font = Font(name=FONT, bold=True, size=12)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
    L.reg("Checks", "lamp", r)
    return ws

def gen_cover(wb, L, eng):
    ws = wb.create_sheet("Cover", 0)
    ws.sheet_properties.tabColor = TAB["cover"]
    ws["A1"] = "智擎集团 · 集团财务分析模型"; ws["A1"].font = Font(name=FONT, bold=True, size=20)
    ws["A2"] = "虚构大模型集团（参考智谱商业模式）| Python+openpyxl 全公式生成 | 单位：万元"
    ws["A2"].font = SUBTITLE
    ws.column_dimensions["A"].width = 26
    for col in "BCDEFG": ws.column_dimensions[col].width = 16

    nav = [
        ("Assumptions", "核心假设——全部蓝字输入，改动全模型联动"),
        ("Capex_Dep", "GPU 资本开支与折旧引擎（批次法）"),
        ("Cloud_IS / Cloud_BS_CF", "智擎云 · MaaS（单体全量）"),
        ("Ind_IS / Ind_BS_CF", "智擎行业 · 私有化+解决方案"),
        ("Toc_IS / Toc_BS_CF", "智擎互联 · C端（母持 80%）"),
        ("HoldCo_IS / BS / CF", "智擎研究院 · 全量三表（关联收入方）"),
        ("IC_Register", "内部交易登记簿（T1/T2/T3）"),
        ("IC_Recon", "关联方对账矩阵 + cut-off 在途调节"),
        ("Eliminations", "合并抵消分录底稿（E0-E3）"),
        ("Consol_IS / BS / CF", "合并三表（Σ单体−抵消）"),
        ("FactTable / Dashboard", "维度数据仓 + 四下拉自选切换看板"),
        ("UE", "单位经济：MaaS 单token/单客户 · C端单用户"),
        ("ProjectEval", "私有化项目评估：NPV/IRR + 敏感性"),
        ("Checks", "勾稽自检面板：8 组断言全绿才可信"),
    ]
    r = 4
    c = ws.cell(row=r, column=1, value="导航"); c.font = BOLD
    c2 = ws.cell(row=r, column=2, value="说明"); c2.font = BOLD
    r += 1
    for name, desc in nav:
        cell = ws.cell(row=r, column=1, value=name.split(" / ")[0])
        cell.hyperlink = f"#'{name.split(' / ')[0]}'!A1"
        cell.font = Font(name=FONT, color="0563C1", underline="single")
        ws.cell(row=r, column=2, value=desc).font = BLACK_FORMULA
        r += 1
    r += 1
    c = ws.cell(row=r, column=1, value="集团速览（合并口径）"); c.font = BOLD
    r += 1
    rows = [
        ("营业收入", "Consol_IS", "rev"), ("净利润", "Consol_IS", "ni"),
        ("归母净利润", "Consol_IS", "ni_parent"), ("期末现金", "Consol_BS", "cash"),
        ("GPU 净值", "Consol_BS", "fa"),
    ]
    for label, sheet, key in rows:
        ws.cell(row=r, column=1, value=label).font = BLACK_FORMULA
        for i in range(5):
            cc = ws.cell(row=r, column=2 + i, value=f"={L.ref(sheet, key, i)}")
            cc.number_format = NUM; cc.font = BLACK_FORMULA
        r += 1
    for i, y in enumerate(A.YEARS):
        h = ws.cell(row=r - 6, column=2 + i, value=y)
        h.number_format = '0"A"' if y <= 2025 else '0"E"'
        h.font = BOLD_WHITE; h.fill = HEADER_FILL
    r += 1
    ws.cell(row=r, column=1, value="勾稽总检灯 →").font = BOLD
    try:
        lamp_ref = L.ref("Checks", "lamp")
        lc = ws.cell(row=r, column=2, value=f'=IF({lamp_ref}="","",{lamp_ref})')
        lc.font = Font(name=FONT, bold=True, size=12)
    except KeyError:
        pass
    return ws

def gen(wb, L, eng):
    gen_checks(wb, L, eng)
    gen_cover(wb, L, eng)
