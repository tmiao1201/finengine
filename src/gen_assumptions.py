"""Assumptions + Capex_Dep 生成器：全部假设落表并注册坐标，附带批次折旧引擎。"""
import assumptions as A
from style import *
from layout import Layout

def _row(ws, L, r, key, label, vals, note=None, numFmt=NUM):
    line(ws, r, label, values=vals, note=note, numFmt=numFmt)
    L.reg("Assumptions", key, r, expected=vals)
    return r + 1

def _const(ws, L, r, key, label, val, note=None, numFmt=None):
    line(ws, r, label, values=[val] * 5, note=note,
         numFmt=numFmt or ('0.0%' if isinstance(val, float) and val < 1 else '#,##0'))
    L.reg("Assumptions", key, r, expected=[val] * 5)
    return r + 1

def gen(wb, L):
    ws = wb.create_sheet("Assumptions")
    sheet_scaffold(ws, "智擎集团 · 核心假设 Assumptions",
                   "蓝字=可调输入 | 单位：万元（另注除外）| 改动全模型联动", A.YEARS, "assum")
    r = 4
    def sec(t):
        nonlocal r
        c = ws.cell(row=r, column=1, value=t); c.font = BOLD; c.fill = SECTION_FILL
        for col in range(2, 8): ws.cell(row=r, column=col).fill = SECTION_FILL
        r += 1

    sec("── 集团共用 ──")
    r = _const(ws, L, r, "TAX_RATE", "所得税率（高新15%）", A.TAX_RATE, numFmt=PCT)
    r = _const(ws, L, r, "DSO", "应收账期（天）", A.DSO)
    r = _const(ws, L, r, "DPO", "应付账期（天）", A.DPO)
    r = _const(ws, L, r, "IC_DAYS", "内部往来账期（天）", A.IC_DAYS)
    r = _const(ws, L, r, "WACC", "WACC（项目评估）", A.WACC, numFmt=PCT)

    sec("── 智擎云 · MaaS ──")
    r = _row(ws, L, r, "CLOUD_TOKENS_B", "日均 token 消耗（B/day）", A.CLOUD_TOKENS_B, "智谱 bigmodel 模式", '#,##0')
    r = _row(ws, L, r, "CLOUD_PRICE", "综合单价（元/M tokens）", A.CLOUD_PRICE, "混合计价", '0.00')
    r = _row(ws, L, r, "CLOUD_UCOST", "推理单位成本（元/M tokens）", A.CLOUD_UCOST, "自持+外租混合", '0.00')
    r = _row(ws, L, r, "CLOUD_BW", "带宽机房费用", A.CLOUD_BW)
    for k, lbl in [("rd", "研发编制（人）"), ("sales", "销售编制（人）"), ("gna", "管理编制（人）")]:
        r = _row(ws, L, r, f"CLOUD_{k.upper()}_HC", lbl, A.CLOUD_HEADCOUNT[k], numFmt='#,##0')
    r = _row(ws, L, r, "CLOUD_CUSTOMERS", "付费客户数（家）", A.CLOUD_CUSTOMERS, "UE 用", '#,##0')
    r = _const(ws, L, r, "CLOUD_CHURN_M", "月流失率", A.CLOUD_CHURN_M, "UE 用", PCT)

    sec("── 智擎行业 · 私有化+解决方案 ──")
    r = _row(ws, L, r, "IND_PROJECTS", "私有化项目数（个）", A.IND_PROJECTS, numFmt='#,##0')
    r = _row(ws, L, r, "IND_CONTRACT", "平均合同额", A.IND_CONTRACT)
    r = _row(ws, L, r, "IND_SUBSCRIBERS", "解决方案签约客户（累计）", A.IND_SUBSCRIBERS, "订阅/运维", '#,##0')
    r = _row(ws, L, r, "IND_SUB_FEE", "解决方案年费", A.IND_SUB_FEE)
    r = _const(ws, L, r, "IND_COGS_PVT", "私有化成本率（硬件+人力）", A.IND_COGS_PVT, numFmt=PCT)
    r = _const(ws, L, r, "IND_COGS_SOL", "解决方案成本率", A.IND_COGS_SOL, numFmt=PCT)
    for k, lbl in [("rd", "研发编制（人）"), ("sales", "销售编制（人）"), ("gna", "管理编制（人）")]:
        r = _row(ws, L, r, f"IND_{k.upper()}_HC", lbl, A.IND_HEADCOUNT[k], numFmt='#,##0')

    sec("── 智擎互联 · C端 ──")
    r = _row(ws, L, r, "TOC_MAU_W", "MAU（万人）", A.TOC_MAU_W, numFmt='#,##0')
    r = _row(ws, L, r, "TOC_PAYRATE", "付费转化率", A.TOC_PAYRATE, numFmt='0.0%')
    r = _row(ws, L, r, "TOC_ARPPU", "ARPPU（元/付费用户/月）", A.TOC_ARPPU, numFmt='0.0')
    r = _row(ws, L, r, "TOC_ADS", "广告收入", A.TOC_ADS)
    r = _row(ws, L, r, "TOC_UCOST", "单位推理+带宽成本（元/MAU/月）", A.TOC_UCOST, numFmt='0.00')
    r = _row(ws, L, r, "TOC_CAC", "获客成本（元/新增用户）", A.TOC_CAC, numFmt='0.0')
    r = _const(ws, L, r, "TOC_NEW_U_MULTIPLIER", "新增用户=MAU增量×", A.TOC_NEW_U_MULTIPLIER, "补流失", '0.0"x"')
    for k, lbl in [("rd", "研发编制（人）"), ("sales", "固定销售编制（人）"), ("gna", "管理编制（人）")]:
        r = _row(ws, L, r, f"TOC_{k.upper()}_HC", lbl, A.TOC_HEADCOUNT[k], numFmt='#,##0')

    sec("── 智擎研究院 ──")
    r = _row(ws, L, r, "HOLD_HC", "研发编制（人）", A.HOLD_HEADCOUNT, "基座模型研发", '#,##0')
    r = _row(ws, L, r, "HOLD_SALARY", "人均全成本（万/年）", A.HOLD_SALARY, numFmt='0.0')
    r = _row(ws, L, r, "HOLD_OTHER_RD", "其他研发费用（数据/外包）", A.HOLD_OTHER_RD)
    r = _row(ws, L, r, "HOLD_GNA_HC", "管理编制（人）", A.HOLD_GNA_HC, numFmt='#,##0')
    r = _const(ws, L, r, "HOLD_GNA_SALARY", "管理人均全成本（万/年）", A.HOLD_GNA_SALARY, numFmt='0.0')

    sec("── 关联交易定价 ──")
    r = _const(ws, L, r, "T1_RATE_CLOUD", "T1 license 费率 · 智擎云", A.T1_RATE["CLOUD"], "收入×费率", PCT)
    r = _const(ws, L, r, "T1_RATE_IND", "T1 license 费率 · 智擎行业", A.T1_RATE["IND"], "收入×费率", PCT)
    r = _const(ws, L, r, "T1_RATE_TOC", "T1 license 费率 · 智擎互联", A.T1_RATE["TOC"], "收入×费率", PCT)
    r = _row(ws, L, r, "T2_FEE", "T2 定制研发服务费（行业→研究院）", A.T2_FEE)
    r = _const(ws, L, r, "T3_SHARE", "T3 互联算力结算占其推理成本", A.T3_SHARE, numFmt=PCT)
    r = _const(ws, L, r, "T3_GM", "T3 云内部毛利率", A.T3_GM, "内部利润随互联对外收入实现", PCT)
    r = _const(ws, L, r, "CUTOFF_2027", "cut-off 在途（互联已确认/云未开票）", A.CUTOFF_2027, "仅2027年末演示")

    sec("── 股权融资计划（当年新增注资）──")
    for e, lbl in [("HOLD", "研究院（老股东直接增资）"), ("CLOUD", "智擎云"), ("IND", "智擎行业"),
                   ("TOC", "智擎互联（含战投20%）")]:
        r = _row(ws, L, r, f"INJECT_{e}", lbl, A.EQUITY_INJECT[e])

    sec("── 收入地区拆分 ──")
    for rg in A.REGIONS:
        r = _row(ws, L, r, f"MIX_{rg}", f"{rg} 占比", A.REGION_MIX[rg], numFmt=PCT)

    sec("── 期初数（2023年初）──")
    for e in A.ENTITIES:
        r = _const(ws, L, r, f"OPEN_RE_{e}", f"{A.ENTITY_NAMES[e]} 期初未分配利润", A.OPENING[e]["RE"])
        r = _const(ws, L, r, f"OPEN_PI_{e}", f"{A.ENTITY_NAMES[e]} 期初实收资本", A.OPENING[e]["paid_in"])
    r = _const(ws, L, r, "OPEN_FA_HOLD", "研究院期初 GPU 原值", A.OPENING_FA["HOLD"], "并入2023批次")
    r = _const(ws, L, r, "OPEN_FA_CLOUD", "智擎云期初 GPU 原值", A.OPENING_FA["CLOUD"], "并入2023批次")
    r = _const(ws, L, r, "OPEN_LT_INV", "研究院期初长投合计", A.OPENING_LT_INV)
    for e, dr in [("CLOUD", A.DEFERRED_RATE["CLOUD"]), ("IND", A.DEFERRED_RATE["IND"]),
                  ("TOC", A.DEFERRED_RATE["TOC"])]:
        r = _const(ws, L, r, f"DEF_RATE_{e}", f"{A.ENTITY_NAMES[e]} 递延收入率", dr, "期末递延/当年收入", PCT)
    return ws


def gen_capex(wb, L, eng):
    """批次折旧引擎：批次明细 → 当年折旧 → 净值。公式全链接。"""
    ws = wb.create_sheet("Capex_Dep")
    sheet_scaffold(ws, "GPU 资本开支与折旧引擎", "批次法：购置当年起 5 年直线 | 大模型公司的资产负债表灵魂",
                   A.YEARS, "engine")
    dep, net = eng["dep"], eng["net"]
    r = 4
    def sec(t):
        nonlocal r
        c = ws.cell(row=r, column=1, value=t); c.font = BOLD; c.fill = SECTION_FILL
        for col in range(2, 9): ws.cell(row=r, column=col).fill = SECTION_FILL
        r += 1

    sec("── 采购计划（输入）──")
    ws.column_dimensions["H"].width = 12
    line(ws, r, "训练 GPU 采购（研究院）", values=A.CAPEX["HOLD_TRAIN"], note="含期初存量并入2023批")
    L.reg("Capex_Dep", "CAPEX_TRAIN", r, expected=A.CAPEX["HOLD_TRAIN"]); r += 1
    line(ws, r, "推理 GPU 采购（智擎云）", values=A.CAPEX["CLOUD_INFER"])
    L.reg("Capex_Dep", "CAPEX_INFER", r, expected=A.CAPEX["CLOUD_INFER"]); r += 2

    sec("── 批次折旧明细（H列=批次原值，B列=购置年）──")
    batch_rows = {}
    for name, capex_key, opening_key, opening in [
            ("训练", "CAPEX_TRAIN", "OPEN_FA_HOLD", A.OPENING_FA["HOLD"]),
            ("推理", "CAPEX_INFER", "OPEN_FA_CLOUD", A.OPENING_FA["CLOUD"])]:
        rows = []
        for i, y in enumerate(A.YEARS):
            orig = opening + A.CAPEX[{"训练": "HOLD_TRAIN", "推理": "CLOUD_INFER"}[name]][0] if i == 0 \
                else A.CAPEX[{"训练": "HOLD_TRAIN", "推理": "CLOUD_INFER"}[name]][i]
            ws.cell(row=r, column=1, value=f"{name}卡 {y}批").font = BLACK_FORMULA
            b = ws.cell(row=r, column=2, value=y); b.font = BLUE_INPUT; b.number_format = '0"年"'
            h = ws.cell(row=r, column=8, value=orig); h.font = BLUE_INPUT; h.number_format = NUM
            # 折旧公式：购置当年起 5 年
            for ci in range(5):
                col = get_column_letter(3 + ci)
                cell = ws.cell(row=r, column=3 + ci)
                cell.value = f'=IF(AND({col}$3>=$B{r},{col}$3<$B{r}+5),$H{r}/5,0)'
                cell.number_format = NUM
            rows.append(r); r += 1
        batch_rows[name] = rows
        r += 1

    sec("── 汇总 ──")
    tr0, tr1 = batch_rows["训练"][0], batch_rows["训练"][-1]
    ir0, ir1 = batch_rows["推理"][0], batch_rows["推理"][-1]
    f = [f'=SUM({get_column_letter(3+i)}{tr0}:{get_column_letter(3+i)}{tr1})' for i in range(5)]
    line(ws, r, "训练 GPU 当年折旧", formulas=f, bold=True, note="→ 研究院研发费用")
    L.reg("Capex_Dep", "DEP_TRAIN", r, expected=dep["HOLD_TRAIN"]); dep_row = r; r += 1
    # 净值 = 累计原值 − 累计折旧
    f = [f'=SUMIFS($H${tr0}:$H${tr1},$B${tr0}:$B${tr1},"<="&{get_column_letter(3+i)}$3)'
         f'-SUM($C${dep_row}:{get_column_letter(3+i)}{dep_row})' for i in range(5)]
    line(ws, r, "训练 GPU 净值", formulas=f, bold=True, note="→ 研究院 BS")
    L.reg("Capex_Dep", "NET_TRAIN", r, expected=net["HOLD_TRAIN"]); r += 1

    f = [f'=SUM({get_column_letter(3+i)}{ir0}:{get_column_letter(3+i)}{ir1})' for i in range(5)]
    line(ws, r, "推理 GPU 当年折旧", formulas=f, bold=True, note="→ 智擎云 COGS 构成")
    L.reg("Capex_Dep", "DEP_INFER", r, expected=dep["CLOUD_INFER"]); dep_infer_row = r; r += 1
    f = [f'=SUMIFS($H${ir0}:$H${ir1},$B${ir0}:$B${ir1},"<="&{get_column_letter(3+i)}$3)'
         f'-SUM($C${dep_infer_row}:{get_column_letter(3+i)}{dep_infer_row})' for i in range(5)]
    line(ws, r, "推理 GPU 净值", formulas=f, bold=True, note="→ 智擎云 BS")
    L.reg("Capex_Dep", "NET_INFER", r, expected=net["CLOUD_INFER"]); r += 2

    # 自检行：折旧引擎 vs 影子值
    line(ws, r, "★ 引擎自检（折旧/净值 vs 影子值，应全 0）", bold=True); r += 1
    for key in ["DEP_TRAIN", "NET_TRAIN", "DEP_INFER", "NET_INFER"]:
        row = L.row_of[("Capex_Dep", key)]
        shadow = {"DEP_TRAIN": dep["HOLD_TRAIN"], "NET_TRAIN": net["HOLD_TRAIN"],
                  "DEP_INFER": dep["CLOUD_INFER"], "NET_INFER": net["CLOUD_INFER"]}[key]
        f = [f"={get_column_letter(3+i)}{row}-({shadow[i]:.6f})" for i in range(5)]
        line(ws, r, f"  diff {key}", formulas=f, note="绿=0", numFmt=NUM1)
        for ci in range(5):
            ws.cell(row=r, column=3+ci).font = GREEN_LINK
        r += 1
    return ws
