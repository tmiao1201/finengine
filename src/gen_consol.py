"""合并层：Eliminations（抵消分录）+ Consol 三表（加总−抵消）。"""
import assumptions as A
from style import *
from layout import Layout

BS_SHEET = {"HOLD": "HoldCo_BS", "CLOUD": "Cloud_BS_CF", "IND": "Ind_BS_CF", "TOC": "Toc_BS_CF"}
IS_SHEET = {"HOLD": "HoldCo_IS", "CLOUD": "Cloud_IS", "IND": "Ind_IS", "TOC": "Toc_IS"}

def gen_elim(wb, L, eng):
    ws = wb.create_sheet("Eliminations")
    sheet_scaffold(ws, "合并抵消分录底稿", "E0 长投×权益 · E1 关联交易 · E2 内部往来 · E3 少数股东 | 万元",
                   A.YEARS, "consol")
    C = eng["C"]["BS"]; r = 4
    def ar(k, i): return L.ref("Assumptions", k, i)
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent)
        L.reg("Eliminations", key, r, expected=expected); r += 1

    c = ws.cell(row=r, column=1, value="── E0 长期股权投资 × 子公司权益（母公司份额）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    for sub, share, init in [("CLOUD", 1.0, 10000), ("IND", 1.0, 8000), ("TOC", 0.8, 5000)]:
        bs = BS_SHEET[sub]; inj_row = L.row_of[("Assumptions", f"INJECT_{sub}")]
        sfx = "" if share == 1.0 else "*0.8"
        frow(f"e0_{sub}_pi", f"借：实收资本 · {A.ENTITY_NAMES[sub]}",
             [f"={L.ref(bs,'paid_in',i)}{sfx}" for i in range(5)],
             [eng["E"][sub]["BS"]["paid_in"][i]*share for i in range(5)])
        frow(f"e0_{sub}_re", f"借：未分配利润 · {A.ENTITY_NAMES[sub]}",
             [f"={L.ref(bs,'RE',i)}{sfx}" for i in range(5)],
             [eng["E"][sub]["BS"]["RE"][i]*share for i in range(5)])
        lt_exp = [init + sum(A.EQUITY_INJECT[sub][:i+1]) * share for i in range(5)]
        frow(f"e0_{sub}_lt", f"贷：长投 · {A.ENTITY_NAMES[sub]}（成本法）",
             [f"={init}+SUM(Assumptions!$C{inj_row}:{chr(67+i)}{inj_row}){sfx}" for i in range(5)],
             lt_exp)
    frow("e0_capadj", "E0 差额 → 合并资本公积（Σ借−Σ贷）",
         [f"={L.ref('Eliminations','e0_CLOUD_pi',i)}+{L.ref('Eliminations','e0_CLOUD_re',i)}"
          f"-{L.ref('Eliminations','e0_CLOUD_lt',i)}+{L.ref('Eliminations','e0_IND_pi',i)}"
          f"+{L.ref('Eliminations','e0_IND_re',i)}-{L.ref('Eliminations','e0_IND_lt',i)}"
          f"+{L.ref('Eliminations','e0_TOC_pi',i)}+{L.ref('Eliminations','e0_TOC_re',i)}"
          f"-{L.ref('Eliminations','e0_TOC_lt',i)}" for i in range(5)],
         C["cap_adj"], bold=True)

    c = ws.cell(row=r, column=1, value="── E1 关联收入 × 关联成本对冲 ──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("e1", "借：关联收入 / 贷：关联成本（全额）",
         [f"={L.ref('IC_Register','total',i)}" for i in range(5)],
         [L.val("IC_Register", "total", i) for i in range(5)], bold=True)

    c = ws.cell(row=r, column=1, value="── E2 内部往来对冲（双边孰低=调节后）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("e2_ap", "借：内部应付合计",
         [f"={L.ref('IC_Register','sum_ap',i)}" for i in range(5)],
         [L.val("IC_Register", "sum_ap", i) for i in range(5)])
    frow("e2_ar", "贷：内部应收合计（孰低=ΣAR）",
         [f"={L.ref('IC_Register','sum_ar',i)}" for i in range(5)],
         [L.val("IC_Register", "sum_ar", i) for i in range(5)])
    frow("e2_op", "差额 → 其他应付款·内部在途（=ΣAP−ΣAR）",
         [f"={L.ref('Eliminations','e2_ap',i)}-{L.ref('Eliminations','e2_ar',i)}" for i in range(5)],
         C["other_payable"], "cut-off 在途挂账", bold=True)

    c = ws.cell(row=r, column=1, value="── E3 少数股东（利润拆分，非借贷）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("e3", "少数股东损益 = 互联 NI × 20%",
         [f"={L.ref('Toc_IS','ni',i)}*0.2" for i in range(5)],
         eng["C"]["IS"]["ni_minority"])
    frow("e3_eq", "少数股东权益 = 互联净资产 × 20%",
         [f"=({L.ref('Toc_BS_CF','paid_in',i)}+{L.ref('Toc_BS_CF','RE',i)})*0.2" for i in range(5)],
         C["minority"])
    return ws

def gen_consol_is(wb, L, eng):
    ws = wb.create_sheet("Consol_IS")
    sheet_scaffold(ws, "智擎集团 · 合并利润表", "Σ四主体 − E1 关联对冲 − E3 归属拆分 | 万元", A.YEARS, "consol")
    I = eng["C"]["IS"]; r = 4
    def s4(sheet_key, i, skip=("HOLD",)):
        return "+".join(L.ref(IS_SHEET[e], sheet_key, i) for e in ["CLOUD", "IND", "TOC", "HOLD"] if e not in skip)
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1, fill=None):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent, fill=fill)
        L.reg("Consol_IS", key, r, expected=expected); r += 1

    frow("rev_ext", "对外收入合计（云+行业+互联）",
         [f"={s4('rev_ext',i)}" for i in range(5)], I["rev_ext"], bold=True)
    frow("rev_elim", "减：关联收入抵消（E1）",
         [f"=-{L.ref('Eliminations','e1',i)}" for i in range(5)],
         [-L.val("Eliminations", "e1", i) for i in range(5)], "T1+T2+T3 全消")
    frow("rev", "合并营业收入", [f"={L.ref('Consol_IS','rev_ext',i)}" for i in range(5)], I["rev"], bold=True)
    frow("cogs_sum", "Σ四主体营业成本",
         [f"={'+'.join(L.ref(IS_SHEET[e],'cogs',i) for e in IS_SHEET)}" for i in range(5)],
         [sum(eng["E"][e]["IS"]["cogs"][i] for e in A.ENTITIES) for i in range(5)])
    frow("cogs_elim", "减：关联成本对冲（E1）",
         [f"=-{L.ref('Eliminations','e1',i)}" for i in range(5)],
         [-L.val("Eliminations", "e1", i) for i in range(5)],
         "含保留云真实服务成本，净效应=−0.45×T3")
    frow("cogs", "合并营业成本",
         [f"={L.ref('Consol_IS','cogs_sum',i)}+{L.ref('Consol_IS','cogs_elim',i)}" for i in range(5)],
         I["cogs"], bold=True)
    frow("gp", "合并毛利润",
         [f"={L.ref('Consol_IS','rev',i)}-{L.ref('Consol_IS','cogs',i)}" for i in range(5)], I["gp"], bold=True)
    frow("rd", "研发费用", [f"={s4('rd',i,skip=())}" for i in range(5)],
         [sum(eng["E"][e]["IS"]["rd"][i] for e in A.ENTITIES) for i in range(5)])
    frow("sales", "销售费用", [f"={s4('sales',i,skip=())}" for i in range(5)],
         [sum(eng["E"][e]["IS"]["sales"][i] for e in A.ENTITIES) for i in range(5)])
    frow("gna", "管理费用", [f"={s4('gna',i,skip=())}" for i in range(5)],
         [sum(eng["E"][e]["IS"]["gna"][i] for e in A.ENTITIES) for i in range(5)])
    frow("ebit", "合并 EBIT",
         [f"={L.ref('Consol_IS','gp',i)}-{L.ref('Consol_IS','rd',i)}-{L.ref('Consol_IS','sales',i)}"
          f"-{L.ref('Consol_IS','gna',i)}" for i in range(5)], I["ebit"], bold=True)
    frow("tax", "所得税（Σ）",
         [f"={s4('tax',i,skip=())}" for i in range(5)],
         [sum(eng["E"][e]["IS"]["tax"][i] for e in A.ENTITIES) for i in range(5)])
    frow("ni", "合并净利润",
         [f"={L.ref('Consol_IS','ebit',i)}-{L.ref('Consol_IS','tax',i)}" for i in range(5)], I["ni"], bold=True)
    frow("ni_minority", "少数股东损益（E3）",
         [f"={L.ref('Eliminations','e3',i)}" for i in range(5)], I["ni_minority"])
    frow("ni_parent", "归属于母公司净利润",
         [f"={L.ref('Consol_IS','ni',i)}-{L.ref('Consol_IS','ni_minority',i)}" for i in range(5)],
         I["ni_parent"], bold=True, fill=TOTAL_FILL)
    return ws

def gen_consol_bs(wb, L, eng):
    ws = wb.create_sheet("Consol_BS")
    sheet_scaffold(ws, "智擎集团 · 合并资产负债表", "Σ四主体 − E0 − E2 + 少数股东权益 | 万元", A.YEARS, "consol")
    B = eng["C"]["BS"]; r = 4
    def b4(key, i):
        parts = []
        for e in BS_SHEET:
            try:
                parts.append(L.ref(BS_SHEET[e], key, i))
            except KeyError:
                pass  # 该主体无此科目（如行业/互联无 GPU）
        return "+".join(parts) if parts else "0"
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1, fill=None):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent, fill=fill)
        L.reg("Consol_BS", key, r, expected=expected); r += 1

    frow("cash", "货币资金", [f"={b4('cash',i)}" for i in range(5)], B["cash"])
    frow("AR", "应收账款", [f"={b4('AR',i)}" for i in range(5)], B["AR"])
    frow("fa", "固定资产 · GPU 净值", [f"={b4('fa',i)}" for i in range(5)], B["fa"])
    frow("ic_ar", "内部应收（抵消后 0）",
         [f"={b4('IC_AR',i)}-{L.ref('Eliminations','e2_ar',i)}" for i in range(5)], [0.0]*5)
    frow("lt_inv", "长期股权投资（抵消后 0）",
         [f"={L.ref('HoldCo_BS','lt_inv',i)}-{L.ref('Eliminations','e0_CLOUD_lt',i)}"
          f"-{L.ref('Eliminations','e0_IND_lt',i)}-{L.ref('Eliminations','e0_TOC_lt',i)}" for i in range(5)],
         [0.0]*5)
    frow("assets", "资产总计",
         [f"={L.ref('Consol_BS','cash',i)}+{L.ref('Consol_BS','AR',i)}+{L.ref('Consol_BS','fa',i)}"
          for i in range(5)], B["c_assets"], bold=True, indent=0)
    r += 1
    frow("AP", "应付账款", [f"={b4('AP',i)}" for i in range(5)], B["AP"])
    frow("deferred", "递延收益", [f"={b4('deferred',i)}" for i in range(5)], B["deferred"])
    frow("ic_ap", "内部应付（抵消后=在途差额）",
         [f"={b4('IC_AP',i)}-{L.ref('Eliminations','e2_ar',i)}" for i in range(5)],
         [B["other_payable"][i] for i in range(5)],
         "按双边孰低抵消，剩余挂 other_pay")
    frow("other_pay", "其他应付款 · 内部在途",
         [f"={L.ref('Eliminations','e2_op',i)}" for i in range(5)], B["other_payable"])
    frow("liab", "负债合计",
         [f"={L.ref('Consol_BS','AP',i)}+{L.ref('Consol_BS','deferred',i)}+{L.ref('Consol_BS','other_pay',i)}"
          for i in range(5)], B["c_liab"], bold=True, indent=0)
    frow("paid_in_hold", "实收资本（研究院口径）",
         [f"={L.ref('HoldCo_BS','paid_in',i)}" for i in range(5)],
         [eng["E"]["HOLD"]["BS"]["paid_in"][i] for i in range(5)], "子公司全抵")
    frow("cap_adj", "合并调整 · 资本公积（E0 差额）",
         [f"={L.ref('Eliminations','e0_capadj',i)}" for i in range(5)], B["cap_adj"])
    frow("re_hold", "未分配利润（研究院口径）",
         [f"={L.ref('HoldCo_BS','RE',i)}" for i in range(5)],
         [eng["E"]["HOLD"]["BS"]["RE"][i] for i in range(5)])
    frow("minority", "少数股东权益（E3）",
         [f"={L.ref('Eliminations','e3_eq',i)}" for i in range(5)], B["minority"])
    frow("equity", "合并所有者权益合计",
         [f"={L.ref('Consol_BS','paid_in_hold',i)}+{L.ref('Consol_BS','cap_adj',i)}"
          f"+{L.ref('Consol_BS','re_hold',i)}+{L.ref('Consol_BS','minority',i)}" for i in range(5)],
         B["consol_eq"], bold=True, indent=0)
    frow("L_E", "负债+权益合计",
         [f"={L.ref('Consol_BS','liab',i)}+{L.ref('Consol_BS','equity',i)}" for i in range(5)],
         [B["c_liab"][i]+B["consol_eq"][i] for i in range(5)], bold=True, indent=0)
    line(ws, r, "★ 合并 BS 平衡差（应全 0）", bold=True); r += 1
    L.reg("Consol_BS", "chk_balance", r)
    line(ws, r, "  资产 − 负债 − 权益",
         formulas=[f"={L.ref('Consol_BS','assets',i)}-{L.ref('Consol_BS','L_E',i)}" for i in range(5)],
         numFmt=NUM1); r += 1
    return ws

def gen_consol_cf(wb, L, eng):
    ws = wb.create_sheet("Consol_CF")
    sheet_scaffold(ws, "智擎集团 · 合并现金流量表（间接法）", "合并 NI+D&A+ΔWC；内部结算不出合并范围自动消失 | 万元",
                   A.YEARS, "consol")
    F = eng["C"]["CF"]; r = 4
    def cref(k, i): return L.ref("Consol_BS", k, i)
    def iref(k, i): return L.ref("Consol_IS", k, i)
    def pref(k, i): return cref(k, i-1) if i > 0 else "0"
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1, fill=None):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent, fill=fill)
        L.reg("Consol_CF", key, r, expected=expected); r += 1

    frow("ni", "合并净利润", [f"={iref('ni',i)}" for i in range(5)], F["ni"])
    frow("da", "加：折旧（训练+推理）",
         [f"={L.ref('Capex_Dep','DEP_TRAIN',i)}+{L.ref('Capex_Dep','DEP_INFER',i)}" for i in range(5)], F["da"])
    frow("d_ar", "减：应收增加", [f"=-({cref('AR',i)}-({pref('AR',i)}))" for i in range(5)], F["d_ar"])
    frow("d_def", "加：递延增加", [f"={cref('deferred',i)}-({pref('deferred',i)})" for i in range(5)], F["d_deferred"])
    frow("d_ap", "加：应付增加", [f"={cref('AP',i)}-({pref('AP',i)})" for i in range(5)], F["d_ap"])
    frow("d_op", "加：在途挂账变动", [f"={cref('other_pay',i)}-({pref('other_pay',i)})" for i in range(5)],
         F.get("d_op", [0.0]*5), "非现金")
    frow("cfo", "经营活动现金流",
         [f"={L.ref('Consol_CF','ni',i)}+{L.ref('Consol_CF','da',i)}+{L.ref('Consol_CF','d_ar',i)}"
          f"+{L.ref('Consol_CF','d_def',i)}+{L.ref('Consol_CF','d_ap',i)}+{L.ref('Consol_CF','d_op',i)}"
          for i in range(5)], F["cfo"], bold=True)
    frow("cfi", "投资活动现金流",
         [f"=-{L.ref('Capex_Dep','CAPEX_TRAIN',i)}-{L.ref('Capex_Dep','CAPEX_INFER',i)}" for i in range(5)],
         F["cfi"])
    frow("cff", "筹资活动（研究院注资+互联战投）",
         [f"={L.ref('Assumptions','INJECT_HOLD',i)}+{L.ref('Assumptions','INJECT_TOC',i)}*0.2" for i in range(5)],
         F["cff"], "子公司注资=内部划转，抵消")
    frow("net", "现金净变动",
         [f"={L.ref('Consol_CF','cfo',i)}+{L.ref('Consol_CF','cfi',i)}+{L.ref('Consol_CF','cff',i)}"
          for i in range(5)], F["net"], bold=True)
    row = r
    frow("end_cash", "期末现金（CF 推导）",
         [f"={F['end_cash'][0]-F['net'][0]}+{L.ref('Consol_CF','net',0)}"] +
         [f"={chr(66+i)}{row}+{L.ref('Consol_CF','net',i)}" for i in range(1,5)],
         F["end_cash"], bold=True)
    frow("bs_cash", "期末现金（Σ单体 BS）",
         [f"={'+'.join(L.ref(BS_SHEET[e],'cash',i) for e in BS_SHEET)}" for i in range(5)],
         F["bs_cash"])
    line(ws, r, "★ 合并 CF=BS 差异（应全 0）", bold=True); r += 1
    L.reg("Consol_CF", "chk_tie", r)
    line(ws, r, "  CF推导 − Σ单体现金",
         formulas=[f"={L.ref('Consol_CF','end_cash',i)}-{L.ref('Consol_CF','bs_cash',i)}" for i in range(5)],
         numFmt=NUM1); r += 1
    return ws

def gen(wb, L, eng):
    gen_elim(wb, L, eng)
    gen_consol_is(wb, L, eng)
    gen_consol_bs(wb, L, eng)
    gen_consol_cf(wb, L, eng)
