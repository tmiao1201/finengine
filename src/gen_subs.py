"""三家子公司生成器：IS 全量 + BS&CF 关键科目（半厚模式，1b）。
每家一张 IS + 一张 BS&CF 合并表，底部自带勾稽差异行。"""
import assumptions as A
from style import *
from layout import Layout

# ─────────────────── 智擎云 ───────────────────
def cloud_is(wb, L, eng):
    ws = wb.create_sheet("Cloud_IS")
    sheet_scaffold(ws, "智擎云 · 利润表", "单体 | MaaS API：token 量 × 单价 | 万元", A.YEARS, "engine")
    E = eng["E"]["CLOUD"]["IS"]; r = 4
    def ar(k, i): return L.ref("Assumptions", k, i)
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent)
        L.reg("Cloud_IS", key, r, expected=expected); r += 1

    frow("rev_ext", "对外收入 · MaaS API",
         [f"={ar('CLOUD_TOKENS_B',i)}*36.5*{ar('CLOUD_PRICE',i)}" for i in range(5)],
         E["rev_ext"], "日均token×365×单价/1000")
    frow("rev_ic", "关联收入 · T3 算力服务（互联）",
         [f"={ar('TOC_MAU_W',i)}*{ar('TOC_UCOST',i)}*12*{ar('T3_SHARE',i)}" for i in range(5)],
         E["rev_ic"], "互联推理成本×70%")
    frow("rev", "营业收入合计",
         [f"={L.ref('Cloud_IS','rev_ext',i)}+{L.ref('Cloud_IS','rev_ic',i)}" for i in range(5)],
         E["rev"], bold=True)
    frow("cogs_ext", "成本 · 推理算力（对外部分）",
         [f"={ar('CLOUD_TOKENS_B',i)}*36.5*{ar('CLOUD_UCOST',i)}" for i in range(5)],
         [E["cogs_ext"][i] - A.CLOUD_BW[i] for i in range(5)])
    frow("cogs_ic", "成本 · T3 服务成本",
         [f"={L.ref('Cloud_IS','rev_ic',i)}*(1-{ar('T3_GM',i)})" for i in range(5)], E["cogs_ic"],
         "内部价毛利率45%")
    frow("cogs_bw", "成本 · 带宽机房", [f"={ar('CLOUD_BW',i)}" for i in range(5)], None)
    L.reg("Cloud_IS", "cogs_bw", L.row_of[("Cloud_IS","cogs_bw")], expected=A.CLOUD_BW)
    frow("cogs_t1", "成本 · T1 license 费",
         [f"={L.ref('Cloud_IS','rev_ext',i)}*{ar('T1_RATE_CLOUD',i)}" for i in range(5)], E["cogs_ic_t1"])
    frow("cogs", "营业成本合计",
         [f"=SUM({L.ref('Cloud_IS','cogs_ext',i)},{L.ref('Cloud_IS','cogs_ic',i)},"
          f"{L.ref('Cloud_IS','cogs_bw',i)},{L.ref('Cloud_IS','cogs_t1',i)})" for i in range(5)],
         E["cogs"], bold=True)
    frow("gp", "毛利润",
         [f"={L.ref('Cloud_IS','rev',i)}-{L.ref('Cloud_IS','cogs',i)}" for i in range(5)], E["gp"], bold=True)
    frow("rd", "研发费用",
         [f"={ar('CLOUD_RD_HC',i)}*70" for i in range(5)], E["rd"], "编制×人均70万")
    frow("sales", "销售费用", [f"={ar('CLOUD_SALES_HC',i)}*60" for i in range(5)], E["sales"], "×60万")
    frow("gna", "管理费用", [f"={ar('CLOUD_GNA_HC',i)}*55" for i in range(5)], E["gna"], "×55万")
    frow("ebit", "EBIT",
         [f"={L.ref('Cloud_IS','gp',i)}-{L.ref('Cloud_IS','rd',i)}-{L.ref('Cloud_IS','sales',i)}"
          f"-{L.ref('Cloud_IS','gna',i)}" for i in range(5)], E["ebit"], bold=True)
    frow("tax", "所得税", [f"=MAX(0,{L.ref('Cloud_IS','ebit',i)})*{ar('TAX_RATE',i)}" for i in range(5)], E["tax"])
    frow("ni", "净利润",
         [f"={L.ref('Cloud_IS','ebit',i)}-{L.ref('Cloud_IS','tax',i)}" for i in range(5)], E["ni"], bold=True)
    return ws

def cloud_bscf(wb, L, eng):
    ws = wb.create_sheet("Cloud_BS_CF")
    sheet_scaffold(ws, "智擎云 · 资产负债表+现金流量表", "单体半厚 | BS在上 CF在下 | 万元", A.YEARS, "engine")
    B, F = eng["E"]["CLOUD"]["BS"], eng["E"]["CLOUD"]["CF"]; r = 4
    def ar(k, i): return L.ref("Assumptions", k, i)
    def iref(k, i): return L.ref("Cloud_IS", k, i)
    def sref(k, i): return L.ref("Cloud_BS_CF", k, i)
    def pref(k, i): return sref(k, i-1) if i > 0 else "0"
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1, fill=None):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent, fill=fill)
        L.reg("Cloud_BS_CF", key, r, expected=expected); r += 1

    c = ws.cell(row=r, column=1, value="── 资产负债表（关键科目）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("AR", "应收账款", [f"={iref('rev_ext',i)}*{ar('DSO',i)}/365" for i in range(5)], B["AR"])
    icar = [f"=({ar('TOC_MAU_W',i)}*{ar('TOC_UCOST',i)}*12*{ar('T3_SHARE',i)})*{ar('IC_DAYS',i)}/365"
            for i in range(5)]
    icar[4] = icar[4] + f"-{ar('CUTOFF_2027',4)}"
    frow("IC_AR", "内部应收（互联 T3）", icar, B["IC_AR"], "2027 含 cut-off 未开票 −220")
    frow("fa", "固定资产 · 推理 GPU 净值", [f"={L.ref('Capex_Dep','NET_INFER',i)}" for i in range(5)], B["fa"])
    frow("assets_nc", "非现金资产合计",
         [f"={sref('AR',i)}+{sref('IC_AR',i)}+{sref('fa',i)}" for i in range(5)],
         [B["AR"][i]+B["IC_AR"][i]+B["fa"][i] for i in range(5)], bold=True, indent=0)
    frow("AP", "应付账款",
         [f"=({iref('cogs',i)}+{iref('rd',i)}+{iref('sales',i)}+{iref('gna',i)}"
          f"-{L.ref('Capex_Dep','DEP_INFER',i)})*{ar('DPO',i)}/365" for i in range(5)], B["AP"])
    frow("IC_AP", "内部应付（研究院 T1）",
         [f"={iref('cogs_t1',i)}*{ar('IC_DAYS',i)}/365" for i in range(5)], B["IC_AP"])
    frow("deferred", "递延收益", [f"={iref('rev',i)}*{ar('DEF_RATE_CLOUD',i)}" for i in range(5)], B["deferred"])
    frow("liab", "负债合计", [f"={sref('AP',i)}+{sref('IC_AP',i)}+{sref('deferred',i)}" for i in range(5)],
         [B["AP"][i]+B["IC_AP"][i]+B["deferred"][i] for i in range(5)], bold=True, indent=0)
    row = r
    frow("paid_in", "实收资本",
         [f"={ar('OPEN_PI_CLOUD',0)}+{ar('INJECT_CLOUD',0)}"] +
         [f"={chr(66+i)}{row}+{ar('INJECT_CLOUD',i)}" for i in range(1,5)], B["paid_in"])
    row = r
    frow("RE", "未分配利润",
         [f"={ar('OPEN_RE_CLOUD',0)}+{iref('ni',0)}"] +
         [f"={chr(66+i)}{row}+{iref('ni',i)}" for i in range(1,5)], B["RE"])
    frow("cash", "货币资金（倒挤）",
         [f"={sref('liab',i)}+{sref('paid_in',i)}+{sref('RE',i)}-{sref('assets_nc',i)}" for i in range(5)],
         B["cash"], bold=True, indent=0, fill=TOTAL_FILL)

    c = ws.cell(row=r, column=1, value="── 现金流量表（间接法）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("cf_ni", "净利润", [f"={iref('ni',i)}" for i in range(5)], F["ni"])
    frow("cf_da", "加：折旧", [f"={L.ref('Capex_Dep','DEP_INFER',i)}" for i in range(5)], F["da"])
    frow("cf_d_ar", "减：应收增加", [f"=-({sref('AR',i)}-({pref('AR',i)}))" for i in range(5)], F["d_ar"])
    frow("cf_d_ic", "内部往来净变动",
         [f"=({sref('IC_AP',i)}-({pref('IC_AP',i)}))-({sref('IC_AR',i)}-({pref('IC_AR',i)}))" for i in range(5)],
         F["d_ic_wc"])
    frow("cf_d_def", "加：递延增加", [f"={sref('deferred',i)}-({pref('deferred',i)})" for i in range(5)], F["d_deferred"])
    frow("cf_d_ap", "加：应付增加", [f"={sref('AP',i)}-({pref('AP',i)})" for i in range(5)], F["d_ap"])
    frow("cfo", "经营现金流",
         [f"={sref('cf_ni',i)}+{sref('cf_da',i)}+{sref('cf_d_ar',i)}+{sref('cf_d_ic',i)}"
          f"+{sref('cf_d_def',i)}+{sref('cf_d_ap',i)}" for i in range(5)], F["cfo"], bold=True)
    frow("cfi", "投资现金流", [f"=-{L.ref('Capex_Dep','CAPEX_INFER',i)}" for i in range(5)], F["cfi"])
    frow("cff", "筹资：股权注资", [f"={ar('INJECT_CLOUD',i)}" for i in range(5)], F["cff"])
    frow("net", "现金净变动",
         [f"={sref('cfo',i)}+{sref('cfi',i)}+{sref('cff',i)}" for i in range(5)], F["net"], bold=True)
    op = A.OPENING["CLOUD"]["paid_in"] + A.OPENING["CLOUD"]["RE"] - A.OPENING_FA["CLOUD"]
    row = r
    frow("end_cash", "期末现金（CF）",
         [f"={op}+{sref('net',0)}"] + [f"={chr(66+i)}{row}+{sref('net',i)}" for i in range(1,5)],
         F["end_cash"], bold=True)
    line(ws, r, "★ CF=BS 差异（应全 0）", bold=True); r += 1
    L.reg("Cloud_BS_CF", "chk_tie", r)
    line(ws, r, "  CF推导 − BS倒挤",
         formulas=[f"={sref('end_cash',i)}-{sref('cash',i)}" for i in range(5)], numFmt=NUM1); r += 1
    return ws

# ─────────────────── 智擎行业 ───────────────────
def ind_is(wb, L, eng):
    ws = wb.create_sheet("Ind_IS")
    sheet_scaffold(ws, "智擎行业 · 利润表", "单体 | 私有化部署 + 行业解决方案 双业务线 | 万元", A.YEARS, "engine")
    E = eng["E"]["IND"]["IS"]; r = 4
    def ar(k, i): return L.ref("Assumptions", k, i)
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent)
        L.reg("Ind_IS", key, r, expected=expected); r += 1

    frow("rev_pvt", "收入 · 私有化部署",
         [f"={ar('IND_PROJECTS',i)}*{ar('IND_CONTRACT',i)}" for i in range(5)],
         [A.IND_PROJECTS[i]*A.IND_CONTRACT[i] for i in range(5)], "项目数×合同额")
    frow("rev_sol", "收入 · 行业解决方案",
         [f"={ar('IND_SUBSCRIBERS',i)}*{ar('IND_SUB_FEE',i)}" for i in range(5)],
         [A.IND_SUBSCRIBERS[i]*A.IND_SUB_FEE[i] for i in range(5)], "签约客户×年费")
    frow("rev_ext", "对外收入合计",
         [f"={L.ref('Ind_IS','rev_pvt',i)}+{L.ref('Ind_IS','rev_sol',i)}" for i in range(5)],
         E["rev_ext"], bold=True)
    frow("cogs_pvt", "成本 · 私有化（硬件+人力）",
         [f"={L.ref('Ind_IS','rev_pvt',i)}*{ar('IND_COGS_PVT',i)}" for i in range(5)],
         [A.IND_PROJECTS[i]*A.IND_CONTRACT[i]*A.IND_COGS_PVT for i in range(5)])
    frow("cogs_sol", "成本 · 解决方案",
         [f"={L.ref('Ind_IS','rev_sol',i)}*{ar('IND_COGS_SOL',i)}" for i in range(5)],
         [A.IND_SUBSCRIBERS[i]*A.IND_SUB_FEE[i]*A.IND_COGS_SOL for i in range(5)])
    frow("cogs_t1", "成本 · T1 license 费",
         [f"={L.ref('Ind_IS','rev_ext',i)}*{ar('T1_RATE_IND',i)}" for i in range(5)], E["cogs_ic_t1"])
    frow("cogs_t2", "成本 · T2 定制研发",
         [f"={ar('T2_FEE',i)}" for i in range(5)], E["cogs_ic_t2"])
    frow("cogs", "营业成本合计",
         [f"=SUM({L.ref('Ind_IS','cogs_pvt',i)},{L.ref('Ind_IS','cogs_sol',i)},"
          f"{L.ref('Ind_IS','cogs_t1',i)},{L.ref('Ind_IS','cogs_t2',i)})" for i in range(5)],
         E["cogs"], bold=True)
    frow("gp", "毛利润",
         [f"={L.ref('Ind_IS','rev_ext',i)}-{L.ref('Ind_IS','cogs',i)}" for i in range(5)], E["gp"], bold=True)
    frow("rd", "研发费用", [f"={ar('IND_RD_HC',i)}*70" for i in range(5)], E["rd"], "×70万")
    frow("sales", "销售费用", [f"={ar('IND_SALES_HC',i)}*60" for i in range(5)], E["sales"], "×60万")
    frow("gna", "管理费用", [f"={ar('IND_GNA_HC',i)}*55" for i in range(5)], E["gna"], "×55万")
    frow("ebit", "EBIT",
         [f"={L.ref('Ind_IS','gp',i)}-{L.ref('Ind_IS','rd',i)}-{L.ref('Ind_IS','sales',i)}"
          f"-{L.ref('Ind_IS','gna',i)}" for i in range(5)], E["ebit"], bold=True)
    frow("tax", "所得税", [f"=MAX(0,{L.ref('Ind_IS','ebit',i)})*{ar('TAX_RATE',i)}" for i in range(5)], E["tax"])
    frow("ni", "净利润",
         [f"={L.ref('Ind_IS','ebit',i)}-{L.ref('Ind_IS','tax',i)}" for i in range(5)], E["ni"], bold=True)
    return ws

def ind_bscf(wb, L, eng):
    ws = wb.create_sheet("Ind_BS_CF")
    sheet_scaffold(ws, "智擎行业 · 资产负债表+现金流量表", "单体半厚 | 无 GPU 资产 | 万元", A.YEARS, "engine")
    B, F = eng["E"]["IND"]["BS"], eng["E"]["IND"]["CF"]; r = 4
    def ar(k, i): return L.ref("Assumptions", k, i)
    def iref(k, i): return L.ref("Ind_IS", k, i)
    def sref(k, i): return L.ref("Ind_BS_CF", k, i)
    def pref(k, i): return sref(k, i-1) if i > 0 else "0"
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1, fill=None):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent, fill=fill)
        L.reg("Ind_BS_CF", key, r, expected=expected); r += 1

    c = ws.cell(row=r, column=1, value="── 资产负债表（关键科目）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("AR", "应收账款", [f"={iref('rev_ext',i)}*{ar('DSO',i)}/365" for i in range(5)], B["AR"])
    frow("IC_AP", "内部应付（T1+T2）",
         [f"=({iref('cogs_t1',i)}+{iref('cogs_t2',i)})*{ar('IC_DAYS',i)}/365" for i in range(5)], B["IC_AP"])
    frow("AP", "应付账款",
         [f"=({iref('cogs',i)}+{iref('rd',i)}+{iref('sales',i)}+{iref('gna',i)})*{ar('DPO',i)}/365"
          for i in range(5)], B["AP"], "无折旧")
    frow("deferred", "递延收益", [f"={iref('rev_ext',i)}*{ar('DEF_RATE_IND',i)}" for i in range(5)], B["deferred"])
    frow("liab", "负债合计", [f"={sref('AP',i)}+{sref('IC_AP',i)}+{sref('deferred',i)}" for i in range(5)],
         [B["AP"][i]+B["IC_AP"][i]+B["deferred"][i] for i in range(5)], bold=True, indent=0)
    row = r
    frow("paid_in", "实收资本",
         [f"={ar('OPEN_PI_IND',0)}+{ar('INJECT_IND',0)}"] +
         [f"={chr(66+i)}{row}+{ar('INJECT_IND',i)}" for i in range(1,5)], B["paid_in"])
    row = r
    frow("RE", "未分配利润",
         [f"={ar('OPEN_RE_IND',0)}+{iref('ni',0)}"] +
         [f"={chr(66+i)}{row}+{iref('ni',i)}" for i in range(1,5)], B["RE"])
    frow("cash", "货币资金（倒挤）",
         [f"={sref('liab',i)}+{sref('paid_in',i)}+{sref('RE',i)}-{sref('AR',i)}" for i in range(5)],
         B["cash"], bold=True, indent=0, fill=TOTAL_FILL)

    c = ws.cell(row=r, column=1, value="── 现金流量表（间接法）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("cf_ni", "净利润", [f"={iref('ni',i)}" for i in range(5)], F["ni"])
    frow("cf_d_ar", "减：应收增加", [f"=-({sref('AR',i)}-({pref('AR',i)}))" for i in range(5)], F["d_ar"])
    frow("cf_d_ic", "内部往来净变动",
         [f"=({sref('IC_AP',i)}-({pref('IC_AP',i)}))" for i in range(5)], F["d_ic_wc"])
    frow("cf_d_def", "加：递延增加", [f"={sref('deferred',i)}-({pref('deferred',i)})" for i in range(5)], F["d_deferred"])
    frow("cf_d_ap", "加：应付增加", [f"={sref('AP',i)}-({pref('AP',i)})" for i in range(5)], F["d_ap"])
    frow("cfo", "经营现金流",
         [f"={sref('cf_ni',i)}+{sref('cf_d_ar',i)}+{sref('cf_d_ic',i)}+{sref('cf_d_def',i)}"
          f"+{sref('cf_d_ap',i)}" for i in range(5)], F["cfo"], bold=True)
    frow("cff", "筹资：股权注资", [f"={ar('INJECT_IND',i)}" for i in range(5)], F["cff"])
    frow("net", "现金净变动", [f"={sref('cfo',i)}+{sref('cff',i)}" for i in range(5)], F["net"], bold=True)
    op = A.OPENING["IND"]["paid_in"] + A.OPENING["IND"]["RE"]
    row = r
    frow("end_cash", "期末现金（CF）",
         [f"={op}+{sref('cfo',0)}+{sref('cff',0)}"] +
         [f"={chr(66+i)}{row}+{sref('cfo',i)}+{sref('cff',i)}" for i in range(1,5)],
         F["end_cash"], bold=True)
    line(ws, r, "★ CF=BS 差异（应全 0）", bold=True); r += 1
    L.reg("Ind_BS_CF", "chk_tie", r)
    line(ws, r, "  CF推导 − BS倒挤",
         formulas=[f"={sref('end_cash',i)}-{sref('cash',i)}" for i in range(5)], numFmt=NUM1); r += 1
    return ws

# ─────────────────── 智擎互联 ───────────────────
def F_cac(i):
    prev = A.TOC_MAU_W[i-1] if i > 0 else 0
    return (A.TOC_MAU_W[i]-prev)*A.TOC_NEW_U_MULTIPLIER*A.TOC_CAC[i]

def toc_is(wb, L, eng):
    ws = wb.create_sheet("Toc_IS")
    sheet_scaffold(ws, "智擎互联 · 利润表", "单体 | C端：MAU×转化×ARPPU + 广告 | 万元", A.YEARS, "engine")
    E = eng["E"]["TOC"]["IS"]; r = 4
    def ar(k, i): return L.ref("Assumptions", k, i)
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent)
        L.reg("Toc_IS", key, r, expected=expected); r += 1

    frow("rev_sub", "收入 · C端订阅",
         [f"={ar('TOC_MAU_W',i)}*{ar('TOC_PAYRATE',i)}*{ar('TOC_ARPPU',i)}*12" for i in range(5)],
         [A.TOC_MAU_W[i]*A.TOC_PAYRATE[i]*A.TOC_ARPPU[i]*12 for i in range(5)])
    frow("rev_ads", "收入 · 广告", [f"={ar('TOC_ADS',i)}" for i in range(5)], A.TOC_ADS)
    frow("rev_ext", "对外收入合计",
         [f"={L.ref('Toc_IS','rev_sub',i)}+{L.ref('Toc_IS','rev_ads',i)}" for i in range(5)],
         E["rev_ext"], bold=True)
    frow("cogs_ext", "成本 · 推理+带宽",
         [f"={ar('TOC_MAU_W',i)}*{ar('TOC_UCOST',i)}*12" for i in range(5)], E["cogs_ext"],
         "含付云 T3（70%）")
    frow("cogs_t1", "成本 · T1 license 费",
         [f"={L.ref('Toc_IS','rev_ext',i)}*{ar('T1_RATE_TOC',i)}" for i in range(5)], E["cogs_ic_t1"])
    frow("cogs", "营业成本合计",
         [f"={L.ref('Toc_IS','cogs_ext',i)}+{L.ref('Toc_IS','cogs_t1',i)}" for i in range(5)],
         E["cogs"], bold=True)
    frow("gp", "毛利润",
         [f"={L.ref('Toc_IS','rev_ext',i)}-{L.ref('Toc_IS','cogs',i)}" for i in range(5)], E["gp"], bold=True)
    frow("rd", "研发费用", [f"={ar('TOC_RD_HC',i)}*70" for i in range(5)], E["rd"], "×70万")
    frow("sales_cac", "销售费用 · 获客",
         [f"=({ar('TOC_MAU_W',i)}-{ar('TOC_MAU_W',i-1) if i>0 else 0})*{ar('TOC_NEW_U_MULTIPLIER',0)}"
          f"*{ar('TOC_CAC',i)}" for i in range(5)], None, "新增用户×CAC")
    L.reg("Toc_IS", "sales_cac", L.row_of[("Toc_IS","sales_cac")],
          expected=[F_cac(i) for i in range(5)])
    frow("sales", "销售费用合计",
         [f"={L.ref('Toc_IS','sales_cac',i)}+{ar('TOC_SALES_HC',i)}*55" for i in range(5)], E["sales"])
    frow("gna", "管理费用", [f"={ar('TOC_GNA_HC',i)}*55" for i in range(5)], E["gna"], "×55万")
    frow("ebit", "EBIT",
         [f"={L.ref('Toc_IS','gp',i)}-{L.ref('Toc_IS','rd',i)}-{L.ref('Toc_IS','sales',i)}"
          f"-{L.ref('Toc_IS','gna',i)}" for i in range(5)], E["ebit"], bold=True)
    frow("tax", "所得税", [f"=MAX(0,{L.ref('Toc_IS','ebit',i)})*{ar('TAX_RATE',i)}" for i in range(5)], E["tax"])
    frow("ni", "净利润",
         [f"={L.ref('Toc_IS','ebit',i)}-{L.ref('Toc_IS','tax',i)}" for i in range(5)], E["ni"],
         "归母80%", bold=True)
    return ws

def toc_bscf(wb, L, eng):
    ws = wb.create_sheet("Toc_BS_CF")
    sheet_scaffold(ws, "智擎互联 · 资产负债表+现金流量表", "单体半厚 | 母公司持股80% | 万元", A.YEARS, "engine")
    B, F = eng["E"]["TOC"]["BS"], eng["E"]["TOC"]["CF"]; r = 4
    def ar(k, i): return L.ref("Assumptions", k, i)
    def iref(k, i): return L.ref("Toc_IS", k, i)
    def sref(k, i): return L.ref("Toc_BS_CF", k, i)
    def pref(k, i): return sref(k, i-1) if i > 0 else "0"
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1, fill=None):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent, fill=fill)
        L.reg("Toc_BS_CF", key, r, expected=expected); r += 1

    c = ws.cell(row=r, column=1, value="── 资产负债表（关键科目）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("AR", "应收账款", [f"={iref('rev_ext',i)}*{ar('DSO',i)}/365" for i in range(5)], B["AR"],
         "广告主账期")
    frow("IC_AP", "内部应付（T3+T1）",
         [f"=({ar('TOC_MAU_W',i)}*{ar('TOC_UCOST',i)}*12*{ar('T3_SHARE',i)}+{iref('cogs_t1',i)})"
          f"*{ar('IC_DAYS',i)}/365" for i in range(5)], B["IC_AP"], "含在途，全挂")
    frow("AP", "应付账款",
         [f"=({iref('cogs',i)}+{iref('rd',i)}+{iref('sales',i)}+{iref('gna',i)})*{ar('DPO',i)}/365"
          for i in range(5)], B["AP"])
    frow("deferred", "递延收益", [f"={iref('rev_sub',i)}*{ar('DEF_RATE_TOC',i)}" for i in range(5)],
         B["deferred"], "订阅年费预收")
    frow("liab", "负债合计", [f"={sref('AP',i)}+{sref('IC_AP',i)}+{sref('deferred',i)}" for i in range(5)],
         [B["AP"][i]+B["IC_AP"][i]+B["deferred"][i] for i in range(5)], bold=True, indent=0)
    row = r
    frow("paid_in", "实收资本",
         [f"={ar('OPEN_PI_TOC',0)}+{ar('INJECT_TOC',0)}"] +
         [f"={chr(66+i)}{row}+{ar('INJECT_TOC',i)}" for i in range(1,5)], B["paid_in"],
         "母80%+战投20%")
    row = r
    frow("RE", "未分配利润",
         [f"={ar('OPEN_RE_TOC',0)}+{iref('ni',0)}"] +
         [f"={chr(66+i)}{row}+{iref('ni',i)}" for i in range(1,5)], B["RE"])
    frow("cash", "货币资金（倒挤）",
         [f"={sref('liab',i)}+{sref('paid_in',i)}+{sref('RE',i)}-{sref('AR',i)}" for i in range(5)],
         B["cash"], bold=True, indent=0, fill=TOTAL_FILL)

    c = ws.cell(row=r, column=1, value="── 现金流量表（间接法）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("cf_ni", "净利润", [f"={iref('ni',i)}" for i in range(5)], F["ni"])
    frow("cf_d_ar", "减：应收增加", [f"=-({sref('AR',i)}-({pref('AR',i)}))" for i in range(5)], F["d_ar"])
    frow("cf_d_ic", "内部往来净变动",
         [f"=({sref('IC_AP',i)}-({pref('IC_AP',i)}))" for i in range(5)], F["d_ic_wc"])
    frow("cf_d_def", "加：递延增加", [f"={sref('deferred',i)}-({pref('deferred',i)})" for i in range(5)], F["d_deferred"])
    frow("cf_d_ap", "加：应付增加", [f"={sref('AP',i)}-({pref('AP',i)})" for i in range(5)], F["d_ap"])
    frow("cfo", "经营现金流",
         [f"={sref('cf_ni',i)}+{sref('cf_d_ar',i)}+{sref('cf_d_ic',i)}+{sref('cf_d_def',i)}"
          f"+{sref('cf_d_ap',i)}" for i in range(5)], F["cfo"], bold=True)
    frow("cff", "筹资：股权注资（含战投）", [f"={ar('INJECT_TOC',i)}" for i in range(5)], F["cff"])
    frow("net", "现金净变动", [f"={sref('cfo',i)}+{sref('cff',i)}" for i in range(5)], F["net"], bold=True)
    op = A.OPENING["TOC"]["paid_in"] + A.OPENING["TOC"]["RE"]
    row = r
    frow("end_cash", "期末现金（CF）",
         [f"={op}+{sref('cfo',0)}+{sref('cff',0)}"] +
         [f"={chr(66+i)}{row}+{sref('cfo',i)}+{sref('cff',i)}" for i in range(1,5)],
         F["end_cash"], bold=True)
    line(ws, r, "★ CF=BS 差异（应全 0）", bold=True); r += 1
    L.reg("Toc_BS_CF", "chk_tie", r)
    line(ws, r, "  CF推导 − BS倒挤",
         formulas=[f"={sref('end_cash',i)}-{sref('cash',i)}" for i in range(5)], numFmt=NUM1); r += 1
    return ws

def gen(wb, L, eng):
    cloud_is(wb, L, eng); cloud_bscf(wb, L, eng)
    ind_is(wb, L, eng); ind_bscf(wb, L, eng)
    toc_is(wb, L, eng); toc_bscf(wb, L, eng)
