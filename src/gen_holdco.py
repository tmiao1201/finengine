"""研究院（HoldCo）全量三表生成器 —— 所有单体 sheet 的样板。
公式与影子引擎严格镜像；sheet 底部自带勾稽差异行（证据在表里）。"""
import assumptions as A
from style import *
from layout import Layout

def a(key, i=0):
    return f"Assumptions!{chr(67+i)}{key if isinstance(key,int) else None}"
# 注：统一走 L.ref，此函数仅占位说明意图

def _is(wb, L, eng):
    ws = wb.create_sheet("HoldCo_IS")
    sheet_scaffold(ws, "智擎研究院 · 利润表", "单体 | 全关联收入，无对外收入 | 万元", A.YEARS, "engine")
    E = eng["E"]["HOLD"]["IS"]
    r = 4
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent)
        L.reg("HoldCo_IS", key, r, expected=expected); r += 1
    def aref(key, i):  # Assumptions 引用
        return L.ref("Assumptions", key, i)

    r = 4
    frow("rev_ic_t1", "关联收入 · T1 模型 license",
         [f"={L.ref('Cloud_IS','rev_ext',i)}*{aref('T1_RATE_CLOUD',i)}"
          f"+{L.ref('Ind_IS','rev_ext',i)}*{aref('T1_RATE_IND',i)}"
          f"+{L.ref('Toc_IS','rev_ext',i)}*{aref('T1_RATE_TOC',i)}" for i in range(5)],
         E["rev_ic_t1"], "三家子公司对外收入×费率")
    frow("rev_ic_t2", "关联收入 · T2 定制研发服务",
         [f"={aref('T2_FEE',i)}" for i in range(5)], E["rev_ic_t2"], "行业子公司")
    frow("rev", "营业收入合计",
         [f"={L.ref('HoldCo_IS','rev_ic_t1',i)}+{L.ref('HoldCo_IS','rev_ic_t2',i)}" for i in range(5)],
         E["rev"], bold=True)
    frow("cogs", "营业成本", [0]*5, E["cogs"], "研究院无直接 COGS")
    frow("gp", "毛利润",
         [f"={L.ref('HoldCo_IS','rev',i)}-{L.ref('HoldCo_IS','cogs',i)}" for i in range(5)],
         E["gp"], bold=True)
    frow("rd", "研发费用",
         [f"={aref('HOLD_HC',i)}*{aref('HOLD_SALARY',i)}+{L.ref('Capex_Dep','DEP_TRAIN',i)}"
          f"+{aref('HOLD_OTHER_RD',i)}" for i in range(5)], E["rd"], "编制+训练折旧+其他")
    frow("sales", "销售费用", [0]*5, E["sales"])
    frow("gna", "管理费用",
         [f"={aref('HOLD_GNA_HC',i)}*{aref('HOLD_GNA_SALARY',i)}" for i in range(5)], E["gna"])
    frow("ebit", "EBIT",
         [f"={L.ref('HoldCo_IS','gp',i)}-{L.ref('HoldCo_IS','rd',i)}-{L.ref('HoldCo_IS','sales',i)}"
          f"-{L.ref('HoldCo_IS','gna',i)}" for i in range(5)], E["ebit"], bold=True)
    frow("tax", "所得税（亏损不确认）",
         [f"=MAX(0,{L.ref('HoldCo_IS','ebit',i)})*{aref('TAX_RATE',i)}" for i in range(5)], E["tax"])
    frow("ni", "净利润",
         [f"={L.ref('HoldCo_IS','ebit',i)}-{L.ref('HoldCo_IS','tax',i)}" for i in range(5)],
         E["ni"], bold=True)
    return ws

def _bs(wb, L, eng):
    ws = wb.create_sheet("HoldCo_BS")
    sheet_scaffold(ws, "智擎研究院 · 资产负债表", "单体 | 现金=倒挤项 → 结构性平衡 | 万元", A.YEARS, "engine")
    B = eng["E"]["HOLD"]["BS"]; CFd = eng["E"]["HOLD"]["CF"]
    r = 4
    def brow(key, label, formulas, expected, note=None, bold=False, indent=1, fill=None):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent, fill=fill)
        L.reg("HoldCo_BS", key, r, expected=expected); r += 1
    def aref(k, i): return L.ref("Assumptions", k, i)
    def iref(k, i): return L.ref("HoldCo_IS", k, i)

    r = 4
    brow("AR", "应收账款", [0]*5, B["AR"], "无对外收入，应收为 0")
    brow("IC_AR", "内部应收（关联往来）",
         [f"=({iref('rev_ic_t1',i)}+{iref('rev_ic_t2',i)})*{aref('IC_DAYS',i)}/365" for i in range(5)],
         B["IC_AR"], "T1+T2 × 60天")
    brow("fa", "固定资产 · GPU 净值", [f"={L.ref('Capex_Dep','NET_TRAIN',i)}" for i in range(5)],
         B["fa"], "训练算力")
    row = r
    brow("lt_inv", "长期股权投资",
         [f"={aref('OPEN_LT_INV',0)}+({aref('INJECT_CLOUD',0)}+{aref('INJECT_IND',0)}"
         f"+{aref('INJECT_TOC',0)}*0.8)"] +
         [f"={chr(66+i)}{row}+({aref('INJECT_CLOUD',i)}+{aref('INJECT_IND',i)}"
          f"+{aref('INJECT_TOC',i)}*0.8)" for i in range(1,5)],
         B["lt_inv"], "成本法：期初+对子公司注资")
    # 非现金资产合计（辅助行）
    L.reg("HoldCo_BS", "assets_nc", r,
          expected=[B["AR"][i]+B["IC_AR"][i]+B["fa"][i]+B["lt_inv"][i] for i in range(5)])
    line(ws, r, "非现金资产合计",
         formulas=[f"={L.ref('HoldCo_BS','AR',i)}+{L.ref('HoldCo_BS','IC_AR',i)}"
                   f"+{L.ref('HoldCo_BS','fa',i)}+{L.ref('HoldCo_BS','lt_inv',i)}" for i in range(5)],
         bold=True, note="（辅助行）"); r += 1
    r += 1
    brow("AP", "应付账款",
         [f"=({iref('cogs',i)}+{iref('rd',i)}+{iref('sales',i)}+{iref('gna',i)}"
          f"-{L.ref('Capex_Dep','DEP_TRAIN',i)})*{aref('DPO',i)}/365" for i in range(5)],
         B["AP"], "现金成本×DPO")
    brow("IC_AP", "内部应付", [0]*5, B["IC_AP"])
    brow("deferred", "递延收益", [0]*5, B["deferred"])
    brow("liab", "负债合计",
         [f"={L.ref('HoldCo_BS','AP',i)}+{L.ref('HoldCo_BS','IC_AP',i)}+{L.ref('HoldCo_BS','deferred',i)}"
          for i in range(5)], [B["AP"][i]+B["IC_AP"][i]+B["deferred"][i] for i in range(5)],
         bold=True, indent=0)
    row = r
    brow("paid_in", "实收资本",
         [f"={aref('OPEN_PI_HOLD',0)}+{aref('INJECT_HOLD',0)}"] +
         [f"={chr(66+i)}{row}+{aref('INJECT_HOLD',i)}" for i in range(1,5)],
         B["paid_in"])
    row = r
    brow("RE", "未分配利润",
         [f"={aref('OPEN_RE_HOLD',0)}+{iref('ni',0)}"] +
         [f"={chr(66+i)}{row}+{iref('ni',i)}" for i in range(1,5)],
         B["RE"], "期初+NI，无分红")
    brow("equity", "所有者权益合计",
         [f"={L.ref('HoldCo_BS','paid_in',i)}+{L.ref('HoldCo_BS','RE',i)}" for i in range(5)],
         [B["paid_in"][i]+B["RE"][i] for i in range(5)], bold=True, indent=0)
    brow("L_E", "负债+权益合计",
         [f"={L.ref('HoldCo_BS','liab',i)}+{L.ref('HoldCo_BS','equity',i)}" for i in range(5)],
         [B["AP"][i]+B["IC_AP"][i]+B["deferred"][i]+B["paid_in"][i]+B["RE"][i] for i in range(5)],
         bold=True, indent=0)
    brow("cash", "货币资金（倒挤）",
         [f"={L.ref('HoldCo_BS','L_E',i)}-{L.ref('HoldCo_BS','assets_nc',i)}" for i in range(5)],
         B["cash"], "负债+权益−非现金资产", bold=True, indent=0, fill=TOTAL_FILL)
    brow("assets", "资产总计",
         [f"={L.ref('HoldCo_BS','cash',i)}+{L.ref('HoldCo_BS','assets_nc',i)}" for i in range(5)],
         [B["cash"][i]+B["AR"][i]+B["IC_AR"][i]+B["fa"][i]+B["lt_inv"][i] for i in range(5)],
         bold=True, indent=0)
    # 勾稽自检行
    line(ws, r, "★ BS 平衡差（应全 0）", bold=True); r += 1
    L.reg("HoldCo_BS", "chk_balance", r)
    line(ws, r, "  资产 − 负债 − 权益",
         formulas=[f"={L.ref('HoldCo_BS','assets',i)}-{L.ref('HoldCo_BS','L_E',i)}" for i in range(5)],
         numFmt=NUM1); chk_row = r; r += 1
    return ws

def _cf(wb, L, eng):
    ws = wb.create_sheet("HoldCo_CF")
    sheet_scaffold(ws, "智擎研究院 · 现金流量表（间接法）", "单体 | 独立推导期末现金，与 BS 倒挤现金互为勾稽 | 万元",
                   A.YEARS, "engine")
    F = eng["E"]["HOLD"]["CF"]
    r = 4
    def crow(key, label, formulas, expected, note=None, bold=False, indent=1):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent)
        L.reg("HoldCo_CF", key, r, expected=expected); r += 1
    def aref(k, i): return L.ref("Assumptions", k, i)
    def iref(k, i): return L.ref("HoldCo_IS", k, i)
    def bref(k, i): return L.ref("HoldCo_BS", k, i)
    def pref(k, i):  # BS 科目的上一年值（i=0 用期初 0）
        return bref(k, i-1) if i > 0 else "0"

    crow("ni", "净利润", [f"={iref('ni',i)}" for i in range(5)], F["ni"], bold=False)
    crow("da", "加：折旧摊销", [f"={L.ref('Capex_Dep','DEP_TRAIN',i)}" for i in range(5)], F["da"])
    crow("d_ar", "减：应收账款增加", [f"=-({bref('AR',i)}-({pref('AR',i)}))" for i in range(5)], F["d_ar"])
    crow("d_ic_wc", "内部往来净变动", [f"=({bref('IC_AP',i)}-({pref('IC_AP',i)}))-({bref('IC_AR',i)}-({pref('IC_AR',i)}))"
                                     for i in range(5)], F["d_ic_wc"], "Δ应付−Δ应收")
    crow("d_def", "加：递延收益增加", [f"={bref('deferred',i)}-({pref('deferred',i)})" for i in range(5)], F["d_deferred"])
    crow("d_ap", "加：应付账款增加", [f"={bref('AP',i)}-({pref('AP',i)})" for i in range(5)], F["d_ap"])
    crow("cfo", "经营活动现金流净额",
         [f"={L.ref('HoldCo_CF','ni',i)}+{L.ref('HoldCo_CF','da',i)}+{L.ref('HoldCo_CF','d_ar',i)}"
          f"+{L.ref('HoldCo_CF','d_ic_wc',i)}+{L.ref('HoldCo_CF','d_def',i)}+{L.ref('HoldCo_CF','d_ap',i)}"
          for i in range(5)], F["cfo"], bold=True)
    crow("capex", "购建固定资产",
         [f"=-{L.ref('Capex_Dep','CAPEX_TRAIN',i)}" for i in range(5)],
         [-x for x in A.CAPEX["HOLD_TRAIN"]])
    inv_exp = [-(A.EQUITY_INJECT["CLOUD"][i]+A.EQUITY_INJECT["IND"][i]+A.EQUITY_INJECT["TOC"][i]*0.8)
               for i in range(5)]
    crow("inv_add", "对子公司增资",
         [f"=-({aref('INJECT_CLOUD',i)}+{aref('INJECT_IND',i)}+{aref('INJECT_TOC',i)}*0.8)" for i in range(5)],
         inv_exp, note="长投增加")
    crow("cfi", "投资活动现金流净额",
         [f"={L.ref('HoldCo_CF','capex',i)}+{L.ref('HoldCo_CF','inv_add',i)}" for i in range(5)],
         F["cfi"], bold=True)
    crow("cff", "筹资：股权注资", [f"={aref('INJECT_HOLD',i)}" for i in range(5)], F["cff"])
    crow("net", "现金净变动",
         [f"={L.ref('HoldCo_CF','cfo',i)}+{L.ref('HoldCo_CF','cfi',i)}+{L.ref('HoldCo_CF','cff',i)}"
          for i in range(5)], F["net"], bold=True)
    opening_cash = A.OPENING["HOLD"]["paid_in"] + A.OPENING["HOLD"]["RE"] - A.OPENING_FA["HOLD"] - A.OPENING_LT_INV
    row = r
    crow("end_cash", "期末现金（CF 推导）",
         [f"={opening_cash}+{L.ref('HoldCo_CF','net',0)}"] +
         [f"={chr(66+i)}{row}+{L.ref('HoldCo_CF','net',i)}" for i in range(1,5)],
         F["end_cash"], bold=True)
    crow("bs_cash", "期末现金（BS 倒挤）", [f"={bref('cash',i)}" for i in range(5)], F["bs_cash"])
    line(ws, r, "★ CF=BS 现金差异（应全 0）", bold=True); r += 1
    L.reg("HoldCo_CF", "chk_tie", r)
    line(ws, r, "  CF推导 − BS倒挤",
         formulas=[f"={L.ref('HoldCo_CF','end_cash',i)}-{L.ref('HoldCo_CF','bs_cash',i)}" for i in range(5)],
         numFmt=NUM1); r += 1
    return ws

def gen(wb, L, eng):
    _bs_pre = None
    _is(wb, L, eng); _bs(wb, L, eng); _cf(wb, L, eng)
