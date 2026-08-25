"""关联层：IC_Register（交易登记）+ IC_Recon（双边对账矩阵+cut-off调节表）。
ic-recon 精神：差异可见、调节过程可见、调节后归零。"""
import assumptions as A
from style import *
from layout import Layout

def gen_register(wb, L, eng):
    ws = wb.create_sheet("IC_Register")
    sheet_scaffold(ws, "内部交易登记簿", "四类关联交易 · 金额公式全链接单体表 | 万元", A.YEARS, "ic")
    r = 4
    def ar(k, i): return L.ref("Assumptions", k, i)
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent)
        L.reg("IC_Register", key, r, expected=expected); r += 1

    c = ws.cell(row=r, column=1, value="── 当年交易额 ──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("t1_cloud", "T1 license · 研究院 → 智擎云",
         [f"={L.ref('Cloud_IS','rev_ext',i)}*{ar('T1_RATE_CLOUD',i)}" for i in range(5)],
         [A.T1_RATE["CLOUD"]*eng["E"]["CLOUD"]["IS"]["rev_ext"][i] for i in range(5)])
    frow("t1_ind", "T1 license · 研究院 → 智擎行业",
         [f"={L.ref('Ind_IS','rev_ext',i)}*{ar('T1_RATE_IND',i)}" for i in range(5)],
         [A.T1_RATE["IND"]*eng["E"]["IND"]["IS"]["rev_ext"][i] for i in range(5)])
    frow("t1_toc", "T1 license · 研究院 → 智擎互联",
         [f"={L.ref('Toc_IS','rev_ext',i)}*{ar('T1_RATE_TOC',i)}" for i in range(5)],
         [A.T1_RATE["TOC"]*eng["E"]["TOC"]["IS"]["rev_ext"][i] for i in range(5)])
    frow("t2", "T2 定制研发 · 研究院 → 智擎行业",
         [f"={ar('T2_FEE',i)}" for i in range(5)], [float(x) for x in A.T2_FEE])
    frow("t3", "T3 算力服务 · 智擎云 → 智擎互联",
         [f"={ar('TOC_MAU_W',i)}*{ar('TOC_UCOST',i)}*12*{ar('T3_SHARE',i)}" for i in range(5)],
         eng["E"]["CLOUD"]["IS"]["rev_ic"])
    frow("total", "关联交易合计",
         [f"=SUM({L.ref('IC_Register','t1_cloud',i)},{L.ref('IC_Register','t1_ind',i)},"
          f"{L.ref('IC_Register','t1_toc',i)},{L.ref('IC_Register','t2',i)},{L.ref('IC_Register','t3',i)})"
          for i in range(5)], None, bold=True)
    L.reg("IC_Register", "total", L.row_of[("IC_Register","total")],
          expected=[A.T1_RATE["CLOUD"]*eng["E"]["CLOUD"]["IS"]["rev_ext"][i]
                    + A.T1_RATE["IND"]*eng["E"]["IND"]["IS"]["rev_ext"][i]
                    + A.T1_RATE["TOC"]*eng["E"]["TOC"]["IS"]["rev_ext"][i]
                    + float(A.T2_FEE[i]) + eng["E"]["CLOUD"]["IS"]["rev_ic"][i] for i in range(5)])

    c = ws.cell(row=r, column=1, value="── 期末往来余额（账期60天挂账）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("bal_hold", "研究院 应收三家（借方）",
         [f"=({L.ref('IC_Register','t1_cloud',i)}+{L.ref('IC_Register','t1_ind',i)}"
          f"+{L.ref('IC_Register','t1_toc',i)}+{L.ref('IC_Register','t2',i)})*{ar('IC_DAYS',i)}/365"
          for i in range(5)], eng["E"]["HOLD"]["BS"]["IC_AR"])
    frow("bal_cloud_ar", "智擎云 应收互联（T3）",
         [f"={L.ref('IC_Register','t3',i)}*{ar('IC_DAYS',i)}/365"
          + ("-%s" % L.ref("Assumptions","CUTOFF_2027",i) if i == 4 else "") for i in range(5)],
         eng["E"]["CLOUD"]["BS"]["IC_AR"], "2027 未含未开票 220")
    frow("bal_cloud_ap", "智擎云 应付研究院（T1）",
         [f"={L.ref('IC_Register','t1_cloud',i)}*{ar('IC_DAYS',i)}/365" for i in range(5)],
         eng["E"]["CLOUD"]["BS"]["IC_AP"])
    frow("bal_ind", "智擎行业 应付研究院（T1+T2）",
         [f"=({L.ref('IC_Register','t1_ind',i)}+{L.ref('IC_Register','t2',i)})*{ar('IC_DAYS',i)}/365"
          for i in range(5)], eng["E"]["IND"]["BS"]["IC_AP"])
    frow("bal_toc", "智擎互联 应付（T3+T1，含在途）",
         [f"=({L.ref('IC_Register','t3',i)}+{L.ref('IC_Register','t1_toc',i)})*{ar('IC_DAYS',i)}/365"
          for i in range(5)], eng["E"]["TOC"]["BS"]["IC_AP"])
    frow("sum_ar", "内部应收合计",
         [f"={L.ref('IC_Register','bal_hold',i)}+{L.ref('IC_Register','bal_cloud_ar',i)}" for i in range(5)],
         [eng["E"]["HOLD"]["BS"]["IC_AR"][i]+eng["E"]["CLOUD"]["BS"]["IC_AR"][i] for i in range(5)], bold=True)
    frow("sum_ap", "内部应付合计",
         [f"={L.ref('IC_Register','bal_cloud_ap',i)}+{L.ref('IC_Register','bal_ind',i)}"
          f"+{L.ref('IC_Register','bal_toc',i)}" for i in range(5)],
         [eng["E"]["CLOUD"]["BS"]["IC_AP"][i]+eng["E"]["IND"]["BS"]["IC_AP"][i]
          +eng["E"]["TOC"]["BS"]["IC_AP"][i] for i in range(5)], bold=True)
    return ws

def gen_recon(wb, L, eng):
    ws = wb.create_sheet("IC_Recon")
    sheet_scaffold(ws, "关联方对账矩阵 · 双边核对 + 在途调节", "差异可见 → 调节 → 归零 | 万元", A.YEARS, "ic")
    r = 4
    def iref(k, i): return L.ref("IC_Register", k, i)
    def frow(key, label, formulas, expected, note=None, bold=False, indent=1, fill=None):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=indent, fill=fill)
        L.reg("IC_Recon", key, r, expected=expected); r += 1

    c = ws.cell(row=r, column=1, value="── 双边核对（A 账上应收 B  vs  B 账上应付 A）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("p1_ar", "研究院 ↔ 智擎云：研究院应收", [f"={iref('t1_cloud',i)}*{L.ref('Assumptions','IC_DAYS',i)}/365" for i in range(5)],
         [A.T1_RATE["CLOUD"]*eng["E"]["CLOUD"]["IS"]["rev_ext"][i]*A.IC_DAYS/365 for i in range(5)])
    frow("p1_ap", "研究院 ↔ 智擎云：云应付", [f"={iref('bal_cloud_ap',i)}" for i in range(5)],
         eng["E"]["CLOUD"]["BS"]["IC_AP"])
    frow("p1_diff", "差异", [f"={L.ref('IC_Recon','p1_ar',i)}-{L.ref('IC_Recon','p1_ap',i)}" for i in range(5)],
         [0.0]*5, bold=True)
    r += 1
    frow("p2_ar", "研究院 ↔ 智擎行业：研究院应收",
         [f"=({iref('t1_ind',i)}+{iref('t2',i)})*{L.ref('Assumptions','IC_DAYS',i)}/365" for i in range(5)],
         [(A.T1_RATE["IND"]*eng["E"]["IND"]["IS"]["rev_ext"][i]+float(A.T2_FEE[i]))*A.IC_DAYS/365 for i in range(5)])
    frow("p2_ap", "研究院 ↔ 智擎行业：行业应付", [f"={iref('bal_ind',i)}" for i in range(5)],
         eng["E"]["IND"]["BS"]["IC_AP"])
    frow("p2_diff", "差异", [f"={L.ref('IC_Recon','p2_ar',i)}-{L.ref('IC_Recon','p2_ap',i)}" for i in range(5)],
         [0.0]*5, bold=True)
    r += 1
    frow("p3_ar", "研究院 ↔ 智擎互联：研究院应收", [f"={iref('t1_toc',i)}*{L.ref('Assumptions','IC_DAYS',i)}/365" for i in range(5)],
         [A.T1_RATE["TOC"]*eng["E"]["TOC"]["IS"]["rev_ext"][i]*A.IC_DAYS/365 for i in range(5)])
    frow("p3_ap", "研究院 ↔ 智擎互联：互联应付（T1部分）",
         [f"={iref('t1_toc',i)}*{L.ref('Assumptions','IC_DAYS',i)}/365" for i in range(5)],
         [A.T1_RATE["TOC"]*eng["E"]["TOC"]["IS"]["rev_ext"][i]*A.IC_DAYS/365 for i in range(5)])
    frow("p3_diff", "差异", [f"={L.ref('IC_Recon','p3_ar',i)}-{L.ref('IC_Recon','p3_ap',i)}" for i in range(5)],
         [0.0]*5, bold=True)
    r += 1
    frow("p4_ar", "智擎云 ↔ 智擎互联：云应收（未开票−220）", [f"={iref('bal_cloud_ar',i)}" for i in range(5)],
         eng["E"]["CLOUD"]["BS"]["IC_AR"])
    frow("p4_ap", "智擎云 ↔ 智擎互联：互联应付（全额挂账）",
         [f"={iref('t3',i)}*{L.ref('Assumptions','IC_DAYS',i)}/365" for i in range(5)],
         [A.TOC_MAU_W[i]*A.TOC_UCOST[i]*12*A.T3_SHARE*A.IC_DAYS/365 for i in range(5)])
    frow("p4_diff", "差异（2027 = 在途 220）",
         [f"={L.ref('IC_Recon','p4_ar',i)}-{L.ref('IC_Recon','p4_ap',i)}" for i in range(5)],
         [0.0, 0.0, 0.0, 0.0, -A.CUTOFF_2027], bold=True, fill=WARN_FILL)

    c = ws.cell(row=r, column=1, value="── 在途调节表（cut-off：互联12月已确认，云次年1月开票）──"); c.font = BOLD; c.fill = SECTION_FILL; r += 1
    frow("adj1", "调节前差异", [f"={L.ref('IC_Recon','p4_diff',i)}" for i in range(5)],
         [0.0, 0.0, 0.0, 0.0, -A.CUTOFF_2027])
    frow("adj2", "加：云未开票在途服务", [0.0, 0.0, 0.0, 0.0, f"={L.ref('Assumptions','CUTOFF_2027',4)}"],
         [0.0, 0.0, 0.0, 0.0, float(A.CUTOFF_2027)], "服务已提供、发票在途")
    frow("adj3", "调节后差异（应全 0）",
         [f"={L.ref('IC_Recon','adj1',i)}+{L.ref('IC_Recon','adj2',i)}" for i in range(5)],
         [0.0]*5, bold=True, fill=TOTAL_FILL)
    r += 1
    frow("elim_base", "合并抵消口径（双边孰低=调节后应收方）",
         [f"={L.ref('IC_Recon','p4_ar',i)}+{L.ref('IC_Recon','adj2',i)}" for i in range(5)],
         [A.TOC_MAU_W[i]*A.TOC_UCOST[i]*12*A.T3_SHARE*A.IC_DAYS/365 for i in range(5)],
         "Eliminations E2 引用此行", bold=True)
    return ws

def gen(wb, L, eng):
    gen_register(wb, L, eng)
    gen_recon(wb, L, eng)
