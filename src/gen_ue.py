"""UE 单位经济模型：上块=智擎云 MaaS（token/客户双视角），下块=智擎互联 C 端单用户。
公式全链接 Assumptions / Cloud_IS / Toc_IS；影子值 Python 直算，recompute 逐格核对。"""
import assumptions as A
from style import *
from layout import Layout

# 云 CAC 假设（万元/客户）：企业级销售团队逐年成熟，获客单价随之抬升（蓝字输入放本表）
CLOUD_CAC = [3.0, 3.5, 4.0, 4.5, 5.0]

def gen(wb, L, eng):
    ws = wb.create_sheet("UE")
    sheet_scaffold(ws, "单位经济模型 Unit Economics",
                   "上：智擎云 MaaS（token / 客户双视角）· 下：智擎互联 C 端单用户 | 单位见各行标签",
                   A.YEARS, "analysis")
    ws["A3"] = "科目（单位见行标签）"
    CIS = eng["E"]["CLOUD"]["IS"]; TIS = eng["E"]["TOC"]["IS"]
    r = 4
    def ar(k, i): return L.ref("Assumptions", k, i)
    def u(k, i): return L.ref("UE", k, i)
    def sec(t):
        nonlocal r
        c = ws.cell(row=r, column=1, value=t); c.font = BOLD; c.fill = SECTION_FILL
        for col in range(2, 8): ws.cell(row=r, column=col).fill = SECTION_FILL
        r += 1
    def frow(key, label, formulas, expected, note=None, bold=False, numFmt=NUM1):
        nonlocal r
        line(ws, r, label, formulas=formulas, note=note, bold=bold, indent=1, numFmt=numFmt)
        L.reg("UE", key, r, expected=expected); r += 1
    def vrow(key, label, vals, note=None, numFmt=NUM1):
        nonlocal r
        line(ws, r, label, values=vals, note=note, numFmt=numFmt)
        L.reg("UE", key, r, expected=vals); r += 1

    # ── 影子值：Python 直算（与公式同逻辑）───────────────
    P, U = A.CLOUD_PRICE, A.CLOUD_UCOST
    m_gp = [P[i] - U[i] for i in range(5)]                       # 单 token 毛利
    m_gm = [m_gp[i] / P[i] for i in range(5)]                    # 单 token 毛利率
    sc_p = [(P[0] - P[i]) / P[0] for i in range(5)]              # 价格累计降幅
    sc_c = [(U[0] - U[i]) / U[0] for i in range(5)]              # 成本累计降幅
    sc_d = [sc_c[i] - sc_p[i] for i in range(5)]                 # 剪刀差

    cust = A.CLOUD_CUSTOMERS
    arr = [CIS["rev_ext"][i] / cust[i] for i in range(5)]        # ARR/客户（万）
    gpc = [CIS["gp"][i] / cust[i] for i in range(5)]             # 毛利/客户（万）
    life = 1.0 / A.CLOUD_CHURN_M                                 # 平均生命周期（月）
    ltv = [gpc[i] / 12 * life for i in range(5)]                 # LTV = 月毛利×生命周期
    ltvcac = [ltv[i] / CLOUD_CAC[i] for i in range(5)]
    cac_pb = [CLOUD_CAC[i] / (gpc[i] / 12) for i in range(5)]    # CAC 回收期（月）

    rsub = [A.TOC_MAU_W[i] * A.TOC_PAYRATE[i] * A.TOC_ARPPU[i] * 12 for i in range(5)]
    rext = [rsub[i] + A.TOC_ADS[i] for i in range(5)]
    tcogs = [A.TOC_MAU_W[i] * A.TOC_UCOST[i] * 12 for i in range(5)]
    arpu = [rext[i] / A.TOC_MAU_W[i] / 12 for i in range(5)]     # 万元÷万人÷月 → 元
    tuc = [tcogs[i] / A.TOC_MAU_W[i] / 12 for i in range(5)]     # 元/人/月
    gpm = [arpu[i] - tuc[i] for i in range(5)]                   # 单用户月毛利
    t_pb = [A.TOC_CAC[i] / gpm[i] for i in range(5)]             # 回本周期（月）
    t_sub = [rsub[i] / rext[i] for i in range(5)]
    t_ads = [A.TOC_ADS[i] / rext[i] for i in range(5)]

    # ── 一、单 token 视角 ────────────────────────────────
    sec("── 一、智擎云 MaaS · 单 token 视角（每列=年份）──")
    frow("m_price", "收入单价（元/M tokens）", [f"={ar('CLOUD_PRICE',i)}" for i in range(5)],
         P, "Assumptions 混合计价", numFmt='0.00')
    frow("m_ucost", "推理单位成本（元/M tokens）", [f"={ar('CLOUD_UCOST',i)}" for i in range(5)],
         U, "自持+外租算力", numFmt='0.00')
    frow("m_gp", "单 token 毛利（元/M tokens）",
         [f"={u('m_price',i)}-{u('m_ucost',i)}" for i in range(5)], m_gp,
         "价−成本", bold=True, numFmt='0.00')
    frow("m_gm", "单 token 毛利率", [f"={u('m_gp',i)}/{u('m_price',i)}" for i in range(5)],
         m_gm, "毛利/单价", numFmt=PCT)
    sec("── 剪刀差：成本降幅 vs 价格降幅（累计 vs 2023A）──")
    frow("sc_p", "价格累计降幅", [f"=({u('m_price',0)}-{u('m_price',i)})/{u('m_price',0)}"
         for i in range(5)], sc_p, numFmt=PCT)
    frow("sc_c", "单位成本累计降幅", [f"=({u('m_ucost',0)}-{u('m_ucost',i)})/{u('m_ucost',0)}"
         for i in range(5)], sc_c, "成本降得比价格快", numFmt=PCT)
    frow("sc_d", "剪刀差 = 成本降幅 − 价格降幅", [f"={u('sc_c',i)}-{u('sc_p',i)}" for i in range(5)],
         sc_d, ">0 → 毛利率逐年扩张", bold=True, numFmt=PCT)
    r += 1

    # ── 二、单客户视角 ──────────────────────────────────
    sec("── 二、智擎云 MaaS · 单客户视角 ──")
    frow("c_n", "付费客户数（家）", [f"={ar('CLOUD_CUSTOMERS',i)}" for i in range(5)],
         A.CLOUD_CUSTOMERS, numFmt='#,##0')
    frow("c_arr", "ARR / 客户（万元/家）",
         [f"={L.ref('Cloud_IS','rev_ext',i)}/{ar('CLOUD_CUSTOMERS',i)}" for i in range(5)],
         arr, "对外收入÷客户数", numFmt='0.00')
    frow("c_gpc", "毛利 / 客户（万元/家）",
         [f"={L.ref('Cloud_IS','gp',i)}/{ar('CLOUD_CUSTOMERS',i)}" for i in range(5)],
         gpc, "毛利÷客户数", numFmt='0.00')
    frow("c_churn", "月流失率 churn（常量）", [f"={ar('CLOUD_CHURN_M',i)}" for i in range(5)],
         [A.CLOUD_CHURN_M] * 5, "1.2%/月", numFmt=PCT)
    frow("c_life", "平均生命周期（月）= 1/churn", [f"=1/{u('c_churn',i)}" for i in range(5)],
         [life] * 5, numFmt='0.0')
    frow("c_ltv", "LTV（万元/客户）= 月毛利 × 生命周期",
         [f"={u('c_gpc',i)}/12*{u('c_life',i)}" for i in range(5)], ltv,
         "月毛利=毛利/客户/12", bold=True, numFmt='0.00')
    vrow("c_cac", "CAC 获客成本（万元/客户，输入）", CLOUD_CAC, "蓝字输入·本表假设", numFmt='0.00')
    frow("c_ltvcac", "LTV / CAC（倍）", [f"={u('c_ltv',i)}/{u('c_cac',i)}" for i in range(5)],
         ltvcac, ">3x 为健康", bold=True, numFmt=RATIO)
    frow("c_pb", "CAC 回收期（月）= CAC / 月毛利",
         [f"={u('c_cac',i)}/({u('c_gpc',i)}/12)" for i in range(5)], cac_pb, numFmt='0.0')
    r += 1

    # ── 三、C 端单用户 ──────────────────────────────────
    sec("── 三、智擎互联 · C 端单用户 ──")
    frow("t_arpu", "ARPU（元/人/月）",
         [f"={L.ref('Toc_IS','rev_ext',i)}/{ar('TOC_MAU_W',i)}/12" for i in range(5)],
         arpu, "收入(万)÷MAU(万人)÷12", numFmt='0.00')
    frow("t_ucost", "推理+带宽成本（元/人/月）",
         [f"={L.ref('Toc_IS','cogs_ext',i)}/{ar('TOC_MAU_W',i)}/12" for i in range(5)],
         tuc, "含付云 T3", numFmt='0.00')
    frow("t_gpm", "单用户月毛利（元/人/月）",
         [f"={u('t_arpu',i)}-{u('t_ucost',i)}" for i in range(5)], gpm,
         "ARPU−单位成本", bold=True, numFmt='0.00')
    frow("t_cac", "获客 CAC（元/新增用户）", [f"={ar('TOC_CAC',i)}" for i in range(5)],
         A.TOC_CAC, numFmt='0.0')
    frow("t_pb", "回本周期（月）= CAC / 月毛利",
         [f"={u('t_cac',i)}/{u('t_gpm',i)}" for i in range(5)], t_pb, bold=True, numFmt='0.0')
    frow("t_sub", "订阅收入占比",
         [f"={L.ref('Toc_IS','rev_sub',i)}/{L.ref('Toc_IS','rev_ext',i)}" for i in range(5)],
         t_sub, numFmt=PCT)
    frow("t_ads", "广告收入占比",
         [f"={L.ref('Toc_IS','rev_ads',i)}/{L.ref('Toc_IS','rev_ext',i)}" for i in range(5)],
         t_ads, numFmt=PCT)
    return ws
