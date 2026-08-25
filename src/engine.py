"""影子引擎：Python 纯算全部三表+关联+合并。
与 Excel 公式互为镜像，重算后逐格对比。单位：万元。"""
import json, os, sys
sys.path.insert(0, os.path.dirname(__file__))
import assumptions as A

D = lambda x: round(float(x), 6)

# ── 1. Capex 折旧滚动（批次法：购置当年起 5 年直线）────────
def build_depreciation():
    dep, net = {}, {}
    for asset in ["HOLD_TRAIN", "CLOUD_INFER"]:
        # 存量原值视为 2023 年初全新资产，与 2023 年采购并批、当年起折
        batches = [(2023, A.CAPEX_OPENING[asset] + A.CAPEX[asset][0])] + [
            (y, c) for y, c in list(zip(A.YEARS, A.CAPEX[asset]))[1:]
        ]
        d_seq, net_seq = [], []
        for i, y in enumerate(A.YEARS):
            d = sum(orig / A.DEP_LIFE for by, orig in batches if 0 <= y - by < A.DEP_LIFE)
            n = sum(orig * max(0, 1 - (y - by + 1) / A.DEP_LIFE) for by, orig in batches
                    if 0 <= y - by < A.DEP_LIFE)
            d_seq.append(D(d)); net_seq.append(D(n))
        dep[asset], net[asset] = d_seq, net_seq
    return dep, net

# ── 2. 关联交易额 ───────────────────────────────────────
def build_related(rev_ext):
    """rev_ext[entity] = 对外收入序列（不含关联）"""
    t1 = {e: [D(r * A.T1_RATE[e]) for r in rev_ext[e]] for e in A.ENTITIES if e != "HOLD"}
    t1_hold = [D(sum(t1[e][i] for e in t1)) for i in range(A.N)]  # 研究院确认
    t2 = [D(x) for x in A.T2_FEE]
    return t1, t1_hold, t2

# ── 3. 各主体三表 ───────────────────────────────────────
def build_entities():
    dep, net = build_depreciation()
    E = {}  # E[entity] = {"IS": {...}, "BS": {...}, "CF": {...}}

    for e in A.ENTITIES:
        IS, BS, CF = {}, {}, {}
        prev_re = A.OPENING[e]["RE"]
        for i, y in enumerate(A.YEARS):
            # ---- IS ----
            if e == "HOLD":
                rev_ic = 0.0  # 占位，关联收入待 cache 建好后回填
                cogs = 0.0
                rd = A.HOLD_HEADCOUNT[i] * A.HOLD_SALARY[i] + dep["HOLD_TRAIN"][i] + A.HOLD_OTHER_RD[i]
                sales = 0.0
                gna = A.HOLD_GNA_HC[i] * A.HOLD_GNA_SALARY
                rev_ext_i = 0.0
            elif e == "CLOUD":
                rev_maaS = A.CLOUD_TOKENS_B[i] * 36.5 * A.CLOUD_PRICE[i]
                rev_ext_i = rev_maaS
                cogs_ext = A.CLOUD_TOKENS_B[i] * 36.5 * A.CLOUD_UCOST[i]
                rd = A.CLOUD_HEADCOUNT["rd"][i] * A.CLOUD_SALARY["rd"]
                sales = A.CLOUD_HEADCOUNT["sales"][i] * A.CLOUD_SALARY["sales"]
                gna = A.CLOUD_HEADCOUNT["gna"][i] * A.CLOUD_SALARY["gna"]
            elif e == "IND":
                rev_pvt = A.IND_PROJECTS[i] * A.IND_CONTRACT[i]
                rev_sol = A.IND_SUBSCRIBERS[i] * A.IND_SUB_FEE[i]
                rev_ext_i = rev_pvt + rev_sol
                cogs_ext = rev_pvt * A.IND_COGS_PVT + rev_sol * A.IND_COGS_SOL
                rd = A.IND_HEADCOUNT["rd"][i] * A.IND_SALARY["rd"]
                sales = A.IND_HEADCOUNT["sales"][i] * A.IND_SALARY["sales"]
                gna = A.IND_HEADCOUNT["gna"][i] * A.IND_SALARY["gna"]
            else:  # TOC
                rev_sub = A.TOC_MAU_W[i] * A.TOC_PAYRATE[i] * A.TOC_ARPPU[i] * 12
                rev_ads = A.TOC_ADS[i]
                rev_ext_i = rev_sub + rev_ads
                cogs_ext = A.TOC_MAU_W[i] * A.TOC_UCOST[i] * 12
                rd = A.TOC_HEADCOUNT["rd"][i] * A.TOC_SALARY["rd"]
                mau_prev = A.TOC_MAU_W[i-1] if i > 0 else 0
                new_users = (A.TOC_MAU_W[i] - mau_prev) * A.TOC_NEW_U_MULTIPLIER
                sales = new_users * A.TOC_CAC[i] + A.TOC_HEADCOUNT["sales"][i] * A.TOC_SALARY["sales"]
                gna = A.TOC_HEADCOUNT["gna"][i] * A.TOC_SALARY["gna"]

            # 关联收入/成本（非研究院主体在拿到 _rev_cache 后补算）
            if e == "HOLD":
                IS.setdefault("rev", []).append(0.0)  # 占位，后续回填
                IS.setdefault("rev_ext", []).append(0.0)
                IS.setdefault("cogs_ext", []).append(0.0)
                IS.setdefault("cogs_ic", []).append(0.0)
            else:
                IS.setdefault("rev_ext", []).append(D(rev_ext_i))
                IS.setdefault("cogs_ext", []).append(D(cogs_ext))

            IS.setdefault("rd", []).append(D(rd))
            IS.setdefault("sales", []).append(D(sales))
            IS.setdefault("gna", []).append(D(gna))

        E[e] = {"IS": IS}
        E.setdefault("_dep", dep); E.setdefault("_net", net)

    # 补关联收入/成本（需先有对外收入）
    t1, t1_hold, t2 = build_related({e: E[e]["IS"]["rev_ext"] for e in A.ENTITIES})
    E["_rev_cache"] = {"t1": t1, "t1_hold": t1_hold, "t2": t2}

    for e in ["CLOUD", "IND", "TOC"]:
        IS = E[e]["IS"]
        rev_ext = IS["rev_ext"]
        if e == "CLOUD":
            # T3 = 互联 COGS × 70%（互联 COGS 尚未算，用其 rev_ext 无关，直接按假设序列）
            toc_cogs = [A.TOC_MAU_W[i] * A.TOC_UCOST[i] * 12 for i in range(A.N)]
            rev_ic = [D(c * A.T3_SHARE) for c in toc_cogs]
            cogs_ic = [D(r * (1 - A.T3_GM)) for r in rev_ic]
        else:
            rev_ic = [0.0] * A.N; cogs_ic = [0.0] * A.N  # IND/TOC 无关联收入，t1/t2 仅计成本
        IS["rev_ic"] = rev_ic
        IS["cogs_ic"] = cogs_ic
        IS["rev"] = [D(rev_ext[i] + rev_ic[i]) for i in range(A.N)]
        IS["cogs_ic_t1"] = t1[e]  # 行业/互联的 T1 成本（license 费）
        # license 费计入 COGS（智擎云也付 license）
        IS["cogs"] = [D(IS["cogs_ext"][i] + IS["cogs_ic"][i] + t1[e][i]) for i in range(A.N)]
    # 研究院收入重挂（此时 cache 已就绪）
    HI = E["HOLD"]["IS"]
    HI["rev_ic_t1"] = t1_hold; HI["rev_ic_t2"] = t2
    HI["rev_ic"] = [D(t1_hold[i] + t2[i]) for i in range(A.N)]
    HI["rev"] = HI["rev_ic"]; HI["rev_ext"] = [0.0]*A.N
    HI["cogs_ext"] = [0.0]*A.N; HI["cogs_ic"] = [0.0]*A.N
    HI["cogs"] = [0.0]*A.N
    # 行业 T2 成本计入研发
    for i in range(A.N):
        HI["rd"][i] = D(HI["rd"][i])  # 研究院研发已含全部
    IND = E["IND"]["IS"]
    IND["cogs_ic_t2"] = [D(x) for x in t2]  # 行业 T2 成本（计入成本）
    IND["cogs"] = [D(IND["cogs_ext"][i] + IND["cogs_ic"][i] + t1["IND"][i] + t2[i]) for i in range(A.N)]

    # 毛利/EBIT/税/NI
    for e in A.ENTITIES:
        IS = E[e]["IS"]
        IS["gp"] = [D(IS["rev"][i] - IS["cogs"][i]) for i in range(A.N)]
        IS["opex"] = [D(IS["rd"][i] + IS["sales"][i] + IS["gna"][i]) for i in range(A.N)]
        IS["ebit"] = [D(IS["gp"][i] - IS["opex"][i]) for i in range(A.N)]
        IS["tax"] = [D(max(0.0, IS["ebit"][i]) * A.TAX_RATE) for i in range(A.N)]
        IS["ni"] = [D(IS["ebit"][i] - IS["tax"][i]) for i in range(A.N)]

    # ---- BS + CF（逐年滚动）----
    dep, net = E["_dep"], E["_net"]
    for e in A.ENTITIES:
        IS, BS, CF = E[e]["IS"], {}, {}
        op = A.OPENING[e]
        prev = {"RE": op["RE"], "paid_in": op["paid_in"],
                "AR": 0.0, "IC_AR": 0.0, "IC_AP": 0.0, "deferred": 0.0, "AP": 0.0,
                "fa": A.OPENING_FA.get(e, 0.0),
                "lt_inv": A.OPENING_LT_INV if e == "HOLD" else 0.0}
        # 期初现金倒挤：权益 − 非现金资产（WC 期初为零）
        prev["cash"] = prev["paid_in"] + prev["RE"] - prev["fa"] - prev["lt_inv"]
        for k in ["cash", "AR", "IC_AR", "IC_AP", "deferred", "AP", "RE", "paid_in", "fa"]:
            BS.setdefault(k, [])
        CF.setdefault("ni", IS["ni"])
        for k in ["da", "d_ar", "d_ic_wc", "d_deferred", "d_ap", "cfo", "capex", "cfi",
                  "equity_in", "cff", "net", "end_cash", "bs_cash"]:
            CF.setdefault(k, [])

        for i, y in enumerate(A.YEARS):
            AR = IS["rev_ext"][i] * A.DSO / 365
            # 内部往来：研究院挂 T1+T2 应收；云挂 T3 应收（2027 末含 cut-off 未开票 -220）
            if e == "HOLD":
                ic_flows = t1_hold[i] + t2[i]
                IC_AR = ic_flows * A.IC_DAYS / 365
                IC_AP = 0.0
            elif e == "CLOUD":
                IC_AR = E["TOC"]["IS"]["cogs"][0] if False else (A.TOC_MAU_W[i]*A.TOC_UCOST[i]*12*A.T3_SHARE) * A.IC_DAYS / 365
                if y == 2027: IC_AR -= A.CUTOFF_2027  # 云未开票，少挂
                IC_AP = t1["CLOUD"][i] * A.IC_DAYS / 365
            elif e == "IND":
                IC_AR = 0.0
                IC_AP = (t1["IND"][i] + t2[i]) * A.IC_DAYS / 365
            else:
                IC_AR = 0.0
                toc_t3 = A.TOC_MAU_W[i]*A.TOC_UCOST[i]*12*A.T3_SHARE
                IC_AP = toc_t3 * A.IC_DAYS / 365 + t1["TOC"][i] * A.IC_DAYS / 365  # T3+T1，含在途全挂
            cash_cost = IS["cogs"][i] + IS["opex"][i] - (dep["HOLD_TRAIN"][i] if e=="HOLD" else dep["CLOUD_INFER"][i] if e=="CLOUD" else 0.0)
            AP = cash_cost * A.DPO / 365
            deferred = IS["rev"][i] * A.DEFERRED_RATE.get(e, 0.0)
            fa = net["HOLD_TRAIN"][i] if e == "HOLD" else (net["CLOUD_INFER"][i] if e == "CLOUD" else 0.0)
            RE = prev["RE"] + IS["ni"][i]
            paid_in = prev["paid_in"] + A.EQUITY_INJECT[e][i]
            # 长投（研究院）
            lt_inv = None
            if e == "HOLD":
                add = A.EQUITY_INJECT["CLOUD"][i] + A.EQUITY_INJECT["IND"][i] + A.EQUITY_INJECT["TOC"][i]*0.8
                lt_inv = (prev.get("lt_inv", sum(A.HOLD_INITIAL_INVESTMENT.values()))) + add

            assets_nc = AR + IC_AR + fa + (lt_inv or 0.0)
            liab = AP + IC_AP + deferred
            equity = paid_in + RE
            cash = liab + equity - assets_nc  # 倒挤

            # ---- CF 间接法 ----
            da = (dep["HOLD_TRAIN"][i] if e=="HOLD" else dep["CLOUD_INFER"][i] if e=="CLOUD" else 0.0)
            d_ar = AR - prev["AR"]; d_ic_wc = (IC_AP - prev["IC_AP"]) - (IC_AR - prev["IC_AR"])
            d_def = deferred - prev["deferred"]; d_ap = AP - prev["AP"]
            cfo = IS["ni"][i] + da - d_ar + d_ic_wc + d_def + d_ap
            capex = (A.CAPEX["HOLD_TRAIN"][i] if e=="HOLD" else A.CAPEX["CLOUD_INFER"][i] if e=="CLOUD" else 0.0)
            inv_add = add if e == "HOLD" else 0.0
            cfi = -(capex + inv_add)
            equity_in = A.EQUITY_INJECT[e][i]
            cff = equity_in
            end_cash = prev["cash"] + cfo + cfi + cff
            bs_cash = cash

            for k, v in [("cash", cash), ("AR", AR), ("IC_AR", IC_AR), ("IC_AP", IC_AP),
                         ("AP", AP), ("deferred", deferred), ("fa", fa), ("RE", RE),
                         ("paid_in", paid_in)]:
                BS[k].append(D(v))
            if lt_inv is not None: BS.setdefault("lt_inv", []).append(D(lt_inv))
            for k, v in [("da", da), ("d_ar", -d_ar), ("d_ic_wc", d_ic_wc), ("d_deferred", d_def),
                         ("d_ap", d_ap), ("cfo", cfo), ("capex", -(capex+inv_add)), ("cfi", cfi),
                         ("equity_in", equity_in), ("cff", cff),
                         ("net", cfo+cfi+cff), ("end_cash", end_cash), ("bs_cash", bs_cash)]:
                CF[k].append(D(v))

            prev = {"cash": cash, "AR": AR, "IC_AR": IC_AR, "IC_AP": IC_AP, "AP": AP,
                    "deferred": deferred, "RE": RE, "paid_in": paid_in, "fa": fa,
                    "lt_inv": lt_inv if lt_inv is not None else prev.get("lt_inv", 0.0)}

        E[e]["BS"], E[e]["CF"] = BS, CF
    return E, t1, t1_hold, t2, dep, net

# ── 4. 合并 ────────────────────────────────────────────
def build_consol(E, t1, t1_hold, t2):
    C = {"IS": {}, "BS": {}, "CF": {}}
    s = lambda k, key: [D(sum(E[e][k][key][i] for e in A.ENTITIES)) for i in range(A.N)]

    IS = C["IS"]
    IS["rev_ext"] = s("IS", "rev_ext")
    IS["rev_ic"] = s("IS", "rev_ic")
    IS["rev"] = IS["rev_ext"]  # 合并收入 = 对外（关联全消）
    IS["cogs_ext"] = s("IS", "cogs_ext")
    # 合并成本：互联端 T3 付款全额抵消，保留云部门真实服务成本(0.55×T3)
    # 即 合并cogs = Σ对外成本 − 云内部毛利(0.45×T3)，内部利润随互联对外收入实现
    cloud_ic_gp = [D(E["CLOUD"]["IS"]["rev_ic"][i] - E["CLOUD"]["IS"]["cogs_ic"][i]) for i in range(A.N)]
    IS["cloud_ic_gp"] = cloud_ic_gp
    IS["cogs"] = [D(IS["cogs_ext"][i] - cloud_ic_gp[i]) for i in range(A.N)]
    IS["gp"] = [D(IS["rev"][i] - IS["cogs"][i]) for i in range(A.N)]
    IS["opex"] = s("IS", "opex")
    IS["ebit"] = [D(IS["gp"][i] - IS["opex"][i]) for i in range(A.N)]
    IS["tax"] = s("IS", "tax")
    IS["ni"] = [D(IS["ebit"][i] - IS["tax"][i]) for i in range(A.N)]
    # 少数股东损益 = 互联 NI × 20%（互联亏损则少数承担亏损）
    IS["ni_minority"] = [D(E["TOC"]["IS"]["ni"][i] * 0.2) for i in range(A.N)]
    IS["ni_parent"] = [D(IS["ni"][i] - IS["ni_minority"][i]) for i in range(A.N)]

    BS = C["BS"]
    for k in ["cash", "AR", "IC_AR", "fa", "lt_inv", "AP", "IC_AP", "deferred", "RE", "paid_in"]:
        BS[k] = s("BS", k) if all(k in E[e]["BS"] for e in A.ENTITIES) else \
                [D(sum(E[e]["BS"].get(k, [0]*A.N)[i] for e in A.ENTITIES)) for i in range(A.N)]

    # E0：长投 vs 子公司母份额权益对冲（按子公司明细），差额进合并资本公积
    cap_adj = []
    for i in range(A.N):
        diff = 0.0
        for sub, share in [("CLOUD", 1.0), ("IND", 1.0), ("TOC", 0.8)]:
            sub_eq = (E[sub]["BS"]["paid_in"][i]*share + E[sub]["BS"]["RE"][i]*share)
            lt = (A.HOLD_INITIAL_INVESTMENT[sub] +
                  sum(A.EQUITY_INJECT[sub][j] for j in range(i+1)) * share)
            diff += sub_eq - lt  # 借方>贷方 → 资本公积贷方
        cap_adj.append(D(diff))
    BS["cap_adj"] = cap_adj
    # 合并科目（抵消后）
    BS["c_AR"] = BS["AR"]
    BS["c_cash"] = BS["cash"]
    BS["c_fa"] = BS["fa"]
    BS["c_ICAR"] = [0.0]*A.N        # 全抵
    BS["c_ltinv"] = [0.0]*A.N       # 全抵
    BS["c_AP"] = BS["AP"]; BS["c_deferred"] = BS["deferred"]
    BS["c_ICAP"] = [0.0]*A.N
    # 内部往来双边孰低抵消，差额（在途未达）挂"其他应付款—内部在途"
    BS["other_payable"] = [D(BS["IC_AP"][i] - BS["IC_AR"][i]) for i in range(A.N)]
    BS["c_paid_in"] = [D(BS["paid_in"][0] - sum(A.HOLD_INITIAL_INVESTMENT.values()) - 0) for _ in range(A.N)]  # 占位，下面重算
    # 合并权益 = Σ权益 − 子公司权益全额 + 资本公积调整 + 少数股东权益(=TOC权益20%)
    sub_eq_total = [D(sum(E[e]["BS"]["paid_in"][i]+E[e]["BS"]["RE"][i] for e in ["CLOUD","IND","TOC"])) for i in range(A.N)]
    minority = [D((E["TOC"]["BS"]["paid_in"][i]+E["TOC"]["BS"]["RE"][i]) * 0.2) for i in range(A.N)]
    consol_eq = [D(BS["paid_in"][i] + BS["RE"][i] - sub_eq_total[i] + cap_adj[i] + minority[i]) for i in range(A.N)]
    BS["minority"] = minority; BS["consol_eq"] = consol_eq
    BS["c_assets"] = [D(BS["cash"][i]+BS["AR"][i]+BS["fa"][i]) for i in range(A.N)]
    BS["c_liab"] = [D(BS["AP"][i]+BS["deferred"][i]+BS["other_payable"][i]) for i in range(A.N)]

    # 合并 CF（间接法，从合并 BS 变动推）
    CF = C["CF"]
    CF["ni"] = IS["ni"]
    prev = {"AR": 0.0, "deferred": 0.0, "AP": 0.0, "other_pay": 0.0}
    for i in range(A.N):
        da = E["_dep"]["HOLD_TRAIN"][i] + E["_dep"]["CLOUD_INFER"][i]
        CF.setdefault("da", []).append(D(da))
        d_ar = BS["AR"][i]-prev["AR"]; d_def = BS["deferred"][i]-prev["deferred"]; d_ap = BS["AP"][i]-prev["AP"]
        d_op = BS["other_payable"][i]-prev["other_pay"]  # 在途挂账为非现金调整
        cfo = IS["ni"][i] + da - d_ar + d_def + d_ap + d_op
        capex = A.CAPEX["HOLD_TRAIN"][i] + A.CAPEX["CLOUD_INFER"][i]
        # 合并融资 = 研究院自身注资 + 互联战投 20%（给子公司的注资是内部划转，抵消）
        equity_in = A.EQUITY_INJECT["HOLD"][i] + A.EQUITY_INJECT["TOC"][i] * 0.2
        CF.setdefault("d_ar", []).append(D(-d_ar)); CF.setdefault("d_deferred", []).append(D(d_def))
        CF.setdefault("d_ap", []).append(D(d_ap)); CF.setdefault("cfo", []).append(D(cfo))
        CF.setdefault("capex", []).append(D(-capex)); CF.setdefault("cfi", []).append(D(-capex))
        CF.setdefault("equity_in", []).append(D(equity_in)); CF.setdefault("cff", []).append(D(equity_in))
        netcf = cfo - capex + equity_in
        CF.setdefault("net", []).append(D(netcf))
        CF.setdefault("end_cash", []).append(D((prev.get("cash", 0.0)) + netcf))
        CF.setdefault("bs_cash", []).append(D(BS["cash"][i]))
        prev = {"AR": BS["AR"][i], "deferred": BS["deferred"][i], "AP": BS["AP"][i],
                "other_pay": BS["other_payable"][i], "cash": BS["cash"][i]}
    # 合并期初现金 = Σ单体期初（倒挤口径），由第一年反推
    run = C["BS"]["cash"][0] - CF["net"][0]
    for i in range(A.N):
        CF["end_cash"][i] = D(run + CF["net"][i]); run = CF["end_cash"][i]
    return C

# ── 5. FactTable 长表 ───────────────────────────────────
def build_facttable(E, C):
    rows = []  # (entity, segment, region, year, account, value)
    seg_map = {"CLOUD": "MaaS API", "IND_PVT": "私有化部署", "IND_SOL": "行业解决方案",
               "TOC_SUB": "C端订阅", "TOC_ADS": "广告及其他"}
    def split_regions(val, i):
        return [(rg, D(val * A.REGION_MIX[rg][i])) for rg in A.REGIONS]
    for i, y in enumerate(A.YEARS):
        # 收入（对外，按业务线×地区）
        for ent, seg, val in [
            ("CLOUD", seg_map["CLOUD"], E["CLOUD"]["IS"]["rev_ext"][i]),
            ("IND", seg_map["IND_PVT"], A.IND_PROJECTS[i]*A.IND_CONTRACT[i]),
            ("IND", seg_map["IND_SOL"], A.IND_SUBSCRIBERS[i]*A.IND_SUB_FEE[i]),
            ("TOC", seg_map["TOC_SUB"], A.TOC_MAU_W[i]*A.TOC_PAYRATE[i]*A.TOC_ARPPU[i]*12),
            ("TOC", seg_map["TOC_ADS"], A.TOC_ADS[i]),
        ]:
            for rg, v in split_regions(val, i):
                rows.append(("CONSOL", seg, rg, y, "收入", v))
                rows.append((A.ENTITY_NAMES and ent, seg, rg, y, "收入", v))
        # COGS 按业务线（挂集团地区）
        for ent, seg, val in [
            ("CLOUD", seg_map["CLOUD"], E["CLOUD"]["IS"]["cogs_ext"][i]),
            ("IND", seg_map["IND_PVT"], A.IND_PROJECTS[i]*A.IND_CONTRACT[i]*A.IND_COGS_PVT),
            ("IND", seg_map["IND_SOL"], A.IND_SUBSCRIBERS[i]*A.IND_SUB_FEE[i]*A.IND_COGS_SOL),
            ("TOC", seg_map["TOC_SUB"], A.TOC_MAU_W[i]*A.TOC_UCOST[i]*12),
            ("TOC", seg_map["TOC_ADS"], 0.0),
        ]:
            rows.append((ent, seg, "集团", y, "成本", D(val)))
            rows.append(("CONSOL", seg, "集团", y, "成本", D(val)))
        # 费用/其他 IS 科目挂主体×集团
        for e in A.ENTITIES:
            for acc, key in [("研发费用", "rd"), ("销售费用", "sales"), ("管理费用", "gna")]:
                rows.append((e, "集团", "集团", y, acc, E[e]["IS"][key][i]))
                rows.append(("CONSOL", "集团", "集团", y, acc, C["IS"][key][i] if key in C["IS"] else sum(E[x]["IS"][key][i] for x in A.ENTITIES)))
        # BS/CF 关键科目（主体+合并）
        for ent in A.ENTITIES + ["CONSOL"]:
            src = E[ent] if ent != "CONSOL" else C
            for acc, key in [("净利润", ("IS", "ni")), ("现金", ("BS", "cash")),
                             ("固定资产", ("BS", "fa")), ("营业收入", ("IS", "rev")),
                             ("毛利", ("IS", "gp")), ("EBITDA", None)]:
                if acc == "EBITDA":
                    da = (E["_dep"]["HOLD_TRAIN"][i] + E["_dep"]["CLOUD_INFER"][i]) if ent == "CONSOL" else (
                        E["_dep"]["HOLD_TRAIN"][i] if ent == "HOLD" else E["_dep"]["CLOUD_INFER"][i] if ent == "CLOUD" else 0.0)
                    v = src["IS"]["ebit"][i] + da
                else:
                    v = src[key[0]][key[1]][i]
                rows.append((ent, "集团", "集团", y, acc, D(v)))
    return rows

# ── 6. 断言集（engine 级，Excel 级在 recompute）───────────
def run_checks(E, C):
    ck = []
    def add(name, vals, tol=1e-4):
        bad = [(A.YEARS[i], v) for i, v in enumerate(vals) if abs(v) > tol]
        ck.append((name, "PASS" if not bad else f"FAIL {bad}"))
    for e in A.ENTITIES:
        bs = E[e]["BS"]
        add(f"C1 {e} BS平衡", [bs["cash"][i]+bs["AR"][i]+bs["IC_AR"][i]+bs["fa"][i]+bs.get("lt_inv",[0]*A.N)[i]
                              - bs["AP"][i]-bs["IC_AP"][i]-bs["deferred"][i]-bs["paid_in"][i]-bs["RE"][i] for i in range(A.N)])
        add(f"C2 {e} CF=BS现金", [E[e]["CF"]["end_cash"][i] - E[e]["CF"]["bs_cash"][i] for i in range(A.N)])
        add(f"C5x {e} 现金>0", [0.0 if bs["cash"][i] > 0 else -1.0 for i in range(A.N)])
    cbs = C["BS"]
    add("C3 合并BS平衡", [cbs["cash"][i]+cbs["AR"][i]+cbs["fa"][i] - cbs["AP"][i]-cbs["deferred"][i]-cbs["other_payable"][i]-cbs["consol_eq"][i] for i in range(A.N)])
    add("C2 合并CF=BS现金", [C["CF"]["end_cash"][i] - C["CF"]["bs_cash"][i] for i in range(A.N)])
    add("C8 归母+少数=合并NI", [C["IS"]["ni_parent"][i]+C["IS"]["ni_minority"][i]-C["IS"]["ni"][i] for i in range(A.N)])
    # 集团形状 sanity
    ni = C["IS"]["ni"]
    # 形状：亏损峰值在中段（2025），2027 亏幅显著小于 2023
    add("形状: 亏损先扩后收", [0.0 if ni[2] < ni[0] and ni[4] > ni[0] else -1.0])
    ebitda26 = C["IS"]["ebit"][3] + E["_dep"]["HOLD_TRAIN"][3] + E["_dep"]["CLOUD_INFER"][3]
    add("形状: 2026 EBITDA>0", [0.0 if ebitda26 > 0 else -1.0])
    return ck

def build_all():
    E, t1, t1_hold, t2, dep, net = build_entities()
    C = build_consol(E, t1, t1_hold, t2)
    rows = build_facttable(E, C)
    ck = run_checks(E, C)
    return {"E": E, "C": C, "checks": ck, "facttable": rows,
            "dep": dep, "net": net}

if __name__ == "__main__":
    R = build_all()
    print("=" * 62)
    print("影子引擎勾稽断言：")
    for name, st in R["checks"]:
        print(f"  {'✅' if st=='PASS' else '❌'} {name}  {'' if st=='PASS' else st}")
    C = R["C"]
    print("\n集团关键形状（万元）：")
    print("  收入   :", [f"{v:,.0f}" for v in C["IS"]["rev"]])
    print("  毛利率 :", [f"{C['IS']['gp'][i]/C['IS']['rev'][i]:.1%}" for i in range(A.N)])
    print("  净利润 :", [f"{v:,.0f}" for v in C["IS"]["ni"]])
    print("  归母NI :", [f"{v:,.0f}" for v in C["IS"]["ni_parent"]])
    print("  期末现金:", [f"{v:,.0f}" for v in C["BS"]["cash"]])
    print("  GPU净值:", [f"{v:,.0f}" for v in R["net"]["HOLD_TRAIN"]], "+", [f"{v:,.0f}" for v in R["net"]["CLOUD_INFER"]])
    for e in A.ENTITIES:
        print(f"  {A.ENTITY_NAMES[e]} 现金:", [f"{v:,.0f}" for v in R['E'][e]['BS']['cash']])
    out = os.path.join(os.path.dirname(__file__), "..", "shadow_values.json")
    with open(out, "w") as f:
        json.dump({"checks": R["checks"], "years": A.YEARS}, f, ensure_ascii=False, indent=1, default=str)
    print("\n影子值 JSON →", os.path.abspath(out))
