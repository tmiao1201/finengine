"""影子引擎：连锁咖啡单主体三表。行业无关骨架 + 咖啡映射。
只认 meta.get 接口（数据与逻辑分离）；倒挤现金法 + CF 间接法勾稽断言。万元。"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
import meta
from industry_config import DEP_LIFE, SEGMENTS, REGIONS

D = lambda x: round(float(x), 6)
YEARS = meta.YEARS; N = meta.N

def build_depreciation():
    """门店批次折旧：期初存量并入首年批，净增门店×单店投入为各年批。"""
    capex = []
    for i in range(N):
        stores_prev = meta.get("stores", i-1) if i > 0 else (meta.get("open_store_fa", 0) /
                                                              meta.get("capex_per_store", 0))
        net_open = meta.get("stores", i) - stores_prev
        capex.append(net_open * meta.get("capex_per_store", i))
    batches = [(YEARS[0], meta.get("open_store_fa", 0) + capex[0])] + \
              [(YEARS[i], capex[i]) for i in range(1, N)]
    dep, net = [], []
    for i, y in enumerate(YEARS):
        d = sum(o / DEP_LIFE for by, o in batches if 0 <= y - by < DEP_LIFE)
        n = sum(o * max(0, 1 - (y - by + 1) / DEP_LIFE) for by, o in batches
                if 0 <= y - by < DEP_LIFE)
        dep.append(D(d)); net.append(D(n))
    return dep, net, capex

def build():
    dep, fa_net, capex = build_depreciation()
    IS, BS, CF = {}, {}, {}
    prev = {"cash": None, "AR": 0.0, "AP": 0.0, "deferred": 0.0,
            "RE": meta.get("open_re", 0), "paid_in": meta.get("open_pi", 0),
            "fa": meta.get("open_store_fa", 0)}
    prev["cash"] = prev["paid_in"] + prev["RE"] - prev["fa"]
    for i in range(N):
        # ---- IS（量×价）----
        bev = meta.get("stores", i) * meta.get("daily_cups", i) * meta.get("avg_ticket", i) * 365 / 10000
        ret = bev * meta.get("retail_ratio", i)
        rev = bev + ret
        cogs_mat = rev * meta.get("cogs_rate", i)
        cogs_store = meta.get("stores", i) * meta.get("store_opex_m", i) * 12
        cogs = cogs_mat + cogs_store + dep[i]
        gp = rev - cogs
        mkt = rev * meta.get("mkt_rate", i)
        rd = meta.get("headcount_rd", i) * meta.get("salary_rd", i)
        sales = meta.get("headcount_sales", i) * meta.get("salary_sales", i)
        gna = meta.get("headcount_gna", i) * meta.get("salary_gna", i)
        ebit = gp - mkt - rd - sales - gna
        tax = max(0.0, ebit) * meta.get("tax_rate", i)
        ni = ebit - tax
        # ---- BS（倒挤现金）----
        AR = rev * meta.get("dso_days", i) / 365
        AP = (cogs + mkt + rd + sales + gna - dep[i]) * meta.get("dpo_days", i) / 365
        deferred = rev * meta.get("deferred_rate", i)
        RE = prev["RE"] + ni
        paid_in = prev["paid_in"] + meta.get("equity_inject", i)
        cash = AP + deferred + paid_in + RE - AR - fa_net[i]
        # ---- CF（间接法）----
        d_ar = AR - prev["AR"]; d_ap = AP - prev["AP"]; d_def = deferred - prev["deferred"]
        cfo = ni + dep[i] - d_ar + d_ap + d_def
        cfi = -capex[i]; cff = meta.get("equity_inject", i)
        end_cash = prev["cash"] + cfo + cfi + cff

        for k, v in [("rev_bev", bev), ("rev_ret", ret), ("rev", rev),
                     ("cogs_mat", cogs_mat), ("cogs_store", cogs_store), ("dep", dep[i]),
                     ("cogs", cogs), ("gp", gp), ("mkt", mkt), ("rd", rd), ("sales", sales),
                     ("gna", gna), ("ebit", ebit), ("tax", tax), ("ni", ni)]:
            IS.setdefault(k, []).append(D(v))
        for k, v in [("AR", AR), ("fa", fa_net[i]), ("AP", AP), ("deferred", deferred),
                     ("RE", RE), ("paid_in", paid_in), ("cash", cash)]:
            BS.setdefault(k, []).append(D(v))
        for k, v in [("cfo", cfo), ("capex", -capex[i]), ("cfi", cfi), ("cff", cff),
                     ("net", cfo + cfi + cff), ("end_cash", end_cash), ("bs_cash", cash)]:
            CF.setdefault(k, []).append(D(v))
        prev = {"cash": cash, "AR": AR, "AP": AP, "deferred": deferred,
                "RE": RE, "paid_in": paid_in, "fa": fa_net[i]}

    # FactTable（业务线×地区 × IS 科目）
    ft = []
    for i in range(N):
        bev = IS["rev_bev"][i]; ret = IS["rev_ret"][i]
        for seg, val, cost in [("现制饮品", bev, IS["cogs"][i] * bev / (bev + ret)),
                               ("零售产品", ret, IS["cogs"][i] * ret / (bev + ret))]:
            for rg in REGIONS:
                mix = meta.get("region_mix", i, region=rg)
                ft.append(("GROUP", seg, rg, YEARS[i], "收入", D(val * mix)))
            ft.append(("GROUP", seg, "集团", YEARS[i], "成本", D(cost)))
        ft.append(("GROUP", "集团", "集团", YEARS[i], "毛利", IS["gp"][i]))
        ft.append(("GROUP", "集团", "集团", YEARS[i], "EBIT", IS["ebit"][i]))
        ft.append(("GROUP", "集团", "集团", YEARS[i], "净利润", IS["ni"][i]))
    return {"IS": IS, "BS": BS, "CF": CF, "FT": ft, "dep": dep, "fa_net": fa_net,
            "capex": capex}

def checks(R):
    IS, BS, CF = R["IS"], R["BS"], R["CF"]
    ck = []
    def add(name, vals, tol=1e-4):
        bad = [(YEARS[i], v) for i, v in enumerate(vals) if abs(v) > tol]
        ck.append((name, "PASS" if not bad else f"FAIL {bad}"))
    add("C1 BS平衡", [BS["cash"][i]+BS["AR"][i]+BS["fa"][i]-BS["AP"][i]-BS["deferred"][i]
                      -BS["paid_in"][i]-BS["RE"][i] for i in range(N)])
    add("C2 CF=BS现金", [CF["end_cash"][i]-CF["bs_cash"][i] for i in range(N)])
    add("C3 现金>0", [0.0 if BS["cash"][i] > 0 else -1.0 for i in range(N)])
    add("C4 RE滚动", [0.0 if i == 0 and abs(BS["RE"][0]-meta.get("open_re",0)-IS["ni"][0]) < 1e-6
                      else (0.0 if i > 0 and abs(BS["RE"][i]-BS["RE"][i-1]-IS["ni"][i]) < 1e-6 else -1.0)
                      for i in range(N)])
    add("形状: 盈利扩张", [0.0 if IS["ni"][0] > 0 and IS["ni"][4] > IS["ni"][0] else -1.0])
    return ck

if __name__ == "__main__":
    R = build()
    print("=" * 60)
    for name, st in checks(R):
        print(f"  {'✅' if st=='PASS' else '❌'} {name}  {'' if st=='PASS' else st}")
    print("\n关键形状（万元）：")
    print("  收入 :", [f"{v:,.0f}" for v in R["IS"]["rev"]])
    print("  毛利率:", [f"{R['IS']['gp'][i]/R['IS']['rev'][i]:.1%}" for i in range(N)])
    print("  净利 :", [f"{v:,.0f}" for v in R["IS"]["ni"]])
    print("  期末现金:", [f"{v:,.0f}" for v in R["BS"]["cash"]])
    print("  门店净值:", [f"{v:,.0f}" for v in R["fa_net"]])
