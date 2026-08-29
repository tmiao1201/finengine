"""指标字典：连锁咖啡实例。每个指标挂中文含义/维度适用/三种收入模式归类。
与 finengine（大模型公司）对照：换行业 = 换这份字典 + industry_config，引擎同构。"""

METRICS = {
    # 量×价（volume_price）——本行业主模式
    "stores":            {"label": "期末门店数（家）", "entity": "GROUP", "region": None, "pattern": "volume_price"},
    "daily_cups":        {"label": "日均杯量/店（杯）", "entity": "GROUP", "region": None, "pattern": "volume_price"},
    "avg_ticket":        {"label": "客单价（元）", "entity": "GROUP", "region": None, "pattern": "volume_price"},
    "retail_ratio":      {"label": "零售产品占现制比例", "entity": "GROUP", "region": None, "pattern": "volume_price"},
    # 成本
    "cogs_rate":         {"label": "原料成本率", "entity": "GROUP", "region": None, "pattern": "cost_rate"},
    "store_opex_m":      {"label": "单店月运营成本（万）", "entity": "GROUP", "region": None, "pattern": "cost_rate"},
    "mkt_rate":          {"label": "营销费率", "entity": "GROUP", "region": None, "pattern": "cost_rate"},
    # 结构
    "headcount_rd":      {"label": "研发编制（数字化）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "headcount_sales":   {"label": "销售市场编制", "entity": "GROUP", "region": None, "pattern": "structure"},
    "headcount_gna":     {"label": "管理编制", "entity": "GROUP", "region": None, "pattern": "structure"},
    "salary_rd":         {"label": "研发人均全成本（万/年）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "salary_sales":      {"label": "销售人均全成本（万/年）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "salary_gna":        {"label": "管理人均全成本（万/年）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "capex_per_store":   {"label": "单店投入（万）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "tax_rate":          {"label": "所得税率", "entity": "GROUP", "region": None, "pattern": "structure"},
    "dso_days":          {"label": "应收账期（天）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "dpo_days":          {"label": "应付账期（天）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "deferred_rate":     {"label": "储值/递延率", "entity": "GROUP", "region": None, "pattern": "structure"},
    "region_mix":        {"label": "收入地区占比", "entity": "GROUP", "region": "分地区", "pattern": "structure"},
    "equity_inject":     {"label": "当年注资（万）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "open_pi":           {"label": "期初实收资本（万）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "open_re":           {"label": "期初未分配利润（万）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "open_store_fa":     {"label": "期初门店资产原值（万）", "entity": "GROUP", "region": None, "pattern": "structure"},
}

REVENUE_PATTERNS = {
    "volume_price": "收入 = 量 × 价。本例：门店 × 杯量 × 客单价（另可：席位×订阅价/产能×出厂价/客流×客单）",
}
