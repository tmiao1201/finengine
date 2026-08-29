"""指标字典 —— 数据契约的真相源。
行业不在代码里，在指标集合里。本文件 = 「大模型公司」这个行业的实例化；
换行业 = 换一份 METRICS 实例 + industry_config，引擎逻辑零改动。

pattern 归类（任何行业收入都逃不出三种模式的组合）：
  volume_price      量×价（MaaS token、SaaS 席位、零售坪效、制造产能）
  cohort_retention  客户×留存（订阅、C端付费、会员）
  project_contract  项目×合同（私有化部署、工程总包、定制开发）
  cost_rate         成本率/单位成本模式
  structure         结构参数（账期、税率、编制、融资）
"""

METRICS = {
    # ── 集团结构参数（structure）──
    "tax_rate":        {"label": "所得税率", "entity": "GROUP", "region": None, "pattern": "structure"},
    "dso_days":        {"label": "应收账期（天）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "dpo_days":        {"label": "应付账期（天）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "ic_days":         {"label": "内部往来账期（天）", "entity": "GROUP", "region": None, "pattern": "structure"},
    "wacc":            {"label": "WACC", "entity": "GROUP", "region": None, "pattern": "structure"},
    "region_mix":      {"label": "收入地区占比", "entity": "GROUP", "region": "分地区", "pattern": "structure"},

    # ── 智擎云 · MaaS（volume_price）──
    "token_volume_day": {"label": "日均 token 消耗（B/day）", "entity": "CLOUD", "region": None, "pattern": "volume_price"},
    "api_price":        {"label": "综合单价（元/M tokens）", "entity": "CLOUD", "region": None, "pattern": "volume_price"},
    "inference_ucost":  {"label": "推理单位成本（元/M tok）", "entity": "CLOUD", "region": None, "pattern": "cost_rate"},
    "bandwidth_fee":    {"label": "带宽机房费用（万）", "entity": "CLOUD", "region": None, "pattern": "cost_rate"},

    # ── 智擎行业（project_contract + cohort_retention）──
    "projects_pvt":     {"label": "私有化项目数（个）", "entity": "IND", "region": None, "pattern": "project_contract"},
    "avg_contract":     {"label": "平均合同额（万）", "entity": "IND", "region": None, "pattern": "project_contract"},
    "cogs_rate_pvt":    {"label": "私有化成本率", "entity": "IND", "region": None, "pattern": "cost_rate"},
    "subscribers_sol":  {"label": "解决方案签约客户（累计）", "entity": "IND", "region": None, "pattern": "cohort_retention"},
    "sub_fee":          {"label": "解决方案年费（万）", "entity": "IND", "region": None, "pattern": "cohort_retention"},
    "cogs_rate_sol":    {"label": "解决方案成本率", "entity": "IND", "region": None, "pattern": "cost_rate"},

    # ── 智擎互联 · C端（cohort_retention）──
    "mau_w":            {"label": "MAU（万人）", "entity": "TOC", "region": None, "pattern": "cohort_retention"},
    "pay_rate":         {"label": "付费转化率", "entity": "TOC", "region": None, "pattern": "cohort_retention"},
    "arppu":            {"label": "ARPPU（元/付费/月）", "entity": "TOC", "region": None, "pattern": "cohort_retention"},
    "ads_rev":          {"label": "广告收入（万）", "entity": "TOC", "region": None, "pattern": "cohort_retention"},
    "ucost_toc":        {"label": "单位推理+带宽成本（元/MAU/月）", "entity": "TOC", "region": None, "pattern": "cost_rate"},
    "cac":              {"label": "获客成本（元/新增用户）", "entity": "TOC", "region": None, "pattern": "cost_rate"},
    "new_user_mult":    {"label": "新增用户=MAU增量×", "entity": "TOC", "region": None, "pattern": "structure"},

    # ── 编制与薪酬（structure，按主体实例化）──
    "headcount_rd":     {"label": "研发编制（人）", "entity": "ALL", "region": None, "pattern": "structure"},
    "headcount_sales":  {"label": "销售编制（人）", "entity": "ALL", "region": None, "pattern": "structure"},
    "headcount_gna":    {"label": "管理编制（人）", "entity": "ALL", "region": None, "pattern": "structure"},
    "salary_rd":        {"label": "研发人均全成本（万/年）", "entity": "ALL", "region": None, "pattern": "structure"},
    "salary_sales":     {"label": "销售人均全成本（万/年）", "entity": "ALL", "region": None, "pattern": "structure"},
    "salary_gna":       {"label": "管理人均全成本（万/年）", "entity": "ALL", "region": None, "pattern": "structure"},

    # ── 智擎研究院专属 ──
    "hold_salary":      {"label": "研究院研发人均全成本（万）", "entity": "HOLD", "region": None, "pattern": "structure"},
    "other_rd":         {"label": "其他研发费用（万）", "entity": "HOLD", "region": None, "pattern": "cost_rate"},

    # ── 关联交易定价（structure）──
    "t1_rate":          {"label": "T1 license 费率", "entity": "BY_SUB", "region": None, "pattern": "structure"},
    "t2_fee":           {"label": "T2 定制研发服务费（万）", "entity": "HOLD", "region": None, "pattern": "structure"},
    "t3_share":         {"label": "T3 算力结算占比", "entity": "GROUP", "region": None, "pattern": "structure"},
    "t3_gm":            {"label": "T3 云内部毛利率", "entity": "GROUP", "region": None, "pattern": "structure"},
    "cutoff_2027":      {"label": "cut-off 在途（万）", "entity": "GROUP", "region": None, "pattern": "structure"},

    # ── 融资与期初（structure）──
    "equity_inject":    {"label": "当年新增注资（万）", "entity": "BY_SUB", "region": None, "pattern": "structure"},
    "open_re":          {"label": "期初未分配利润（万）", "entity": "BY_SUB", "region": None, "pattern": "structure"},
    "open_pi":          {"label": "期初实收资本（万）", "entity": "BY_SUB", "region": None, "pattern": "structure"},
    "open_fa":          {"label": "期初 GPU 原值（万）", "entity": "HOLD_CLOUD", "region": None, "pattern": "structure"},
    "open_lt_inv":      {"label": "期初长投合计（万）", "entity": "HOLD", "region": None, "pattern": "structure"},
    "deferred_rate":    {"label": "递延收入率", "entity": "BY_SUB", "region": None, "pattern": "structure"},

    # ── GPU 资本开支 ──
    "capex_gpu":        {"label": "GPU 采购（万）", "entity": "HOLD_CLOUD", "region": None, "pattern": "structure"},

    # ── UE 参考（不进三表，挂分析层）──
    "customers":        {"label": "付费客户数（家）", "entity": "CLOUD", "region": None, "pattern": "cohort_retention"},
    "churn_m":          {"label": "月流失率", "entity": "CLOUD", "region": None, "pattern": "cohort_retention"},
}

# 三种收入模式的识别问题（skill 方法论用）
REVENUE_PATTERNS = {
    "volume_price":     "收入 = 量 × 价。量：token/席位/产能/客流；价：单价。先规模后价格通缩是常态",
    "cohort_retention": "收入 = 客户数 × ARPU × 留存。核心看 LTV/CAC 与 cohort 斜率",
    "project_contract": "收入 = 项目数 × 合同额 + 尾部运维年金。核心看交付成本率与现金节奏",
}
