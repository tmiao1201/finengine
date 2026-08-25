# finengine · 大模型集团财务分析模型

虚构「智擎集团」（参考智谱商业模式），从零手搓的集团财务 Excel package。
Python + openpyxl **全公式**生成——改一个假设，全模型联动。

## 一分钟看懂

```bash
.venv/bin/python src/build.py        # 重建 dist/finengine.xlsx（23 张 sheet）
.venv/bin/python src/engine.py       # 影子引擎：17 项勾稽断言
.venv/bin/python src/recompute.py    # formulas 重算 1500+ 格 vs 影子值逐格对比
.venv/bin/python tests/read_checks.py # 读 Checks 面板 PASS/FAIL
```

## 集团架构

| 主体 | 业务 | 持股 |
|---|---|---|
| 智擎研究院 | 基座研发（收入全为关联交易） | — |
| 智擎云 | MaaS API（token 计费） | 100% |
| 智擎行业 | 私有化 + 解决方案 | 100% |
| 智擎互联 | C 端订阅 + 广告 | **80%**（少数股东） |

内部交易：T1 license / T2 定制研发 / T3 算力转售 / T4 往来挂账（60 天）。
合并抵消：E0 长投×权益、E1 关联对冲、E2 往来孰低、E3 少数股东。
彩蛋：IC_Recon 埋了一笔 220 万 cut-off 在途（互联已确认/云未开票），走完调节表归零。

## 勾稽体系（倒挤现金法）

1. BS 现金 = 负债+权益−非现金资产（**结构性平衡**）
2. CF 间接法独立推导期末现金
3. 断言 CF现金 = BS现金 —— 全模型最深的勾稽线
4. 合并层：BS 平衡、CF=Σ单体现金、归母+少数=合并NI、关联对账调节后=0

## 双引擎交叉验证

- `engine.py` Python 影子计算（17 断言全绿）
- Excel 公式复刻同一逻辑
- `formulas` 库重算 → 逐格对比 → 全部一致才交付

## 分工（graph engineering 实践）

M0 地基(串行) → Gate1 假设审 → M1 单体×4 → Gate2 BS平 → M2 关联层 → M3 合并层
→ Gate3 合并勾稽 → M4 分析层×3(subagent 并行) → M5 检查美化 → Gate4 端到端全绿。

## 已知简化（明示在模型内）

- 内部 license 当期结转，无未实现利润抵消
- 所得税不结转亏损；财务费用/利息收入不计
- 母公司长投成本法，合并差额进资本公积（不追溯权益法）
- BS/CF 科目无业务线/地区维度（Dashboard 仅 IS 四维联动）
