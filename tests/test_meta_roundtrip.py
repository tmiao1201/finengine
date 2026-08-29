"""Meta roundtrip 证明：数据与逻辑分离不是口号。
从生成的 xlsx 读回 Meta 宽表 → bind 进引擎 → 重算 → 与原影子值一致。
等价于：进公司后把 ERP 数据灌进 Meta 表，三表引擎零改动直接出数。"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import engine, meta

# 1) 基线（配置源）
base = engine.build_all()

# 2) 从 xlsx 读回 Meta（模拟外部灌数）
path = os.path.join(os.path.dirname(__file__), "..", "dist", "finengine.xlsx")
rows = meta.load_meta_from_xlsx(path)
print(f"从 xlsx 读回 Meta：{len(rows)} 行")

# 3) 绑定新数据源，引擎重算
meta.bind(rows)
again = engine.build_all()

# 4) 逐格对比核心科目
bad = 0
checked = 0
for ent in meta.ENTITIES:
    for stmt in ["IS", "BS", "CF"]:
        for k, vals in base["E"][ent][stmt].items():
            for i, v in enumerate(vals):
                v2 = again["E"][ent][stmt][k][i]
                checked += 1
                if abs(float(v) - float(v2)) > max(1e-6, abs(float(v)) * 1e-9):
                    bad += 1
for k, vals in base["C"]["IS"].items():
    for i, v in enumerate(vals):
        checked += 1
        if abs(float(v) - float(again["C"]["IS"][k][i])) > 1e-6:
            bad += 1

print(f"对比 {checked} 格 → {'✅ 全部一致' if bad == 0 else f'❌ {bad} 格不一致'}")
print("结论：换数据源（xlsx 灌数）→ 引擎零改动 → 三表输出完全一致" if bad == 0 else "数据与逻辑未解耦！")
sys.exit(0 if bad == 0 else 1)
