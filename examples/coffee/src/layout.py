"""坐标注册器：全模型唯一的『单元格地址 ↔ 语义 key ↔ 影子期望值』真相源。
生成器写公式时注册 (sheet, row, key)，期望值来自影子引擎；
recompute.py 用它做逐格交叉验证。"""
from openpyxl.utils import get_column_letter

class Layout:
    def __init__(self):
        self.row_of = {}     # (sheet, key) -> row
        self.expected = {}   # (sheet, key, col_idx) -> value  col_idx 0-4

    def reg(self, sheet, key, row, expected=None):
        """登记一个科目行。expected: 5 元组影子值（可 None）"""
        self.row_of[(sheet, key)] = row
        if expected is not None:
            for i, v in enumerate(expected):
                self.expected[(sheet, key, i)] = v

    def ref(self, sheet, key, i=0, absolute_sheet=True):
        """拼引用：默认当年列。key 也可直接是行号 int。"""
        row = key if isinstance(key, int) else self.row_of[(sheet, key)]
        s = f"{sheet}!" if absolute_sheet else ""
        return f"{s}{get_column_letter(3 + i)}{row}"

    def refs(self, sheet, key, s=None, e=None):
        """整行 5 列引用列表"""
        return [self.ref(sheet, key, i) for i in range(5)]

    def val(self, sheet, key, i):
        return self.expected.get((sheet, key, i))
