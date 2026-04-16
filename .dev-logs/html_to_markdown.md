# html_to_markdown.py 开发经验

## 目标

将 docx 文件转化成 markdown 文件，便于 AI 后续的阅读。

## 尝试过程

### 1. 初次尝试：直接使用 docx 库和 pandoc 库

首先尝试使用 Python 的 `docx` 库和 `pandoc` 库来生成 markdown 文件，但是遇到一个无法解决的问题：**docx 文件中的自动编号（如 1.1.1、1.1.2、1.2.1 等）会缺失**。多次尝试均无法解决此问题。

### 2. 曲线救国：PDF 中转方案

想到一个思路：
1. 先将 docx 文件转换为 PDF 文件
2. 再将 PDF 文件转换为 markdown 文件
3. 最后让 AI 将 PDF 转换的 markdown 文件中的编号，对应修改 python 用 docx 库导出的 markdown 文件

这个方案虽然可行，但过于繁琐，需要多次转换和 AI 参与修正。

### 3. 最终方案：HTML 中转

突然想到一个更直接的思路：既然可以导出 PDF，**能不能直接导出 HTML 文件呢**？

验证后发现：HTML 转 markdown 非常简单和准确。最终成功，形成了 `html_to_markdown.py` 文件。

## 关键洞察

- 直接从 docx 转换会丢失自动编号格式
- HTML 格式能够完整保留文档结构和编号
- HTML → markdown 的转换有成熟且可靠的工具

## Bug 修复记录

### 表格 rowspan 导致列错位（2026-04-16）

**现象**：含 rowspan 的表格，"返回值说明"列内容丢失或错位。例如：

```
| 返回类型 | 返回值 |  |  |  |  |       ← 缺少"返回值说明"表头
| int | 0 |  |  |  |  |               ← 缺少说明文字
| -1 |  |  |  |  | 加载套接字库失败 |  ← int 的 rowspan 未展开，-1 跑到第1列
```

**根本原因**：`max_cols` 用 `len(raw_cells[row])` 计算，取的是每行 td 元素数量，而非展开后的实际列数。有 rowspan 的行，其后续行 td 数量少（因为被 rowspan 占用的列没有 td），导致 grid 宽度不足，填充时列偏移错误。

**修复方案**：改用 `dict` 作为 grid（key 为 `(row, col)`），解析每个 td 时主动跳过已被 rowspan 占用的位置，`max_cols` 从 grid 实际键值推导。彻底正确处理任意 rowspan/colspan 组合。

**核心代码变化**：
- 旧：`temp_grid = [[None] * max_cols ...]`，max_cols 基于 raw cell 数
- 新：`grid = {}`，按 `(row_idx, col_idx)` 直接定位，col_idx 自动跳过已占用位置