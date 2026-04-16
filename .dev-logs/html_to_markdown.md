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