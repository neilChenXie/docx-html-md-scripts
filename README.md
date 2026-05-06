# docx-html-md-scripts

文档格式转换工具集，用于 Word 文档（.docx）、HTML 和 Markdown 之间的相互转换。

## 目标

在需要 AI 来配合进行文档写作、结合 SDK 进行编程等工作时，需要把现在常见的 docx 文档转化成 markdown 文档，便于用作 AI 的 context。而目前，让 AI 每次写脚本转化，不能 handle 一些复杂的情况。比如自动编号、复杂表格等问题（详见 `.dev-logs/html_to_markdown.md`）。

## 两条技术路线

| 步骤 | 路线 1：WPS/Word COM（仅 Windows） | 路线 2：Pandoc（跨平台） |
|------|-----------------------------------|--------------------------|
| docx → html | `docx2html-win` CLI / `docx_to_html_win()` | `docx2html` CLI / `py_docx_to_html()` |
| html → md | `html2md` CLI / `html_to_markdown()` | `html2md` CLI / `py_html_to_markdown()` |
| 特点 | 对中文排版保真度更高 | 无需安装 Office，跨平台可用 |

**md 转 docx 文档：**
- md → docx：`md2docx` 命令 / `markdown_to_docx()`

## 项目文件结构

```
docx-html-md-scripts/
├── pyproject.toml               # 包元数据、依赖、CLI 入口点
├── src/
│   └── docx_md_tools/
│       ├── __init__.py          # 公共 API 导出
│       ├── py_docx_to_html.py   # docx → HTML（pypandoc + 编号标题注入）
│       ├── html_to_markdown.py  # HTML → Markdown（路线 1：WPS/Word 导出 HTML）
│       ├── py_html_to_markdown.py  # HTML → Markdown（路线 2：Pandoc 导出 HTML）
│       └── md_to_docx.py        # Markdown → docx
├── scripts/
│   └── docx_to_html.ps1         # PowerShell 脚本（Windows，由 docx2html-win 调用）
├── .gitignore
├── README.md
├── .dev-logs/                   # 开发日志
│   ├── doc_to_html.md
│   └── html_to_markdown.md
└── test-case/                   # 测试用例
    ├── docx转md用例1-保留自动编码.docx
    ├── docx转md用例1-保留自动编码-参考答案.md
    ├── docx转md用例2-复杂表格处理.html
    └── docx转md用例2-复杂表格处理-参考答案.md
```

## 安装

### 本地开发安装

```bash
git clone <repo-url>
cd docx-html-md-scripts

# 创建虚拟环境
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate  # Linux/macOS

# 可编辑模式安装（修改代码立即生效）
pip install -e .
```

### 在其他项目中安装

**方式 A — 从 Git 安装（推荐，pin 版本）**

```bash
# 初次安装
pip install git+https://github.com/you/docx-html-md-scripts.git@v0.1.0

# 更新到新版本
pip install --upgrade git+https://github.com/you/docx-html-md-scripts.git@v1.0.0
```

**方式 B — 本地 editable 模式（实时同步开发）**

```bash
pip install -e /path/to/docx-html-md-scripts
```
修改本项目代码后，其他项目立即生效，无需重新安装。

### 依赖

所有依赖已在 `pyproject.toml` 中声明，`pip install` 时自动安装：
- `beautifulsoup4`：HTML 解析
- `chardet`：编码检测
- `lxml`：XML/HTML 解析引擎
- `pypandoc`：Pandoc Python 绑定
- `python-docx`：Word 文档读写

## 使用方式

### 方式 1：CLI 命令（安装后全局可用）

```bash
# docx 转 html（Pandoc 跨平台路线）
docx2html document.docx                  # 输出 document.html
docx2html document.docx output.html      # 指定输出路径

# docx 转 html（WPS/Word COM 路线，仅 Windows，中文保真度更高）
docx2html-win document.docx              # 输出 document.html
docx2html-win document.docx output.html  # 指定输出路径

# html 转 markdown
html2md document.html                    # 输出 document.md
html2md document.html output.md          # 指定输出路径

# markdown 转 docx
md2docx document.md                      # 输出 document.docx
md2docx document.md -o output.docx       # 指定输出路径
md2docx document.md --output output.docx
```

### 方式 2：作为 Python 库导入

```python
from docx_md_tools import py_docx_to_html, docx_to_html_win, html_to_markdown, py_html_to_markdown, markdown_to_docx

# 路线 1：WPS/Word COM 导出 HTML → Markdown（仅 Windows）
html_path = docx_to_html_win("document.docx", "output.html")  # docx → html (COM)
html_to_markdown("output.html", "output.md")                   # html → md

# 路线 2：Pandoc 跨平台 → Markdown（含自动编号注入）
py_docx_to_html("document.docx", "output.html")  # docx → html + 编号标题
py_html_to_markdown("output.html", "output.md")   # html → md
```

### 方式 3：直接运行模块文件

```bash
python src/docx_md_tools/py_docx_to_html.py document.docx
python src/docx_md_tools/html_to_markdown.py document.html
python src/docx_md_tools/py_html_to_markdown.py document.html
python src/docx_md_tools/md_to_docx.py document.md
```

### 路线 1：WPS/Word COM（仅 Windows）

```bash
# 安装后直接使用 CLI 命令
docx2html-win 文档.docx
docx2html-win 文档.docx output.html

# 或在 Python 中调用
python -c "from docx_md_tools import docx_to_html_win; docx_to_html_win('文档.docx')"
```

## 完整示例

### 路线 1：WPS/Word COM（Windows）

```bash
# 1. docx -> html（WPS/Word COM）
docx2html-win 报告.docx

# 2. html -> md（WPS 路径）
html2md 报告.html

# 现在可以使用 "报告.md" 作为 AI 的 context
```

### 路线 2：Pandoc 跨平台

```bash
# 1. docx -> html（含自动编号注入）
docx2html 报告.docx

# 2. html -> md（Pandoc 路径）
html2md 报告.html

# 现在可以使用 "报告.md" 作为 AI 的 context
```

### 在 Python 脚本中

```python
from docx_md_tools import py_docx_to_html, py_html_to_markdown

# 路线 2：Pandoc 跨平台一键转换
html_path = py_docx_to_html("报告.docx")
md_path = py_html_to_markdown(html_path)
print(f"Markdown saved to: {md_path}")
```

## 版本管理与更新同步

- 语义化版本，通过 Git tag 管理（`v0.1.0`, `v0.2.0`, `v1.0.0` 等）
- Breaking changes 增加大版本号，新功能增加次版本号，修复增加补丁号
- 其他项目通过 `git+...@vX.Y.Z` 锁定版本，或使用 editable 模式实时同步

## 常见问题

**Q: 为什么不直接用 pandoc 将 docx 转换为 markdown？**

A: 直接转换会丢失 Word 文档中的自动编号（如 1.1.1、1.1.2）。通过 HTML 中转可以完整保留这些格式。详见 `.dev-logs/html_to_markdown.md`。

**Q: 转换后中文显示乱码怎么办？**

A: `html_to_markdown.py` 会自动尝试多种编码（GB2312、GBK、UTF-8 等），如果仍有问题，请确保原 docx 文件编码正确。

**Q: docx2html 命令和 PowerShell 脚本有什么区别？**

A: 这是两条不同的技术路线：

- **路线 1（`scripts/docx_to_html.ps1`）**：通过 WPS/Word COM 对象导出 HTML，对中文文档的格式保真度更高（尤其是复杂排版），但仅限 Windows 平台。后续使用 `html_to_markdown()` 转换为 md。

- **路线 2（`docx2html` / `py_docx_to_html()`）**：使用 pypandoc 跨平台转换，会自动从 docx XML 提取编号标题并注入 HTML，保留类似 1.1.1 的章节编号。后续使用 `py_html_to_markdown()` 转换为 md。

## 相关资源

- [开发日志：docx 转 HTML](.dev-logs/doc_to_html.md)
- [开发日志：HTML 转 Markdown](.dev-logs/html_to_markdown.md)
