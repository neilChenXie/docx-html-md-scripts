# docx-html-md-scripts

文档格式转换工具集，用于 Word 文档（.docx）、HTML 和 Markdown 之间的相互转换。

## 目标

在需要 AI 来配合进行文档写作、结合 SDK 进行编程等工作时，需要把现在常见的 docx 文档转化成 markdown 文档，便于用作 AI 的 context。而目前，让 AI 每次写脚本转化，不能 handle 一些复杂的情况。比如自动编号、复杂表格等问题（详见 `.dev-logs/html_to_markdown.md`）。

## 转换流程

**docx 文档转 md 流程：**
1. docx → html：`docx_to_html.ps1`
2. html → md：`html_to_markdown.py`

**md 转 docx 文档：**
- md → docx：`md_to_docx.py`

## 项目文件结构

```
docx-html-md-scripts/
├── docx_to_html.ps1       # PowerShell 脚本：Word 文档转 HTML
├── html_to_markdown.py    # Python 脚本：HTML 转 Markdown
├── md_to_docx.py          # Python 脚本：Markdown 转 Word 文档
├── requirements.txt       # Python 依赖
├── 将text文件转化成word文件.md  # 参考提示词
├── README.md              # 项目说明文档
├── .dev-logs/             # 开发日志
│   ├── doc_to_html.md
│   └── html_to_markdown.md
└── test-case/             # 测试用例
    ├── docx转md用例1-保留自动编码.docx
    └── docx转md用例1-保留自动编码-参考答案.md
```

## 安装依赖

### Python 脚本

```bash
# 创建并激活虚拟环境（推荐）
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# .venv\Scripts\activate   # Windows

# 安装依赖
pip install -r requirements.txt
```

**依赖说明：**
- `beautifulsoup4`：HTML 解析
- `lxml`：XML/HTML 解析引擎
- `python-docx`：Word 文档读写

### PowerShell 脚本

需要安装以下软件之一：
- **WPS Office**（推荐，中文文档兼容性最好）
- **Microsoft Word**
- **Pandoc**（跨平台，命令行工具）
- **LibreOffice**

## 脚本用途及执行命令

### docx_to_html.ps1

将 Word 文档 (.docx) 转换为 HTML 格式，基本等同 WPS/Word "另存为 HTML" 的效果。

**基本用法：**

```powershell
# 默认输出同名 .html 文件
.\docx_to_html.ps1 "文档.docx"

# 指定输出路径
.\docx_to_html.ps1 "文档.docx" "输出.html"
```

**在CMD、Bash中用法：**

```bash
# 方法1: 直接调用 PowerShell
powershell.exe -ExecutionPolicy Bypass -File "docx_to_html.ps1" "文档.docx"

# 方法2: 使用 pwsh (PowerShell 7+)
pwsh -File "docx_to_html.ps1" "文档.docx"
```

**注意事项：**

1. **首次运行 PowerShell 脚本**可能需要设置执行策略：
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
   ```
2. **转换过程中请勿关闭 WPS/Word** 窗口（虽然脚本会自动隐藏）
3. **输出文件会覆盖**已存在的同名文件
4. **图片资源**会自动提取到 `文件名.files` 目录

### html_to_markdown.py

将 Word 导出的 HTML 文件转换为 Markdown 格式，能够正确处理自动编号、复杂表格等结构。

**基本用法：**

```bash
# 默认输出同名 .md 文件
python html_to_markdown.py document.html

# 指定输出路径
python html_to_markdown.py document.html output.md
```

**功能特点：**
- 自动检测编码（支持 GB2312、GBK、UTF-8）
- 正确处理表格结构
- 保留有序/无序列表
- 支持标题、加粗、斜体等格式
- 自动清理 Word 特有的标记

### md_to_docx.py

将 Markdown 文件转换为 Word 文档（.docx），支持中文字体设置。

**基本用法：**

```bash
# 默认输出同名 .docx 文件
python md_to_docx.py document.md

# 指定输出路径
python md_to_docx.py document.md -o output.docx
python md_to_docx.py document.md --output output.docx
```

**功能特点：**
- 支持多级标题（H1-H6）
- 支持 Markdown 表格
- 支持行内格式（加粗等）
- 自动设置中文字体（宋体、黑体）

## 完整示例

```bash
# 1. 将 docx 转换为 HTML
pwsh -File docx_to_html.ps1 "报告.docx"

# 2. 将 HTML 转换为 Markdown
python html_to_markdown.py "报告.html"

# 现在可以使用 "报告.md" 作为 AI 的 context
```

## 常见问题

**Q: 为什么不直接用 pandoc 将 docx 转换为 markdown？**

A: 直接转换会丢失 Word 文档中的自动编号（如 1.1.1、1.1.2）。通过 HTML 中转可以完整保留这些格式。详见 `.dev-logs/html_to_markdown.md`。

**Q: 转换后中文显示乱码怎么办？**

A: `html_to_markdown.py` 会自动尝试多种编码，如果仍有问题，请确保原 docx 文件编码正确。

## 相关资源

- [开发日志：docx 转 HTML](.dev-logs/doc_to_html.md)
- [开发日志：HTML 转 Markdown](.dev-logs/html_to_markdown.md)
