# docx-html-md-scripts

文档格式转换工具集，用于 Word 文档（.docx）、HTML 和 Markdown 之间的相互转换。

## 目标

在需要 AI 来配合进行文档写作、结合 SDK 进行编程等工作时，需要把现在常见的 docx 文档转化成 markdown 文档，便于用作 AI 的 context。而目前，让 AI 每次写脚本转化，不能 handle 一些复杂的情况。比如自动编号、复杂表格等问题（详见 `.dev-logs/html_to_markdown.md`）。

## 转换流程

**docx 文档转 md 流程：**
1. docx → html：`docx_to_html.py`（跨平台）或 `docx_to_html.ps1`（仅 Windows，需要 WPS/Word）
2. html → md：`html_to_markdown.py`

**md 转 docx 文档：**
- md → docx：`md_to_docx.py`

## 项目文件结构

```
docx-html-md-scripts/
├── docx_to_html.py         # Python 脚本：Word 文档转 HTML（跨平台，使用 pypandoc）
├── docx_to_html.ps1        # PowerShell 脚本：Word 文档转 HTML（Windows，使用 WPS/Word）
├── html_to_markdown.py     # Python 脚本：HTML 转 Markdown
├── md_to_docx.py           # Python 脚本：Markdown 转 Word 文档
├── requirements.txt        # Python 依赖
├── .gitignore              # Git 忽略规则
├── 将text文件转化成word文件.md  # 参考提示词
├── README.md               # 项目说明文档
├── .dev-logs/              # 开发日志
│   ├── doc_to_html.md
│   └── html_to_markdown.md
└── test-case/              # 测试用例
    ├── docx转md用例1-保留自动编码.docx
    ├── docx转md用例1-保留自动编码-参考答案.md
    ├── docx转md用例2-复杂表格处理.html
    └── docx转md用例2-复杂表格处理-参考答案.md
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
- `chardet`：编码检测
- `lxml`：XML/HTML 解析引擎
- `pypandoc`：Pandoc Python 绑定，用于 docx → HTML 转换
- `python-docx`：Word 文档读写

### PowerShell 脚本

`docx_to_html.ps1` 需要安装以下软件之一：
- **WPS Office**（推荐，中文文档兼容性最好）
- **Microsoft Word**
- **Pandoc**（跨平台，命令行工具）
- **LibreOffice**

## 脚本用途及执行命令

### docx_to_html.py

将 Word 文档 (.docx) 转换为 HTML 格式，使用 pypandoc（Pandoc Python 绑定）。跨平台可用，无需安装 WPS/Word。

**基本用法：**

```bash
# 默认输出同名 .html 文件
python docx_to_html.py document.docx

# 指定输出路径
python docx_to_html.py document.docx output.html
```

**功能特点：**
- 跨平台（Windows / macOS / Linux 均可使用）
- 无需安装 WPS 或 Microsoft Word
- 图片资源自动提取到 `文件名.files` 目录

### docx_to_html.ps1

将 Word 文档 (.docx) 转换为 HTML 格式，基本等同 WPS/Word "另存为 HTML" 的效果。仅适用于 Windows 平台。

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

将 HTML 文件转换为 Markdown 格式。支持两种 HTML 来源：

- **WPS/Word 导出 HTML**（GB2312/GBK 编码，含 `mso-list` 标记）
- **Pandoc 导出 HTML**（UTF-8 编码，标准 HTML5 结构）

能够正确处理自动编号、复杂表格等结构。

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

### 跨平台（推荐，使用 docx_to_html.py）

```bash
# 1. 将 docx 转换为 HTML
python docx_to_html.py 报告.docx

# 2. 将 HTML 转换为 Markdown
python html_to_markdown.py 报告.html

# 现在可以使用 "报告.md" 作为 AI 的 context
```

### Windows（使用 PowerShell 脚本）

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

**Q: docx_to_html.py 和 docx_to_html.ps1 有什么区别？**

A: `docx_to_html.py` 使用 pypandoc，跨平台可用且无需安装 Office 软件；`docx_to_html.ps1` 通过 WPS/Word COM 对象导出，对中文文档的格式保真度更高（尤其是复杂排版），但仅限 Windows 平台。

## 相关资源

- [开发日志：docx 转 HTML](.dev-logs/doc_to_html.md)
- [开发日志：HTML 转 Markdown](.dev-logs/html_to_markdown.md)
