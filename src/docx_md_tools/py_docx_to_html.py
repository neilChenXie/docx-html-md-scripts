#!/usr/bin/env python3
"""
Convert Word document (.docx) to HTML format using pypandoc,
with automatic numbered-heading extraction and injection.
"""

import os
import re as _re
import subprocess
import sys
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

import pypandoc
from bs4 import BeautifulSoup


def _extract_numbered_headings(docx_path):
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    with zipfile.ZipFile(docx_path) as z:
        doc = ET.fromstring(z.read('word/document.xml'))
        num_xml = ET.fromstring(z.read('word/numbering.xml'))

    num_to_abstract = {}
    for num in num_xml.findall('w:num', ns):
        nid = num.get(f'{{{W}}}numId')
        abst = num.find('w:abstractNumId', ns)
        if abst is not None:
            num_to_abstract[nid] = abst.get(f'{{{W}}}val')

    abstract_lvl_text = {}
    abstract_lvl_start = {}
    for abst in num_xml.findall('w:abstractNum', ns):
        aid = abst.get(f'{{{W}}}abstractNumId')
        abstract_lvl_text[aid] = {}
        abstract_lvl_start[aid] = {}
        for lvl in abst.findall('w:lvl', ns):
            ilvl = lvl.get(f'{{{W}}}ilvl')
            txt = lvl.find('w:lvlText', ns)
            start = lvl.find('w:start', ns)
            if txt is not None:
                abstract_lvl_text[aid][ilvl] = txt.get(f'{{{W}}}val', '')
            abstract_lvl_start[aid][ilvl] = int(start.get(f'{{{W}}}val', 1)) if start is not None else 1

    counters = {}
    result = []

    body = doc.find('.//w:body', ns)
    for p in body.findall('w:p', ns):
        num_pr = p.find('.//w:numPr', ns)
        if num_pr is None:
            continue
        num_id_el = num_pr.find('w:numId', ns)
        ilvl_el = num_pr.find('w:ilvl', ns)
        if num_id_el is None or ilvl_el is None:
            continue

        nid = num_id_el.get(f'{{{W}}}val')
        ilvl = ilvl_el.get(f'{{{W}}}val')
        ilvl_int = int(ilvl)

        if nid not in counters:
            counters[nid] = {}
        for k in list(counters[nid].keys()):
            if int(k) > ilvl_int:
                del counters[nid][k]
        counters[nid][ilvl] = counters[nid].get(ilvl, 0) + 1

        aid = num_to_abstract.get(nid)
        lvl_text_tpl = abstract_lvl_text.get(aid, {}).get(ilvl, '')
        number_str = lvl_text_tpl
        for i in range(ilvl_int + 1):
            if str(i) in counters[nid]:
                cnt = counters[nid][str(i)]
            else:
                cnt = abstract_lvl_start.get(aid, {}).get(str(i), 1)
            number_str = number_str.replace(f'%{i+1}', str(cnt))

        text = ''.join(t.text or '' for t in p.findall('.//w:t', ns)).strip()
        if text:
            if _re.match(r'^\d+\)$', number_str):
                continue
            result.append((text, number_str, ilvl_int + 1))

    return result


def _inject_numbered_headings(html_path, numbered_headings):
    with open(html_path, encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')

    heading_map = defaultdict(list)
    for text, number_str, level in numbered_headings:
        heading_map[text].append((number_str, level))
    heading_idx = defaultdict(int)

    for p in soup.find_all('p'):
        children = [c for c in p.children if str(c).strip()]
        if len(children) != 1:
            continue
        child = children[0]
        if getattr(child, 'name', None) == 'strong':
            text = child.get_text().strip()
        elif child.name is None and p.parent and p.parent.name == 'li':
            text = str(child).strip()
            if len(text) > 30:
                continue
        else:
            continue
        if text not in heading_map:
            continue
        idx2 = heading_idx[text]
        if idx2 >= len(heading_map[text]):
            continue
        number_str, level = heading_map[text][idx2]
        heading_idx[text] += 1
        h_level = min(level + 1, 6)
        new_tag = soup.new_tag(f'h{h_level}')
        span = soup.new_tag('span', attrs={'class': 'header-section-number'})
        span.string = number_str
        new_tag.append(span)
        new_tag.append(' ' + text)
        p.replace_with(new_tag)
        while new_tag.parent and new_tag.parent.name in ('li', 'ol', 'ul'):
            container = new_tag.parent
            if container.name == 'li':
                list_tag = container.parent
                list_tag.insert_before(new_tag.extract())
                if not container.get_text(strip=True):
                    container.decompose()
            else:
                container.insert_before(new_tag.extract())

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(str(soup))


def py_docx_to_html(input_file, output_file=""):
    """
    Convert Word document (.docx) to HTML using pypandoc with numbered-heading injection.

    Args:
        input_file: Path to the .docx input file
        output_file: Path to the .html output file. If empty, defaults to
                     same name as input file with .html extension.

    Returns:
        output_file path on success

    Raises:
        FileNotFoundError: If input_file does not exist
        ValueError: If input_file does not have .docx extension
    """
    if not os.path.isfile(input_file):
        raise FileNotFoundError(f"File not found: {input_file}")

    input_file = os.path.realpath(input_file)

    if not input_file.lower().endswith(".docx"):
        raise ValueError(f"Input file must have .docx extension: {input_file}")

    input_dir = os.path.dirname(input_file)
    input_basename = os.path.splitext(os.path.basename(input_file))[0]

    if not output_file:
        output_file = os.path.join(input_dir, f"{input_basename}.html")
    else:
        output_file = os.path.expanduser(output_file)
        if not os.path.isabs(output_file):
            output_file = os.path.join(os.getcwd(), output_file)

    output_dir = os.path.dirname(output_file)
    os.makedirs(output_dir, exist_ok=True)

    print(f"Converting: {input_file}")
    print(f"Output:     {output_file}")

    pandoc_version = pypandoc.get_pandoc_version()
    print(f"Using Pandoc ({pandoc_version}) via pypandoc...")

    media_dir = os.path.splitext(output_file)[0] + ".files"

    output = pypandoc.convert_file(
        input_file,
        "html",
        format="docx",
        extra_args=[
            "--standalone",
            "--extract-media=" + media_dir,
        ],
        outputfile=output_file,
    )

    print("[OK] Pandoc conversion successful")

    numbered_headings = _extract_numbered_headings(input_file)
    if numbered_headings:
        _inject_numbered_headings(output_file, numbered_headings)
        print(f"[OK] Injected {len(numbered_headings)} numbered headings")

    if os.path.isfile(output_file):
        file_size = os.path.getsize(output_file)
        print()
        print("=" * 40)
        print("[OK] Conversion successful!")
        print(f"[OK] Output: {output_file}")
        print(f"[OK] Size: {file_size} bytes")
        print("=" * 40)

    return output_file


def main():
    if len(sys.argv) < 2:
        print("Usage: python py_docx_to_html.py <input.docx> [output.html]")
        print("Example: python py_docx_to_html.py document.docx")
        print("         python py_docx_to_html.py document.docx output.html")
        sys.exit(1)

    try:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) >= 3 else ""
        py_docx_to_html(input_file, output_file)
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")
        sys.exit(1)


def docx_to_html_win(input_file, output_file=None):
    """通过 WPS/Word/LibreOffice COM（PowerShell）将 .docx 转为 HTML（路线 1）

    与 py_docx_to_html()（路线 2：Pandoc 跨平台）不同，此函数调用
    scripts/docx_to_html.ps1，通过 COM 接口驱动 WPS/Word/LibreOffice
    另存为 HTML，对中文排版保真度更高。仅 Windows 可用。

    Args:
        input_file: 输入 .docx 文件路径
        output_file: 输出 .html 文件路径（可选，默认与输入同目录同主名）

    Returns:
        output_file 路径字符串

    Raises:
        FileNotFoundError: 找不到 ps1 脚本或输入文件
        RuntimeError: PowerShell 转换失败
    """
    scripts_dir = Path(__file__).resolve().parent.parent / "scripts"
    ps1_path = scripts_dir / "docx_to_html.ps1"

    if not ps1_path.is_file():
        raise FileNotFoundError(f"PowerShell script not found: {ps1_path}")
    if not os.path.isfile(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")

    cmd = [
        "powershell.exe", "-ExecutionPolicy", "Bypass",
        "-File", str(ps1_path), str(Path(input_file).resolve()),
    ]
    if output_file:
        cmd.append(str(Path(output_file).resolve()))

    proc = subprocess.run(cmd, capture_output=True, text=True)
    sys.stdout.write(proc.stdout)
    if proc.returncode != 0:
        sys.stderr.write(proc.stderr)
        raise RuntimeError("PowerShell conversion failed")

    if output_file:
        return output_file
    return str(Path(input_file).with_suffix(".html"))


def win_main():
    """CLI 入口点：使用 WPS/Word COM 路线（路线 1，仅 Windows）

    用法: docx2html-win <input.docx> [output.html]
    """
    if sys.platform != "win32":
        print("Error: docx2html-win requires Windows (WPS/Word COM)")
        sys.exit(1)

    if len(sys.argv) < 2:
        print("Usage: docx2html-win <input.docx> [output.html]")
        print("Example: docx2html-win document.docx")
        print("         docx2html-win document.docx output.html")
        sys.exit(1)

    try:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) >= 3 else None
        docx_to_html_win(input_file, output_file)
    except (FileNotFoundError, RuntimeError) as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
