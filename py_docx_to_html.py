#!/usr/bin/env python3
#
# Convert Word document (.docx) to HTML format
#
# Uses pandoc Python package (pypandoc) instead of command-line tool.
#
# Usage:
#   python docx_to_html_py.py <input.docx> [output.html]

import os
import sys
import zipfile
from xml.etree import ElementTree as ET

import pypandoc
from bs4 import BeautifulSoup


def extract_numbered_headings(docx_path):
    """从docx XML提取自动编号段落，返回 {段落文本: (编号字符串, 层级)} 列表（保序）"""
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    with zipfile.ZipFile(docx_path) as z:
        doc = ET.fromstring(z.read('word/document.xml'))
        num_xml = ET.fromstring(z.read('word/numbering.xml'))

    # 构建 numId -> abstractNumId
    num_to_abstract = {}
    for num in num_xml.findall('w:num', ns):
        nid = num.get(f'{{{W}}}numId')
        abst = num.find('w:abstractNumId', ns)
        if abst is not None:
            num_to_abstract[nid] = abst.get(f'{{{W}}}val')

    # 构建 abstractNumId -> {ilvl -> lvlText} 和 {ilvl -> start}
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

    # 遍历段落，跟踪每个numId的各级计数器
    counters = {}  # numId -> {ilvl -> count}
    result = []   # [(text, number_str, heading_level)]

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
        # 重置更深层级的计数器
        for k in list(counters[nid].keys()):
            if int(k) > ilvl_int:
                del counters[nid][k]
        counters[nid][ilvl] = counters[nid].get(ilvl, 0) + 1

        # 生成编号字符串：对未出现的父级用其start值
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
            # 只保留章节标题格式：包含点号分隔（1.1）或全角括号（（1））的编号
            # 排除纯数字+半角括号格式（1)、2)）这类列表项
            import re as _re
            if _re.match(r'^\d+\)$', number_str):
                continue
            result.append((text, number_str, ilvl_int + 1))

    return result


def inject_numbered_headings(html_path, numbered_headings):
    """将自动编号标题注入HTML：把 <p><strong>文字</strong></p> 替换为带编号的标题标签"""
    with open(html_path, encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')

    # 建立文本->编号信息的映射（文本可能重复，用列表+索引）
    from collections import defaultdict
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


def main():
    if len(sys.argv) < 2:
        print("Usage: python docx_to_html_py.py <input.docx> [output.html]")
        print("Example: python docx_to_html_py.py document.docx")
        print("         python docx_to_html_py.py document.docx output.html")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) >= 3 else ""

    if not os.path.isfile(input_file):
        print(f"Error: File not found: {input_file}")
        sys.exit(1)

    input_file = os.path.realpath(input_file)

    if not input_file.lower().endswith(".docx"):
        print(f"Error: Input file must have .docx extension: {input_file}")
        sys.exit(1)

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

    numbered_headings = extract_numbered_headings(input_file)
    if numbered_headings:
        inject_numbered_headings(output_file, numbered_headings)
        print(f"[OK] Injected {len(numbered_headings)} numbered headings")

    if os.path.isfile(output_file):
        file_size = os.path.getsize(output_file)
        print()
        print("=" * 40)
        print("[OK] Conversion successful!")
        print(f"[OK] Output: {output_file}")
        print(f"[OK] Size: {file_size} bytes")
        print("=" * 40)


if __name__ == "__main__":
    main()