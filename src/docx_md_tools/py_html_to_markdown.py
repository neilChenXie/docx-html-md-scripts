#!/usr/bin/env python3
"""
HTML to Markdown converter (Pandoc HTML path).
Supports WPS/Word HTML exports (GB2312/GBK) and Pandoc HTML (UTF-8 standard HTML5).
"""

from bs4 import BeautifulSoup
import chardet
import re
import sys
import os


def _td_content(td):
    items = td.find_all('li')
    if items:
        sep = 'Ø'
        return sep + sep.join(li.get_text().strip() for li in items)
    text = td.get_text(separator=' ').strip()
    text = re.sub(r'\s+', ' ', text)
    return text


def _convert_table(table):
    tbody = table.find('tbody') or table
    trs = tbody.find_all('tr')
    if not trs:
        return ""

    num_rows = len(trs)
    grid = {}

    for row_idx, tr in enumerate(trs):
        col_idx = 0
        for td in tr.find_all(['td', 'th']):
            while (row_idx, col_idx) in grid:
                col_idx += 1
            colspan = int(td.get('colspan', 1))
            rowspan = int(td.get('rowspan', 1))
            text = _td_content(td)
            cell = {'text': text, 'colspan': colspan}
            for r in range(rowspan):
                for c in range(colspan):
                    grid[(row_idx + r, col_idx + c)] = cell
            col_idx += colspan

    if not grid:
        return ""

    max_cols = max(c for (_, c) in grid.keys()) + 1

    result_rows = []
    for row_idx in range(num_rows):
        final_row = []
        col_idx = 0
        while col_idx < max_cols:
            cell = grid.get((row_idx, col_idx))
            if cell:
                colspan = cell['colspan']
                final_row.append(cell['text'])
                for _ in range(colspan - 1):
                    final_row.append('')
                col_idx += colspan
            else:
                final_row.append('')
                col_idx += 1
        result_rows.append(final_row[:max_cols])

    md = []
    md.append("| " + " | ".join(result_rows[0]) + " |")
    md.append("|" + "|".join(["---" for _ in range(max_cols)]) + "|")
    for row in result_rows[1:]:
        md.append("| " + " | ".join(row) + " |")

    return "\n".join(md) + "\n"


def _convert_list(lst, ordered=False, level=0):
    result = []
    items = lst.find_all('li', recursive=False)
    for i, item in enumerate(items):
        indent = "  " * level
        if ordered:
            prefix = f"{indent}{i+1})"
        else:
            prefix = "Ø"
        text = item.get_text().strip()
        result.append(prefix + text)
    sep = chr(10) * 2
    return sep.join(result) + chr(10)


def _convert_element(elem):
    result = []
    for child in elem.children:
        if child.name is None:
            text = " ".join(str(child).split())
            if text:
                result.append(text)
        elif child.name == 'p':
            para_text = _convert_element(child).strip()
            if para_text:
                result.append(para_text + "\n\n")
        elif child.name == 'br':
            result.append(" ")
        elif child.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            num_span = child.find('span', class_='header-section-number')
            if num_span:
                number = num_span.get_text().strip()
                num_span.decompose()
                title = child.get_text().strip()
                result.append(f"{number} **{title}**\n\n")
            else:
                level = int(child.name[1])
                text = child.get_text().strip()
                result.append("#" * level + " " + text + "\n\n")
        elif child.name in ['b', 'strong']:
            text = child.get_text().strip()
            if text:
                result.append(f"**{text}**")
        elif child.name in ['i', 'em']:
            text = child.get_text().strip()
            if text:
                result.append(f"*{text}*")
        elif child.name == 'table':
            result.append(_convert_table(child) + "\n")
        elif child.name in ['ul', 'ol']:
            ordered = child.name == 'ol'
            result.append(_convert_list(child, ordered) + "\n")
        elif child.name == 'div':
            result.append(_convert_element(child))
        elif child.name == 'section':
            result.append(_convert_element(child))
        elif child.name == 'pre':
            code_text = child.get_text().rstrip()
            if code_text:
                result.append(f"```\n{code_text}\n```\n\n")
        elif child.name == 'code':
            text = child.get_text().strip()
            if text:
                result.append(f"`{text}`")
        elif child.name == 'span':
            text = _convert_element(child)
            if text.strip():
                result.append(text)
        else:
            result.append(_convert_element(child))
    return "".join(result)


def py_html_to_markdown(html_path, md_path=None):
    """
    Convert a Word-exported HTML file to Markdown (Pandoc path).

    Args:
        html_path: Path to the HTML input file
        md_path:   Path to the Markdown output file. Defaults to .md extension.

    Returns:
        True on success, False on failure
    """
    if not os.path.exists(html_path):
        print(f"Error: File not found: {html_path}")
        return False

    if md_path is None:
        md_path = html_path.replace('.html', '.md').replace('.htm', '.md')

    html = None
    encodings = []
    with open(html_path, "rb") as f:
        raw = f.read()
    detected = chardet.detect(raw)
    if detected and detected['encoding']:
        encodings.append(detected['encoding'].lower())
    encodings.extend(['utf-8', 'gb2312', 'gbk', 'gb18030'])

    for enc in encodings:
        try:
            html = raw.decode(enc)
            print(f"Using encoding: {enc}")
            break
        except (UnicodeDecodeError, LookupError):
            continue

    if html is None:
        print("Error: Unable to read file, encoding not supported")
        return False

    soup = BeautifulSoup(html, 'html.parser')

    for tag in soup(["script", "style", "meta", "link"]):
        tag.decompose()

    body = soup.find('body')
    content = _convert_element(body) if body else _convert_element(soup)

    content = re.sub(r'\[if !supportLists\]', '', content)
    content = re.sub(r'\[endif\]', '', content)
    content = re.sub(r'StartFragment|EndFragment', '', content)
    content = re.sub(r'^(\d+(?:\.\d+)*)(\*\*)', r'\1 \2', content, flags=re.MULTILINE)
    content = re.sub(r'\*\*\*\*', '', content)
    content = re.sub(r'\n{3,}', '\n\n', content)
    content = '\n'.join(line.rstrip() for line in content.split('\n'))

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"[OK] Conversion complete: {md_path}")
    print(f"[OK] File size: {len(content)} characters")
    return True


def main():
    if len(sys.argv) < 2:
        print("Usage: python py_html_to_markdown.py <html_file> [output_md_file]")
        print("Example: python py_html_to_markdown.py document.html")
        print("      python py_html_to_markdown.py document.html output.md")
        sys.exit(1)

    html_path = sys.argv[1]
    md_path = sys.argv[2] if len(sys.argv) > 2 else None

    success = py_html_to_markdown(html_path, md_path)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
