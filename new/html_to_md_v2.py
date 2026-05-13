#!/usr/bin/env python3
"""
HTML 转 Markdown v2
支持 WPS/Word 导出 HTML (GB2312/GBK) 和 Pandoc HTML (UTF-8 标准 HTML5)

v2 修复:
1. 自动编号标题: 移除 mso-list 自动编号 span, 保留正确标题文本
2. 条件注释清理: 移除 [if supportFields] 等条件注释产生的 PAGEREF/raw HTML
3. 图片处理: 确保 <p> 内的 <img> 正确输出
4. --debug 诊断模式
"""

from bs4 import BeautifulSoup, Comment, ProcessingInstruction
import chardet
import re
import sys
import os
import argparse


def _td_content(td):
    parts = []
    for child in td.descendants:
        if child.name == 'img':
            src = child.get('src', '')
            alt = child.get('alt', '').strip()
            if src:
                if not alt:
                    alt = os.path.basename(src)
                parts.append(f"![{alt}]({src})")
        elif child.name is None:
            text = str(child).strip()
            if text:
                parts.append(text)
    return ' '.join(parts).strip()


def convert_table(table):
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


def convert_list(lst, ordered=False, level=0):
    result = []
    items = lst.find_all('li', recursive=False)

    for i, item in enumerate(items):
        indent = "  " * level
        if ordered:
            prefix = f"{indent}{i+1}. "
        else:
            prefix = f"{indent}- "

        text = item.get_text().strip()
        if text:
            result.append(prefix + text)

    return "\n".join(result) + "\n"


def _get_heading_text(elem, debug=False):
    style = elem.get('style', '')
    has_mso_list = 'mso-list' in style

    if not has_mso_list:
        text = elem.get_text().strip()
        text = text.replace('\xa0', ' ')
        return text

    if debug:
        print(f"[DEBUG] _get_heading_text: <{elem.name}> style='{style[:100]}'")
        for child in elem.children:
            if hasattr(child, 'name') and child.name:
                sty = child.get('style', '')[:60] if child.get('style') else 'N/A'
                print(f"[DEBUG]   child <{child.name}> style='{sty}' text='{child.get_text().strip()[:40]}'")
            elif hasattr(child, 'name') and child.name is None:
                print(f"[DEBUG]   child NavString: '{str(child).strip()[:40]}'")
            elif isinstance(child, Comment):
                print(f"[DEBUG]   child Comment: '{str(child).strip()[:40]}'")

    in_conditional = False
    parts = []
    for child in elem.children:
        if isinstance(child, (Comment, ProcessingInstruction)):
            comment_text = str(child)
            if 'supportLists' in comment_text or 'supportFields' in comment_text:
                in_conditional = True
                continue
            if 'endif' in comment_text:
                in_conditional = False
                continue
            continue

        if in_conditional:
            if child.name == 'a':
                continue
            if child.name == 'span':
                span_style = child.get('style', '')
                if 'mso-list:Ignore' in span_style:
                    numbering_text = child.get_text().strip().replace('\xa0', ' ')
                    if numbering_text:
                        parts.append(numbering_text)
                    continue
            if child.name == 'b':
                inner = child.get_text().strip().replace('\xa0', ' ')
                if inner:
                    parts.append(inner)
                continue
            inner = child.get_text().strip().replace('\xa0', ' ')
            if inner:
                parts.append(inner)
            continue

        if child.name == 'a':
            continue
        if child.name == 'span':
            span_style = child.get('style', '')
            if 'mso-list:Ignore' in span_style:
                numbering_text = child.get_text().strip().replace('\xa0', ' ')
                if numbering_text:
                    parts.append(numbering_text)
                continue
            inner = convert_element(child).strip().replace('\xa0', ' ')
            if inner:
                parts.append(inner)
            continue
        if child.name is None:
            text = str(child).strip().replace('\xa0', ' ')
            if text:
                parts.append(text)
            continue
        if child.name in ('b', 'strong'):
            inner = child.get_text().strip().replace('\xa0', ' ')
            if inner:
                parts.append(inner)
            continue
        inner = child.get_text().strip().replace('\xa0', ' ')
        if inner:
            parts.append(inner)

    result = ' '.join(parts)
    if debug:
        print(f"[DEBUG]   _get_heading_text result: '{result[:80]}'")
    return result


def convert_element(elem, debug=False):
    result = []
    for child in elem.children:
        if isinstance(child, (Comment, ProcessingInstruction)):
            continue
        if child.name is None:
            text = str(child).strip()
            if text:
                result.append(text)
        elif child.name == 'p':
            para_text = convert_element(child, debug).strip()
            if para_text:
                result.append(para_text + "\n\n")
        elif child.name == 'br':
            result.append("\n")
        elif child.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            level = int(child.name[1])
            text = _get_heading_text(child, debug)
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
            result.append(convert_table(child) + "\n")
        elif child.name in ['ul', 'ol']:
            ordered = child.name == 'ol'
            result.append(convert_list(child, ordered) + "\n")
        elif child.name == 'div':
            result.append(convert_element(child, debug))
        elif child.name == 'section':
            result.append(convert_element(child, debug))
        elif child.name == 'pre':
            code_text = child.get_text().rstrip()
            if code_text:
                result.append(f"```\n{code_text}\n```\n\n")
        elif child.name == 'code':
            text = child.get_text().strip()
            if text:
                result.append(f"`{text}`")
        elif child.name == 'img':
            src = child.get('src', '')
            alt = child.get('alt', '').strip()
            if src:
                if not alt:
                    alt = os.path.basename(src)
                result.append(f"![{alt}]({src})\n")
        elif child.name == 'span':
            text = convert_element(child, debug)
            if text.strip():
                result.append(text)
        else:
            result.append(convert_element(child, debug))
    return "".join(result)


def _clean_word_markup(soup, debug=False):
    for tag in soup(["script", "style", "meta", "link"]):
        tag.decompose()

    for op in soup.find_all('o:p'):
        op.decompose()

    for comment in soup.find_all(string=lambda text: isinstance(text, Comment)):
        comment_text = str(comment)
        if 'supportFields' in comment_text:
            _remove_conditional_block(comment)
            continue
        if 'supportLists' in comment_text:
            parent = comment.parent
            if parent and parent.name and parent.name in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
                _remove_heading_list_block(comment)
            else:
                _remove_conditional_block(comment)
            continue
        if 'endif' in comment_text:
            comment.extract()

    for pi in soup.find_all(string=lambda text: isinstance(text, ProcessingInstruction)):
        pi.extract()

    if debug:
        _debug_report(soup)


def _remove_conditional_block(start_comment):
    current = start_comment.next_sibling
    removed = []
    while current is not None:
        if isinstance(current, (Comment, ProcessingInstruction)):
            if 'endif' in str(current):
                current.extract()
                break
            else:
                next_sib = current.next_sibling
                current.extract()
                current = next_sib
                continue
        next_sib = current.next_sibling
        if current.name:
            removed.append(current.name)
            current.decompose()
        else:
            current.extract()
        current = next_sib
    start_comment.extract()


def _remove_heading_list_block(start_comment):
    from bs4 import NavigableString
    current = start_comment.next_sibling
    while current is not None:
        if isinstance(current, (Comment, ProcessingInstruction)):
            if 'endif' in str(current):
                current.extract()
                break
            else:
                next_sib = current.next_sibling
                current.extract()
                current = next_sib
                continue
        next_sib = current.next_sibling
        if current.name == 'a':
            current.decompose()
        elif current.name == 'span':
            span_style = current.get('style', '')
            if 'mso-list:Ignore' in span_style:
                text_parts = []
                for content in current.contents:
                    if isinstance(content, NavigableString):
                        text_parts.append(str(content).strip())
                numbering = ' '.join(text_parts).strip()
                if numbering:
                    current.replace_with(NavigableString(numbering + ' '))
                else:
                    current.decompose()
        current = next_sib
    start_comment.extract()


def _debug_report(soup):
    import re as _re
    headings = soup.find_all(_re.compile(r'^h[1-6]$'))
    mso_count = 0
    plain_count = 0
    for h in headings:
        style = h.get('style', '')
        if 'mso-list' in style:
            mso_count += 1
        else:
            plain_count += 1
    print(f"[DEBUG] Headings: {len(headings)} total, {mso_count} auto-numbered (mso-list), {plain_count} plain-text")

    imgs = soup.find_all('img')
    print(f"[DEBUG] Images: {len(imgs)} <img> tags found")
    missing_src = sum(1 for img in imgs if not img.get('src', ''))
    if missing_src:
        print(f"[DEBUG] WARNING: {missing_src} images have empty src attribute")

    for i, img in enumerate(imgs[:5]):
        src = img.get('src', '')
        parent = img.parent.name if img.parent else '?'
        print(f"[DEBUG]   img[{i}]: parent=<{parent}>, src_len={len(src)}")

    comments = soup.find_all(string=lambda t: isinstance(t, Comment))
    support_fields = sum(1 for c in comments if 'supportFields' in str(c))
    support_lists = sum(1 for c in comments if 'supportLists' in str(c))
    print(f"[DEBUG] Comments remaining: {len(comments)} ({support_lists} supportLists, {support_fields} supportFields)")


def _post_process(content):
    content = re.sub(r'\[if\s+!supportLists\]', '', content)
    content = re.sub(r'\[if\s+supportFields\]', '', content)
    content = re.sub(r'\[endif\]', '', content)
    content = re.sub(r'\[if\s+[^\]]*\]', '', content)
    content = re.sub(r'StartFragment|EndFragment', '', content)

    content = re.sub(r'PAGEREF\s+\S+\s+\\h', '', content)
    content = re.sub(r'mso-element:field-\w+', '', content)

    content = re.sub(r'font-family:[^;"]+;?', '', content)
    content = re.sub(r"mso-spacerun:'yes';?\s*", '', content)
    content = re.sub(r"mso-fareast-font-family:[^;\"']+;?\s*", '', content)
    content = re.sub(r"mso-ansi-font-weight:\w+;?\s*", '', content)
    content = re.sub(r"mso-bidi-font-weight:\w+;?\s*", '', content)
    content = re.sub(r"font-size:\d+\.\d+pt;?\s*", '', content)
    content = re.sub(r"font-variant:\w+;?\s*", '', content)
    content = re.sub(r"font-style:\w+;?\s*", '', content)
    content = re.sub(r"tab-stops:[^;\"']+;?\s*", '', content)
    content = re.sub(r"mso-list:[^;\"']+;?\s*", '', content)

    content = re.sub(r'<span[^>]*>\s*</span>', '', content)
    content = re.sub(r'<span[^>]*>\s*$', '', content, flags=re.MULTILINE)
    content = re.sub(r'^\s*</span>', '', content, flags=re.MULTILINE)
    content = re.sub(r'<span[^>]*>', '', content)
    content = re.sub(r'</span>', '', content)

    content = re.sub(r'<a\s+name="[^"]*">\s*</a>', '', content)
    content = re.sub(r'<a[^>]*>\s*</a>', '', content)

    content = re.sub(r'<o:p>\s*</o:p>', '', content)

    content = re.sub(r'^(\d+(?:\.\d+)*)(\*\*)', r'\1 \2', content, flags=re.MULTILINE)
    content = re.sub(r'\*\*\*\*', '', content)

    content = re.sub(r'</?[a-zA-Z][^>]*>', '', content)

    content = re.sub(r'\n{5,}', '\n\n\n\n', content)
    content = '\n'.join(line.rstrip() for line in content.split('\n'))

    return content


def html_to_markdown(html_path, md_path=None, debug=False):
    if not os.path.exists(html_path):
        print(f"错误: 文件不存在: {html_path}")
        return False

    if md_path is None:
        md_path = html_path.replace('.html', '.md').replace('.htm', '.md')

    with open(html_path, "rb") as f:
        raw = f.read()

    encodings = []
    detected = chardet.detect(raw)
    if detected and detected['encoding']:
        det_enc = detected['encoding'].lower()
        if det_enc == 'gb2312':
            det_enc = 'gb18030'
        encodings.append(det_enc)
    encodings.extend(['utf-8', 'gb18030', 'gbk', 'gb2312'])

    html = None
    used_enc = None
    for enc in encodings:
        try:
            html = raw.decode(enc)
            used_enc = enc
            break
        except (UnicodeDecodeError, LookupError):
            continue

    if html is None:
        print("错误: 无法读取文件，编码不支持")
        return False

    if debug:
        print(f"[DEBUG] File encoding: {used_enc} (detected: {detected})")
        print(f"[DEBUG] File size: {len(raw)} bytes, decoded: {len(html)} chars")

    soup = BeautifulSoup(html, 'html.parser')

    _clean_word_markup(soup, debug)

    body = soup.find('body')
    content = convert_element(body if body else soup, debug)

    if debug:
        img_count = content.count('![')
        print(f"[DEBUG] Image references in convert_element output: {img_count}")

    content = _post_process(content)

    if debug:
        img_count = content.count('![')
        print(f"[DEBUG] Image references after post-processing: {img_count}")

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"[OK] 转换完成: {md_path}")
    print(f"[OK] 文件大小: {len(content)} 字符")
    return True


def main():
    parser = argparse.ArgumentParser(description='HTML 转 Markdown v2')
    parser.add_argument('html_path', help='HTML 文件路径')
    parser.add_argument('md_path', nargs='?', default=None, help='输出 Markdown 路径')
    parser.add_argument('--debug', action='store_true', help='输出诊断信息')
    args = parser.parse_args()

    success = html_to_markdown(args.html_path, args.md_path, debug=args.debug)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
