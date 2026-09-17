"""
HTML 需求文档提取脚本 — 单文件 HTML → 正文 Markdown + 内嵌图片，支持按章节分批

用法：
    python extract_html.py --input requirements.html --info          # 先看有哪些章节
    python extract_html.py --input requirements.html --sections 2,3  # 只提这两章
    python extract_html.py --input requirements.html                 # 全文提取

输出：
    outdir/prd.md（或 prd_s2-3.md）  正文，标题/表格/列表转 Markdown，图片位置留 [[IMAGE: ...]] 锚点
    outdir/images/img_NN.ext         从 data:image base64 还原出的图片（原型图等）

为什么需要这个脚本：
    需求 HTML 往往几 MB，其中 95%+ 是 CSS 与内嵌 base64 图片，正文只有一两万字符。
    直接 Read 原文件会一次性撑爆上下文。本脚本把「正文」与「图片」拆开，
    正文一次读完，图片按 prd.md 里的锚点 **按需** Read。

关于 --sections（对应 PPT 路径的 --slides 页码）：
    HTML 没有页码，但有章节。--info 会列出章节目录，--sections 按编号挑章节，
    用法与「先做 1-10 页，再做 11-15 页」的分批节奏完全一致。
    图片编号是 **全文全局** 的，分批提取时不会互相覆盖或错位。

依赖：
    - pip: beautifulsoup4
"""

import argparse
import base64
import os
import re
import sys

INLINE_TAGS = [
    'b', 'strong', 'i', 'em', 'span', 'a', 'code', 'small',
    'sup', 'sub', 'u', 'mark', 'abbr', 'font',
]

MIME_EXT = {
    'jpeg': 'jpg',
    'svg+xml': 'svg',
}

ANCHOR_RE = re.compile(r'\[\[IMAGE: images/img_(\d+)\.')


def load_soup(path):
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        print("ERROR: 缺少依赖 beautifulsoup4。请先安装：")
        print("  pip install beautifulsoup4 --break-system-packages -q")
        sys.exit(1)

    with open(path, encoding='utf-8', errors='replace') as f:
        html = f.read()
    soup = BeautifulSoup(html, 'html.parser')
    for tag in soup(['style', 'script', 'noscript']):
        tag.decompose()
    return soup, len(html)


def parse_range(spec, total):
    """'1-3,5' → {1,2,3,5}。与 extract_prd.py 的页码语法保持一致。"""
    if not spec:
        return set(range(1, total + 1))
    result = set()
    for part in spec.split(','):
        part = part.strip().lstrip('sS§')
        m = re.match(r'^(\d+)\s*-\s*(\d+)$', part)
        if m:
            lo, hi = int(m.group(1)), int(m.group(2))
            result.update(range(max(lo, 1), min(hi, total) + 1))
        elif part.isdigit():
            n = int(part)
            if 1 <= n <= total:
                result.add(n)
    return result


# ── 分节 ──────────────────────────────────────────────────────────────────────

def section_title(nodes):
    """
    节标题 = 节内开头的前几行文本。

    刻意不取「节内第一个 h 标签」：需求 HTML 的章节标题常做成带 class 的 div
    （如 <span class="num">§2</span> 全局宽布局框架），而节内正文里反倒有
    <h4>▸ 布局规则</h4> 这类小标题——按 h 标签取会把每节都标成「▸ 布局规则」。
    """
    text = '\n'.join(node_text(n) for n in nodes)
    lines = [' '.join(l.split()) for l in text.split('\n')]
    lines = [l for l in lines if l]
    title = ''
    for line in lines[:3]:
        title = (title + ' ' + line).strip()
        if len(title) >= 6:
            break
    return title[:70] or '(无标题)'


def sections_from_headings(soup, level):
    """以 hN 为界切节：标题 + 其后续兄弟节点，直到下一个同级标题。"""
    result = []
    for h in soup.find_all('h%d' % level):
        nodes = [h]
        for sib in h.next_siblings:
            if getattr(sib, 'name', None) == h.name:
                break
            nodes.append(sib)
        result.append(nodes)
    return result


def split_sections(soup):
    """返回 (节列表, 依据说明)。每节是一个节点列表。逐级降级，拿不到就整篇一节。"""
    secs = soup.find_all('section')
    if len(secs) >= 2:
        return [[s] for s in secs], '<section> 标签'

    divs = [d for d in soup.find_all('div', id=True)
            if re.match(r'^[a-z]*\d+$', (d.get('id') or '').strip(), re.I)]
    # 只保留顶层的，避免父子 div 都被当成节
    divs = [d for d in divs if not any(d is not o and o in d.parents for o in divs)]
    if len(divs) >= 2:
        return [[d] for d in divs], 'div[id] 锚点'

    for level in range(1, 7):
        groups = sections_from_headings(soup, level)
        if len(groups) >= 2:
            return groups, 'h%d 标题' % level

    return [[soup]], None


# ── 内容转换 ──────────────────────────────────────────────────────────────────

def extract_images(soup):
    """
    全文扫描 img，按 **全局顺序** 编号并替换为锚点，返回图片清单。
    先全局编号再按节筛选，保证分批提取时编号稳定、互不覆盖。
    """
    images, external = [], []

    for tag in soup.find_all('img'):
        src = tag.get('src', '') or ''
        alt = ' '.join((tag.get('alt') or '').split())
        m = re.match(r'data:image/([\w+.-]+);base64,(.*)', src, re.S)
        if not m:
            # 外链图片：本地拿不到，如实报告，不静默吞掉
            external.append(src[:120])
            tag.replace_with(soup.new_string(
                "\n[[IMAGE-EXTERNAL: %s]]\n" % (src[:200] or '(空 src)')
            ))
            continue

        idx = len(images) + 1
        ext = MIME_EXT.get(m.group(1).lower(), m.group(1).lower())
        rel = 'images/img_%02d.%s' % (idx, ext)
        images.append({'idx': idx, 'rel': rel, 'alt': alt, 'b64': m.group(2)})
        tag.replace_with(soup.new_string(
            "\n[[IMAGE: %s]]%s\n" % (rel, ("  alt=" + alt) if alt else '')
        ))

    return images, external


def write_images(images, used_idx, outdir):
    img_dir = os.path.join(outdir, 'images')
    os.makedirs(img_dir, exist_ok=True)
    written = []
    for img in images:
        if img['idx'] not in used_idx:
            continue
        try:
            data = base64.b64decode(re.sub(r'\s+', '', img['b64']))
        except Exception as e:
            print("WARN: 第 %d 张图 base64 解码失败，已跳过：%s" % (img['idx'], e))
            continue
        with open(os.path.join(outdir, img['rel']), 'wb') as f:
            f.write(data)
        written.append(img)
    return written


def tables_to_markdown(root, soup):
    """<table> → Markdown 表格。需求文档的数值规则多半在表里，必须无损保留。"""
    count = 0
    for tb in root.find_all('table'):
        rows = []
        for tr in tb.find_all('tr'):
            cells = [
                ' '.join(td.get_text(' ', strip=True).split()).replace('|', '\\|')
                for td in tr.find_all(['td', 'th'])
            ]
            if cells:
                rows.append(cells)
        if not rows:
            tb.decompose()
            continue
        width = max(len(r) for r in rows)
        lines = []
        for i, cells in enumerate(rows):
            cells = cells + [''] * (width - len(cells))
            lines.append('| ' + ' | '.join(cells) + ' |')
            if i == 0:
                lines.append('|' + '---|' * width)
        tb.replace_with(soup.new_string('\n' + '\n'.join(lines) + '\n'))
        count += 1
    return count


def headings_to_markdown(root, soup):
    for tag in root.find_all(re.compile(r'^h[1-6]$')):
        level = int(tag.name[1])
        text = ' '.join(tag.get_text(' ', strip=True).split())
        tag.replace_with(soup.new_string('\n' + '#' * level + ' ' + text + '\n'))


def lists_to_markdown(root, soup):
    for li in root.find_all('li'):
        text = ' '.join(li.get_text(' ', strip=True).split())
        li.replace_with(soup.new_string('\n- ' + text))


def flatten_inline(root):
    """先摊平 inline 标签再取文本，否则 <b>/<span> 会把一句话切成好几行。"""
    for name in INLINE_TAGS:
        for tag in root.find_all(name):
            tag.unwrap()
    root.smooth()


def node_text(node):
    return node.get_text('\n') if hasattr(node, 'get_text') else str(node)


def render_nodes(nodes, soup):
    """一节 → Markdown 文本。"""
    n_tables = 0
    for node in nodes:
        if not hasattr(node, 'find_all'):
            continue
        n_tables += tables_to_markdown(node, soup)
        headings_to_markdown(node, soup)
        lists_to_markdown(node, soup)
        flatten_inline(node)
    text = '\n'.join(node_text(n) for n in nodes)
    text = re.sub(r'[ \t ]+', ' ', text)
    text = re.sub(r' *\n *', '\n', text)
    text = re.sub(r'\n{3,}', '\n\n', text).strip()
    return text, n_tables


# ── 主流程 ────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', required=True, help='HTML 文件路径')
    parser.add_argument('--outdir', default='html_output', help='输出目录（默认 html_output）')
    parser.add_argument('--sections', default=None,
                        help='章节编号范围，如 1-3,5（默认全文）。编号见 --info')
    parser.add_argument('--info', action='store_true', help='只列章节目录和规模，不写文件')
    args = parser.parse_args()

    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        print("ERROR: 文件不存在：%s" % input_path)
        sys.exit(1)

    soup, raw_len = load_soup(input_path)
    title_tag = soup.find('title')
    doc_title = ' '.join(title_tag.get_text(' ', strip=True).split()) if title_tag else ''

    # 图片必须在分节前全局编号，否则分批提取时编号会错位
    images, external = extract_images(soup)
    groups, basis = split_sections(soup)
    titles = [section_title(g) for g in groups]

    # ── --info：列章节目录 ────────────────────────────────────────────────
    if args.info:
        print("📊 %s" % (doc_title or os.path.basename(input_path)))
        print("   原始文件 %d 字符 | 内嵌图片 %d 张" % (raw_len, len(images)))
        if external:
            print("   ⚠️ 外链图片 %d 张（本地取不到）" % len(external))
        if basis is None:
            print("\n⚠️ 这份 HTML 没有可识别的章节结构，无法分批提取，只能整篇处理。")
            print("   正文约 %d 字符。" % len(soup.get_text(' ')))
            sys.exit(0)

        print("\n📑 章节目录（分节依据：%s，共 %d 节）：" % (basis, len(groups)))
        for i, (g, t) in enumerate(zip(groups, titles), 1):
            raw = ' '.join(node_text(n) for n in g)
            n_img = len(set(int(x) for x in ANCHOR_RE.findall(raw)))
            print("   %2d. %s" % (i, t))
            print("       约 %d 字符 | 图 %d 张" % (len(' '.join(raw.split())), n_img))
        print("\n用 --sections 挑章节，例如：--sections 1-3  或  --sections 2,5")
        sys.exit(0)

    # ── 正式提取 ──────────────────────────────────────────────────────────
    selected = sorted(parse_range(args.sections, len(groups)))
    if not selected:
        print("ERROR: --sections %s 没有匹配到任何章节（共 %d 节，先跑 --info 看目录）"
              % (args.sections, len(groups)))
        sys.exit(1)

    outdir = os.path.abspath(args.outdir)
    os.makedirs(outdir, exist_ok=True)

    parts, total_tables = [], 0
    if doc_title:
        parts.append('# ' + doc_title + '\n')
    if args.sections:
        parts.append('> 本文件只包含第 %s 节（全文共 %d 节）。其余章节未提取。\n'
                     % (','.join(str(i) for i in selected), len(groups)))

    for i in selected:
        text, n_tab = render_nodes(groups[i - 1], soup)
        total_tables += n_tab
        if len(groups) > 1:
            parts.append('\n<!-- ── 第 %d 节：%s ── -->\n' % (i, titles[i - 1]))
        parts.append(text)

    body = '\n'.join(parts)
    used_idx = set(int(x) for x in ANCHOR_RE.findall(body))
    written = write_images(images, used_idx, outdir)

    if written:
        parts.append('\n\n---\n\n## 图片清单（按需 Read，不要全部读）\n')
        for img in written:
            parts.append('- `%s` — %s' % (img['rel'], img['alt'] or '(无说明)'))
        body = '\n'.join(parts)

    suffix = ''
    if args.sections and len(selected) < len(groups):
        suffix = '_s' + '-'.join(str(i) for i in selected)
    md_path = os.path.join(outdir, 'prd%s.md' % suffix)
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(body + '\n')

    print("✅ 提取完成：%s" % (doc_title or os.path.basename(input_path)))
    if len(groups) > 1:
        print("   章节：%s / 共 %d 节（依据：%s）"
              % (','.join(str(i) for i in selected), len(groups), basis))
    print("   原始 %d 字符 → 正文 %d 字符" % (raw_len, len(body)))
    print("📄 正文：%s" % md_path)
    print("🖼  图片：%d 张（本次章节内）→ %s" % (len(written), os.path.join(outdir, 'images')))
    print("📋 表格：%d 个（已转 Markdown，数值无损）" % total_tables)

    # 抽取质量告警：宁可吵，也不要静默产出空壳让后续分析踩空
    if len(body) < 500:
        print("\n⚠️ 正文过短（<500 字符），这份 HTML 可能是 JS 动态渲染的，"
              "静态解析拿不到内容。请改用整页截图，或让需求方导出静态版。")
    if total_tables == 0:
        print("⚠️ 未解析到任何表格。若原文档肉眼可见表格，可能是 div 伪表格"
              "（Notion / 飞书导出常见），正文里的数值排布需人工复核。")
    if not written:
        print("⚠️ 本次提取的章节里没有图片。若预期有原型图，确认是否选错章节。")
    if external:
        print("⚠️ 有 %d 张外链图片未取回（正文中标为 [[IMAGE-EXTERNAL]]）：" % len(external))
        for u in external[:5]:
            print("     - %s" % u)


if __name__ == '__main__':
    main()
