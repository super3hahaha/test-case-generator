"""
Test Case Generator - 标准五列格式专用脚本
列：模块 | 用例名称 | 描述 | 预期 | 备注

支持输入格式：CSV（UTF-8）、HTML（Google Sheets 导出）

用法：
    python generate.py --input cases.csv --output output.xlsx --changes changes.json
    python generate.py --input cases.html --output output.xlsx --changes changes.json

changes.json 格式：
{
  "modified": [
    {
      "row": 2,
      "col": "C",
      "runs": [
        {"text": "1. 原有步骤\n", "red": false},
        {"text": "2. 新增步骤\n", "red": true},
        {"text": "3. 原有步骤续", "red": false}
      ]
    }
  ],
  "new_rows": [
    {
      "after_module": "模块名称",
      "data": {"模块": "Go Premium入口", "用例名称": "会员入口文案", "描述": "...", "预期": "...", "备注": ""}
    }
  ],
  "deprecated": [2, 5]
}

注意：
- new_row_excel_rows 字段已废弃，脚本自动追踪新增行位置，无需手动传入
- deprecated 为 CSV 数据行索引（0-based）
"""

import argparse
import json
import os
import re
import shutil
import zipfile

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Font, PatternFill


COLUMNS = ['模块', '用例名称', '描述', '预期', '备注']
COL_WIDTHS = {'A': 18, 'B': 16, 'C': 50, 'D': 45, 'E': 14}
HEADER_COLOR = '2F4F4F'
RED_COLOR = 'EA4335'
DEPRECATED_NOTE = '已废弃'
NEW_ROW_MARKER = '__is_new__'   # 内部标记列，不写入 xlsx


# ── 富文本修复 ──────────────────────────────────────────────

def fix_rich_text_xlsx(filepath):
    """修复 openpyxl 写入富文本后的颜色透明和字号错误问题"""
    tmp = filepath + '.tmp'
    shutil.copy(filepath, tmp)

    with zipfile.ZipFile(tmp, 'r') as z:
        all_files = {n: z.read(n) for n in z.namelist()}

    _alpha_re = rb'rgb="00([0-9A-Fa-f]{6})"'
    _alpha_fix = lambda m: b'rgb="FF' + m.group(1).upper() + b'"'

    for target in ('xl/worksheets/sheet1.xml', 'xl/styles.xml'):
        if target in all_files:
            data = re.sub(_alpha_re, _alpha_fix, all_files[target])
            if target == 'xl/worksheets/sheet1.xml':
                data = data.replace(b'<sz val="1000"/>', b'<sz val="10"/>')
            all_files[target] = data

    with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)

    os.remove(tmp)


# ── 富文本写入 ──────────────────────────────────────────────

def make_rich_cell(cell, runs):
    """
    runs: list of {"text": str, "red": bool}
    写入富文本（黑红混排），需配合 fix_rich_text_xlsx 修复颜色
    """
    blocks = []
    for run in runs:
        color = RED_COLOR if run['red'] else '000000'
        ifont = InlineFont(rFont='Arial', sz=1000, color=color)
        blocks.append(TextBlock(ifont, run['text']))
    cell.value = CellRichText(*blocks)
    cell.alignment = Alignment(wrap_text=True, vertical='top')


def make_rich_cell_with_orig_format(cell, runs, orig_meta):
    """Write modified rich text, preserving original formatting for non-red runs.

    For red runs: always red color.
    For non-red runs: match text against original segments to restore
    strikethrough/color from the HTML source.
    """
    orig_segments = orig_meta.get('segments', [])
    bg_color = orig_meta.get('bg_color')

    if not any(seg.get('strikethrough') or seg.get('color') for seg in orig_segments):
        make_rich_cell(cell, runs)
        if bg_color:
            cell.fill = PatternFill('solid', start_color=bg_color)
        return

    orig_text = ''.join(seg['text'] for seg in orig_segments)
    char_fmts = []
    for seg in orig_segments:
        fmt = {'strikethrough': seg.get('strikethrough', False), 'color': seg.get('color')}
        char_fmts.extend([fmt] * len(seg['text']))

    blocks = []
    search_pos = 0

    for run in runs:
        if run['red']:
            ifont = InlineFont(rFont='Arial', sz=1000, color=RED_COLOR)
            blocks.append(TextBlock(ifont, run['text']))
        else:
            run_text = run['text']
            match_pos = orig_text.find(run_text, search_pos)
            if match_pos == -1:
                match_pos = orig_text.find(run_text)

            if match_pos >= 0 and match_pos + len(run_text) <= len(char_fmts):
                i = 0
                while i < len(run_text):
                    cur_fmt = char_fmts[match_pos + i]
                    j = i + 1
                    while j < len(run_text) and char_fmts[match_pos + j] == cur_fmt:
                        j += 1
                    color = cur_fmt.get('color') or '000000'
                    strike = cur_fmt.get('strikethrough', False)
                    ifont = InlineFont(
                        rFont='Arial', sz=1000, color=color,
                        strike=True if strike else None,
                    )
                    blocks.append(TextBlock(ifont, run_text[i:j]))
                    i = j
                search_pos = match_pos + len(run_text)
            else:
                ifont = InlineFont(rFont='Arial', sz=1000, color='000000')
                blocks.append(TextBlock(ifont, run['text']))

    if blocks:
        cell.value = CellRichText(*blocks)
    cell.alignment = Alignment(wrap_text=True, vertical='top')

    if bg_color:
        cell.fill = PatternFill('solid', start_color=bg_color)


# ── 样式工具 ────────────────────────────────────────────────

def style_header(ws):
    fill = PatternFill('solid', start_color=HEADER_COLOR)
    font = Font(bold=True, color='FFFFFF', name='Arial', size=10)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def style_data_cell(cell, red=False):
    cell.font = Font(name='Arial', size=10, color=RED_COLOR if red else '000000')
    cell.alignment = Alignment(wrap_text=True, vertical='top')


def _apply_original_format(cell, cell_meta):
    """Apply original HTML formatting (strikethrough/color/bg) to an unmodified cell."""
    segments = cell_meta.get('segments', [])
    bg_color = cell_meta.get('bg_color')

    has_rich = any(seg.get('strikethrough') or seg.get('color') for seg in segments)

    if has_rich and segments:
        blocks = []
        for seg in segments:
            color = seg.get('color') or '000000'
            strike = seg.get('strikethrough', False)
            ifont = InlineFont(
                rFont='Arial', sz=1000, color=color,
                strike=True if strike else None,
            )
            blocks.append(TextBlock(ifont, seg['text']))
        cell.value = CellRichText(*blocks)
        cell.alignment = Alignment(wrap_text=True, vertical='top')
    else:
        style_data_cell(cell, red=False)

    if bg_color:
        cell.fill = PatternFill('solid', start_color=bg_color)


# ── 主流程 ──────────────────────────────────────────────────

def load_csv(path):
    df = pd.read_csv(path, dtype=str, encoding='utf-8').fillna('')
    missing = [c for c in COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"CSV 缺少列：{missing}，请使用临时脚本模式处理非标准格式。")
    return df[COLUMNS].copy()


# ── HTML 格式提取 ─────────────────────────────────────────────

def _parse_css_classes(soup):
    """Extract formatting info from <style> CSS class definitions."""
    css_map = {}
    style_tag = soup.find('style')
    if not style_tag or not style_tag.string:
        return css_map

    for match in re.finditer(r'\.(s\d+)\s*\{([^}]+)\}', style_tag.string):
        cls_name = match.group(1)
        props_str = match.group(2)
        fmt = {}

        for prop in props_str.split(';'):
            prop = prop.strip()
            if ':' not in prop:
                continue
            name, value = prop.split(':', 1)
            name = name.strip().lower()
            value = value.strip().lower()

            if name == 'text-decoration' and 'line-through' in value:
                fmt['strikethrough'] = True
            elif name == 'color' and value.startswith('#') and len(value) == 7:
                color = value[1:].upper()
                if color != '000000':
                    fmt['color'] = color
            elif name == 'background-color' and value.startswith('#') and len(value) == 7:
                bg = value[1:].upper()
                if bg != 'FFFFFF':
                    fmt['bg_color'] = bg

        if fmt:
            css_map[cls_name] = fmt

    return css_map


def _walk_children(element, inherited_fmt, segments):
    """Recursively walk HTML element children, collecting text segments with formatting."""
    from bs4 import NavigableString, Tag

    for child in element.children:
        if isinstance(child, NavigableString):
            text = str(child)
            if text:
                segments.append({
                    'text': text,
                    'strikethrough': inherited_fmt.get('strikethrough', False),
                    'color': inherited_fmt.get('color'),
                })
        elif isinstance(child, Tag):
            if child.name == 'br':
                segments.append({
                    'text': '\n',
                    'strikethrough': inherited_fmt.get('strikethrough', False),
                    'color': inherited_fmt.get('color'),
                })
            else:
                child_fmt = dict(inherited_fmt)
                style = child.get('style', '')
                if style:
                    for prop in style.split(';'):
                        prop = prop.strip()
                        if ':' not in prop:
                            continue
                        name, value = prop.split(':', 1)
                        name = name.strip().lower()
                        value = value.strip().lower()
                        if name == 'text-decoration' and 'line-through' in value:
                            child_fmt['strikethrough'] = True
                        elif name == 'color' and value.startswith('#') and len(value) == 7:
                            color = value[1:].upper()
                            if color != '000000':
                                child_fmt['color'] = color
                            else:
                                child_fmt.pop('color', None)
                _walk_children(child, child_fmt, segments)


def _merge_segments(segments):
    """Merge adjacent segments with identical formatting."""
    if not segments:
        return []
    merged = [dict(segments[0])]
    for seg in segments[1:]:
        prev = merged[-1]
        if (prev.get('strikethrough', False) == seg.get('strikethrough', False)
                and prev.get('color') == seg.get('color')):
            prev['text'] += seg['text']
        else:
            merged.append(dict(seg))
    return merged


def _extract_cell_rich(cell, css_map):
    """Extract text and rich formatting from an HTML cell.

    Returns:
        (plain_text, bg_color, segments)
        - plain_text: str
        - bg_color: str (hex like 'EAD1DC') or None
        - segments: list of {'text', 'strikethrough', 'color'}
    """
    cell_fmt = {}
    bg_color = None
    for cls in cell.get('class', []):
        if cls in css_map:
            class_fmt = css_map[cls]
            if 'bg_color' in class_fmt:
                bg_color = class_fmt['bg_color']
            if 'strikethrough' in class_fmt:
                cell_fmt['strikethrough'] = True
            if 'color' in class_fmt:
                cell_fmt['color'] = class_fmt['color']

    segments = []
    _walk_children(cell, cell_fmt, segments)
    merged = _merge_segments(segments)

    if merged:
        merged[0] = dict(merged[0])
        merged[0]['text'] = merged[0]['text'].lstrip()
        if not merged[0]['text']:
            merged.pop(0)
    if merged:
        merged[-1] = dict(merged[-1])
        merged[-1]['text'] = merged[-1]['text'].rstrip()
        if not merged[-1]['text']:
            merged.pop()

    plain_text = ''.join(seg['text'] for seg in merged)
    return plain_text, bg_color, merged


# ── 加载输入 ──────────────────────────────────────────────────

def load_html(path):
    """Parse Google Sheets HTML export into DataFrame with format metadata.

    Returns:
        (df, format_meta) where format_meta maps (row_idx, col_name) to
        {'bg_color': str|None, 'segments': list} for cells with formatting.
    """
    from bs4 import BeautifulSoup

    with open(path, encoding='utf-8') as f:
        html_content = f.read()

    soup = BeautifulSoup(html_content, 'html.parser')
    css_map = _parse_css_classes(soup)

    table = soup.find('table')
    if not table:
        raise ValueError("HTML 文件中未找到表格。")

    tbody = table.find('tbody')
    if not tbody:
        raise ValueError("HTML 表格中未找到 tbody。")

    rows = tbody.find_all('tr')
    if not rows:
        raise ValueError("HTML 表格中没有数据行。")

    header_cells = rows[0].find_all('td')
    headers = [cell.get_text(strip=True) for cell in header_cells]

    data_rows = rows[1:]
    num_cols = len(headers)
    rowspan_remaining = {}

    grid = []
    format_meta = {}

    for row_idx, tr in enumerate(data_rows):
        cells = tr.find_all('td')
        row_data = []
        cell_idx = 0

        for col_idx in range(num_cols):
            if col_idx in rowspan_remaining and rowspan_remaining[col_idx] > 0:
                row_data.append('')
                rowspan_remaining[col_idx] -= 1
            else:
                if cell_idx < len(cells):
                    td = cells[cell_idx]
                    plain_text, bg_color, segments = _extract_cell_rich(td, css_map)

                    rowspan = int(td.get('rowspan', 1))
                    if rowspan > 1:
                        rowspan_remaining[col_idx] = rowspan - 1

                    row_data.append(plain_text)

                    if col_idx < len(headers):
                        col_name = headers[col_idx]
                        if col_name in COLUMNS:
                            has_fmt = bg_color or any(
                                seg.get('strikethrough') or seg.get('color')
                                for seg in segments
                            )
                            if has_fmt:
                                format_meta[(row_idx, col_name)] = {
                                    'bg_color': bg_color,
                                    'segments': segments,
                                }

                    cell_idx += 1
                else:
                    row_data.append('')

        grid.append(row_data)

    df = pd.DataFrame(grid, columns=headers).fillna('')
    missing = [c for c in COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"HTML 缺少列：{missing}，请确认表格包含标准五列。")
    return df[COLUMNS].copy(), format_meta


def load_input(path):
    """Auto-detect file format and load. Returns (df, format_meta)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.html', '.htm'):
        return load_html(path)
    return load_csv(path), None


def insert_new_rows(df, new_rows):
    """
    将新增行插入到对应模块最后一行的紧下方。
    在 df 中用内部标记列 __is_new__ 标识新增行，供后续标红使用。
    """
    df = df.copy()
    df[NEW_ROW_MARKER] = False

    for item in new_rows:
        module = item['after_module']
        row_data = item['data']

        # 还原合并单元格：把空的模块列向下填充后找最后一行
        df_filled = df.copy()
        df_filled['模块'] = df_filled['模块'].replace('', None).ffill()
        mask = df_filled['模块'] == module

        insert_pos = df_filled[mask].index[-1] + 1 if mask.any() else len(df)

        new_row = {c: row_data.get(c, '') for c in COLUMNS}
        new_row[NEW_ROW_MARKER] = True   # 打标记
        new_df = pd.DataFrame([new_row])

        df = pd.concat([df.iloc[:insert_pos], new_df, df.iloc[insert_pos:]], ignore_index=True)

    return df


def build_xlsx(df, changes, output_path, format_meta=None):
    deprecated_rows = set(changes.get('deprecated', []))  # CSV 原始 0-based 行号

    # 插入新增行（自动打标记，无需外部传入行号）
    df = insert_new_rows(df, changes.get('new_rows', []))

    # 建立修改行的富文本映射 (excel_row, col_letter) -> runs
    modified_map = {}
    for mod in changes.get('modified', []):
        modified_map[(mod['row'], mod['col'])] = mod['runs']

    wb = Workbook()
    ws = wb.active

    # 列宽
    for col, width in COL_WIDTHS.items():
        ws.column_dimensions[col].width = width

    # 表头
    ws.append(COLUMNS)
    style_header(ws)

    # 数据行
    # deprecated 指向 CSV 原始行，插入新行后行号偏移，需要在插入前记录原始 df_idx
    # 方案：在插入前给原始行加 _orig_idx 标记
    # 但 insert_new_rows 已经重置 index，所以这里直接用 NEW_ROW_MARKER 判断是否新增行，
    # deprecated 仍按原始 CSV 顺序（跳过新增行）计数
    orig_csv_counter = 0  # 仅计原始行的计数器（跳过新增行）

    for df_idx, row in df.iterrows():
        excel_row = df_idx + 2  # 1=header
        is_new = bool(row.get(NEW_ROW_MARKER, False))

        ws.append([row[c] for c in COLUMNS])

        for col_idx, col_letter in enumerate(['A', 'B', 'C', 'D', 'E'], start=1):
            cell = ws.cell(row=excel_row, column=col_idx)
            key = (excel_row, col_letter)
            col_name = COLUMNS[col_idx - 1]
            fmt_key = (orig_csv_counter, col_name)
            has_orig_fmt = format_meta and fmt_key in format_meta

            if key in modified_map:
                if has_orig_fmt:
                    make_rich_cell_with_orig_format(
                        cell, modified_map[key], format_meta[fmt_key])
                else:
                    make_rich_cell(cell, modified_map[key])
            elif is_new:
                style_data_cell(cell, red=True)
            else:
                is_deprecated_note = (
                    orig_csv_counter in deprecated_rows and col_letter == 'E')

                if is_deprecated_note:
                    orig_val = cell.value or ''
                    if has_orig_fmt:
                        meta = format_meta[fmt_key]
                        segs = list(meta.get('segments', []))
                        segs.append({
                            'text': ' ' + DEPRECATED_NOTE,
                            'strikethrough': False, 'color': None,
                        })
                        _apply_original_format(
                            cell, {'bg_color': meta.get('bg_color'), 'segments': segs})
                    else:
                        cell.value = (orig_val + ' ' + DEPRECATED_NOTE).strip()
                        style_data_cell(cell, red=False)
                elif has_orig_fmt:
                    _apply_original_format(cell, format_meta[fmt_key])
                else:
                    style_data_cell(cell, red=False)

        if not is_new:
            orig_csv_counter += 1

        ws.row_dimensions[excel_row].height = 60

    wb.save(output_path)
    fix_rich_text_xlsx(output_path)
    print(f"Done: {output_path}")


# ── 入口 ────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', required=True, help='输入 CSV 或 HTML 路径')
    parser.add_argument('--output', required=True, help='输出 xlsx 路径')
    parser.add_argument('--changes', required=True, help='变更 JSON 路径')
    args = parser.parse_args()

    df, format_meta = load_input(args.input)
    with open(args.changes, encoding='utf-8') as f:
        changes = json.load(f)

    build_xlsx(df, changes, args.output, format_meta)


if __name__ == '__main__':
    main()
