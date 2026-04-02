"""
Test Case Generator - 标准五列格式专用脚本
列：模块 | 用例名称 | 描述 | 预期 | 备注

用法：
    python generate.py --input cases.csv --output output.xlsx --changes changes.json

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
import shutil
import zipfile

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


COLUMNS = ['模块', '用例名称', '描述', '预期', '备注']
COL_WIDTHS = {'A': 18, 'B': 16, 'C': 50, 'D': 45, 'E': 14}
HEADER_COLOR = '2F4F4F'
DEPRECATED_NOTE = '已废弃'
NEW_ROW_MARKER = '__is_new__'   # 内部标记列，不写入 xlsx


# ── 富文本修复 ──────────────────────────────────────────────

def fix_rich_text_xlsx(filepath):
    """修复 openpyxl 写入富文本后的颜色透明和字号错误问题"""
    tmp = filepath + '.tmp'
    shutil.copy(filepath, tmp)

    with zipfile.ZipFile(tmp, 'r') as z:
        sheet_bytes = z.read('xl/worksheets/sheet1.xml')
        all_files = {n: z.read(n) for n in z.namelist()}

    fixed = sheet_bytes.replace(b'rgb="00000000"', b'rgb="FF000000"')
    fixed = fixed.replace(b'rgb="00FF0000"', b'rgb="FFFF0000"')
    fixed = fixed.replace(b'<sz val="1000"/>', b'<sz val="10"/>')

    all_files['xl/worksheets/sheet1.xml'] = fixed

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
        color = 'FF0000' if run['red'] else '000000'
        ifont = InlineFont(rFont='Arial', sz=1000, color=color)
        blocks.append(TextBlock(ifont, run['text']))
    cell.value = CellRichText(*blocks)
    cell.alignment = Alignment(wrap_text=True, vertical='top')


# ── 样式工具 ────────────────────────────────────────────────

def style_header(ws):
    fill = PatternFill('solid', start_color=HEADER_COLOR)
    font = Font(bold=True, color='FFFFFF', name='Arial', size=10)
    border = make_border()
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def make_border():
    thin = Side(style='thin', color='AAAAAA')
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def style_data_cell(cell, red=False):
    cell.font = Font(name='Arial', size=10, color='FF0000' if red else '000000')
    cell.alignment = Alignment(wrap_text=True, vertical='top')
    cell.border = make_border()


# ── 主流程 ──────────────────────────────────────────────────

def load_csv(path):
    df = pd.read_csv(path, dtype=str, encoding='utf-8').fillna('')
    missing = [c for c in COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"CSV 缺少列：{missing}，请使用临时脚本模式处理非标准格式。")
    return df[COLUMNS].copy()


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


def build_xlsx(df, changes, output_path):
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

            if key in modified_map:
                # 修改行：黑红混排富文本
                make_rich_cell(cell, modified_map[key])
                cell.border = make_border()
            elif is_new:
                # 新增行：整行红色
                style_data_cell(cell, red=True)
            else:
                # 原有行：黑色，废弃行在备注列追加标记
                if orig_csv_counter in deprecated_rows and col_letter == 'E':
                    orig_val = cell.value or ''
                    cell.value = (orig_val + ' ' + DEPRECATED_NOTE).strip()
                style_data_cell(cell, red=False)

        if not is_new:
            orig_csv_counter += 1

        ws.row_dimensions[excel_row].height = 60

    wb.save(output_path)
    fix_rich_text_xlsx(output_path)
    print(f"✅ 已生成：{output_path}")


# ── 入口 ────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', required=True, help='输入 CSV 路径')
    parser.add_argument('--output', required=True, help='输出 xlsx 路径')
    parser.add_argument('--changes', required=True, help='变更 JSON 路径')
    args = parser.parse_args()

    df = load_csv(args.input)
    with open(args.changes, encoding='utf-8') as f:
        changes = json.load(f)

    build_xlsx(df, changes, args.output)


if __name__ == '__main__':
    main()
