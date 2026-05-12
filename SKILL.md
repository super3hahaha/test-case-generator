---
name: test-case-generator
description: >
  Use this skill whenever the user provides an existing test case CSV and a new requirements screenshot/description/pptx,
  and wants to generate an updated Excel (.xlsx) test case file. The skill covers the full workflow:
  analyzing existing cases, comparing with new requirements, classifying changes (new / modified / deprecated),
  and outputting a formatted .xlsx where new or modified content is highlighted in red while original content
  stays black. Trigger whenever the user mentions "测试用例", "用例更新", "需求变更", "标红新增", or uploads a
  CSV alongside a requirements image/pptx asking for an Excel output.
---

# Test Case Generator: CSV + 需求截图/PPTX → 标红 xlsx

**当前版本：v1.5.0**

## ⚠️ 关键规则

### 1. 禁止自行编写 xlsx 生成脚本

**检测到标准五列格式时，严禁从头编写 Python 脚本。** 必须直接使用本技能内置的 `scripts/generate.py`，你的工作只是构造 `changes.json` 然后调用它。自行编写脚本 = 错误执行。

### 2. 输入用例必须是 CSV，拒绝 xlsx 输入

**当用户提供 `.xlsx` / `.xls` 文件作为已有测试用例输入时，禁止直接读取。** 必须提醒用户转为 CSV 后再提供：

> "xlsx 文件包含多 sheet、合并单元格、格式信息等，直接解析不可控且浪费 token。请先将用例 sheet 另存为 CSV（UTF-8 编码）后再提供，我会用 CSV 作为输入。"

只有拿到 CSV 文件后才继续执行后续步骤。

### 3. 禁止捏造信息，不确定就问用户

**严禁编造用户未提供的信息。** 当用例内容涉及以下情况时，必须停下来询问用户，不得自行推测或填充：

- **版本号 / 版本说明**：CSV 备注列含版本信息（如"2.4.7版本-广告重构"），但用户未说明本次版本号
- **具体数值 / 价格 / 时间**：需求中未明确的金额、天数、百分比等
- **业务术语 / 专有名词**：无法从需求文档或已有用例中确认含义的术语
- **功能逻辑不明确**：需求描述模糊，存在多种理解方式

**处理方式：**
1. 在 chat 中列出不确定的具体内容
2. 向用户提问，说明需要什么信息
3. **等待用户回复后再继续**，不得跳过或用占位符替代

> 原则：宁可多问一句，不可编造一字。输出的每一条用例内容都必须有明确依据——来自需求文档、已有 CSV、或用户的回复。

### 4. modified 条目必须用 module+case 定位，严禁使用 row 行号

**这是最容易出错的规则，必须严格遵守。**

`generate.py` 处理 `new_rows` 时会向 DataFrame 插入新行并重置 index，导致之后所有行的 Excel 行号发生偏移。如果 `modified` 里写死 `row` 数字，插入新行后必然命中错误的行，造成"修改没有生效"的 bug——**而且不会报错，极难排查**。

```python
# ✅ 正确：用 module+case 定位，generate.py 插入新行后自动重新查找，永远不偏移
{"module": "撤销/重做", "case": "重做（Redo）", "col": "C", "runs": [...]}

# ❌ 严禁：写死行号，只要有 new_rows 插入就必然偏移出错
{"row": 47, "col": "C", "runs": [...]}
```

`generate.py` 在完成所有 `new_rows` 插入后，会自动构建 `(module填充值, case) → excel_row` 查找表，`modified` 条目按此表精确定位，无论插入多少新行、插在哪里都不受影响。

### 5. changes.json 必须用 Python dict + json.dump() 生成，禁止手写 JSON 文本

直接手写 JSON 文件时，字符串内的中文引号（如 `"Cutter"`）、反斜杠等特殊字符会破坏 JSON 语法，导致解析报错。

```python
# ✅ 正确：始终用 Python dict 构建数据，json.dump() 负责所有转义
import json
changes = {
    "modified": [
        {"module": "裁剪页", "case": "UI", "col": "C", "runs": [
            {"text": '文案"Cutter"（删除文件名）', "red": True}
        ]}
    ],
    "new_rows": [],
    "deprecated": []
}
with open('changes.json', 'w', encoding='utf-8') as f:
    json.dump(changes, f, ensure_ascii=False, indent=2)

# ❌ 禁止：手写 JSON 文本，中文引号、特殊字符会破坏语法
# create_file("changes.json", '{"text": "文案"Cutter"", "red": true}')
```

---

## 整体流程

```
0. 【可选】若需求来源为 PPTX，先运行提取脚本（见 Step 0）
1. 读取 CSV，检测列格式（bash 执行）
2. 逐条阅读 CSV 全部用例内容，理解已有用例覆盖的细节
3. 分析新需求（截图/文字描述/PPTX 提取内容），结合 CSV 已有内容对比补充
3.5 覆盖率自检，确保所有需求功能点都有用例覆盖
4. 分类变更，用 Python dict 构造 changes.json（标准格式）或准备用例数据（非标准格式）
4.5 回溯 3.5 核对表，验证每个功能点在最终产物中有对应落地（标准 & 非标准均执行）
5. 【标准格式】直接调用 scripts/generate.py 生成 xlsx ← 唯一正确路径
   【非标准格式】临时生成脚本（见 Step 6）
```

---

## Step 0：PPTX 需求文档预处理（可选）

当用户提供 `.pptx` 文件作为需求来源时，按以下流程处理：

### Step 0.1：查询页数，确认范围

**用户未指定页码时，禁止直接全量提取。** 必须先查询页数并询问用户：

```bash
pip install python-pptx PyMuPDF --break-system-packages -q

python /mnt/skills/user/test-case-generator/scripts/extract_pptx.py \
  --input <pptx路径> --info
```

拿到页数后，**必须询问用户**：
> "这个 PPTX 共 N 页，请问需要分析哪几页？（如 1-3,5）全部提取会消耗较多 token，建议只选需求相关的页。"

如果用户已明确指定了页码范围（如"帮我看第 3-7 页"），可跳过询问直接进入 Step 0.2。

### Step 0.2：导出为整页截图

```bash
python /mnt/skills/user/test-case-generator/scripts/extract_pptx.py \
  --input <pptx路径> \
  --slides <用户指定的页码范围> \
  --outdir pptx_output
```

**系统依赖：** 需安装 [LibreOffice](https://www.libreoffice.org/download/)（用于 pptx → pdf 转换）

**参数说明：**
- `--slides`：页码范围（如 `1-3,5,7-9`）
- `--outdir`：图片输出目录，默认 `pptx_output`
- `--dpi`：导出分辨率，默认 200

**提取后的处理方式：**
脚本将每页幻灯片导出为一张完整的 PNG 图片（包含文字、原型图、排版布局），效果等同于手动截图。导出后**必须用 Read 工具逐张查看这些图片**，然后进入 Step 3 的需求分析。

---

## Step 1：检测 CSV 格式，选择处理路径

**用 bash 读取列名，立即判断走哪条路：**

```bash
python3 -c "
import pandas as pd
df = pd.read_csv('path/to/cases.csv', dtype=str, encoding='utf-8', nrows=0)
cols = df.columns.tolist()
standard = {'模块','用例名称','描述','预期','备注'}
print('STANDARD' if standard.issubset(set(cols)) else 'CUSTOM')
print(cols)
"
```

**执行后必须在 chat 中输出检测结果，例如：**
> 📂 CSV 格式检测：STANDARD，共 N 条用例，列名：模块、用例名称、描述、预期、备注

### ✅ 输出 STANDARD → 强制走固定脚本路径

列名包含以下五项即为标准格式（顺序不限，允许有额外列）：

| 列名 | 说明 |
|------|------|
| 模块 | 功能模块，子用例合并单元格（空值表示同上）|
| 用例名称 | 测试类型：界面检查 / 页面跳转 / 按钮状态 等 |
| 描述 | 测试步骤，编号列举 |
| 预期 | 期望结果，对应步骤编号 |
| 备注 | 补充说明 |

**→ 必须执行 Step 4 → Step 5，禁止跳过，禁止自己写脚本。**

### ⚠️ 输出 CUSTOM → 走临时脚本路径

列名不匹配 → 根据实际列名临时生成 Python 脚本，参考 Step 6。

---

## Step 2：逐条阅读 CSV，理解已有用例内容

> ⚠️ **CSV 不仅是格式模板，已有用例的内容也是新需求分析的重要参考和补充。**

在分析新需求之前，必须先完整阅读 CSV 中每一条用例的内容，理解已有用例覆盖了哪些细节：

1. **逐条读取** CSV 中所有用例（模块、用例名称、描述、预期）
2. **理解已有细节**：CSV 中的用例可能包含需求文档/截图中未明确提到的具体信息（如具体的操作项、状态值、边界条件等）
3. **作为新需求的补充**：当新需求只描述了大方向但缺少细节时，已有用例中的相关内容应作为补充纳入新生成的用例中

**执行后必须在 chat 中输出 CSV 内容摘要，例如：**
> 📖 CSV 已有用例概览：
> - 模块A（N条）：用例1、用例2、...
> - 模块B（N条）：用例1、用例2、...
> - CSV 中包含但需求文档未提及的内容：xxx

**原则：CSV 中已有的用例是新需求分析的参考依据。新需求文档/截图未涉及的细节，如果 CSV 已有用例中包含，应视为有效需求保留或融入新用例。不得因为新需求未提及就丢弃 CSV 中已有的有效内容。**

---

## Step 3：分析新需求，分类变更

在已理解 CSV 已有用例内容的前提下，结合新需求与已有用例进行对比分析，输出三类变更。

> ⚠️ **需求截图必须同时分析文字描述和原型图两部分，缺一不可。**
>
> - **文字描述**：功能列表、交互逻辑、状态说明等
> - **原型图**：UI 布局、图标样式、元素排列、按钮状态、视觉层级等
>
> 原型图中可能包含文字未提及的细节（如具体的图标类型、元素位置关系、不同状态下的视觉变化），这些同样需要转化为测试用例。分析时应逐个 UI 元素检查，不要只依赖文字部分。
>
> **Before/After 对比**：当需求截图包含改版前后的原型图时，必须逐元素对比：
> - **新增元素**：After 有但 Before 没有 → 新增用例覆盖
> - **移除元素**：Before 有但 After 没有 → 在相关用例中标注移除（见下方处理方式）
> - **变化元素**：样式、位置、文案等发生变化 → 修改对应用例

**执行后必须在 chat 中分别输出从文字和原型图中提取的信息，例如：**
> 📝 文字描述提取：
> 1. 功能点A ...
> 2. 功能点B ...
>
> 🎨 原型图提取：
> 1. UI元素A（图标类型、位置）...
> 2. 状态变化B ...
> 3. 文字中未提及但原型图展示的细节C ...

### 🔄 引用性需求的处理（"与线上一致"/"与现有逻辑相同"等）

当需求描述中出现引用现有行为的表述（如"与线上版本一致"、"表现与现有逻辑相同"、"保持不变"等），**必须回到 CSV 中查找该模块的已有用例，然后询问用户如何处理**。

**处理步骤：**
1. 定位 CSV 中该模块的所有已有用例
2. 在 chat 中**列出找到的已有用例概要**，并询问用户：

> "需求提到「xxx与线上版本一致」，CSV 中该模块已有以下用例：
> - 用例1：xxx
> - 用例2：xxx
>
> 请问这部分用例要怎么处理？例如：
> 1. 参考旧用例的详细步骤，融入到新用例中
> 2. 不展开写，备注里简单注明「参考xx模块」即可
> 3. 其他（请说明）"

3. **等待用户回复后**，按用户指示处理这些用例
4. 不得自行决定保留、删除或修改这些用例

### ✏️ 修改用例

原有用例逻辑发生变化，找到对应用例，修改描述或预期。
- 记录：哪个模块（`module`）、哪条用例（`case`）、哪列（C/D/E）、原文、新增内容

> ⚠️ **严格禁止删除原有步骤**：修改某一行时，必须保留该单元格内**所有原有步骤**，不得删除或省略任何未发生变更的内容。只将新增或变更的部分插入到正确位置并标红，其余保持原样（黑色）。

#### 🗑️ 元素移除的处理方式

当 Before/After 对比发现某个 UI 元素被移除时，**不要删除原有用例中的相关描述**，而是在原文后追加标红的「（已移除）」标注。

示例：

```
描述列原文：检查顶部 pro 图标以黄色皇冠显示
修改后：    检查顶部 pro 图标以黄色皇冠显示（已移除）  ← 「（已移除）」为红色
```

对应 changes.json 写法（用 module+case，不用 row）：

```python
{
    "module": "裁剪页", "case": "UI", "col": "C",
    "runs": [
        {"text": "1. 检查顶部 pro 图标以黄色皇冠显示", "red": False},
        {"text": "（已移除）", "red": True}
    ]
}
```

### ➕ 新增用例

新需求中未被任何原有用例覆盖的逻辑，包括：
- 正向路径（happy path）
- 异常/边界分支
- 不同状态/权限下的行为差异

> ⚠️ **新增行的插入位置**：新增用例必须插入到其所属模块在表格中最后一行的**紧下方**，不得追加到整个表格末尾。

#### 📝 描述与预期的写作风格

**描述列**：写操作步骤，动词用"查看"、"检查"、"点击"、"输入"等，不用"是否为"、"验证是否"等判断句式。

✅ 正确：`检查横幅组成元素：皇冠图标、标题"Go Premium"`
❌ 错误：`检查标题是否为"Go Premium"`

**预期列**：写直接的结论/结果，不重复描述动作，不用"正确显示"等模糊表述，直接说明应呈现的状态或内容。

✅ 正确：`横幅包含皇冠图标和"Go Premium"标题`
❌ 错误：`标题正确显示"Go Premium"`

**对应关系**：描述第N步 → 预期第N条，编号一一对应。

### 🚫 废弃用例

新版本下线的功能，标记对应用例备注列注明"已废弃"。

**用词风格要求：与原有用例保持绝对一致，颗粒度相同。**

---

## Step 3.5：覆盖率自检（必须执行）

> ⚠️ **在构建 changes.json 之前，必须逐条核对需求文档的每一个功能点，确认全部被用例覆盖。**

完成 Step 3 的变更分析后，执行以下自检流程：

1. **提取需求清单**：从需求文档/截图中逐条列出所有功能点、交互逻辑、状态变化等
2. **逐条核对**：对每个功能点，检查是否有对应的用例（已有用例或新增用例）覆盖
3. **标记遗漏**：未被任何用例覆盖的功能点，必须立即补充为新增用例
4. **输出核对表**：在 chat 中输出核对结果，格式如下：

| 需求功能点 | 覆盖状态 | 对应用例 |
|-----------|---------|---------|
| 铃声专辑-Set as Ringtone | 已覆盖 | 设置面板-铃声专辑 |
| 提示音-Set as Low Battery Alert | 遗漏 | -> 新增用例 |

**只有全部功能点都标记为 ✅ 后，才能进入 Step 4。**

---

## Step 4：构建 changes.json（标准格式专用）

**必须用 Python dict + `json.dump()` 生成，禁止手写 JSON 文本（见关键规则5）。**

```python
import json

changes = {
    "modified": [
        {
            "module": "裁剪页",    # 模块名，与 CSV 模块列填充后的值一致
            "case": "UI",          # 用例名称，与 CSV 用例名称列完全一致
            "col": "C",            # 列字母（A/B/C/D/E）
            "runs": [
                {"text": "1. 原有步骤一\n", "red": False},
                {"text": "2. 新增步骤\n", "red": True},
                {"text": "3. 原有步骤三", "red": False}
            ]
        }
    ],
    "new_rows": [
        {
            "after_module": "Go Premium入口",
            "data": {
                "模块": "Go Premium入口",
                "用例名称": "会员入口文案",
                "描述": "1. 检查入口文案\n2. 检查跳转",
                "预期": "1. 文案正确\n2. 正常跳转",
                "备注": ""
            }
        }
    ],
    "deprecated": [3]   # CSV 原始数据行索引（0-based，不计新增行）
}

with open('changes.json', 'w', encoding='utf-8') as f:
    json.dump(changes, f, ensure_ascii=False, indent=2)
```

字段说明：
- `modified[].module`：模块名（CSV 模块列填充后的实际值，不能是空字符串）
- `modified[].case`：用例名称（与 CSV 用例名称列完全一致）
- `modified[].col`：列字母（A/B/C/D/E）
- `modified[].runs`：富文本段落，`red: True` 为红色新增内容
- `new_rows[].after_module`：插入到哪个模块的最后一行下方（脚本自动追踪行号并标红）
- `deprecated`：CSV 原始数据行索引（0-based，不计新增行），对应行备注列追加"已废弃"

---

## Step 4.5：覆盖率复核（标准格式和非标准格式均必须执行）

> ⚠️ **在生成最终产物之前，必须回溯 Step 3.5 的核对表，逐条验证每个需求功能点在即将生成的内容中有对应落地。**

1. **回溯 Step 3.5 核对表**：逐条检查每个标记为"已覆盖"或"新增用例"的功能点
2. **在即将生成的内容中定位对应条目**：检查 changes.json 中是否有对应的 `modified` / `new_rows` / `deprecated` 条目
3. **标记遗漏**：找不到对应条目的功能点，必须立即补充
4. **输出复核结果**：

| 需求功能点 | Step 3.5 状态 | 落地状态 | 位置 |
|-----------|-------------|---------|------|
| 功能A | 修改用例 | ✅ | modified module=裁剪页 case=UI col=C |
| 功能B | 新增用例 | ✅ | new_rows after_module=裁剪页 |
| 功能C | 新增用例 | ❌ 遗漏 | → 已补充 |

**只有全部功能点都确认落地后，才能进入 Step 5 或 Step 6。**

---

## Step 5：调用固定脚本（标准格式，必须执行此步骤）

**不要自己写生成脚本。将 changes.json 写入磁盘后，直接运行以下命令：**

```bash
pip install openpyxl pandas --break-system-packages -q

python /mnt/skills/user/test-case-generator/scripts/generate.py \
  --input <csv路径> \
  --output output.xlsx \
  --changes changes.json
```

脚本已内置处理所有细节，无需额外代码：
- 插入所有 `new_rows` 后自动构建 `(module, case) → excel_row` 查找表，`modified` 按 module+case 精确定位，不受行号偏移影响
- 新增行插入到正确模块位置并整行标红
- 修改行黑红混排富文本
- 废弃行备注标注"已废弃"
- 富文本颜色/字号 bug 自动修复
- 表头样式、列宽、行高

---

## Step 6：非标准格式 → 临时生成脚本

当 CSV 列名不是标准五列时，根据实际列名临时编写 Python 脚本。需参考以下要点：

### 富文本标红的已知 bug（必须处理）

openpyxl 写入富文本有两个 bug，**必须用 XML 修补方式绕过**：

| Bug | 现象 | 根因 |
|-----|------|------|
| 颜色透明 | 内容存在但不可见 | `InlineFont(color='EA4335')` 写入 `rgb="00EA4335"`，alpha=00 透明 |
| 字体超大 | 内容撑满格子看不见 | `InlineFont(sz=1000)` 单位是半点，1000=500pt |

**修复函数（必须包含在临时脚本中）：**

```python
import zipfile, shutil, os

def fix_rich_text_xlsx(filepath):
    tmp = filepath + '.tmp'
    shutil.copy(filepath, tmp)
    with zipfile.ZipFile(tmp, 'r') as z:
        sheet_bytes = z.read('xl/worksheets/sheet1.xml')
        all_files = {n: z.read(n) for n in z.namelist()}
    fixed = sheet_bytes.replace(b'rgb="00000000"', b'rgb="FF000000"')
    fixed = fixed.replace(b'rgb="00EA4335"', b'rgb="FFEA4335"')
    fixed = fixed.replace(b'<sz val="1000"/>', b'<sz val="10"/>')
    all_files['xl/worksheets/sheet1.xml'] = fixed
    with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)
    os.remove(tmp)
```

**富文本写入方式：**

```python
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText

def make_rich_cell(cell, runs):
    # runs: list of (text, is_red)
    blocks = []
    for text, is_red in runs:
        color = 'EA4335' if is_red else '000000'
        ifont = InlineFont(rFont='Arial', sz=1000, color=color)
        blocks.append(TextBlock(ifont, text))
    cell.value = CellRichText(*blocks)
    cell.alignment = Alignment(wrap_text=True, vertical='top')

# 保存后必须调用修复
wb.save(output_path)
fix_rich_text_xlsx(output_path)
```

---

## Step 7：输出变更说明

在 chat 中附上变更摘要：

```
✏️ 修改用例（N条）
| 原用例 | 修改内容 |
|--------|---------|
| 模块-用例名 | 具体改动描述 |

➕ 新增用例（N条）
| 用例名称 | 覆盖逻辑 |
|---------|---------|
| ... | ... |

🚫 废弃用例：无 / 列出废弃项
```

---

## 注意事项

1. **格式检测优先**：每次收到 CSV 后，第一步必须检测列名，决定走哪条路径
2. **修改行必须保留全部原有内容**：只插入新增/变更部分并标红，不得丢弃原有步骤
3. **颜色修复必须在 `wb.save()` 之后执行**
4. **新增行整行用 `Font(color='EA4335')`**，不需要富文本
5. **修改行只标红新增部分**，原有文字保持黑色
6. **用词风格**：描述用简洁操作动词（查看/检查/点击），不用「是否为」判断句；预期直接写结论状态，不用模糊的「正确显示」，与原有用例颗粒度和句式保持一致
7. **换行符**：步骤之间用 `\n` 分隔，配合 `wrap_text=True`
8. **module 字段必须是填充后的值**：CSV 中模块列空值表示同上，`module` 字段必须填写实际模块名（如 `"裁剪"`），不能是空字符串
