# 11.41B Status Report Generator

## 用户手册 · User Guide

---

### 目录 · Contents

- [简介 · Introduction](#简介--introduction)
- [系统要求 · Requirements](#系统要求--requirements)
- [安装 · Installation](#安装--installation)
- [启动 · Starting the App](#启动--starting-the-app)
- [使用步骤 · How to Use](#使用步骤--how-to-use)
- [字段说明 · Field Reference](#字段说明--field-reference)
- [生成文件说明 · Output File](#生成文件说明--output-file)
- [常见问题 · FAQ](#常见问题--faq)

---

### 简介 · Introduction

**中文**

本工具是一个本地 Web 应用，用于根据 ITU 11.41B 条款自动生成协调状态更新报告（Word 文档）。只需输入目标通知 ID，工具将自动读取对应的 MDB 数据库文件和历史审查 Excel 文件，填充到标准模板中并输出供打印的 `.docx` 文件。

**English**

This is a local web application that automatically generates the 11.41B coordination-status update report (Word document). Enter a Target Notice ID and the tool will read the corresponding MDB database and `(history_space)` Excel file, populate the standard template, and produce a ready-to-print `.docx` file.

---

### 系统要求 · Requirements

| 项目 / Item | 要求 / Requirement |
|---|---|
| 操作系统 / OS | Windows（需要 Microsoft Access ODBC 驱动）/ Windows (Microsoft Access ODBC driver required) |
| Python | 3.9 + |
| M: 盘 / M: drive | 需要映射到 `M:\BRSSD\CSS\` / Must be mapped to `M:\BRSSD\CSS\` |
| 所需 Python 包 | 见 `requirements.txt` |

---

### 安装 · Installation

```bash
cd 11_41B
py -m pip install -r requirements.txt
```

---

### 启动 · Starting the App

```bash
cd 11_41B
py app.py
```

启动后，在浏览器中访问：**http://127.0.0.1:5050**

After starting, open a browser and go to: **http://127.0.0.1:5050**

---

### 使用步骤 · How to Use

**中文**

1. 在浏览器中打开 `http://127.0.0.1:5050`。
2. 页面会自动列出 `M:\BRSSD\CSS\11.41B\2026\` 下所有可用的通知 ID（含 `.mdb` 文件的文件夹）。
3. 从列表中选择一个或多个目标通知 ID（按住 **Ctrl** 可多选）；若无列表，可手动输入一个或多个 ID，逗号分隔。
4. （可选）填写以下字段：
   - **会议编号（WM）**：如 `1668`；若留空，文件名中将使用 `xxxx`。
   - **文件编号（D）**：如 `72317`；若留空，文件名中将使用 `xxxxx`。
   - **会议日期**：如 `4 June 2026`；若留空，将使用今天的日期。
5. 点击 **Generate Report** 按钮。
6. 报告将自动下载，文件名格式为 `11.41B_WM{xxxx}_D{xxxxx}_Draft.docx`。

**English**

1. Open `http://127.0.0.1:5050` in a browser.
2. The page will automatically list all available Notice IDs (folders containing a `.mdb` file) under `M:\BRSSD\CSS\11.41B\2026\`.
3. Select one or more Notice IDs from the list (hold **Ctrl** to multi-select); if no list appears, type one or more IDs separated by commas.
4. (Optional) Fill in the following fields:
   - **Meeting No. (WM)**: e.g. `1668`; if blank, the filename will use `xxxx`.
   - **Document No. (D)**: e.g. `72317`; if blank, the filename will use `xxxxx`.
   - **Meeting Date**: e.g. `4 June 2026`; if blank, today's date is used.
5. Click **Generate Report**.
6. The report downloads automatically. Filename format: `11.41B_WM{xxxx}_D{xxxxx}_Draft.docx`.

---

### 字段说明 · Field Reference

| 字段 / Field | 必填 / Required | 说明 / Description |
|---|---|---|
| 目标通知 ID / Target Notice ID | ✅ | 可多选；文件夹名称，如 `103500417`；无列表时输入逗号分隔的多个 ID / Multi-selectable; folder name e.g. `103500417`; comma-separated IDs when no list |
| 会议编号（WM）/ Meeting No. (WM) | ❌ | 显示在标题中，如 `1668` → `1668th meeting` / Shown in heading |
| 文件编号（D）/ Document No. (D) | ❌ | 显示在文件名中 / Used in filename |
| 会议日期 / Meeting Date | ❌ | 显示在标题中，如 `4 June 2026` / Shown in heading |

---

### 生成文件说明 · Output File

生成的 Word 文档包含以下修改内容：

The generated Word document contains the following updated content:

| 位置 / Location | 内容 / Content |
|---|---|
| 文档标题 / Document heading | `{N}th meeting on {date}`（序数上角标）/ ordinal suffix is superscript |
| 文件名 / Filename | `11.41B_WM{WM}_D{D}_Draft.docx` |
| 正文第二段 / Body paragraph 2 | 条款编号（如 `No. 9.7 \|O\|`）替换为实际值，**黄色高亮** / provision numbers replaced with actual values, **yellow highlighted** |
| 表格第一张 / Table 1 | 每条通知各自一行或多行；最后一列为该通知修改的条款 / One or more rows per notice; last column shows per-notice provisions |
| 页脚文档编号 / Footer document No. | `D {D编号}`，与输入的文件编号一致 / `D {D number}`, matching the entered document number |

**表格列映射 / Table Column Mapping**

| 列 / Column | 数据来源 / Source |
|---|---|
| Notice ID | `com_el.ntc_id` |
| ADM | `com_el.adm` |
| STATION | `com_el.sat_name` |
| TYPE | `G` if `com_el.long_nom` has a value; `N` if empty |
| Date of Receipt | `com_el.d_rcv`（格式 `DD.MM.YYYY`） |
| ADM Coordination completed | `history_space` 文件中 `adm` 列的去重值 / Distinct `adm` values from `(history_space)` file |
| 11.41 still exist | `provn.coord_prov` 中任一值包含 `11.41` 则为 `YES`，否则 `NO` / `YES` if any `provn.coord_prov` value contains `11.41` (e.g. `11.41`, `11.41/9.13`); `NO` otherwise |
| Provision updated from 11.41\|X\| | `(history_space)` 文件中 `provn_target` 列去重值，格式 `No. X.Y \|O\|` / Distinct `provn_target` values formatted as `No. X.Y \|O\|` |

---

### 常见问题 · FAQ

**Q: 下拉框为空 / The drop-down is empty.**

A（中文）: 请确认 M: 盘已正确映射，且 `M:\BRSSD\CSS\11.41B\2026\` 下存在含 `.mdb` 文件的文件夹。

A (English): Confirm that the M: drive is mapped and that `M:\BRSSD\CSS\11.41B\2026\` contains subfolders with `.mdb` files.

---

**Q: 提示"MDB file not found"。**

A（中文）: 确认所选通知 ID 的文件夹下存在同名的 `.mdb` 文件（如 `103500417\103500417.mdb`）。

A (English): Confirm that the notice ID folder contains an `.mdb` file with the same name (e.g. `103500417\103500417.mdb`).

---

**Q: 提示"No (history_space) file found"。**

A（中文）: 确认 `M:\BRSSD\CSS\11.41B\2026\BR_TEXT_RESULTS\{通知ID}\Status Review\` 目录下存在以 `(history_space).xlsx` 结尾的文件。

A (English): Confirm that a file ending with `(history_space).xlsx` exists under `M:\BRSSD\CSS\11.41B\2026\BR_TEXT_RESULTS\{noticeID}\Status Review\`.

---

## 开发文档 · Developer Guide

---

### 目录 · Contents

- [项目结构 · Project Structure](#项目结构--project-structure)
- [技术栈 · Tech Stack](#技术栈--tech-stack)
- [数据流 · Data Flow](#数据流--data-flow)
- [核心函数 · Core Functions](#核心函数--core-functions)
- [Word 文档生成逻辑 · Document Generation Logic](#word-文档生成逻辑--document-generation-logic)
- [字体规范 · Font Specification](#字体规范--font-specification)
- [扩展指南 · Extension Guide](#扩展指南--extension-guide)
- [依赖说明 · Dependencies](#依赖说明--dependencies)

---

### 项目结构 · Project Structure

```
11_41B/
├── app.py               # Flask 后端 / Flask backend
├── requirements.txt     # Python 依赖 / Python dependencies
├── README.md            # 本文档 / This document
└── templates/
    └── index.html       # 前端页面 / Frontend page
```

**数据路径（固定，在 `app.py` 顶部配置）/ Data paths (configured at top of `app.py`)**

```
M:\BRSSD\CSS\11.41B\2026\
├── {notice_id}\
│   └── {notice_id}.mdb          # Access 数据库 / Access DB
└── BR_TEXT_RESULTS\
    └── {notice_id}\
        └── Status Review\
            └── **(history_space).xlsx
```

---

### 技术栈 · Tech Stack

| 组件 / Component | 技术 / Technology |
|---|---|
| Web 框架 / Web framework | Flask 3+ |
| MDB 读取 / MDB reading | pyodbc + Microsoft Access ODBC (32-bit) |
| Excel 读取 / Excel reading | openpyxl |
| Word 生成 / Word generation | python-docx |
| XML 操作 / XML manipulation | lxml (via python-docx internals) |
| 前端 / Frontend | Vanilla HTML/CSS/JS (fetch API) |

---

### 数据流 · Data Flow

```
User Input (notice_id, wm_num, d_num, meeting_date)
    │
    ▼
get_mdb_data(notice_id)
    ├─ com_el: ntc_id, adm, sat_name, long_nom→type, d_rcv
    └─ provn:  agree_st → still_exist (YES/NO)
    │
    ▼
get_history_data(notice_id)
    ├─ distinct provn_target → sorted numerically
    └─ distinct adm
    │
    ▼
generate_report(...)
    ├─ Load template .docx (fresh copy each call)
    ├─ _update_meeting_para()   → heading with superscript ordinal
    ├─ _replace_provn_section() → yellow-highlighted provision refs
    └─ _replace_table_data()    → cloned template row + actual data
    │
    ▼
BytesIO buffer → send_file() → browser download
```

---

### 核心函数 · Core Functions

#### `_make_run_xml(text, bold, highlight, superscript, font_name, sz_cs)`

创建 `w:r` XML 元素，包含完整的字体属性。

Creates a `w:r` XML element with full font properties.

- `font_name`：字体名称，默认 `'Calibri'` / Font name, defaults to `'Calibri'`
- `sz_cs`：复杂脚本字号（half-points），默认 `22`（= 11 pt）/ Complex-script size in half-points, default `22`
- `superscript`：设置 `w:vertAlign val="superscript"` / Sets superscript alignment
- `highlight`：高亮颜色，如 `'yellow'` / Highlight colour

#### `_update_meeting_para(para, wm_num, meeting_date)`

将标题段落重建为 **3 个 run**：

Rebuilds the heading paragraph as **3 runs**:

1. `"Presented at the {N}"` — bold, Calibri
2. `"{th/st/nd/rd}"` — bold, Calibri, **superscript**
3. `" meeting on {date}"` — bold, Calibri

#### `_replace_provn_section(para, provn_targets)`

在协调段落中，定位 `|X| to` 与 `or removed from` 之间的 runs，将其替换为带**黄色高亮** Calibri 字体的 provision 编号 runs。

Locates the runs between `|X| to` and `or removed from` in the coordination paragraph and replaces them with yellow-highlighted Calibri provision-number runs.

格式 / Format: `No. {value} |O|`（中间的数字加粗 / the number is bold）

#### `_set_cell_text(cell, text)`

从克隆的模板行中**保留原有 `rPr`**（包含 `minorHAnsi` 主题字体 + 10 pt 字号），只替换文字内容。

**Preserves the existing `rPr`** from the cloned template row (contains `minorHAnsi` theme font + 10 pt size), replacing only the text content.

#### `_update_footer(doc, wm_num, d_num)`

遍历所有 section 的页脚 run，用正则将 `D\s*\d+`（非字母前缀）替换为 `D {d_num}`，将 `WM\d+` 替换为 `WM{wm_num}`，保持页脚其他内容（如 `No.11.41B`、页码）不变。

Iterates all section footer runs; uses regex to replace `D\s*\d+` (not preceded by a letter) with `D {d_num}` and `WM\d+` with `WM{wm_num}`, leaving the rest of the footer (e.g. `No.11.41B`, page number) unchanged.

#### `_replace_table_data(table, notice_data_list)`

接受 `(mdb_data, history_data)` 元组列表，每个 notice 的每条 `com_el` 记录输出一行，最后一列为该 notice 的条款列表（`No. X.Y |O|` 格式）。

Accepts a list of `(mdb_data, history_data)` tuples; outputs one row per `com_el` record per notice, with the last column containing that notice's provision list (`No. X.Y |O|` format).

1. 将第一个数据行的 XML 深拷贝为模板。
2. 删除所有数据行（保留表头行）。
3. 遍历所有 notice，为每条 `com_el` 记录克隆该模板行并填入数据（含最后一列条款）。

1. Deep-copies the first data row's XML as a template.  
2. Deletes all data rows (keeps header row 0).  
3. Iterates all notices; clones the template row for each `com_el` record, filling in data including the provision list in the last column.

---

### Word 文档生成逻辑 · Document Generation Logic

模板文件每次请求都会**重新加载**（`Document(TEMPLATE_PATH)`），确保多次生成互不干扰。

The template file is **reloaded fresh on every request** (`Document(TEMPLATE_PATH)`), ensuring no state leaks between requests.

段落定位使用**文本子串匹配**，而非硬编码的段落索引，以便在模板内容微调后仍能正常工作。

Paragraphs are located by **text substring matching**, not hardcoded indices, so minor template edits won't break generation.

---

### 字体规范 · Font Specification

| 元素 / Element | 字体 / Font | 字号 / Size | 备注 / Notes |
|---|---|---|---|
| 标题段 / Heading para | Calibri (`w:ascii/hAnsi/cs`) | szCs 22 (11 pt) | Bold; ordinal suffix is superscript |
| 正文高亮 runs / Body highlighted runs | Calibri | szCs 22 (11 pt) | `w:highlight val="yellow"` |
| 表格数据 / Table data | Calibri via `minorHAnsi` theme | sz/szCs 20 (10 pt) | Preserved from cloned template row |

---

### 扩展指南 · Extension Guide

**添加新的数据列 / Adding a new table column**

在 `_replace_table_data` 的 `values` 列表中追加值，并确保表格模板也有对应的列。

Append the value to the `values` list in `_replace_table_data`, and ensure the template table has the corresponding column.

**更改模板文件 / Changing the template file**

修改 `app.py` 顶部的 `TEMPLATE_PATH` 常量。

Change the `TEMPLATE_PATH` constant at the top of `app.py`.

**更改基准目录 / Changing the base directory**

修改 `app.py` 顶部的 `BASE_DIR` 常量。

Change the `BASE_DIR` constant at the top of `app.py`.

**部署为生产服务 / Deploying to production**

将 Flask 开发服务器替换为 `waitress` 或 `gunicorn`：

Replace the Flask dev server with `waitress` or `gunicorn`:

```bash
# Windows (waitress)
pip install waitress
waitress-serve --port=5050 app:app

# Linux (gunicorn)
gunicorn -w 2 -b 0.0.0.0:5050 app:app
```

---

### 依赖说明 · Dependencies

| 包 / Package | 用途 / Purpose |
|---|---|
| `flask` | HTTP 路由 / HTTP routing |
| `python-docx` | Word 文档读写 / Word document read/write |
| `openpyxl` | Excel 读取 / Excel reading |
| `pyodbc` | ODBC 连接 Access MDB / ODBC connection to Access MDB |
| `lxml` | python-docx 的 XML 依赖 / XML dependency for python-docx |

> **注意 / Note**: `pyodbc` 连接 `.mdb` 文件需要 Windows 上安装的 **32-bit Microsoft Access Database Engine**（与 Python 位数匹配）。
>
> `pyodbc` connecting to `.mdb` files requires the **32-bit Microsoft Access Database Engine** installed on Windows (matching the Python bitness).
