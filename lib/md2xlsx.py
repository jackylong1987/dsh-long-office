#!/usr/bin/env python3
"""md2xlsx.py — Markdown 表格 → Excel(.xlsx)

排版口径与 dsh-office-reader 的 md2docx 一致（用户硬规则）：
  · 中文宋体 / 纯西文单元格 Times New Roman；全黑字；**无底纹（不加填充色）**
  · 表格：细边框、内容水平+垂直居中、表头加粗、列宽按内容自适应、首行冻结
  · 以 `=` 开头的单元格写成**真正的 Excel 公式**；数字串转数值
  · 每个 Markdown 表格 → 一个 sheet，表名取该表格上方最近的标题（无则 Sheet1/2…）

用法:
  python3 md2xlsx.py <input.md> [output.xlsx] [--force]
  默认输出同名 .xlsx；默认**不覆盖**已存在文件（防误覆盖），--force 才覆盖。
"""
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

EAST = "宋体"
LATIN = "Times New Roman"
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
HEAD_FONT_EA = Font(name=EAST, bold=True, color="FF000000")
BODY_FONT_EA = Font(name=EAST, color="FF000000")
HEAD_FONT_LA = Font(name=LATIN, bold=True, color="FF000000")
BODY_FONT_LA = Font(name=LATIN, color="FF000000")

NUM_RE = re.compile(r"^-?\d+(?:\.\d+)?%?$")
ASCII_RE = re.compile(r"^[\x00-\x7F]*$")
BAD_SHEET = re.compile(r"[\\/*?:\[\]]")


def is_sep_row(cells):
    """| --- | :--: | 这种分隔行"""
    return all(re.fullmatch(r":?-{2,}:?", c.strip()) for c in cells if c.strip() != "") and any(
        c.strip() for c in cells
    )


def split_row(line):
    raw = line.strip()
    if raw.startswith("|"):
        raw = raw[1:]
    if raw.endswith("|"):
        raw = raw[:-1]
    return [c.strip() for c in raw.split("|")]


def parse_tables(md_text):
    """→ [(sheet_name, rows)]  rows = [[cell, ...], ...]"""
    out, cur, heading, pending_heading = [], [], None, None
    in_table = False
    for line in md_text.splitlines():
        s = line.strip()
        if s.startswith("#"):
            pending_heading = s.lstrip("#").strip() or None
            if in_table:
                in_table = False
            continue
        if s.startswith("|") and s.count("|") >= 2:
            cells = split_row(s)
            if not in_table:
                in_table = True
                cur = []
                heading = pending_heading
            if is_sep_row(cells):
                continue
            cur.append(cells)
        else:
            if in_table and cur:
                out.append((heading, cur))
                cur, in_table = [], False
    if in_table and cur:
        out.append((heading, cur))
    return out


def clean_sheet_name(name, used, idx):
    base = BAD_SHEET.sub("-", (name or "").strip()) or f"Sheet{idx}"
    base = base[:31]
    cand, n = base, 2
    while cand in used:
        suffix = f"({n})"
        cand = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(cand)
    return cand


def put_cell(ws, r, c, raw, bold):
    val = raw
    if val.startswith("="):
        ws.cell(row=r, column=c, value=val)          # 真公式
    elif NUM_RE.fullmatch(val):
        try:
            v = float(val[:-1]) / 100 if val.endswith("%") else float(val)
            ws.cell(row=r, column=c, value=v)
        except ValueError:
            ws.cell(row=r, column=c, value=val)
    else:
        ws.cell(row=r, column=c, value=val)
    cell = ws.cell(row=r, column=c)
    ascii_only = bool(ASCII_RE.fullmatch(val)) and val != ""
    if bold:
        cell.font = HEAD_FONT_LA if ascii_only else HEAD_FONT_EA
    else:
        cell.font = BODY_FONT_LA if ascii_only else BODY_FONT_EA
    cell.alignment = CENTER
    cell.border = BORDER
    # 清掉任何填充（用户硬规则：不许有底纹）
    cell.fill = PatternFill(fill_type=None)


def write_sheet(wb, name, rows, idx):
    ws = wb.create_sheet(title=name)
    for ri, row in enumerate(rows, start=1):
        for ci, raw in enumerate(row, start=1):
            put_cell(ws, ri, ci, raw, bold=(ri == 1))
    # 列宽自适应（中文按 2 个字符宽估）
    for ci in range(1, max(len(r) for r in rows) + 1):
        widest = 0
        for row in rows:
            if ci <= len(row):
                txt = row[ci - 1]
                w = sum(2 if ord(ch) > 127 else 1 for ch in txt)
                widest = max(widest, w)
        ws.column_dimensions[get_column_letter(ci)].width = min(max(widest + 2, 6), 60)
    # 行高与冻结表头
    for ri in range(1, len(rows) + 1):
        ws.row_dimensions[ri].height = 20
    ws.freeze_panes = "A2"
    return ws


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    force = "--force" in sys.argv
    if not args:
        print("用法: md2xlsx.py <input.md> [output.xlsx] [--force]", file=sys.stderr)
        return 2
    src = os.path.abspath(args[0])
    out = os.path.abspath(args[1]) if len(args) > 1 else re.sub(r"\.md$", "", src, flags=re.I) + ".xlsx"
    if os.path.exists(out) and not force:
        print(f"已存在，未覆盖（加 --force 覆盖）: {out}", file=sys.stderr)
        return 3
    if not os.path.isfile(src):
        print(f"找不到输入: {src}", file=sys.stderr)
        return 2

    with open(src, encoding="utf-8") as fh:
        tables = parse_tables(fh.read())
    if not tables:
        print("输入里没有 Markdown 表格（| … | 且含分隔行）", file=sys.stderr)
        return 4

    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    for i, (heading, rows) in enumerate(tables, start=1):
        write_sheet(wb, clean_sheet_name(heading, used, i), rows, i)
    wb.save(out)
    print("Saved:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
