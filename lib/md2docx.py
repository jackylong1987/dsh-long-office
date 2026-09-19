#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generic Markdown -> styled .docx converter with a page-number footer.

Usage:
  python3 md2docx.py <input.md> [output.docx]

If output.docx is omitted, it defaults to <input-name>.docx next to the input.

Renders headings (h1-h3), bold/italic inline, tables (pipe syntax),
ordered/unordered lists, blockquotes, and horizontal rules. Adds a centered
footer with an auto-updating PAGE field (updates when opened in Word or
exported to PDF).

Requires: python3 + python-docx (`pip install python-docx`).
"""
import re
import sys
import os

from docx import Document
from docx.shared import Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def die(msg):
    print(f"md2docx: {msg}", file=sys.stderr)
    sys.exit(1)


if len(sys.argv) < 2:
    die("usage: md2docx.py <input.md> [output.docx]")

SRC = os.path.abspath(sys.argv[1])
if not os.path.isfile(SRC):
    die(f"input file not found: {SRC}")
OUT = (
    os.path.abspath(sys.argv[2])
    if len(sys.argv) > 2
    else os.path.splitext(SRC)[0] + ".docx"
)

with open(SRC, encoding="utf-8") as f:
    lines = f.read().splitlines()


def add_page_number(paragraph):
    """Insert an auto-updating '第 N 页' page field into a footer paragraph."""
    run = paragraph.add_run("第 ")
    set_east_asia(run)
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    # 2026-09-20 修：**域字符与缓存结果必须分在不同 run**。
    # 浏览器的 docx 预览器（docx-preview）会把"含 fldChar/instrText 的整个 run"丢掉；
    # 若缓存数字与域字符同 run，数字会一起消失 → 预览里页码位置错乱（表现为「第 1页」）。
    # 拆开后：域 run 被丢，数字 run 保留，页码位置永远正确。
    def _run(text=None):
        rr = paragraph.add_run() if text is None else paragraph.add_run(text)
        set_east_asia(rr)
        return rr
    _run()._r.append(fld)
    _run()._r.append(instr)
    _run()._r.append(sep)
    _run("1")
    _run()._r.append(end)
    _run(" 页")


doc = Document()
# ── 纸张：A4（210×297mm）────────────────────────────────────────────
# python-docx 的默认模板是 US Letter（215.9×279.4mm），中文公文/专利文件必须 A4；
# 2026-09-20 修：此前所有由本脚本生成的 docx 都是 Letter（页数与版式会与 Word/A4 不一致）。
for _sec in doc.sections:
    _sec.page_width = Mm(210)
    _sec.page_height = Mm(297)


# --- base styles ---
normal = doc.styles["Normal"]
normal.font.name = "宋体"
normal.font.size = Pt(10.5)
normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
# --- 2026-09-14 用户反馈「生成的 docx 行距被压（预览里比我自己放上去的文件紧）」---
# 根因：python-docx 默认模板的 docDefaults 是 <w:spacing w:line="276" w:lineRule="auto"/>（≈1.15 倍），
# 而用户自己的文件基本都是 1.5 倍（w:line="360"）。这里统一为 1.5 倍，Word 打开与预览观感一致。
normal.paragraph_format.line_spacing = 1.5


# --- 2026-09-12 用户要求：全篇只用黑色字体、任何底纹都不能有 ---
normal.font.color.rgb = RGBColor(0, 0, 0)
for _sn in ("Title", "Heading 1", "Heading 2", "Heading 3", "Heading 4",
            "Intense Quote", "List Bullet", "List Number", "Table Grid"):
    try:
        _st = doc.styles[_sn]
    except KeyError:
        continue
    try:
        _st.font.color.rgb = RGBColor(0, 0, 0)
    except Exception:
        pass
    try:  # 2026-09-15 用户要求：全部报告/文件用宋体（含标题，避免继承主题 major font）
        _st.font.name = "宋体"
        _st.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    except Exception:
        pass
    try:  # 2026-09-14：标题/列表等样式同样统一 1.5 倍行距（否则仍继承 docDefaults 的 1.15 倍）
        _st.paragraph_format.line_spacing = 1.5
    except Exception:
        pass
    # 去掉该样式自带的底纹（w:shd）
    for _el in list(_st.element.iter()):
        if _el.tag == qn("w:shd"):
            _el.getparent().remove(_el)
    # 2026-09-15 用户要求：报告不要页眉、标题下不要装饰线 → 去掉样式自带的段落下边框
    for _el in list(_st.element.iter()):
        if _el.tag == qn("w:pBdr"):
            _el.getparent().remove(_el)
# 清空页眉（只用页脚页码）
try:
    for _hp in doc.sections[0].header.paragraphs:
        for _r in list(_hp.runs):
            _r._element.getparent().remove(_r._element)
except Exception:
    pass


def set_east_asia(run):
    run.font.name = "宋体"
    run.font.color.rgb = RGBColor(0, 0, 0)   # 2026-09-12：只用黑色
    r = run._element
    rPr = r.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.append(rf)
    rf.set(qn("w:ascii"), "Times New Roman")     # 西文/数字：Times New Roman（宋体空格是全角）
    rf.set(qn("w:hAnsi"), "Times New Roman")
    rf.set(qn("w:eastAsia"), "宋体")


def add_runs_with_bold(par, text):
    """2026-09-15 用户要求：**正文一律不加粗，只有标题加粗**。

    md 里的 `**…**` / `*…*` 仅作标记清理，不再转成加粗/斜体
    （标题由 Heading 样式自带加粗；表格表头由 flush_table 单独置粗）。
    """
    r = par.add_run(re.sub(r"\*+", "", text))
    set_east_asia(r)


def add_body_paragraph(text, style=None):
    p = doc.add_paragraph(style=style)
    add_runs_with_bold(p, text)
    return p


def flush_table():
    global table_rows, in_table
    if not table_rows:
        in_table = False
        return
    ncols = max(len(r) for r in table_rows)
    tbl = doc.add_table(rows=len(table_rows), cols=ncols)
    tbl.style = "Table Grid"   # 2026-09-12：原来 Light Grid Accent 1 带底纹/彩色框线
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ri, row in enumerate(table_rows):
        for ci in range(ncols):
            cell = tbl.cell(ri, ci)
            cell.text = ""
            cp = cell.paragraphs[0]
            text = row[ci] if ci < len(row) else ""
            add_runs_with_bold(cp, text)
            if ri == 0:
                for run in cp.runs:
                    run.bold = True
    tbl.autofit = True
    table_rows = []
    in_table = False


table_rows = []
in_table = False

i = 0
while i < len(lines):
    stripped = lines[i].strip()
    if not stripped:
        i += 1
        continue
    if re.fullmatch(r"-{3,}", stripped):
        flush_table()
        doc.add_paragraph()
        i += 1
        continue
    if in_table and "|" in stripped and re.fullmatch(r"\|?[\s:|-]+\|?", stripped):
        i += 1
        continue
    if stripped.startswith("|"):
        in_table = True
        table_rows.append([c.strip() for c in stripped.strip("|").split("|")])
        i += 1
        continue
    if in_table:
        flush_table()
    if stripped.startswith("### "):
        h = doc.add_heading(level=3)
        add_runs_with_bold(h, stripped[4:])
        i += 1
        continue
    if stripped.startswith("## "):
        h = doc.add_heading(level=2)
        add_runs_with_bold(h, stripped[3:])
        i += 1
        continue
    if stripped.startswith("# "):
        h = doc.add_heading(level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER          # ★文档标题居中（用户要求）
        add_runs_with_bold(h, stripped[2:])
        i += 1
        continue
    if stripped.startswith("> "):
        p = doc.add_paragraph(style="Intense Quote")
        add_runs_with_bold(p, stripped[2:] + " ")
        i += 1
        continue
    m = re.match(r"^(\d+)\.\s+(.*)$", stripped)
    if m:
        p = doc.add_paragraph(style="List Number")
        add_runs_with_bold(p, m.group(2))
        i += 1
        continue
    if stripped.startswith("- "):
        p = doc.add_paragraph(style="List Bullet")
        add_runs_with_bold(p, stripped[2:])
        i += 1
        continue
    p = doc.add_paragraph()
    add_runs_with_bold(p, stripped)
    i += 1

if in_table:
    flush_table()

# --- footer with page number field ---
footer = doc.sections[0].footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_page_number(fp)
# 2026-09-15 用户要求：页码字体与正文一致（宋体）、居中
for _r in fp.runs:
    _r.font.name = "宋体"
    _r.font.color.rgb = RGBColor(0, 0, 0)
    _rp = _r._element.get_or_add_rPr()
    _rp.get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
try:
    _fs = fp.style
    _fs.font.name = "宋体"
    _fs.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
except Exception:
    pass

# --- 2026-09-15 用户要求：**所有生成文件**都必须符合 skill `doc-heading-format` ---
# 在源头统一：正文首行缩进 2 字符、表格内容居中（水平＋垂直）、全黑字、无底纹、行距 1.5。
# 放在这里，任何技能/任何调用 md2docx 得到的 docx 都自动合规，不依赖调用方是否记得。
def apply_doc_format(doc):
    for p in doc.paragraphs:
        st = (p.style.name or "")
        is_head = st.startswith(("Heading", "Title")) or p.alignment is not None
        pPr = p._p.get_or_add_pPr()
        for _shd in pPr.findall(qn("w:shd")):
            pPr.remove(_shd)
        for r in p.runs:
            r.font.color.rgb = RGBColor(0, 0, 0)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(6)
        # 只有普通正文段落才首行缩进（标题/列表项/引用不缩进）
        if (not is_head) and p.text.strip() and st in ("Normal", ""):
            ind = pPr.get_or_add_ind()
            ind.set(qn("w:firstLineChars"), "200")   # 2 字符（中文习惯）
            ind.set(qn("w:firstLine"), "420")
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                try:
                    c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                except Exception:
                    pass
                tcPr = c._tc.get_or_add_tcPr()
                for _shd in tcPr.findall(qn("w:shd")):
                    tcPr.remove(_shd)
                for p in c.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER   # ★表格内容一律居中
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)
                    for r in p.runs:
                        r.font.color.rgb = RGBColor(0, 0, 0)


apply_doc_format(doc)

doc.save(OUT)
print("Saved:", OUT)
