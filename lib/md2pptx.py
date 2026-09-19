#!/usr/bin/env python3
"""md2pptx.py — Markdown 大纲 → PowerPoint(.pptx)，16:9

完全自研（只依赖 python-pptx + Pillow），排版口径沿用本机硬规则：
  · 白底、**无底纹/无花哨填充**、文字全黑；中文宋体、纯西文 Times New Roman
  · 16:9（13.333 × 7.5 英寸）；标题 28pt 粗、正文 18pt、备注 12pt
  · **演讲者备注**：`> 备注：...` 行写进当前页 notes
  · **图片**：`![说明](路径)` 默认按原图比例放入内容区；`![说明](路径 "4x3")` 指定英寸宽×高
  · **坐标回显**：生成后打印每页每个元素的落点（英寸坐标 + 画布尺寸 + 文本线框）

Markdown 约定（自上而下）:
  # 标题              → 封面（副标题取 `> ...`）
  ## 小节标题         → 新的一页（内容页）
  - 要点 /   - 子要点  → 项目符号（1/2 级）
  > 备注：xxx         → 当前页备注（封面页则作副标题）
  ![说明](图片路径)    → 当前页插图
  ---                → 换页（无标题）
  （普通段落）         → 当前页正文段落

用法:
  python3 md2pptx.py <input.md> [output.pptx] [--force] [--echo-json out.json]
  默认输出同名 .pptx；默认不覆盖已存在文件，--force 才覆盖。
"""
import json
import os
import re
import sys

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

EAST, LATIN = "宋体", "Times New Roman"
BLACK = RGBColor(0, 0, 0)
SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN, TITLE_H = 0.6, 1.2
TITLE_PT, BODY_PT, NOTE_PT = 28, 18, 12
BULLET_CHARS = "•", "–"
ASCII_RE = re.compile(r"^[\x00-\x7F]*$")
SEG_RE = re.compile(r"[\x00-\x7F]+|[^\x00-\x7F]+")     # 按 ASCII / 非 ASCII 分段，西文走 TNR
IMG_RE = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<path>[^)\s]+)(?:\s+\"(?P<size>[0-9.]+x[0-9.]+)\")?\)\s*$")
NOTE_RE = re.compile(r"^>\s*(?:备注|Notes?|note)\s*[:：]?\s*(.*)$", re.I)
QUOTE_RE = re.compile(r"^>\s*(.*)$")


def parse(md_text):
    """→ [ {title, subtitle, bullets:[(level,text)], images:[(path,alt,w,h)], paragraphs:[str], notes:str} ]"""
    slides, cur = [], None

    def new(title=None):
        s = {"title": title, "subtitle": None, "bullets": [], "images": [],
             "paragraphs": [], "notes": ""}
        slides.append(s)
        return s

    for raw in md_text.splitlines():
        line = raw.rstrip()
        s = line.strip()
        if not s:
            continue
        m_img = IMG_RE.match(s)
        if s == "---":
            new()
            continue
        if s.startswith("# "):
            cur = new(s[2:].strip())          # 封面
            continue
        if s.startswith("## "):
            cur = new(s[3:].strip())          # 内容页
            continue
        if cur is None:
            cur = new()
        m_note = NOTE_RE.match(s)
        if m_note:
            if cur["title"] and not cur["bullets"] and not cur["paragraphs"] and not cur["notes"]:
                cur["subtitle"] = m_note.group(1).strip()   # 封面页的 > 备注： → 副标题
            else:
                cur["notes"] = (cur["notes"] + "\n" + m_note.group(1).strip()).strip()
            continue
        m_quote = QUOTE_RE.match(s)
        if m_quote:
            cur["notes"] = (cur["notes"] + "\n" + m_quote.group(1).strip()).strip()
            continue
        if m_img:
            size = m_img.group("size")
            w, h = (float(x) for x in size.split("x")) if size else (None, None)
            cur["images"].append((m_img.group("path"), m_img.group("alt"), w, h))
            continue
        m_b = re.match(r"^(\s*)[-*+]\s+(.*)$", raw)
        if m_b:
            level = 1 if len(m_b.group(1)) < 2 else 2
            cur["bullets"].append((level, m_b.group(2).strip()))
            continue
        cur["paragraphs"].append(s)
    return [s for s in slides if any([s["title"], s["bullets"], s["images"], s["paragraphs"]])]


def style_run(run, size, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = BLACK
    txt = run.text or ""
    run.font.name = LATIN if ASCII_RE.fullmatch(txt) and txt else EAST
    # 东亚字体显式设置（python-pptx 的 font.name 只管 latin）
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", EAST)


def add_styled_runs(paragraph, text, size, bold=False):
    """按'纯西文段／中日韩段'拆成多个 run：西文 Times New Roman、中文宋体（用户硬规则）"""
    wrote = False
    for seg in SEG_RE.findall(text):
        if not seg:
            continue
        run = paragraph.add_run()
        run.text = seg
        style_run(run, size, bold)
        wrote = True
    if not wrote:
        run = paragraph.add_run()
        run.text = ""
        style_run(run, size, bold)


def add_textbox(slide, left, top, width, height):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    return tf


def fill_title(slide, title, subtitle):
    tf = add_textbox(slide, MARGIN, 0.35, SLIDE_W - 2 * MARGIN, TITLE_H)
    add_styled_runs(tf.paragraphs[0], title or "", TITLE_PT, bold=True)
    if subtitle:
        add_styled_runs(tf.add_paragraph(), subtitle, 14)


def picture_size(path, w, h):
    with Image.open(path) as im:
        iw, ih = im.size
    ratio = iw / ih if ih else 1.0
    if w and h:
        return w, h
    box_w, box_h = 5.4, 4.6
    if not w and not h:
        w = box_w
        h = w / ratio
        if h > box_h:
            h = box_h
            w = h * ratio
    elif w and not h:
        h = w / ratio
    else:
        w = h * ratio
    return w, h


def build(md_text, echo_json=None):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SLIDE_W), Inches(SLIDE_H)
    blank = prs.slide_layouts[6]
    geometry = []

    for spec in parse(md_text):
        slide = prs.slides.add_slide(blank)
        si = len(prs.slides._sldIdLst)          # 1-based 页码
        is_cover = spec["subtitle"] is not None or (spec["title"] and not spec["bullets"]
                                                   and not spec["images"] and not spec["paragraphs"]
                                                   and not spec["notes"])
        if spec["title"]:
            if is_cover:
                tf = add_textbox(slide, MARGIN, 2.6, SLIDE_W - 2 * MARGIN, 1.6)
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                add_styled_runs(p, spec["title"], 40, bold=True)
                if spec["subtitle"]:
                    p2 = tf.add_paragraph()
                    p2.alignment = PP_ALIGN.CENTER
                    add_styled_runs(p2, spec["subtitle"], 16)
            else:
                fill_title(slide, spec["title"], None)

        has_img = bool(spec["images"])
        text_w = 6.2 if has_img else SLIDE_W - 2 * MARGIN
        top = 1.75 if spec["title"] and not is_cover else 1.2

        if spec["bullets"] or spec["paragraphs"]:
            tf = add_textbox(slide, MARGIN, top, text_w, SLIDE_H - top - MARGIN)
            first = True
            for level, text in spec["bullets"]:
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.level = min(level - 1, 4)
                p.space_after = Pt(6)
                bullet = BULLET_CHARS[min(level, 2) - 1]
                add_styled_runs(p, f"{bullet} {text}", BODY_PT)
            for text in spec["paragraphs"]:
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                add_styled_runs(p, text, BODY_PT)

        for idx, (path, alt, w, h) in enumerate(spec["images"]):
            if not os.path.isfile(path):
                print(f"  警告：图片不存在，已跳过 → {path}", file=sys.stderr)
                continue
            w, h = picture_size(path, w, h)
            left = SLIDE_W - MARGIN - w if has_img else MARGIN
            pic_top = max(top, (SLIDE_H - h) / 2)
            pic = slide.shapes.add_picture(path, Inches(left), Inches(pic_top), Inches(w), Inches(h))
            pic.name = alt or os.path.basename(path)

        if spec["notes"]:
            tf = slide.notes_slide.notes_text_frame
            tf.text = spec["notes"]
            for p in tf.paragraphs:
                for r in p.runs:
                    style_run(r, NOTE_PT)

        for sh in slide.shapes:
            kind = str(sh.shape_type).split()[0] if sh.shape_type is not None else "TEXT_BOX"
            text = sh.text_frame.text.replace("\n", " / ")[:60] if sh.has_text_frame else (
                sh.name if kind == "PICTURE" else "")
            geometry.append({
                "slide": si,
                "shape": kind,
                "name": sh.name,
                "left_in": round(sh.left / 914400, 3), "top_in": round(sh.top / 914400, 3),
                "width_in": round(sh.width / 914400, 3), "height_in": round(sh.height / 914400, 3),
                "text": text,
            })
    if echo_json:
        with open(echo_json, "w", encoding="utf-8") as fh:
            json.dump({"canvas_in": [SLIDE_W, SLIDE_H], "slides": geometry}, fh, ensure_ascii=False, indent=2)
    return prs, geometry


def main():
    argv = [a for a in sys.argv[1:] if not a.startswith("--")]
    force = "--force" in sys.argv
    echo = None
    if "--echo-json" in sys.argv:
        echo = sys.argv[sys.argv.index("--echo-json") + 1]
    if not argv:
        print("用法: md2pptx.py <input.md> [output.pptx] [--force] [--echo-json out.json]", file=sys.stderr)
        return 2
    src = os.path.abspath(argv[0])
    out = os.path.abspath(argv[1]) if len(argv) > 1 else re.sub(r"\.md$", "", src, flags=re.I) + ".pptx"
    if os.path.exists(out) and not force:
        print(f"已存在，未覆盖（加 --force 覆盖）: {out}", file=sys.stderr)
        return 3
    if not os.path.isfile(src):
        print(f"找不到输入: {src}", file=sys.stderr)
        return 2
    with open(src, encoding="utf-8") as fh:
        prs, geo = build(fh.read(), echo)
    prs.save(out)
    print("Saved:", out)
    print(f"  画布 {SLIDE_W}x{SLIDE_H} 英寸（16:9），共 {len(prs.slides._sldIdLst)} 页")
    cur = None
    for g in geo:
        print("  [%s] %-12s left=%.2f top=%.2f w=%.2f h=%.2f in  %s" % (
            g["slide"], g["shape"][:12], g["left_in"], g["top_in"], g["width_in"], g["height_in"], g["text"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
