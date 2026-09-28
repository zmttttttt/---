# -*- coding: utf-8 -*-
"""
_build_docx.py — 将「作品说明文档.md」渲染为格式规范、可直接编辑的 Word 文档。

设计要点
--------
* 字体：标题 微软雅黑 / 正文 宋体 / 代码 Consolas（均为 Windows 通用字体，便于二次编辑）
* 页面：A4，左右页边距 2.5cm，上下 2.54cm，正文 1.5 倍行距 + 首行缩进 2 字符
* 封面：大标题 + 副标题 + 元数据表 + 摘要框 + 分页
* 表格：Table Grid + 深蓝表头 + 隔行浅蓝底纹 + 列宽按内容自适应
* 图片：等比缩放（最大宽 15.5cm / 最大高 18cm）+ 居中 + 自动图注
* 页眉：非首页显示文档名；页脚：居中页码
"""

import os
import re

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

# ---------------------------------------------------------------- 路径
BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "作品说明文档.md")
OUT = os.path.join(BASE, "作品说明文档.docx")

# ---------------------------------------------------------------- 主题
C_PRIMARY = RGBColor(0x1F, 0x4E, 0x79)   # 深蓝 — 一级标题 / 表头
C_SECONDARY = RGBColor(0x2E, 0x75, 0xB6)  # 中蓝 — 二级标题
C_TEXT = RGBColor(0x26, 0x26, 0x26)      # 正文
C_GRAY = RGBColor(0x59, 0x59, 0x59)      # 辅助文字
C_CODE = RGBColor(0xC7, 0x25, 0x4E)      # 行内代码

FONT_HEI = "微软雅黑"
FONT_SONG = "宋体"
FONT_MONO = "Consolas"

FILL_HEADER = "1F4E79"   # 表头底纹
FILL_ZEBRA = "F5F9FD"    # 隔行底纹
FILL_CODE = "F2F2F2"     # 行内代码底纹
FILL_QUOTE = "F4F8FC"    # 引用底纹
FILL_NOTE = "FFF8E8"     # 封面摘要框

CN_NUM = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]


# ================================================================ 底层工具
def set_font(run, ascii_font=FONT_SONG, ea_font=None, size=None,
             bold=None, italic=None, color=None):
    """统一设置 run 的中西文字体与字号。"""
    ea_font = ea_font or ascii_font
    run.font.name = ascii_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    rFonts.set(qn("w:eastAsia"), ea_font)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def shade_par(par, fill):
    pPr = par._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def shade_run(run, fill):
    rPr = run._element.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    rPr.append(shd)


def shade_cell(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def par_border(par, edge="bottom", sz=8, color="1F4E79", space=4):
    pPr = par._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    el = OxmlElement("w:" + edge)
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    pBdr.append(el)


def set_first_line_chars(par, chars=200):
    """真正的『首行缩进 N 字符』（Word 优先识别 firstLineChars）。"""
    pPr = par._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(qn("w:firstLineChars"), str(chars))


def add_field(par, code):
    """插入域代码（如 PAGE）。"""
    run = par.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = code
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)
    return run


def repeat_header_row(table):
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader")
    hdr.set(qn("w:val"), "true")
    trPr.append(hdr)


# ================================================================ 行内解析
INLINE_RE = re.compile(r"(\*\*.+?\*\*|`[^`]+`)")


def add_inline(par, text, ascii_font=FONT_SONG, ea_font=FONT_SONG, size=11,
               bold=False, color=C_TEXT, italic=False):
    """把 Markdown 行内语法（**粗体** / `代码`）写入段落。"""
    # 行内三反引号代码（如 ```c ... ```）先降级为单反引号
    text = re.sub(r"```(\w*)\s*\.\.\.\s*```", lambda m: "`%s ...`" % m.group(1), text)
    for part in INLINE_RE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            r = par.add_run(part[2:-2])
            set_font(r, ascii_font, ea_font, size, True, italic, color)
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            r = par.add_run(part[1:-1])
            set_font(r, FONT_MONO, FONT_MONO, size - 0.5, bold, False, C_CODE)
            shade_run(r, FILL_CODE)
        else:
            r = par.add_run(part)
            set_font(r, ascii_font, ea_font, size, bold, italic, color)


# ================================================================ 元素构建
def add_h2(doc, text):
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.space_before = Pt(18)
    pf.space_after = Pt(8)
    pf.line_spacing = 1.25
    pf.keep_with_next = True
    add_inline(par, text, FONT_HEI, FONT_HEI, 15, True, C_PRIMARY)
    par_border(par, "bottom", sz=10, color="1F4E79", space=3)
    return par


def add_h3(doc, text):
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.25
    pf.keep_with_next = True
    add_inline(par, text, FONT_HEI, FONT_HEI, 12.5, True, C_SECONDARY)
    return par


def add_body(doc, text):
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = 1.5
    pf.space_after = Pt(6)
    pf.first_line_indent = Pt(22)   # 兜底：2 字符 @11pt
    set_first_line_chars(par, 200)
    add_inline(par, text, size=11)
    return par


def add_list_item(doc, text, ordered=False, level=0):
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.left_indent = Cm(0.85 + level * 0.6)
    pf.first_line_indent = Cm(-0.85)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.45
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if not ordered:
        r = par.add_run("•  ")
        set_font(r, FONT_SONG, FONT_SONG, 11, False, False, C_PRIMARY)
    add_inline(par, text, size=11)
    return par


def add_quote(doc, text):
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.left_indent = Cm(0.9)
    pf.right_indent = Cm(0.5)
    pf.space_before = Pt(6)
    pf.space_after = Pt(8)
    pf.line_spacing = 1.4
    add_inline(par, text, size=10.5, color=RGBColor(0x33, 0x33, 0x33))
    par_border(par, "left", sz=18, color="2E75B6", space=8)
    shade_par(par, FILL_QUOTE)
    return par


def add_code_block(doc, lines):
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.left_indent = Cm(0.6)
    pf.space_before = Pt(6)
    pf.space_after = Pt(8)
    pf.line_spacing = 1.2
    for idx, line in enumerate(lines):
        if idx:
            par.add_run().add_break()
        r = par.add_run(line if line else " ")
        set_font(r, FONT_MONO, FONT_MONO, 9.5, False, False, RGBColor(0x1A, 0x1A, 0x1A))
    par_border(par, "left", sz=12, color="BFBFBF", space=6)
    shade_par(par, FILL_CODE)
    return par


def add_hr(doc):
    par = doc.add_paragraph()
    par.paragraph_format.space_before = Pt(6)
    par.paragraph_format.space_after = Pt(10)
    par.add_run("")
    par_border(par, "bottom", sz=6, color="C9C9C9", space=1)
    return par


def disp_len(s):
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in s)


def estimate_widths(rows, ncol, total_cm=15.8):
    weights = [1.0] * ncol
    for row in rows:
        for i in range(min(ncol, len(row))):
            weights[i] = max(weights[i], disp_len(row[i]))
    total = sum(weights)
    widths = [total_cm * w / total for w in weights]
    floor = total_cm * 0.09
    widths = [max(w, floor) for w in widths]
    scale = total_cm / sum(widths)
    return [w * scale for w in widths]


def add_table(doc, rows):
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        return None
    ncol = len(rows[0])
    table = doc.add_table(rows=0, cols=ncol)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    widths = estimate_widths(rows, ncol)

    for ri, row in enumerate(rows):
        cells = table.add_row().cells
        for ci in range(ncol):
            val = row[ci] if ci < len(row) else ""
            cell = cells[ci]
            cell.width = Cm(widths[ci])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            par = cell.paragraphs[0]
            pf = par.paragraph_format
            pf.space_before = Pt(3)
            pf.space_after = Pt(3)
            pf.line_spacing = 1.2
            pf.alignment = (WD_ALIGN_PARAGRAPH.CENTER if ri == 0
                            else WD_ALIGN_PARAGRAPH.LEFT)
            if ri == 0:
                add_inline(par, val, FONT_HEI, FONT_HEI, 10, True,
                           RGBColor(0xFF, 0xFF, 0xFF))
                shade_cell(cell, FILL_HEADER)
            else:
                add_inline(par, val, FONT_SONG, FONT_SONG, 10)
                if ri % 2 == 0:
                    shade_cell(cell, FILL_ZEBRA)

    repeat_header_row(table)
    # 表后留白
    tail = doc.add_paragraph()
    tail.paragraph_format.space_after = Pt(4)
    tail.paragraph_format.line_spacing = 0.6
    return table


def add_image(doc, rel_path, caption=None, max_w_cm=15.5, max_h_cm=18.0):
    path = os.path.join(BASE, *rel_path.split("/"))
    if not os.path.exists(path):
        print("  [warn] 图片缺失:", path)
        return
    with Image.open(path) as im:
        w_px, h_px = im.size
    ratio = h_px / float(w_px)
    w_cm = max_w_cm
    h_cm = w_cm * ratio
    if h_cm > max_h_cm:
        h_cm = max_h_cm
        w_cm = h_cm / ratio

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(8)
    par.paragraph_format.space_after = Pt(2)
    par.paragraph_format.keep_with_next = True
    par.add_run().add_picture(path, width=Cm(w_cm))

    if caption:
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_after = Pt(14)
        r = cp.add_run(caption)
        set_font(r, FONT_HEI, FONT_HEI, 9, False, False, C_GRAY)


def add_note_box(doc, text, fill=FILL_NOTE, border="E8B84B"):
    """摘要框：单元格 + 底纹 + 边框。"""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    cell.width = Cm(15.8)
    shade_cell(cell, fill)
    # 自定义边框
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "8")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), border)
        borders.append(el)
    tcPr.append(borders)
    par = cell.paragraphs[0]
    par.paragraph_format.line_spacing = 1.45
    par.paragraph_format.space_before = Pt(4)
    par.paragraph_format.space_after = Pt(4)
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_inline(par, text, FONT_SONG, FONT_SONG, 10.5)
    return table


def add_meta_table(doc, pairs):
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for label, value in pairs:
        cells = table.add_row().cells
        cells[0].width = Cm(3.6)
        cells[1].width = Cm(12.2)
        cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shade_cell(cells[0], "EAF2FA")

        p0 = cells[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(3)
        p0.paragraph_format.space_after = Pt(3)
        r0 = p0.add_run(label)
        set_font(r0, FONT_HEI, FONT_HEI, 10.5, True, False, C_PRIMARY)

        p1 = cells[1].paragraphs[0]
        p1.paragraph_format.space_before = Pt(3)
        p1.paragraph_format.space_after = Pt(3)
        add_inline(p1, value, FONT_SONG, FONT_SONG, 10.5)
    return table


# ================================================================ 封面
COVER_NOTE = (
    "本文档依据大赛作品提交要求编写，共八个部分：一、作品简介；二、创作思路；"
    "三、AI 工具名称及版本；四、提示词（Prompts）设计思路；五、人工干预与修改部分说明；"
    "六、应用场景；七、功能介绍；八、使用说明，并附作品文件清单。全文约 4700 字，"
    "其中「三、AI 工具名称及版本」下设的「AI 工具运用情况」专节约 1000 字，"
    "第七部分嵌入 8 张真实运行界面截图。"
)


def build_cover(doc, meta_pairs):
    for _ in range(2):
        par = doc.add_paragraph()
        par.paragraph_format.space_after = Pt(0)

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_after = Pt(10)
    r = par.add_run("作品说明文档")
    set_font(r, FONT_HEI, FONT_HEI, 30, True, False, C_PRIMARY)

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_after = Pt(4)
    r = par.add_run("C语智学 —— C语言选择结构课程智能助教")
    set_font(r, FONT_HEI, FONT_HEI, 15, True, False, C_SECONDARY)

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_after = Pt(18)
    r = par.add_run("（「小C老师」）")
    set_font(r, FONT_HEI, FONT_HEI, 12, False, False, C_GRAY)

    rule = doc.add_paragraph()
    rule.paragraph_format.space_after = Pt(20)
    rule.add_run("")
    par_border(rule, "bottom", sz=12, color="1F4E79", space=1)

    add_meta_table(doc, meta_pairs)

    gap = doc.add_paragraph()
    gap.paragraph_format.space_after = Pt(8)

    add_note_box(doc, COVER_NOTE)

    doc.add_page_break()


# ================================================================ 页眉页脚
def build_header_footer(doc):
    section = doc.sections[0]
    section.different_first_page_header_footer = True

    hp = section.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = hp.add_run("C语智学 · C语言选择结构课程智能助教 · 作品说明文档")
    set_font(hr, FONT_HEI, FONT_HEI, 8.5, False, False, C_GRAY)
    par_border(hp, "bottom", sz=4, color="BFBFBF", space=2)

    fp = section.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = fp.add_run("— 第 ")
    set_font(r0, FONT_SONG, FONT_SONG, 9, False, False, C_GRAY)
    rf = add_field(fp, "PAGE")
    set_font(rf, FONT_SONG, FONT_SONG, 9, False, False, C_GRAY)
    r1 = fp.add_run(" 页 / 共 ")
    set_font(r1, FONT_SONG, FONT_SONG, 9, False, False, C_GRAY)
    rf2 = add_field(fp, "NUMPAGES")
    set_font(rf2, FONT_SONG, FONT_SONG, 9, False, False, C_GRAY)
    r2 = fp.add_run(" 页 —")
    set_font(r2, FONT_SONG, FONT_SONG, 9, False, False, C_GRAY)


# ================================================================ 主流程
# 源 md 文件被外部编辑器锁定时无法直接改写，故在构建时打补丁，
# 使 7.8 节与重新采集的截图（08_答题解析与错因诊断.png）保持一致。
PATCH_78 = {
    "old": (
        '### 7.8 📋 任务详情页\n'
        '\n'
        '![任务详情学习页](作品说明文档_截图/08_任务详情学习页.png)\n'
        '\n'
        '任务详情页展示完整任务描述、思路引导、参考代码、运行结果，'
        '便于学生从"看题 → 思考 → 编码 → 验证"完整闭环。'
    ),
    "new": (
        '### 7.8 🔍 答题解析与错因诊断\n'
        '\n'
        '![答题解析与错因诊断](作品说明文档_截图/08_答题解析与错因诊断.png)\n'
        '\n'
        '学生提交答案后，系统即时判分并用红/绿双色标注对错选项，'
        '随后"张老师"给出针对该题的文字解析，说明"为什么选、为什么错"。'
        '答错时额外出现"🤖 追问张老师（AI 深度讲解）"按钮，'
        '点击可让 AI 结合本题考点、学生作答记录与历史错题进行个性化错因诊断，'
        '生成"总体表现点评 → 逐知识点掌握分析 → 错题回顾与错因推断 → '
        '专属三步提升计划 → 单元测验成绩预测"五段式诊断报告，'
        '实现从"判分"到"纠因"再到"提升"的完整闭环。'
    ),
}


def apply_patches(raw):
    old, new = PATCH_78["old"], PATCH_78["new"]
    if old in raw:
        raw = raw.replace(old, new)
        print("[patch] 7.8 节已更新为“答题解析与错因诊断”")
    elif new in raw:
        print("[patch] 7.8 节已是最新，跳过")
    else:
        print("[patch][warn] 未找到 7.8 节原文，请人工核对")
    return raw


def build():
    with open(SRC, "r", encoding="utf-8") as fh:
        raw = fh.read()
    raw = apply_patches(raw)
    lines = raw.split("\n")

    doc = Document()
    # 全局默认样式
    normal = doc.styles["Normal"]
    normal.font.name = FONT_SONG
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_SONG)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.5

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)

    meta_pairs = []
    cover_built = False
    seen_h2 = False
    h2_idx = 0
    fig_idx = 0

    i = 0
    total = len(lines)
    while i < total:
        line = lines[i].rstrip()
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        # ---- 分隔线
        if stripped == "---":
            if not cover_built:
                build_cover(doc, meta_pairs)
                cover_built = True
            else:
                add_hr(doc)
            i += 1
            continue

        # ---- 文档主标题（已用于封面）
        if line.startswith("# "):
            i += 1
            continue

        # ---- 封面元数据（**标签**：值）
        if not seen_h2 and not cover_built:
            m = re.match(r"^\*\*(.+?)\*\*[：:]\s*(.*)$", stripped)
            if m:
                meta_pairs.append((m.group(1), m.group(2)))
                i += 1
                continue

        # ---- 标题
        if line.startswith("## "):
            seen_h2 = True
            h2_idx += 1
            fig_idx = 0
            add_h2(doc, line[3:].strip())
            i += 1
            continue

        if line.startswith("### "):
            add_h3(doc, line[4:].strip())
            i += 1
            continue

        # ---- 引用块
        if line.startswith("> "):
            buf = []
            while i < total and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip().lstrip(">").strip())
                i += 1
            add_quote(doc, " ".join(x for x in buf if x))
            continue

        # ---- 代码块
        if stripped.startswith("```"):
            fence = stripped[:3]
            i += 1
            buf = []
            while i < total and not lines[i].strip().startswith("```"):
                buf.append(lines[i].rstrip())
                i += 1
            i += 1
            add_code_block(doc, buf)
            continue

        # ---- 图片
        m = re.match(r"^!\[(.*?)\]\((.+?)\)\s*$", stripped)
        if m:
            alt, rel = m.group(1), m.group(2)
            fig_idx += 1
            prefix = CN_NUM[h2_idx - 1] if 1 <= h2_idx <= len(CN_NUM) else str(h2_idx)
            caption = "图 %s-%d  %s" % (prefix, fig_idx, alt)
            add_image(doc, rel, caption)
            i += 1
            continue

        # ---- 表格
        if stripped.startswith("|"):
            block = []
            while i < total and lines[i].strip().startswith("|"):
                block.append(lines[i].strip())
                i += 1
            rows = []
            for row in block:
                cells = [c.strip() for c in row.strip("|").split("|")]
                if all(re.fullmatch(r":?-{2,}:?", c) or c == "" for c in cells):
                    continue  # 分隔行
                rows.append(cells)
            add_table(doc, rows)
            continue

        # ---- 有序列表
        m = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        if m:
            add_list_item(doc, m.group(2), ordered=True)
            i += 1
            continue

        # ---- 无序列表
        m = re.match(r"^[-*+]\s+(.*)$", stripped)
        if m:
            add_list_item(doc, m.group(1), ordered=False)
            i += 1
            continue

        # ---- 普通段落（合并连续行）
        buf = [stripped]
        i += 1
        while i < total:
            nxt = lines[i].strip()
            if (not nxt or nxt == "---" or nxt.startswith("#")
                    or nxt.startswith("|") or nxt.startswith(">")
                    or nxt.startswith("```") or nxt.startswith("![")
                    or re.match(r"^[-*+]\s+", nxt) or re.match(r"^\d+\.\s+", nxt)):
                break
            buf.append(nxt)
            i += 1
        add_body(doc, "".join(buf))

    if not cover_built:
        build_cover(doc, meta_pairs)

    build_header_footer(doc)
    doc.save(OUT)
    print("已生成:", OUT)
    print("大小: %.1f KB" % (os.path.getsize(OUT) / 1024.0))


if __name__ == "__main__":
    build()
