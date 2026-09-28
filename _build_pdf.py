#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
作品说明文档 - Markdown 转 PDF（reportlab）
- 嵌入 8 张功能截图
- 中文字体（Noto Sans SC）
- A4 页面，1.5cm 边距
"""

import os
import re
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak, KeepTogether,
    Table, TableStyle, Flowable
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY

WORKDIR = r"C:\Users\zmt\WorkBuddy\2026-09-08-20-01-33"
MD_PATH = os.path.join(WORKDIR, "作品说明文档.md")
PDF_PATH = os.path.join(WORKDIR, "作品说明文档.pdf")

# Register Chinese fonts
NOTO_REG = r"C:\Windows\Fonts\NotoSansSC-Regular.otf"
NOTO_BOLD = r"C:\Windows\Fonts\NotoSansSC-Bold.otf"
SIMHEI = r"C:\Windows\Fonts\simhei.ttf"

# Try Noto Sans SC first
if os.path.exists(NOTO_REG):
    pdfmetrics.registerFont(TTFont("NotoSC", NOTO_REG))
    if os.path.exists(NOTO_BOLD):
        pdfmetrics.registerFont(TTFont("NotoSC-Bold", NOTO_BOLD))
    CN_FONT = "NotoSC"
    CN_FONT_BOLD = "NotoSC-Bold" if os.path.exists(NOTO_BOLD) else "NotoSC"
elif os.path.exists(SIMHEI):
    pdfmetrics.registerFont(TTFont("SimHei", SIMHEI))
    CN_FONT = "SimHei"
    CN_FONT_BOLD = "SimHei"
else:
    raise RuntimeError("No Chinese font found")

# Styles
styles = getSampleStyleSheet()

style_h1 = ParagraphStyle(
    'H1_CN', parent=styles['Heading1'],
    fontName=CN_FONT_BOLD, fontSize=18, leading=24,
    textColor=colors.HexColor('#1a5490'),
    spaceBefore=16, spaceAfter=10, alignment=TA_LEFT,
)

style_h2 = ParagraphStyle(
    'H2_CN', parent=styles['Heading2'],
    fontName=CN_FONT_BOLD, fontSize=14, leading=20,
    textColor=colors.HexColor('#2c3e50'),
    spaceBefore=12, spaceAfter=8, alignment=TA_LEFT,
)

style_h3 = ParagraphStyle(
    'H3_CN', parent=styles['Heading3'],
    fontName=CN_FONT_BOLD, fontSize=12, leading=18,
    textColor=colors.HexColor('#34495e'),
    spaceBefore=8, spaceAfter=6, alignment=TA_LEFT,
)

style_body = ParagraphStyle(
    'Body_CN', parent=styles['BodyText'],
    fontName=CN_FONT, fontSize=10.5, leading=17,
    alignment=TA_JUSTIFY, spaceAfter=6,
    firstLineIndent=21,  # 2-character indent (Chinese style)
)

style_body_noindent = ParagraphStyle(
    'BodyNoIndent_CN', parent=style_body,
    firstLineIndent=0,
)

style_caption = ParagraphStyle(
    'Caption_CN', parent=styles['BodyText'],
    fontName=CN_FONT, fontSize=9, leading=12,
    alignment=TA_CENTER, spaceAfter=10, spaceBefore=4,
    textColor=colors.HexColor('#666666'),
)

style_code = ParagraphStyle(
    'Code_CN', parent=styles['Code'],
    fontName=CN_FONT, fontSize=9, leading=14,
    backColor=colors.HexColor('#f4f4f4'),
    borderColor=colors.HexColor('#dddddd'), borderWidth=0.5, borderPadding=6,
    leftIndent=10, rightIndent=10,
    spaceBefore=4, spaceAfter=4,
)

style_quote = ParagraphStyle(
    'Quote_CN', parent=styles['BodyText'],
    fontName=CN_FONT, fontSize=10, leading=16,
    leftIndent=18, rightIndent=18,
    textColor=colors.HexColor('#555555'),
    backColor=colors.HexColor('#fafafa'),
    borderColor=colors.HexColor('#cccccc'), borderWidth=0,
    borderPadding=8, spaceBefore=4, spaceAfter=6,
)

style_table_cell = ParagraphStyle(
    'TableCell_CN', parent=styles['BodyText'],
    fontName=CN_FONT, fontSize=9, leading=13,
    alignment=TA_LEFT, spaceAfter=0, spaceBefore=0,
)

style_table_header = ParagraphStyle(
    'TableHeader_CN', parent=styles['BodyText'],
    fontName=CN_FONT_BOLD, fontSize=9, leading=13,
    alignment=TA_CENTER, textColor=colors.white,
    spaceAfter=0, spaceBefore=0,
)

style_meta = ParagraphStyle(
    'Meta_CN', parent=styles['BodyText'],
    fontName=CN_FONT, fontSize=10, leading=14,
    textColor=colors.HexColor('#444444'),
    alignment=TA_LEFT, spaceAfter=2,
)

style_cover_title = ParagraphStyle(
    'CoverTitle', parent=styles['Title'],
    fontName=CN_FONT_BOLD, fontSize=28, leading=38,
    alignment=TA_CENTER, textColor=colors.HexColor('#1a5490'),
    spaceAfter=12,
)

style_cover_sub = ParagraphStyle(
    'CoverSub', parent=styles['BodyText'],
    fontName=CN_FONT, fontSize=14, leading=22,
    alignment=TA_CENTER, textColor=colors.HexColor('#666666'),
    spaceAfter=4,
)


def parse_inline(text):
    """Convert Markdown inline to reportlab inline tags."""
    if not text:
        return ""
    # Escape XML
    text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    # Bold
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    # Italic
    text = re.sub(r'\*(.+?)\*', r'<i>\1</i>', text)
    # Inline code
    text = re.sub(r'`([^`]+?)`', r'<font name="' + CN_FONT + '" color="#c0392b">\1</font>', text)
    return text


def make_table(lines):
    """Parse a Markdown table (lines starting with |) into reportlab Table."""
    rows = []
    for i, line in enumerate(lines):
        cells = [c.strip() for c in line.strip().strip('|').split('|')]
        rows.append(cells)
    # Detect header alignment row (---)
    if len(rows) >= 2 and re.match(r'^[\s\-:|]+$', lines[1]):
        header = rows[0]
        data_rows = rows[2:]
    else:
        header = None
        data_rows = rows

    # Build paragraph cells
    flow_rows = []
    if header:
        flow_rows.append([Paragraph(parse_inline(c), style_table_header) for c in header])
    for r in data_rows:
        flow_rows.append([Paragraph(parse_inline(c), style_table_cell) for c in r])

    # Column widths (auto-fit A4 width ~17cm)
    ncols = max(len(r) for r in flow_rows)
    avail_w = 17 * cm
    col_w = avail_w / ncols

    tbl = Table(flow_rows, colWidths=[col_w] * ncols, repeatRows=1 if header else 0)
    tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a5490')) if header else (),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('ROWBACKGROUNDS', (0, 1 if header else 0), (-1, -1),
         [colors.white, colors.HexColor('#f7f9fc')]),
    ]))
    return tbl


def add_image(path, max_w_cm=15.5, max_h_cm=20):
    """Insert an image, scaled to fit within max_w × max_h while keeping aspect ratio."""
    from PIL import Image as PILImage
    pil = PILImage.open(path)
    w_px, h_px = pil.size
    dpi = 96
    w_in = w_px / dpi
    h_in = h_px / dpi
    w_cm = w_in * 2.54
    h_cm = h_in * 2.54
    # Scale to fit within bounding box, preserving aspect ratio
    ratio_w = max_w_cm / w_cm if w_cm > max_w_cm else 1
    ratio_h = max_h_cm / h_cm if h_cm > max_h_cm else 1
    ratio = min(ratio_w, ratio_h)
    if ratio < 1:
        w_cm = w_cm * ratio
        h_cm = h_cm * ratio
    img = Image(path, width=w_cm * cm, height=h_cm * cm)
    return img


def build_pdf():
    doc = SimpleDocTemplate(
        PDF_PATH, pagesize=A4,
        leftMargin=1.8 * cm, rightMargin=1.8 * cm,
        topMargin=1.8 * cm, bottomMargin=1.8 * cm,
        title="C语智学—作品说明文档",
        author="小C老师·C语言选择结构智能助教",
    )

    story = []

    with open(MD_PATH, 'r', encoding='utf-8') as f:
        lines = f.read().splitlines()

    i = 0
    n = len(lines)

    # Cover page
    story.append(Spacer(1, 4 * cm))
    story.append(Paragraph("作品说明文档", style_cover_title))
    story.append(Spacer(1, 0.4 * cm))
    story.append(Paragraph("C语智学 — C语言选择结构课程智能助教", style_cover_sub))
    story.append(Paragraph('"小C老师"', style_cover_sub))
    story.append(Spacer(1, 1.5 * cm))

    # Meta block
    meta_rows = [
        ["参赛赛道", "赛道一 · 智能体开发"],
        ["作品类别", "教师组 / 高等教育（含高职） / 个人参赛"],
        ["文档版本", "v1.0（2026年9月）"],
        ["对应文件", "作品说明文档.pdf"],
    ]
    meta_tbl = Table(
        [[Paragraph(f"<b>{k}</b>", style_table_cell),
          Paragraph(v, style_table_cell)] for k, v in meta_rows],
        colWidths=[4 * cm, 12 * cm]
    )
    meta_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#eaf2fa')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(meta_tbl)

    story.append(Spacer(1, 1.5 * cm))

    # Cover summary box
    cover_note = (
        "本文档对应大赛要求「作品说明文档（PDF，≥1500字，其中AI工具运用≥300字）」。"
        "全文共 8 个章节，覆盖：作品简介、创作思路、AI工具名称及版本、"
        "提示词设计思路、人工干预与修改说明、应用场景、功能介绍、使用说明，"
        "并嵌入 8 张功能演示截图。"
    )
    story.append(Paragraph(cover_note, ParagraphStyle(
        'CoverNote', parent=style_body_noindent,
        fontSize=10, leading=16,
        textColor=colors.HexColor('#333333'),
        backColor=colors.HexColor('#fffbe6'),
        borderColor=colors.HexColor('#ffe58f'),
        borderWidth=1, borderPadding=10,
        alignment=TA_LEFT,
    )))

    story.append(PageBreak())

    # Skip the H1 title in body since we have cover page
    skip_until_h2 = True
    in_table = False
    table_buffer = []

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # Skip horizontal rule
        if stripped == '---':
            i += 1
            continue

        # Heading 1
        if stripped.startswith('# '):
            if skip_until_h2:
                i += 1
                continue
            text = parse_inline(stripped[2:])
            story.append(Spacer(1, 0.3 * cm))
            story.append(Paragraph(text, style_h1))
            i += 1
            continue

        # Heading 2
        if stripped.startswith('## '):
            text = parse_inline(stripped[3:])
            story.append(Spacer(1, 0.3 * cm))
            story.append(Paragraph(text, style_h2))
            i += 1
            continue

        # Heading 3
        if stripped.startswith('### '):
            text = parse_inline(stripped[4:])
            story.append(Paragraph(text, style_h3))
            i += 1
            continue

        # Image: ![alt](path)
        m = re.match(r'!\[(.*?)\]\((.*?)\)', stripped)
        if m:
            alt = m.group(1)
            img_path = m.group(2)
            full_path = img_path if os.path.isabs(img_path) else os.path.join(WORKDIR, img_path)
            if os.path.exists(full_path):
                try:
                    img = add_image(full_path, max_w_cm=15)
                    story.append(Spacer(1, 0.2 * cm))
                    story.append(KeepTogether([img, Paragraph(parse_inline(alt), style_caption)]))
                except Exception as e:
                    print(f"Image load failed: {full_path} - {e}")
            i += 1
            continue

        # Table
        if stripped.startswith('|') and stripped.endswith('|'):
            table_buffer = [stripped]
            j = i + 1
            while j < n and lines[j].strip().startswith('|') and lines[j].strip().endswith('|'):
                table_buffer.append(lines[j].strip())
                j += 1
            try:
                story.append(Spacer(1, 0.15 * cm))
                story.append(make_table(table_buffer))
                story.append(Spacer(1, 0.15 * cm))
            except Exception as e:
                print(f"Table failed: {e}")
                for tb in table_buffer:
                    story.append(Paragraph(parse_inline(tb), style_body))
            i = j
            continue

        # Code block
        if stripped.startswith('```'):
            lang = stripped[3:].strip()
            j = i + 1
            code_lines = []
            while j < n and not lines[j].strip().startswith('```'):
                code_lines.append(lines[j])
                j += 1
            code_text = '\n'.join(code_lines)
            # Replace < and > in code
            code_text = code_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            code_text = code_text.replace('\n', '<br/>')
            code_text = code_text.replace(' ', '&nbsp;').replace('\t', '&nbsp;' * 4)
            story.append(Paragraph(code_text, style_code))
            i = j + 1
            continue

        # Block quote
        if stripped.startswith('> '):
            quote_lines = []
            j = i
            while j < n and (lines[j].strip().startswith('> ') or lines[j].strip() == ''):
                if lines[j].strip():
                    quote_lines.append(lines[j].strip()[2:])
                j += 1
            quote_text = ' '.join(quote_lines)
            story.append(Paragraph(parse_inline(quote_text), style_quote))
            i = j
            continue

        # Empty line
        if not stripped:
            i += 1
            continue

        # Regular paragraph
        para_lines = [stripped]
        j = i + 1
        while j < n and lines[j].strip() and not (
            lines[j].strip().startswith('#') or
            lines[j].strip().startswith('|') or
            lines[j].strip().startswith('```') or
            lines[j].strip().startswith('> ') or
            lines[j].strip().startswith('![') or
            lines[j].strip() == '---'
        ):
            para_lines.append(lines[j].strip())
            j += 1
        text = ' '.join(para_lines)
        # Detect bullets
        if re.match(r'^\d+\.\s', text):
            story.append(Paragraph(parse_inline(text), style_body_noindent))
        elif text.startswith('- '):
            story.append(Paragraph('• ' + parse_inline(text[2:]), style_body_noindent))
        else:
            story.append(Paragraph(parse_inline(text), style_body))
        i = j

    doc.build(story)
    print(f"PDF generated: {PDF_PATH}")
    size_kb = os.path.getsize(PDF_PATH) / 1024
    print(f"File size: {size_kb:.1f} KB")


if __name__ == "__main__":
    build_pdf()