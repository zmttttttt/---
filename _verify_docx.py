# -*- coding: utf-8 -*-
"""校验生成的 docx：段落/标题/表格/图片/字数。"""
import os
import re
import sys
import zipfile

from docx import Document

BASE = os.path.dirname(os.path.abspath(__file__))
DOCX = os.path.join(BASE, "作品说明文档.docx")

doc = Document(DOCX)
paras = doc.paragraphs
tables = doc.tables

print("段落总数:", len(paras))
print("表格总数:", len(tables))

with zipfile.ZipFile(DOCX) as z:
    media = [n for n in z.namelist() if n.startswith("word/media/")]
print("嵌入图片:", len(media))
for m in sorted(media):
    print("   -", m)

print("\n--- 标题结构 ---")
for p in paras:
    t = p.text.strip()
    if not t:
        continue
    r = p.runs[0] if p.runs else None
    if r is not None and r.font.size and r.font.size.pt >= 12.4 and r.font.bold:
        print("  [%spt] %s" % (r.font.size.pt, t[:60]))

print("\n--- 图片图注 ---")
for p in paras:
    if p.text.strip().startswith("图 "):
        print("  ", p.text.strip())

# 字数统计
def body_text():
    out = []
    for p in paras:
        out.append(p.text)
    for t in tables:
        for row in t.rows:
            for c in row.cells:
                out.append(c.text)
    return "\n".join(out)

txt = body_text()
cjk = len(re.findall(r"[\u3400-\u9FFF]", txt))
latin = len(re.findall(r"[A-Za-z]+", txt))
print("\n汉字数: %d" % cjk)
print("英文词数: %d" % latin)
print("合计字数(汉字+英文词): %d" % (cjk + latin))

# AI 工具运用情况小节字数
m = re.search(r"（一）作品原型开发的.*?（四）AI 工具不可替代的环节。.*?不是\"创意来源\"。",
              txt, re.S)
print("AI 小节命中:", bool(m))
if m:
    seg = m.group(0)
    c = len(re.findall(r"[\u3400-\u9FFF]", seg))
    l = len(re.findall(r"[A-Za-z]+", seg))
    print("AI 小节字数: %d (汉字 %d + 英文词 %d)" % (c + l, c, l))
