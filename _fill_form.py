# -*- coding: utf-8 -*-
import sys, io, copy, shutil, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from docx import Document
from docx.text.paragraph import Paragraph

SRC = r'C:/Users/zmt/Downloads/锡教函〔2026〕171号  市教育局关于举办无锡市“人工智能+教育“创新应用技能大赛的通知_641620260829092757/锡教函〔2026〕171号  市教育局关于举办无锡市“人工智能+教育“创新应用技能大赛的通知_641620260829092757/附件2：报名表及承诺书.docx'
OUT_DIR = r'C:\Users\zmt\WorkBuddy\2026-09-08-20-01-33'
OUT = os.path.join(OUT_DIR, '无锡市“人工智能+教育”大赛-报名表及承诺书（C语智学）.docx')

doc = Document(SRC)
t = doc.tables[0]
rows = t.rows


def val_cell(ri):
    """取第 ri 行的值单元格（第2列）"""
    return rows[ri].cells[1]


# 取一个内容单元格的 run 格式作为模板（R1C1 p0 第一个 run，正文体、不加粗）
TEMPLATE_RPR = copy.deepcopy(val_cell(1).paragraphs[0].runs[0]._r.rPr)


def set_text(p, text):
    """保留段落首个 run 的格式，替换文字并删除多余 run；无 run 时新建并套用模板格式"""
    runs = p.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        newr = p.add_run(text)
        if TEMPLATE_RPR is not None:
            newr._r.insert(0, copy.deepcopy(TEMPLATE_RPR))
    return p


def fill_multiline(cell, lines, drop_extra=True):
    """把 cell 填成多段：复用原段落，不够则克隆原段落格式追加"""
    paras = list(cell.paragraphs)
    base = paras[0]
    made = []
    for i, line in enumerate(lines):
        if i < len(paras):
            p = paras[i]
        else:
            newp_el = copy.deepcopy(base._p)
            paras[-1]._p.addnext(newp_el)
            paras.append(Paragraph(newp_el, cell))
            p = paras[-1]
        set_text(p, line)
        made.append(p)
    # 处理多余的占位段落
    for p in paras[len(lines):]:
        if drop_extra and not p.text.strip():
            p._p.getparent().remove(p._p)   # 纯空行直接删掉，避免行高浪费
        else:
            set_text(p, '')
    return made


# ========== R0 作品名称 ==========
p = set_text(val_cell(0).paragraphs[0], 'C语智学——C语言选择结构课程智能助教')
for r in p.runs:
    r.font.bold = True

# ========== R1 参赛组别 ==========
set_text(val_cell(1).paragraphs[0], '☑ 教师组         □ 学生组')

# ========== R2 参赛赛道 ==========
fill_multiline(val_cell(2), [
    '☑ 赛道一：智能体开发                        □ 赛道二：AI研发创作',
    '□ 赛道三：AI创新创业（OPC）',
])

# ========== R3 参赛人员/团队负责人 ==========
fill_multiline(val_cell(3), [
    '姓名：【请填写】',
    '单位/学校：【请填写】',
    '年级班级：／（教师组不适用）',
    '手机号码：【请填写】',
])

# ========== R4 团队成员 ==========
fill_multiline(val_cell(4), ['无（个人参赛，共 1 人）'])

# ========== R5 指导教师（学生组填）==========
fill_multiline(val_cell(5), ['／（教师组不适用）'])

# ========== R6 使用的AI工具/平台 ==========
fill_multiline(val_cell(6), [
    '1. 腾讯 WorkBuddy（大赛推荐开发平台）：用于需求拆解、代码生成、测试用例生成与文档撰写；',
    '2. 智谱 AI GLM-4-Flash（国产大模型，免费开放）：作为智能体对话引擎，承担多轮答疑、代码纠错与学情诊断；',
    '3. 腾讯混元 hunyuan-lite（国产大模型）：备选对话引擎，保障服务可用性；',
    '4. HTML5 + 原生 JavaScript：智能体前端实现，无第三方框架；Python 3：作品资料与文档处理。',
])

# ========== R7 赛道方向 ==========
fill_multiline(val_cell(7), ['AI+学科教学（C语言程序设计学科教学与个性化学习辅导）'])

# ========== R8 应用网址 ==========
fill_multiline(val_cell(8), [
    '暂无公开线上地址。作品为单文件离线应用，支持本地直接运行与校园局域网部署，'
    '部署环境、运行方式及使用手册详见作品说明文档《使用说明》章节。'
])

# ========== R9 适用学段 ==========
c9 = val_cell(9)
for p in c9.paragraphs:
    for r in p.runs:
        if '高等教育（含高职）' in r.text:
            r.text = r.text.replace('□ 高等教育（含高职）', '☑ 高等教育（含高职）')

# ========== R10 作品简介 ==========
INTRO = [
    '《C语言程序设计》“选择结构”单元抽象难懂、易错点密集（if 误用赋值号、switch 漏 break、else 配对、'
    '70<x<80 连写比较等），教师难以逐人追踪错因。本作品从单元真实教学需求出发，开发“C语智学”'
    'C语言选择结构课程智能助教。',
    '作品基于国产大模型（智谱 GLM-4-Flash、腾讯混元），通过提示词工程塑造“张老师”教师人格：亲切耐心、'
    '启发式引导、善用生活化案例（身高预测、闰年判断、门票打折），支持多轮追问、代码纠错与知识点溯源。'
    '内置 5 大任务闯关、7 个知识点讲堂、20 道原创练习与即时批改，内容全部源自本单元教案与学生高频错题。',
    '学情分析模块自动采集答题数据，生成知识点掌握雷达图、薄弱点诊断与测验成绩预测；接入大模型后可一键生成'
    'AI 深度诊断报告，逐题分析错因并给出三步提升计划。作品为单文件离线应用，断网亦可演示，API 异常自动'
    '回退本地知识库，课堂使用稳定可靠。',
    '应用效果：学生自主闯关并获得个性化讲解，教师获取班级薄弱点数据用于精准教学，显著减轻重复答疑负担。',
]
fill_multiline(val_cell(10), INTRO)

n_chars = sum(len(s) for s in INTRO)
print('作品简介字数：', n_chars)

doc.save(OUT)
print('已保存：', OUT)

# 校验
d2 = Document(OUT)
t2 = d2.tables[0]
print('\n===== 填写结果预览 =====')
for ri in range(12):
    cells = t2.rows[ri].cells
    seen = set()
    out = []
    for c in cells:
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc))
        out.append(c.text.replace('\n', ' ⏎ '))
    print(f'R{ri}: ' + ' || '.join(out))
