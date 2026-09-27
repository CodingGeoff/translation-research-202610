# -*- coding: utf-8 -*-
"""
问卷调查法 正反例 · 学院模板版 V4
- 保留原模板完整 logo（不删、不裁剪）
- 04 对照 / 05 启示 拆细分页，用大白话讲清楚，让没看过论文的人也能看懂
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import deck_lib as D
from deck_lib import Presentation, Inches, MSO_ANCHOR, MSO_SHAPE, qn

D.F_REG = r"C:\Windows\Fonts\msyh.ttc"
D.F_BOLD = r"C:\Windows\Fonts\msyhbd.ttc"
D.IDX = 0
D.EA = "思源黑体 CN Light"
D.LAT = "Arial"
D.MONO = "Consolas"
EA = D.EA

NAVY  = "2D2D89"
INK   = "1F2438"
SUB   = "4A5266"
GREY  = "7A8293"
FAINT = "A5ACBA"
TEAL  = "0E7C86"
TEAL_LT = "E7F1F1"
GOOD  = "1E7A5C"
GOOD_LT = "E9F2EE"
WARN  = "B4451F"
WARN_LT = "F7EBE5"
GOLD  = "C08A3E"
GOLD_LT = "F5EDDB"
LINE  = "DDE3E9"

TPL = r"d:\10_Workspace\translation-research-202610\翻译研究方法\template\高级翻译学院专用PPT模板修改版.pptx"
OUT = r"d:\10_Workspace\translation-research-202610\翻译研究方法\第1组\问卷调查法正反例_学院模板版_v4.pptx"
ASD = r"d:\10_Workspace\translation-research-202610\翻译研究方法\第1组\assets"

W, H = 13.3333, 7.5
ML = 0.72
CR = W - ML
TOTAL = 35

prs = Presentation(TPL)
blank_layout = prs.slides[0].slide_layout   # 保留原模板完整 logo

sldIdLst = prs.slides._sldIdLst
for sid in list(sldIdLst):
    rId = sid.get(qn('r:id'))
    prs.part.drop_rel(rId)
    sldIdLst.remove(sid)


def S():
    return prs.slides.add_slide(blank_layout)


def rect(slide, x, y, w, h, fill=None, ln=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE, adj=None):
    return D.rect(slide, x, y, w, h, fill, ln, lw, shape, adj)


def txt(slide, x, y, w, h, items, align="l", anchor="top"):
    tb, tf = D.add_tb(slide, x, y, w, h, anchor=anchor)
    if isinstance(items, str):
        items = [(items, 14, INK, False, EA, 1.3, None)]
    for n, it in enumerate(items):
        t, sz, col, bd, fnt, lh, spc = (list(it) + [None] * 7)[:7]
        D.para(tf, t, size=sz, color=col, bold=bd, font=fnt or EA, lh=lh or 1.3,
               align=align, first=(n == 0), spc=spc)
    return tb, tf


def foot(slide, page_no, chap_label):
    txt(slide, ML, 7.02, 9.0, 0.24,
        [("问卷调查法正反例 · %s" % chap_label, 9, FAINT, False, EA, 1.2, 0.5)])
    txt(slide, 9.0, 7.02, CR - 9.0, 0.24,
        [("%02d / %02d" % (page_no, TOTAL), 9, FAINT, False, D.MONO, 1.2, 0.5)], align="r")


def add_pic(slide, path, x, y, h, caption=None):
    from PIL import Image
    im = Image.open(path)
    w = h * im.width / im.height
    slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    rect(slide, x - 0.03, y - 0.03, w + 0.06, h + 0.06, None, LINE, 1.0)
    if caption:
        txt(slide, x, y + h + 0.08, w, 0.3, [(caption, 9, GREY, False, EA, 1.2, None)], align="c")


def divider(num, cn, en, page_no, chap_label):
    s = S()
    txt(s, ML, 1.7, 4.2, 2.4, [(num, 120, NAVY, True, D.MONO, 1.0, None)])
    rect(s, ML + 4.6, 1.95, 0.06, 2.1, TEAL)
    txt(s, ML + 5.0, 2.05, 7.0, 0.95, [(cn, 34, NAVY, True, EA, 1.12, None)])
    txt(s, ML + 5.0, 3.15, 7.0, 0.4, [(en, 15, GREY, False, D.LAT, 1.3, None)])
    foot(s, page_no, chap_label)
    return s


def cpage(page_no, chap_label, title, items, img=None, caption=None, accent=NAVY, title_size=25):
    s = S()
    rect(s, ML, 1.22, 0.09, 0.46, accent)
    txt(s, ML + 0.24, 1.16, 11.0, 0.26, [(chap_label, 10, GREY, True, EA, 1.2, 1.1)])
    txt(s, ML + 0.22, 1.46, 11.4, 0.66, [(title, title_size, NAVY, True, EA, 1.1, None)])
    lx = ML + 0.24
    lw = 6.35 if img else 11.8
    yy = 2.15
    for it in items:
        if isinstance(it, tuple):
            lead, body = it
            txt(s, lx, yy, lw, 0.4, [(lead, 16, accent, True, EA, 1.25, 0.35)])
            yy += 0.4
            if body:
                lines, sz, used = D.fit_wrap(body, lw, 1.4, 14.5, lh=1.32, min_size=13)
                txt(s, lx, yy, lw, used + 0.05, [(body, sz, SUB, False, EA, 1.32, None)])
                yy += used + 0.22
        else:
            lines, sz, used = D.fit_wrap(it, lw, 1.4, 15, lh=1.32, min_size=13.5)
            txt(s, lx, yy, lw, used + 0.05, [("·  " + it, sz, INK, False, EA, 1.32, 0.3)])
            yy += used + 0.28
    if img:
        p = os.path.join(ASD, img)
        if os.path.exists(p):
            add_pic(s, p, 8.1, 2.1, 4.4, caption)
    foot(s, page_no, chap_label)
    return s


def tpage(page_no, chap_label, title, items, accent=NAVY, title_size=25):
    s = S()
    rect(s, ML, 1.22, 0.09, 0.46, accent)
    txt(s, ML + 0.24, 1.16, 11.0, 0.26, [(chap_label, 10, GREY, True, EA, 1.2, 1.1)])
    txt(s, ML + 0.22, 1.46, 11.4, 0.66, [(title, title_size, NAVY, True, EA, 1.1, None)])
    lx = ML + 0.24
    lw = 11.8
    yy = 2.15
    for it in items:
        if isinstance(it, tuple):
            lead, body = it
            txt(s, lx, yy, lw, 0.4, [(lead, 17, accent, True, EA, 1.25, 0.4)])
            yy += 0.44
            if body:
                lines, sz, used = D.fit_wrap(body, lw, 1.5, 15, lh=1.35, min_size=13.5)
                txt(s, lx, yy, lw, used + 0.05, [(body, sz, SUB, False, EA, 1.35, None)])
                yy += used + 0.24
        else:
            lines, sz, used = D.fit_wrap(it, lw, 1.5, 16, lh=1.35, min_size=14)
            txt(s, lx, yy, lw, used + 0.05, [("·  " + it, sz, INK, False, EA, 1.35, 0.32)])
            yy += used + 0.3
    foot(s, page_no, chap_label)
    return s


def compare_page(page_no, chap_label, title, pos, neg, note):
    """对照页：左=正例(绿卡)，右=反例(红卡)，底部=通俗解释"""
    s = S()
    rect(s, ML, 1.22, 0.09, 0.46, NAVY)
    txt(s, ML + 0.24, 1.16, 11.0, 0.26, [(chap_label, 10, GREY, True, EA, 1.2, 1.1)])
    txt(s, ML + 0.22, 1.46, 11.4, 0.66, [(title, 25, NAVY, True, EA, 1.1, None)])
    # 左卡 正例
    lx, lw = ML, 6.0
    D.card(s, lx, 2.2, lw, 3.5, GOOD_LT, None, 0.75)
    rect(s, lx, 2.2, lw, 0.42, GOOD)
    txt(s, lx, 2.2, lw, 0.42, [("正例 · 口译笔记量表", 14, "FFFFFF", True, EA, 1.15, None)], align="c", anchor="mid")
    yy = 2.78
    for it in pos:
        lines, sz, used = D.fit_wrap(it, lw - 0.5, 1.4, 13.5, lh=1.32, min_size=12)
        txt(s, lx + 0.25, yy, lw - 0.5, used + 0.05, [("·  " + it, sz, INK, False, EA, 1.32, 0.26)])
        yy += used + 0.26
    # 右卡 反例
    rx, rw = 6.72, 5.9
    D.card(s, rx, 2.2, rw, 3.5, WARN_LT, None, 0.75)
    rect(s, rx, 2.2, rw, 0.42, WARN)
    txt(s, rx, 2.2, rw, 0.42, [("反例 · MTPE 问卷", 14, "FFFFFF", True, EA, 1.15, None)], align="c", anchor="mid")
    yy = 2.78
    for it in neg:
        lines, sz, used = D.fit_wrap(it, rw - 0.5, 1.4, 13.5, lh=1.32, min_size=12)
        txt(s, rx + 0.25, yy, rw - 0.5, used + 0.05, [("·  " + it, sz, INK, False, EA, 1.32, 0.26)])
        yy += used + 0.26
    # 底部通俗解释
    D.card(s, ML, 5.85, 12.02, 0.95, GOLD_LT, None, 0.75)
    rect(s, ML, 5.85, 0.06, 0.95, GOLD)
    txt(s, ML + 0.3, 5.85, 11.5, 0.95, [(note, 14.5, INK, False, EA, 1.4, None)], anchor="mid")
    foot(s, page_no, chap_label)
    return s


# ============================================================ P1 封面
s = S()
txt(s, ML, 2.1, 12.0, 0.34, [("《翻译研究方法》第 1 组 · 方法 A 案例部分", 12, GREY, False, EA, 1.3, 0.5)])
txt(s, ML, 2.5, 12.0, 0.8, [("问卷调查法在翻译研究中的应用", 40, NAVY, True, EA, 1.12, None)])
txt(s, ML, 3.4, 12.0, 0.8, [("正例与反例", 40, NAVY, True, EA, 1.12, None)])
txt(s, ML, 4.35, 12.0, 0.36, [("Questionnaire Survey in Translation Studies: A Positive and a Negative Case", 14, GREY, False, D.LAT, 1.3, None)])
rect(s, ML, 5.05, 4.6, 0.03, TEAL)
txt(s, ML, 5.3, 12.0, 0.9, [
    ("正例（口译）｜ 周金华、董燕萍 2019 · 口译笔记熟练度量表的开发", 12.5, INK, False, EA, 1.5, None),
    ("反例（笔译）｜ Yang & Mustafa 2022 · On Postediting of Machine Translation", 12.5, INK, False, EA, 1.5, None)])
txt(s, ML, 6.2, 12.0, 0.4, [("汇报人：＿＿ ｜ 2026.10.12", 12, GREY, False, EA, 1.4, None)])

# ============================================================ P2 目录
s = S()
txt(s, ML, 1.3, 6.0, 0.6, [("目录", 30, NAVY, True, EA, 1.1, None)])
txt(s, ML, 2.0, 6.0, 0.36, [("Contents", 13, GREY, False, D.LAT, 1.3, None)])
toc = [
    ("01", "问卷法：定位与判断标准", "Survey in Translation Studies"),
    ("02", "正例：口译笔记熟练度量表", "周金华、董燕萍 2019"),
    ("03", "反例：MTPE 态度问卷", "Yang & Mustafa 2022"),
    ("04", "对照：方法学上的差距", "Side-by-side comparison"),
    ("05", "启示：设计与施测的规范", "Guidelines for questionnaire design"),
]
y = 2.55
for num, cn, en in toc:
    txt(s, ML, y, 1.2, 0.5, [(num, 24, TEAL, True, D.MONO, 1.1, None)])
    txt(s, ML + 1.5, y + 0.02, 8.5, 0.44, [(cn, 18, INK, True, EA, 1.2, None)])
    txt(s, ML + 1.5, y + 0.48, 8.5, 0.26, [(en, 12, GREY, False, D.LAT, 1.2, None)])
    rect(s, ML, y + 0.78, 11.9, 0.012, LINE)
    y += 0.86

# ============================================================ P3 分隔 01
divider("01", "问卷法：定位与判断标准", "Questionnaire Survey in Translation Studies", 3, "问卷法")

# ============================================================ P4 问卷法定位
tpage(4, "问卷法", "什么是问卷调查法", [
    ("定义", "以标准化问卷为工具，系统收集翻译相关数据的定量研究方法。"),
    ("适用场景", "读者对译文的反应与评价、译者行为与策略、翻译教学效果、翻译行业现状。"),
    ("优势", "样本量大、操作标准化、结果便于定量统计。"),
    ("局限", "回收率难保证，缺乏深入定性，结论受问卷质量影响大。"),
], accent=TEAL)

# ============================================================ P5 证据链
tpage(5, "问卷法", "判断标准：一条证据链", [
    ("研究问题", "想回答什么，先写成一句可检验的话。"),
    ("构念与题项", "把抽象概念拆成可回答的题目。"),
    ("抽样与施测", "向谁发放、如何回收、有效多少。"),
    ("统计与结论", "数据能支撑到哪一步，就说到哪一步。"),
], accent=TEAL)

# ============================================================ P6 分隔 02
divider("02", "正例：口译笔记熟练度量表", "Zhou & Dong 2019, Foreign Language Teaching and Research", 6, "正例")

# ============================================================ P7 正例论文
cpage(7, "正例", "研究问题", [
    ("文献", "周金华、董燕萍（2019）《外语教学与研究》51(6): 925-937"),
    ("问题", "口译笔记熟练度「看不见、摸不着」，能否用一份自评量表测出来？"),
    ("思路", "先给它一个可操作的定义，再编制一份能被复核的自评量表。"),
], img="zheng_p01_title.png", caption="论文首页", accent=TEAL)

# ============================================================ P8 正例·明确测量对象
cpage(8, "正例", "第一步：明确测量对象", [
    ("两阶段", "笔记记录（note writing）、笔记理解（note reading）"),
    ("四维度", "听记协调性、记录系统性、时效性、笔记使用"),
    ("原则", "先把抽象概念拆成可观察的维度，再落笔写题项。"),
], img="zheng_p06_itemtable.png", caption="表1 量表维度与题项", accent=TEAL)

# ============================================================ P9 正例题库
tpage(9, "正例", "第二步：生成题库", [
    ("来源", "学员学习日志 + 译员访谈（4 名学员、6 名教师）"),
    ("规模", "初始题库 74 题"),
    ("语言", "题干尽量采用译员自己的表达，避免学术腔。"),
], accent=TEAL)

# ============================================================ P10 正例评审
tpage(10, "正例", "第三步：专家评审", [
    ("专家", "5 名学员译员 + 3 名教师译员"),
    ("过程", "三轮讨论，历时半年"),
    ("结果", "收到 28 题，李克特 6 点量表"),
], accent=TEAL)

# ============================================================ P11 正例删题
cpage(11, "正例", "第四步：试测与删题", [
    ("试测", "54 名翻译专业学生参与试测"),
    ("方法", "极端组比较 + 同质性检验（项目分析）"),
    ("结果", "删 7 题，留 21 题，重新随机排序"),
    ("处理", "改名「自我描述量表」，压低社会赞许性"),
], img="zheng_p07_itemanalysis.png", caption="表2 题项分析结果", accent=TEAL)

# ============================================================ P12 正例信度
cpage(12, "正例", "第五步：正式施测与信度", [
    ("被试", "144 名有近一学年笔记交传训练的本科生"),
    ("信度", "总量表 Cronbach's α = .901，各维度信度也较高"),
], img="zheng_p08_reliability.png", caption="表3 信度检验", accent=TEAL)

# ============================================================ P13 正例效度
cpage(13, "正例", "第六步：效度证据", [
    ("内容效度", "I-CVI / S-CVI 均为 1.00"),
    ("结构效度", "EFA 抽出 4 因子，累计解释 60.49%"),
    ("效标关联", "练习量 .473、动机 .539、绩效 .556（均 p<.01）"),
], img="zheng_p09_efa.png", caption="因素分析结果", accent=TEAL)

# ============================================================ P14 正例小结
tpage(14, "正例", "这篇研究交付了什么", [
    ("工具", "一份可复用的测量工具，21 题随附录公开。"),
    ("证据链", "构念 → 题项 → 抽样 → 信效度，每一步都留下可核查的记录。"),
    ("分寸", "只言「可测、能区分训练阶段」，不下因果结论。"),
], accent=GOOD)

# ============================================================ P15 分隔 03
divider("03", "反例：MTPE 态度问卷", "Yang & Mustafa 2022, Human Behavior and Emerging Technologies", 15, "反例")

# ============================================================ P16 反例论文
cpage(16, "反例", "研究问题", [
    ("文献", "Yang & Mustafa (2022)，Human Behavior and Emerging Technologies"),
    ("问题", "本科翻译专业是否应引入「机器翻译 + 译后编辑（MTPE）」流程？"),
    ("工具", "6 条李克特陈述 + 4 道开放题，借 TAM 框架"),
], img="fan_p01_title.png", caption="论文首页", accent=WARN)

# ============================================================ P17 反例工具
cpage(17, "反例", "工具设计", [
    ("结构", "感知有用性 3 题、感知易用性 3 题，6 点量表"),
    ("开放题", "4 道（如「是否知道 MTPE 已是行业标准」）"),
    ("框架", "借用技术接受模型（TAM），但只借概念、未借量表"),
], img="fan_p05_questionnaire.png", caption="问卷题项原文", accent=WARN)

# ============================================================ P18 反例抽样
cpage(18, "反例", "抽样与施测", [
    ("样本", "127 名本科生（大三 39、大四 88），西北一所高校"),
    ("方式", "便利抽样，未说明抽样框"),
    ("缺口", "发放数、回收数、有效数均未报告"),
], img="fan_p04_participants.png", caption="被试与变量", accent=WARN)

# ============================================================ P19 反例分析
cpage(19, "反例", "数据分析", [
    ("方法", "描述统计 + 单题 Pearson 相关"),
    ("结果", "6 题均值 4.02–4.55，均高于中点 3.5"),
    ("信度", "分维度 α = .714 / .835"),
], img="fan_p06_data.png", caption="数据分析结果", accent=WARN)

# ============================================================ P20 反例断点①
cpage(20, "反例", "断点① 构念—题项脱钩", [
    ("现象", "借了 TAM 的框架，没借它的量表结构"),
    ("问题", "6 条自拟题，无内容评审、无前测"),
    ("后果", "三题是否真落在两个构念上，从未检验"),
], img="fan_p05_questionnaire.png", caption="6 条自拟题原文", accent=WARN)

# ============================================================ P21 反例断点②
cpage(21, "反例", "断点② 抽样口径模糊", [
    ("现象", "便利抽样，发放 / 回收 / 有效均未报告"),
    ("后果", "应答率、自选择偏差无从判断"),
    ("施测", "线上一次填答，无照应题、无注意力检查"),
], img="fan_p04_participants.png", caption="被试说明（仅 127 人）", accent=WARN)

# ============================================================ P22 反例断点③④
tpage(22, "反例", "断点③④ 分析层级错位 · 结论超出数据", [
    ("单题当构念用", "拿单题与课程成绩做相关，均值高于中点即称「积极」。"),
    ("横断面推因果", "从一次性自陈，直接推出「应纳入课程」。"),
], accent=WARN)

# ============================================================ P23 反例小结
tpage(23, "反例", "这篇研究的问题在哪", [
    ("不是造假", "数据是真的，统计也报了。"),
    ("是证据链断裂", "六步都做了，但每一步只写了半句。"),
    ("核心", "结论超出了数据能承担的范围。"),
], accent=WARN)

# ============================================================ P24 分隔 04
divider("04", "对照：方法学上的差距", "A side-by-side comparison", 24, "对照")

# ============================================================ P25-P30 对照（六维度，大白话）
compare_page(25, "对照", "① 题目从哪来", [
    "从学生的学习日志、译员访谈里收集素材",
    "先攒 74 题，专家三轮评审收到 28 题",
    "试测后再删 7 题，定稿 21 题",
], [
    "6 道题全部由作者自己拟写",
    "没有请专家评审",
    "也没有先小规模试测",
], "出考卷的差别：正例是「出题 → 审题 → 试考 → 定稿」，反例是「出题 → 直接拿去考」。")

compare_page(26, "对照", "② 题目测的是什么", [
    "先把「熟练度」定义成「记录 / 理解两阶段 × 四维度」",
    "按维度逐条写题",
    "再用因子分析验证：题目确实落在预设的维度上",
], [
    "借 TAM 框架，声称测「有用性」和「易用性」",
    "但 6 道自拟题从没验证过",
    "说不清它们到底在测什么",
], "题目和「想测的东西」对不对得上：正例对得上、还做了验证；反例没验证，等于心里没底。")

compare_page(27, "对照", "③ 谁来填问卷", [
    "144 名「有近一学年笔记交传训练」的本科生",
    "样本条件写得清清楚楚",
    "读者一看就知道研究对象是谁",
], [
    "127 名「某高校翻译专业」本科生（大三 39、大四 88）",
    "发出去多少份、收回多少份、有效多少份都没写",
    "抽样方式只说「便利抽样」",
], "问卷发给谁、收回多少，决定结论能代表谁：正例说得清，反例说不清。")

compare_page(28, "对照", "④ 结果稳不稳定（信度）", [
    "总量表 Cronbach's α = .901，很高",
    "各维度的信度也都较高",
    "说明反复测，结果稳定可靠",
], [
    "α = .714 / .835，偏低",
    "而且只按两个分量表报告",
    "没有报总量表整体信度",
], "信度就是「尺子量得稳不稳」：α 在 .9 以上很稳，.7 左右只能算勉强及格。")

compare_page(29, "对照", "⑤ 测得到底对不对（效度）", [
    "内容效度：专家评分 CVI = 1.00",
    "结构效度：因子分析验证了四维度",
    "效标关联：与练习量 .473、动机 .539、绩效 .556 都相关",
], [
    "只有「单题 × 课程成绩」的相关",
    "没有系统的效度检验",
    "也没有说明题目是否覆盖了构念",
], "效度就是「测的到底是不是想测的东西」：正例有三重证据，反例几乎只有一条相关。")

compare_page(30, "对照", "⑥ 结论有没有越界", [
    "结论停在「量表可测、能区分训练阶段」",
    "每一步都在数据能支撑的范围内",
], [
    "从「127 人态度偏积极」",
    "直接推出「应把 MTPE 纳入本科教学」",
    "中间跳过了好几步",
], "正例说「这把尺子好用」；反例说「尺子好用，所以应该改课程」——后者多走了两步。")

# ============================================================ P31 分隔 05
divider("05", "启示：设计与施测的规范", "Guidelines for questionnaire design", 31, "启示")

# ============================================================ P32 清单（上）
s = S()
rect(s, ML, 1.22, 0.09, 0.46, GOLD)
txt(s, ML + 0.24, 1.16, 11.0, 0.26, [("启示", 10, GREY, True, EA, 1.2, 1.1)])
txt(s, ML + 0.22, 1.46, 11.4, 0.66, [("交稿前十问（上）", 25, NAVY, True, EA, 1.1, None)])
ck1 = [
    ("1", "我要回答什么问题？", "先把它写成一句可检验的问话，再动手。"),
    ("2", "题目和概念对得上吗？", "先定义要测的概念，再写题，列一张「概念 → 维度 → 题目」对照表。"),
    ("3", "题目是哪来的？", "沿用（注明出处）/ 改编（说明改动）/ 自编（说明怎么来的）。"),
    ("4", "问卷发给谁？", "写清总体、抽样范围，谁符合、谁不符合。"),
    ("5", "三份数字报了吗？", "发放、回收、有效各多少；先定「怎样算无效」。"),
]
y = 2.2
for n, lead, body in ck1:
    rect(s, ML, y, 0.4, 0.4, GOLD, None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.2)
    txt(s, ML, y, 0.4, 0.4, [(n, 13, "FFFFFF", True, D.MONO, 1.2, None)], align="c", anchor="mid")
    txt(s, ML + 0.58, y, 11.3, 0.36, [(lead, 16, NAVY, True, EA, 1.2, None)])
    txt(s, ML + 0.58, y + 0.4, 11.3, 0.4, [(body, 13.5, SUB, False, EA, 1.3, None)])
    y += 0.98
foot(s, 32, "启示")

# ============================================================ P33 清单（下）
s = S()
rect(s, ML, 1.22, 0.09, 0.46, GOLD)
txt(s, ML + 0.24, 1.16, 11.0, 0.26, [("启示", 10, GREY, True, EA, 1.2, 1.1)])
txt(s, ML + 0.22, 1.46, 11.4, 0.66, [("交稿前十问（下）", 25, NAVY, True, EA, 1.1, None)])
ck2 = [
    ("6", "先试测了吗？", "正式发放前，找几个人试填，或请专家把一次关。"),
    ("7", "信效度报了吗？", "报 α 或 ω；问卷分多个维度，还要报 EFA / CFA。"),
    ("8", "指导语写清楚了吗？", "说明用途、匿名、可以中途退出。"),
    ("9", "题型和统计配吗？", "单选用频数、量表用均值、开放题用编码。"),
    ("10", "结论有没有越界？", "横断面问卷只能说「有关联」，不能说「导致」。"),
]
y = 2.2
for n, lead, body in ck2:
    rect(s, ML, y, 0.4, 0.4, GOLD, None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.2)
    txt(s, ML, y, 0.4, 0.4, [(n, 13, "FFFFFF", True, D.MONO, 1.2, None)], align="c", anchor="mid")
    txt(s, ML + 0.58, y, 11.3, 0.36, [(lead, 16, NAVY, True, EA, 1.2, None)])
    txt(s, ML + 0.58, y + 0.4, 11.3, 0.4, [(body, 13.5, SUB, False, EA, 1.3, None)])
    y += 0.98
foot(s, 33, "启示")

# ============================================================ P34 最小改法
tpage(34, "启示", "反例三处，怎么改", [
    ("题目来源", "换成已验证的量表题目并注明出处；坚持自拟，就补一次因子分析（EFA）或报告「每题和总分的相关」。"),
    ("单题当构念", "先把同一维度的题合成平均分、报 α，再做相关；结论措辞用「有关联」，不用「影响」。"),
    ("结论越界", "把结论收缩到「这 127 名学生（这所学校）」，并补上发放 / 回收 / 有效三份数字。"),
], accent=GOLD)

# ============================================================ P35 结尾
s = S()
txt(s, ML, 2.3, 12.0, 0.9, [("谢谢观看", 44, NAVY, True, EA, 1.1, None)])
txt(s, ML, 3.4, 12.0, 0.4, [("Thank You", 18, GREY, False, D.LAT, 1.3, None)])
rect(s, ML, 4.1, 4.6, 0.03, TEAL)
txt(s, ML, 4.35, 12.0, 1.2, [
    ("周金华、董燕萍（2019）. 口译笔记熟练度量表的开发. 外语教学与研究 51(6): 925-937.", 11.5, INK, False, EA, 1.5, None),
    ("Yang, Z. & Mustafa, H. R. (2022). On postediting of machine translation and workflow. HBET, 5793054.", 11.5, INK, False, EA, 1.5, None)])
txt(s, ML, 5.6, 12.0, 0.4, [("汇报人：＿＿", 12, GREY, False, EA, 1.4, None)])

prs.save(OUT)
print("saved:", OUT)
print("slides:", len(prs.slides._sldIdLst))
