# -*- coding: utf-8 -*-
"""
问卷调查法 正反例 · 学院模板版 V3
- 相比 V2：去掉左上角 logo；标题词更通俗；每页适当多放一点字、要点式说清楚
- 结构：章节式（5 章）+ 大编号分隔页 + 每页一个小主题 + 右栏论文截图
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
LINE  = "DDE3E9"

TPL = r"d:\10_Workspace\translation-research-202610\翻译研究方法\template\高级翻译学院专用PPT模板修改版.pptx"
OUT = r"d:\10_Workspace\translation-research-202610\翻译研究方法\第1组\问卷调查法正反例_学院模板版_v3.pptx"
ASD = r"d:\10_Workspace\translation-research-202610\翻译研究方法\第1组\assets"

W, H = 13.3333, 7.5
ML = 0.72
CR = W - ML
TOTAL = 30

prs = Presentation(TPL)
blank_layout = prs.slides[0].slide_layout

# 删除左上角 logo 图片（保留底部格言）
for sh in list(blank_layout.shapes):
    if sh.shape_type == 13:  # PICTURE
        sh._element.getparent().remove(sh._element)

sldIdLst = prs.slides._sldIdLst
for sid in list(sldIdLst):
    rId = sid.get(qn('r:id'))
    prs.part.drop_rel(rId)
    sldIdLst.remove(sid)


def S():
    slide = prs.slides.add_slide(blank_layout)
    # 左上角文字 logo（已裁剪掉左侧图标）
    p = os.path.join(ASD, 'logo_text.png')
    if os.path.exists(p):
        from PIL import Image
        im = Image.open(p)
        lh = 0.5
        lw = lh * im.width / im.height
        slide.shapes.add_picture(p, Inches(0.35), Inches(0.2), Inches(lw), Inches(lh))
    return slide


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
    txt(s, ML, 1.6, 4.2, 2.4, [(num, 120, NAVY, True, D.MONO, 1.0, None)])
    rect(s, ML + 4.6, 1.85, 0.06, 2.1, TEAL)
    txt(s, ML + 5.0, 1.95, 7.0, 0.95, [(cn, 34, NAVY, True, EA, 1.12, None)])
    txt(s, ML + 5.0, 3.05, 7.0, 0.4, [(en, 15, GREY, False, D.LAT, 1.3, None)])
    foot(s, page_no, chap_label)
    return s


def cpage(page_no, chap_label, title, items, img=None, caption=None, accent=NAVY, title_size=25):
    s = S()
    rect(s, ML, 1.12, 0.09, 0.46, accent)
    txt(s, ML + 0.24, 1.06, 11.0, 0.26, [(chap_label, 10, GREY, True, EA, 1.2, 1.1)])
    txt(s, ML + 0.22, 1.36, 11.4, 0.66, [(title, title_size, NAVY, True, EA, 1.1, None)])
    lx = ML + 0.24
    lw = 6.35 if img else 11.8
    yy = 2.05
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
            add_pic(s, p, 8.1, 1.95, 4.55, caption)
    foot(s, page_no, chap_label)
    return s


def tpage(page_no, chap_label, title, items, accent=NAVY, title_size=25):
    s = S()
    rect(s, ML, 1.12, 0.09, 0.46, accent)
    txt(s, ML + 0.24, 1.06, 11.0, 0.26, [(chap_label, 10, GREY, True, EA, 1.2, 1.1)])
    txt(s, ML + 0.22, 1.36, 11.4, 0.66, [(title, title_size, NAVY, True, EA, 1.1, None)])
    lx = ML + 0.24
    lw = 11.8
    yy = 2.05
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


# ============================================================ P1 封面
s = S()
txt(s, ML, 2.0, 12.0, 0.34, [("《翻译研究方法》第 1 组 · 方法 A 案例部分", 12, GREY, False, EA, 1.3, 0.5)])
txt(s, ML, 2.4, 12.0, 0.8, [("问卷调查法在翻译研究中的应用", 40, NAVY, True, EA, 1.12, None)])
txt(s, ML, 3.3, 12.0, 0.8, [("正例与反例", 40, NAVY, True, EA, 1.12, None)])
txt(s, ML, 4.25, 12.0, 0.36, [("Questionnaire Survey in Translation Studies: A Positive and a Negative Case", 14, GREY, False, D.LAT, 1.3, None)])
rect(s, ML, 4.95, 4.6, 0.03, TEAL)
txt(s, ML, 5.2, 12.0, 0.9, [
    ("正例（口译）｜ 周金华、董燕萍 2019 · 口译笔记熟练度量表的开发", 12.5, INK, False, EA, 1.5, None),
    ("反例（笔译）｜ Yang & Mustafa 2022 · On Postediting of Machine Translation", 12.5, INK, False, EA, 1.5, None)])
txt(s, ML, 6.2, 12.0, 0.4, [("汇报人：＿＿ ｜ 2026.10.12", 12, GREY, False, EA, 1.4, None)])

# ============================================================ P2 目录
s = S()
txt(s, ML, 1.1, 6.0, 0.6, [("目录", 30, NAVY, True, EA, 1.1, None)])
txt(s, ML, 1.8, 6.0, 0.36, [("Contents", 13, GREY, False, D.LAT, 1.3, None)])
toc = [
    ("01", "问卷法：定位与判断标准", "Survey in Translation Studies"),
    ("02", "正例：口译笔记熟练度量表", "周金华、董燕萍 2019"),
    ("03", "反例：MTPE 态度问卷", "Yang & Mustafa 2022"),
    ("04", "对照：方法学上的差距", "Side-by-side comparison"),
    ("05", "启示：设计与施测的规范", "Guidelines for questionnaire design"),
]
y = 2.4
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

# ============================================================ P25 对照上
tpage(25, "对照", "两篇并排（上）", [
    ("题项从哪来", "正例：日志＋访谈＋三轮评审，74 → 28 → 21　｜　反例：自拟 6 条，无评审"),
    ("测的是什么", "正例：4 维度构念，EFA 复核　｜　反例：2 构念各 3 题，未检验"),
    ("谁在填", "正例：144 人，纳入条件写明　｜　反例：127 人，回收口径未报"),
], accent=NAVY)

# ============================================================ P26 对照下
tpage(26, "对照", "两篇并排（下）", [
    ("信度", "正例：α = .901　｜　反例：.714 / .835"),
    ("效度", "正例：CVI ＋ EFA ＋ 三项效标　｜　反例：单题 × 课程成绩"),
    ("结论", "正例：可测、能区分阶段　｜　反例：可行、应纳入课程（超出数据）"),
], accent=NAVY)

# ============================================================ P27 分隔 05
divider("05", "启示：设计与施测的规范", "Guidelines for questionnaire design", 27, "启示")

# ============================================================ P28 十条清单
s = S()
rect(s, ML, 1.12, 0.09, 0.46, GOLD)
txt(s, ML + 0.24, 1.06, 11.0, 0.26, [("启示", 10, GREY, True, EA, 1.2, 1.1)])
txt(s, ML + 0.22, 1.36, 11.4, 0.66, [("交稿前的十行自检", 25, NAVY, True, EA, 1.1, None)])
ck = ["研究问题写成一句可回答的话",
      "构念 → 维度 → 题项：一张对齐表",
      "题项来源写清（沿用 / 改编 / 自编）",
      "总体、抽样框、纳入排除各一句",
      "发放 / 回收 / 有效三份数字",
      "无效判定规则先定后用",
      "前测与专家评审留痕",
      "信度 α / ω；多维补 EFA / CFA",
      "指导语含匿名、用途、自愿退出",
      "结论只说关联，不说因果"]
for i, c in enumerate(ck):
    col = i // 5
    row = i % 5
    x = ML + col * 6.05
    y = 2.05 + row * 0.86
    rect(s, x, y, 0.34, 0.34, NAVY if col == 0 else TEAL, None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.2)
    txt(s, x, y, 0.34, 0.34, [("%02d" % (i + 1), 10, "FFFFFF", True, D.MONO, 1.2, None)], align="c", anchor="mid")
    txt(s, x + 0.5, y + 0.03, 5.5, 0.4, [(c, 15, INK, False, EA, 1.25, None)])
txt(s, ML, 6.3, 11.9, 0.4, [("这十条在正例里都有对应段落 —— 抄它的结构，别抄它的内容。", 13, GREY, False, EA, 1.4, None)])
foot(s, 28, "启示")

# ============================================================ P29 最小改法
tpage(29, "启示", "反例的最小改法", [
    ("题项", "改用已验证量表条目并标出处；自编则补 EFA 或报题总相关。"),
    ("分析", "先按维度合成均分、报 α，再做相关；只写「关联」不写「影响」。"),
    ("结论", "收缩到「该校学生」，并报发放 / 回收 / 有效三份数字。"),
], accent=GOLD)

# ============================================================ P30 结尾
s = S()
txt(s, ML, 2.2, 12.0, 0.9, [("谢谢观看", 44, NAVY, True, EA, 1.1, None)])
txt(s, ML, 3.3, 12.0, 0.4, [("Thank You", 18, GREY, False, D.LAT, 1.3, None)])
rect(s, ML, 4.0, 4.6, 0.03, TEAL)
txt(s, ML, 4.25, 12.0, 1.2, [
    ("周金华、董燕萍（2019）. 口译笔记熟练度量表的开发. 外语教学与研究 51(6): 925-937.", 11.5, INK, False, EA, 1.5, None),
    ("Yang, Z. & Mustafa, H. R. (2022). On postediting of machine translation and workflow. HBET, 5793054.", 11.5, INK, False, EA, 1.5, None)])
txt(s, ML, 5.6, 12.0, 0.4, [("汇报人：＿＿", 12, GREY, False, EA, 1.4, None)])

prs.save(OUT)
print("saved:", OUT)
print("slides:", len(prs.slides._sldIdLst))
