# -*- coding: utf-8 -*-
"""
问卷调查法 正反例 · 学院模板版（第 3 稿）
- 以《高级翻译学院专用PPT模板修改版.pptx》为基底（保留左上 logo + 底部格言水印 + 思源黑体字体）
- 配色参考 ref 文件夹的学术蓝绿色系（深蓝 + 青绿 + 浅青）
- 内容：11 页（封面 + 正文 10 页），对应《一份问卷，两种命运》逐字稿
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import deck_lib as D
from deck_lib import Presentation, Inches, Pt, RGBColor, PP_ALIGN, MSO_ANCHOR, MSO_SHAPE, qn

# ---- 字体（Windows 测量字体 + 模板字体名）----
D.F_REG = r"C:\Windows\Fonts\msyh.ttc"
D.F_BOLD = r"C:\Windows\Fonts\msyhbd.ttc"
D.IDX = 0
D.EA = "思源黑体 CN Light"   # 模板中文字体
D.LAT = "Arial"
D.MONO = "Consolas"
EA = D.EA

# ---- 配色（ref 蓝绿学术风）----
NAVY   = "2D2D89"
INK    = "1F2438"
SUB    = "4A5266"
GREY   = "7A8293"
FAINT  = "A5ACBA"
TEAL   = "0E7C86"
TEAL_LT= "E7F1F1"
GOOD   = "1E7A5C"
GOOD_LT= "E9F2EE"
WARN   = "B4451F"
WARN_LT= "F7EBE5"
GOLD   = "C08A3E"
LINE   = "DDE3E9"
PAPER  = "FFFFFF"

TPL = r"d:\10_Workspace\translation-research-202610\翻译研究方法\template\高级翻译学院专用PPT模板修改版.pptx"
OUT = r"d:\10_Workspace\translation-research-202610\翻译研究方法\第1组\问卷调查法正反例_学院模板版.pptx"

W, H = 13.3333, 7.5
ML = 0.72          # 左边距（略大于 logo 区，避免遮挡）
CR = W - ML
TOTAL = 11

prs = Presentation(TPL)
blank_layout = prs.slides[0].slide_layout   # 空白版式（含 logo + 底部格言）

# 删除模板原有 6 页（同时解除 part 关系，避免重名）
sldIdLst = prs.slides._sldIdLst
for sid in list(sldIdLst):
    rId = sid.get(qn('r:id'))
    prs.part.drop_rel(rId)
    sldIdLst.remove(sid)


def S():
    return prs.slides.add_slide(blank_layout)


def rect(slide, x, y, w, h, fill=None, ln=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE, adj=None):
    return D.rect(slide, x, y, w, h, fill, ln, lw, shape, adj)


def txt(slide, x, y, w, h, items, align="l", anchor="top", valign="top"):
    """items: [(text,size,color,bold,font,lh,spc), ...] 或 str"""
    tb, tf = D.add_tb(slide, x, y, w, h, anchor=anchor)
    if isinstance(items, str):
        items = [(items, 12.5, INK, False, EA, 1.34, None)]
    for n, it in enumerate(items):
        t, sz, col, bd, fnt, lh, spc = (list(it) + [None] * 7)[:7]
        D.para(tf, t, size=sz, color=col, bold=bd, font=fnt or EA, lh=lh or 1.34,
               align=align, first=(n == 0), spc=spc)
    return tb, tf


def head(slide, eyebrow, title, no, accent=NAVY, sub=None):
    """统一页头：左上角眉题 + 标题（避开 logo 区，logo 在 y<1.0 左上角）"""
    rect(slide, ML, 1.28, 0.10, 0.52, accent)          # 标题左侧竖条
    txt(slide, ML + 0.26, 1.18, 11.0, 0.26, [(eyebrow, 10, GREY, True, EA, 1.2, 1.1)])
    txt(slide, ML + 0.24, 1.44, 11.6, 0.62, [(title, 25, NAVY, True, EA, 1.08, None)])
    if sub:
        txt(slide, ML + 0.26, 2.06, 12.0, 0.28, [(sub, 11, GREY, False, EA, 1.3, None)])
    foot(slide, no)


def foot(slide, no):
    txt(slide, ML, 7.06, 9.0, 0.24,
        [("《翻译研究方法》第 1 组 · 问卷调查法正反例 · 2026.10.12", 8.5, FAINT, False, EA, 1.2, 0.5)])
    txt(slide, 9.0, 7.06, CR - 9.0, 0.24,
        [("%02d / %02d" % (no, TOTAL - 1), 8.5, FAINT, False, D.MONO, 1.2, 0.5)], align="r")


def notes(slide, t):
    slide.notes_slide.notes_text_frame.text = t


def card(slide, x, y, w, h, fill=TEAL_LT, bar=None, ln=None, lw=0.75):
    """浅色卡片 + 可选左侧竖条"""
    D.card(slide, x, y, w, h, fill, ln, lw)
    if bar:
        rect(slide, x, y, 0.055, h, bar)


# 逐字稿（每页一段，同时写入备注与 md）
SCRIPT = [
 ("封面", "大家好，我是＿＿，接下来十分钟由我汇报我们组“方法 A 正反例”的部分：问卷调查法。正例是一篇口译研究，中文核心期刊；反例是一篇笔译研究，国际期刊，两篇都在翻译研究领域，类型一中一英。这一页只有一句话，也是今天全部的判断标准：研究问题 → 构念与题项 → 抽样与施测 → 统计与结论，这条链闭住了没有。"),
 ("第01页", "四步：研究问题问什么；构念与题项，怎么把问题变成可答的题；抽样与施测，发给谁、收回来什么；统计与结论，证据能撑到哪一步。两个案例按同一把尺看。闭住，就是一份好问卷；断在任何一环，结论就悬空。我们下面看到的两处断点，恰好一环在第 2 步、一环在第 4 步。"),
 ("第02页", "交替传译里有个东西叫笔记熟练度，大家都想测，可它看不见摸不着：你不能靠看一份笔记成品打分，也很难让别人替他评。作者就干了两件事——给它一个可操作的定义，再编一份可以被复核的自评问卷，也就是量表。所以这一页的重点不是“发现了什么规律”，而是：它交付了一件别人还能接着用的测量工具。周金华、董燕萍 2019 年，发在《外语教学与研究》，文末附了全部 21 个题项，构念是两阶段乘四个维度。"),
 ("第03页", "七步。请大家只看右边那一列：每一步，论文里都留下一行可以核查的证据，这一列就是你写“研究方法”那一节时要写的句子。前三步：界定构念，把“熟练”拆成记录和理解两个阶段、四个维度；生成题库七十四题，题干几乎全取自学习日志和访谈里译员的原话；专家评审，五名学员译员加三名教师译员，三轮讨论历时半年，收到二十八题、李克特六点。第四步试测，五十四名学生，用极端组比较和同质性检验删掉七题。改名“口译笔记自我描述量表”，为的是压低社会赞许性；正式施测一百四十四名近一学年有笔记交传训练的本科生；α 等于 .901，内容效度指数全为 1，四个因子解释 60.49%，再和三项效标相关。"),
 ("第04页", "三点可以直接搬进我们的论文。第一，构念先于题项：先写操作定义，再写“构念—维度—题项”对齐表，最后才落笔写题，不要想到什么问什么。第二，删题是设计，不是失败：七十四到二十八到二十一，每一轮都留下依据——删了哪些、按哪个指标、阈值多少，同时把信度和内容效度量化报告出来。第三，结论只走一步：量表和练习量、动机、绩效对照，所以它敢说“可测、能区分训练阶段”，但绝不说“笔记好所以口译好”。"),
 ("第05页", "反例换成笔译方向，题目其实很像我们会写的：本科翻译专业要不要引入“机器翻译加译后编辑”，也就是 MTPE 的流程？作者自拟一份问卷：六条李克特陈述，感知有用性和感知易用性各三题、六点量表，再加四道开放题，理论框架借的是技术接受模型 TAM。被试是西北一所高校的一百二十七名本科生，一次横断面施测，统计上只做描述和相关。数据都报了：样本量、两个分量表的 α 是 .714 和 .835，六题均值四点零二到四点五五，都高于中点。所以这篇的问题不是假，而是它撑不起被要求撑的那个结论。"),
 ("第06页", "六步都做了，每一步只写了半句。第一、二步：目的是论证流程可行，可问卷只回答态度；借了 TAM 的框架，没借它的结构，六个自拟题到底测不测得到那两个构念，没检验。第三、四步：一所高校一百二十七人、便利抽样，发放和回收都没报，于是应答率、自选择偏差都无从判断；线上一次填答，没有照应题，也没有注意力检查。第五、六步：拿单个题项和课程成绩做相关，均值高于中点就叫“积极”，等于把题项当构念分数用；最后从横断面自陈，一步走到“应纳入课程”。"),
 ("第07页", "如果要救，只需要三处最小改动。第一，题项来自已验证的量表：自拟三条不等于“感知有用性”；沿用就注明出处，自编就至少补一次 EFA，或者报告题总相关。第二，别用单题代表构念：先把题目按维度合成均分、报 α，再做相关，措辞只写“关联”，不写“影响”。第三，结论收缩到样本：一百二十七人只能说明“该校学生”，同时报出发放、回收、有效三份数字和判定规则。三条都不需要重做研究，当天就能补。"),
 ("第08页", "把两篇放在同一张桌子上，它们问的其实是同一件事：这份问卷，凭什么得出这个结论？题项哪来的：一边是日志、访谈加三轮评审，七十四到二十八到二十一；一边自拟六条，无评审、无前测。测的是什么：一边四个维度的构念，并用因子分析复核过；一边两个构念各三题，结构未检验。谁在填：一边一百四十四名同一训练阶段的学生、条件写明；一边两个年级一百二十七人、回收口径没写。信度 .901 对 .714 和 .835；效度一边有内容效度、因子分析和三项效标，一边只有单题和成绩的相关；结论一边停在“可测”，一边写到“可行”。"),
 ("第09页", "这是我们组整理的十条，交稿前逐条打勾，勾不上的那一行就是最可能被老师追问的地方。研究问题写成一句可回答的话；构念、维度、题项要有一张能对齐的表；题项来源写清——沿用注出处、改编列改动、自编说过程；总体、抽样框、纳入排除各一句话；发放、回收、有效三份数字，先定无效判定规则；前测或专家评审留痕；信度报 α 或 ω，多维就报 EFA 或 CFA；指导语交代匿名、用途和自愿退出；题型与统计对应；最后，横断面自陈只说关联、不说因果。"),
 ("第10页", "最后四道题留给大家，大约四分钟：第一题，预算只剩三十份问卷，四道检验保哪两道、舍哪一道；第二题，“均值高于中点”要补哪个统计量才站得住；第三题，把结论推广到 MTI 群体，最少补哪三条信息；第四题留给我们自己，对照刚才那几处断点，你最担心踩到哪一条。资料清单在这页右边，两篇原文都能直接点开核对数字。我们的汇报到这里，谢谢。"),
]


# ============================================================ 封面
s = S()
# 顶部居中分隔（标题上方留白，logo 在左上）
txt(s, ML, 2.10, 11.9, 0.34,
    [("《翻译研究方法》第 1 组 · 方法 A 案例部分", 12, GREY, False, EA, 1.3, 0.5)])
txt(s, ML, 2.48, 12.0, 1.0, [("一份问卷，两种命运", 42, NAVY, True, EA, 1.12, None)])
txt(s, ML, 3.62, 11.9, 0.36,
    [("A Questionnaire, Two Fates · Interpreting × Translation", 14.5, GREY, False, D.LAT, 1.3, None)])
rect(s, ML, 4.28, 4.6, 0.03, TEAL)
txt(s, ML, 4.50, 12.0, 1.0, [
    ("正例（口译研究）｜ 周金华、董燕萍（2019）. 口译笔记熟练度量表的开发. 外语教学与研究 51(6): 925-937.", 12.5, INK, False, EA, 1.55, None),
    ("反例（笔译研究）｜ Yang & Mustafa (2022). On postediting of machine translation and workflow. HBET, 5793054.", 12.5, INK, False, EA, 1.55, None)])
txt(s, ML, 5.62, 12.0, 0.34,
    [("全程只用一把尺：研究问题 → 构念与题项 → 抽样与施测 → 统计与结论", 12.5, NAVY, True, EA, 1.4, 0.4)])
txt(s, ML, 6.12, 12.0, 0.8, [
    ("小组分工（4 人）：于萍［组长］框架与主线 · 陈冠臻 正例核对 · 熊芮 反例核对 · 方燕 对照与清单 · 汇报人：＿＿", 10, FAINT, False, EA, 1.45, None)])
notes(s, SCRIPT[0][1])

# ============================================================ 01 先看链条
s = S()
head(s, "THE CHAIN · 四步", "先看链条", 1, NAVY, "闭住就是好问卷；断在任何一环，结论就悬空。")
chain = [("研究问题", "要回答什么", "Research question"),
         ("构念与题项", "怎么变成可答的题", "Construct & items"),
         ("抽样与施测", "向谁发、怎么收", "Sampling & admin."),
         ("统计与结论", "证据撑到哪一步", "Analysis & claim")]
bw, g = 2.82, 0.24
for i, (a, b, c) in enumerate(chain):
    x = ML + i * (bw + g)
    card(s, x, 2.78, bw, 1.78, TEAL_LT, bar=TEAL)
    txt(s, x + 0.26, 3.04, bw - 0.5, 1.4, [
        ("%02d" % (i + 1), 11, TEAL, False, D.MONO, 1.3, 0.5),
        (a, 17, INK, True, EA, 1.3, 0.3),
        (b, 11.5, SUB, False, EA, 1.4, None),
        (c, 10, FAINT, False, D.LAT, 1.4, None)])
txt(s, ML, 4.86, 11.6, 0.34,
    [("断点最常出现在第 2 与第 4 步：题项测的不是那个构念；结论比数据多走了几步。", 11.5, GREY, False, EA, 1.4, None)])
notes(s, SCRIPT[1][1])

# ============================================================ 02 正例概况
s = S()
head(s, "POSITIVE CASE · 口译研究", "把“熟练度”变成可测的构念", 2, TEAL,
     "周金华、董燕萍（2019）. 口译笔记熟练度量表的开发. 外语教学与研究 51(6): 925-937.（CSSCI 核心期刊）")
card(s, ML, 2.78, 11.9, 1.5, TEAL_LT, bar=TEAL)
txt(s, ML + 0.36, 3.04, 11.3, 1.05, [
    ("笔记熟练度，看不见、摸不着 —— 能测吗？", 13.5, INK, True, EA, 1.5, 0.3),
    ("给一个可操作的定义，再编一份可被复核的自评问卷（量表）。目标不是“发现规律”，而是交付一件能接着用的工具。", 12, SUB, False, EA, 1.5, None)])
txt(s, ML, 4.62, 11.9, 0.4,
    [("构念：两阶段（记录 / 理解）× 四维度（听记协调 · 系统性 · 时效性 · 使用）", 12.5, INK, False, EA, 1.4, None)])
txt(s, ML, 5.16, 11.9, 0.4,
    [("附录公开 21 个题项，工具可复用、可被复核。", 11, GREY, False, EA, 1.4, None)])
notes(s, SCRIPT[2][1])

# ============================================================ 03 正例流程
s = S()
head(s, "POSITIVE CASE · PROCEDURE", "七步，每步留下一件证据", 3, TEAL,
     "右栏 ＝ 你写“研究方法”一节时要写的句子。")
rows = [("① 界定构念", "两阶段 × 四维度", "构念定义＋维度表"),
        ("② 生成题库", "74 题，来自日志与访谈", "题项来源说明"),
        ("③ 专家评审", "3 轮 → 28 题（6 点）", "评审人数与依据"),
        ("④ 试测删题", "54 人 → 删 7 题", "删题指标与阈值"),
        ("⑤ 正式施测", "144 人", "总体口径与份数"),
        ("⑥ 信度效度", "α = .901；CVI = 1", "信效度指标表"),
        ("⑦ 三项效标", ".473 · .539 · .556", "效标＋组间比较")]
y = 2.74
rect(s, ML, y, 11.9, 0.36, TEAL)
txt(s, ML + 0.2, y, 1.7, 0.36, [("步骤", 11.5, "FFFFFF", True, EA, 1.2, None)], anchor="mid")
txt(s, 2.6, y, 5.9, 0.36, [("做了什么", 11.5, "FFFFFF", True, EA, 1.2, None)], anchor="mid")
txt(s, 8.6, y, 3.9, 0.36, [("留下什么证据", 11.5, "FFFFFF", True, EA, 1.2, None)], anchor="mid")
y += 0.36
for i, (k, a, b) in enumerate(rows):
    hh = 0.44
    fill = TEAL_LT if i % 2 == 0 else "FFFFFF"
    rect(s, ML, y, 11.9, hh, fill)
    txt(s, ML + 0.2, y, 1.7, hh, [(k, 11.5, INK, True, EA, 1.2, None)], anchor="mid")
    txt(s, 2.6, y, 5.9, hh, [(a, 11.5, SUB, False, EA, 1.2, None)], anchor="mid")
    txt(s, 8.6, y, 3.9, hh, [(b, 11.5, INK, False, EA, 1.2, None)], anchor="mid")
    y += hh
txt(s, ML, y + 0.12, 11.9, 0.3,
    [("结果：21 题 · 4 因子解释 60.49% ｜ 初级组与高级组差异显著", 11.5, GREY, False, EA, 1.4, None)])
notes(s, SCRIPT[3][1])

# ============================================================ 04 正例三点
s = S()
head(s, "POSITIVE CASE · TAKEAWAYS", "三点值得照搬", 4, GOOD,
     "括号里是你在“研究方法”一节里的写法。")
takes = [("构念先于题项", "先有操作定义与对齐表，再写题项。", "Construct first, items second."),
         ("删题是设计，不是失败", "74 → 28 → 21，每轮都留下依据。", "Item analysis, reported."),
         ("结论只走一步", "只说“可测、能区分阶段”，不说“笔记好所以口译好”。", "Claim no more than data.")]
y = 2.72
for i, (a, b, c) in enumerate(takes):
    card(s, ML, y, 11.9, 1.22, GOOD_LT, bar=GOOD)
    txt(s, ML + 0.36, y + 0.22, 11.2, 0.95, [
        ("%d. %s" % (i + 1, a), 15, INK, True, EA, 1.34, 0.3),
        (b, 11.8, SUB, False, EA, 1.5, None),
        (c, 10, FAINT, False, D.LAT, 1.4, None)])
    y += 1.38
txt(s, ML, y + 0.06, 11.9, 0.26,
    [("另：21 题随附录公开 —— 工具可复用，研究才能被复核、被引用。", 10.5, FAINT, False, EA, 1.4, None)])
notes(s, SCRIPT[4][1])

# ============================================================ 05 反例概况
s = S()
head(s, "NEGATIVE CASE · 笔译研究", "六道题，去撑一个新流程", 5, WARN,
     "Yang & Mustafa (2022). On postediting of MT and workflow. HBET, 5793054.（Wiley 开放获取）")
card(s, ML, 2.78, 11.9, 1.5, WARN_LT, bar=WARN)
txt(s, ML + 0.36, 3.04, 11.3, 1.05, [
    ("问题：本科翻译专业要不要引入“机器翻译 ＋ 译后编辑（MTPE）”的流程？", 13.5, INK, True, EA, 1.5, 0.3),
    ("工具：6 条李克特陈述（有用性 3 · 易用性 3，6 点）＋ 4 道开放题；借技术接受模型（TAM）；与两门课成绩做相关。", 12, SUB, False, EA, 1.5, None)])
txt(s, ML, 4.62, 11.9, 0.4,
    [("它也报了数字：N = 127 ｜ α = .714 / .835 ｜ 均值 4.02–4.55（都高于中点）", 12.5, INK, False, EA, 1.4, None)])
txt(s, ML, 5.16, 11.9, 0.4,
    [("先说清楚：这不是造假，数据是真的、统计也报了 —— 只是撑不起被要求撑的结论。", 11, WARN, True, EA, 1.4, None)])
notes(s, SCRIPT[5][1])

# ============================================================ 06 反例断点
s = S()
head(s, "NEGATIVE CASE · WHERE IT BREAKS", "六步都做了，每步只半句", 6, WARN,
     "左：论文写了什么 ｜ 右：缺掉的那一环")
brk = [("① 目的与工具错位", "要论证“可行”，只答了态度", "可行性需评审或试点"),
       ("② 借框架不借结构", "TAM 两构念，6 条自拟题", "结构未经检验"),
       ("③ 抽样口径模糊", "一所高校，便利抽样", "未报发放 / 回收 / 有效"),
       ("④ 施测一次成型", "线上一次填答", "无照应题与偏差控制"),
       ("⑤ 分析层级错位", "单题 × 课程成绩", "把题项当构念用"),
       ("⑥ 写作代替解释", "均值高于中点 → “可行”", "横断面推不出因果")]
y = 2.74
rect(s, ML, y, 11.9, 0.36, WARN)
txt(s, ML + 0.2, y, 2.2, 0.36, [("环节", 11.5, "FFFFFF", True, EA, 1.2, None)], anchor="mid")
txt(s, 3.2, y, 4.5, 0.36, [("论文写了什么", 11.5, "FFFFFF", True, EA, 1.2, None)], anchor="mid")
txt(s, 7.8, y, 4.8, 0.36, [("缺掉的一环", 11.5, "FFFFFF", True, EA, 1.2, None)], anchor="mid")
y += 0.36
for i, (k, a, b) in enumerate(brk):
    hh = 0.44
    fill = WARN_LT if i % 2 == 0 else "FFFFFF"
    rect(s, ML, y, 11.9, hh, fill)
    txt(s, ML + 0.2, y, 2.2, hh, [(k, 11.5, INK, True, EA, 1.2, None)], anchor="mid")
    txt(s, 3.2, y, 4.5, hh, [(a, 11.5, SUB, False, EA, 1.2, None)], anchor="mid")
    txt(s, 7.8, y, 4.8, hh, [(b, 11.5, INK, False, EA, 1.2, None)], anchor="mid")
    y += hh
txt(s, ML, y + 0.12, 11.9, 0.3,
    [("这不是造假：数据是真的，统计也报了 —— 它只是撑不起被要求撑的那个结论。", 12.5, WARN, True, EA, 1.5, None)])
notes(s, SCRIPT[6][1])

# ============================================================ 07 反例三处最小改法
s = S()
head(s, "NEGATIVE CASE · LESSONS", "三处最容易踩", 7, WARN,
     "每处一条“最小改法”，不需要重做研究，当天就能补。")
bad = [("题项要来自已验证的量表", "自拟 3 条不等于“感知有用性”。",
        "改：沿用已验证条目并标出处；补 EFA 或报题总相关。"),
       ("别用单题代表构念", "Q6 × 成绩 r = .746 → “能力强”？",
        "改：先合成维度均分并报 α，再做相关；只写“关联”。"),
       ("结论要收缩到样本", "127 人 → “教学可行”？",
        "改：称“该校学生”；报发放 / 回收 / 有效三份数字。")]
y = 2.72
for a, b, c in bad:
    card(s, ML, y, 11.9, 1.22, WARN_LT, bar=WARN)
    txt(s, ML + 0.36, y + 0.22, 11.2, 0.95, [
        (a, 15, INK, True, EA, 1.34, 0.3),
        (b, 11.8, SUB, False, EA, 1.5, None),
        (c, 11.8, INK, False, EA, 1.5, None)])
    y += 1.38
txt(s, ML, y + 0.06, 11.9, 0.26,
    [("底线：剔卷标准与编码规则要预先写清（此处 18 份被剔除，依据未披露）。", 10.5, FAINT, False, EA, 1.4, None)])
notes(s, SCRIPT[7][1])

# ============================================================ 08 两案并排
s = S()
head(s, "SIDE BY SIDE", "两案并排", 8, NAVY, "两篇问的其实是同一件事：这份问卷，凭什么得出这个结论？")
cmp = [("题项哪来的", "日志＋访谈＋评审 74→28→21", "自拟 6 条，无评审前测"),
       ("测的是什么", "4 维度，EFA 复核吻合", "2 构念各 3 题，未检验"),
       ("谁在填", "144 人，条件写明", "127 人，回收口径未报"),
       ("信度", "α = .901", ".714 / .835"),
       ("效度", "CVI ＋ EFA ＋ 三项效标", "单题 × 课程成绩"),
       ("结论走到哪", "可测、能区分训练阶段", "流程可行、纳入课程")]
y = 2.78
rect(s, ML, y, 2.3, 0.36, NAVY)
txt(s, ML + 0.2, y, 2.1, 0.36, [("维度", 11.5, "FFFFFF", True, EA, 1.2, None)], anchor="mid")
rect(s, ML + 2.3, y, 4.7, 0.36, TEAL)
txt(s, ML + 2.5, y, 4.4, 0.36, [("口译 · 正例", 11.5, "FFFFFF", True, EA, 1.2, None)], anchor="mid")
rect(s, ML + 7.0, y, 4.9, 0.36, WARN)
txt(s, ML + 7.2, y, 4.6, 0.36, [("笔译 MTPE · 反例", 11.5, "FFFFFF", True, EA, 1.2, None)], anchor="mid")
y += 0.36
for i, (k, a, b) in enumerate(cmp):
    hh = 0.46
    fill = "FFFFFF" if i % 2 == 0 else TEAL_LT
    rect(s, ML, y, 2.3, hh, fill)
    txt(s, ML + 0.2, y, 2.1, hh, [(k, 11.5, INK, True, EA, 1.2, None)], anchor="mid")
    rect(s, ML + 2.3, y, 4.7, hh, "F3FAF9")
    txt(s, ML + 2.5, y, 4.4, hh, [(a, 11.3, INK, False, EA, 1.2, None)], anchor="mid")
    rect(s, ML + 7.0, y, 4.9, hh, "FBF5F2")
    txt(s, ML + 7.2, y, 4.6, hh, [(b, 11.3, SUB, False, EA, 1.2, None)], anchor="mid")
    y += hh
txt(s, ML, y + 0.14, 11.9, 0.34,
    [("倒过来读，就是你的自检表", 13, NAVY, True, EA, 1.4, None)])
txt(s, ML, y + 0.52, 11.9, 0.34,
    [("题项哪来的？测的还是它吗？谁填的、留下多少？信度在哪一行？结论多走了几步？", 11, GREY, False, EA, 1.4, None)])
notes(s, SCRIPT[8][1])

# ============================================================ 09 十行自检
s = S()
head(s, "TEN CHECKS", "交稿前十行清单", 9, NAVY, "勾不上的那一行，就是你论文里最可能被追问的地方。")
ck = ["研究问题：一句可回答的话（描述 / 比较 / 关联）",
      "构念 → 维度 → 题项：一张对齐表",
      "题项来源：沿用注出处｜改编列改动｜自编说过程",
      "总体、抽样框、纳入与排除：各一句话",
      "发放 / 回收 / 有效：三份数字＋回收率",
      "无效判定规则：先定后用",
      "前测与专家评审：留痕",
      "信度 α / ω；多维则 EFA / CFA",
      "指导语：匿名、用途、自愿退出；敏感题后置",
      "结论分寸：横断面只说关联，不说因果"]
col_gap = 6.05
for i, c in enumerate(ck):
    col = i // 5
    row = i % 5
    x = ML + col * col_gap
    y = 2.72 + row * 0.62
    rect(s, x, y, 0.34, 0.34, NAVY if col == 0 else TEAL, None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.2)
    txt(s, x, y, 0.34, 0.34, [("%02d" % (i + 1), 9.5, "FFFFFF", True, D.MONO, 1.2, None)], align="c", anchor="mid")
    txt(s, x + 0.5, y + 0.02, col_gap - 0.6, 0.34, [(c, 11.5, INK, False, EA, 1.25, None)])
txt(s, ML, 5.98, 11.9, 0.4,
    [("这十条在正例那篇里都有对应段落 —— 抄它的结构，别抄它的内容。", 11, GREY, False, EA, 1.4, None)])
notes(s, SCRIPT[9][1])

# ============================================================ 10 思考题 + 资料
s = S()
head(s, "DISCUSSION & SOURCES", "四道思考题 ＋ 资料清单", 10, NAVY, "讨论约 4 分钟；每题本组先自答一句，再点人。")
qs = [("Q1 取舍", "30 份问卷的预算，四道检验保哪两道？"),
      ("Q2 补一个统计量", "“均值高于中点”＝“总体积极”？你补哪个、为什么？"),
      ("Q3 追问样本", "要推广到 MTI，最少补哪三条信息？"),
      ("Q4 反身", "对照六病，你最怕踩哪一条？最小改法是什么？")]
y = 2.72
for i, (k, q) in enumerate(qs):
    num, label = k.split(" ", 1)
    txt(s, ML, y, 0.7, 0.32, [(num, 12, TEAL, True, D.MONO, 1.3, 0.3)])
    txt(s, ML + 0.85, y, 1.7, 0.32, [(label, 12, INK, True, EA, 1.3, 0.2)])
    txt(s, ML + 2.8, y, 4.4, 0.32, [(q, 11.5, SUB, False, EA, 1.34, None)])
    rect(s, ML, y + 0.38, 7.3, 0.012, LINE)
    y += 0.58
txt(s, ML, 5.10, 7.3, 0.5,
    [("讨论 4 分钟：本组每题先自答一句，再点人。", 10.5, FAINT, False, EA, 1.5, None)])
# 右侧资料
txt(s, 8.5, 2.66, 4.2, 0.3, [("论文与学习资料", 11, INK, True, EA, 1.3, 0.3)])
refs = [("正例全文＋量表附录", "bcdlab.gdufs.edu.cn/info/1025/2025.htm"),
        ("反例原文（开放获取）", "onlinelibrary.wiley.com/doi/10.1155/2022/5793054"),
        ("同类笔译态度调查", "doi.org/10.18298/ijlet.3242"),
        ("口译问卷范例", "张威（2013）《中国翻译》(2): 17-25"),
        ("方法读物", "吴明隆《问卷统计分析实务》"),
        ("检索式", "CNKI: 翻译 AND（问卷 OR 量表）AND 信效度")]
y = 3.02
for a, b in refs:
    txt(s, 8.5, y, 4.2, 0.28, [(a, 10.2, INK, False, EA, 1.32, None)])
    txt(s, 8.5, y + 0.25, 4.2, 0.30, [(b, 8.6, FAINT, False, D.MONO, 1.25, None)])
    y += 0.66
txt(s, 8.5, y + 0.06, 4.2, 0.4,
    [("两条链接都可直接点开核对数字；检索式可复制到知网 / Web of Science。", 9.5, FAINT, False, EA, 1.4, None)])
notes(s, SCRIPT[10][1])

prs.save(OUT)
print("saved:", OUT)
print("slides:", len(prs.slides._sldIdLst))

# 生成逐字稿 md
md = ["# 《一份问卷，两种命运》· 10 分钟逐字稿（学院模板版）",
      "",
      "对应课件：`问卷调查法正反例_学院模板版.pptx`（11 页：封面 ＋ 正文 10 页）",
      "语速：口播约 2300 字，272 字/分钟 ≈ 8 分 27 秒，红线 10 分钟。",
      "行首（问题）（设计）等为提示词，不念；`〔…〕`为动作提示，不念。",
      ""]
for tag, t in SCRIPT:
    md.append("## %s" % tag)
    md.append("")
    md.append(t)
    md.append("")
MD_OUT = r"d:\10_Workspace\translation-research-202610\翻译研究方法\第1组\逐字稿_学院模板版.md"
with open(MD_OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(md))
print("saved:", MD_OUT)
