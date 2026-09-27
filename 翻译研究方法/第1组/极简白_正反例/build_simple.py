# -*- coding: utf-8 -*-
"""极简纯白版：问卷调查法 正反例分析（口译 1 例 + 笔译 1 例），10 分钟汇报。
复用 ../deck_lib.py 的排版/测宽能力；构建后删除所有装饰色块 → 版面只剩文字与留白。
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from pptx import Presentation
from deck_lib import (new_prs, add_slide, para, add_tb, W, H, CR, INK,
                      wrap, tw, fit_wrap, EA, LAT, MONO)

OUT = "问卷调查法正反例分析_极简白_10min.pptx"
INK_ = "0F1523"
SUB = "555F70"
FAINT = "9AA1AE"
F = "B4451F"          # 唯一强调色，只用于少量短标签

prs = new_prs()
TOTAL = 11


def S():
    return add_slide(prs)


def page(s, no, tag=None, right=None):
    if tag:
        tb, tf = add_tb(s, 0.62, 0.40, 8.6, 0.24)
        para(tf, tag, size=9.5, color=FAINT, bold=True, first=True, spc=1.1, font=LAT)
    tb, tf = add_tb(s, CR - 2.4, H - 0.50, 2.4, 0.22)
    para(tf, right or ("第 %02d / %02d 页" % (no, TOTAL)), size=8.5, color=FAINT,
         align="r", first=True, font=MONO)


def title(s, text, sub=None, y=0.66, size=26, w=None):
    w = w or (CR - 0.62)
    tb, tf = add_tb(s, 0.62, y, w, 0.66)
    para(tf, text, size=size, color=INK_, bold=True, first=True, lh=1.16)
    if sub:
        lines, sz, _ = fit_wrap(sub, w, 0.44, 11.0, lh=1.30)
        tb, tf = add_tb(s, 0.62, y + 0.66, w, 0.46)
        para(tf, "\n".join(lines), size=sz, color=SUB, first=True, lh=1.30)


def stack3(s, rows, x, w, y, gap=0.24, ksize=12.0, bsize=10.8, lw=None,
           kw=1.9, right=None, close=None, csize=10.5, ccolor=None, rsize=None, bottom=6.92):
    """三栏行（关键词｜左说明｜右说明），逐行按实际行数堆叠。right=右栏左边界。"""
    lw = lw or (11.7 - w)
    rw = (x + w) - right - 0.20
    avail = bottom - y - (0.52 if close else 0)
    LH3 = 1.30
    for _ in range(90):
        hs = [max(_lines(a, lw, bsize, LH3)[1], _lines(b, rw, rsize or bsize, LH3)[1]) for k, a, b in rows]
        need = sum(hs) + gap * len(rows) + (0.0 if not close else 0.52)
        if need <= avail or (bsize <= 9.4 and gap <= 0.08):
            break
        if bsize > 9.4:
            bsize = round(bsize - 0.15, 2)
        else:
            gap = round(gap - 0.015, 3)
    yy = y
    for k, a, b in rows:
        ka, ha = _lines(a, lw, bsize, LH3)
        kb, hb = _lines(b, rw, rsize or bsize, LH3)
        hh = max(ha, hb)
        tb, tf = add_tb(s, x, yy, kw, 0.4)
        para(tf, k, size=ksize, color=ccolor or INK_, bold=True, first=True, lh=1.18)
        tb, tf = add_tb(s, x + kw + 0.22, yy, lw, hh + 0.04)
        para(tf, ka, size=bsize, color=SUB, first=True, lh=LH3)
        tb, tf = add_tb(s, right, yy, rw, hh + 0.04)
        para(tf, kb, size=rsize or bsize, color=INK_, first=True, lh=LH3)
        yy += hh + gap
    yy -= gap
    if close:
        cl, cs, cu = _lines(close, 12.1 - X0, csize, LH3, ret=True)
        tb, tf = add_tb(s, X0, yy + 0.20, 12.1, cu + 0.04)
        para(tf, cl, size=cs, color=ccolor or FAINT, first=True, lh=LH3)
        yy = max(yy, y) + cu + 0.06
    return yy


def _lines(txt, w, size, lh, ret=False):
    ls, sz, used = fit_wrap(txt, w, 4.0, size, lh=lh)
    return (("\n".join(ls), sz, used) if ret else ("\n".join(ls), used))


def stack2(s, rows, x, w, y, gap=0.10, ksize=12.4, bsize=11.0, kcol=None, bcol=None, bottom=6.92):
    """小标题 + 正文（同栏），逐条按实际行数堆叠。"""
    kcol = kcol or w
    bcol = bcol or w
    LH2 = 1.34
    for _ in range(90):
        us = [fit_wrap(b, bcol, 4.0, bsize, lh=LH2)[2] for k, b in rows]
        need = sum(us) + (ksize * 1.2 / 72.0 + 0.06) * len([k for k, b in rows if k]) + gap * len(rows)
        if need <= bottom - y or bsize <= 9.0:
            break
        bsize = round(bsize - 0.15, 2)
    yy = y
    for k, body in rows:
        if k:
            tb, tf = add_tb(s, x, yy, kcol, 0.32)
            para(tf, k, size=ksize, color=F, bold=True, first=True, lh=1.2)
            yy += ksize * 1.2 / 72.0 + 0.05
        bl, bs, bu = _lines(body, bcol, bsize, LH2, ret=True)
        tb, tf = add_tb(s, x, yy, bcol, bu + 0.04)
        para(tf, bl, size=bs, color=INK_, first=True, lh=LH2)
        yy += bu + gap
    return yy - gap


def block(s, items, bounds, size=13.2, gap=0.34, lh=1.34, dsize=11.2,
          close=None, csize=10.8, ccolor=None, cgap=0.26):
    """[（小标题, 正文）…] 按实际需要的高度自上而下堆叠；close=页脚收束句（自动紧跟正文，不留大片空白，也不会压到正文）。"""
    x, y, w, h = bounds
    geo = []
    yy = y
    for it in items:
        ttl, body = it if isinstance(it, (list, tuple)) else (None, it)
        hh = size * 1.2 / 72.0 + 0.05 if ttl else 0.0
        lines, sz, used = fit_wrap(body, w, 3.0, dsize, lh=lh)
        geo.append((ttl, lines, sz, used, hh))
        yy += hh + used + gap
    ch = 0.0
    if close:
        clines, csz, cused = fit_wrap(close, w, 1.2, csize, lh=1.28)
        ch = cused + cgap
    total = yy - gap - y + ch
    k = 1.0
    if total > h:                                  # 极端情况：压缩间距，避免越界
        k = max(0.30, 1 - (total - h) / max(0.001, gap * (len(items) - 1)))
    yy = y
    for ttl, lines, sz, used, hh in geo:
        tb, tf = add_tb(s, x, yy, w, 1.8)
        if ttl:
            para(tf, ttl, size=size, color=INK_, bold=True, first=True, lh=1.2)
            para(tf, "\n".join(lines), size=sz, color=SUB, lh=lh, before=2)
        else:
            para(tf, "\n".join(lines), size=sz, color=INK_, first=True, lh=lh)
        yy += hh + used + gap * k
    if close:
        tb, tf = add_tb(s, x, yy + cgap * 0.4, w, 0.5)
        para(tf, "\n".join(clines), size=csz, color=ccolor or INK_, bold=True, first=True, lh=1.28)
    return yy


def lines2(s, rows, x0, w0, y, rh=0.70, gapw=0.42, ksize=12.0, bsize=10.6):
    """三栏行：关键词 ｜ 左说明 ｜ 右说明"""
    w1 = w0 * 0.34
    w2 = w0 - 2.42 - w1 - gapw
    for k, a, b in rows:
        tb, tf = add_tb(s, x0, y, 2.32, 0.4)
        para(tf, k, size=ksize, color=INK_, bold=True, first=True, lh=1.18)
        tb, tf = add_tb(s, x0 + 2.42, y, w1, rh)
        ls, sz, _ = fit_wrap(a, w1 - 0.04, rh, bsize, lh=1.28)
        para(tf, "\n".join(ls), size=sz, color=SUB, first=True, lh=1.28)
        tb, tf = add_tb(s, x0 + 2.42 + w1 + gapw, y, w2, rh)
        ls, sz, _ = fit_wrap(b, w2 - 0.04, rh, bsize, lh=1.28)
        para(tf, "\n".join(ls), size=sz, color=INK_, first=True, lh=1.28)
        y += rh + 0.04
    return y


X0, WB = 0.62, CR - 0.62

# ============================================================ 00 封面
s = S()
tb, tf = add_tb(s, X0, 1.34, 11.6, 0.3)
para(tf, "《翻译研究方法》第 1 组 ｜ 方法 A 案例部分：问卷调查法的正例与反例", size=11.5,
     color=SUB, first=True, spc=0.5)
tb, tf = add_tb(s, X0, 1.90, 11.8, 1.6)
para(tf, "一份问卷，两种命运", size=40, color=INK_, bold=True, first=True, lh=1.12)
para(tf, "Questionnaires, well used and badly used · Interpreting × Translation",
     size=14.5, color=SUB, before=10, font=LAT)
tb, tf = add_tb(s, X0, 3.94, 11.7, 1.0)
para(tf, "正例（口译研究）周金华、董燕萍（2019）. 口译笔记熟练度量表的开发. 外语教学与研究 51(6): 925-937.",
     size=12.5, color=INK_, first=True, lh=1.55)
para(tf, "反例（笔译研究）Yang, Z. & Mustafa, H. R. (2022). On postediting of machine translation and workflow… "
         "Human Behavior and Emerging Technologies, 2022: 5793054.",
     size=12.5, color=INK_, lh=1.55)
tb, tf = add_tb(s, X0, 5.14, 11.6, 0.5)
para(tf, "全程只回答一个问题：这条链——研究问题 → 构念与题项 → 抽样与施测 → 统计与结论——闭了没有。",
     size=11.5, color=SUB, first=True, lh=1.5)
tb, tf = add_tb(s, X0, 5.86, 11.9, 0.8)
para(tf, "小组分工（成员 4 人）", size=10, color=FAINT, bold=True, first=True, spc=0.6)
para(tf, "于萍［组长］：整体框架与逻辑主线、正反例评析（第 04、07 页）｜ 汇报人：＿＿（本页 10 分钟，正文 10 页）",
     size=10.5, color=INK_, lh=1.5, before=3)
para(tf, "陈冠臻：正例（口译研究）文献核对与流程梳理（第 02–03 页）｜ 熊芮：反例（笔译研究）文献核对与数据摘记（第 05–06 页）",
     size=10.5, color=INK_, lh=1.5)
para(tf, "方燕：交稿自检清单、4 道思考题与资料汇总（第 09–10 页）｜ 全员参与 PPT 定稿、教室试播与 U 盘备份（.pptx ＋ PDF）",
     size=10.5, color=INK_, lh=1.5)
page(s, 0, right="封面 · 不计页码")

# ============================================================ 01 路线
s = S()
title(s, "10 分钟怎么走", "两个案例走同样四步，方便逐条对照；每页 50–70 秒（时间见讲稿）")
block(s, [("① 案例概况", "研究问题 / 方法设计 / 主要发现 —— 口译、笔译各一页（第 02、05 页）"),
          ("② 操作流程", "他们一步一步怎么做的，每步留下了什么（第 03、06 页）"),
          ("③ 评析", "正例六点可迁移（第 04 页）· 反例六病引以为戒（第 07 页）"),
          ("④ 收束", "两案并排对照 → 交稿前 10 行自检 → 4 道思考题与资料清单（第 08–10 页）")],
      (X0, 2.14, 11.6, 3.55), size=14, gap=0.40, dsize=11.9)
tb, tf = add_tb(s, X0, 5.70, 11.6, 0.6)
para(tf, "为什么挑这两篇：口译那篇把问卷法的全部环节都写在纸面上，最容易被复制；笔译那篇题少、发表快、也报了统计，"
         "是同学作业里最常见的“看起来像实证”的问卷。两篇都是翻译研究领域，一中文一英文、一份是核心期刊、一份是国际 OA 期刊。",
     size=10.8, color=SUB, first=True, lh=1.45)
page(s, 1, tag="ROADMAP")

# ============================================================ 02 正例概况
s = S()
title(s, "正例｜口译研究：把“说不清的熟练度”做成可测的问卷",
      "周金华、董燕萍（2019）. 口译笔记熟练度量表的开发. 外语教学与研究 51(6): 925-937.（CSSCI 核心期刊）")
block(s, [("研究问题 / 目标", "交替传译的“笔记熟练度”看不见摸不着：不能靠看笔记成品打分，也很难由他人评估。"
           "能否给出可操作的定义，并编一份可以被复核的自评问卷（量表）来测它？"),
          ("研究方法设计", "质性材料生成题库（学员学习日志；4 名学员译员 + 6 名教师译员的开放型访谈）→ 多轮专家评议 → 预试量表 → "
           "试测做项目分析 → 正式量表 → 正式施测 → 信度 + 内容效度 + 建构效度 + 效标关联效度四道检验。"),
          ("主要发现", "量表 21 题、4 个维度（听记协调性 / 记录系统性 / 时效性 / 笔记使用）；内部一致性 α = .901；"
           "探索性因子分析抽出 4 个公因子，与预设维度吻合，累计解释变异量 60.49%；总分与每周有笔记交传练习小时数 "
           "r = .473、与口译学习动机 r = .539、与英汉交译测试绩效 r = .556（均 p < .01）；"
           "初级组与高级组在全量表及各维度上差异显著。结论：笔记熟练度确实存在、可以被自评问卷有效测量、并且随练习发展。")],
      (X0, 2.34, 11.7, 4.05), size=13.6, gap=0.42, dsize=12.0,
      close="这一页的重点不是“发现了什么规律”，而是：它交付了一件别人还能接着用的测量工具。", csize=11.6)
page(s, 2, tag="POSITIVE CASE · OVERVIEW")

# ============================================================ 03 正例流程
s = S()
title(s, "正例｜操作流程：七步，每步都留下一件可核查的东西", "先做什么 → 交出什么证据（这一列就是你论文里要写的句子）")
rows = [("① 界定构念", "把“熟练”拆成两阶段（记录 / 理解）、四个维度，写成可观察的操作定义", "构念定义 + 维度表"),
        ("② 生成题库", "从日志与访谈中提取表述，初始题库 74 题；题干尽量用译员的语言", "题项来源说明"),
        ("③ 专家评审", "5 名学员 + 3 名教师译员评议，3 轮讨论历时半年 → 预试量表 28 题（李克特 6 点）", "评审人数、轮次、删改依据"),
        ("④ 试测删题", "54 名翻译专业大三学生试测；极端组比较 + 同质性检验 → 删 7 题，留 21 题并重新随机排序", "项目分析指标与删题阈值"),
        ("⑤ 正式施测", "144 名已完成近一学年有笔记交传训练的本科生；条件与人数一并写清", "总体口径 + 有效份数"),
        ("⑥ 信度效度", "α = .901；内容效度指数 I-CVI、S-CVI 全部为 1；EFA 四因子解释 60.49% 变异", "信效度指标表"),
        ("⑦ 以数据作答", "与练习量、动机、口译绩效三项效标相关；初级/高级组独立样本 t 检验显著", "效标证据 + 组间比较")]
stack3(s, rows, X0, 11.7, 2.22, gap=0.20, ksize=12.2, bsize=11.2, kw=1.85, lw=6.50, right=9.42, bottom=6.44,
       close="七步的共同点：每问一个问题，论文里就有一行对应的证据——题项数、被试数、删题理由、α、CVI、方差解释率、效标相关。",
       csize=10.5)
page(s, 3, tag="POSITIVE CASE · PROCEDURE")

# ============================================================ 04 正例优点
s = S()
title(s, "正例｜六点值得照搬", "括号里是：你在“研究方法”一节里怎么把它写出来")
block(s, [("① 构念先于题项", "先有操作定义，再写题项，不“想到什么问什么”。（写一段构念界定，配一张构念—维度—题项对齐表）"),
          ("② 题库大于量表", "74 → 28 → 21，删题是设计的一环而不是失败。（附题项筛选表：删了哪些、依据哪个指标）"),
          ("③ 量表形式经过选择", "李克特 6 点、去掉中点，逼出倾向；避免“中立即逃避”。（说明为何选 6 点而非 5 点）"),
          ("④ 内容效度量化", "不是“请专家看了一下”，而是 I-CVI / S-CVI 计数并报告评审者一致性。（列专家人数与判定标准）"),
          ("⑤ 效度有外部之锚", "把问卷分数与练习量、动机、口译绩效对照，问卷才不是自说自话。（写清效标从何而来、如何测）"),
          ("⑥ 结论只走一步", "只说“可测、能区分训练阶段”，不说“笔记好所以口译好”。（局限段交代样本与横断面性质）")],
      (X0, 2.26, 11.7, 4.10), size=13.4, gap=0.30, dsize=12.0,
      close="补一句：量表的 21 个题项随论文附录公开——工具可复用，研究才能被复核、被引用。", csize=11.4)
page(s, 4, tag="POSITIVE CASE · WHY IT WORKS")

# ============================================================ 05 反例概况
s = S()
title(s, "反例｜笔译研究：用 6 道题的问卷，去支撑一个新教学流程",
      "Yang, Z. & Mustafa, H. R. (2022). On postediting of machine translation and workflow for undergraduate "
      "translation program in China. Human Behavior and Emerging Technologies, 2022: 5793054.（Wiley，开放获取）")
block(s, [("研究问题 / 目标", "本科翻译专业是否应引入“机器翻译 + 译后编辑（MTPE）”流程？作者列三问：学生对 MTPE 的正面与负面反应是什么？"
           "这些反应与两门课（计算机辅助翻译、翻译实践）成绩有无相关？阻碍使用的主要因素有哪些？"),
          ("研究方法设计", "混合方法：问卷含 6 条李克特陈述（感知有用性 3 题、感知易用性 3 题，6 点量表）+ 4 道开放题；"
           "被试为我国西北一所高校翻译专业本科生 127 人（大三 39、大四 88）；定量用 SPSS 做描述统计与相关，"
           "开放题用 NVivo 编码为主题；理论框架取自技术接受模型（TAM）。"),
          ("主要发现", "6 题均值 4.02–4.55，作者读作“总体积极”；α 分维度 .714 / .835；Q6 与机辅课程成绩相关 r = .746，"
           "Q2 与笔译课成绩相关 r = −.419（语言能力越强，越怀疑 MTPE 能提升质量）；开放题归纳出 4 个正面主题与 4 个负面主题"
           "（后编辑标准不清、质量不稳定、引擎难选、技术障碍）；结论：面向本科的 MTPE 流程可行，应作为子能力纳入教学。")],
      (X0, 2.44, 11.7, 3.95), size=13.6, gap=0.42, dsize=12.0,
      close="先说清楚：这篇不算“造假”——数据是真的，统计也报了。问题只有一个：这份问卷撑不起它被要求撑的结论。",
      csize=11.4, ccolor=F)
page(s, 5, tag="NEGATIVE CASE · OVERVIEW")

# ============================================================ 06 反例流程
s = S()
title(s, "反例｜操作流程：六步都“做了”，每步只写了半句", "左栏：论文里实际呈现的做法 ｜ 右栏：因此缺掉的那一环")
rows = [("① 目的与工具错位", "研究目标是“提出并论证可行流程”，问卷只回答态度与相关",
         "可行性判断需要专家评审或小样本试点，问卷数据无法替代"),
        ("② 借框架不借结构", "套用 TAM 两个构念，各配 3 条自拟陈述", "无内容评审、无前测、未检验 3 题是否真落在两个维度上（只报了 α）"),
        ("③ 抽样口径模糊", "同一所高校 127 名本科生（39 大三 / 88 大四）", "便利抽样；未报发放与回收、无有效问卷判定规则 → 应答率与自选择偏差无从评估"),
        ("④ 施测一次成型", "线上一次填答，题目直接呈现", "无前后照应题、无注意力检查、未谈共同方法偏差与社会赞许"),
        ("⑤ 分析层级错位", "拿单个题项与课程成绩做 Pearson 相关；均值高于中点即称“积极”",
         "把题项当构念分数用；6 点量表是序数量表，且 4.02–4.55 无比较基线、无检验量"),
        ("⑥ 写作代替解释", "三组发现并列陈列，收束到“建议纳入课程”", "未说明课程成绩受评分标准等影响，也未交代横断面自陈不可推因果")]
stack3(s, rows, X0, 11.7, 2.30, gap=0.24, ksize=12.0, bsize=11.2, kw=1.95, lw=2.62, right=6.90, bottom=6.94)
page(s, 6, tag="NEGATIVE CASE · WHERE IT BREAKS")

# ============================================================ 07 反例六病
s = S()
title(s, "反例｜六病：每一条都有一个“最小改法”", "不必推翻研究设计，只补该补的那一行；这也是答辩时最容易被追问的六处")
tb, tf = add_tb(s, 3.06, 2.06, 4.0, 0.3)
para(tf, "病灶", size=10, color=FAINT, bold=True, first=True)
tb, tf = add_tb(s, 7.40, 2.06, 4.0, 0.3)
para(tf, "最小改法", size=10, color=FAINT, bold=True, first=True)
ills = [("病① 构念—题项脱钩", "“感知有用性”＝3 条自拟题，结构未经检验",
         "引用已验证量表的条目并在附录标来源；至少补一次 EFA 或报告题总相关"),
        ("病② 单题当构念用", "用 Q6 一题与成绩相关，推出“态度好＝技术能力强”",
         "先按维度合成均分（并报告 α），再做相关；措辞只写“关联”不写“影响”"),
        ("病③ 拿代理效标解释", "课程成绩里混着评分松紧、出勤、文本难度差异",
         "交代成绩口径并控制年级与性别，或改用盲评的翻译测试卷作为效标"),
        ("病④ 样本与总体错位", "一所学校 127 人 → 结论写“本科翻译教学可行”",
         "把结论收缩到“该校学生”，或补分层抽样说明；报出发放/回收/有效三份数字"),
        ("病⑤ 横断面冒充纵向", "大三 vs 大四的差异被读作“一年内态度变化”",
         "改写成“两个年级之间存在差异”，或真正做同组前后测（pre-post）"),
        ("病⑥ 过程不可复核", "18 份开放题因“相关性有限”被剔除，编码规则未披露",
         "预先写清理废标准与编码步骤，附两人独立编码的一致性或对账记录")]
stack3(s, ills, X0, 11.7, 2.36, gap=0.22, ksize=12.0, bsize=10.9, kw=2.30, lw=2.04, right=7.40, ccolor=F, bottom=6.94)
page(s, 7, tag="NEGATIVE CASE · SIX FAULTS")

# ============================================================ 08 并排对照
s = S()
title(s, "同一道题，两种做法", "两篇问的其实是同一件事：这份问卷，凭什么得出这个结论？")
tb, tf = add_tb(s, 6.86, 2.02, 2.4, 0.3)
para(tf, "口译（正例）", size=10.5, color=INK_, bold=True, first=True)
tb, tf = add_tb(s, 9.44, 2.02, 2.9, 0.3)
para(tf, "笔译 MTPE（反例）", size=10.5, color=INK_, bold=True, first=True)
rows = [("题项从哪来", "日志＋访谈＋三轮评审，74→28→21", "自拟 6 条，无评审、无前测"),
        ("测的是什么结构", "4 维度构念，EFA 复核吻合", "2 构念各 3 题，结构未检验"),
        ("谁在填、来了多少", "144 名同一训练阶段学员，纳入条件写明", "127 名两个年级学生，回收口径未写"),
        ("信度", "α = .901，并附题项级分析", "分维度 α = .714 / .835"),
        ("效度", "CVI ＋ EFA  三项效标相关", "以课程成绩做单题相关"),
        ("结论走到哪一步", "可测、可区分训练阶段、可用于教学诊断", "流程可行、应纳入课程（超出数据）")]
yy = stack3(s, rows, X0, 11.7, 2.44, gap=0.20, ksize=11.4, bsize=10.8, kw=2.05, lw=2.55, right=6.86, bottom=5.95)
yy = stack2(s, [("把这六行倒过来，就是你的自检表",
             "题项哪来的？测的还是我说的那个东西吗？谁填的、来了多少、留下多少？信度写在哪一行？"
             "效度证据是哪一条？我的结论是不是比数据多走了两步？")],
      X0, 5.75, yy + 0.16, ksize=12.6, bsize=11.0, gap=0.04, bottom=6.90)
tb, tf = add_tb(s, X0, 5.86, 5.7, 1.0)
para(tf, "把这六行倒过来，就是你的自检表", size=12.6, color=INK_, bold=True, first=True)
para(tf, "题项哪来的？测的还是我说的那个东西吗？谁填的、来了多少、留下多少？信度写在哪一行？"
         "效度证据是哪一条？我的结论是不是比数据多走了两步？",
     size=11.0, color=SUB, lh=1.44, before=4)
page(s, 8, tag="SIDE BY SIDE")

# ============================================================ 09 自检清单
s = S()
title(s, "交稿前的 10 行清单", "勾不上的那一行，就是你论文里最可能被追问的地方")
items = ["1. 研究问题写成一句可回答的话，并标明它要的是“描述 / 比较 / 关联”哪一类答案",
         "2. 构念 → 维度 → 题项，有一张能互相对齐的表（哪怕只有 6 题）",
         "3. 题项来源写清：沿用（注出处）／改编（列改动）／自编（说生成过程）",
         "4. 总体、抽样框、纳入与排除条件，各一句话",
         "5. 发放 / 回收 / 有效三份数字与回收率；无效判定规则先定后用",
         "6. 前测或专家评审留痕：几人、评什么、删了几题、为什么",
         "7. 信度（α 或 ω）；有多维结构就报 EFA 或 CFA 指标",
         "8. 指导语含身份、目的、时长、匿名与用途、自愿与退出；敏感题后置",
         "9. 题型与统计对应：单选多选→频数与卡方，量表→均值与 α，排序→秩相关，开放→编码后转主题",
         "10. 结论分寸：横断面自陈只说关联与分布，不说因果与机制"]
y = 1.96
for it in items:
    ls, sz, used = fit_wrap(it, 11.5, 0.34, 12.4, lh=1.30)
    tb, tf = add_tb(s, X0, y, 11.6, used + 0.02)
    para(tf, "\n".join(ls), size=sz, color=INK_, first=True, lh=1.30)
    y += used + 0.075
tb, tf = add_tb(s, X0, y + 0.06, 11.6, 0.4)
para(tf, "课堂提示：这十条全部能在正例那篇里找到对应段落——去抄它的结构，别抄它的内容。",
     size=10.8, color=SUB, first=True, lh=1.4)
page(s, 9, tag="CHECKLIST")

# ============================================================ 10 思考题 + 资料
s = S()
title(s, "4 道思考题 ＋ 资料清单", "讨论约 4 分钟；每题本组先自答一句，再点人")
qs = [("Q1 取舍", "口译那篇为“笔记熟练度”做了四道检验（内容效度、结构效度、效标效度、信度）。若你只有一学期、30 份问卷的预算，"
       "哪两道必须保住、哪一道可以先不做？说出取舍理由。"),
      ("Q2 补一个统计量", "反例里“6 题均值都高于中点”被读作“总体积极”。如果审稿人只允许你补一个统计量来堵住这个漏洞，"
       "你补什么？为什么是它？"),
      ("Q3 追问样本", "由 127 名本校本科生推出“本科翻译教学应引入 MTPE”。要把结论推广到 MTI 群体，"
       "最少需要补哪三条关于样本与施测的信息？"),
      ("Q4 反身", "我们自己的课程论文也在用问卷。对照刚才六病，你现在最担心自己踩到哪一条？"
       "请给出那一条对应的“最小改法”（一句话即可）。")]
stack2(s, qs, X0, 6.5, 2.04, gap=0.14, ksize=12.4, bsize=11.0)
tb, tf = add_tb(s, 7.44, 2.00, 4.6, 0.3)
para(tf, "论文与学习资料（可直接点开核对）", size=10.6, color=INK_, bold=True, first=True)
refs = [("正例全文与量表附录（广外实验室成果页，含方法与数据）",
         "bcdlab.gdufs.edu.cn/info/1025/2025.htm"),
        ("正例原文：《外语教学与研究》51(6): 925-937（知网/期刊官网检索标题）",
         "外语教学与研究 · 2019 年第 6 期"),
        ("反例原文（Wiley 开放获取，全文与 4 张统计表可读）",
         "onlinelibrary.wiley.com/doi/10.1155/2022/5793054"),
        ("同类笔译态度调查（可作正例加强）：Çetiner & İşisağ (2019) IJLET 7(1): 110-120",
         "doi.org/10.18298/ijlet.3242"),
        ("口译问卷研究范例：张威（2013）会议口译员职业角色自我认定的调查研究，《中国翻译》(2): 17-25（135 份有效问卷）",
         "知网检索标题"),
        ("方法读物：吴明隆《问卷统计分析实务》；Fink, A. Conducting Research Surveys（项目分析与删题标准）",
         "图书馆社科借阅区"),
        ("检索式：CNKI 主题=翻译 AND（问卷 OR 量表）AND（信效度）；WoS TS=((translation OR interpreting) AND questionnaire AND validity)",
         "可直接复制使用")]
yy = 2.36
for a, b in refs:
    al, asz, au = _lines(a, 4.5, 9.5, 1.32, ret=True)
    bl, bsz, bu = _lines(b, 4.5, 8.6, 1.28, ret=True)
    tb, tf = add_tb(s, 7.44, yy, 4.6, au + bu + 0.05)
    para(tf, al, size=asz, color=INK_, first=True, lh=1.32)
    para(tf, bl, size=bsz, color=FAINT, lh=1.28, font=MONO, before=1)
    yy += au + bu + 0.12
page(s, 10, tag="DISCUSSION & SOURCES")

# ---------------------------------------------------------------- 纯白化：删除所有装饰色块
prs.save(OUT)
prs2 = Presentation(OUT)
removed = 0
for sl in prs2.slides:
    for sh in list(sl.shapes):
        try:
            if sh.has_text_frame and sh.text_frame.text.strip():
                continue
        except Exception:
            pass
        sh._element.getparent().remove(sh._element)
        removed += 1
prs2.save(OUT)
print("saved:", OUT, "slides:", len(prs2.slides._sldIdLst), "decor shapes removed:", removed)
