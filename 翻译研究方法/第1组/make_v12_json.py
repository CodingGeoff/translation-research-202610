# -*- coding: utf-8 -*-
"""从 deck_content_v11.json 生成 deck_content_v12.json：
1. 正例补充「生成题库」的题项来源原文截图（日志+访谈 / 译员表达+评审+74题）
2. 重构反例：每个批评点配论文原文截图，并修正判断
   - 不再说「缺乏信度检验」→ 论文其实报了 α，改为「只有信度、没有效度」
   - 不再说「便利抽样」→ 论文根本没写抽样方法，改为「抽样方法 / 发放回收未报告」
"""
import json, os

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, "deck_content_v11.json"), encoding="utf-8") as f:
    data = json.load(f)

slides = data["slides"]
by_page = {s["page"]: s for s in slides}

# ---- 1) 正例：第一步 + 第二步补题项来源截图 ----
by_page[9]["bullets"] = [
    "作者先给「笔记熟练度」下可操作定义：拆成笔记记录、笔记理解两个阶段，再细分为四个维度。",
    "构念从哪来？不是凭空想——分析了 72 份学员学习日志，访谈了 4 位学生译员、6 位口译教师。",
    "先把「要测什么」定义清楚，才能据此写出对应的题项。",
]
by_page[9]["image"] = "zheng_p06_itemtable.png"
by_page[9]["image2"] = "zheng_p04_construct.png"
by_page[9]["caption"] = "四个维度（表 1）"
by_page[9]["caption2"] = "构念来源：日志 + 访谈（p4）"

by_page[10]["bullets"] = [
    "题项来源：借助上面的学习日志和访谈数据来写，不是凭空编。",
    "题干尽量采用译员自己的表达，避免学术腔。",
    "写完后请 5 名口译学员、3 名口译教师评审，删改后再讨论；初始题库 74 题，经 3 轮评审历时半年，得到 28 题预试量表。",
]
by_page[10]["image"] = "zheng_p05_itemwrite.png"
by_page[10]["caption"] = "题项编写：译员表达 + 评审 + 74 题（论文 p5）"

# ---- 2) 重构反例部分（page 16 是 section，17-23 重排为 17-24）----
new_anti = [
    {
        "page": 17, "type": "slide",
        "title": "反例论文与研究问题",
        "bullets": [
            "文献：Yang, Z. & Mustafa, H. R. (2022). On postediting of machine translation and workflow for undergraduate translation program in China. HBET, 5793054.",
            "研究问题：RQ1 学生对 MTPE 的正面/负面反应；RQ2 反应是否与两门课程成绩相关；RQ3 使用障碍是什么。",
            "方法：混合方法——问卷（6 道题 + 4 道开放题）+ 定性编码。",
        ],
        "image": "fan_p01_title.png", "caption": "论文首页",
    },
    {
        "page": 18, "type": "slide",
        "title": "环节一：借了 TAM 的框架",
        "bullets": [
            "框架：借 TAM，声称测「有用性」和「易用性」两个构念。",
            "依据这两个构念来设计问卷，但题目却对不上框架的「骨架」。",
        ],
        "image": "fan_p03_tam.png", "caption": "理论框架 TAM（论文 p3）",
    },
    {
        "page": 19, "type": "slide",
        "title": "环节一：6 道题自拟、未验证",
        "bullets": [
            "题项：6 个陈述句（有用性 3 题 + 易用性 3 题）+ 4 道开放题，6 点李克特。",
            "问题：全文只列出这 6 题，没交代沿用哪个现成量表、没有专家评审、没有试测。",
            "等于题目从没验证过是不是真的在测那两个概念。",
        ],
        "image": "fan_p04_design.png", "caption": "问卷 6 题 + 4 开放题（论文 p4）",
    },
    {
        "page": 20, "type": "slide",
        "title": "环节二：抽样与回收未报告",
        "bullets": [
            "样本：西北一所高校翻译专业 127 人（大三 39、大四 88）。",
            "问题：只写「127 名学生参与」，没写抽样方法（便利/随机），也没写发放、回收、有效各多少。",
            "后果：无法判断应答率、代表性，也无法排除自选择偏差。",
        ],
        "image": "fan_p04_participants.png", "caption": "被试说明（论文 p4）",
    },
    {
        "page": 21, "type": "slide",
        "title": "环节三：只有信度、没有效度",
        "bullets": [
            "信度：报了 Cronbach's α（有用性 .714、易用性 .835），这一步做到了。",
            "效度：全文没有内容效度、没有因子分析，6 题是否真测「有用性/易用性」两个构念，未验证。",
            "问题：只有信度、没有效度，量表的「有效性」悬空。",
        ],
        "image": "fan_p06_alpha.png", "caption": "信度 α=.714 / .835（论文 p6）",
    },
    {
        "page": 22, "type": "slide",
        "title": "环节四：把一道题当一个构念",
        "bullets": [
            "做法：拿 Q6 单题与 CAT 课程成绩相关（r=.746），Q2 单题与翻译实践成绩相关（r=-.419）。",
            "问题：把「一道题」当成「一个构念」的分数来用，而不是先把同维度 3 题合成平均分再算。",
            "后果：单题的相关系数不稳定，结论容易失真。",
        ],
        "image": "fan_p05_corr.png", "caption": "单题相关分析（论文 p5）",
    },
    {
        "page": 23, "type": "slide",
        "title": "环节五：结论越过了数据",
        "bullets": [
            "结论：从「127 名学生一次横断面自陈、态度偏积极」，得出「应把 MTPE 纳入本科教学」。",
            "问题：横断面问卷 + 单题相关，只能说明「有关联」，推不出「应该做」。",
        ],
    },
    {
        "page": 24, "type": "flow_break",
        "title": "反例：证据链在哪一环断裂",
        "steps": [
            {"name": "构念与题项", "issue": "自拟 6 题、未验证"},
            {"name": "抽样与施测", "issue": "抽样方法 / 回收未报告"},
            {"name": "信度与效度", "issue": "只有信度、没有效度"},
            {"name": "统计与结论", "issue": "单题当构念 + 越界"},
        ],
    },
    {
        "page": 25, "type": "slide",
        "title": "反例小结：不是造假，是证据链断裂",
        "bullets": [
            "反例的数据是真的、信度也报了，并不是造假。",
            "问题在于证据链断裂：题项没来源、抽样没交代、只有信度没有效度、单题当构念。",
            "最根本的一点：结论超出了数据能承担的范围。",
        ],
    },
]

# 删除原反例正文（page 17-23），替换为新序列；section(16) 保留
kept = [s for s in slides if not (17 <= s["page"] <= 23)]
kept = [s for s in kept if not (s["type"] in ("flow_break",)) or s["page"] != 23]
# 上面一行多余，直接重建：先删除 17-23，再插入新反例
kept = [s for s in slides if s["page"] < 17 or s["page"] > 23]
# 找到 section 16 的位置，在其后插入 new_anti
result = []
for s in kept:
    result.append(s)
    if s["page"] == 16:
        result.extend(new_anti)

# 重新编号
for i, s in enumerate(result, 1):
    s["page"] = i

data["slides"] = result
out = os.path.join(BASE, "deck_content_v12.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
print("saved:", out, "| 页数:", len(result))
for s in result:
    if s["type"] in ("flow", "flow_break") or s["page"] in (9, 10, 18, 19, 20, 21):
        imgs = s.get("image", "") + (" + " + s.get("image2", "") if s.get("image2") else "")
        print("  S%02d %-12s %s %s" % (s["page"], s["type"], s["title"][:14], imgs))
