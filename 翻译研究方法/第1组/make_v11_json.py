# -*- coding: utf-8 -*-
"""从 deck_content_v9.json 生成 deck_content_v11.json：
1. 补充概念解释（Cronbach's α、自拟题、未报告）
2. 新增两张流程图页：正例研究流程总览 / 反例证据链断裂
"""
import json, os

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, "deck_content_v9.json"), encoding="utf-8") as f:
    data = json.load(f)

slides = data["slides"]
by_page = {s["page"]: s for s in slides}

# ---- 1) 补充概念解释 ----
by_page[11]["bullets"] = [
    "被试：144 名有近一学年笔记交传训练的本科生。",
    "信度：总量表 Cronbach's α = .901，各维度也较高。α 是「内部一致性」指标，取值 0–1，越接近 1 说明题目之间越一致、测量越稳定；通常 ≥ .7 可接受，≥ .9 优秀。",
]

by_page[16]["bullets"] = [
    "借了 TAM 框架说测「有用性」和「易用性」，但 6 道题是作者自己拟写的——没有沿用现成量表的条目，也没有请专家评审、没有先试测。",
    "换句话说，这 6 题从没验证过是不是真的在测那两个概念，题目质量没有保证。",
    "简单说：借了框架的名字，没借框架的骨架，题目到底测的是什么说不清。",
]

by_page[17]["bullets"] = [
    "样本：127 名本科生（大三 39、大四 88），西北一所高校。",
    "抽样方式是便利抽样，但「发出去多少份、收回多少份、有效多少份」都没有报告——少了这三份数字，就不知道有多少人没回、被剔除，也就无法判断结果到底能代表谁。",
    "线上一次填答，也没有照应题、注意力检查来筛掉乱填的问卷。",
]

# ---- 2) 新增流程图页 ----
flow_pos = {
    "page": 8, "type": "flow",
    "title": "正例研究流程：从构念到一份可复用的量表",
    "steps": [
        "界定构念\n两阶段 × 四维度",
        "编制题项\n74 → 28 → 21",
        "正式施测\n144 名学生",
        "信度检验\nα = .901",
        "效度验证\nCVI / EFA / 效标",
    ],
}
flow_break = {
    "page": 21, "type": "flow_break",
    "title": "反例：证据链在哪一环断裂",
    "steps": [
        {"name": "构念与题项", "issue": "自拟 6 题，未验证"},
        {"name": "抽样与施测", "issue": "发放 / 回收未报告"},
        {"name": "统计与分析", "issue": "单题当构念用"},
        {"name": "结论", "issue": "横断面推因果"},
    ],
}

# 重建列表：插入 flow(在 page 8 前，即原来 page 8 位置)，插入 flow_break(在 page 21 前)
new_slides = []
for s in slides:
    if s["page"] == 8:
        new_slides.append(flow_pos)
    if s["page"] == 21:
        new_slides.append(flow_break)
    new_slides.append(s)

# 重新编号
for i, s in enumerate(new_slides, 1):
    s["page"] = i

data["slides"] = new_slides
out = os.path.join(BASE, "deck_content_v11.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
print("saved:", out, "| 页数:", len(new_slides))
for s in new_slides:
    if s["type"] in ("flow", "flow_break"):
        print("  新增流程图:", s["page"], s["type"], s["title"])
