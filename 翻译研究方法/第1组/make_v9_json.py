# -*- coding: utf-8 -*-
"""从 deck_content_v5.json 生成 deck_content_v9.json：
- P9 加图（74题）、P10 换图（试测删题）、P12 拆成两页（内容+结构效度 / 效标关联）、P18 加第二图（反例α）
"""
import json, os

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, "deck_content_v5.json"), encoding="utf-8") as f:
    data = json.load(f)

slides = data["slides"]

# 按 page 建索引
by_page = {s["page"]: s for s in slides}

# P9 加图
by_page[9]["image"] = "zheng_p05_pool.png"
by_page[9]["caption"] = "初始题项库 74 题（论文 p5）"

# P10 换图（试测删题在 p5，替换原表2）
by_page[10]["image"] = "zheng_p05_pilot.png"
by_page[10]["caption"] = "评审 28 题 · 试测 54 人 · 删 7 留 21（论文 p5）"

# P11 信度图 caption 更新
by_page[11]["image"] = "zheng_p08_reliability.png"
by_page[11]["caption"] = "总量表 Cronbach's α = .901（论文 p8）"

# P12 拆成两页
p12 = by_page[12]
p12a = {
    "page": 12, "type": "slide",
    "title": "第五步：效度证据（一）内容与结构效度",
    "bullets": [
        "内容效度：I-CVI、S-CVI 均为 1.00",
        "结构效度：EFA 抽出 4 因子，累计解释 60.49%",
    ],
    "image": "zheng_p08_cvi.png",
    "image2": "zheng_p09_efa.png",
    "caption": "内容效度 CVI=1（p8）",
    "caption2": "因素分析 60.49%（p9）",
}
p12b = {
    "page": 13, "type": "slide",
    "title": "第五步：效度证据（二）效标关联效度",
    "bullets": [
        "效标关联：与练习量 r=.473、动机 r=.539、绩效 r=.556 显著相关（均 p<.01）",
    ],
    "image": "zheng_p10_criterion.png",
    "caption": "效标系数 .473 / .539 / .556（论文 p10）",
}

# P18 加第二图（反例 α）
by_page[18]["image2"] = "fan_p06_alpha.png"
by_page[18]["caption2"] = "信度 Cronbach's α=.714 / .835（论文 p6）"

# 重建 slides 列表：P12 替换为 p12a + p12b，后续 page+1
new_slides = []
for s in slides:
    if s["page"] < 12:
        new_slides.append(s)
    elif s["page"] == 12:
        new_slides.append(p12a)
        new_slides.append(p12b)
    else:
        s2 = dict(s)
        s2["page"] = s["page"] + 1
        new_slides.append(s2)

data["slides"] = new_slides

out = os.path.join(BASE, "deck_content_v9.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
print("saved:", out, "| 页数:", len(new_slides))
