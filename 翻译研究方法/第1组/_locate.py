# -*- coding: utf-8 -*-
import fitz, os

ESS = r"d:\10_Workspace\translation-research-202610\翻译研究方法\essay"
ZH = os.path.join(ESS, "口译笔记熟练度量表的开发_周金华.pdf")
FAN = os.path.join(ESS, "On_Postediting_of_Machine_Translation_and_Workflow.pdf")

# (pdf, 页码0-based, 搜索关键词)
tasks = [
    (ZH, 5, "维度"),
    (ZH, 5, "表1"),
    (ZH, 6, "表2"),
    (ZH, 7, "Cronbach"),
    (ZH, 8, "因素分析"),
    (ZH, 8, "60.49"),
    (FAN, 3, "127"),
    (FAN, 3, "Participants"),
    (FAN, 4, "Q1"),
    (FAN, 5, "mean"),
]

for pdf, idx, kw in tasks:
    doc = fitz.open(pdf)
    page = doc[idx]
    rects = page.search_for(kw)
    print("== %s p%d 搜索 %r -> %d 处" % (os.path.basename(pdf), idx+1, kw, len(rects)))
    for r in rects[:5]:
        print("    rect: x0=%.0f y0=%.0f x1=%.0f y1=%.0f (页高 %.0f)" % (r.x0, r.y0, r.x1, r.y1, page.rect.height))
    doc.close()
