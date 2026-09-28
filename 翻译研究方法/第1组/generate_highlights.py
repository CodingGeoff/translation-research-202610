# -*- coding: utf-8 -*-
"""
论文截图精确定位 + 高亮 + 裁剪生成器
对每张图：定位到论文里能佐证该页内容的原文，用黄色高亮标出，再裁剪出该区域。
运行：python generate_highlights.py
"""
import fitz, os

ESS = r"d:\10_Workspace\translation-research-202610\翻译研究方法\essay"
ASD = r"d:\10_Workspace\translation-research-202610\翻译研究方法\第1组\assets"
ZH = os.path.join(ESS, "口译笔记熟练度量表的开发_周金华.pdf")
FAN = os.path.join(ESS, "On_Postediting_of_Machine_Translation_and_Workflow.pdf")

# (输出文件名, pdf, 页码0-based, 高亮关键词, 裁剪rect(x0,y0,x1,y1))
JOBS = [
    ("zheng_p01_title.png", ZH, 0, ["口译笔记熟练度量表的开发"], (48, 38, 445, 300)),
    ("zheng_p06_itemtable.png", ZH, 5, ["维度"], (55, 90, 438, 392)),
    ("zheng_p07_itemanalysis.png", ZH, 6, ["表2", "题项分析结果"], (55, 72, 438, 392)),
    ("zheng_p08_reliability.png", ZH, 7, ["Cronbach"], (55, 72, 438, 262)),
    ("zheng_p09_efa.png", ZH, 8, ["因素分析", "60.49"], (55, 68, 438, 242)),
    ("fan_p01_title.png", FAN, 0, ["On Postediting"], (42, 36, 560, 300)),
    ("fan_p04_participants.png", FAN, 3, ["127"], (42, 185, 560, 340)),
    ("fan_p05_questionnaire.png", FAN, 3, ["Q1", "Q2", "Q3", "Q4", "Q5", "Q6"], (42, 400, 560, 640)),
    ("fan_p06_data.png", FAN, 4, ["Mean", "4.55"], (42, 325, 560, 445)),
]

os.makedirs(ASD, exist_ok=True)

for outname, pdf, idx, keywords, clip in JOBS:
    doc = fitz.open(pdf)
    page = doc[idx]
    # 高亮关键词（半透明黄色）
    shape = page.new_shape()
    hit = 0
    for kw in keywords:
        for rect in page.search_for(kw):
            shape.draw_rect(rect)
            hit += 1
    shape.finish(color=(1, 0.6, 0), fill=(1, 0.85, 0.2), fill_opacity=0.5)
    shape.commit()
    # 裁剪 + 放大
    r = fitz.Rect(clip)
    pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=r, alpha=False)
    outp = os.path.join(ASD, outname)
    pix.save(outp)
    print("%-28s 高亮%d处  尺寸%dx%d" % (outname, hit, pix.width, pix.height))
    doc.close()

print("done")
