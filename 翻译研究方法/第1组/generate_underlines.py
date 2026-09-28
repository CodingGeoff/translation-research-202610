# -*- coding: utf-8 -*-
"""
V9 论文截图生成器：精确定位 + 红色下划线标注 + 裁剪
- 用下划线（而非色块）标注关键数据，同行矩形合并，避免数字/汉字分格
- 覆盖 PPT 中出现的所有关键数据（α、CVI、r、样本量、删题数、均值等）
运行：python generate_underlines.py
"""
import fitz, os

ESS = r"d:\10_Workspace\translation-research-202610\翻译研究方法\essay"
ASD = r"d:\10_Workspace\translation-research-202610\翻译研究方法\第1组\assets"
ZH = os.path.join(ESS, "口译笔记熟练度量表的开发_周金华.pdf")
FAN = os.path.join(ESS, "On_Postediting_of_Machine_Translation_and_Workflow.pdf")

RED = (0.78, 0, 0)

# (输出名, pdf, 页码0-based, 下划线关键词列表, 裁剪rect(x0,y0,x1,y1))
# 原则：中文论文数字直接可见的不再划线，只对核心统计量（α/CVI/r/均值）红色加粗标注；
#       英文论文保留少量下划线（数字不显眼，需定位）。
JOBS = [
    ("zheng_p01_title.png", ZH, 0, [], (48, 38, 445, 300)),
    ("zheng_p06_itemtable.png", ZH, 5, [], (55, 90, 438, 392)),
    ("zheng_p05_pool.png", ZH, 4, [], (55, 250, 438, 315)),
    ("zheng_p05_pilot.png", ZH, 4, [], (55, 210, 438, 440)),
    ("zheng_p08_reliability.png", ZH, 7, [".901"], (55, 78, 438, 270)),
    ("zheng_p08_cvi.png", ZH, 7, ["CVI"], (55, 288, 438, 545)),
    ("zheng_p09_efa.png", ZH, 8, ["60.49"], (55, 68, 438, 242)),
    ("zheng_p10_criterion.png", ZH, 9, [".473", ".539", ".556"], (55, 210, 438, 445)),
    ("fan_p01_title.png", FAN, 0, ["On Postediting"], (42, 36, 560, 300)),
    ("fan_p04_participants.png", FAN, 3, ["127"], (42, 185, 560, 340)),
    ("fan_p05_questionnaire.png", FAN, 3, ["Q1", "Q2", "Q3", "Q4", "Q5", "Q6"], (42, 400, 560, 640)),
    ("fan_p06_data.png", FAN, 4, ["Mean", "4.55", "4.02"], (42, 325, 560, 445)),
    ("fan_p06_alpha.png", FAN, 5, ["0.714", "0.835"], (42, 180, 560, 240)),
]

os.makedirs(ASD, exist_ok=True)


def merge_rects(rects, y_gap=4, x_gap=8):
    """合并同一行、相邻的矩形，避免数字和汉字分格。"""
    rects = sorted(rects, key=lambda r: (round(r.y0), r.x0))
    merged = []
    for r in rects:
        if merged and abs(r.y0 - merged[-1].y0) <= y_gap and r.x0 - merged[-1].x1 <= x_gap:
            m = merged[-1]
            merged[-1] = fitz.Rect(m.x0, min(m.y0, r.y0), r.x1, max(m.y1, r.y1))
        else:
            merged.append(fitz.Rect(r))
    return merged


for outname, pdf, idx, keywords, clip in JOBS:
    doc = fitz.open(pdf)
    page = doc[idx]
    shape = page.new_shape()
    drawn = 0
    for kw in keywords:
        rects = page.search_for(kw)
        for r in merge_rects(rects):
            y = r.y1 + 1.5  # 文字底部稍下
            shape.draw_line(fitz.Point(r.x0 - 2, y), fitz.Point(r.x1 + 2, y))
            drawn += 1
    shape.finish(color=RED, width=4)
    shape.commit()
    r = fitz.Rect(clip)
    pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=r, alpha=False)
    outp = os.path.join(ASD, outname)
    pix.save(outp)
    print("%-28s 下划线%d条  尺寸%dx%d" % (outname, drawn, pix.width, pix.height))
    doc.close()

print("done")
