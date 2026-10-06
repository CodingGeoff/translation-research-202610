# -*- coding: utf-8 -*-
"""
V12 论文截图生成器（手动精确裁剪版）：
- 每张截图指定裁剪 rect，把证据句完整截进来（而非只截数字）
- 标注关键词画红色加粗下划线：中文只标题项来源/核心统计量，英文保留少量
运行：python generate_underlines.py
"""
import fitz, os

ESS = r"d:\10_Workspace\translation-research-202610\翻译研究方法\essay"
ASD = r"d:\10_Workspace\translation-research-202610\翻译研究方法\第1组\assets"
ZH = os.path.join(ESS, "口译笔记熟练度量表的开发_周金华.pdf")
FAN = os.path.join(ESS, "On_Postediting_of_Machine_Translation_and_Workflow.pdf")

RED = (0.78, 0, 0)

# (输出名, pdf, 页码0-based, 裁剪rect(x0,y0,x1,y1), 标注关键词)
JOBS = [
    # ---- 正例 ----
    ("zheng_p01_title.png", ZH, 0, (100, 140, 415, 405), []),
    ("zheng_p04_construct.png", ZH, 3, (85, 162, 425, 262), ["72 份", "4 位", "6 位"]),
    ("zheng_p05_itemwrite.png", ZH, 4, (75, 66, 470, 306), ["译员表达", "5 名", "3 名", "74 个", "3 轮"]),
    ("zheng_p06_itemtable.png", ZH, 5, (62, 40, 478, 318), []),
    ("zheng_p05_pilot.png", ZH, 4, (75, 295, 425, 436), []),
    ("zheng_p08_reliability.png", ZH, 7, (190, 82, 435, 280), [".901"]),
    ("zheng_p08_cvi.png", ZH, 7, (50, 432, 420, 542), ["CVI"]),
    ("zheng_p09_efa.png", ZH, 8, (50, 66, 435, 176), ["60.49"]),
    ("zheng_p10_criterion.png", ZH, 9, (138, 192, 478, 478), [".473", ".539", ".556"]),
    # ---- 反例 ----
    ("fan_p01_title.png", FAN, 0, (36, 140, 565, 300), ["On Postediting"]),
    ("fan_p03_tam.png", FAN, 2, (295, 290, 568, 412), ["technology acceptance model", "perceived usefulness", "perceived ease of use"]),
    ("fan_p04_participants.png", FAN, 3, (36, 205, 392, 372), ["127"]),
    ("fan_p04_design.png", FAN, 3, (308, 286, 570, 625), ["6 statements", "3 items"]),
    ("fan_p05_corr.png", FAN, 4, (38, 496, 306, 616), ["Q6", "Q2"]),
    ("fan_p06_alpha.png", FAN, 5, (92, 182, 412, 252), ["0.714", "0.835"]),
]

os.makedirs(ASD, exist_ok=True)


def merge_rects(rects, y_gap=4, x_gap=8):
    rects = sorted(rects, key=lambda r: (round(r.y0), r.x0))
    merged = []
    for r in rects:
        if merged and abs(r.y0 - merged[-1].y0) <= y_gap and r.x0 - merged[-1].x1 <= x_gap:
            m = merged[-1]
            merged[-1] = fitz.Rect(m.x0, min(m.y0, r.y0), r.x1, max(m.y1, r.y1))
        else:
            merged.append(fitz.Rect(r))
    return merged


for outname, pdf, idx, clip_t, mark_kws in JOBS:
    doc = fitz.open(pdf)
    page = doc[idx]
    clip = fitz.Rect(*clip_t)
    pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=clip, alpha=False)
    outp = os.path.join(ASD, outname)
    pix.save(outp)
    print("%-28s 尺寸%dx%d" % (outname, pix.width, pix.height))
print("done")
