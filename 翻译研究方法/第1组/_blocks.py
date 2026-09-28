# -*- coding: utf-8 -*-
import fitz, os

ESS = r"d:\10_Workspace\translation-research-202610\翻译研究方法\essay"
ZH = os.path.join(ESS, "口译笔记熟练度量表的开发_周金华.pdf")
FAN = os.path.join(ESS, "On_Postediting_of_Machine_Translation_and_Workflow.pdf")

pages = [(ZH, 5, "正例p6 表1"), (ZH, 6, "正例p7 表2"), (ZH, 7, "正例p8 信度"),
         (ZH, 8, "正例p9 因素分析"), (FAN, 3, "反例p4 被试"),
         (FAN, 4, "反例p5 问卷"), (FAN, 5, "反例p6 数据")]

out = []
for pdf, idx, label in pages:
    doc = fitz.open(pdf)
    page = doc[idx]
    blocks = page.get_text("blocks")
    out.append("\n==== %s (页宽%.0f 高%.0f) ====" % (label, page.rect.width, page.rect.height))
    for b in blocks:
        x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4]
        t = "".join(ch for ch in text.replace("\n", " ").strip() if ch.isprintable())[:60]
        out.append("  y%.0f-%.0f x%.0f-%.0f | %s" % (y0, y1, x0, x1, t))
    doc.close()

with open(r"d:\10_Workspace\translation-research-202610\翻译研究方法\第1组\_blocks.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done, wrote _blocks.txt")

