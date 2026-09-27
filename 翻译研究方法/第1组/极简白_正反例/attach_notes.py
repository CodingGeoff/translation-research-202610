# -*- coding: utf-8 -*-
"""把讲稿（讲稿_10min_逐字稿.md）按页写进 pptx 的“演讲者备注”窗格，放映/排练时可直接看稿。
用法：先 python3 build_simple.py ，再 python3 attach_notes.py （每次重排版后都要重跑本脚本）。
"""
import io, os, re, sys
from pptx import Presentation
from pptx.util import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
PPTX = os.path.join(HERE, "问卷调查法_极简白_黑字_10min.pptx")
MD = os.path.join(HERE, "讲稿_极简白_10min.md")
if not os.path.exists(PPTX):
    sys.exit("找不到 pptx，请先运行 build_simple.py")

lines = io.open(MD, encoding="utf-8").read().split("\n")
cut = next((i for i, l in enumerate(lines) if l.startswith("# 四道思考题")), len(lines))
lines = lines[:cut]                                   # 备注里只放逐字稿

secs, cur = [], None
for ln in lines:
    if ln.startswith("## "):
        cur = {"head": ln[3:].strip(), "body": []}
        secs.append(cur)
    elif cur is not None:
        st = ln.strip()
        if st == "---":
            cur = None                                # 正文结束，后面是主持人口径
        elif st:
            cur["body"].append(st)

prs = Presentation(PPTX)
slides = list(prs.slides)
if len(secs) != len(slides):
    print("注意：讲稿段落数 %d 与幻灯片数 %d 不一致（按较少者写入）" % (len(secs), len(slides)))
spoken = 0
for i, sl in enumerate(slides[:len(secs)]):
    sec = secs[i]
    head = sec["head"]                       # 讲稿小标题整行：页码 ｜ 环节 ｜ 时间 ｜ 最迟 ｜ 字数
    paras = []
    for p in sec["body"]:                                # 一行＝一口气（一段）
        p = re.sub(r"\*\*(.+?)\*\*", r"\1", p).strip()
        p = re.sub(r"^（[^）]{1,6}）", "", p)             # 去掉（问题）（收束）等提示标签
        if p:
            paras.append(p)
    spoken += sum(len(p) for p in paras)
    tf = sl.notes_slide.notes_text_frame
    tf.clear()
    tf.text = "\n".join([head] + paras)
    for para in tf.paragraphs:
        for r in para.runs:
            r.font.size = Pt(12)
prs.save(PPTX)
print("notes written:", min(len(secs), len(slides)), "slides ｜ 口播字数", spoken,
      "｜ 272字/分 ≈ %d:%02d ｜ 页数 %d" % (int(spoken/272*60)//60, round(spoken/272*60%60), len(slides)))
