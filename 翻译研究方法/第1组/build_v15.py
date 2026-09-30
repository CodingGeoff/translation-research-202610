# -*- coding: utf-8 -*-
"""
V15 构建脚本：把「问卷调查法_翻译研究_优化版PPT.json」做成 PPTX。

- 通过环境变量设置汇报人与学号（第一页展示）：
      SPEAKER_NAME  默认 陈冠臻
      STUDENT_ID    默认 20261210050
- 读取 JSON 后：
      1) 替换封面「汇报人 / 学号」占位符
      2) 补全封面中文副标题（来自 meta.subtitle）
      3) 把没有配图的 caption（关键提醒）合并进正文，避免丢失
- 复用 build_v11.py 的渲染模板生成 V15

用法：
    python build_v15.py
    SPEAKER_NAME=张三 STUDENT_ID=123 python build_v15.py
"""
import os, json, subprocess, sys

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "问卷调查法_翻译研究_优化版PPT.json")
TMP = os.path.join(BASE, "deck_content_v15.json")

NAME = os.environ.get("SPEAKER_NAME", "陈冠臻")
SID = os.environ.get("STUDENT_ID", "20261210050")
speaker = "汇报人：%s    学号：%s" % (NAME, SID)

with open(SRC, encoding="utf-8") as f:
    data = json.load(f)

# 1) 替换 meta 与封面里的汇报人/学号
data["meta"]["speaker"] = speaker

for sd in data["slides"]:
    if sd.get("type") == "cover":
        sd["lines"] = [speaker, data["meta"].get("date", "2026 年 10 月 12 日")]
        # 2) 补全中文副标题（若封面缺失）
        if not sd.get("subtitle") and data["meta"].get("subtitle"):
            sd["subtitle"] = data["meta"]["subtitle"]
    # 3) 无配图的 caption 合并进正文，标为「注：」
    if sd.get("type") == "slide" and sd.get("caption") and not sd.get("image") and not sd.get("image2"):
        sd["bullets"] = list(sd.get("bullets", [])) + ["注：" + sd["caption"]]

with open(TMP, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("汇报人：%s    学号：%s" % (NAME, SID))
subprocess.run([sys.executable, os.path.join(BASE, "build_v11.py"), TMP], check=True)
