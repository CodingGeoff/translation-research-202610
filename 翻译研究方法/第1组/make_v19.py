# -*- coding: utf-8 -*-
"""生成 V19：
1. 引号规范化：英文直引号成对替换为中文引号（确保中文引号）
2. 去掉 caption 合并时产生的重复「注：」前缀（注：注意： → 注意： 等）
"""
import json, os

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "deck_content_v18.json")
DST = os.path.join(BASE, "deck_content_v19.json")

with open(SRC, encoding="utf-8") as f:
    d = json.load(f)

# 去掉「注：」前缀（这些是 caption 合并时加的，而 caption 本身已有「注意：/优点：/关键提醒：/更稳妥的表达：」等前缀）
FIX = 0
for sd in d["slides"]:
    for i, b in enumerate(list(sd.get("bullets", []))):
        if b.startswith("注："):
            sd["bullets"][i] = b[2:]
            FIX += 1

# 引号规范化（英文直引号 → 中文引号）；成对交替
def norm(s):
    out = []
    dq_open = True
    sq_open = True
    for ch in s:
        if ch == '"':
            out.append('\u201c' if dq_open else '\u201d')
            dq_open = not dq_open
        elif ch == "'":
            out.append('\u2018' if sq_open else '\u2019')
            sq_open = not sq_open
        else:
            out.append(ch)
    return "".join(out)

def walk(obj):
    if isinstance(obj, str):
        return norm(obj)
    if isinstance(obj, list):
        return [walk(x) for x in obj]
    if isinstance(obj, dict):
        return {k: walk(v) for k, v in obj.items()}
    return obj

d = walk(d)

with open(DST, "w", encoding="utf-8") as f:
    json.dump(d, f, ensure_ascii=False, indent=2)

print("去掉「注：」前缀 %d 处" % FIX)
print("已写:", DST)
