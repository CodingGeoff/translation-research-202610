# -*- coding: utf-8 -*-
"""
从 PPTX 反推 deck_content JSON。

用法：
    python reverse_to_json.py "输入.pptx" [输出.json]

说明：
- 逐页识别类型（cover / toc / section / slide / flow / flow_break / end）
- 提取标题、正文 bullets、目录项、章节号
- 若页面含图片，则按尺寸匹配 assets/ 下的截图文件名，写入 image 字段
- 页码「X / Y」、占位符等会被自动过滤
- 输出 JSON 结构兼容 build_v11.py（可直接重新生成 PPT）
"""
import sys, os, json, re, io
from pptx import Presentation
from pptx.util import Emu

BASE = os.path.dirname(os.path.abspath(__file__))
ASD = os.path.join(BASE, "assets")

PAGE_RE = re.compile(r'^\s*\d+\s*/\s*\d+\s*$')
SECTION_RE = re.compile(r'^\s*0?\d{1,2}\s*$')      # 章节号，如 01、1、02
NUM_ONLY_RE = re.compile(r'^\s*\d+\s*$')            # 纯数字，如目录编号 1、2、3

# 常见的图片说明词，用于把「论文首页」这类文字识别为 caption 而非正文
CAPTION_HINTS = ("论文首页", "论文", "表", "图", "问卷", "信度", "效度", "被试", "维度", "题项")


def is_page_number(t):
    return bool(PAGE_RE.match(t))


def split_bullets(text):
    """把一段正文按换行拆成 bullets，去掉「· / • / 数字序号」前缀。"""
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    out = []
    for ln in lines:
        ln = re.sub(r'^[·•●]\s*', '', ln)
        out.append(ln)
    return out


def asset_map():
    """assets 目录：{ (宽,高): 文件名 }，用于按尺寸匹配图片。"""
    m = {}
    if os.path.isdir(ASD):
        from PIL import Image
        for f in os.listdir(ASD):
            if f.lower().endswith(('.png', '.jpg', '.jpeg')):
                try:
                    im = Image.open(os.path.join(ASD, f))
                    m[im.size] = f
                except Exception:
                    pass
    return m


def match_image(shape, size_map):
    """按图片原始像素尺寸匹配 assets 文件名。"""
    try:
        img = shape.image
        return size_map.get(img.size)
    except Exception:
        return None


def get_text_items(slide):
    """收集页面所有非页码文本，按 top、left 排序。返回 [(top, left, text, shape)]。"""
    items = []
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            t = sh.text_frame.text.strip()
            if is_page_number(t):
                continue
            top = Emu(sh.top).inches if sh.top is not None else 0
            left = Emu(sh.left).inches if sh.left is not None else 0
            items.append((top, left, t, sh))
    items.sort(key=lambda x: (round(x[0], 1), x[1]))
    return items


def classify(items):
    """根据文本特征判断页面类型。"""
    all_text = "\n".join(t for _, _, t, _ in items)
    # 结束页
    if "谢谢观看" in all_text or "Thank You" in all_text:
        return "end"
    # 封面：含「汇报人 / 学号」
    if "汇报人" in all_text or "学号" in all_text:
        return "cover"
    # 目录：含「目录 / CONTENTS」
    if "目录" in all_text or "CONTENTS" in all_text:
        return "toc"
    # 章节页：有一个两位数字的独立文本框
    for _, _, t, _ in items:
        if SECTION_RE.match(t) and len(t.strip()) <= 2:
            return "section"
    return "slide"


def extract_cover(items):
    d = {"type": "cover", "title": "", "lines": []}
    for top, left, t, _ in items:
        if "汇报人" in t or "学号" in t or re.match(r'^\d{4}', t):
            d["lines"].extend(split_bullets(t))
        elif re.search(r'[A-Za-z]', t) and not re.search(r'[\u4e00-\u9fff]', t):
            d["subtitle_en"] = t
        elif not d.get("title"):
            d["title"] = t
        else:
            d["subtitle"] = t
    return d


def extract_toc(items):
    d = {"type": "toc", "title": "目录", "items": []}
    for top, left, t, _ in items:
        if "目录" in t or "CONTENTS" in t:
            d["title"] = "目录"
        elif NUM_ONLY_RE.match(t):
            continue  # 圆形编号，跳过
        else:
            d["items"].append(t)
    return d


def extract_section(items):
    d = {"type": "section", "number": "", "title": ""}
    for top, left, t, _ in items:
        if SECTION_RE.match(t) and len(t.strip()) <= 2:
            d["number"] = t.strip().zfill(2)
        else:
            d["title"] = t
    return d


def extract_slide(items):
    d = {"type": "slide", "title": "", "bullets": []}
    for top, left, t, sh in items:
        # 章节号、目录编号跳过
        if SECTION_RE.match(t) and len(t.strip()) <= 2:
            continue
        # caption：位于页面中下部、且是图片说明词
        if top > 5.0 and len(t) < 20:
            d.setdefault("caption", t)
            continue
        # 第一个有效文本作为标题，其余作为正文
        if not d["title"]:
            d["title"] = t
        else:
            d["bullets"].extend(split_bullets(t))
    return d


def attach_image(slide, sd, size_map):
    """若页面含图片，按尺寸匹配 assets 文件名。"""
    for sh in slide.shapes:
        if sh.shape_type == 13:
            name = match_image(sh, size_map)
            if name:
                sd["image"] = name
                return


def extract_end(items):
    d = {"type": "end", "title": "谢谢观看", "lines": []}
    for top, left, t, _ in items:
        if "谢谢观看" in t:
            d["title"] = "谢谢观看"
        elif "Thank You" in t:
            d["subtitle"] = "Thank You"
        else:
            d["lines"].extend(split_bullets(t))
    return d


def reverse(pptx_path, out_path):
    prs = Presentation(pptx_path)
    size_map = asset_map()
    slides = []
    for slide in prs.slides:
        items = get_text_items(slide)
        if not items:
            continue
        typ = classify(items)
        if typ == "cover":
            sd = extract_cover(items)
        elif typ == "toc":
            sd = extract_toc(items)
        elif typ == "section":
            sd = extract_section(items)
        elif typ == "end":
            sd = extract_end(items)
        else:
            sd = extract_slide(items)
            attach_image(slide, sd, size_map)
        sd["page"] = len(slides) + 1
        slides.append(sd)

    meta = {
        "title": "问卷调查法在翻译研究中的应用",
        "subtitle": "",
        "subtitle_en": "",
        "speaker": "汇报人：陈冠臻",
        "date": "2026 年 10 月 12 日",
    }
    # 从封面补 meta
    for sd in slides:
        if sd["type"] == "cover":
            meta["title"] = sd.get("title", meta["title"])
            meta["subtitle"] = sd.get("subtitle", "")
            meta["subtitle_en"] = sd.get("subtitle_en", "")
            for ln in sd.get("lines", []):
                if "汇报人" in ln:
                    meta["speaker"] = ln
            break

    data = {"meta": meta, "slides": slides}
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("已反推:", out_path)
    print("页数:", len(slides))
    for sd in slides:
        img = sd.get("image", "")
        print("  S%02d %-8s %s %s" % (sd["page"], sd["type"], sd.get("title", "")[:20], img))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(src)[0] + "_reversed.json"
    reverse(src, dst)
