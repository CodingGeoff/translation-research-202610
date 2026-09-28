# -*- coding: utf-8 -*-
"""
工具链稳健性测试：验证 deck_content_v5.json 完整性 + build_v5/build_v6 可重复生成。
运行：python test_deck.py
"""
import os, sys, json, subprocess, tempfile

BASE = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(BASE, "deck_content_v5.json")
ASD = os.path.join(BASE, "assets")

VALID_TYPES = {"cover", "toc", "section", "slide", "end"}
TEXT_FIELDS = {"title", "subtitle", "subtitle_en", "bullets", "lines", "items", "caption"}

passed = 0
failed = 0


def check(cond, msg):
    global passed, failed
    if cond:
        passed += 1
        print("  [OK] %s" % msg)
    else:
        failed += 1
        print("  [FAIL] %s" % msg)


def validate_json(path):
    print("== 校验 %s ==" % os.path.basename(path))
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    check("meta" in data, "存在 meta")
    check("slides" in data and isinstance(data["slides"], list), "存在 slides 列表")
    slides = data["slides"]
    check(len(slides) > 0, "slides 非空（共 %d 页）" % len(slides))

    pages = []
    for sd in slides:
        pages.append(sd.get("page"))
        t = sd.get("type")
        check(t in VALID_TYPES, "第 %s 页 type 合法（%s）" % (sd.get("page"), t))
        # 结构字段存在
        check("page" in sd, "第 %s 页有 page" % sd.get("page"))
        # slide 必须有 title + bullets
        if t == "slide":
            check(bool(sd.get("title")), "第 %s 页 slide 有 title" % sd.get("page"))
            check("bullets" in sd and isinstance(sd["bullets"], list) and len(sd["bullets"]) > 0,
                  "第 %s 页 slide 有 bullets" % sd.get("page"))
        # 图片文件存在
        if sd.get("image"):
            p = os.path.join(ASD, sd["image"])
            check(os.path.exists(p), "第 %s 页图片存在（%s）" % (sd.get("page"), sd["image"]))
        # 文字字段类型
        for k, v in sd.items():
            if k in TEXT_FIELDS:
                if isinstance(v, list):
                    check(all(isinstance(x, str) for x in v), "第 %s 页 %s 列表元素为字符串" % (sd.get("page"), k))
                else:
                    check(isinstance(v, str), "第 %s 页 %s 为字符串" % (sd.get("page"), k))

    # 页码连续
    check(pages == list(range(1, len(slides) + 1)), "页码连续 1..%d" % len(slides))
    return data


def run_build(script, outname):
    print("== 运行 %s ==" % script)
    out = os.path.join(BASE, outname)
    r = subprocess.run([sys.executable, os.path.join(BASE, script)],
                       cwd=BASE, capture_output=True, text=True, encoding="utf-8", errors="replace")
    ok = r.returncode == 0 and os.path.exists(out)
    check(ok, "%s 成功生成 %s" % (script, outname))
    if not ok:
        print("    stdout:", r.stdout[-500:])
        print("    stderr:", r.stderr[-500:])
        return None
    # 用 python-pptx 重新打开验证
    try:
        from pptx import Presentation
        prs = Presentation(out)
        check(len(prs.slides._sldIdLst) == len(json.load(open(JSON_PATH, encoding="utf-8"))["slides"]),
              "%s 页数与 JSON 一致" % outname)
    except Exception as e:
        check(False, "%s 可被 python-pptx 打开（%s）" % (outname, e))
    return out


if __name__ == "__main__":
    print("######## 问卷调查法正反例 PPT 工具链测试 ########\n")
    validate_json(JSON_PATH)
    run_build("build_v5.py", "问卷调查法正反例_v5.pptx")
    run_build("build_v6.py", "问卷调查法正反例_v6.pptx")

    print("\n======== 结果：%d 通过，%d 失败 ========" % (passed, failed))
    sys.exit(1 if failed else 0)
