# -*- coding: utf-8 -*-
"""逐行排版盒校验：文字是否越出所在卡片 / 卡片叠压 / 页脚 / 页面。"""
import sys
from pptx import Presentation
from deck_lib import wrap, tw
from pptx.oxml.ns import qn

EMU_IN = 914400.0


def E(v):
    return float(v) / EMU_IN if v is not None else 0.0


def sp(p, tag):
    e = p._pPr.find(qn(tag)) if p._pPr is not None else None
    if e is None:
        return 0.0
    pt = e.find(qn("a:spcPts"))
    return float(pt.get("val")) / 100.0 / 72.0 if pt is not None else 0.0


def line_rects(sh):
    """返回 (逐行矩形列表, 文本框 x, 文本框右边界)"""
    tf = sh.text_frame
    x, y, w, h = E(sh.left), E(sh.top), E(sh.width), E(sh.height)
    ml, mr = E(tf.margin_left), E(tf.margin_right)
    mt, mb = E(tf.margin_top), E(tf.margin_bottom)
    st = str(sh.shape_type)
    if "TEXT_BOX" not in st:
        bp = tf._txBody.find(qn("a:bodyPr"))
        g = lambda k, d: (float(bp.get(k)) / EMU_IN if bp is not None and bp.get(k) else d)
        ml = ml or g("lIns", 0.10)
        mr = mr or g("rIns", 0.10)
        mt = mt or g("tIns", 0.05)
        mb = mb or g("bIns", 0.05)
    iw = max(0.2, w - ml - mr)
    paras = []
    total = 0.0
    for p in tf.paragraphs:
        txt = "".join(r.text for r in p.runs)
        if not txt.strip():
            continue
        r0 = p.runs[0]
        sz = r0.font.size.pt if r0.font.size else 18.0
        bold = bool(r0.font.bold)
        ls = p.line_spacing if isinstance(p.line_spacing, float) and p.line_spacing > 0 else 1.2
        lines = wrap(txt, iw, sz, bold)
        hh = len(lines) * sz * ls / 72.0
        sb, sa = sp(p, "a:spcBef"), sp(p, "a:spcAft")
        paras.append((lines, sz, bold, ls, hh, sb, sa))
        total += hh + sb + sa
    va = str(tf.vertical_anchor)
    yy = y + mt
    if "MIDDLE" in va:
        yy = y + mt + (h - mt - mb - total) / 2.0
    elif "BOTTOM" in va:
        yy = y + h - mb - total
    out = []
    for lines, sz, bold, ls, hh, sb, sa in paras:
        yy += sb
        for ln in lines:
            out.append((x + ml, yy, tw(ln, sz, bold), sz * ls / 72.0, ln))
            yy += sz * ls / 72.0
        yy += sa
    return out, x, x + w


def run(path, foot_first=6.68):
    prs = Presentation(path)
    W, H = E(prs.slide_width), E(prs.slide_height)
    probs = []
    for si, slide in enumerate(prs.slides, 1):
        cards = []
        for zi, sh in enumerate(slide.shapes):
            x, y, w, h = E(sh.left), E(sh.top), E(sh.width), E(sh.height)
            if w >= W - 0.15 or h >= H - 0.15:
                continue
            try:
                st = str(sh.shape_type)
            except Exception:
                st = ""
            if "TEXT_BOX" in st or w < 0.55 or h < 0.30:
                continue
            cards.append((x, y, w, h, zi))
        for sh in slide.shapes:
            if not sh.has_text_frame or not any(p.runs for p in sh.text_frame.paragraphs):
                continue
            try:
                st = str(sh.shape_type)
            except Exception:
                st = ""
            is_box = "TEXT_BOX" in st
            rects, bx0, bx1 = line_rects(sh)
            by0 = min([r[1] for r in rects]) if rects else 0.0
            sx, sy = E(sh.left), E(sh.top)
            sw, shh = E(sh.width), E(sh.height)
            for (x, y, lw, lh, txt) in rects:
                if y + lh > H - 0.04:
                    probs.append("S%02d 越出页面底 %.2f “%s”" % (si, y + lh - H, txt[:14]))
                if not is_box:
                    if y + lh > sy + shh + 0.04:
                        probs.append("S%02d 形状内纵溢 %.2f “%s”" % (si, y + lh - sy - shh, txt[:14]))
                    if x + lw > sx + sw + 0.03:
                        probs.append("S%02d 形状内横溢 %.2f “%s”" % (si, x + lw - sx - sw, txt[:14]))
                host = None
                for (cx, cy, cw, chh, cz) in cards:
                    if y >= foot_first:
                        break
                    if (bx0 >= cx - 0.06 and bx1 <= cx + cw + 0.10
                            and cy - 0.06 <= by0 <= cy + chh - 0.02):
                        if host is None or cw * chh < host[3] * host[4]:
                            host = (cx, cy, cw, chh, cz)
                if host:
                    cx, cy, cw, chh, cz = host
                    if y + lh > cy + chh + 0.02:
                        probs.append("S%02d 卡片纵溢 %.2f @y%.2f “%s”" % (si, y + lh - cy - chh, y, txt[:14]))
                    if x + lw > cx + cw + 0.03:
                        probs.append("S%02d 卡片横溢 %.2f “%s”" % (si, x + lw - cx - cw, txt[:14]))
    probs = list(dict.fromkeys(probs))
    if not probs:
        print("OK — 未发现文字越界/溢出/叠压（slides=%d, %.3f×%.3fin）"
              % (len(prs.slides._sldIdLst), W, H))
    else:
        print("PROBLEMS (%d):" % len(probs))
        for p in probs[:40]:
            print("  " + p)
    return probs


def overlap_scan(path):
    """补充检查：文本行之间（不同形状）是否明显上下叠压"""
    prs = Presentation(path)
    W, H = E(prs.slide_width), E(prs.slide_height)
    bad = []
    for si, slide in enumerate(prs.slides, 1):
        items = []
        for sh in slide.shapes:
            if not sh.has_text_frame or not any(p.runs for p in sh.text_frame.paragraphs):
                continue
            rects, _, _ = line_rects(sh)
            for r in rects:
                items.append((r, id(sh)))
        cardlist = []
        for sh in slide.shapes:
            try:
                if "TEXT_BOX" in str(sh.shape_type):
                    continue
            except Exception:
                continue
            cx, cy, cw, chh = E(sh.left), E(sh.top), E(sh.width), E(sh.height)
            if cw * chh > 1.2 and cw < W - 0.15 and chh < H - 0.15:
                cardlist.append((cx, cy, cw, chh))
        def host_of(x, y, lw):
            best = None
            for (cx, cy, cw, chh) in cardlist:
                if x >= cx - 0.03 and x + lw <= cx + cw + 0.30 and cy - 0.06 <= y <= cy + chh + 0.06:
                    if best is None or cw * chh < best:
                        best = cw * chh
            return best
        items.sort(key=lambda z: z[0][1])
        n = len(items)
        for i in range(n):
            (x1, y1, w1, h1, t1), s1 = items[i]
            for j in range(i + 1, n):
                (x2, y2, w2, h2, t2), s2 = items[j]
                if y2 > y1 + max(h1, h2):
                    break
                if s1 == s2 or host_of(x1, y1, w1) != host_of(x2, y2, w2):
                    continue
                ox = min(x1 + w1, x2 + w2) - max(x1, x2)
                oy = min(y1 + h1, y2 + h2) - max(y1, y2)
                if ox > 0.15 and oy > min(h1, h2) * 0.55:
                    bad.append("S%02d 文字叠压: “%s” × “%s”" % (si, t1[:12], t2[:12]))
    return list(dict.fromkeys(bad))


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "第1组_问卷调查法与访谈法_v1.pptx"
    probs = run(src)
    ov = overlap_scan(src)
    if ov:
        print("OVERLAPS (%d):" % len(ov))
        for o in ov[:25]:
            print("  " + o)
