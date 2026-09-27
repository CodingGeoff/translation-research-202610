# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation

TPL = r"d:\10_Workspace\translation-research-202610\翻译研究方法\template\高级翻译学院专用PPT模板修改版.pptx"
prs = Presentation(TPL)
print("slide size:", prs.slide_width / 914400, prs.slide_height / 914400)
print("num slide_layouts:", len(prs.slide_layouts))
for i, ly in enumerate(prs.slide_layouts):
    print("  layout[%d] name=%r" % (i, ly.name))
print("num masters:", len(prs.slide_masters))
# 测试字体加载
import deck_lib as D
D.F_REG = r"C:\Windows\Fonts\msyh.ttc"
D.F_BOLD = r"C:\Windows\Fonts\msyhbd.ttc"
D.IDX = 0
try:
    print("tw test 思源黑体 12pt:", D.tw("思源黑体测试", 12, False))
    print("wrap test:", D.wrap("这是一段比较长的中文测试文本用来验证折行功能是否正常", 2.0, 12))
except Exception as e:
    print("font error:", repr(e))
