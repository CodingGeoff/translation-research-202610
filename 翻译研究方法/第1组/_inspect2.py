# -*- coding: utf-8 -*-
import zipfile, re, io, sys
from PIL import Image
from collections import Counter

TPL = 'template/高级翻译学院专用PPT模板修改版.pptx'
z = zipfile.ZipFile(TPL)

print("### slideMaster1.xml ###")
s = z.read('ppt/slideMasters/slideMaster1.xml').decode('utf-8')
print("has blip/image:", 'blip' in s)
for m in re.finditer(r'r:embed="([^"]+)"', s):
    print("  embed:", m.group(1))
for m in re.finditer(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', s):
    print("  color #", m.group(1))
# find off/ext for shapes in master
for m in re.finditer(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(-?\d+)" cy="(-?\d+)"', s):
    print("  shape off=(%s,%s) ext=(%s,%s) -> x=%.2f y=%.2f w=%.2f h=%.2f in" % (
        m.group(1), m.group(2), m.group(3), m.group(4),
        int(m.group(1))/914400, int(m.group(2))/914400, int(m.group(3))/914400, int(m.group(4))/914400))

print()
print("### slideMaster1 rels ###")
try:
    r = z.read('ppt/slideMasters/_rels/slideMaster1.xml.rels').decode('utf-8')
    print(r)
except KeyError:
    print("none")

print()
print("### which layouts reference image / fill ###")
for i in range(1, 12):
    nm = 'ppt/slideLayouts/slideLayout%d.xml' % i
    try:
        ls = z.read(nm).decode('utf-8')
    except KeyError:
        continue
    hasimg = 'blip' in ls
    colors = re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', ls)
    print("layout%d: img=%s colors=%s" % (i, hasimg, colors[:6]))

print()
print("### slide rels (background image?) ###")
for i in range(1, 7):
    try:
        r = z.read('ppt/slides/_rels/slide%d.xml.rels' % i).decode('utf-8')
    except KeyError:
        continue
    rels = re.findall(r'Target="([^"]+)"', r)
    print("slide%d rels targets: %s" % (i, rels))
