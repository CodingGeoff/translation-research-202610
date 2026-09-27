# -*- coding: utf-8 -*-
import zipfile, re
z = zipfile.ZipFile('template/高级翻译学院专用PPT模板修改版.pptx')

print("### slideLayout1.xml full ###")
s = z.read('ppt/slideLayouts/slideLayout1.xml').decode('utf-8')
print(s)
