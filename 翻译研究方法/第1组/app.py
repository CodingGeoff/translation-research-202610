# -*- coding: utf-8 -*-
"""
PPT 内容可视化编辑器（Flask）
- 网页上编辑每页文字（标题/要点等）
- 保存回 deck_content_v5.json
- 一键生成 V5 / V6 PPT，下载
- 导出 / 导入 JSON

运行：python app.py  然后浏览器打开 http://127.0.0.1:5000
"""
import os, sys, json, subprocess, io
from flask import Flask, request, render_template_string, send_file, redirect, url_for, flash

BASE = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(BASE, "deck_content_v5.json")

app = Flask(__name__)
app.secret_key = "deck-editor"

# 可编辑的字符串字段 / 列表字段（其余视为结构字段，不编辑）
STR_FIELDS = {"title", "subtitle", "subtitle_en", "caption", "number"}
LIST_FIELDS = {"bullets", "lines", "items"}


def load_deck():
    with open(JSON_PATH, encoding="utf-8") as f:
        return json.load(f)


def save_deck(data):
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def build_pptx(script, outname):
    r = subprocess.run([sys.executable, os.path.join(BASE, script)],
                       cwd=BASE, capture_output=True, text=True, encoding="utf-8", errors="replace")
    out = os.path.join(BASE, outname)
    if r.returncode != 0 or not os.path.exists(out):
        return False, (r.stdout or "") + (r.stderr or "")
    return True, ""


HTML = """
<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>PPT 内容编辑器</title>
<style>
body{font-family:'Segoe UI',sans-serif;margin:0;background:#f5f6f8;color:#222;}
.top{position:sticky;top:0;background:#1f2937;color:#fff;padding:14px 24px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;z-index:10;}
.top h1{font-size:18px;margin:0;margin-right:auto;}
button,.btn{background:#2563eb;color:#fff;border:none;padding:8px 14px;border-radius:6px;cursor:pointer;font-size:14px;}
.btn.green{background:#059669}.btn.gray{background:#6b7280}
.slide{background:#fff;margin:16px 24px;padding:16px 20px;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);}
.slide .tag{display:inline-block;background:#eef2ff;color:#4338ca;padding:2px 8px;border-radius:4px;font-size:12px;margin-right:8px;}
.slide label{display:block;font-size:12px;color:#666;margin:10px 0 4px;}
.slide input[type=text],.slide textarea{width:100%;box-sizing:border-box;padding:8px;border:1px solid #d1d5db;border-radius:6px;font-size:14px;font-family:inherit;}
.slide textarea{min-height:90px;resize:vertical;line-height:1.5;}
.slide .readonly{background:#f9fafb;color:#999;border-style:dashed;}
.msg{padding:10px 24px;}
.msg.ok{color:#059669}.msg.err{color:#dc2626}
</style>
</head>
<body>
<div class="top">
  <h1>PPT 内容编辑器（27 页）</h1>
  <button onclick="document.getElementById('form').submit();document.getElementById('action').value='save'">保存到 JSON</button>
  <button class="green" onclick="sub('v5')">生成 V5</button>
  <button class="green" onclick="sub('v6')">生成 V6</button>
  <a class="btn gray" href="/export" style="text-decoration:none">导出 JSON</a>
  <form method="post" action="/import" enctype="multipart/form-data" style="display:inline">
    <input type="file" name="file" accept=".json" onchange="this.form.submit()" style="color:#fff">
  </form>
</div>
{% if msg %}<div class="msg {{msgtype}}">{{msg}}</div>{% endif %}
<form id="form" method="post" action="/save">
<input type="hidden" name="action" id="action" value="save">
{% for s in slides %}
<div class="slide">
  <span class="tag">第 {{s.page}} 页</span><span class="tag">{{s.type}}</span>
  {% if s.image %}<span class="tag">图：{{s.image}}</span>{% endif %}
  {% for k,v in s.items() %}
    {% if k in ('title','subtitle','subtitle_en','caption','number') %}
      <label>{{k}}</label>
      <input type="text" name="p{{s.page}}__{{k}}" value="{{v}}">
    {% elif k in ('bullets','lines','items') %}
      <label>{{k}}（每行一条）</label>
      <textarea name="p{{s.page}}__{{k}}">{{v|join('\n')}}</textarea>
    {% endif %}
  {% endfor %}
</div>
{% endfor %}
</form>
<script>
function sub(v){
  document.getElementById('action').value = 'build_'+v;
  document.getElementById('form').submit();
}
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML, slides=load_deck()["slides"], msg=None, msgtype="")


@app.route("/save", methods=["POST"])
def save():
    action = request.form.get("action", "save")
    data = load_deck()
    # 1) 收集表单修改并写回 JSON
    for s in data["slides"]:
        page = s["page"]
        for k in list(s.keys()):
            key = "p%d__%s" % (page, k)
            if k in STR_FIELDS and key in request.form:
                s[k] = request.form[key]
            elif k in LIST_FIELDS and key in request.form:
                lines = [ln.rstrip("\r") for ln in request.form[key].split("\n")]
                lines = [ln for ln in lines if ln.strip() != ""]
                s[k] = lines
    save_deck(data)
    # 2) 处理构建请求
    if action == "build_v5":
        ok, err = build_pptx("build_v5.py", "问卷调查法正反例_v5.pptx")
        if ok:
            return redirect("/download/问卷调查法正反例_v5.pptx")
        return render_template_string(HTML, slides=data["slides"], msg="生成失败：" + err, msgtype="err")
    if action == "build_v6":
        ok, err = build_pptx("build_v6.py", "问卷调查法正反例_v6.pptx")
        if ok:
            return redirect("/download/问卷调查法正反例_v6.pptx")
        return render_template_string(HTML, slides=data["slides"], msg="生成失败：" + err, msgtype="err")
    return render_template_string(HTML, slides=data["slides"], msg="已保存到 deck_content_v5.json", msgtype="ok")


@app.route("/download/<path:name>")
def download(name):
    return send_file(os.path.join(BASE, name), as_attachment=True)


@app.route("/export")
def export():
    return send_file(JSON_PATH, as_attachment=True, download_name="deck_content.json")


@app.route("/import", methods=["POST"])
def import_json():
    f = request.files.get("file")
    if not f:
        return redirect("/")
    data = json.load(io.BytesIO(f.read()))
    if "slides" not in data:
        return render_template_string(HTML, slides=load_deck()["slides"], msg="导入失败：缺少 slides 字段", msgtype="err")
    save_deck(data)
    return render_template_string(HTML, slides=data["slides"], msg="导入成功，共 %d 页" % len(data["slides"]), msgtype="ok")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
