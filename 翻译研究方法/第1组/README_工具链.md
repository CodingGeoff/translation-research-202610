# 问卷调查法正反例 PPT 工具链 · 技术文档

一套「内容与生成分离」的 PPT 工具：所有文字放在一个 JSON 里，脚本读 JSON 生成 PPT。改 JSON 即可重新生成，也可交给别的 AI 润色 JSON 后再导入。

---

## 一、文件清单

| 文件 | 作用 |
|------|------|
| `deck_content_v5.json` | **唯一内容源**：27 页的全部文字（标题、要点、图片路径等） |
| `build_v5.py` | 读 JSON → 生成白底黑字版（`问卷调查法正反例_v5.pptx`） |
| `build_v6.py` | 读 JSON → 生成学院模板版（`问卷调查法正反例_v6.pptx`，带 logo + 底部格言） |
| `app.py` | Flask 可视化编辑器：网页改字 → 保存 → 生成 |
| `test_deck.py` | 稳健性测试：校验 JSON + 跑通两个 build |
| `assets/` | 论文截图（`zheng_*.png` 正例、`fan_*.png` 反例） |

## 二、JSON 格式

顶层结构：

```json
{
  "meta": { "title": "...", "note": "..." },
  "slides": [
    { "page": 1, "type": "cover", "title": "...", "subtitle": "...", "lines": ["...", "..."] },
    { "page": 2, "type": "toc", "title": "...", "items": ["...", "..."] },
    { "page": 3, "type": "section", "number": "01", "title": "..." },
    { "page": 4, "type": "slide", "title": "...", "bullets": ["...", "..."], "image": "zheng_p01_title.png", "caption": "..." },
    { "page": 27, "type": "end", "title": "...", "lines": ["..."] }
  ]
}
```

字段说明：

- **结构字段（不要改）**：`page`（页码）、`type`（cover/toc/section/slide/end）、`image`（图片文件名）、`number`（章节号）
- **文字字段（可润色）**：`title`、`subtitle`、`subtitle_en`、`caption`、`bullets`（要点列表）、`items`、`lines`

## 三、命令行用法

```bash
# 生成两个版本（用默认 deck_content_v5.json）
python build_v5.py
python build_v6.py

# 用润色后的 JSON 重新生成
python build_v5.py 润色后.json
python build_v6.py 润色后.json
```

## 四、Flask 可视化

```bash
python app.py
# 浏览器打开 http://127.0.0.1:5000
```

功能：网页上直接编辑每页标题/要点 → 「保存到 JSON」→ 「生成 V5 / V6」下载；右上角可「导出 JSON」「导入 JSON」。

## 五、交给 AI 润色的流程（核心闭环）

1. **导出**：把 `deck_content_v5.json` 发给另一个 AI（或点网页右上角「导出 JSON」）。
2. **润色提示**（可直接复制给 AI）：
   > 请润色这个 JSON 里每个 slide 的文字（title / subtitle / bullets / lines / items / caption），
   > 要求中文更自然、更口语化。只改文字内容，不要改 type、page、number、image 等字段，
   > 不要增删 slide，保持 JSON 结构完全一致，直接返回润色后的 JSON。
3. **导入**：把润色后的 JSON 保存为文件，用
   `python build_v5.py 润色后.json`（或 `build_v6.py`）重新生成 PPT；
   也可以在网页右上角「导入 JSON」上传。

## 六、测试

```bash
python test_deck.py
```

会校验：JSON 字段完整、页码连续、图片文件存在、两个 build 能跑通且页数与 JSON 一致。全部通过输出「0 失败」。

## 七、两个版本的区别

| | V5 | V6 |
|---|---|---|
| 基底 | 新建空白 | 复制学院模板 |
| logo / 格言 | 无 | 有（左上 logo + 底部格言） |
| 版式安全区 | 全页 | 避开左上 logo（y<1.05）与底部格言（y>6.57） |
| 字体 | Times New Roman + 宋体/黑体 | 同 V5 |

## 八、依赖

- `python-pptx`、`Pillow`（生成）
- `flask`（可视化编辑器，可选）
- 图片需放在 `assets/`，与 JSON 中 `image` 字段一致
