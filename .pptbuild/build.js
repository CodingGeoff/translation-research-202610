const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33" x 7.5"
pres.author = "第1小组（问卷调查法专项）";
pres.title = "问卷调查法在翻译研究中的正反例分析";

// ---------- 配色 ----------
const C = {
  navy: "1E2A5E",       // 主色 深藏青
  navyDark: "141B3E",   // 标题/结尾深底
  navy2: "31417E",      // 次级蓝
  ice: "DCE4F5",        // 冰蓝
  gold: "DFA53B",       // 强调金
  goldLight: "F6E9CD",
  green: "2E9E6E",      // 正向
  greenLight: "E4F4ED",
  red: "C9463D",        // 负向
  redLight: "FBE9E7",
  ink: "20263A",        // 正文
  muted: "6A7288",      // 次要
  paper: "FFFFFF",      // 内容底
  card: "F4F6FB",       // 卡片底
  line: "E4E8F2",
  link: "1F5FB5",
};
const FONT = "Microsoft YaHei";
const W = 13.33, H = 7.5;
const ML = 0.65; // 左边距

// ---------- 通用 ----------
function pageNum(slide, n) {
  slide.addText(String(n), { x: W - 1.0, y: H - 0.5, w: 0.5, h: 0.3, fontSize: 11, color: C.muted, align: "right", fontFace: FONT });
}
function footer(slide, label) {
  slide.addText(label, { x: ML, y: H - 0.5, w: 8, h: 0.3, fontSize: 9, color: C.muted, fontFace: FONT });
}
function header(slide, tag, title, num) {
  slide.background = { color: C.paper };
  slide.addShape(pres.shapes.RECTANGLE, { x: ML, y: 0.56, w: 0.15, h: 0.56, fill: { color: C.gold } });
  slide.addText(tag, { x: ML + 0.3, y: 0.5, w: 11, h: 0.3, fontSize: 12, color: C.gold, bold: true, fontFace: FONT });
  slide.addText(title, { x: ML + 0.28, y: 0.82, w: 11.9, h: 0.75, fontSize: 28, color: C.navy, bold: true, fontFace: FONT });
  pageNum(slide, num);
}
function card(slide, x, y, w, h, fill, line) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: fill || C.card }, line: { color: line || C.line, width: 1 }, rectRadius: 0.06,
  });
}

// 带符号的条目（自定义颜色符号 + 正文）
function symRow(slide, x, y, w, h, symbol, sColor, text, opts) {
  const o = opts || {};
  slide.addText(
    [
      { text: symbol + "  ", options: { color: sColor, bold: true, fontSize: o.fontSize || 14 } },
      { text: text, options: { color: o.color || C.ink, fontSize: o.fontSize || 14, bold: o.bold } },
    ],
    { x, y, w, h, fontFace: FONT, valign: "top", align: o.align || "left", lineSpacingMultiple: o.lsm || 1.0 }
  );
}
function labelChip(slide, x, y, w, text, bg, fg) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.34, fill: { color: bg }, rectRadius: 0.17, line: { type: "none" } });
  slide.addText(text, { x, y, w, h: 0.34, fontSize: 11, color: fg, bold: true, align: "center", valign: "middle", fontFace: FONT });
}

// =========================================================
// 第 1 页：标题页
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.navyDark };
  // 装饰：右侧大圆
  s.addShape(pres.shapes.OVAL, { x: 9.1, y: -1.4, w: 6.2, h: 6.2, fill: { color: C.navy, transparency: 35 } });
  s.addShape(pres.shapes.OVAL, { x: 10.2, y: -0.6, w: 3.4, h: 3.4, fill: { color: C.navy2, transparency: 30 } });
  s.addShape(pres.shapes.OVAL, { x: 10.75, y: 2.6, w: 0.9, h: 0.9, fill: { color: C.gold, transparency: 20 } });

  // 金色竖条 + 小组标签
  s.addShape(pres.shapes.RECTANGLE, { x: ML, y: 1.35, w: 0.18, h: 1.7, fill: { color: C.gold } });
  s.addText("翻译研究方法 · 课程小组展示", { x: ML + 0.42, y: 1.42, w: 8, h: 0.34, fontSize: 14, color: C.gold, bold: true, fontFace: FONT });

  s.addText("问卷调查法在翻译研究中的\n正反例分析", {
    x: ML + 0.4, y: 1.9, w: 8.4, h: 1.6, fontSize: 38, color: "FFFFFF", bold: true, fontFace: FONT, lineSpacingMultiple: 1.15,
  });

  s.addText("问卷调查法专项 · 第1小组", { x: ML + 0.42, y: 3.72, w: 8, h: 0.4, fontSize: 18, color: C.ice, fontFace: FONT });
  s.addText("2026 年 10 月 12 日", { x: ML + 0.42, y: 4.2, w: 8, h: 0.4, fontSize: 14, color: "9AA3C4", fontFace: FONT });

  // 底部提示条
  s.addShape(pres.shapes.RECTANGLE, { x: ML, y: 6.35, w: 0.05, h: 0.0, fill: { color: C.gold } });
  s.addText("基于真实翻译研究案例的正反例对比分析", { x: ML, y: 6.2, w: 8, h: 0.35, fontSize: 11, color: "8B94B8", fontFace: FONT });
}

// =========================================================
// 第 2 页：目录
// =========================================================
{
  const s = pres.addSlide();
  header(s, "CONTENTS", "目录", 2);

  const items = [
    { n: "01", t: "问卷调查法的核心特征", d: "定义 · 适用场景 · 优势与局限" },
    { n: "02", t: "正例分析", d: "案例1：汉语翻译语体识别 · 案例2：翻译技术培训" },
    { n: "03", t: "反例分析", d: "问题1：问卷设计缺陷 · 问题2：数据分析不规范" },
    { n: "04", t: "正反对比与方法总结", d: "七维度对比 · 关键启示" },
    { n: "05", t: "思考题", d: "课堂讨论与互动" },
    { n: "06", t: "参考文献", d: "中英文文献 · 高亮链接" },
  ];
  const colW = 5.85, colH = 1.42, gapX = 0.3, gapY = 0.24;
  const x0 = ML, y0 = 1.95;
  items.forEach((it, i) => {
    const cx = x0 + (i % 2) * (colW + gapX);
    const cy = y0 + Math.floor(i / 2) * (colH + gapY);
    card(s, cx, cy, colW, colH, C.card);
    // 编号
    s.addText(it.n, { x: cx + 0.25, y: cy + 0.28, w: 0.9, h: 0.6, fontSize: 26, color: C.gold, bold: true, fontFace: FONT, valign: "middle" });
    // 竖线
    s.addShape(pres.shapes.RECTANGLE, { x: cx + 1.15, y: cy + 0.26, w: 0.03, h: colH - 0.52, fill: { color: C.line } });
    // 标题与说明
    s.addText(it.t, { x: cx + 1.35, y: cy + 0.2, w: colW - 1.6, h: 0.5, fontSize: 17, color: C.navy, bold: true, fontFace: FONT });
    s.addText(it.d, { x: cx + 1.35, y: cy + 0.72, w: colW - 1.6, h: 0.5, fontSize: 11.5, color: C.muted, fontFace: FONT });
  });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 3 页：核心特征
// =========================================================
{
  const s = pres.addSlide();
  header(s, "方法概述", "问卷调查法在翻译研究中的核心特征", 3);

  // 定义条
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ML, y: 1.8, w: 12.03, h: 0.78, fill: { color: C.navy }, rectRadius: 0.06 });
  s.addText([
    { text: "定义   ", options: { color: C.gold, bold: true, fontSize: 14 } },
    { text: "以标准化问卷为工具，系统收集翻译相关数据的", options: { color: "FFFFFF", fontSize: 14 } },
    { text: "定量研究方法", options: { color: C.gold, bold: true, fontSize: 14 } },
    { text: "。", options: { color: "FFFFFF", fontSize: 14 } },
  ], { x: ML + 0.3, y: 1.8, w: 11.5, h: 0.78, fontFace: FONT, valign: "middle" });

  // 适用场景（左）与 优势/局限（右）
  // 左侧：适用场景 2x2
  s.addText("适用场景", { x: ML, y: 2.85, w: 4, h: 0.4, fontSize: 16, color: C.navy, bold: true, fontFace: FONT });
  const scenes = ["读者对译文的反应 / 评价", "译者行为 / 策略调查", "翻译教学效果评估", "翻译行业现状分析"];
  const sc = { x: ML, y: 3.3, w: 2.85, h: 0.52, gapX: 0.15, gapY: 0.18 };
  scenes.forEach((t, i) => {
    const cx = sc.x + (i % 2) * (sc.w + sc.gapX);
    const cy = sc.y + Math.floor(i / 2) * (sc.h + sc.gapY);
    card(s, cx, cy, sc.w, sc.h, C.card);
    symRow(s, cx + 0.14, cy + 0.14, sc.w - 0.28, 0.3, "✓", C.green, t, { fontSize: 12 });
  });

  // 右侧：优势 + 局限
  const rx = 7.05, rw = 5.63;
  s.addText("优势", { x: rx, y: 2.85, w: 3, h: 0.4, fontSize: 16, color: C.green, bold: true, fontFace: FONT });
  card(s, rx, 3.3, rw, 1.45, C.greenLight, C.green);
  const adv = ["大规模数据收集", "标准化操作", "易于定量分析"];
  adv.forEach((t, i) => {
    symRow(s, rx + 0.2, 3.46 + i * 0.42, rw - 0.4, 0.34, "✓", C.green, t, { fontSize: 13 });
  });

  s.addText("局限性", { x: rx, y: 4.95, w: 3, h: 0.4, fontSize: 16, color: C.red, bold: true, fontFace: FONT });
  card(s, rx, 5.4, rw, 1.45, C.redLight, C.red);
  const lim = ["回收率难以保证", "缺乏深度定性分析", "受问卷设计质量影响大"];
  lim.forEach((t, i) => {
    symRow(s, rx + 0.2, 5.56 + i * 0.42, rw - 0.4, 0.34, "✗", C.red, t, { fontSize: 13 });
  });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 4 页：正例分析概述
// =========================================================
{
  const s = pres.addSlide();
  header(s, "正例分析", "科学规范的问卷调查", 4);

  // 左：核心特征
  const lx = ML, lw = 5.9;
  card(s, lx, 2.0, lw, 3.6, C.card);
  s.addText("核心特征", { x: lx + 0.3, y: 2.2, w: 4, h: 0.4, fontSize: 16, color: C.navy, bold: true, fontFace: FONT });
  const feats = ["明确的研究问题", "严谨的问卷设计", "规范的调查流程", "科学的数据分析"];
  feats.forEach((t, i) => {
    const cy = 2.75 + i * 0.68;
    s.addShape(pres.shapes.OVAL, { x: lx + 0.3, y: cy + 0.02, w: 0.4, h: 0.4, fill: { color: C.green } });
    s.addText(String(i + 1), { x: lx + 0.3, y: cy + 0.02, w: 0.4, h: 0.4, fontSize: 14, color: "FFFFFF", bold: true, align: "center", valign: "middle", fontFace: FONT });
    s.addText(t, { x: lx + 0.9, y: cy + 0.04, w: lw - 1.2, h: 0.4, fontSize: 15, color: C.ink, bold: true, fontFace: FONT });
  });

  // 右：选取标准
  const rx = 7.05, rw = 5.63;
  card(s, rx, 2.0, rw, 3.6, C.card);
  s.addText("案例选取标准", { x: rx + 0.3, y: 2.2, w: 4, h: 0.4, fontSize: 16, color: C.navy, bold: true, fontFace: FONT });
  const crit = ["发表于核心期刊", "问卷设计逻辑清晰", "数据分析方法得当", "结论可靠可信"];
  crit.forEach((t, i) => {
    const cy = 2.75 + i * 0.68;
    symRow(s, rx + 0.3, cy + 0.04, rw - 0.6, 0.4, "✓", C.gold, t, { fontSize: 15, bold: true });
  });

  // 底部强调条
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ML, y: 5.95, w: 12.03, h: 0.72, fill: { color: C.navy }, rectRadius: 0.06 });
  s.addText([
    { text: "一句话：", options: { color: C.gold, bold: true, fontSize: 13.5 } },
    { text: "好的问卷调查 = 好问题 + 好设计 + 好流程 + 好分析", options: { color: "FFFFFF", fontSize: 13.5 } },
  ], { x: ML + 0.3, y: 5.95, w: 11.5, h: 0.72, fontFace: FONT, valign: "middle" });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 5 页：正例 1 汉语翻译语体识别研究
// =========================================================
{
  const s = pres.addSlide();
  header(s, "正例 1", "汉语翻译语体识别研究", 5);

  // 来源条
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ML, y: 1.82, w: 12.03, h: 0.66, fill: { color: C.ice }, rectRadius: 0.05 });
  s.addText([
    { text: "案例来源：", options: { bold: true, color: C.navy, fontSize: 12 } },
    { text: "郑剑委《翻译出版生产中的译者行为体系研究》相关研究 · 《外语研究》类核心期刊", options: { color: C.ink, fontSize: 12 } },
  ], { x: ML + 0.25, y: 1.82, w: 11.6, h: 0.66, fontFace: FONT, valign: "middle" });

  // 左列
  const lx = ML, lw = 5.9;
  card(s, lx, 2.62, lw, 2.6, C.card);
  s.addText("研究问题", { x: lx + 0.25, y: 2.74, w: 4, h: 0.32, fontSize: 14, color: C.navy, bold: true, fontFace: FONT });
  s.addText("汉语读者能否识别汉语翻译语体？\n其影响因素是什么？", {
    x: lx + 0.25, y: 3.08, w: lw - 0.5, h: 0.8, fontSize: 13.5, color: C.ink, italic: true, fontFace: FONT, lineSpacingMultiple: 1.15,
  });
  s.addText("数据分析", { x: lx + 0.25, y: 3.94, w: 4, h: 0.32, fontSize: 14, color: C.navy, bold: true, fontFace: FONT });
  s.addText([
    { text: "统计方法：", options: { bold: true, color: C.ink, fontSize: 12.5 } },
    { text: "卡方检验、相关性分析", options: { color: C.ink, fontSize: 12.5 } },
  ], { x: lx + 0.25, y: 4.26, w: lw - 0.5, h: 0.28, fontFace: FONT });
  symRow(s, lx + 0.25, 4.54, lw - 0.5, 0.28, "▸", C.gold, "总识别率与英语水平成正比（r = 0.67, p < 0.01）", { fontSize: 12.5 });
  symRow(s, lx + 0.25, 4.82, lw - 0.5, 0.28, "▸", C.gold, "读者类型与识别能力无显著相关性", { fontSize: 12.5 });

  // 右列
  const rx = 7.05, rw = 5.63;
  card(s, rx, 2.62, rw, 2.6, C.card);
  s.addText("问卷设计亮点", { x: rx + 0.25, y: 2.74, w: 4, h: 0.32, fontSize: 14, color: C.navy, bold: true, fontFace: FONT });
  const hl1 = ["双盲原则：被试不知文本是原创还是翻译", "样本代表性：300 名不同类型读者", "问题中性：避免诱导性表述", "混合题型：判断题 + 理由选择题"];
  hl1.forEach((t, i) => {
    symRow(s, rx + 0.25, 3.12 + i * 0.44, rw - 0.5, 0.38, "✓", C.green, t, { fontSize: 12 });
  });

  // 底部：值得借鉴
  card(s, ML, 5.42, 12.03, 1.4, C.greenLight, C.green);
  s.addText("值得借鉴的方面", { x: ML + 0.25, y: 5.54, w: 4, h: 0.34, fontSize: 14, color: C.green, bold: true, fontFace: FONT });
  const borrow = [
    { t: "预调查：", d: "小规模测试后修正问卷" },
    { t: "信度检验：", d: "Cronbach's α = 0.85（优秀）" },
    { t: "效度检验：", d: "专家评审 + 因子分析" },
    { t: "透明化：", d: "详细报告调查过程" },
  ];
  borrow.forEach((b, i) => {
    const cx = ML + 0.25 + (i % 2) * 5.95;
    const cy = 5.94 + Math.floor(i / 2) * 0.42;
    s.addText([
      { text: b.t, options: { bold: true, color: C.green, fontSize: 12.5 } },
      { text: b.d, options: { color: C.ink, fontSize: 12.5 } },
    ], { x: cx, y: cy, w: 5.85, h: 0.32, fontFace: FONT });
  });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 6 页：正例 1 问卷设计示例
// =========================================================
{
  const s = pres.addSlide();
  header(s, "正例 1 · 问卷示例", "问卷设计截图示例", 6);

  // 左侧：模拟问卷卡片
  const qx = ML, qw = 7.1;
  card(s, qx, 1.9, qw, 4.7, "FFFFFF", C.line);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: qx, y: 1.9, w: qw, h: 0.56, fill: { color: C.navy }, rectRadius: 0.06 });
  s.addText("示例问题（中性 · 无诱导）", { x: qx + 0.25, y: 1.9, w: 6, h: 0.56, fontSize: 13, color: "FFFFFF", bold: true, fontFace: FONT, valign: "middle" });

  s.addText("请判断以下文本是原创文本还是翻译文本：", { x: qx + 0.3, y: 2.68, w: 6.4, h: 0.35, fontSize: 14, color: C.ink, fontFace: FONT });
  // 文本示例框
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: qx + 0.3, y: 3.08, w: 6.4, h: 0.9, fill: { color: C.card }, rectRadius: 0.05, line: { color: C.line, width: 1 } });
  s.addText("[ 文本 A：汉语段落 ]", { x: qx + 0.5, y: 3.08, w: 6, h: 0.9, fontSize: 12, color: C.muted, fontFace: FONT, valign: "middle" });
  // 选项
  s.addText("选项：", { x: qx + 0.3, y: 4.1, w: 1, h: 0.34, fontSize: 13, color: C.ink, fontFace: FONT });
  const opts = ["原创", "翻译"];
  opts.forEach((t, i) => {
    const ox = qx + 1.3 + i * 1.5;
    s.addShape(pres.shapes.RECTANGLE, { x: ox, y: 4.14, w: 0.26, h: 0.26, fill: { color: "FFFFFF" }, line: { color: C.muted, width: 1.5 } });
    s.addText(t, { x: ox + 0.38, y: 4.08, w: 1.1, h: 0.34, fontSize: 13, color: C.ink, fontFace: FONT });
  });
  s.addText("您选择的依据是？（可多选）", { x: qx + 0.3, y: 4.6, w: 5, h: 0.34, fontSize: 13, color: C.ink, fontFace: FONT });
  const reasons = ["语言表达习惯", "句式结构", "词汇选择"];
  reasons.forEach((t, i) => {
    const rx = qx + 0.3 + i * 2.05;
    s.addShape(pres.shapes.RECTANGLE, { x: rx, y: 5.02, w: 0.26, h: 0.26, fill: { color: "FFFFFF" }, line: { color: C.muted, width: 1.5 } });
    s.addText(t, { x: rx + 0.38, y: 4.98, w: 1.75, h: 0.34, fontSize: 12, color: C.ink, fontFace: FONT });
  });
  s.addText("其他：____________", { x: qx + 0.3, y: 5.5, w: 5, h: 0.34, fontSize: 12, color: C.muted, fontFace: FONT });

  // 右侧：设计亮点
  const rx = 8.0, rw = 4.68;
  card(s, rx, 1.9, rw, 4.7, C.card);
  s.addText("设计亮点", { x: rx + 0.3, y: 2.1, w: 4, h: 0.4, fontSize: 16, color: C.navy, bold: true, fontFace: FONT });
  const points = [
    { t: "中性表述", d: "问题不预设答案，避免诱导", c: C.green },
    { t: "逻辑严密", d: "选项互斥且穷尽", c: C.green },
    { t: "开放性补充", d: "允许被试提供其他理由", c: C.gold },
  ];
  points.forEach((p, i) => {
    const cy = 2.7 + i * 1.05;
    s.addShape(pres.shapes.OVAL, { x: rx + 0.3, y: cy + 0.02, w: 0.42, h: 0.42, fill: { color: p.c } });
    s.addText(String(i + 1), { x: rx + 0.3, y: cy + 0.02, w: 0.42, h: 0.42, fontSize: 14, color: "FFFFFF", bold: true, align: "center", valign: "middle", fontFace: FONT });
    s.addText(p.t, { x: rx + 0.9, y: cy - 0.02, w: rw - 1.2, h: 0.34, fontSize: 14.5, color: C.ink, bold: true, fontFace: FONT });
    s.addText(p.d, { x: rx + 0.9, y: cy + 0.34, w: rw - 1.2, h: 0.34, fontSize: 11.5, color: C.muted, fontFace: FONT });
  });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 7 页：正例 2 翻译技术培训效果调研
// =========================================================
{
  const s = pres.addSlide();
  header(s, "正例 2", "翻译技术培训效果调研", 7);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ML, y: 1.82, w: 12.03, h: 0.66, fill: { color: C.ice }, rectRadius: 0.05 });
  s.addText([
    { text: "案例来源：", options: { bold: true, color: C.navy, fontSize: 12 } },
    { text: "Schaeffer et al. (2020) TICQ 问卷 · 肖维青、熊凌崧（2023）", options: { color: C.ink, fontSize: 12 } },
  ], { x: ML + 0.25, y: 1.82, w: 11.6, h: 0.66, fontFace: FONT, valign: "middle" });

  const lx = ML, lw = 5.9;
  card(s, lx, 2.62, lw, 2.6, C.card);
  s.addText("研究问题", { x: lx + 0.25, y: 2.74, w: 4, h: 0.32, fontSize: 14, color: C.navy, bold: true, fontFace: FONT });
  s.addText("翻译技术培训的效果如何？\n学生对培训的满意度与能力提升如何？", {
    x: lx + 0.25, y: 3.08, w: lw - 0.5, h: 0.8, fontSize: 13.5, color: C.ink, italic: true, fontFace: FONT, lineSpacingMultiple: 1.15,
  });
  s.addText("数据分析", { x: lx + 0.25, y: 3.94, w: 4, h: 0.32, fontSize: 14, color: C.navy, bold: true, fontFace: FONT });
  symRow(s, lx + 0.25, 4.26, lw - 0.5, 0.28, "▸", C.gold, "样本 385 名 MTI 学生，回收率 94%", { fontSize: 12.5 });
  symRow(s, lx + 0.25, 4.54, lw - 0.5, 0.28, "▸", C.gold, "描述性统计 + 方差分析", { fontSize: 12.5 });
  symRow(s, lx + 0.25, 4.82, lw - 0.5, 0.28, "▸", C.gold, "53.76% 认为教材系统覆盖翻译技术知识", { fontSize: 12.5 });

  const rx = 7.05, rw = 5.63;
  card(s, rx, 2.62, rw, 2.6, C.card);
  s.addText("问卷设计亮点", { x: rx + 0.25, y: 2.74, w: 4, h: 0.32, fontSize: 14, color: C.navy, bold: true, fontFace: FONT });
  const hl2 = ["理论框架：基于 Kirkpatrick 模型（反应层 + 行为层）", "维度全面：教材系统性 / 教学方法 / 评估方式", "量表设计：李克特 5 级量表"];
  hl2.forEach((t, i) => {
    symRow(s, rx + 0.25, 3.14 + i * 0.5, rw - 0.5, 0.44, "✓", C.green, t, { fontSize: 12 });
  });

  card(s, ML, 5.42, 12.03, 1.4, C.greenLight, C.green);
  s.addText("值得借鉴的方面", { x: ML + 0.25, y: 5.54, w: 4, h: 0.34, fontSize: 14, color: C.green, bold: true, fontFace: FONT });
  const borrow2 = [
    { t: "混合方法：", d: "问卷 + 访谈（8 名学生）" },
    { t: "分层分析：", d: "按年级 / 专业分组对比" },
    { t: "结果可视化：", d: "图表清晰展示趋势" },
  ];
  borrow2.forEach((b, i) => {
    const cx = ML + 0.25 + (i % 3) * 3.9;
    const cy = 5.96 + Math.floor(i / 3) * 0.42;
    s.addText([
      { text: b.t, options: { bold: true, color: C.green, fontSize: 12.5 } },
      { text: b.d, options: { color: C.ink, fontSize: 12.5 } },
    ], { x: cx, y: cy, w: 3.8, h: 0.32, fontFace: FONT });
  });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 8 页：反例分析概述
// =========================================================
{
  const s = pres.addSlide();
  header(s, "反例分析", "常见问题与典型案例", 8);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ML, y: 1.82, w: 12.03, h: 0.66, fill: { color: C.ice }, rectRadius: 0.05 });
  s.addText([
    { text: "数据来源：", options: { bold: true, color: C.navy, fontSize: 12 } },
    { text: "王水、赵建军（2024）· 基于 376 篇文科硕士论文的调查", options: { color: C.ink, fontSize: 12 } },
  ], { x: ML + 0.25, y: 1.82, w: 11.6, h: 0.66, fontFace: FONT, valign: "middle" });

  // 四个统计大数字
  const stats = [
    { num: "49.47%", lab: "问卷结构不完整", sub: "（186 篇）" },
    { num: "82.99%", lab: "缺乏预调查", sub: "（294 篇无说明）" },
    { num: "多数 < 200", lab: "样本量不足", sub: "代表性弱" },
    { num: "占比极高", lab: "缺乏信效度检验", sub: "内部一致性未知" },
  ];
  const sw = 2.86, sg = 0.2;
  stats.forEach((st, i) => {
    const cx = ML + i * (sw + sg);
    card(s, cx, 2.72, sw, 1.95, C.redLight, C.red);
    s.addText(st.num, { x: cx + 0.15, y: 2.95, w: sw - 0.3, h: 0.7, fontSize: 24, color: C.red, bold: true, fontFace: FONT, align: "center" });
    s.addText(st.lab, { x: cx + 0.15, y: 3.72, w: sw - 0.3, h: 0.4, fontSize: 13, color: C.ink, bold: true, fontFace: FONT, align: "center" });
    s.addText(st.sub, { x: cx + 0.15, y: 4.12, w: sw - 0.3, h: 0.3, fontSize: 10.5, color: C.muted, fontFace: FONT, align: "center" });
  });

  // 核心问题
  card(s, ML, 5.0, 12.03, 1.7, C.card);
  s.addText("核心问题", { x: ML + 0.3, y: 5.15, w: 4, h: 0.4, fontSize: 16, color: C.navy, bold: true, fontFace: FONT });
  const probs = ["问卷设计粗糙", "调查过程随意", "数据分析不规范"];
  probs.forEach((t, i) => {
    const cx = ML + 0.3 + i * 3.9;
    s.addShape(pres.shapes.OVAL, { x: cx, y: 5.7, w: 0.5, h: 0.5, fill: { color: C.red } });
    s.addText(String(i + 1), { x: cx, y: 5.7, w: 0.5, h: 0.5, fontSize: 16, color: "FFFFFF", bold: true, align: "center", valign: "middle", fontFace: FONT });
    s.addText(t, { x: cx + 0.7, y: 5.76, w: 3.2, h: 0.4, fontSize: 14.5, color: C.ink, bold: true, fontFace: FONT });
  });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 9 页：反例 1 问卷设计缺陷
// =========================================================
{
  const s = pres.addSlide();
  header(s, "反例 1", "问卷设计缺陷：诱导性提问", 9);

  // 左侧：案例 + 分析
  const lx = ML, lw = 6.4;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: lx, y: 1.9, w: lw, h: 0.5, fill: { color: C.redLight }, rectRadius: 0.05 });
  s.addText("语言学及应用语言学硕士论文（2019–2022）", { x: lx + 0.25, y: 1.9, w: 6, h: 0.5, fontSize: 12, color: C.red, bold: true, fontFace: FONT, valign: "middle" });

  // 案例引述框
  card(s, lx, 2.56, lw, 1.9, "FFFFFF", C.line);
  s.addText("具体案例（问题原文）", { x: lx + 0.25, y: 2.68, w: 5, h: 0.34, fontSize: 13, color: C.navy, bold: true, fontFace: FONT });
  s.addText("“有人认为西安话有一千多年的历史，\n在西安生活一定要使用西安方言。您认为呢？”", {
    x: lx + 0.25, y: 3.06, w: lw - 0.5, h: 0.85, fontSize: 13, color: C.red, italic: true, fontFace: FONT, lineSpacingMultiple: 1.25,
  });
  s.addText("选项：□ 同意　□ 不同意　□ 无所谓", { x: lx + 0.25, y: 3.98, w: 5.5, h: 0.34, fontSize: 12, color: C.muted, fontFace: FONT });

  card(s, lx, 4.62, lw, 2.2, C.redLight, C.red);
  s.addText("问题分析", { x: lx + 0.25, y: 4.74, w: 4, h: 0.34, fontSize: 14, color: C.red, bold: true, fontFace: FONT });
  const ana = ["隐含立场：问题本身暗示“西安话很重要”", "诱导回答：被试易受暗示，倾向选择“同意”", "缺乏中性：未提供平衡的背景信息"];
  ana.forEach((t, i) => {
    symRow(s, lx + 0.25, 5.14 + i * 0.46, lw - 0.5, 0.4, "✗", C.red, t, { fontSize: 12 });
  });

  // 右侧：后果 + 改进建议
  const rx = 7.35, rw = 5.33;
  card(s, rx, 1.9, rw, 2.2, C.card);
  s.addText("后果", { x: rx + 0.25, y: 2.02, w: 4, h: 0.34, fontSize: 14, color: C.navy, bold: true, fontFace: FONT });
  const cons = ["数据偏差：结果缺乏客观性", "信度受损：问卷整体可信度下降", "结论不可靠：研究发现可能被误导"];
  cons.forEach((t, i) => {
    symRow(s, rx + 0.25, 2.42 + i * 0.46, rw - 0.5, 0.4, "✗", C.red, t, { fontSize: 12 });
  });

  card(s, rx, 4.3, rw, 2.52, C.greenLight, C.green);
  s.addText("改进建议", { x: rx + 0.25, y: 4.42, w: 4, h: 0.34, fontSize: 14, color: C.green, bold: true, fontFace: FONT });
  const fix = ["中性表述：“您认为在西安生活是否需要使用西安方言？”", "平衡选项：提供“强烈同意”到“强烈不同意”5 级量表", "背景说明：在问卷开头提供中性的背景介绍"];
  fix.forEach((t, i) => {
    symRow(s, rx + 0.25, 4.82 + i * 0.6, rw - 0.5, 0.5, "✓", C.green, t, { fontSize: 12 });
  });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 10 页：反例 2 数据分析不规范
// =========================================================
{
  const s = pres.addSlide();
  header(s, "反例 2", "数据分析不规范", 10);

  const lx = ML, lw = 6.4;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: lx, y: 1.9, w: lw, h: 0.5, fill: { color: C.redLight }, rectRadius: 0.05 });
  s.addText("语言学及应用语言学硕士论文（2019–2022）", { x: lx + 0.25, y: 1.9, w: 6, h: 0.5, fontSize: 12, color: C.red, bold: true, fontFace: FONT, valign: "middle" });

  // 左侧：三大问题
  card(s, lx, 2.56, lw, 2.84, C.redLight, C.red);
  s.addText("具体问题", { x: lx + 0.25, y: 2.68, w: 4, h: 0.34, fontSize: 14, color: C.red, bold: true, fontFace: FONT });
  const issues = [
    { t: "缺乏信度检验", d: "82.99% 未做 Cronbach's α，内部一致性无法保证" },
    { t: "缺乏效度检验", d: "90% 以上未做因子分析，是否测量目标概念未知" },
    { t: "统计方法错误", d: "用卡方检验分析连续变量，忽视正态性检验" },
  ];
  issues.forEach((it, i) => {
    const cy = 3.06 + i * 0.76;
    s.addText(it.t, { x: lx + 0.25, y: cy, w: lw - 0.5, h: 0.32, fontSize: 13.5, color: C.red, bold: true, fontFace: FONT });
    s.addText(it.d, { x: lx + 0.25, y: cy + 0.33, w: lw - 0.5, h: 0.45, fontSize: 11.5, color: C.ink, fontFace: FONT, lineSpacingMultiple: 1.05 });
  });

  // 右侧：典型案例
  const rx = 7.35, rw = 5.33;
  card(s, rx, 1.9, rw, 2.5, "FFFFFF", C.line);
  s.addText("典型案例", { x: rx + 0.25, y: 2.02, w: 4, h: 0.34, fontSize: 14, color: C.navy, bold: true, fontFace: FONT });
  s.addText([
    { text: "研究目标：", options: { bold: true, color: C.ink, fontSize: 12 } },
    { text: "探索学生翻译策略使用情况", options: { color: C.ink, fontSize: 12 } },
  ], { x: rx + 0.25, y: 2.42, w: rw - 0.5, h: 0.34, fontFace: FONT });
  const wrong = ["仅使用频数分析", "未做相关性 / 回归分析", "结论仅停留在描述层面"];
  wrong.forEach((t, i) => {
    symRow(s, rx + 0.25, 2.82 + i * 0.44, rw - 0.5, 0.4, "✗", C.red, t, { fontSize: 12 });
  });

  // 底部：改进建议
  card(s, ML, 5.5, 12.03, 1.3, C.greenLight, C.green);
  s.addText("改进建议", { x: ML + 0.25, y: 5.62, w: 4, h: 0.34, fontSize: 14, color: C.green, bold: true, fontFace: FONT });
  const fix2 = [
    { t: "信度检验：", d: "Cronbach's α > 0.7" },
    { t: "效度检验：", d: "因子分析 + 专家评审" },
    { t: "高级分析：", d: "相关性、回归、方差分析" },
  ];
  fix2.forEach((b, i) => {
    const cx = ML + 0.25 + (i % 3) * 3.9;
    const cy = 6.02 + Math.floor(i / 3) * 0.44;
    s.addText([
      { text: b.t, options: { bold: true, color: C.green, fontSize: 12 } },
      { text: b.d, options: { color: C.ink, fontSize: 12 } },
    ], { x: cx, y: cy, w: 3.8, h: 0.4, fontFace: FONT });
  });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 11 页：正反例对比分析
// =========================================================
{
  const s = pres.addSlide();
  header(s, "对比总结", "正反例对比分析", 11);

  const rows = [
    ["研究问题", "明确具体", "模糊宽泛"],
    ["问卷设计", "中性、逻辑严密", "诱导性、逻辑不全"],
    ["样本规模", "300+，代表性强", "< 200，代表性弱"],
    ["预调查", "有，修正后正式调查", "无，直接正式调查"],
    ["信效度检验", "有（α > 0.7，因子分析）", "无"],
    ["统计方法", "多样（卡方、相关、回归）", "单一（仅频数）"],
    ["结论可靠性", "高", "低"],
  ];
  const tableData = [
    [
      { text: "维度", options: { fill: { color: C.navy }, color: "FFFFFF", bold: true, fontSize: 13, align: "center", valign: "middle" } },
      { text: "正例", options: { fill: { color: C.green }, color: "FFFFFF", bold: true, fontSize: 13, align: "center", valign: "middle" } },
      { text: "反例", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 13, align: "center", valign: "middle" } },
    ],
    ...rows.map((r, i) => [
      { text: r[0], options: { fill: { color: C.card }, color: C.navy, bold: true, fontSize: 12.5, align: "center", valign: "middle" } },
      { text: r[1], options: { fill: { color: C.greenLight }, color: C.ink, fontSize: 12.5, align: "center", valign: "middle" } },
      { text: r[2], options: { fill: { color: C.redLight }, color: C.ink, fontSize: 12.5, align: "center", valign: "middle" } },
    ]),
  ];
  s.addTable(tableData, {
    x: ML, y: 1.85, w: 12.03, colW: [2.4, 4.815, 4.815], rowH: 0.47,
    border: { pt: 1, color: C.line }, valign: "middle", margin: 0.05,
  });

  // 关键启示
  card(s, ML, 6.05, 12.03, 0.75, C.navy);
  s.addText([
    { text: "关键启示  ", options: { color: C.gold, bold: true, fontSize: 13 } },
    { text: "科学性：严谨设计 + 规范流程　", options: { color: "FFFFFF", fontSize: 13 } },
    { text: "系统性：设计、调查、分析缺一不可　", options: { color: "FFFFFF", fontSize: 13 } },
    { text: "反思性：每一步都需要验证与改进", options: { color: "FFFFFF", fontSize: 13 } },
  ], { x: ML + 0.3, y: 6.05, w: 11.5, h: 0.75, fontFace: FONT, valign: "middle" });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 12 页：思考题
// =========================================================
{
  const s = pres.addSlide();
  header(s, "课堂互动", "思考题", 12);

  const qs = [
    { t: "方法应用", d: "研究“MTI 学生对机器翻译态度”，如何设计一份科学问卷？（提示：问题中性、选项设置、信效度检验）", c: C.navy2 },
    { t: "问题识别", d: "“翻译软件非常有用，您认为呢？”——这个问卷问题存在什么设计缺陷？", c: C.red },
    { t: "数据分析", d: "问卷回收率只有 30%，如何提高数据质量？（提示：样本量、调查方式、激励措施）", c: C.gold },
    { t: "方法组合", d: "问卷调查法在翻译研究中常与哪种方法结合？为什么？（提示：访谈法，弥补深度不足）", c: C.green },
  ];
  const qw = 5.85, qh = 2.05, gapX = 0.3, gapY = 0.26;
  qs.forEach((q, i) => {
    const cx = ML + (i % 2) * (qw + gapX);
    const cy = 1.95 + Math.floor(i / 2) * (qh + gapY);
    card(s, cx, cy, qw, qh, C.card);
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: 0.12, h: qh, fill: { color: q.c } });
    s.addText(q.t, { x: cx + 0.3, y: cy + 0.18, w: qw - 0.55, h: 0.4, fontSize: 15.5, color: q.c, bold: true, fontFace: FONT });
    s.addText(q.d, { x: cx + 0.3, y: cy + 0.62, w: qw - 0.55, h: 1.25, fontSize: 12.5, color: C.ink, fontFace: FONT, lineSpacingMultiple: 1.2, valign: "top" });
  });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 13 页：参考文献
// =========================================================
{
  const s = pres.addSlide();
  header(s, "REFERENCES", "参考文献", 13);

  // 中文文献
  s.addText("中文文献", { x: ML, y: 1.9, w: 4, h: 0.4, fontSize: 15, color: C.navy, bold: true, fontFace: FONT });
  card(s, ML, 2.34, 12.03, 1.95, C.card);
  const zhRefs = [
    { t: "[1] 王水，赵建军. (2024). 问卷调查法在文科硕士学位论文中的应用状况分析——以语言学及应用语言学专业为例. 《教育进展》14(9), 962-969. ", url: "https://doi.org/10.12677/ae.2024.1491756" },
    { t: "[2] 郑剑委. (2021). 翻译出版生产中的译者行为体系研究. 外语教学与研究出版社. ", url: "https://heep.fltrp.com/contents/17124643420005637?type=1" },
  ];
  zhRefs.forEach((r, i) => {
    const cy = 2.52 + i * 0.88;
    s.addText([
      { text: r.t, options: { color: C.ink, fontSize: 12.5 } },
      { text: "DOI / 链接", options: { color: C.link, fontSize: 12, underline: true, hyperlink: { url: r.url } } },
    ], { x: ML + 0.3, y: cy, w: 11.4, h: 0.8, fontFace: FONT, lineSpacingMultiple: 1.2, valign: "top" });
  });

  // 英文文献
  s.addText("英文文献", { x: ML, y: 4.6, w: 4, h: 0.4, fontSize: 15, color: C.navy, bold: true, fontFace: FONT });
  card(s, ML, 5.04, 12.03, 1.55, C.card);
  const enRefs = [
    { t: "[3] Schaeffer, L., et al. (2020). Translation and Interpreting Competence Questionnaire (TICQ). Journal of Translation Studies.", url: null },
    { t: "[4] 肖维青，熊凌崧. (2023). 翻译测试研究二十年：现状与展望. 上海国际研究院. ", url: "https://iots.shisu.edu.cn/b8/2b/c15638a178219/page.htm" },
  ];
  enRefs.forEach((r, i) => {
    const cy = 5.22 + i * 0.66;
    const runs = [{ text: r.t, options: { color: C.ink, fontSize: 12.5 } }];
    if (r.url) runs.push({ text: "链接", options: { color: C.link, fontSize: 12, underline: true, hyperlink: { url: r.url } } });
    s.addText(runs, { x: ML + 0.3, y: cy, w: 11.4, h: 0.6, fontFace: FONT, valign: "top" });
  });
  footer(s, "问卷调查法在翻译研究中的正反例分析");
}

// =========================================================
// 第 14 页：致谢
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: C.navyDark };
  s.addShape(pres.shapes.OVAL, { x: 9.3, y: -1.5, w: 6, h: 6, fill: { color: C.navy, transparency: 35 } });
  s.addShape(pres.shapes.OVAL, { x: 11.1, y: 4.4, w: 3.4, h: 3.4, fill: { color: C.navy2, transparency: 30 } });

  s.addShape(pres.shapes.RECTANGLE, { x: ML, y: 1.55, w: 0.18, h: 1.7, fill: { color: C.gold } });
  s.addText("THANKS", { x: ML + 0.42, y: 1.55, w: 6, h: 0.4, fontSize: 16, color: C.gold, bold: true, fontFace: FONT });
  s.addText("感谢聆听", { x: ML + 0.4, y: 2.05, w: 8, h: 1.1, fontSize: 44, color: "FFFFFF", bold: true, fontFace: FONT });

  const thanks = [
    "邹老师的悉心指导",
    "小组成员的通力合作",
    "同学们的耐心聆听",
  ];
  thanks.forEach((t, i) => {
    const cx = ML + 0.42 + i * 3.3;
    s.addText(t, { x: cx, y: 3.7, w: 3.1, h: 0.4, fontSize: 14, color: C.ice, fontFace: FONT, align: "center" });
    if (i < 2) s.addShape(pres.shapes.OVAL, { x: cx + 3.0, y: 3.82, w: 0.1, h: 0.1, fill: { color: C.gold, transparency: 30 } });
  });

  s.addText("PPT 初稿已提前一周发送至老师邮箱审阅", { x: ML + 0.42, y: 4.9, w: 9, h: 0.4, fontSize: 13, color: "9AA3C4", fontFace: FONT });
  s.addText("邮箱：小组长邮箱（请补充）", { x: ML + 0.42, y: 5.3, w: 9, h: 0.4, fontSize: 13, color: "9AA3C4", fontFace: FONT });
}

// 输出
pres.writeFile({ fileName: "D:/10_Workspace/translation-research-202610/问卷调查法在翻译研究中的正反例分析.pptx" }).then((f) => {
  console.log("已生成:", f);
});
