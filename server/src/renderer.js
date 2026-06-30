import fs from 'node:fs/promises';
import path from 'node:path';
import PptxGenJS from 'pptxgenjs';

const layout = { name: 'LAYOUT_WIDE', width: 13.333, height: 7.5 };

// PPT 渲染入口 by AI.Coding：把结构化页面计划转换为可编辑 PPTX 文件。
export async function renderDeck(deck, filePath) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'AI PPT Agent';
  pptx.subject = deck.subtitle;
  pptx.title = deck.title;
  pptx.company = 'Local AI PPT Agent';

  deck.slides.forEach((slide, index) => {
    renderSlide(pptx, deck, slide, index);
  });

  await pptx.writeFile({ fileName: filePath });
}

// 单页渲染方法 by AI.Coding：根据页面类型选择版式，保证首版视觉统一。
function renderSlide(pptx, deck, slide, index) {
  const page = pptx.addSlide();
  const primary = normalizeColor(deck.theme.primaryColor, '4169E1');
  const accent = normalizeColor(deck.theme.accentColor, '4FD1C5');

  page.background = { color: index === 0 ? primary : 'F7F9FE' };

  if (slide.type === 'cover') {
    renderCover(page, deck, slide, primary, accent);
    return;
  }

  renderHeader(page, slide, primary, index);

  if (slide.type === 'chart' && slide.chart) {
    renderChartSlide(page, slide, primary, accent);
  } else if (slide.type === 'timeline') {
    renderTimelineSlide(page, slide, primary, accent);
  } else if (slide.type === 'text_image') {
    renderTextImageSlide(page, slide, primary, accent);
  } else {
    renderBulletSlide(page, slide, primary);
  }

  renderFooter(page, index + 1, deck.slides.length, primary);
}

// 封面渲染方法 by AI.Coding：用形状和文本组成可编辑封面，避免整页图片不可编辑。
function renderCover(page, deck, slide, primary, accent) {
  page.addShape('rect', { x: 0, y: 0, w: layout.width, h: layout.height, fill: { color: primary }, line: { color: primary } });
  page.addShape('arc', { x: 8.2, y: -0.8, w: 6, h: 6, fill: { color: accent, transparency: 28 }, line: { color: accent, transparency: 100 } });
  page.addShape('rect', { x: 0.65, y: 0.62, w: 1.05, h: 0.12, fill: { color: accent }, line: { color: accent } });
  page.addText('AI PPT AGENT', { x: 0.65, y: 0.85, w: 4, h: 0.3, fontSize: 12, bold: true, color: 'DCE7FF', charSpace: 2 });
  page.addText(slide.title || deck.title, { x: 0.65, y: 2.05, w: 7.5, h: 1.55, fontSize: 38, bold: true, color: 'FFFFFF', breakLine: false, fit: 'shrink' });
  page.addText(slide.subtitle || deck.subtitle, { x: 0.7, y: 3.82, w: 7.2, h: 0.55, fontSize: 18, color: 'DCE7FF' });
  page.addText((slide.bullets || []).join('   |   '), { x: 0.7, y: 5.7, w: 8.6, h: 0.45, fontSize: 14, color: 'EEF4FF' });
  page.addShape('roundRect', { x: 9.0, y: 4.65, w: 3.25, h: 1.25, rectRadius: 0.12, fill: { color: 'FFFFFF', transparency: 8 }, line: { color: 'FFFFFF', transparency: 60 } });
  page.addText('Generated locally\nEditable PPTX', { x: 9.22, y: 4.9, w: 2.8, h: 0.76, fontSize: 18, bold: true, color: primary, align: 'center' });
}

// 标题栏渲染方法 by AI.Coding：所有内容页共用页眉，保证风格统一。
function renderHeader(page, slide, primary, index) {
  page.addText(String(index + 1).padStart(2, '0'), { x: 0.55, y: 0.46, w: 0.6, h: 0.3, fontSize: 11, bold: true, color: primary });
  page.addShape('rect', { x: 1.25, y: 0.58, w: 0.82, h: 0.05, fill: { color: primary }, line: { color: primary } });
  page.addText(slide.title, { x: 0.65, y: 0.95, w: 9.8, h: 0.62, fontSize: 28, bold: true, color: '172033', fit: 'shrink' });
  if (slide.subtitle) {
    page.addText(slide.subtitle, { x: 0.68, y: 1.58, w: 9.8, h: 0.36, fontSize: 12, color: '6B7890' });
  }
}

// 要点页渲染方法 by AI.Coding：使用可编辑文本框和编号块表现重点内容。
function renderBulletSlide(page, slide, primary) {
  const bullets = slide.bullets.length ? slide.bullets : ['待补充内容'];
  bullets.forEach((bullet, index) => {
    const y = 2.15 + index * 0.78;
    page.addShape('roundRect', { x: 0.8, y, w: 0.42, h: 0.42, rectRadius: 0.08, fill: { color: primary }, line: { color: primary } });
    page.addText(String(index + 1), { x: 0.8, y: y + 0.08, w: 0.42, h: 0.18, fontSize: 10, bold: true, color: 'FFFFFF', align: 'center' });
    page.addText(bullet, { x: 1.48, y: y - 0.02, w: 10.5, h: 0.48, fontSize: 19, color: '25324A', fit: 'shrink' });
  });
}

// 图文页渲染方法 by AI.Coding：右侧先用占位视觉，后续可替换为 gpt-image2 输出图片。
function renderTextImageSlide(page, slide, primary, accent) {
  renderBulletSlide(page, slide, primary);
  page.addShape('roundRect', { x: 8.25, y: 2.04, w: 4.15, h: 3.5, rectRadius: 0.18, fill: { color: 'EAF0FF' }, line: { color: 'CFD9F5' } });
  page.addShape('arc', { x: 8.72, y: 2.36, w: 2.1, h: 2.1, fill: { color: accent, transparency: 24 }, line: { color: accent, transparency: 100 } });
  page.addShape('rect', { x: 9.28, y: 4.44, w: 2.2, h: 0.12, fill: { color: primary }, line: { color: primary } });
  page.addText('Image slot\nfor gpt-image2', { x: 8.72, y: 3.22, w: 3.18, h: 0.86, fontSize: 21, bold: true, color: primary, align: 'center' });
}

// 图表页渲染方法 by AI.Coding：首版用 PPT 原生形状画简单柱状图，保证可编辑。
function renderChartSlide(page, slide, primary, accent) {
  const chart = slide.chart;
  const labels = Array.isArray(chart.labels) ? chart.labels : [];
  const values = Array.isArray(chart.values) ? chart.values : [];
  const maxValue = Math.max(...values, 1);

  renderBulletsAt(page, slide.bullets.slice(0, 3), 0.8, 2.15, 5.1, primary);
  page.addShape('roundRect', { x: 6.6, y: 2.0, w: 5.65, h: 3.7, rectRadius: 0.14, fill: { color: 'FFFFFF' }, line: { color: 'DDE5F5' } });

  labels.forEach((label, index) => {
    const barWidth = 0.72;
    const gap = 0.52;
    const x = 7.0 + index * (barWidth + gap);
    const height = Math.max((values[index] / maxValue) * 2.2, 0.15);
    const y = 4.85 - height;
    page.addShape('rect', { x, y, w: barWidth, h: height, fill: { color: index % 2 ? accent : primary }, line: { color: index % 2 ? accent : primary } });
    page.addText(String(values[index] || 0), { x: x - 0.05, y: y - 0.34, w: barWidth + 0.1, h: 0.2, fontSize: 10, bold: true, color: '25324A', align: 'center' });
    page.addText(label, { x: x - 0.18, y: 5.08, w: barWidth + 0.36, h: 0.42, fontSize: 9, color: '65748C', align: 'center', fit: 'shrink' });
  });
}

// 时间线页渲染方法 by AI.Coding：把步骤转成横向里程碑，适合路线图表达。
function renderTimelineSlide(page, slide, primary, accent) {
  const items = slide.bullets.length ? slide.bullets : ['阶段一', '阶段二', '阶段三'];
  page.addShape('rect', { x: 1.05, y: 3.65, w: 10.9, h: 0.05, fill: { color: 'CBD6EF' }, line: { color: 'CBD6EF' } });
  items.forEach((item, index) => {
    const x = 1.1 + index * (10.6 / Math.max(items.length - 1, 1));
    page.addShape('ellipse', { x: x - 0.22, y: 3.43, w: 0.44, h: 0.44, fill: { color: index % 2 ? accent : primary }, line: { color: 'FFFFFF', width: 2 } });
    page.addText(item, { x: x - 1.0, y: index % 2 ? 4.05 : 2.62, w: 2, h: 0.56, fontSize: 12, bold: true, color: '25324A', align: 'center', fit: 'shrink' });
  });
}

// 指定位置要点渲染方法 by AI.Coding：图表页复用该方法控制文本区域大小。
function renderBulletsAt(page, bullets, x, y, width, primary) {
  bullets.forEach((bullet, index) => {
    page.addShape('ellipse', { x, y: y + index * 0.72, w: 0.16, h: 0.16, fill: { color: primary }, line: { color: primary } });
    page.addText(bullet, { x: x + 0.34, y: y + index * 0.72 - 0.08, w: width, h: 0.4, fontSize: 15, color: '25324A', fit: 'shrink' });
  });
}

// 页脚渲染方法 by AI.Coding：展示页码和产品标识，便于预览检查完整性。
function renderFooter(page, pageNumber, total, primary) {
  page.addShape('rect', { x: 0.65, y: 6.85, w: 11.95, h: 0.01, fill: { color: 'DDE5F5' }, line: { color: 'DDE5F5' } });
  page.addText('AI PPT Agent', { x: 0.65, y: 6.98, w: 2.2, h: 0.2, fontSize: 9, color: primary, bold: true });
  page.addText(`${pageNumber}/${total}`, { x: 11.65, y: 6.98, w: 0.9, h: 0.2, fontSize: 9, color: '6B7890', align: 'right' });
}

// 颜色清洗方法 by AI.Coding：PptxGenJS 需要不带井号的十六进制颜色。
function normalizeColor(value, fallback) {
  if (typeof value !== 'string') return fallback;
  const cleaned = value.replace('#', '').trim();
  return /^[0-9a-fA-F]{6}$/.test(cleaned) ? cleaned.toUpperCase() : fallback;
}
