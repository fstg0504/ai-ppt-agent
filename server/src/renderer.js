import fs from 'node:fs/promises';
import path from 'node:path';
import PptxGenJS from 'pptxgenjs';

const layout = { width: 13.333, height: 7.5 };

// PPT 渲染入口 by AI.Coding：按 AI 返回的 blocks 坐标渲染，不再根据 type 套固定版式。
export async function renderDeck(deck, filePath) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'AI PPT Agent';
  pptx.subject = deck.subtitle;
  pptx.title = deck.title;
  pptx.company = 'Local AI PPT Agent';

  deck.slides.forEach((slide) => renderSlide(pptx, deck, slide));
  await pptx.writeFile({ fileName: filePath });
}

// 单页渲染方法 by AI.Coding：只消费 AI 布局块，保证预览和导出使用同一份结构。
function renderSlide(pptx, deck, slide) {
  const page = pptx.addSlide();
  const theme = normalizeTheme(deck.theme);
  page.background = { color: normalizeColor(slide.backgroundColor || theme.backgroundColor, 'F7F9FE') };

  slide.blocks.forEach((block) => renderBlock(page, block, theme));
}

// 布局块渲染方法 by AI.Coding：根据 block.kind 执行通用元素渲染，而不是固定页面模板。
function renderBlock(page, block, theme) {
  const box = toPptBox(block);
  if (block.kind === 'shape') renderShape(page, block, box, theme);
  if (block.kind === 'text') renderText(page, block, box, theme);
  if (block.kind === 'bullets') renderBullets(page, block, box, theme);
  if (block.kind === 'card') renderCard(page, block, box, theme);
  if (block.kind === 'chart') renderChart(page, block, box, theme);
  if (block.kind === 'image_slot') renderImageSlot(page, block, box, theme);
}

// 文本渲染方法 by AI.Coding：用可编辑文本框承载 AI 指定的文本内容。
function renderText(page, block, box, theme) {
  const style = blockStyle(block, theme);
  page.addText(block.text || block.title || block.body, {
    ...box,
    fontFace: theme.fontFamily,
    fontSize: style.fontSize,
    bold: style.bold,
    color: style.color,
    align: style.align,
    valign: 'mid',
    fit: 'shrink',
    margin: 0.06,
    breakLine: false
  });
}

// 要点渲染方法 by AI.Coding：按 AI 指定位置渲染可编辑项目符号列表。
function renderBullets(page, block, box, theme) {
  const style = blockStyle(block, theme);
  page.addText(block.items.map((item) => ({ text: item, options: { bullet: { type: 'ul' } } })), {
    ...box,
    fontFace: theme.fontFamily,
    fontSize: style.fontSize,
    color: style.color,
    fit: 'shrink',
    breakLine: false,
    margin: 0.08,
    paraSpaceAfterPt: 8
  });
}

// 卡片渲染方法 by AI.Coding：卡片是通用内容块，标题和正文都保持可编辑。
function renderCard(page, block, box, theme) {
  const style = blockStyle(block, theme);
  page.addShape('roundRect', {
    ...box,
    rectRadius: 0.08,
    fill: { color: normalizeColor(style.backgroundColor, 'FFFFFF'), transparency: style.opacity },
    line: { color: normalizeColor(theme.primaryColor, '4169E1'), transparency: 72 }
  });
  if (block.title) {
    page.addText(block.title, { x: box.x + 0.18, y: box.y + 0.16, w: box.w - 0.36, h: Math.min(0.42, box.h), fontFace: theme.fontFamily, fontSize: Math.min(style.fontSize + 2, 28), bold: true, color: style.color, fit: 'shrink' });
  }
  if (block.body) {
    page.addText(block.body, { x: box.x + 0.18, y: box.y + 0.68, w: box.w - 0.36, h: Math.max(box.h - 0.84, 0.2), fontFace: theme.fontFamily, fontSize: style.fontSize, color: style.color, fit: 'shrink', breakLine: false, margin: 0.02 });
  }
}

// 形状渲染方法 by AI.Coding：作为 AI 设计中的背景色块、装饰线和视觉分隔。
function renderShape(page, block, box, theme) {
  const style = blockStyle(block, theme);
  page.addShape('roundRect', {
    ...box,
    rectRadius: 0.08,
    fill: { color: normalizeColor(style.backgroundColor || style.color, theme.accentColor), transparency: style.opacity },
    line: { color: normalizeColor(style.backgroundColor || style.color, theme.accentColor), transparency: 100 }
  });
}

// 图表渲染方法 by AI.Coding：当前用可编辑形状表达简单柱状图，图表数据由 AI 提供。
function renderChart(page, block, box, theme) {
  const chart = block.chart;
  if (!chart) return;
  const maxValue = Math.max(...chart.values, 1);
  const barCount = chart.labels.length;
  const gap = box.w * 0.08;
  const barWidth = (box.w - gap * (barCount + 1)) / Math.max(barCount, 1);
  page.addShape('roundRect', { ...box, rectRadius: 0.06, fill: { color: 'FFFFFF', transparency: 8 }, line: { color: 'DDE5F5' } });

  chart.labels.forEach((label, index) => {
    const value = chart.values[index] || 0;
    const barHeight = Math.max((value / maxValue) * (box.h * 0.58), 0.08);
    const x = box.x + gap + index * (barWidth + gap);
    const y = box.y + box.h * 0.68 - barHeight;
    page.addShape('rect', { x, y, w: barWidth, h: barHeight, fill: { color: normalizeColor(theme.primaryColor, '4169E1') }, line: { color: normalizeColor(theme.primaryColor, '4169E1') } });
    page.addText(String(value), { x: x - 0.05, y: y - 0.28, w: barWidth + 0.1, h: 0.18, fontSize: 8, bold: true, color: '25324A', align: 'center', fit: 'shrink' });
    page.addText(label, { x: x - 0.12, y: box.y + box.h * 0.74, w: barWidth + 0.24, h: box.h * 0.2, fontSize: 8, color: '65748C', align: 'center', fit: 'shrink' });
  });
}

// 图片占位渲染方法 by AI.Coding：尚未接入图片模型时，按 AI 指定位置展示可替换素材槽。
function renderImageSlot(page, block, box, theme) {
  page.addShape('roundRect', { ...box, rectRadius: 0.08, fill: { color: 'EAF0FF' }, line: { color: normalizeColor(theme.accentColor, '4FD1C5'), transparency: 30 } });
  page.addText(block.imagePrompt ? `AI image\n${block.imagePrompt}` : 'AI image slot', { ...box, fontSize: 12, bold: true, color: normalizeColor(theme.primaryColor, '4169E1'), align: 'center', valign: 'mid', fit: 'shrink' });
}

// 坐标转换方法 by AI.Coding：AI 使用百分比坐标，PptxGenJS 使用英寸坐标。
function toPptBox(block) {
  return {
    x: (block.x / 100) * layout.width,
    y: (block.y / 100) * layout.height,
    w: (block.w / 100) * layout.width,
    h: (block.h / 100) * layout.height
  };
}

// 块样式方法 by AI.Coding：融合 AI 样式和主题默认值，避免样式缺失导致不可读。
function blockStyle(block, theme) {
  return {
    fontSize: block.style?.fontSize || 16,
    bold: Boolean(block.style?.bold),
    color: normalizeColor(block.style?.color, theme.textColor),
    backgroundColor: normalizeColor(block.style?.backgroundColor, 'FFFFFF'),
    align: block.style?.align || 'left',
    opacity: block.style?.opacity || 0
  };
}

// 主题清洗方法 by AI.Coding：统一 PPTX 渲染使用的颜色和字体默认值。
function normalizeTheme(theme = {}) {
  return {
    primaryColor: theme.primaryColor || '#4169E1',
    accentColor: theme.accentColor || '#4FD1C5',
    backgroundColor: theme.backgroundColor || '#F7F9FE',
    textColor: '#172033',
    fontFamily: theme.fontFamily || 'Microsoft YaHei'
  };
}

// 颜色清洗方法 by AI.Coding：PptxGenJS 需要不带井号的十六进制颜色。
function normalizeColor(value, fallback) {
  const source = typeof value === 'string' && value ? value : fallback;
  const cleaned = source.replace('#', '').trim();
  return /^[0-9a-fA-F]{6}$/.test(cleaned) ? cleaned.toUpperCase() : fallback.replace('#', '');
}
