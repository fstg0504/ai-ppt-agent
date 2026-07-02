const deckTypeLabels = {
  business: '商务方案',
  report: '工作汇报',
  training: '培训课件',
  pitch: '路演融资'
};

const allowedSlideTypes = new Set(['cover', 'agenda', 'section', 'bullets', 'text_image', 'chart', 'timeline', 'summary']);
const allowedBlockKinds = new Set(['text', 'bullets', 'card', 'shape', 'chart', 'image_slot']);
const allowedChartKinds = new Set(['bar', 'line', 'pie']);

// PPT 规划入口 by AI.Coding：完全依赖文本模型生成内容，模型不可用时直接失败。
export async function buildDeckPlan(input) {
  if (!process.env.LLM_API_KEY) {
    throw new Error('未配置 LLM_API_KEY，无法生成 PPT。请先在 server/.env 配置文本模型。');
  }

  const deck = await buildDeckWithModel(input);
  const repairedDeck = repairDeck(deck, input);
  const validation = validateDeck(repairedDeck, input.slideCount);
  if (!validation.valid) {
    throw new Error(`模型返回结构不合法：${validation.errors.join('；')}`);
  }

  return {
    deck: repairedDeck,
    modelMode: process.env.LLM_MODEL || 'llm',
    planningWarning: validation.warnings.join('；')
  };
}

// 文本模型规划方法 by AI.Coding：要求模型只返回 JSON，便于后续稳定渲染 PPTX。
async function buildDeckWithModel(input) {
  const controller = new AbortController();
  const timeoutMs = Number(process.env.LLM_TIMEOUT_MS || 60000);
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  const endpoint = normalizeBaseUrl(process.env.LLM_BASE_URL || 'https://api.openai.com/v1');

  const requestBody = {
    model: process.env.LLM_MODEL || 'gpt-5.5',
    messages: [
      { role: 'system', content: systemPrompt() },
      { role: 'user', content: userPrompt(input) }
    ],
    temperature: Number(process.env.LLM_TEMPERATURE || 0.4),
    response_format: { type: 'json_object' }
  };

  const response = await fetch(`${endpoint}/chat/completions`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${process.env.LLM_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(requestBody),
    signal: controller.signal
  }).finally(() => clearTimeout(timeout));

  if (!response.ok) {
    const errorText = await safeReadText(response);
    throw new Error(`文本模型请求失败：${response.status} ${truncateText(errorText, 240)}`);
  }

  const data = await response.json();
  const content = data.choices?.[0]?.message?.content || '';
  if (!content) {
    throw new Error('文本模型响应为空');
  }

  return parseModelJson(content);
}

// 系统提示词 by AI.Coding：约束模型产出可编辑 PPT 所需的结构化页面数据。
function systemPrompt() {
  return [
    '你是一个成熟的 AI PPT 设计师和内容规划器。',
    '只返回 JSON 对象，不要 Markdown，不要解释。',
    'JSON 结构：{ "title": string, "subtitle": string, "theme": { "primaryColor": string, "accentColor": string, "backgroundColor": string, "fontFamily": string }, "slides": Slide[] }。',
    'Slide 结构：{ "type": "cover|agenda|section|bullets|text_image|chart|timeline|summary", "title": string, "backgroundColor"?: string, "blocks": Block[] }。',
    'Block 结构：{ "kind": "text|bullets|card|shape|chart|image_slot", "x": number, "y": number, "w": number, "h": number, "text"?: string, "items"?: string[], "title"?: string, "body"?: string, "style"?: Style, "chart"?: Chart, "imagePrompt"?: string }。',
    '坐标使用百分比，x/y/w/h 范围 0 到 100，必须保证元素不重叠、不超出页面、文字区域足够。',
    'Style 可包含：fontSize、bold、color、backgroundColor、align、radius、opacity。',
    'Chart 结构：{ "kind": "bar|line|pie", "labels": string[], "values": number[] }。',
    'slides 数量必须等于用户要求的 slideCount。',
    '如果 slideCount 为 1，唯一页面必须是 cover；如果 slideCount 大于 1，第一张必须是 cover，最后一张必须是 summary。',
    '每页由你决定最合适的排版，不要依赖固定模板；封面可大标题，内容页可卡片、左右分栏、图表聚焦、时间线等。',
    '每页至少包含 2 个 blocks，必须包含一个标题 text block。',
    '没有用户数据时，图表必须表达为示意数据，并在相邻 text 或 bullets 中说明。'
  ].join('\n');
}

// 用户提示词 by AI.Coding：把业务输入包装成明确任务，减少模型自由发挥。
function userPrompt(input) {
  return JSON.stringify({
    task: '根据输入生成可编辑 PPTX 的结构化页面计划',
    constraints: {
      slideCount: input.slideCount,
      deckType: deckTypeLabels[input.deckType] || input.deckType,
      audience: input.audience || '未指定',
      editablePptx: true,
      language: 'zh-CN'
    },
    input
  });
}

// 模型 JSON 解析方法 by AI.Coding：兼容代码块、前后解释文本和轻微格式噪声。
function parseModelJson(content) {
  const jsonText = extractJson(content);
  try {
    return JSON.parse(jsonText);
  } catch (error) {
    throw new Error(`模型 JSON 解析失败：${error.message}`);
  }
}

// JSON 提取方法 by AI.Coding：从模型响应中截取第一个完整 JSON 对象。
function extractJson(content) {
  const trimmed = content.trim();
  if (trimmed.startsWith('```')) {
    return trimmed.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```$/i, '').trim();
  }

  const firstBrace = trimmed.indexOf('{');
  const lastBrace = trimmed.lastIndexOf('}');
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    return trimmed.slice(firstBrace, lastBrace + 1);
  }

  return trimmed;
}

// 地址规范化方法 by AI.Coding：避免用户配置末尾斜杠导致接口路径重复。
function normalizeBaseUrl(baseUrl) {
  return baseUrl.replace(/\/+$/, '');
}

// 响应文本读取方法 by AI.Coding：错误响应不保证是 JSON，统一安全转文本。
async function safeReadText(response) {
  try {
    return await response.text();
  } catch (error) {
    return error.message;
  }
}

// 文本截断方法 by AI.Coding：避免后端错误信息把完整响应刷到前端。
function truncateText(text, maxLength) {
  if (!text) return '';
  return text.length > maxLength ? `${text.slice(0, maxLength)}...` : text;
}

// 计划修复方法 by AI.Coding：只清洗非内容字段，不编造 PPT 标题、页标题或要点。
function repairDeck(deck, input) {
  const slides = Array.isArray(deck.slides) ? deck.slides : [];
  return {
    title: String(deck.title || '').trim(),
    subtitle: deck.subtitle || '',
    theme: {
      primaryColor: deck.theme?.primaryColor || '#4169E1',
      accentColor: deck.theme?.accentColor || '#4FD1C5',
      backgroundColor: deck.theme?.backgroundColor || '#F7F9FE',
      fontFamily: deck.theme?.fontFamily || 'Microsoft YaHei'
    },
    slides: slides.map((slide, index) => repairSlide(slide, index))
  };
}

// 单页修复方法 by AI.Coding：只清洗 AI 返回的布局块，不补充固定排版内容。
function repairSlide(slide, index) {
  const blocks = Array.isArray(slide.blocks) ? slide.blocks.map(repairBlock).filter(Boolean) : [];

  return {
    type: String(slide.type || '').trim(),
    title: String(slide.title || '').trim(),
    backgroundColor: slide.backgroundColor || '',
    blocks
  };
}

// 布局块修复方法 by AI.Coding：限制坐标和样式范围，防止 AI 输出破坏渲染器。
function repairBlock(block) {
  if (!block || typeof block !== 'object') return null;
  const kind = String(block.kind || '').trim();
  return {
    kind,
    x: clampNumber(block.x, 0, 100),
    y: clampNumber(block.y, 0, 100),
    w: clampNumber(block.w, 1, 100),
    h: clampNumber(block.h, 1, 100),
    text: String(block.text || '').trim(),
    title: String(block.title || '').trim(),
    body: String(block.body || '').trim(),
    items: Array.isArray(block.items) ? block.items.map((item) => String(item || '').trim()).filter(Boolean).slice(0, 8) : [],
    style: repairStyle(block.style),
    chart: repairChart(block.chart),
    imagePrompt: String(block.imagePrompt || '').trim()
  };
}

// 样式修复方法 by AI.Coding：只允许渲染器支持的安全样式字段。
function repairStyle(style = {}) {
  return {
    fontSize: clampNumber(style.fontSize, 8, 54),
    bold: Boolean(style.bold),
    color: normalizeHexColor(style.color),
    backgroundColor: normalizeHexColor(style.backgroundColor),
    align: ['left', 'center', 'right'].includes(style.align) ? style.align : 'left',
    radius: clampNumber(style.radius, 0, 32),
    opacity: clampNumber(style.opacity, 0, 100)
  };
}

// 图表修复方法 by AI.Coding：只保留支持的图表类型和数值数组。
function repairChart(chart) {
  if (!chart || typeof chart !== 'object') return undefined;
  const labels = Array.isArray(chart.labels) ? chart.labels.map((item) => String(item || '').trim()).filter(Boolean) : [];
  const values = Array.isArray(chart.values) ? chart.values.map((item) => Number(item)).filter(Number.isFinite) : [];
  if (!labels.length || labels.length !== values.length) return undefined;

  return {
    kind: allowedChartKinds.has(chart.kind) ? chart.kind : 'bar',
    labels,
    values
  };
}

// 数字夹取方法 by AI.Coding：把 AI 输出约束在可渲染范围内。
function clampNumber(value, min, max) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return min;
  return Math.min(Math.max(numberValue, min), max);
}

// 颜色清洗方法 by AI.Coding：保留合法十六进制颜色，否则交给主题默认值处理。
function normalizeHexColor(value) {
  if (typeof value !== 'string') return '';
  const cleaned = value.trim();
  return /^#[0-9a-fA-F]{6}$/.test(cleaned) ? cleaned : '';
}

// 计划校验方法 by AI.Coding：对关键字段做硬校验，对可修复问题给出警告。
function validateDeck(deck, slideCount) {
  const errors = [];
  const warnings = [];

  if (!deck.title) errors.push('缺少 title');
  if (!Array.isArray(deck.slides)) errors.push('slides 不是数组');
  if (deck.slides?.length !== slideCount) errors.push(`slides 数量不是 ${slideCount}`);
  if (deck.slides?.[0]?.type !== 'cover') errors.push('第一张不是 cover');
  if (slideCount > 1 && deck.slides?.[deck.slides.length - 1]?.type !== 'summary') errors.push('最后一张不是 summary');

  deck.slides?.forEach((slide, index) => {
    if (!allowedSlideTypes.has(slide.type)) errors.push(`第 ${index + 1} 页 type 不支持`);
    if (!slide.title) errors.push(`第 ${index + 1} 页缺少 title`);
    if (!Array.isArray(slide.blocks) || slide.blocks.length < 2) errors.push(`第 ${index + 1} 页缺少布局 blocks`);
    if (!slide.blocks?.some((block) => block.kind === 'text' && block.text)) errors.push(`第 ${index + 1} 页缺少标题文本 block`);
    slide.blocks?.forEach((block, blockIndex) => {
      if (!allowedBlockKinds.has(block.kind)) errors.push(`第 ${index + 1} 页第 ${blockIndex + 1} 个 block kind 不支持`);
      if (block.kind === 'chart' && !block.chart) warnings.push(`第 ${index + 1} 页 chart block 缺少有效图表数据`);
      if ((block.kind === 'text' || block.kind === 'card') && !block.text && !block.title && !block.body) warnings.push(`第 ${index + 1} 页第 ${blockIndex + 1} 个文本类 block 内容为空`);
    });
  });

  return { valid: errors.length === 0, errors, warnings };
}
