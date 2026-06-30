const deckTypeLabels = {
  business: '商务方案',
  report: '工作汇报',
  training: '培训课件',
  pitch: '路演融资'
};

// PPT 规划入口 by AI.Coding：优先调用文本模型，失败或未配置时使用本地兜底。
export async function buildDeckPlan(input) {
  if (process.env.LLM_API_KEY) {
    try {
      const deck = await buildDeckWithModel(input);
      return { deck: sanitizeDeck(deck, input), modelMode: process.env.LLM_MODEL || 'llm' };
    } catch (error) {
      console.warn('LLM planning failed, fallback to local plan:', error.message);
    }
  }

  return { deck: buildFallbackDeck(input), modelMode: 'local-fallback' };
}

// 文本模型规划方法 by AI.Coding：要求模型只返回 JSON，便于后续稳定渲染 PPTX。
async function buildDeckWithModel(input) {
  const response = await fetch(`${process.env.LLM_BASE_URL || 'https://api.openai.com/v1'}/chat/completions`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${process.env.LLM_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: process.env.LLM_MODEL || 'gpt-5.5',
      messages: [
        { role: 'system', content: systemPrompt() },
        { role: 'user', content: JSON.stringify(input) }
      ],
      temperature: 0.7
    })
  });

  if (!response.ok) {
    throw new Error(`文本模型请求失败：${response.status}`);
  }

  const data = await response.json();
  const content = data.choices?.[0]?.message?.content || '';
  return JSON.parse(extractJson(content));
}

// 系统提示词 by AI.Coding：约束模型产出可编辑 PPT 所需的结构化页面数据。
function systemPrompt() {
  return [
    '你是一个成熟的 AI PPT 产品规划器。',
    '只返回 JSON，不要 Markdown，不要解释。',
    'JSON 结构：{ "title": string, "subtitle": string, "theme": { "primaryColor": string, "accentColor": string }, "slides": Slide[] }。',
    'Slide 结构：{ "type": "cover|agenda|section|bullets|text_image|chart|timeline|summary", "title": string, "subtitle"?: string, "bullets": string[], "chart"?: { "kind": "bar|line|pie", "labels": string[], "values": number[] }, "imagePrompt"?: string }。',
    '每页 bullets 控制在 3 到 5 条，标题要适合直接放进 PPT。',
    '没有用户数据时，图表必须表达为示意数据，并在 bullets 中说明。'
  ].join('\n');
}

// JSON 提取方法 by AI.Coding：兼容模型偶尔包裹代码块的情况。
function extractJson(content) {
  const trimmed = content.trim();
  if (trimmed.startsWith('```')) {
    return trimmed.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```$/i, '').trim();
  }
  return trimmed;
}

// 本地兜底规划 by AI.Coding：无模型配置也能跑通完整 PPT 生成链路。
function buildFallbackDeck(input) {
  const label = deckTypeLabels[input.deckType] || '演示文稿';
  const baseSlides = [
    {
      type: 'cover',
      title: input.topic,
      subtitle: `${label} | 面向${input.audience || '目标受众'}`,
      bullets: ['结构化生成', '可编辑 PPTX', '本地开发闭环'],
      imagePrompt: `professional presentation cover about ${input.topic}`
    },
    {
      type: 'agenda',
      title: '目录',
      bullets: ['背景与目标', '核心方案', '价值评估', '落地路线', '风险与保障']
    },
    {
      type: 'section',
      title: '背景与目标',
      bullets: ['明确当前业务痛点', '定义可衡量的改进目标', '对齐受众关注点']
    },
    {
      type: 'text_image',
      title: '核心方案概览',
      bullets: ['文本模型负责规划和文案', '图片模型补齐封面与插图', '渲染引擎输出可编辑 PPTX'],
      imagePrompt: `clean workflow diagram for ${input.topic}`
    },
    {
      type: 'chart',
      title: '预期收益示意',
      bullets: ['以下为示意数据，正式汇报需替换为真实业务数据', '效率、质量和成本作为首批评估维度', '后续可接入 Excel/CSV 自动制图'],
      chart: { kind: 'bar', labels: ['效率', '质量', '成本优化'], values: [42, 28, 35] }
    },
    {
      type: 'timeline',
      title: '落地路线',
      bullets: ['第 1 阶段：跑通生成闭环', '第 2 阶段：接入素材和图表', '第 3 阶段：模板库与单页重生成']
    },
    {
      type: 'bullets',
      title: '风险与保障',
      bullets: ['避免图表数据幻觉', '保留素材来源和生成日志', '失败时使用占位内容保证任务完成']
    },
    {
      type: 'summary',
      title: '总结',
      bullets: ['先保证稳定生成可编辑 PPTX', '再逐步增强图片、图表和模板能力', '最终形成可产品化的本地 AI PPT Agent']
    }
  ];

  return sanitizeDeck({
    title: input.topic,
    subtitle: `${label} 自动生成`,
    theme: { primaryColor: '#4169E1', accentColor: '#4FD1C5' },
    slides: fitSlideCount(baseSlides, input.slideCount)
  }, input);
}

// 页数适配方法 by AI.Coding：根据用户页数增删中间页，保持封面和总结稳定存在。
function fitSlideCount(slides, slideCount) {
  if (slides.length === slideCount) return slides;
  if (slides.length > slideCount) {
    return [...slides.slice(0, Math.max(slideCount - 1, 1)), slides[slides.length - 1]];
  }

  const result = [...slides];
  while (result.length < slideCount) {
    result.splice(result.length - 1, 0, {
      type: 'bullets',
      title: `补充分析 ${result.length - 1}`,
      bullets: ['结合用户材料展开细节', '提炼可执行动作', '补充衡量指标和负责人']
    });
  }
  return result;
}

// 计划清洗方法 by AI.Coding：兜住模型字段缺失，保证渲染阶段不崩溃。
function sanitizeDeck(deck, input) {
  const slides = Array.isArray(deck.slides) ? deck.slides : [];
  return {
    title: deck.title || input.topic,
    subtitle: deck.subtitle || '',
    theme: {
      primaryColor: deck.theme?.primaryColor || '#4169E1',
      accentColor: deck.theme?.accentColor || '#4FD1C5'
    },
    slides: fitSlideCount(slides, input.slideCount).map((slide, index) => ({
      type: slide.type || (index === 0 ? 'cover' : 'bullets'),
      title: slide.title || `第 ${index + 1} 页`,
      subtitle: slide.subtitle || '',
      bullets: Array.isArray(slide.bullets) ? slide.bullets.slice(0, 5) : [],
      chart: slide.chart,
      imagePrompt: slide.imagePrompt || ''
    }))
  };
}
