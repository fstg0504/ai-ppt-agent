import React, { useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import './styles.css';

const initialForm = {
  topic: 'AI 办公自动化方案',
  deckType: 'business',
  slideCount: 8,
  batchCount: 1,
  audience: '企业管理层',
  notes: '强调效率提升、成本降低、落地路线和风险控制。'
};

// 主应用 by AI.Coding：管理批量生成、任务队列轮询、历史记录和下载结果。
function App() {
  const [form, setForm] = useState(initialForm);
  const [tasks, setTasks] = useState([]);
  const [selectedTaskId, setSelectedTaskId] = useState('');
  const [error, setError] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const nextSubmitAt = useRef(0);

  // 任务轮询方法 by AI.Coding：统一拉取任务列表，支持历史记录和批量队列进度展示。
  useEffect(() => {
    let cancelled = false;

    async function fetchTasks() {
      try {
        const response = await fetch('/api/tasks');
        const data = await response.json();
        if (cancelled) return;
        const list = data.tasks || [];
        setTasks(list);
        setSelectedTaskId((current) => current || list[0]?.taskId || '');
      } catch (pollError) {
        if (!cancelled) setError(pollError.message);
      }
    }

    fetchTasks();
    const timer = window.setInterval(fetchTasks, 1200);
    return () => {
      cancelled = true;
      window.clearInterval(timer);
    };
  }, []);

  // 表单更新方法 by AI.Coding：统一处理字段变更，避免每个输入框重复状态逻辑。
  function updateField(field, value) {
    setForm((current) => ({ ...current, [field]: value }));
  }

  // 创建任务方法 by AI.Coding：带 1.5 秒防抖，避免用户连续点击重复创建批量任务。
  async function createDeck(event) {
    event.preventDefault();
    const now = Date.now();
    if (now < nextSubmitAt.current) return;
    nextSubmitAt.current = now + 1500;

    setError('');
    setIsSubmitting(true);

    try {
      const response = await fetch('/api/decks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...form,
          slideCount: Number(form.slideCount),
          batchCount: Number(form.batchCount)
        })
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || '创建任务失败');
      setTasks((current) => [...(data.tasks || []), ...current]);
      setSelectedTaskId(data.tasks?.[0]?.taskId || selectedTaskId);
    } catch (submitError) {
      setError(submitError.message);
    } finally {
      window.setTimeout(() => setIsSubmitting(false), 1500);
    }
  }

  const selectedTask = tasks.find((task) => task.taskId === selectedTaskId) || tasks[0];
  const activeTasks = tasks.filter((task) => task.status === 'queued' || task.status === 'running').length;

  return (
    <main className="page-shell">
      <section className="hero-panel">
        <div>
          <p className="eyebrow">Local AI PPT Agent</p>
          <h1>用结构化 Agent 队列生成可编辑 PPTX</h1>
          <p className="hero-copy">批量任务会在本地按顺序排队，文本模型规划内容，后端渲染 PPTX，历史记录保存到输出目录。</p>
        </div>
        <div className="status-card">
          <span>队列状态</span>
          <strong>{activeTasks ? `${activeTasks} 个进行中` : '空闲'}</strong>
        </div>
      </section>

      <section className="workspace-grid">
        <form className="control-panel" onSubmit={createDeck}>
          <label>
            PPT 主题
            <input value={form.topic} onChange={(event) => updateField('topic', event.target.value)} required />
          </label>

          <div className="form-row three-columns">
            <label>
              类型
              <select value={form.deckType} onChange={(event) => updateField('deckType', event.target.value)}>
                <option value="business">商务方案</option>
                <option value="report">工作汇报</option>
                <option value="training">培训课件</option>
                <option value="pitch">路演融资</option>
              </select>
            </label>
            <label>
              页数
              <input type="number" min="1" max="15" value={form.slideCount} onChange={(event) => updateField('slideCount', event.target.value)} />
            </label>
            <label>
              批量数
              <input type="number" min="1" max="10" value={form.batchCount} onChange={(event) => updateField('batchCount', event.target.value)} />
            </label>
          </div>

          <label>
            受众
            <input value={form.audience} onChange={(event) => updateField('audience', event.target.value)} />
          </label>

          <label>
            补充材料
            <textarea value={form.notes} onChange={(event) => updateField('notes', event.target.value)} rows="6" />
          </label>

          <button disabled={isSubmitting}>{isSubmitting ? '已加入队列...' : '批量生成 PPT'}</button>
          {error ? <p className="error-text">{error}</p> : null}

          <TaskList tasks={tasks} selectedTaskId={selectedTask?.taskId} onSelect={setSelectedTaskId} />
        </form>

        <section className="result-panel">
          <div className="result-header">
            <div>
              <p className="eyebrow">生成结果</p>
              <h2>{selectedTask?.deck?.title || selectedTask?.payload?.topic || '还没有生成 PPT'}</h2>
            </div>
            {selectedTask?.status === 'done' ? <a className="download-link" href={selectedTask.downloadUrl}>下载 PPTX</a> : null}
          </div>

          {selectedTask ? <Progress task={selectedTask} /> : <p className="empty-state">填写左侧表单后开始生成，内容规划完全依赖已配置的文本模型。</p>}

          {selectedTask?.planningWarning ? <p className="warning-text">{selectedTask.planningWarning}</p> : null}

          <SlidePreview deck={selectedTask?.deck} />
        </section>
      </section>
    </main>
  );
}

// 幻灯片预览组件 by AI.Coding：复用 AI 返回的 blocks 在浏览器中预览，避免额外消耗模型 token。
function SlidePreview({ deck }) {
  const slides = deck?.slides || [];
  if (!slides.length) return null;

  return (
    <div className="preview-list">
      {slides.map((slide, index) => (
        <article className="preview-card" key={`${slide.title}-${index}`}>
          <div className="preview-title">
            <span>{String(index + 1).padStart(2, '0')}</span>
            <strong>{slide.title}</strong>
          </div>
          <div className="slide-preview" style={{ background: slide.backgroundColor || deck.theme?.backgroundColor || '#F7F9FE' }}>
            {slide.blocks?.map((block, blockIndex) => <PreviewBlock block={block} theme={deck.theme || {}} key={`${block.kind}-${blockIndex}`} />)}
          </div>
        </article>
      ))}
    </div>
  );
}

// 预览块组件 by AI.Coding：把 AI 百分比坐标转成 CSS absolute 定位，贴近 PPTX 导出效果。
function PreviewBlock({ block, theme }) {
  const style = {
    left: `${block.x}%`,
    top: `${block.y}%`,
    width: `${block.w}%`,
    height: `${block.h}%`,
    color: block.style?.color || '#172033',
    background: block.style?.backgroundColor || 'transparent',
    fontSize: `${block.style?.fontSize || 16}px`,
    fontWeight: block.style?.bold ? 800 : 500,
    textAlign: block.style?.align || 'left',
    borderRadius: `${block.style?.radius || 0}px`,
    opacity: block.style?.opacity ? 1 - block.style.opacity / 100 : 1,
    fontFamily: theme.fontFamily || 'Microsoft YaHei'
  };

  if (block.kind === 'shape') return <div className="preview-block preview-shape" style={style} />;
  if (block.kind === 'bullets') return <ul className="preview-block preview-bullets" style={style}>{block.items?.map((item) => <li key={item}>{item}</li>)}</ul>;
  if (block.kind === 'card') return <div className="preview-block preview-card-block" style={style}><strong>{block.title}</strong><p>{block.body}</p></div>;
  if (block.kind === 'chart') return <PreviewChart block={block} theme={theme} style={style} />;
  if (block.kind === 'image_slot') return <div className="preview-block preview-image-slot" style={style}>AI image<br />{block.imagePrompt}</div>;
  return <div className="preview-block preview-text" style={style}>{block.text || block.title || block.body}</div>;
}

// 图表预览组件 by AI.Coding：浏览器侧用轻量 CSS 柱状图展示 AI 图表数据，不再次调用 AI。
function PreviewChart({ block, theme, style }) {
  const chart = block.chart;
  const maxValue = Math.max(...(chart?.values || [1]), 1);
  return (
    <div className="preview-block preview-chart" style={style}>
      {chart?.labels?.map((label, index) => {
        const value = chart.values[index] || 0;
        return (
          <div className="preview-chart-item" key={label}>
            <span>{value}</span>
            <div style={{ height: `${Math.max((value / maxValue) * 70, 5)}%`, background: theme.primaryColor || '#4169E1' }} />
            <small>{label}</small>
          </div>
        );
      })}
    </div>
  );
}

// 任务列表组件 by AI.Coding：展示本地历史记录、排队状态和每个完成任务的下载入口。
function TaskList({ tasks, selectedTaskId, onSelect }) {
  return (
    <section className="task-list-panel">
      <div className="task-list-header">
        <strong>任务历史</strong>
        <span>{tasks.length} 条</span>
      </div>
      <div className="task-list">
        {tasks.length ? tasks.map((task) => (
          <div className={`task-item ${selectedTaskId === task.taskId ? 'active' : ''}`} key={task.taskId} role="button" tabIndex="0" onClick={() => onSelect(task.taskId)} onKeyDown={(event) => event.key === 'Enter' && onSelect(task.taskId)}>
            <span className={`task-dot ${task.status}`} />
            <span className="task-meta">
              <strong>{task.deck?.title || task.payload?.topic || '未命名 PPT'}</strong>
              <small>{formatTaskTime(task.createdAt)} · {task.status} · {task.progress}%</small>
            </span>
            {task.status === 'done' ? <a href={task.downloadUrl} onClick={(event) => event.stopPropagation()}>下载</a> : null}
          </div>
        )) : <p className="empty-mini">暂无历史记录</p>}
      </div>
    </section>
  );
}

// 进度组件 by AI.Coding：把后端任务状态转换为用户能理解的状态条。
function Progress({ task }) {
  const labelMap = {
    queued: '排队中',
    running: '生成中',
    done: '已完成',
    failed: '失败'
  };

  return (
    <div className={`progress-box ${task.status}`}>
      <div className="progress-copy">
        <strong>{labelMap[task.status] || task.status}</strong>
        <span>{task.message}</span>
      </div>
      <p className="model-line">模型模式：{task.modelMode || 'pending'}{task.fileName ? ` · ${task.fileName}` : ''}</p>
      <div className="progress-track">
        <div style={{ width: `${task.progress || 0}%` }} />
      </div>
    </div>
  );
}

// 时间展示方法 by AI.Coding：把历史任务创建时间转换为本地短格式。
function formatTaskTime(value) {
  if (!value) return '未知时间';
  return new Date(value).toLocaleString();
}

createRoot(document.getElementById('root')).render(<App />);
