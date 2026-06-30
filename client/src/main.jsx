import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import './styles.css';

const initialForm = {
  topic: 'AI 办公自动化方案',
  deckType: 'business',
  slideCount: 8,
  audience: '企业管理层',
  notes: '强调效率提升、成本降低、落地路线和风险控制。'
};

// 主应用 by AI.Coding：管理输入、生成任务、轮询状态和下载结果。
function App() {
  const [form, setForm] = useState(initialForm);
  const [taskId, setTaskId] = useState('');
  const [task, setTask] = useState(null);
  const [error, setError] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);

  // 轮询方法 by AI.Coding：生成任务耗时不固定，前端用 taskId 定期刷新状态。
  useEffect(() => {
    if (!taskId) return undefined;

    let cancelled = false;
    const timer = window.setInterval(async () => {
      try {
        const response = await fetch(`/api/tasks/${taskId}`);
        const data = await response.json();
        if (!cancelled) setTask(data);
        if (data.status === 'done' || data.status === 'failed') {
          window.clearInterval(timer);
        }
      } catch (pollError) {
        if (!cancelled) setError(pollError.message);
      }
    }, 1000);

    return () => {
      cancelled = true;
      window.clearInterval(timer);
    };
  }, [taskId]);

  // 表单更新方法 by AI.Coding：统一处理字段变更，避免每个输入框重复状态逻辑。
  function updateField(field, value) {
    setForm((current) => ({ ...current, [field]: value }));
  }

  // 创建任务方法 by AI.Coding：提交用户输入，让后端异步生成 PPT。
  async function createDeck(event) {
    event.preventDefault();
    setError('');
    setTask(null);
    setTaskId('');
    setIsSubmitting(true);

    try {
      const response = await fetch('/api/decks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...form, slideCount: Number(form.slideCount) })
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || '创建任务失败');
      setTaskId(data.taskId);
      setTask(data);
    } catch (submitError) {
      setError(submitError.message);
    } finally {
      setIsSubmitting(false);
    }
  }

  const slides = task?.deck?.slides || [];

  return (
    <main className="page-shell">
      <section className="hero-panel">
        <div>
          <p className="eyebrow">Local AI PPT Agent</p>
          <h1>用结构化 Agent 流程生成可编辑 PPTX</h1>
          <p className="hero-copy">首版先跑通本地闭环：文本模型规划内容，后端渲染 PPTX，前端支持热更新和下载。</p>
        </div>
        <div className="status-card">
          <span>当前模式</span>
          <strong>{task?.modelMode || '等待生成'}</strong>
        </div>
      </section>

      <section className="workspace-grid">
        <form className="control-panel" onSubmit={createDeck}>
          <label>
            PPT 主题
            <input value={form.topic} onChange={(event) => updateField('topic', event.target.value)} required />
          </label>

          <div className="form-row">
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
              <input type="number" min="4" max="15" value={form.slideCount} onChange={(event) => updateField('slideCount', event.target.value)} />
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

          <button disabled={isSubmitting}>{isSubmitting ? '创建中...' : '生成 PPT'}</button>
          {error ? <p className="error-text">{error}</p> : null}
        </form>

        <section className="result-panel">
          <div className="result-header">
            <div>
              <p className="eyebrow">生成结果</p>
              <h2>{task?.deck?.title || '还没有生成 PPT'}</h2>
            </div>
            {task?.status === 'done' ? <a className="download-link" href={task.downloadUrl}>下载 PPTX</a> : null}
          </div>

          {task ? <Progress task={task} /> : <p className="empty-state">填写左侧表单后开始生成，本地无模型配置也会使用兜底内容。</p>}

          <div className="slide-list">
            {slides.map((slide, index) => (
              <article className="slide-card" key={`${slide.title}-${index}`}>
                <div className="slide-number">{String(index + 1).padStart(2, '0')}</div>
                <div>
                  <span className="slide-type">{slide.type}</span>
                  <h3>{slide.title}</h3>
                  <ul>
                    {slide.bullets?.map((bullet) => <li key={bullet}>{bullet}</li>)}
                  </ul>
                </div>
              </article>
            ))}
          </div>
        </section>
      </section>
    </main>
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
      <div className="progress-track">
        <div style={{ width: `${task.progress || 0}%` }} />
      </div>
    </div>
  );
}

createRoot(document.getElementById('root')).render(<App />);
