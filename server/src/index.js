import cors from 'cors';
import dotenv from 'dotenv';
import express from 'express';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { nanoid } from 'nanoid';
import { buildDeckPlan } from './planner.js';
import { renderDeck } from './renderer.js';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const outputDir = path.resolve(__dirname, '../output');
const historyFile = path.join(outputDir, 'tasks.json');
const app = express();
const tasks = new Map();
const taskQueue = [];
let isQueueRunning = false;
const port = Number(process.env.PORT || 5174);

await loadTaskHistory();

app.set('etag', false);
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use('/downloads', express.static(outputDir));

// API 缓存控制中间件 by AI.Coding：任务状态需要实时轮询，禁止浏览器用 304 复用旧响应。
app.use('/api', (request, response, next) => {
  response.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  response.set('Pragma', 'no-cache');
  response.set('Expires', '0');
  next();
});

// 健康检查接口 by AI.Coding：本地开发时快速确认后端服务是否可用。
app.get('/api/health', (request, response) => {
  response.json({ ok: true });
});

// 创建生成任务接口 by AI.Coding：支持批量创建任务，后台按队列顺序生成 PPT。
app.post('/api/decks', (request, response) => {
  const payload = normalizePayload(request.body);
  if (!payload.topic) {
    response.status(400).json({ error: '请填写 PPT 主题' });
    return;
  }

  const batchCount = Math.min(Math.max(Number(request.body.batchCount || 1), 1), 10);
  const createdTasks = Array.from({ length: batchCount }, (_, index) => createQueuedTask(payload, index + 1, batchCount));
  persistTaskHistory();
  drainQueue();

  response.status(202).json({ tasks: createdTasks });
});

// 查询任务列表接口 by AI.Coding：前端轮询该接口展示历史记录和排队进度。
app.get('/api/tasks', (request, response) => {
  response.json({ tasks: getTaskList() });
});

// 查询任务接口 by AI.Coding：前端根据 taskId 轮询生成进度和结果。
app.get('/api/tasks/:taskId', (request, response) => {
  const task = tasks.get(request.params.taskId);
  if (!task) {
    response.status(404).json({ error: '任务不存在' });
    return;
  }

  response.json(task);
});

// 创建排队任务方法 by AI.Coding：统一生成任务元信息并写入执行队列。
function createQueuedTask(payload, batchIndex, batchTotal) {
  const taskId = nanoid(10);
  const task = {
    taskId,
    status: 'queued',
    progress: 5,
    message: batchTotal > 1 ? `批量任务 ${batchIndex}/${batchTotal} 已排队` : '任务已创建，等待生成',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    modelMode: 'pending',
    payload: { ...payload },
    batchIndex,
    batchTotal
  };
  tasks.set(taskId, task);
  taskQueue.push(taskId);
  return task;
}

// 队列执行方法 by AI.Coding：串行调用模型，避免批量生成时并发压垮本地或模型服务。
async function drainQueue() {
  if (isQueueRunning) return;
  isQueueRunning = true;

  while (taskQueue.length) {
    const taskId = taskQueue.shift();
    const task = tasks.get(taskId);
    if (!task || task.status !== 'queued') continue;

    try {
      await runDeckTask(taskId, task.payload);
    } catch (error) {
      updateTask(taskId, {
        status: 'failed',
        progress: 100,
        message: error.message || '生成失败'
      });
    }
  }

  isQueueRunning = false;
}

// 任务执行方法 by AI.Coding：串联内容规划和 PPTX 渲染，保持接口层简洁。
async function runDeckTask(taskId, payload) {
  updateTask(taskId, { status: 'running', progress: 20, message: '正在生成 PPT 结构' });
  const planResult = await buildDeckPlan(payload);

  updateTask(taskId, {
    progress: 58,
    message: '正在渲染 PPTX 文件',
    deck: planResult.deck,
    modelMode: planResult.modelMode,
    planningWarning: planResult.planningWarning || ''
  });

  const fileName = `${formatTimestamp(new Date())}-${sanitizeFilePart(planResult.deck.title || payload.topic)}.pptx`;
  await renderDeck(planResult.deck, path.join(outputDir, fileName));

  updateTask(taskId, {
    status: 'done',
    progress: 100,
    message: '生成完成，可以下载 PPTX',
    fileName,
    downloadUrl: `/downloads/${encodeURIComponent(fileName)}`
  });
}

// 任务更新方法 by AI.Coding：集中合并任务状态，避免覆盖已有生成结果。
function updateTask(taskId, patch) {
  const current = tasks.get(taskId);
  if (!current) return;
  tasks.set(taskId, { ...current, ...patch, updatedAt: new Date().toISOString() });
  persistTaskHistory();
}

// 任务列表方法 by AI.Coding：按创建时间倒序返回，便于前端展示历史生成记录。
function getTaskList() {
  return Array.from(tasks.values()).sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
}

// 时间文件名方法 by AI.Coding：用生成完成时间命名 PPT，避免随机字符串不便区分。
function formatTimestamp(date) {
  const pad = (value, size = 2) => String(value).padStart(size, '0');
  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
    '-',
    pad(date.getHours()),
    pad(date.getMinutes()),
    pad(date.getSeconds()),
    '-',
    pad(date.getMilliseconds(), 3)
  ].join('');
}

// 文件名清洗方法 by AI.Coding：保留可读标题，同时移除跨平台文件名非法字符。
function sanitizeFilePart(value) {
  const cleaned = String(value || 'ai-ppt')
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 48);
  return cleaned || 'ai-ppt';
}

// 历史加载方法 by AI.Coding：服务启动时从 output/tasks.json 恢复已生成记录。
async function loadTaskHistory() {
  await fs.mkdir(outputDir, { recursive: true });
  try {
    const content = await fs.readFile(historyFile, 'utf8');
    const history = JSON.parse(content);
    if (Array.isArray(history.tasks)) {
      history.tasks.forEach((task) => {
        const restoredTask = task.status === 'queued' || task.status === 'running'
          ? { ...task, status: 'failed', progress: 100, message: '服务重启后任务已中断，请重新生成。' }
          : task;
        tasks.set(task.taskId, restoredTask);
      });
      persistTaskHistory();
    }
  } catch (error) {
    if (error.code !== 'ENOENT') console.warn('Failed to load task history:', error.message);
  }
}

// 历史保存方法 by AI.Coding：用 JSON 文件替代数据库保存本地生成记录。
async function persistTaskHistory() {
  try {
    await fs.mkdir(outputDir, { recursive: true });
    await fs.writeFile(historyFile, JSON.stringify({ tasks: getTaskList() }, null, 2));
  } catch (error) {
    console.warn('Failed to persist task history:', error.message);
  }
}

// 入参清洗方法 by AI.Coding：限制页数范围，防止一次生成过大的模型任务。
function normalizePayload(body) {
  const slideCount = Math.min(Math.max(Number(body.slideCount || 8), 1), 15);
  return {
    topic: String(body.topic || '').trim(),
    deckType: String(body.deckType || 'business'),
    audience: String(body.audience || '').trim(),
    notes: String(body.notes || '').trim(),
    slideCount
  };
}

app.listen(port, () => {
  console.log(`AI PPT Agent server listening on http://localhost:${port}`);
});
