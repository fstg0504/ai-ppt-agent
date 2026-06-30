import cors from 'cors';
import dotenv from 'dotenv';
import express from 'express';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { nanoid } from 'nanoid';
import { buildDeckPlan } from './planner.js';
import { renderDeck } from './renderer.js';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const outputDir = path.resolve(__dirname, '../output');
const app = express();
const tasks = new Map();
const port = Number(process.env.PORT || 5174);

app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use('/downloads', express.static(outputDir));

// 健康检查接口 by AI.Coding：本地开发时快速确认后端服务是否可用。
app.get('/api/health', (request, response) => {
  response.json({ ok: true });
});

// 创建生成任务接口 by AI.Coding：立即返回 taskId，后台继续生成 PPT。
app.post('/api/decks', (request, response) => {
  const payload = normalizePayload(request.body);
  if (!payload.topic) {
    response.status(400).json({ error: '请填写 PPT 主题' });
    return;
  }

  const taskId = nanoid(10);
  const task = {
    taskId,
    status: 'queued',
    progress: 5,
    message: '任务已创建，等待生成',
    createdAt: new Date().toISOString(),
    modelMode: 'pending'
  };
  tasks.set(taskId, task);

  runDeckTask(taskId, payload).catch((error) => {
    updateTask(taskId, {
      status: 'failed',
      progress: 100,
      message: error.message || '生成失败'
    });
  });

  response.status(202).json(task);
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

// 任务执行方法 by AI.Coding：串联内容规划和 PPTX 渲染，保持接口层简洁。
async function runDeckTask(taskId, payload) {
  updateTask(taskId, { status: 'running', progress: 20, message: '正在生成 PPT 结构' });
  const planResult = await buildDeckPlan(payload);

  updateTask(taskId, {
    progress: 58,
    message: '正在渲染 PPTX 文件',
    deck: planResult.deck,
    modelMode: planResult.modelMode
  });

  const fileName = `${taskId}.pptx`;
  await renderDeck(planResult.deck, path.join(outputDir, fileName));

  updateTask(taskId, {
    status: 'done',
    progress: 100,
    message: '生成完成，可以下载 PPTX',
    downloadUrl: `/downloads/${fileName}`
  });
}

// 任务更新方法 by AI.Coding：集中合并任务状态，避免覆盖已有生成结果。
function updateTask(taskId, patch) {
  const current = tasks.get(taskId);
  if (!current) return;
  tasks.set(taskId, { ...current, ...patch, updatedAt: new Date().toISOString() });
}

// 入参清洗方法 by AI.Coding：限制页数范围，防止一次生成过大的本地任务。
function normalizePayload(body) {
  const slideCount = Math.min(Math.max(Number(body.slideCount || 8), 4), 15);
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
