# AI PPT Agent

本项目是一个本地开发版 AI PPT 生成器：React 前端负责输入主题和展示生成结果，Node 后端负责生成结构化 PPT 计划并导出 `.pptx` 文件。

## 本地启动

```bash
npm install
cp .env.example server/.env
npm run dev
```

前端默认地址：`http://localhost:5173`

后端默认地址：`http://localhost:5174`

## 首版能力

- 输入主题、PPT 类型、页数和补充材料
- 后端生成结构化幻灯片计划
- 必须配置 `LLM_API_KEY` 后调用真实文本模型生成计划
- 文本模型直接输出每页布局 blocks，决定元素位置、大小和样式
- 文本模型输出会经过 JSON 提取、结构清洗和基础校验
- 使用 PptxGenJS 输出可编辑 `.pptx`
- 支持设置批量生成数量，后端按队列顺序生成
- 前端基于同一份 AI blocks 做网页预览，不额外消耗模型 token
- 前端展示任务历史、排队进度、幻灯片预览和下载入口
- 历史记录保存在 `server/output/tasks.json`，PPTX 文件保存在 `server/output/`

## 文本模型配置

后端兼容 OpenAI 风格的 `chat/completions` 接口。复制 `.env.example` 到 `server/.env` 后配置：

```env
LLM_API_KEY=你的文本模型 Key
LLM_BASE_URL=https://api.openai.com/v1
LLM_MODEL=gpt-5.5
LLM_TIMEOUT_MS=60000
LLM_TEMPERATURE=0.4
```

如果未配置模型、模型请求失败、超时、返回空内容或 JSON 不合法，任务会失败并在前端展示原因，不会使用本地模板兜底生成内容。

## 预览方式

当前预览不再次调用 AI。文本模型生成一次 `slides[].blocks` 后：

- 前端用 blocks 的百分比坐标渲染 HTML 预览。
- 后端用同一份 blocks 通过 PptxGenJS 导出 PPTX。

这种方式最省 token，且能在下载前看到整体排版。后续如果需要和 PowerPoint 完全一致的预览，可以再增加 LibreOffice Headless，把 PPTX 转成 PNG 预览图。

## 本地历史记录

当前版本不依赖数据库。后端会把任务历史写入：

```text
server/output/tasks.json
```

生成的 PPTX 文件会写入：

```text
server/output/
```

文件名使用生成完成时间和 PPT 标题，例如：

```text
20260701-103012-125-AI办公自动化方案.pptx
```

`server/output/tasks.json` 和生成的 `.pptx` 已被 `.gitignore` 忽略，不会提交到仓库。

## 后续建议

- 接入 `gpt-image2` 生成封面图、插图和背景图
- 增加 ECharts/Mermaid 图表图片生成
- 增加 LibreOffice 预览转图
- 增加模板库和单页重生成能力
