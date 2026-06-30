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
- 无模型配置时使用本地兜底计划，保证流程可跑
- 配置 `LLM_API_KEY` 后可调用文本模型生成计划
- 使用 PptxGenJS 输出可编辑 `.pptx`
- 前端展示任务状态、页面概要和下载入口

## 后续建议

- 接入 `gpt-image2` 生成封面图、插图和背景图
- 增加 ECharts/Mermaid 图表图片生成
- 增加 LibreOffice 预览转图
- 增加模板库和单页重生成能力
