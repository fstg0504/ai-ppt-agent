import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Vite 配置 by AI.Coding：前端开发时把 API 请求代理到本地 Node 服务。
export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      '/api': 'http://localhost:5174',
      '/downloads': 'http://localhost:5174'
    }
  }
});
