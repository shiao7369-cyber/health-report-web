import { defineConfig } from 'astro/config';
import react from '@astrojs/react';
import vercel from '@astrojs/vercel';   // v8 不再有 /serverless 子路徑

export default defineConfig({
  output: 'static',  // Astro v5：static + adapter = 頁面靜態，API routes 加 prerender=false
  adapter: vercel(),
  integrations: [react()],
  vite: {
    build: {
      rollupOptions: {
        output: {
          // 將重量級函式庫拆成獨立 chunk，改善快取與載入
          manualChunks(id) {
            if (id.includes('node_modules/xlsx')) return 'xlsx';
            if (id.includes('node_modules/jszip') || id.includes('node_modules/docx-preview')) return 'docx';
          },
        },
      },
    },
  },
});
