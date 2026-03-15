import { defineConfig } from 'astro/config';
import react from '@astrojs/react';
import vercel from '@astrojs/vercel/serverless';

export default defineConfig({
  output: 'hybrid',   // 頁面預設靜態，API routes 加 prerender=false 啟用 SSR
  adapter: vercel(),
  integrations: [react()],
});
