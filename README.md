# 成人健檢報告產生程式（網頁版）

純靜態 Astro 網頁，所有處理在瀏覽器端完成，部署於 Cloudflare Pages 免費方案。

## 功能
- 上傳 Excel 健檢資料
- 上傳 Word (.docx) 報告範本
- 預覽個案列表（可搜尋、排序、篩選代謝症候群）
- 點選個案產生 docx 報告（單份或批次下載為 ZIP）

## 本機開發
```bash
npm install
npm run dev
```

## 建置
```bash
npm run build
# 輸出於 dist/
```

## Cloudflare Pages 部署設定
| 項目 | 值 |
|------|-----|
| Build command | `npm run build` |
| Build output directory | `dist` |
| Node.js version | `20` |
