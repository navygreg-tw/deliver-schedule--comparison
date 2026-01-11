<div align="center">
<img width="1200" height="475" alt="GHBanner" src="https://github.com/user-attachments/assets/0aa67016-6eaf-458a-adb2-6e31a0763ed6" />
</div>

# Smart Excel Comparer (AI Studio App)

這是一個基於 React 的 Excel 比較工具，使用 AI 輔助分析。

## 功能

- 上傳並比較 Excel 檔案
- 使用 Google Gemini API 進行 AI 分析
- 支援 Excel (.xlsx) 格式

## 本地開發 (Run Locally)

**前置需求:** Node.js (建議 v20+)

1. 安裝依賴套件:
   ```bash
   npm install
   ```

2. 設定環境變數:
   - 複製 `.env.local.example` (若有) 或建立 `.env.local`
   - 設定 `GEMINI_API_KEY`:
     ```
     GEMINI_API_KEY=your_api_key_here
     ```

3. 啟動開發伺服器:
   ```bash
   npm run dev
   ```

## 部署 (Deployment)

本專案使用 GitHub Actions 進行自動部署。

1. 推送程式碼到 GitHub。
2. 確保 GitHub Repository 的 Settings > Pages 設定中，Build and deployment source 設為 **GitHub Actions**。
3. 每次推送到 `main` 分支時，將會自動建置並部署到 GitHub Pages。

## 專案結構

- `src/`: 原始碼
- `public/`: 靜態資源
- `.github/workflows/`: GitHub Actions 設定檔

## 注意事項

- 請勿將 `.env` 或包含 API Key 的檔案上傳至 GitHub。
