# 續訂回覆合併台

這是一個可部署到 GitHub Pages 的純前端工具。它從 Google Drive 讀取業務回覆 Excel，並在瀏覽器記憶體中更新使用者選擇的本機 Excel 複本。

正式網站：<https://wilson100043-byte.github.io/renewal-reply-merge/>

## 使用者流程

1. 貼上 Google Drive Excel 連結並登入 Google。
2. 選擇月份分頁。
3. 選擇本機 `未績訂` Excel。
4. 預覽依 `編號細目` 配對的變更。
5. 下載新版 Excel；原檔不覆蓋。

合併規則：

- 本月回覆空白：沿用上個月回覆。
- 上個月回覆空白：使用加上年月的本月回覆，例如 `202606:下半年才會辦理`。
- 兩邊都有：`YYYYMM:本月回覆;上個月回覆`。
- 兩邊都空白：不清除本機既有內容。

年月會優先從月份分頁讀取；分頁只有「6月」時，年份取自來源檔名（例如 `未續訂清單_知文_2026.xlsx`）。若檔名沒有年份，才使用執行當年的年份。

## 隱私設計

- Excel 只在瀏覽器記憶體中處理。
- 不設後端、不上傳 Excel、不將資料寫入 GitHub。
- Google Access Token 只保留於目前頁面記憶體，不寫入 `localStorage`。
- `localStorage` 只保存 Drive 連結與月份分頁名稱。
- Repository 不得放入公司資料、真實測試檔、Token 或密鑰。

## 本機啟動

需要 Python 3；它只用來提供靜態網頁，不會接收或處理 Excel。

```bash
python3 serve.py
```

瀏覽器開啟 `http://127.0.0.1:8000`。

不登入 Google 時，可以展開「改用已下載的雲端檔」，用本機來源檔測試完整合併流程。

## Google OAuth 設定

1. 在 Google Cloud 建立或選擇專案。
2. 啟用 Google Drive API。
3. 設定 OAuth consent screen。若公司使用 Google Workspace，優先設為 Internal。
4. 建立 OAuth Client，Application type 選擇 `Web application`。
5. 在 Authorized JavaScript origins 加入：
   - 本機測試：`http://localhost:8000` 與 `http://127.0.0.1:8000`
   - 正式站：`https://<github-account>.github.io`
6. 將 Client ID 填入 `config.js` 的 `googleClientId`。

`googleClientId` 是公開識別碼，不是 Client Secret。不要把 Client Secret 放進前端程式。

此原型使用 `drive.readonly`，因為來源是使用者既有的私人 Drive 檔案。正式上線前應由公司管理者確認 OAuth scope 與同意畫面政策。

## GitHub Pages 部署

此專案不需要建置步驟。Repository 的 Pages 設定可選擇從 `main` branch 根目錄部署。

一般 GitHub Pages 網址預設對網際網路公開。公開的只有工具程式；私人 Drive 檔仍需要 Google 授權才能下載。若工具程式本身也必須限制公司內部存取，需要 GitHub Enterprise Cloud 私有 Pages 或其他有登入保護的主機。

## Excel 更新策略

`.xlsx` 是 ZIP 封裝。本工具只重寫目標工作表 XML 中需要修改的 `上個月回覆`儲存格，其他 ZIP 內容保持相同位元組，避免重建整份 workbook 造成樞紐分析、格式或關聯遺失。

更新前會阻擋以下情況：

- 缺少必要工作表或欄位。
- 雲端同一 `編號細目` 出現不同合併結果。
- 目標儲存格含公式。
- 雲端編號在本機找不到。

本機相同 `編號細目` 有多列時，會更新所有匹配列。

## 第三方元件

- `fflate 0.8.2`：MIT License，用於瀏覽器內解壓縮與重新封裝 `.xlsx`。
- Google Identity Services：由 Google 官方網域在執行時載入，用於使用者登入與 Drive 唯讀授權。

## 原型限制

- 已部署至 GitHub Pages，並配置 External／Testing Google OAuth Client ID；只有列在 Google Cloud Test users 的帳號可以直接讀取 Drive 來源檔。
- Access Token 到期後，使用者需要再次按登入按鈕。
- GitHub Pages 無排程與後端，更新必須由使用者開啟網站後執行。
- Windows 與 Mac 使用相同瀏覽器程式碼；仍需以 Windows Excel 實機開啟輸出檔完成正式驗收。
