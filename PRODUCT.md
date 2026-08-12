# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

公司內部使用 Chrome 的同仁，需要把 Google Drive 上的業務續訂回覆合併進本機 Excel 原檔。

## Product Purpose

讓使用者先預覽可更新與異常資料，再安全產生更新後的 Excel。成功代表使用者能確認修改內容、保留可還原的原始資料，並取得可繼續使用的更新檔。

## Operating Context

來源是 Google Drive 上的 `.xlsx`，目標是使用者在第二步選擇的本機「未績訂」Excel。工具在瀏覽器記憶體中處理檔案，不上傳內容至後端。

## Capabilities and Constraints

- 依 `編號細目` 配對來源與目標資料，同步更新 `上個月回覆`與`續訂/停訂`；來源狀態空白時不清除本機值，本機 `續訂/停訂` 已標記為 `OK` 時保留本機值。
- 預設輸出方式是下載一份新版 Excel，不覆寫原檔。
- Chrome 使用者可選擇「下載備份並覆寫原檔」；覆寫前必須再次確認。
- 備份必須在覆寫前完整讀入為獨立的記憶體快照，不能繼續引用即將被改寫的磁碟檔案。流程穩定後，可由產品決策移除強制備份，但目前不可省略。
- 覆寫只作用於本機目標檔，不修改 Google Drive 來源檔。
- 不支援直接寫檔的瀏覽器必須保留下載新版的安全流程。

## Product Principles

- 先預覽，再寫入。
- 可逆操作優先，任何覆寫都有明確備份。
- 危險操作需要清楚命名與再次確認。
- 不把 Excel、Token 或客戶資料送往非必要服務。

## Accessibility & Inclusion

所有主要操作必須可用鍵盤完成，並提供清楚的成功、錯誤、停用與取消回饋。
