# Universe Invoice PWA v0.14.12

## v0.14.12 移除「匯出更新後資料包」

- 完全移除「匯出更新後資料包」按鍵及 `exportUpdatedPackage()` 程式。
- 移除只為此功能使用的 JSZip 依賴及 Service Worker 離線快取項目。
- 「款式搜尋」頁不再提供整個展覽 Folder 的 ZIP 輸出。
- Confirm Invoice／Consignment／Quotation 仍會輸出最新 `jmsdata.xlsx`；正式文件記錄繼續保存在 Sheet 2「Universe Records」。
- 保留「資料匯入」頁的「匯出目前 jmsdata.xlsx」作手動備份／補輸出。
- 手動選圖記錄會寫入最新 `jmsdata.xlsx`；新上傳／拍攝圖片檔本身不會嵌入 Excel，因此相關提示改為只說明目前 PWA 暫存狀態，不再提示匯出資料包。
- 保留 v0.14.11 的「款式搜尋」名稱、新展覽 Records 隔離及 cache-busting / network-first 更新機制。
