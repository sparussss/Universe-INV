# Universe Invoice PWA v0.14.8

## v0.14.8 Records 健康檢查與更新快取修正

- 正式文件記錄唯一來源仍是 `jmsdata.xlsx` 的隱藏工作表 `Universe Records`。
- 新展覽的 `jmsdata.xlsx` 可以只有原本庫存工作表；沒有 `Universe Records` 屬正常情況，不再當成資料包錯誤。
- 沒有 `Universe Records` 時會顯示「新展覽」正常提示：Recall 0、Invoice／Consignment／Quotation 流水號由 0001 開始；第一次 Confirm 後自動建立隱藏的第二個工作表。
- 資料包完整性／健康檢查不再要求任何獨立 Records JSON 檔案。
- 加強 PWA 更新快取：本版使用版本化 app script 檔名，Service Worker 對本機 PWA 程式改為 network-first、離線時才 fallback cache，降低 iPhone 出現新版頁首配舊版 JavaScript 的機會。
- 其餘 v0.14.7 新展覽 Records 強制隔離、草稿 Session 隔離及 Recall Exhibition ID 篩選維持不變。
