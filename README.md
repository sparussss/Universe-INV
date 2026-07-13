# Universe Invoice PWA v0.4

本版重點：

- 手機選擇倉存 Excel
- 手機選擇客戶 Excel
- 選擇整個 Pictures Folder；圖片不複製進 PWA，只建立本次檔名索引
- 圖片檔名第一段配對 ARTNO，例如 `BG-34686 MG.JPG`
- 支援圖片變體及移除檔名末尾重複標記 `(1)`
- Barcode 掃描加入 1× / 2× / 3× / 4×（按裝置實際支援）
- Confirm Invoice 後，移除本次 LOTNO 並自動匯出 `Remaining_Stock_...xlsx`
- 已 Confirm 的 LOTNO 在同一工作階段再次掃描會顯示已售 Invoice

## GitHub Pages 更新

將本資料夾內所有檔案上傳到 Repository 根目錄，覆蓋舊版本。不要多包一層資料夾。

若仍看到舊版：

1. 在 Safari 開啟網站並重新整理。
2. 刪除主畫面的舊 PWA 圖示。
3. 再由 Safari「加入主畫面」。

## 圖片 Folder

圖片必須已經下載在「我的 iPhone／我的 iPad」或可離線讀取的位置。重新開啟 PWA 後，iOS 可能要求再次選擇 Pictures Folder。
