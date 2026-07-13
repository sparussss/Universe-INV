# Universe Invoice PWA v0.6

功能測試版，主要修正：

- 選擇 Pictures Folder 後只顯示統計，不列出圖片清單。
- 修正相機重複啟動造成 `already under transition` 的問題。
- 圖片自動選擇會依據倉存表 `DESC2`，使用 `Stone List & Shape & Cutting.xlsx` 的 Breakdown → Quotation 代碼規則配對。
- 新加入貨品顯示在清單最上方，編號保留加入次序，例如 3、2、1。
- Invoice 貨品區為固定高度，可在區內上下捲動，適合超過 100 件貨。
- LOTNO 支援語音輸入；如瀏覽器不支援直接語音辨識，會開啟文字鍵盤供使用 iPhone 系統聽寫。
- Confirm Invoice 後匯出 Remaining Stock Excel。

## GitHub Pages 更新

把 ZIP 解壓後的所有檔案直接上傳到 Repository 根目錄，覆蓋舊版。若仍顯示舊版本，刪除主畫面舊 PWA，再由 Safari 重新加入主畫面。
