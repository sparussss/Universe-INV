# Universe Invoice PWA v0.7

此 ZIP 固定包含 9 個檔案：

1. index.html
2. styles.css
3. app.js
4. manifest.webmanifest
5. sw.js
6. icon-192.png
7. icon-512.png
8. icon.svg
9. README.md

## 主要功能

- 手機選擇倉存 Excel、客戶 Excel、Pictures Folder
- Pictures Folder 只建立當次索引，不複製圖片
- Customer Code 自動移除空格；Column L 空白時 Sales Rate 使用 0.34
- 手動、語音、Barcode 掃描三種 LOTNO 輸入
- 掃描預設 3×；支架連續掃描
- 成功顯示綠色大 LOTNO；重複橙色；找不到紅色
- DESC2 配合 Stone List 對照自動選圖片
- 無 (1) 的圖片優先；圖片完整顯示不裁切
- 最新貨品在最上方；固定高度清單內捲動
- 新增、插入、刪除、復原、改 Qty／Unit Price／圖片版本
- Confirm 後輸出可編輯 Invoice XLSX 及 Remaining Stock XLSX

## GitHub Pages

解壓後把以上 9 個檔案直接放到 Repository 根目錄並覆蓋舊檔案。不要多包一層資料夾。

第一次使用需有網絡載入 Excel 與 Barcode 程式；成功載入後 Service Worker 會嘗試快取。相機必須在 HTTPS 網址使用。
