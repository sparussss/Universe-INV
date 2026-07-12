# Universe Invoice PWA Prototype

這是可放到 GitHub Pages 的原型版，不需要 Mac、Xcode 或 App Store。

## 已完成的原型功能

- 匯入每次展覽的 `.xls` / `.xlsx` 倉存表
- Barcode = `LOTNO`
- 產品圖片以 `ARTNO` 檔名共用
- 支援 `PT-37499.jpg`、`PT-37499 AM.jpg`、`PT-37499 CT.jpg`
- 無後綴圖為預設，可逐件切換圖片版本
- 客戶、地址、Currency、Sales Rate
- `Unit Price = ceil(PRICE × Sales Rate)`
- 手動輸入 LOTNO
- 相機 Barcode 掃描（視 Safari / iOS 支援而定）
- 新增、指定位置插入、刪除、拖拉排序
- 修改 Quantity、Unit Price、Discount Amount
- Invoice 預覽
- 使用 iOS 列印功能「儲存為 PDF」
- Service Worker 基本離線外殼

## 立即測試

內置示範資料：

- LOTNO：`133685`
- ARTNO：`PT-37499`
- U價：`3092`
- Sales Rate：`0.34`
- 計算：`ceil(3092 × 0.34) = 1052`

在「建立 Invoice」頁手動輸入 `133685` 即可測試。

## GitHub Pages 上傳

1. 在 GitHub 建立一個新 Repository，例如 `universe-invoice-pwa`。
2. 將這個資料夾內所有檔案上傳到 Repository 根目錄。
3. 進入 Repository 的 `Settings` → `Pages`。
4. `Build and deployment` 選擇 `Deploy from a branch`。
5. Branch 選 `main`，Folder 選 `/ (root)`，按 `Save`。
6. 等待約一至數分鐘，GitHub 會提供網址。
7. 用 iPhone / iPad Safari 開啟網址，按分享 →「加入主畫面」。

## 重要限制

- Excel 解析使用 SheetJS CDN。首次開啟及首次載入解析程式時需要網絡；載入後瀏覽器通常會快取。正式版應把 SheetJS 檔案直接放入專案，做到完全離線。
- Safari 的 BarcodeDetector 支援因 iOS 版本而異，因此手動 LOTNO 永遠保留。
- 目前資料主要存在頁面記憶體；重新整理會清除 Invoice 草稿及已匯入圖片。正式版下一步應加入 IndexedDB 永久保存。
- 目前 PDF 使用瀏覽器列印功能，版面已作 A4 原型，但未完全逐像素複製正式 Invoice。

## 下一階段建議

1. IndexedDB 保存客戶、圖片、展覽、產品、Invoice 草稿及歷史。
2. 將 SheetJS 本地化，完整離線匯入 `.xls`。
3. 完整複製正式 Sales Invoice PDF 的分頁及版面。
4. 增加客戶 Excel 匯入。
5. 加入展覽清單、封存及 Invoice History。
