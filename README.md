# Universe Invoice PWA v0.9

UI 改善：
- 移除頂部說明句。
- 展覽名稱預設為 Jewellery Show。
- 移除示範貨品、圖片 Folder 說明及示範 LOTNO。
- 客戶欄位預設留空。
- Invoice Date 使用 YYYY-MM-DD 文字格式。
- 「加入貨品」移到客人 Invoice 資料的同一組別內。

把所有檔案直接上傳到 GitHub Repository 根目錄。


## v0.9
- Invoice No. 自動使用 INV + YY + 4位流水號，由 0001 開始。
- Confirm 後才增加流水號；草稿不消耗號碼。
- LOTNO 先使用 Numpad（數字鍵盤）作實機測試；是否顯示咪高峰由 iPhone 系統鍵盤決定。
- 語音結果會自動移除空格、標點，並把中文數字一至九轉成阿拉伯數字。
- 移除拍照／選照片掃碼按鈕。

- 建立 Invoice 頁面的貨品區標題改為「輸入模式」。
- Barcode 掃描優先鎖定後鏡頭；如相機名稱無法識別，會以 environment 模式再嘗試。


## v0.9 改善
- 客戶輸入至少 2 個字元後自動彈出最多 10 個符合結果。
- Barcode 掃描使用後鏡頭，並加入真正相機 1× / 2× / 3× / 4× 控制；裝置不支援的倍率會停用。
- LOTNO 改用一般 iPhone 鍵盤，可使用系統咪高峰聽寫。
- 加入或輸入錯誤後，自動回到 LOTNO 輸入框。
