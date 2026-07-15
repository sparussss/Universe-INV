# Universe Invoice PWA v0.10.3

## v0.10.3 Excel output changes

- Header layout remains unchanged.
- Invoice Date is written as fixed English text, for example `15 July, 2026`, so PDF conversion will not localise the month into Chinese.
- LOTNO and ARTNO use normal font weight.
- A non-zero Discount displays in brackets, for example `($100.00)`; zero displays as `$0.00`.
- Smart pagination calculates every item height before writing the workbook.
- Each item remains an indivisible block; page breaks are inserted before a complete item.
- The final page reserves enough room for the complete footer. When necessary, the footer moves to its own page.
- Maximum 10 items per page remains a safety limit.


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


## v0.10.2 修正
- 修正相機已支援 1×–10×，但 2×、3×、4× 按鈕仍被停用的問題。
- 不再依賴 html5-qrcode 的 isScanning 屬性，改由 PWA 自行追蹤相機運作狀態。
- 切換倍率後重新讀取實際相機 Zoom 設定。


## v0.10.2
- 修正 iPhone 小螢幕表單、Invoice 貨品控制及 Invoice 預覽橫向溢出。
- PDF 預覽表格使用固定欄寬及自動換行。


## v0.10.2
- Invoice 預覽頁新增「匯出 Excel Invoice」
- Excel 包含公司抬頭、客戶資料、款號圖片、LOTNO、DESC1–DESC6、Qty、Unit、Unit Price、Amount、Subtotal、Discount、Total
- 圖片會按 Invoice 當時選擇的版本嵌入 xlsx


## v0.10.2
- 可匯入 Stone List & Shape & Cutting.xlsx，按 BREAKDOWN → QUOTATION 更新圖片自動選擇。
- 可匯入 Invoice .xlsx 範本；匯出 Excel Invoice 時優先套用範本。
- 未匯入範本時仍保留原有標準 Excel 輸出。


## v0.10.2
- 移除 PWA 直接輸出／列印 PDF 功能。
- 新增 Article Mapping.xlsx 匯入。
- Invoice 範本輸出時，ARTNO 前綴自動轉為 Article Description。
- 範本圖片移到原有右側圖片區（G:I），不再遮蓋 LOTNO／ARTNO。
- 圖片保持原比例，不拉伸、不裁切。


## v0.10.2
- 六個資料匯入區塊可手動展開／收合。
- 匯入成功後自動收合，只保留檔名及匯入結果。
- 匯入失敗時保持展開，方便重新選擇檔案。


## v0.10.2 Excel Template 3(2) dynamic item layout

- Each item uses at least five content rows.
- ARTICLE and non-empty DESC1–DESC6 expand the item automatically when needed.
- One empty separator row is always inserted after each item.
- Product image is restricted to the first fixed five rows and keeps its aspect ratio.
- Footer is moved below the dynamic item area.
- Item blocks and footer are given page-break protection for Excel-to-PDF output.


## v0.10.2
- 修正匯出 Excel 時 `numberToWords` 未定義的錯誤。
- Total Amount 可輸出英文大寫金額。
- 繼續直接使用已匯入的 Invoice Template，不重新建立範本。


## v0.10.2 Template Map driven export
- Reads the imported `Template Map` sheet for header, item columns, image area and footer cells.
- Supports Invoice Master Template 3(3).xlsx without hard-coding the previous cell positions.
- Removes the Template Map sheet from the exported customer invoice.


## v0.10.2 Excel 輸出更新
- Article Mapping 變為選用；未匯入時不顯示 Article。
- 所有 Invoice 儲存格垂直置中並關閉自動換行。
- 圖片限定在 Template Map 指定的 2 欄 × 5 行圖片區，保持比例、置中、不超界。
- 每款最少 5 行，有更多 DESC 時自動增行，之後固定加 1 行空白。
- 分頁以實際列高計算，整個 Item 不拆頁，簡單款每頁可自然容納更多項目。
- Footer 不拆頁。
- 紙張預設 A4 直向，Fit to Width 1 page。


## v0.10.2 final field-test changes

- Each printed A4 page holds at most 10 items.
- Actual row height is still checked; an item that does not fit moves wholly to the next page.
- Item blocks and the final totals/footer are not intentionally split across pages.
- Existing images embedded in the imported Invoice Template, including a letterhead image, are preserved.
- Product images remain contained inside the Template Map image range.


## v0.10.2 Excel layout changes

- Minimum 4 content rows per item.
- Extra DESC rows are added only when required.
- One 10.5 pt separator row after every item.
- All item content rows are fixed at 10.5 pt.
- Product images are contained within the first 4 rows of the D:E image area.
- Manual page breaks are placed before the next item (on the preceding row), preventing a single item line from being left at the bottom of a page.
- Footer page break uses the same before-block logic.
