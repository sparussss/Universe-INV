# Universe Invoice PWA v0.11.28

## v0.11.28 MULTI image matching for five-or-more stones

- 使用更新版 `Stone List & Shape & Cutting.xlsx` 的 BREAKDOWN 代號辨識每件貨品的石種。
- 同一貨品在 DESC2–DESC6 合共有 **5 個或以上不同的非鑽石石頭代號**時，自動把圖片目標判定為 `MULTI`。
- `CDM`／`DIA` 不計入 MULTI 門檻；重複石頭代號只計一次，Shape、Cutting、尺寸亦不計。
- 有 `MULTI` 圖時會優先選取，並在同類候選中盡量配對原始 18K 成色；沒有 `MULTI` 圖時沿用同款參考圖並自動轉黑白。
- 此共用選圖邏輯同步套用於 Invoice、Consignment、Quotation 及配套搜尋；手動圖片／黑白選擇仍保留。


## v0.11.26 strict Stone List filters + status filtering

- 石頭篩選只使用匯入 Stone List 的 BREAKDOWN 代號，並依表內次序顯示。
- 修正 `GGTRD` 誤判為 `GGTR`、`PSARD` 誤判為 `SAR`：現在會辨識為 `GGT`、`PSA`。
- `SKY BT` 的石頭篩選按鈕簡化顯示為 `SKY`。
- 石頭／款式按鈕改為自動寬度及換行，不再超出按鈕或橫向擠迫。
- `Avail / Consign / Sold-OH / Deliv` 改為可多選的狀態篩選按鈕；沒有選擇狀態時顯示全部。


## v0.11.25 wildcard all-stock search + compact filters

- 配套搜尋輸入 `*` 可顯示全部庫存／歷史貨品。
- `*` 結果排序：核心款號數字由大至小 → 石頭依 Stone List 順序 → 同石頭內依 `RG → ER → PT → BR → NL → BL → BG → 其他`。
- 款式／石頭多選篩選改成可收合區；搜尋後預設收起，只顯示一行目前篩選摘要，按「展開篩選」才顯示 chips。
- 輸入 `*` 時同樣可使用既有款式多選、石頭多選、庫存狀態與加入目前文件功能。


## v0.11.24 unified image selection + single edit + tab order

- **石頭優先、成色次選**：圖片石種評分會先移除 `(18KY)` / `(18KR)` / `(18KW)` / `(reg)` / 編號等括號 metadata，再比較石種。因此 `BT`、`BT (18KR)`、`BT (18KY)` 會先視為同一個 BT 石種，然後以原始 DESC1 的成色選最合適版本；例如 `1.20Y750` 會優先 `BT (18KY)`。
- **黑白手動選項**：每款編輯的圖片下拉選單新增 `圖片：黑白`。自動找不到對應石種時仍會自動使用同 ARTNO 參考圖黑白顯示；手動亦可強制把目前資料庫款式圖切成黑白。Preview / Excel / PDF 沿用相同黑白狀態。
- **同一時間只展開一款編輯**：打開另一款「編輯」時，上一款會自動收起；修改仍即時保留。任何一款正在編輯時，全部 ☰ 維持灰色停用，全部收起後才恢復拖曳排序。
- **四個功能共用同一自動選圖核心**：Invoice、Consignment、Quotation、配套搜尋都使用 `chooseImageMatch()`；14K Quotation 仍以原始 18K DESC1 判斷成色。
- **頁面順序**：資料匯入 → 建立文件 → 配套搜尋 → 文件預覽。


## v0.11.23 stone-first image fallback

- 圖片自動配對以**石頭為第一優先**。完全／組合石種匹配成功時維持彩色圖片；同石種有多張候選時，再盡量按原始 18K DESC1 的 Y / W / R 成色對應 `18KY / 18KW / 18KR`。
- 若同 ARTNO **完全找不到對應石種圖片**，PWA 會選同款最合適的參考圖並自動以**黑白**顯示，避免以錯誤石色誤導客人。
- 黑白 fallback 同步套用到建立文件縮圖、文件預覽、配套搜尋，以及 Excel / PDF 輸出；Pictures 原始圖片檔不會被修改。
- `CDM` 不參與主圖片石種選擇。14K Quotation 的圖片成色仍依原始庫存 DESC1 判斷，而不是 585 轉換後文字。

## v0.11.22 combined-stone image matching

- 圖片變體配對加入「包含石種」邏輯：若 DESC2 的主石沒有完全同名圖片，但存在組合圖片，例如 `MCT+CT+LQZ`，系統會把它視為可匹配 `MCT`、`CT`、`LQZ` 的圖片。
- 配對優先順序維持：完全同名變體 → 包含主石的組合變體（優先較少組合項）→ Default／第一張圖片。
- 下拉選單仍顯示實際檔名變體，不建立不存在的虛擬 `MCT.JPG`。

## v0.11.21 quotation / companion search / drag refinement

- **14K Quotation / USD**：自動計算的 14K USD Unit Price 改為向上進位至整數美元，不保留尾數；例如 `$593.39 → $594.00`、`$1,191.82 → $1,192.00`。14K 模式下手動修改 USD Unit Price 亦會向上進位。
- **配套搜尋圖片**：每個 LOTNO 會沿用建立文件時的石種／圖片變體配對邏輯，依 DESC2 與 Stone List alias 選擇相應圖片；不再固定顯示該 ARTNO 的第一張圖片。
- **iPhone 拖曳排序**：加入 300ms 長按、10px 移動門檻，降低換位敏感度，並把自動捲動觸發範圍與速度降低。原有「編輯展開時全部 ≡ 灰掉停用」規則保留。

## v0.11.20 item list interaction revision

- 「刪除」固定顯示在每件貨品圖片下方，不再需要先按「編輯」。
- 右側 ≡ 改為 iPhone 觸控拖曳排序；拖曳後會重建正式 Item 1 → N 次序，Preview / Excel / PDF 跟隨新順序。
- 「編輯」與排序互斥：只要任何貨品的編輯區展開，全部 ≡ 會變灰並暫停拖曳；全部收起後才恢復排序。
- 拖曳只由 ≡ 觸控區啟動，避免一般上下捲動貨品清單時誤排序。





## v0.11.19 FX pricing revision

- EUR reference FX source is changed back to **Frankfurter v2**. Frankfurter v1 remains a fallback if v2 cannot be reached; cached last-known rates and manual FX input remain available.
- With an **automatic Frankfurter** rate (online or cached), EUR Unit Price drops the cents instead of rounding: e.g. `€637.92 → €637.00`.
- With a **manually entered FX Rate**, EUR Unit Price rounds to the nearest whole euro: e.g. `€637.50 → €638.00`.
- Explicit manual EUR Unit Price edits are also rounded to the nearest whole euro. EUR display/Excel format remains fixed at two decimals. USD pricing is unchanged.

## v0.11.18 exhibition workflow refinements

- EUR reference FX source changed from Frankfurter to **Bank of China (Hong Kong) / BOCHK** T/T rates against USD. For EUR/USD, the PWA takes the midpoint of BOCHK Customer Sell and Customer Buy, then uses its reciprocal as the USD→EUR quotation rate. Manual FX override and cached fallback remain available.
- EUR Unit Price is rounded to the nearest whole euro after all pricing/FX/14K calculations, while all EUR money displays and Excel formats keep two decimal places, e.g. `€637.00`.
- Companion-search stone filters no longer show `CDM`; CDM remains visible in product descriptions. Stone filter chip labels are centred.
- Companion-search status summary is kept on one compact line using `Avail`, `Consign`, `Sold-OH`, and `Deliv`.
- Companion-search result actions stay in the right-side action column on iPhone instead of adding a bottom action row.
- Build-document item `編輯` action is moved to the right-side action rail below the `≡` order button, reducing item-card height.

## v0.11.17 work-trial fixes

- Fixes the iPhone `≡` move-to-position action: after entering a new formal position and tapping OK, the requested order is now preserved and all item sequence numbers are rebuilt continuously.
- EUR selling prices are rounded to whole euros after all USD / FX / 14K calculations. Preview, item cards, Amount, Sub Total, Discount, Total, Excel and PDF use the same zero-decimal EUR money format. USD behaviour is unchanged.
- 配套搜尋 款式 and 石頭 filters now support multiple selection. Multiple choices inside one group use OR logic (e.g. ER + PT); the 款式 and 石頭 groups combine with AND logic. `全部` clears that group.
- Existing v0.11.16 14K Quotation, Kitco Ask lookup and stock-search status logic are otherwise unchanged.


## v0.11.16 14K Quotation reference + companion stock search

### Quotation 14K reference
- Only **Quotation** shows a document-level karat selector: **18K 原款** (default) or **14K 參考報價**. Invoice and Consignment always stay on original stock data.
- Choosing 14K automatically requests the current **Kitco Gold Ask** online. The PWA first tries Kitco's live quote gateway and has a Kitco public-page fallback; the latest cached Ask and manual Ask input remain available if the online request is blocked.
- Company gold quotation rule: `Kitco Ask × 1.01`, rounded upward to the next USD 10. Production gold value then includes 12% consumption.
- 18K gold value / g: `Company Gold / 31.1034768 × 0.750 × 1.12`.
- 14K gold value / g: `Company Gold / 31.1034768 × 0.585 × 1.12`.
- Same-design 14K weight is estimated as `18K weight × 0.83`, then rounded **up** to the next 0.05g. Example: `3.15g → 2.65g`.
- DESC1 is rewritten only for the 14K Quotation, while preserving the original in brackets, e.g. `3.15Y750 → 2.65Y585 (3.15Y750)`. Preview, Excel and PDF use the same transformed description and reference Unit Price.
- Returning to Invoice / Consignment immediately restores original DESC1, gold weight, fineness and Unit Price. Stock source data is never rewritten.

### 配套／庫存搜尋
- Adds a fourth top-level tab after 文件預覽: **配套搜尋**.
- Search by full ARTNO or core number. Example: `34686` or `RG-34686` finds the whole `34686` family across RG / PT / ER / BL / NL / BR / BG variants.
- Dynamic chip filters are generated for **款式** and **石頭** from actual matched data.
- Results preserve all statuses: **Available**, **Consigned**, **Sold - On Hand**, and **Sold - Delivered**. Sold items stay visible for exhibition fulfilment decisions.
- Invoice confirmation records sold items as Sold - On Hand; Consignment records Consigned. Sold - On Hand can be marked Sold - Delivered from search.
- Available items, and Sold - On Hand items when appropriate, can be added directly to the current document from the search results.
- Inventory status history is stored locally on the device so confirmed exhibition movements remain searchable after reopening the PWA.

## v0.11.15 On-site image capture

- Item edit controls now include **上傳圖片** and **📷 即時拍照**.
- On iPhone, 即時拍照 uses the rear camera (`capture=environment`).
- Uploaded / captured images immediately replace that item's thumbnail and are used by document preview, Excel export, and PDF output.
- On-site images are compressed in-browser to JPEG (maximum side about 1400 px) to keep working files manageable while preserving aspect ratio and without cropping.
- The original Pictures-folder image is not overwritten. When an on-site image is active, **使用原資料庫圖片** restores the original variant; if no database image exists, the button removes the on-site image.
- On-site image object URLs are released when the item/draft is deleted or the document is confirmed.
- Existing v0.11.14 page layout, 100% scaling, row heights, column widths, margins, pagination, date label, and USD/EUR logic are unchanged.


## v0.11.14 Work-trial finishing touches

- Invoice, Consignment and Quotation now use the concise `Date :` label in the form, document preview and exported Excel/PDF output.
- The visible Total label follows the selected currency consistently: `Total : (USD)` or `Total : (EUR)` in preview and Excel/PDF output.
- All v0.11.13 layout, 100% scale, row height, approved column widths, margins, pagination and right-aligned page numbering are otherwise unchanged.

## v0.11.13 Currency cleanup

- Currency selector now only offers USD and EUR.
- GBP, CNY, JPY, and HKD have been removed from the selectable currency list.
- All v0.11.12 invoice layout, 100% scaling, row height, column widths, margins, pagination, and bottom-right page numbering remain unchanged.
- v0.11.9 remains the stable fallback baseline.

## v0.11.12 iPhone 100% pagination / width trial

- Built from v0.11.11; v0.11.9 remains the stable fallback baseline.
- Keeps true 100% Excel print scale and 1.0 cm margins.
- Uses the iPhone Excel column widths confirmed from INV260003: A 48 pt, B 69.75 pt, C 123.75 pt, D 24.75 pt, E 24.75 pt, F 48.75 pt, G 33.75 pt, H 50.25 pt, I 61.5 pt.
- Reduces Item-page capacity from 56 to 54 rows to match the observed iPhone Excel 100% physical page height.
- Footer reserves 16 rows, so a final page sharing the Footer allows up to 38 Item rows; otherwise the Item page may use the full 54 rows.
- Moves `Page &P of &N` from the centre footer to the right footer.
- Row heights are unchanged.



## v0.11.11 1.0 cm side-margin trial

- Built directly from v0.11.10; v0.11.9 remains the stable fallback baseline.
- Changes only the Excel print left/right margins from 1.4 cm to 1.0 cm for this test.
- Keeps true 100% Excel print scale; Fit-to-Width remains disabled.
- Keeps the same A:I column widths, row heights, top/bottom 1.0 cm margins, header/footer 0.8 cm margins, and 56 / 40+16 pagination logic as v0.11.10.
- Purpose: test whether iPhone Excel can export at 100% without manually choosing “Fit all columns on one page”.


## v0.11.10 100% print-space trial

- Keeps v0.11.9 as the stable fallback baseline.
- Keeps true 100% Excel print scale; Fit-to-Width is still disabled.
- Keeps the existing v0.11.9 item/separator row-height setting unchanged.
- Increases full Item-page capacity from 52 to 56 rows.
- Footer still reserves 16 rows, so a final page sharing the Footer allows up to 40 Item rows.
- If the Footer cannot fit, it moves to the next page and the Item page can use the full 56 rows.
- Print margins: top/bottom 1.0 cm, left/right 1.4 cm, header/footer 0.8 cm.
- A:I widths are rebalanced to 7.5 / 14.7 / 24 / 8.57 / 8.57 / 9.5 / 6.5 / 12 / 11.9. Picture columns D:E remain unchanged and Unit Price H remains 12.

## v0.11.9 Footer-aware pagination

- Keeps the approved Template print geometry from v0.11.8.
- Treats the Item area as 52 available rows on pages that do not contain the Footer.
- The Footer reserves 16 rows, so a final page that also contains the Footer may use up to 36 Item rows.
- If the final Item page needs more than 36 rows, the Footer is forced to the next page and the previous page may use the full 52 Item rows.
- Item separator rows count toward the Item-row total; Header and Footer do not.
- An Item is never split across a page break.

## v0.11.8 Template print lock

- Sales Rate is shown above Currency in the customer section.
- Template-based Excel export locks Item row height to 10.2 pt, including separator rows.
- Each printed page allows at most 36 Item rows; the separator row counts, Header and Footer do not.
- Items are kept intact and move to the next page if they would exceed the 36-row Item area.
- A:I column widths are locked to 13 / 18.71 / 28 / 8.57 / 8.57 / 10.86 / 7.29 / 12 / 12.14.
- A4 portrait export uses 100% scale instead of Fit-to-Width.
- Print margins are locked to the approved Template values: top/bottom 1.7 cm, left/right 0 cm, header/footer 0.8 cm.


## v0.11.0 Excel output changes

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


## v0.11.0 Consignment test
- Added Invoice / Consignment document type switch.
- Separate INV and CON number sequences.
- Consignment preview and Excel labels use Sales Consign / Consign Date.
- Confirming Consignment exports Available Stock and Consignment Out workbooks.

## v0.11.0 測試重點

- 新增 Invoice / Consignment / Quotation 三種文件類型與獨立流水號。
- 新增「匯入展覽資料包」：辨認 jmsdata、客戶表、Stone List、Invoice Template、選用 Article Mapping 及 Pictures。
- 預覽自動更新並顯示產品圖片；恢復 PDF 列印／儲存功能。
- 貨品清單改為精簡卡片，右側拖曳把手可改變預覽及輸出次序。
- Excel 產品圖片放大至完整四行圖片區。
- Remark 支援分行，Discount 支援括號顯示。

## v0.11.1 UI refinements
- Removed the visible “Document Type” label and kept Invoice / Consignment / Quotation on one row.
- Compact item cards now place LOTNO beside ARTNO; expanded editing shows every non-empty DESC1–DESC6 line.
- Delete action moved below the product image and only appears while editing; drag handle remains on the right.
- On-screen preview hides Vendor's Banker. Print/PDF shows the banker block on the right of the invoice/customer information.
- Preview/output table order is now No. / Article No. / Description / Picture / Quantity / Unit / Unit Price / Amount.


## v0.11.7 item sequence refinements
- Formal document order follows scan order: first scanned item is No. 1.
- The working list still shows the newest item at the top.
- Delete/re-add and drag reorder now renumber consistently across working list, preview, Excel and PDF.
- Qty label renamed to Quantity and Quantity / Unit Price edit inputs enlarged.


## v0.11.7
- Added GBP and CNY currency choices.
- Added online USD-based reference FX rates using Frankfurter v2, with cached/offline fallback and manual override.
- Foreign-currency prices are calculated from the USD selling price after Sales Rate.
- Excel/preview totals use the selected currency.


## v0.11.7 FX reliability fix
- Service Worker now ignores cross-origin requests so iPhone/Safari fetches the FX API directly instead of routing it through the PWA cache handler.
- Frankfurter v2 remains the primary source, with Frankfurter v1 as a fallback.
- FX request timeout increased to 12 seconds.
