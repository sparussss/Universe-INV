# Universe Invoice PWA v0.14.31

## v0.14.31
- 「匯入 Invoice PDF」按鍵文字改為「匯入 PDF」。
- 「匯入 PDF」按鍵直接沿用「資料匯入」按鍵的字體大小與高度／padding 規則。
- 「客人 Invoice 資料」標題列與下方客戶摘要區增加正常垂直間距，避免按鍵貼住「尚未選擇客戶」區塊。



## v0.14.30
- Invoice 文件預覽內所有文字節點強制統一為 10px，包含 Header、貨品、Totals、Payment Term、Remarks、聲明、Stone Description、Signature。
- 使用 `!important` 覆蓋舊有各區獨立預覽字級，只影響 PWA 預覽，不改 Excel 輸出格式。

## v0.14.30
- 將「輸出 Excel 附加資料」由匯出前彈出視窗移到「文件預覽」頁面。
- Invoice 附加資料全部預設未選；勾選／取消時，下方 Invoice 預覽即時更新。
- 預覽會模擬實際 Excel footer 的位置與結構：Payment Term、Remarks 重量、Gross Weight 上單線／下雙線、Natural Stone 聲明、Stone Description 兩欄。
- 正式按「匯出 Excel Invoice」時直接使用目前預覽頁的勾選狀態，不再重複彈出附加資料選擇視窗。
- Quotation／Consignment 會隱藏這組 Invoice 附加資料選項。


## Excel export add-on chooser

Invoice Excel export now opens an **optional add-on** dialog before creating the workbook. Every checkbox resets to **unchecked** on every export, for every customer.

Available add-ons:

- Payment Term
- Remark weight details
  - Gold Weight
  - Semi-Precious
  - Diamond
  - All Stone
- Natural-stone declaration
- Stone Description

The previous customer-code restriction is removed. `Total Amount` uses the normal shared Invoice Master Template logic and is not part of the optional add-ons.

When weight details are selected:

- Gold Weight = DESC1 gold weight × Qty
- Semi-Precious = non-diamond DESC2–DESC6 carat weight × Qty
- Diamond = diamond DESC2–DESC6 carat weight × Qty
- All Stone = all DESC2–DESC6 carat weight × Qty
- carats are also shown in grams using ct × 0.2

Diamond / Semi-Precious classification and Stone Description use the **currently imported Stone List**. Stone Description keeps A–Z ordering and two-column output. Optional footer formatting inherits the current Invoice Master Template Remark styles.

If any add-on is selected, the current Invoice Master Template is required. Stone Description and Diamond / Semi-Precious weight options also require the current Stone List.

All other v0.14.20 functionality remains unchanged.


## v0.14.22
- Invoice Excel 附加資料 Remark 新增 Gross Weight 選項；Gross Weight = Gold Weight + All Stone grams，數值格上單線、下雙線。
- 所有附加資料區行高（包括空白行）統一為 10.5 pt。
- 重量標題更新：TOTAL SEMI-PRECIOUS STONES WEIGHT、TOTAL DIAMOND WEIGHT、TOTAL STONES WEIGHT。

## v0.14.30
- 文件預覽整張 Invoice 的標題與內容字體大小統一為 Remarks / Stone Description 內容的預覽字級。
- 包括公司 Header、Invoice 標題、公司／客戶資料、表頭、貨品資料、Totals、Total Amount、Payment Term、Remarks、聲明、Stone Description 及 Signature。
- 只影響 PWA 文件預覽；實際 Excel 字體及格式仍由 Invoice Master Template / 既有 Excel 輸出規則控制。

## v0.14.30
- 文件預覽頂部移除文字公司名／地址，改用提供的 Universe Gems & Jewellery Company 公司標誌。
- 公司標誌在預覽內靠左顯示，方向跟 Excel Invoice 一致。
- 只修改 PWA 預覽；Excel 輸出仍按 Invoice Master Template 原有內容及格式。
