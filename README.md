# Universe Invoice PWA v0.14.15

## v0.14.15 Template style inheritance

- Stone Description keeps the approved C:D / E:H layout and final blank row, but no longer hard-codes Arial 8 pt in the Template export path.
- `STONE DESCRIPTION:` and all stone-description cells now copy the Invoice Master Template Remark content cell's full style (font, size, bold, alignment, wrap, fill, border, number format, etc.).
- LOTNO / ARTNO item cells no longer force `bold = false`; their appearance is inherited entirely from the Invoice Master Template just like the other item fields.
- The no-Template fallback workbook keeps its own built-in styling because there is no Template style to inherit.


## v0.14.14 Stone Description footer refinement
- Excel Stone Description left column is fixed to `C:D`; right column is fixed to `E:H`.
- Stone Description header and stone-name text were set to 8 pt in v0.14.14; v0.14.15 replaces that hard-coded font styling with Template Remark style inheritance.
- Always keep one blank row after the final Stone Description row before the signature line.


## v0.14.13 Stone Description

- Excel Invoice / Consignment / Quotation 會根據目前文件貨品 DESC1–DESC6 自動辨認 Stone List BREAKDOWN。
- 英文石名完全由 Stone List「英文石名」欄讀取，不在 PWA 寫死。
- Remark 自動加入 `STONE DESCRIPTION:`，每行最多兩個石種，左右以 Excel 真正兩個儲存格區塊排列，不靠空格對齊。
- 同一石種代號只列一次，按文件貨品首次出現次序排列。
- 超過 Template 原有 Remark 空間時會自動增加 Remark 行，再將 Signature / Accept By 往下移。
- PWA 仍只輸出 Excel，不新增 PDF 輸出。
- 其他 v0.14.12 功能維持不變。
