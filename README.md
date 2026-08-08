# Universe Invoice PWA v0.14.13

## v0.14.13 Stone Description

- Excel Invoice / Consignment / Quotation 會根據目前文件貨品 DESC1–DESC6 自動辨認 Stone List BREAKDOWN。
- 英文石名完全由 Stone List「英文石名」欄讀取，不在 PWA 寫死。
- Remark 自動加入 `STONE DESCRIPTION:`，每行最多兩個石種，左右以 Excel 真正兩個儲存格區塊排列，不靠空格對齊。
- 同一石種代號只列一次，按文件貨品首次出現次序排列。
- 超過 Template 原有 Remark 空間時會自動增加 Remark 行，再將 Signature / Accept By 往下移。
- PWA 仍只輸出 Excel，不新增 PDF 輸出。
- 其他 v0.14.12 功能維持不變。
