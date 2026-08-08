# Universe Invoice PWA v0.14.3

## v0.14.3 Universe Records 單一正式文件紀錄來源

- 新增 `Universe Records.json`，作為每次展覽會唯一正式的 Invoice／Consignment／Quotation 歷史、Recall、Revision、Cancelled、Exhibition ID／Name 及文件流水號來源。
- 每個新展覽資料包都應放入一份全新的空白 `Universe Records.json` Template；Template 的 `exhibitionId`／`exhibitionName` 可留空，首次匯入時 PWA 會產生新的 Exhibition ID，並以資料包 Folder 名稱作 Exhibition Name。
- 匯入新資料包時，PWA 會先清空上一個資料包在記憶體中的正式文件記錄，再完全以新資料包的 `Universe Records.json` 載入；因此新展覽不會再顯示上一個展覽的 Recall 記錄。
- `jmsdata.xlsx` 不再保存 Invoice Header／Items、Consignment Header／Items、Quotation Header／Items 或 Transaction History 工作表；舊版 jmsdata 如仍有這些工作表，重新輸出時會移除，避免雙重正式紀錄。
- Confirm 後會分別輸出最新 `jmsdata.xlsx` 及 `Universe Records.json`；請把兩個檔案取代回同一個展覽資料包。Quotation 只更新 Records，不改 I:J:K:L 庫存。
- 「匯出更新後資料包」會同時把最新 `jmsdata.xlsx` 及 `Universe Records.json` 放回 ZIP 內。
- PWA 本機不再保存 Confirm 文件歷史作第二套正式資料庫；本機只保留未完成草稿、Image Override 暫存及必要設定。
- 「新展覽會・重設文件編號」現在亦會建立一套全新的空白文件記錄與新 Exhibition ID，下一號回到 `INVYY0001 / CONYY0001 / QUOYY0001`。


## v0.14.2 配套搜尋切換／大量清單效能

- 配套搜尋頁切走再切回時，如資料／篩選／目前文件沒有變化，保留既有 DOM，不再重新建立整個庫存清單。
- 配套搜尋結果改為分批載入：首次只建立 42 張貨品卡，接近清單底部時再自動加入下一批 54 張；所有篩選／排序結果數量仍以完整庫存計算。
- 保留配套搜尋原本卷軸位置，切回頁面時返回上次位置。
- Stone List 最長前綴、Stone Order、Stone Group、貨品石種組合加入快取，避免排序比較時反覆重新解析 DESC2–6。
- 圖片配對結果加入配套搜尋專用快取；「圖片待處理」數量改為閒置時分批計算，避免進入配套搜尋時一次做完整圖片分析。
- 資料包完成匯入 Pictures 後，先建立第一批配套搜尋畫面快取，使首次按「配套搜尋」更快。
- 圖片／Stone List／庫存狀態變更時會自動失效相關快取，避免顯示舊資料。

## v0.14.1 客戶優先導覽／Invoice 交貨狀態／jmsdata.xlsx

- 「資料匯入」移到「展覽文件工作台」標題右側；主導覽改為「選擇客戶／建立文件／配套搜尋／文件預覽」，預設停在「選擇客戶」。
- 配套搜尋只顯示 Available／Consigned／Sold - On Hand／Sold - Delivered 等庫存狀態，不再提供「標記已交貨／取消已交貨」操作。
- Invoice 每款貨品在「編輯」下方加入「未交貨／已交貨 ✓」按鍵；底部加入「全部標記已交貨」。這些選擇只在 Confirm Invoice 時正式寫入 jmsdata。
- Recall Invoice 會按目前 jmsdata 自動帶入每款實際交貨狀態；Delivered 貨品要刪除／取消文件時，必須先在該 Invoice 明確改成「未交貨」，不再要求到配套搜尋處理。
- Confirm 前最後核對新增 Invoice 已交貨／未交貨款數。
- 庫存主檔正式統一為 `jmsdata.xlsx`；PWA 庫存匯入只接受 `.xlsx`，Confirm／手動匯出／更新資料包亦統一輸出 `jmsdata.xlsx`。


## v0.14.0 展覽資料管理／健康檢查／備份／診斷／圖片效能

- 新增內部 Exhibition Name：開始新展覽並重設文件編號時輸入名稱；只供 PWA 管理及 Recall 分組，不印在 Invoice／Consignment／Quotation。
- Recall 依 Exhibition Name 分組；同一年不同展覽即使文件號同樣由 YY0001 開始亦可清楚區分。
- 資料匯入頁新增「目前正在使用的資料」，顯示 jmsdata、Stone List、GoldSilver、Customer、Pictures、Template 等實際載入版本／摘要。
- 新增「展覽資料包健康檢查」：重複 LOTNO、I:J:K:L／Balance、Stone List 一致性、DESC2–DESC6 未辨認石種、Customer 預設 Sales Rate、圖片缺漏、GoldSilver 新鮮度、Template 狀態等。只提示，不自動改來源資料。
- Stone List 自身一致性檢查會區分「同 BREAKDOWN 但石類／GROUP 矛盾」與「同 BREAKDOWN 的多個 QUOTATION Alias（正常）」；亦檢查缺 GROUP／石類／QUOTATION Alias，以及正常的前綴重疊資訊；前綴仍採最長配對。
- GoldSilver.xlsx 新增「可能過期」提示；以工作日計算，週末不會被當作交易日。
- v0.14.0 曾把 Confirm 歷史納入 PWA 本機備份；自 v0.14.3 起正式文件歷史改由 `Universe Records.json` 單一保存，本機備份只保留草稿、Image Override 暫存及設定。
- 「更新資料包」與「清空所有 PWA 本機資料」完全分開。自 v0.14.3 起 Confirm 文件歷史不再與本機合併，而是完全跟隨資料包的 `Universe Records.json`；Image Overrides 仍保留既有安全合併機制。
- 新增 LOTNO 診斷報告，可查看 Stone List 解析、鑽石／色石、GROUP、MULTI、金色、圖片候選、最終選圖、Image Override 及 Quotation 計算資料，亦可匯出文字報告。
- 大量圖片效能改善：圖片索引分批建立並顯示進度；不再在匯入時為全部圖片建立 Object URL；配套搜尋及文件縮圖改用 Lazy Load，只在接近畫面時載入；關閉圖片編輯器後釋放非選用候選圖的 Object URL。


## v0.13.2 Stone List 單一石種資料來源

- 鑽石代號不再寫死於 PWA；由 Stone List 的「石類」欄判斷，石類含「鑽」的 BREAKDOWN 代號會自動視為鑽石。
- 移除 PWA 內置的石種後備對照表；BREAKDOWN / QUOTATION 對照完全跟 Stone List。
- 移除 PWA 內置的 Stone Group 後備表；GROUP 完全跟 Stone List。
- 之後新增石種、修改 QUOTATION Alias、修改 GROUP 或鑽石代號，只需更新資料包內的 Stone List，毋須再修改 PWA。

## v0.13.1 Stone List 最長前綴辨認／配套排序

- DESC2–DESC6 的石種辨認統一以 `Stone List` 的 `BREAKDOWN` 作最長前綴配對：例如 `PTQOV → PTQ`、`PAMPR → PAM`、`GTQ... → GTQ`、`BOX... → BOX`，避免被較短代號 `GT`／`BO` 截錯。
- `BREAKDOWN` 代號即使 QUOTATION 欄暫時留空，仍會被視為有效石種代號；QUOTATION 只作圖片／輸出別名。
- 後備 Stone List 增補 RL／VAQ／TIO／MOP／BMOP／BO，與新版資料表的重要代號一致。
- 配套／庫存搜尋在同一款號核心及同一顏色 Group 內，先按完整石種組合 Signature 排序，再按 RG／ER／PT 等款式排序；相同配色／石種組合會放在一起。
- 例如 37397 的 `PTQ` 一組會先集中排列，`PAM + PTQ` 另一組再集中排列，不會再因 RG／ER／PT 次序而交錯。


## v0.13.0 展覽可靠性／Recall／草稿／核對

- Quotation Confirm 後寫入 `Quotation Header`／`Quotation Items`，可像 Invoice／Consignment 一樣 Recall、修訂或取消；Quotation Recall 不改 jmsdata I:J:K:L 庫存。
- Invoice／Consignment／Quotation 未完成文件自動儲存草稿；重新整理或 PWA 被 iPhone 中斷後，重新匯入展覽資料包可選擇恢復。
- Confirm 前新增最後核對：Document No.、客戶、Sales Rate、Currency、Quantity、Sub Total、Discount、Total、圖片待處理及人手改價數量。
- Customer、Sales Rate、Document No. 等關鍵資料不完整時會先阻止 Confirm。
- 文件貨品區新增圖片狀態摘要；有黑白 fallback／無圖時，Confirm 及 Excel 匯出前會提醒，但不強制阻止。
- Unit Price 經人手修改後顯示「人手修改 Unit Price」標記，最後核對亦會列出數量。
- 新增「新展覽會・重設文件編號」，可將當年 Invoice／Consignment／Quotation 下一號重設至 `INVYY0001`／`CONYY0001`／`QUOYY0001`；內部以展覽 Session 區分同年重複文件號，舊文件不改並仍可 Recall。
- Service Worker 會在版本安裝時預先快取 SheetJS、html5-qrcode、ExcelJS、SortableJS、JSZip 等核心 library；成功安裝／更新一次後，展覽核心流程可在離線情況繼續使用。


## v0.12.14 Excel BL 圖片闊度 4 公分

- Excel 匯出遇到 BL 款式時，Picture 區會擴闊至足以容納 4 公分圖片。
- BL 圖片目標闊度為 4.00 cm，高度按原圖比例使用同一縮放倍率自動計算，不拉高、不壓扁、不裁切。
- 圖片繼續在固定四行 Picture 區域內水平及垂直置中；若個別 BL 圖片的高度會超出四行區域，會按高度等比例縮細。
- 只有文件內含 BL 款式時才會啟用較闊 Picture 版面；沒有 BL 的 Excel 保持原有欄寬。
- Invoice Template 路徑及無範本後備輸出均套用相同規則。


## v0.12.13 NL `(1)` 圖片優先及 u價顯示強化

- NL 款式在石種及金色配對分數相同時，優先自動選用檔案名最後帶 `(1)` 的圖片。
- `(1)` 優先只作同等配對的排序條件，不會凌駕石種或金色配對，亦不影響 RG／PT／ER／BL 等其他款式。
- 配套／庫存搜尋的 u價取消千位分隔逗號，例如 `3,396u` 改為 `3396u`。
- u價改為黑色粗體；LOTNO 及分隔點維持原本灰色樣式。

## v0.12.12 Excel 圖片保持原比例

- Excel 匯出圖片統一使用 `contain` 方式，圖片寬度及高度採用同一縮放比例。
- 修正極橫 BL 圖片被拉高，以及直向圖片被壓扁的問題。
- 圖片完整顯示、不裁切，並在固定四行 Picture 區域內水平及垂直置中。
- 使用 Invoice Template 及沒有範本的 Excel 匯出路徑，均使用同一套圖片放置邏輯。

## v0.12.11 配套搜尋顯示 u價

- 配套／庫存搜尋的每張貨品卡，會在 `LOTNO` 旁顯示 jmsdata `PRICE` 原始 u價，例如 `LOTNO 139418 · 3092u`。
- u價只作庫存搜尋參考，不會乘 Sales Rate，亦不會改動 Invoice／Quotation 計算。
- 如 PRICE 無有效數值，卡片只顯示 LOTNO。

## v0.12.10 London PM 輸入及 MULTI 判定修正

- Quotation 三個成色按鍵繼續同一行平排；兩個 14K 選項固定分成兩行顯示，避免手機自行斷字。
- 移除最新 London PM 日期輸入框。日期只由資料包 `GoldSilver.xlsx` 最新有效紀錄讀取；金價輸入框保留作臨時手動覆蓋。
- 更新金價時直接更新資料包內 `GoldSilver.xlsx`，PWA 不會寫回或修改該 Excel。
- 擴闊 MULTI 判定：保留「3 個或以上顏色 Group」規則；另外，3 種或以上不同非鑽石石種、並橫跨至少 2 個顏色 Group，也視為 MULTI。
- 例如 QAM + GT + YCT 會選用 `MULTI.JPG` 彩色圖；多種石頭但全部同一顏色 Group 不會受影響。

## v0.12.9 介面修正與 MULTI SA 自動選圖

- Quotation 三種成色選項在手機上固定同一行平排：18K 原款、14K 參考報價、14K 同金重報價。
- London PM 歷史金價區預設收起，以展開／收起按鈕控制；收起時仍顯示最新日期及價錢摘要。
- 修正 iPhone 日期輸入框超出卡片及日期文字置中的問題。
- 移除 Quotation 貨品標題列的「回到最新」按鈕，只保留「清空草稿」。
- 自動選圖會把檔名內獨立出現的 `MULTI` 視為多色圖片，包括 `MULTI SA`；例如 BL-37120 可按 Y750 自動選用 18KY 圖片。




## v0.12.7 最新已公布 London PM（完全取代現貨 Ask）

- Quotation 金價來源只使用 Kitco 歷史頁最近一個已公布的 London PM，不再讀取或使用即時現貨 Ask。
- 最新 London PM 必須同時輸入日期及金價；例如 2026-08-04 使用 2026-08-03 的 USD 4026.60 / oz。
- 最新 London PM 日期不會自動假設為今天或前一天，方便遇到星期六、星期日及休市日時輸入真正最新的已公布日期。
- 最新 London PM 及各款完成日 London PM 都採用相同公司金價、18K／14K 金值及 Quotation 差額公式。
- 如果貨品完成日剛好等於最新 London PM 日期，系統會直接共用該最新金價，毋須重複輸入。
- 使用全新本機快取欄位，不會沿用舊版可能保存的現貨 Ask 數值。

## v0.12.6 London PM Quotation 金價重算 / 配套搜尋圖片操作

- 18K 及 14K Quotation 均以最新 Kitco London PM 對比貨品完成日 London PM。
- 完成日期讀取 jmsdata Column R（LDATE）；同一完成日只需輸入一次歷史 London PM，並儲存在本機。
- 公司金價：London PM × 1.01，再向上取至下一個 USD 10。
- 18K／14K 每克金值均加入 12% 生產消耗；14K 金重按 0.83 並每 0.05g 向上進位。
- Quotation 調整：（最新成色金總值－完成日 18K 金總值）× 4 × Sales Rate，再加入 PRICE × Sales Rate 後向上取整。
- 金價區改為兩欄排列，最新 London PM 輸入值靠左。
- 配套搜尋的「編輯圖片」移到左邊縮圖下方。


## v0.12.5 continuous preview / Excel-only export / item image actions

- 完整移除 PWA 的 PDF 輸出功能：按鈕、產生程式、下載流程、第三方 PDF／畫面擷取依賴及相關提示全部移除；文件預覽只保留 Excel 匯出。
- 文件預覽取消 A4 分頁及頁碼，改為單一連續長頁面；Excel 的正式 A4 分頁、每頁 10 款及現有版面完全不變。
- Invoice、Consignment、Quotation 的每件貨品左側統一改為「圖片 → 編輯圖片 → 刪除」。
- 右側「編輯／完成」只處理 Quantity 與 Unit Price；兩個輸入欄位繼續平排。
- 圖片編輯仍共用配套搜尋的候選圖片視窗，保留彩色／黑白、上傳圖片、即時拍照及恢復自動選圖。
- NL 自動選圖仍優先使用同候選檔名帶 `(1)` 的版本。
- RQZ／OPAL 選圖不使用重量：真正多色款若沒有 MULTI 圖，會使用 Description 中實際存在的精確單石圖片，再考慮黑白 fallback。

## v0.12.4 Excel-aligned A4 preview / direct PDF trial

- 文件預覽改用固定 794 × 1123 px 的 A4 畫布，再按手機寬度等比例縮放；不再因手機寬度重新排列欄位。
- 預覽表格欄寬重新按 Excel 版面配置，Description、Picture、Quantity、Unit、Unit Price、Amount 的比例固定。
- 預覽採用 Excel 相同的分頁概念：一般頁最多 10 款；最後一頁最多 7 款時同頁顯示 Footer，8–10 款時 Footer 另開一頁。
- PDF 不再使用瀏覽器列印網頁；改以同一張 A4 預覽逐頁轉成 PDF，因此 PDF 與畫面預覽一致。
- Excel 輸出程式及已確認的 Excel Template 版面不作修改。
- 圖片自動判定不使用重量；石頭代號本身也會作圖片別名，若款式屬多色但沒有 MULTI 圖，會改用實際存在的精確單石圖（例如 RQZ／OPAL），而不是直接轉黑白。
- 這是預覽／PDF 對齊 Excel 的試驗版，實際字型渲染仍可能與 Microsoft Excel 有細微差異。

## v0.12.3 操作與預覽更新

- 配套搜尋進入後直接顯示全部貨品；輸入框改為「輸入款號」，清空後回復全部。
- 配套搜尋與文件貨品共用圖片彈窗；彩色／黑白平排，並加入上傳圖片／即時拍照。
- Quantity 與 Unit Price 平排。
- 只有 NL 自動選圖時，同一候選優先使用檔名 `(1)` 版本。
- Confirm Invoice／Consignment 匯出的 jmsdata 會包含最後手動選圖的 Image Overrides。
- 新增「匯出更新後資料包」，將更新後 jmsdata 及現場新增圖片一併打包。
- 文件預覽改成 A4 分頁、每頁最多 10 款，版面與 Excel 欄位及分頁方向一致。

## v0.12.2 展覽前圖片整理

- 配套搜尋新增「圖片待處理」篩選：包含完全沒有圖片，以及自動選圖只能使用黑白 fallback 的貨品。
- 每件配套搜尋結果新增「編輯圖片」，可從該 ARTNO 的 Pictures 候選中選用彩色或黑白圖片。
- 手動選圖以 LOTNO 儲存，優先於 GROUP／MULTI／成色／黑白自動選圖，並同步用於文件、預覽、Excel 及 PDF。
- `jmsdata.xls` 新增 `Image Overrides` 工作表，記錄 LOTNO、ARTNO、圖片檔名、Variant、黑白狀態及更新時間。
- 匯入更新後的 jmsdata 及原 Pictures 資料夾後，其他同事亦可取得相同選圖結果。
- 配套搜尋顯示尚未匯出的圖片修改數量；完成展覽前整理後一次過匯出更新後的 `jmsdata.xls` 即可。
- 提供「恢復自動選圖」，刪除該 LOTNO 的手動覆蓋記錄。

## v0.12.1 下一輪現場流程修正

- Quotation 可加入 Available、Sold on hand、Consignment、Delivered 四種狀態的 LOTNO；可由配套搜尋或 LOTNO 輸入加入，Confirm Quotation 不改動 jmsdata I:M。
- Recall Invoice／Consignment 後可刪除全部貨品再 Confirm；最新 Revision 會保存為 `CANCELLED`，所有原貨品恢復到首次加入文件前的 I:M 狀態。
- Cancelled 文件仍可再次 Recall；重新加入當時仍 Available 的貨品並 Confirm 後，狀態回復為 `CONFIRMED`，原 Document No. 不變。
- Delivered 貨品新增「取消已交貨」；必須在 Transaction History 找到交貨前狀態，才會回復到 Sold on hand／Consignment，並輸出更新後 jmsdata.xls。
- 建立文件的 Invoice／Consignment／Quotation 貨品列表，以及配套搜尋長列表，加入右下角浮動 `↑` 回到頂部按鈕。
- Recall 修訂不再推進下一個文件流水號，避免每次修改舊文件都跳過新文件號碼。

## v0.12.0 jmsdata document records + Recall

- 匯入及輸出檔名繼續使用公司沿用多年的 `jmsdata.xls`，不改名為其他 Database 名稱。
- 第一個工作表保留原本 jmsdata 庫存資料與 53 欄結構，只按 LOTNO 更新 I:M：I Delivered、J Available、K Sold on hand、L Consignment、M Balance。
- 每次狀態更新遵守 `SUM(I:L)=M`；目前每個 LOTNO 一件，狀態只是在 I:L 之間移動數值 1，M 維持 1。
- 更新後 `jmsdata.xls` 會加入五個工作表：`Invoice Header`、`Invoice Items`、`Consignment Header`、`Consignment Items`、`Transaction History`。
- Invoice／Consignment Confirm 後會保存文件 Header、逐件 Items、庫存轉移及 Revision，並輸出更新後的 `jmsdata.xls`。
- 建立文件頁新增 Recall 搜尋，可按文件號碼、客戶或日期重新開啟原 Invoice／Consignment。
- Recall 保留原 Document No.；刪除貨品會恢復首次加入文件前的 I:M 狀態，新加入貨品重新 Confirm 後才轉成 Sold on hand／Consignment，保留貨品不會重複扣庫存。
- 每次 Recall 重新 Confirm 會新增 Revision，不覆蓋舊版本；最新版本用於下一次 Recall。
- 若某件貨在原文件之後已轉成其他狀態，Recall 不會強行覆蓋後續交易；系統會停止 Confirm 並提示先檢查該 LOTNO。
- Header 不保存 Total Qty／Total Amount；開啟文件時由 Items、Discount 重新計算，避免 Header 與 Items 不同步。
- 配套搜尋只容許 Available 貨品加入新文件；已在 Recall 文件中的非 Available 貨品會由 Recall 自動載入。
- 資料匯入頁新增「匯出目前 jmsdata.xls」，方便在標記 Delivered 或其他狀態更新後重新輸出。
- 現場上傳／拍照的自訂圖片只記錄檔名；資料包圖片的 variant／黑白選擇可 Recall。自訂圖片本身不會嵌入 jmsdata。
- v0.11.32 的圖片尾碼容錯、v0.11.31 配套顏色排序、v0.11.30 GROUP／MULTI 選圖、拖曳排序及正式 Excel/PDF 版面全部保留。

### 建議測試流程

1. 匯入現有 `jmsdata.xls` 及完整展覽資料包。
2. 建立一張兩件貨的 Invoice，Confirm 後儲存新下載的 `jmsdata.xls`。
3. 重新匯入該 `jmsdata.xls`，在 Recall 搜尋原 Invoice No.。
4. 刪除其中一件，再重新 Confirm。
5. 檢查刪除貨品回到 J=1，保留貨品維持 K=1；`Invoice Header` 有 Revision 0 和 1，`Invoice Items` 的 Revision 1 只保留一件。
6. 以 Consignment 重複測試，確認 L 欄及 `Consignment Header/Items`。

# Previous versions

## v0.11.32 Missing-dot suffix tolerance for image filenames

- 正式 ARTNO 仍以 Stock 資料為準，例如 `PT-35774.A`、`PT-35774.B`。
- 圖片檔名如因人手輸入漏了尾碼前的句點，`PT-35774A ...JPG` 會配對到正式 `PT-35774.A`；`PT-35774B ...JPG` 會配對到正式 `PT-35774.B`。
- 規則適用於任何英數尾碼，不只 `.A`／`.B`。
- 無尾碼 `PT-35774` 仍只配對自己的圖片，不會誤取 `PT-35774A`、`PT-35774B` 或帶句點尾碼的圖片。
- 畫面、搜尋、Excel、PDF 仍顯示 Stock 的正式 ARTNO；容錯只用於 Pictures 圖片索引。
- v0.11.31 的配套搜尋顏色排序，以及 v0.11.30 的 GROUP／MULTI／成色／黑白選圖邏輯全部保留。


## v0.11.31 Unified colour/stone sorting for Companion Search

- The same colour-group sorting is now applied to every Companion Search, not only `*`.
- Applies to core article searches, full ARTNO searches, style filters, stone filters, and status filters.
- Order: core number descending → Stone List `GROUP` order → Stone List stone order → `RG → ER → PT → BR → NL → BL → BG → other` → ARTNO / LOTNO.
- When stone filters are selected, matching selected stones are used as the sorting key so the filtered results stay grouped by the chosen colours/stones.
- Existing GROUP-driven image matching and all v0.11.30 behaviour are preserved.


## v0.11.30 Stone List GROUP-driven image matching

- 匯入 `Stone List & Shape & Cutting.xlsx` 時會讀取新增的 `GROUP` 欄。
- MULTI 不再按石頭數量硬判斷；會按不同顏色群組判斷。三個或以上不同 GROUP 才視為真正多色組合。
- BTO + LBT + IO + TZ（即使再加入 BSA）都屬 Blue，同色系不會判作 MULTI。
- QAM + BTO + YCT + GPS 分別跨 Purple/Pink/Rose、Blue、Yellow/Orange、Green，會判作 MULTI。
- 完整石頭組合圖片仍是最高優先，例如 `BT+L.BT` 會先於 MULTI。
- 非 MULTI 款會檢查所有 DESC 的有效石頭及 Stone List 內全部 QUOTATION 別名，例如 RQZ、PAM、BT / L.BT。
- 真正多色款：完整組合圖 → MULTI 圖 → 黑白 fallback。
- 非多色款：正常石頭／組合圖 → 黑白 fallback。
- 成色只在石頭配對層級相同時作次級選擇，例如 BT 候選中以 `Y750 → BT (18KY)`。
- Invoice、Consignment、Quotation、配套搜尋共用同一選圖核心；14K Quotation 仍以原始 18K 成色選圖。

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
- The v0.11.16 stock-search status logic remains; its former live-price quotation method is retired and removed from the current build.


## v0.11.16 14K Quotation reference + companion stock search

> Historical release note: the quotation gold-price method below is retained for version history only and is superseded by v0.12.6.

### Quotation 14K reference

> 舊版即時現貨金價流程已完全停用及移除。現行版本只使用上方 v0.12.7 所述的最新已公布 London PM 與完成日 London PM。

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


## v0.12.8 GoldSilver.xlsx and Quotation modes

- The exhibition package can include `GoldSilver.xlsx`. The PWA reads the `Date` and `Gold PM` columns from the `GoldSilver.com` sheet.
- The newest valid Gold PM date/value is automatically placed in the editable latest London PM fields.
- For each item LDATE, the PWA uses the same date or the nearest earlier trading date in GoldSilver.xlsx.
- Historical values remain manually editable as overrides.
- Added a third quotation mode: `14K 同金重報價`, which converts 18K to 14K while retaining the original 18K metal weight.
- Existing `18K 原款` and `14K 參考報價` calculations are retained.
