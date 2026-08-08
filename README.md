# Universe Invoice PWA v0.14.9

## v0.14.9 快取檔案整理

- 移除 v0.14.8 重複的 `app-v0.14.8.js`；PWA 只保留單一主程式 `app.js`。
- `index.html` 以 `app.js?v=0.14.9` 載入主程式，保留版本化 cache-busting 效果。
- Service Worker 使用 `universe-invoice-v0.14.9` 新 cache，並繼續對本機 PWA 程式採用 network-first；離線時才 fallback 到 cache。
- Service Worker 註冊改用 `updateViaCache: none` 並主動檢查更新，降低 iPhone 沿用舊程式的機會。
- 其餘 v0.14.8 的 Universe Records / 新展覽隔離邏輯及所有功能維持不變。
