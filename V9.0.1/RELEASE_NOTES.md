# SmartFinder v9.0.1 — Release Notes

发布日期 / Release Date: 2026-05-11

## 🐛 修復 (Bug Fixes)

### 1. PDF 壓縮 UI 無回應問題
- **原因**：原版用 `threading.Thread` + `QTimer.singleShot` 更新 UI，worker thread 無法可靠刷新進度文字，用戶看到「壓縮中...」卡住不動，誤以為當機。
- **修復**：全面改用 **QProcess**（PyQt5 原生異步進程管理），主線程直接監聽 Ghostscript 輸出，UI 更新即時、可靠。
- **新增即時進度顯示**：每 2 秒顯示已過時間（「已過 30 秒...」），並顯示 GS 即時輸出。
- **新增取消按鈕**：壓縮期間可按 Cancel 立即中斷（`process.kill()`）。

### 2. 時間預估改善
- 根據檔案大小動態計算 timeout（30 秒基數 + 每 MB 3 秒）。

## 🏷 界面調整

- 介面**左下角**新增淺灰色版本號標記（`v9.0.1`），位於語言切換拉桿與右下角 App 標籤之間。

## 📦 文件结构 (File Structure)

```
V9.0.1/
├── Icon/
├── SmartFinder_v9_0_1.py     # 主程式 (v9.0.1)
├── pdf_compressor.py          # PDF 壓縮模組 (QProcess 版本)
├── SmartFinder.spec           # PyInstaller 設定檔
├── build_app.sh               # 一鍵打包腳本
├── README.md                  # 使用說明書
├── RELEASE_NOTES.md           # 本文件
└── GITHUB_RELEASE_v9.0.1.md   # GitHub Release 用文案
```

## 🛠 打包方法 (Build)

```bash
cd V9.0.1
chmod +x build_app.sh
./build_app.sh
# => dist/SmartFinder.app
```

## 🔄 與 V9.0.0 的差異 (Diff vs V9.0.0)

| 項目 | V9.0.0 | V9.0.1 |
|------|--------|--------|
| PDF 背景執行 | `threading.Thread` + QTimer | ✅ **QProcess** (可靠 UI 更新) |
| 即時進度顯示 | ❌ 卡在「壓縮中...」 | ✅ 顯示已過時間 + GS 輸出 |
| 取消按鈕 | ❌ | ✅ Cancel (kill process) |
| 介面版本號 | ❌ | ✅ `v9.0.1` 淺灰色左下角 |
| V9.0.0 保留 | — | ✅ 完整保留於 `V9.0.0/` 目錄 |