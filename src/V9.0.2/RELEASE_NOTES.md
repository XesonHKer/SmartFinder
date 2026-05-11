# SmartFinder v9.0.2 — Release Notes

发布日期 / Release Date: 2026-05-11

## 🚀 新功能 (New Features)

### 1. Ghostscript 內建打包，無需自行安裝！
- **重大改善**：將 Ghostscript 完整打包進 `.app` 內部，使用者**不再需要**透過 `brew install ghostscript` 手動安裝。
- **自動偵測路徑**：程式會依序尋找 bundled GS → `/opt/homebrew/bin/gs` → `/usr/local/bin/gs` → `/usr/bin/gs` → `which gs`，確保在任何環境下都能正確定位。
- **詳細除錯資訊**：若找不到 GS，會明確顯示每個候選路徑的失敗原因（檔案不存在、無執行權限、執行錯誤等）。

### 2. PDF 壓縮支援取消
- 壓縮過程中可按 **Cancel** 按鈕立即中斷（`QProcess.kill()`），按鈕會顯示「取消中...」回饋。

### 3. 進度顯示優化
- 壓縮期間同時顯示品質設定（如 `50% / 100 DPI`）與 GS 即時輸出。
- 每 2 秒更新已過時間，確保使用者知道程式仍在運作。

## 🐛 修復 (Bug Fixes)

- GS Bundle 的動態庫依賴使用 `@loader_path`，確保在不同 macOS 版本下都能正確載入。
- 修正 build_app.sh 中的版本號標籤（v9.0.1 → v9.0.2）。

## 📦 文件結構 (File Structure)

```
V9.0.2/
├── Icon/
│   ├── program_icon.png
│   └── icon-windowed.icns
├── gs_bundle/
│   ├── gs                    # Ghostscript 靜態二進位
│   ├── gs-wrapper.sh         # 環境變數包裝腳本
│   └── Resource/             # GS 初始化腳本與資源
├── SmartFinder_v9_0_2.py     # 主程式 (v9.0.2)
├── pdf_compressor.py         # PDF 壓縮模組 (QProcess + GS Bundle)
├── SmartFinder.spec          # PyInstaller 設定檔
├── build_app.sh              # 一鍵打包腳本
├── README.md                 # 使用說明書
├── RELEASE_NOTES.md          # 本文件
└── GITHUB_RELEASE_v9.0.2.md  # GitHub Release 用文案
```

## 🛠 打包方法 (Build)

```bash
cd V9.0.2
chmod +x build_app.sh
./build_app.sh
# => dist/SmartFinder.app (~28 MB, 含 GS Bundle)
```

## 🔄 與 V9.0.1 的差異 (Diff vs V9.0.1)

| 項目 | V9.0.1 | V9.0.2 |
|------|--------|--------|
| Ghostscript 安裝 | ❌ 需自行 `brew install ghostscript` | ✅ **內建 GS Bundle，開箱即用** |
| PDF 壓縮後端 | QProcess (system GS) | QProcess (bundled GS → system GS) |
| GS 找不到時 | 單純提示安裝 | ✅ 顯示詳細偵測路徑與失敗原因 |
| 取消按鈕 | ✅ Cancel | ✅ Cancel + 「取消中...」回饋 |
| 即時進度 | ✅ 已過時間 + GS 輸出 | ✅ 已過時間 + GS 輸出 + 品質設定 |
| 介面版本號 | `v9.0.1` 左下角 | `v9.0.2` 左下角 |
| V9.0.1 保留 | — | ✅ 完整保留於 `打包v9/V9.0.1/` 目錄 |