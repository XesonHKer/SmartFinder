# SmartFinder v9.0.0 — Release Notes

发布日期 / Release Date: 2026-05-11

## 🎯 新增功能 (New Features)

### 1. PDF 压缩功能 / PDF Compressor
- 在主視窗操作工具列新增「**壓縮 PDF**」按鈕。
- 打開獨立壓縮對話框，提供 **品質拉桿** (0%~100%)，拖動即可看到預估大小變化。
- 支援 Ghostscript 後端壓縮，可將 PDF 壓縮至原大小的 15%~70%。
- 支援 **English / 繁體中文** 雙語介面。
- 背景執行壓縮，壓縮期間 UI 不會卡死。

### 2. 技術細節
- 新的獨立模組 `pdf_compressor.py`，可被主程式匯入。
- 壓縮前自動檢查 Ghostscript 安裝狀態，未安裝時提示 `brew install ghostscript`。
- 使用 QSlider 即時預估壓縮後大小，用戶決定後才開始壓縮。

## 📦 文件结构 (File Structure)

```
V9.0.0/
├── Icon/
│   ├── program_icon.png
│   └── icon-windowed.icns
├── SmartFinder_v9_0_0.py     # 主程式 (含 i18n + PDF 壓縮)
├── pdf_compressor.py          # PDF 壓縮模組 (全新)
├── SmartFinder.spec           # PyInstaller 設定檔
├── build_app.sh               # 一鍵打包腳本
├── README.md                  # 使用說明書 (已更新)
├── RELEASE_NOTES.md           # 本文件
└── GITHUB_RELEASE_v9.0.0.md   # GitHub Release 用文案
```

## 🛠 打包方法 (Build)

```bash
cd V9.0.0
chmod +x build_app.sh
./build_app.sh
# => dist/SmartFinder.app
```

**相依套件 (PyInstaller 打包時):**
```bash
pip install pyqt5 pyinstaller
```

**運行時相依 (用戶端):**
```bash
brew install ghostscript
# Ghostscript 用於 PDF 壓縮核心功能
```

## 🔄 與 V8.1.0 的差異 (Diff vs V8.1.0)

| 項目 | V8.1.0 | V9.0.0 |
|------|--------|--------|
| PDF 壓縮 | ❌ | ✅ 品質拉桿 + 預估大小 |
| 按鈕行 | 4 個按鈕 | 5 個按鈕 (+壓縮 PDF) |
| 背景執行 | — | PDF 壓縮在背景執行 (UI 不凍結) |
| 模組檔案 | 1 個 (.py) | 2 個 (+ pdf_compressor.py) |
| Ghostscript 依賴 | ❌ | ✅ 選用 (brew install ghostscript) |