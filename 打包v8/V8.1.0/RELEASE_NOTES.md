# SmartFinder v8.1.0 — Release Notes

发布日期 / Release Date: 2026-04-29

## 🌐 新增功能 (New Features)

### 1. 双语界面 / Bilingual UI
- 新增完整的 **English** 介面，所有按鈕、標籤、訊息框、欄位標題等均已國際化。
- 保留原有 **繁體中文** 介面。

### 2. 左下角语言选择拉桿 / Language Slider (Bottom-Left)
- 在主視窗左下角加入了一個 `EN ⇄ 中文` 的拉桿（QSlider）。
- 拉桿向左 = English，向右 = 中文。
- 切換語言會即時刷新整個 UI（無需重啟 App）。
- 語言偏好會保存在 `Smart_Finder_settings.json`，下次啟動會自動沿用上次選擇。

### 3. 默認語言 / Default Language
- 首次啟動或設定檔不存在時，**默認語言為 English**。

## 📦 文件结构 (File Structure)

```
V8.1.0/
├── Icon/
│   ├── program_icon.png
│   └── icon-windowed.icns
├── SmartFinder_v8_1_0.py     # 主程式 (含 i18n)
├── SmartFinder.spec          # PyInstaller 設定檔
├── build_app.sh              # 一鍵打包腳本
└── RELEASE_NOTES.md          # 本文件
```

## 🛠 打包方法 (Build)

```bash
cd V8.1.0
chmod +x build_app.sh
./build_app.sh
```

打包後的 App 位於 `V8.1.0/dist/SmartFinder.app`。

## 🔄 與 V8 的差異 (Diff vs V8)

| 項目 | V8 | V8.1.0 |
|------|----|--------|
| 介面語言 | 繁體中文 | 英文 + 繁體中文 |
| 語言切換 | 不支援 | 左下角拉桿 (即時切換) |
| 默認語言 | 中文 | **English** |
| 語言偏好保存 | — | `Smart_Finder_settings.json` |

> **保留 V8**：原始 V8 內容 (`SmartFinder_v8.py`、`SmartFinder.spec`、`build_app.sh`、`Icon/`) 仍然原封不動地放在 `打包v8/` 目錄下，方便日後重用或回溯。
