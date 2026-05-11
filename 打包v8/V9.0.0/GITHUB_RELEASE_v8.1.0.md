# SmartFinder v8.1.0

> Release date: 2026-04-29
> Platform: macOS (Apple Silicon / arm64)

A bilingual (English / 繁體中文) macOS file-finder utility, with an instant language toggle slider at the bottom-left corner.

---

## ✨ What's new in v8.1.0

### 🌐 Bilingual UI (English + 繁體中文)
- 全新完整的 **English** 介面，所有按鈕、標籤、訊息框、欄位標題等均已國際化。
- 保留原有 **繁體中文** 介面。

### 🎚 Language slider (bottom-left)
- 主視窗左下角新增 `EN ⇄ 中文` 拉桿（QSlider）。
- 拉桿向左 = English，向右 = 中文。
- 切換語言會 **即時刷新整個 UI**（無需重啟 App）。

### 💾 Persistent language preference
- 語言偏好會保存於 `Smart_Finder_settings.json`，下次啟動會自動沿用上次選擇。

### 🇬🇧 Default language = English
- 首次啟動或設定檔不存在時，**預設語言為 English**。

---

## 📦 Download

| Asset | Size | Notes |
|------|------|-------|
| `SmartFinder-v8.1.0-macOS-arm64.zip` | ~24 MB | macOS app bundle (Apple Silicon) |

**SHA-256**
```
4521d3e457e8727087acc22433002d8d15b1f3b89c762c738ac056484e88b7fd  SmartFinder-v8.1.0-macOS-arm64.zip
```

---

## 🚀 How to install / run

1. Download `SmartFinder-v8.1.0-macOS-arm64.zip` and double-click to unzip — you'll get `SmartFinder.app`.
2. Move `SmartFinder.app` to `/Applications` (optional).
3. First launch: macOS Gatekeeper may show *"unidentified developer"* because the app is unsigned. To open it:
   - **Right-click** on `SmartFinder.app` → **Open** → confirm **Open**.
   - Or run once in Terminal:
     ```bash
     xattr -dr com.apple.quarantine /Applications/SmartFinder.app
     ```

---

## 🔄 Diff vs v8

| Item | v8 | **v8.1.0** |
|------|----|-----------|
| UI language | 繁體中文 | English + 繁體中文 |
| Language switch | ❌ | ✅ Bottom-left slider (live) |
| Default language | 中文 | **English** |
| Preference persistence | — | `Smart_Finder_settings.json` |

---

## 🛠 Build from source

```bash
cd V8.1.0
chmod +x build_app.sh
./build_app.sh
# => dist/SmartFinder.app
```

Requires Python 3.11+ and PyQt5, plus PyInstaller 6.x.

---

## 🗒 Notes

- The original v8 source/spec are still kept in the parent `打包v8/` folder, untouched, so v8 can be rebuilt at any time.
- Settings files written next to the working directory:
  - `Smart_Finder_settings.json` — language preference
  - `Smart_Finder_recent_directories.json` — recent search directories
