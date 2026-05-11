# 📋 Release Notes

所有重要變更都會記錄在此文件中。

---

## � v8.1.0 — 雙語切換功能 (Latest)

> 發佈日期：2026-04-29  
> 適用平台：macOS 10.13+  
> 下載：[`SmartFinder_v8.1.dmg`](../../releases/latest) / [`SmartFinder_v8.1.zip`](../../releases/latest)

本版本以 **使用者體驗** 為核心，新增完整的中英文介面切換，並將設定持久化保存。

### ✨ 新功能 (Features)

- 🌐 **中英文雙語介面**  
  視窗左下角新增 **語言切換拉桿 (Language Slider)**，可在 `EN` 與 `中文` 之間即時切換，無須重啟程式。
- 💾 **語言偏好自動儲存**  
  選擇後會寫入 `Smart_Finder_settings.json`，下次啟動會記住你的選擇。**預設語言為英文 (English)**。
- 🏷 **版本號顯示**  
  右下角品牌標籤升級為 `Smart Finder v8.1.0 @ Xeson`，方便辨識所用版本。

### 🛠 改進 (Improvements)

- 🧩 **完整 i18n 架構**  
  全部 UI 文字（按鈕、標籤、欄位標題、對話框、警告訊息）皆透過 `TRANSLATIONS` 字典集中管理，未來新增語言只需擴充字典即可。
- 🎨 **視覺提示**  
  目前選中的語言會以 **粗體 + 高亮色** 顯示，未選中的另一個語言則使用次要色，狀態一目瞭然。
- 🔄 **即時 Live-update**  
  切換語言時，**結果列表的「類型」欄（Folder ↔ 資料夾）也會同步刷新**，無須重新搜尋。
- 🪪 **設定檔分離**  
  原本的 `Smart_Finder_recent_directories.json` 不變，新增獨立的 `Smart_Finder_settings.json` 儲存使用者偏好。

### 🐛 Bug 修復 (Bug Fixes)

- 修復 RenameDialog 在中文模式下按鈕文字截斷的問題（已調整視窗大小至 320×170）。
- 修復語言切換後 `Selected: N file(s)` 計數標籤未即時更新的問題。

### 📥 安裝方式

下載下方 **Assets** 區任一格式：

| 格式 | 適合對象 | 大小 |
| --- | --- | --- |
| `SmartFinder_v8.1.dmg` | 想要傳統 macOS 安裝體驗（拖入 Applications） | ~28 MB |
| `SmartFinder_v8.1.zip` | 偏好直接解壓使用，或需要在多台 Mac 同步部署 | ~24 MB |

⚠️ 首次啟動若被 Gatekeeper 阻擋，請在終端機執行：
```bash
xattr -cr /Applications/SmartFinder.app
```

### 🔄 從 v8.0.0 升級

直接下載新版覆蓋安裝即可，原本的「最近目錄」會被保留。  
**首次啟動 v8.1.0 將預設為英文**，可隨時透過左下角拉桿切換為中文。

---

## 🎉 v8.0.0 — 穩定版正式發佈

> 發佈日期：2026-04-29  
> 適用平台：macOS 10.13+  
> 下載：[`SmartFinder_v8.dmg`](../../releases/latest)

這是 **Xeson's Smart Finder** 首次於 GitHub 公開發佈的版本！本次更新著重於 **穩定性、容錯性與 macOS 原生體驗**，是日常使用的最推薦版本。

### ✨ 新功能 (Features)

- 🍎 **macOS 原生視窗 / Dock 圖示**  
  新增自訂 App Icon (`program_icon.png` / `icon-windowed.icns`)，視窗左上角與 Dock 皆會顯示精美圖示。
- 🏷 **右下角品牌標籤**  
  視窗右下角新增「Smart Finder @ Xeson」字樣與小型 icon，icon 大小會依字體自動縮放對齊，視覺更整潔。
- 📦 **一鍵打包腳本 `build_app.sh`**  
  整合 PyInstaller 流程，開發者只要 `./build_app.sh` 即可產出 `.app`。
- 💿 **官方 DMG 安裝檔**  
  首次提供 `SmartFinder_v8.dmg`，拖曳即可安裝。

### 🛠 改進 (Improvements)

- 🔁 **資源路徑相容性**  
  全新的 `resource_path()` 會在多個候選位置中尋找圖示檔，確保 PyInstaller 打包後仍能正確載入資源。
- 🧠 **更聰明的檔案類型判斷**  
  `get_file_kind()` 改善 Word / Word-old / Excel / Excel-MacroEnable / MOV / MP3 / Photoshop / AI 等檔案的識別文字，列表更易讀。
- 📐 **欄位寬度最佳化**  
  結果列表的 File Name / Kind / Date / Size / Location 五個欄位採用更合理的初始寬度。
- 🪪 **最近目錄記憶**  
  搜尋過的目錄會以 `Smart_Finder_recent_directories.json` 儲存，最多保留 5 個並自動排序。

### 🛡 穩定性 (Stability)

本版最大的升級！全面新增 **try / except 保護網**，避免在以下情境發生閃退：

- 🚫 遍歷時遇到 **無權限** 的資料夾 → 自動跳過並繼續搜尋
- 📁 檔案在搜尋後 **被刪除 / 移動** → 跳過該筆記錄而非崩潰
- 🧾 `Smart_Finder_recent_directories.json` 損毀 → 重置為空清單
- 📦 一次打開 / 移動 **超過 5 個** 檔案 → 跳出警告對話框，避免系統卡死
- 🏷 批量改名時偵測到 **重名衝突** → 中止操作以保護原始檔案
- ⚠️ UI 元件初始化失敗 → 回退至最小可用視窗，仍可關閉程式

### 🐛 Bug 修復 (Bug Fixes)

- 修復某些情境下 `os.path.relpath` 對跨磁碟路徑會丟出 `ValueError` 的問題（改為 fallback 至絕對路徑）。
- 修復 JSON 寫入失敗時會中斷搜尋流程的問題。
- 修復多選後若部分檔案已不存在，會導致整批操作失敗的問題。

### 📥 安裝方式

1. 至本頁面下方 **Assets** 區下載 `SmartFinder_v8.dmg`
2. 雙擊掛載後，將 `SmartFinder.app` 拖入「應用程式」資料夾
3. 首次啟動若被 Gatekeeper 阻擋，請至「系統設定 → 隱私權與安全性」點擊 **仍要打開**，或在終端機執行：
   ```bash
   xattr -cr /Applications/SmartFinder.app
   ```

### ⚠️ 已知限制

- 本 App **尚未經 Apple 公證 (Notarization)**，首次開啟需手動允許。
- 目前僅針對 **macOS** 優化，Windows / Linux 用戶建議從原始碼執行。
- 搜尋邏輯為單執行緒，超大目錄（10 萬+ 檔案）搜尋時 UI 可能短暫無反應。

### 🙏 致謝

感謝所有試用過 v1 ~ v7 版本並提供回饋的朋友們，你們的建議讓 SmartFinder 變得更好！

---

## 📜 歷史版本（Internal Builds）

> 以下為公開發佈前的內部測試版，未對外提供下載，僅作開發紀錄。

### v7 — 內部測試
- 重構 UI Layout，使用 `QHBoxLayout` 整理動作按鈕
- 加入「更改存放位置」功能，支援批量搬移與重名偵測

### v6 — 內部測試
- 加入「批量改名」對話框 (`RenameDialog`)
- 支援 Prefix / Suffix / 全名 + 自動編號 三種改名模式
- 改名後可選擇是否一併變更存放位置

### v5 — 內部測試
- 多選結果列表 (`ExtendedSelection`)
- 新增「選擇的檔案」與「搜尋結果」即時計數標籤
- 一次打開 >5 個檔案 / 文件夾時跳出警告

### v4 — 內部測試
- 新增最近目錄下拉清單 (`QComboBox` + JSON 持久化)
- 結果列表改用 `QTreeWidget` 多欄顯示

### v3 — 內部測試
- 加入檔案 Kind 智能判斷（Word / Excel / MOV / MP3 / PSD / AI…）
- 加入檔案大小自動換算（B / KB / MB / GB / TB）

### v2 — 內部測試
- PyQt5 GUI 雛型完成，支援關鍵字模糊搜尋
- 顯示檔名、修改日期、相對路徑

### v1 — 概念驗證
- 純 CLI 版本，使用 `os.walk` 遞迴搜尋
- 將結果輸出至終端機

---

<div align="center">

**[⬆ 回到最新版本](#-v800--穩定版正式發佈-latest)**

Made with ❤️ by **Xeson**

</div>
