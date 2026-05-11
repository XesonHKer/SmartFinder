# 📋 GitHub Release 文案範本（直接複製貼上用）

> 這份檔案是給你建立 GitHub Release 時用的「精簡版說明」，  
> 直接複製到 **Releases → Draft a new release → 描述欄** 即可。  
> 完整版本說明請見 [`RELEASE_NOTES.md`](RELEASE_NOTES.md)。

---

## ✏️ 建議的 Release 設定

| 欄位 | 建議內容 |
| --- | --- |
| **Tag version** | `v8.0.0` |
| **Release title** | `🎉 SmartFinder v8.0.0 — 首次公開發佈` |
| **Target** | `main` |
| **Set as the latest release** | ✅ 勾選 |
| **Assets（附件）** | `SmartFinder_v8.dmg` |

---

## 📝 Release 描述（複製以下內容）

```markdown
# 🎉 Xeson's Smart Finder v8.0.0

一款專為 macOS 設計的智能檔案搜尋與批量管理工具，現在正式於 GitHub 公開！  
此為首個對外發佈的穩定版本，著重於 **穩定性、容錯性與 macOS 原生體驗**。

## ✨ 主要亮點

- 🔍 **智能模糊搜尋**：在指定資料夾遞迴搜尋符合關鍵字的所有檔案
- ✏️ **批量改名**：支援 Prefix / Suffix / 全名 + 自動編號 三種模式
- 🚚 **批量搬移**：一次將多個檔案移至新位置，自動偵測重名衝突
- 📂 **批量打開**：可一次以系統預設程式開啟多個檔案或所在路徑
- 🕘 **最近目錄記憶**：自動保存最近 5 個搜尋目錄
- 🍎 **macOS 原生體驗**：附自訂 Dock / 視窗 Icon，已打包成獨立 `.app`

## 🛡 穩定性升級

- 全面 try/except 保護，**不再因權限不足、檔案被刪除而閃退**
- JSON 設定檔損毀時自動重置
- 一次操作 >5 個檔案會跳出警告，避免系統卡死
- 批量改名遇重名自動中止，**保護原始檔案**

## 📥 安裝方式

1. 下載下方 `SmartFinder_v8.dmg`
2. 雙擊掛載後，將 `SmartFinder.app` 拖入「應用程式」資料夾
3. 首次啟動若被 Gatekeeper 阻擋，於終端機執行：
   ```bash
   xattr -cr /Applications/SmartFinder.app
   ```
   或在「系統設定 → 隱私權與安全性」點擊 **仍要打開**

## ⚠️ 已知限制

- 本 App 尚未經 Apple 公證，首次開啟需手動允許
- 目前僅針對 macOS 優化（Windows / Linux 用戶可從原始碼執行）

## 📖 完整文件

- 詳細說明：[README.md](../blob/main/README.md)
- 完整變更紀錄：[RELEASE_NOTES.md](../blob/main/RELEASE_NOTES.md)

---

**Full Changelog**：首次公開發佈，無前一版本可比較。

> ⭐ 如果這個工具對你有幫助，請給個 Star 支持一下！  
> 🐛 Bug 回報 / 功能建議：[Issues](../../issues)

Made with ❤️ by **Xeson**
```

---

## 🚀 建立 Release 的快速指令

如果你已經把專案推上 GitHub，可以直接用 [GitHub CLI](https://cli.github.com/) 一行建立 Release：

```bash
# 確認已安裝 gh：brew install gh
# 確認已登入：gh auth login

gh release create v8.0.0 \
  SmartFinder_v8.dmg \
  --title "🎉 SmartFinder v8.0.0 — 首次公開發佈" \
  --notes-file RELEASE_NOTES.md \
  --latest
```

或者用網頁手動建立：

1. 進入 GitHub Repo → 右側 **Releases** → **Draft a new release**
2. **Choose a tag** 輸入 `v8.0.0` → **Create new tag**
3. **Release title** 填上「🎉 SmartFinder v8.0.0 — 首次公開發佈」
4. 描述欄貼上上方 markdown 內容
5. 把 `SmartFinder_v8.dmg` 拖入 **Attach binaries** 區域
6. 勾選 **Set as the latest release**
7. 點擊 **Publish release** 🚀
