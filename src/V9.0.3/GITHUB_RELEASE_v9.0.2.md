# SmartFinder v9.0.2

> Release date: 2026-05-11
> Platform: macOS (Apple Silicon / arm64)

A bilingual (English / 繁體中文) macOS file-finder utility with **built-in PDF compression** — no more manual Ghostscript installation.

**v9.0.2** is a major usability release: Ghostscript is now **bundled inside the `.app`**, so PDF compression works out of the box. No `brew install` needed.

---

## ✨ What's new in v9.0.2

### 🚀 Ghostscript Bundled — No More `brew install`!
- Ghostscript is now packed inside the `.app` bundle (~28 MB total).
- **Zero setup**: download → open → compress PDFs immediately.
- Smart fallback: bundled GS → Homebrew → system paths → `which gs`.
- Detailed diagnostics: if GS can't be found, shows exactly why for each path.

### 🗜️ PDF Compressor Improvements
- **Cancel with feedback**: press Cancel to kill GS immediately; button text changes to "Cancelling...".
- Real-time progress shows quality setting (e.g. `50% / 100 DPI`) alongside elapsed seconds.
- Same reliable QProcess backend from v9.0.1 (no UI freezing).

### 🔄 Diff vs v9.0.1

| Item | v9.0.1 | **v9.0.2** |
|------|--------|-----------|
| Ghostscript setup | ❌ Must `brew install ghostscript` | ✅ **Bundled — works out of the box** |
| PDF compressor backend | QProcess (system GS) | QProcess (bundled → system GS) |
| GS-not-found info | Simple message | ✅ Detailed path-by-path diagnostics |
| Cancel button | ✅ Cancel | ✅ Cancel + "Cancelling..." state |
| Real-time progress | ✅ Elapsed + GS output | ✅ Elapsed + GS output + quality info |
| Version label | `v9.0.1` bottom-left | `v9.0.2` bottom-left |
| v9.0.1 preserved | — | ✅ Kept in `打包v9/V9.0.1/` |

---

## 📦 Download

| Asset | Size | Notes |
|------|------|-------|
| `SmartFinder-v9.0.2-macOS-arm64.zip` | ~28 MB | macOS app bundle (Apple Silicon) — GS bundled inside |

**No prerequisites required** (Ghostscript is included!).

---

## 🚀 How to install / run

1. Download `SmartFinder-v9.0.2-macOS-arm64.zip` and double-click to unzip.
2. Move `SmartFinder.app` to `/Applications`.
3. First launch: **Right-click** → **Open** → confirm **Open**.
   - Or run: `xattr -dr com.apple.quarantine /Applications/SmartFinder.app`
4. Done! PDF compression works immediately.

---

## 🖼 Features at a glance

- 🔍 **Smart fuzzy search**: recursively search directories for files matching keywords.
- 🗂 **Detailed file info**: name, type, date modified, size, relative path.
- 📂 **Batch open**: open multiple files at once (with warning for >5).
- 📍 **Batch open location**: jump to Finder folders of selected files (deduplicated).
- ✏️ **Batch rename**: add prefix/suffix or full rename with auto-numbering.
- 🚚 **Batch move**: move files to a new location with duplicate detection.
- 🗜️ **PDF compression**: quality slider with real-time size estimation — **no external install**.
- 🌐 **Bilingual UI**: switch between English / 繁體中文 instantly (slider at bottom-left).
- 💾 **Language persistence**: remembers your choice across restarts.
- 🕘 **Recent directories**: remembers last 5 search directories for quick access.

---

## 🛠 Build from source

```bash
cd V9.0.2
chmod +x build_app.sh
./build_app.sh
# => dist/SmartFinder.app (~28 MB)
```

### Requirements to build
- Python 3.11
- PyQt5 (`pip install pyqt5`)
- PyInstaller (`pip install pyinstaller`)

---

## 🗒 Notes

- v9.0.0 and v9.0.1 source is fully preserved in `打包v8/` and `打包v9/V9.0.1/` for rollback.
- Settings files: `Smart_Finder_settings.json`, `Smart_Finder_recent_directories.json`.
- The bundled GS is statically-linked (no system library dependencies needed).