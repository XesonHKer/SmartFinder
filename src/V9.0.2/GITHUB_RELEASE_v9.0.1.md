# SmartFinder v9.0.1

> Release date: 2026-05-11
> Platform: macOS (Apple Silicon / arm64)

A bilingual (English / 繁體中文) macOS file-finder utility with **PDF compression** support.

**v9.0.1** is a bugfix release — the PDF compressor now shows real-time progress and can be cancelled.

---

## ✨ What's new in v9.0.1

### 🐛 Fixed: PDF Compressor "freezes" during compression
- **Root cause**: `threading.Thread` + `QTimer.singleShot` couldn't reliably update the UI from a worker thread — the progress label appeared stuck.
- **Fix**: Rewrote the compressor to use **QProcess** (native PyQt async process management) — the UI updates in real-time.
- Real-time progress: elapsed seconds + GS stderr output shown on screen.
- **Cancel button**: kill the GS process immediately if needed.

### 🏷 Version label in UI
- Bottom-left corner now shows `v9.0.1` in light gray (`#aaaaaa`), between the language slider and the app label.

### 🗜️ PDF Compressor (same as v9.0.0)
- Quality slider with real-time size estimation.
- Background compression via Ghostscript (requires `brew install ghostscript`).

---

## 📦 Download

| Asset | Size | Notes |
|------|------|-------|
| `SmartFinder-v9.0.1-macOS-arm64.zip` | ~24 MB | macOS app bundle (Apple Silicon) |

**Prerequisite (for PDF compression):**
```bash
brew install ghostscript
```

---

## 🚀 How to install / run

1. Download `SmartFinder-v9.0.1-macOS-arm64.zip` and double-click to unzip.
2. Move `SmartFinder.app` to `/Applications`.
3. First launch: **Right-click** → **Open** → confirm **Open**.
   - Or run: `xattr -dr com.apple.quarantine /Applications/SmartFinder.app`
4. For PDF compression: `brew install ghostscript`

---

## 🔄 Diff vs v9.0.0

| Item | v9.0.0 | **v9.0.1** |
|------|--------|-----------|
| PDF compressor backend | threading | **QProcess** (reliable UI) |
| Real-time progress | ❌ Stuck label | ✅ Elapsed + GS output |
| Cancel button | ❌ | ✅ Kill GS process |
| Version label in UI | ❌ | ✅ `v9.0.1` bottom-left |
| v9.0.0 preserved | — | ✅ Kept in `V9.0.0/` |

---

## 🛠 Build from source

```bash
cd V9.0.1
chmod +x build_app.sh
./build_app.sh
# => dist/SmartFinder.app
```

---

## 🗒 Notes

- v9.0.0 source is fully preserved in `打包v8/V9.0.0/` for rollback.
- Settings files: `Smart_Finder_settings.json`, `Smart_Finder_recent_directories.json`.