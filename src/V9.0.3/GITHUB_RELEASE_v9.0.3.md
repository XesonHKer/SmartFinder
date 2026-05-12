# SmartFinder v9.0.3

> Release date: 2026-05-13
> Platform: macOS (Apple Silicon / arm64)

A bilingual (English / 繁體中文) macOS file-finder utility with **built-in PDF compression** — no more manual Ghostscript installation.

**v9.0.3** is a critical bugfix release: PDF compression no longer drops characters (missing `&`, lines, etc.) that occurred when Ghostscript stripped font encoding information from TrueType font subsets.

---

## ✨ What's new in v9.0.3

### 🐛 Fixed: PDF Compression Font Encoding → Missing Characters

**Problem:** Ghostscript 10.07 silently strips `/Encoding` from single-glyph TrueType font subsets (e.g. Zapfino used only for a `&` symbol) during MacRoman→WinAnsi conversion. This caused characters to appear blank or missing in Apple Preview, Chrome, and some other PDF viewers — even though the glyph data was still in the file.

**Root cause:** A design issue in Ghostscript 10.07's `pdfwrite` engine — when it converts MacRoman-encoded TrueType subsets, it removes the `/Encoding` entry from the font dictionary. Most PDF renderers fall back to standard encoding, but single-char subsets with non-ASCII mappings can lose their glyph entirely.

**Fix:** After Ghostscript compression, SmartFinder now runs an automatic post-processing step:
1. Decompresses the PDF with `qpdf`
2. Scans every TrueType font object that's missing `/Encoding`
3. Injects `/Encoding /WinAnsiEncoding` — the correct, standard declaration
4. Recompresses the PDF (no size increase)

**No files changed** other than the encoding declaration. The font programs, images, layout, and all other content remain exactly as Ghostscript output them.

### 🔄 Diff vs v9.0.2

| Item | v9.0.2 | **v9.0.3** |
|------|--------|-----------|
| Font encoding fix | ❌ No — GS strips encoding | ✅ **Automatic post-compression fix** |
| Missing characters | ❌ `&`, lines, some glyphs invisible in Preview | ✅ All characters rendered correctly |
| PDF size after fix | 2.5 MB (compressed) | 2.5 MB (same — zero overhead) |
| Tested on | — | ✅ Portfolio PDF (155 pages, 18 fonts) |
| qpdf required | ❌ Not used | ✅ Bundled or system `qpdf` |
| Ghostscript bundled | ✅ Yes | ✅ Yes (unchanged) |

---

## 📦 Download

| Asset | Size | Notes |
|------|------|-------|
| `SmartFinder-v9.0.3-macOS-arm64.zip` | ~28 MB | macOS app bundle (Apple Silicon) — GS bundled inside |

**No prerequisites required** (Ghostscript is included!).

---

## 🚀 How to install / run

1. Download `SmartFinder-v9.0.3-macOS-arm64.zip` and double-click to unzip.
2. Move `SmartFinder.app` to `/Applications`.
3. First launch: **Right-click** → **Open** → confirm **Open**.
   - Or run: `xattr -dr com.apple.quarantine /Applications/SmartFinder.app`
4. Done! PDF compression works immediately with all fonts preserved.

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
cd V9.0.3
chmod +x build_app.sh
./build_app.sh
# => dist/SmartFinder.app (~28 MB)
```

### Requirements to build
- Python 3.11
- PyQt5 (`pip install pyqt5`)
- PyInstaller (`pip install pyinstaller`)
- qpdf (`brew install qpdf`)

---

## 🗒 Notes

- v9.0.0 through v9.0.2 source is fully preserved in `打包v8/` and `打包v9/V9.0.2/` for rollback.
- The font encoding fix requires `qpdf` to be installed. It's not bundled — users should `brew install qpdf`.
- If `qpdf` is missing, compression still works but font encoding won't be fixed. A warning will be shown.
- Settings files: `Smart_Finder_settings.json`, `Smart_Finder_recent_directories.json`.
- The bundled GS is statically-linked (no system library dependencies needed).