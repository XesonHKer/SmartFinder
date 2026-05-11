# SmartFinder v9.0.0

> Release date: 2026-05-11
> Platform: macOS (Apple Silicon / arm64)

A bilingual (English / 繁體中文) macOS file-finder utility with **PDF compression** support.

---

## ✨ What's new in v9.0.0

### 🗜️ PDF Compressor (Interactive Slide Bar)
- New **"Compress PDF"** button in the action toolbar.
- Opens a dedicated dialog with a **quality slider** (0%–100%):
  - Drag to see estimated output size in real-time.
  - Click "Compress PDF" to run in the background (UI stays responsive).
- Uses Ghostscript (`gs`) as the compression backend — shrinks image-heavy PDFs by 30%–85%.
- Full bilingual support (English / 繁體中文).

### 🧩 Modular design
- `pdf_compressor.py` is a standalone module — easy to maintain and test.
- Auto-detects Ghostscript installation and guides the user to `brew install ghostscript` if missing.

---

## 📦 Download

| Asset | Size | Notes |
|------|------|-------|
| `SmartFinder-v9.0.0-macOS-arm64.zip` | ~24 MB | macOS app bundle (Apple Silicon) |

**Prerequisite (for PDF compression):**
```bash
brew install ghostscript
```

---

## 🚀 How to install / run

1. Download `SmartFinder-v9.0.0-macOS-arm64.zip` and double-click to unzip — you'll get `SmartFinder.app`.
2. Move `SmartFinder.app` to `/Applications` (optional).
3. First launch: macOS Gatekeeper may show *"unidentified developer"* because the app is unsigned. To open it:
   - **Right-click** on `SmartFinder.app` → **Open** → confirm **Open**.
   - Or run once in Terminal:
     ```bash
     xattr -dr com.apple.quarantine /Applications/SmartFinder.app
     ```
4. **For PDF compression**, install Ghostscript:
   ```bash
   brew install ghostscript
   ```

---

## 🔄 Diff vs v8.1.0

| Item | v8.1.0 | **v9.0.0** |
|------|--------|-----------|
| PDF compression | ❌ | ✅ Quality slider + size estimate |
| Action buttons | 4 | 5 (+ "Compress PDF") |
| Background processing | — | PDF compression runs in background |
| Module files | 1 (.py) | 2 (+ pdf_compressor.py) |
| Ghostscript dependency | ❌ | ✅ Optional (`brew install ghostscript`) |

---

## 🛠 Build from source

```bash
cd V9.0.0
chmod +x build_app.sh
./build_app.sh
# => dist/SmartFinder.app
```

Requires Python 3.11+, PyQt5, PyInstaller 6.x.

---

## 🗒 Notes

- The original v8 / v8.1.0 source are still kept in `打包v8/` folder, untouched.
- PDF compression requires Ghostscript (`brew install ghostscript`).
- Settings files written next to the working directory:
  - `Smart_Finder_settings.json` — language preference
  - `Smart_Finder_recent_directories.json` — recent search directories