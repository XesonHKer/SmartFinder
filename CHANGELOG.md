# Changelog

All notable changes to Smart Finder will be documented in this file.

## v9.0.2 (2026-05-11)

### 🚀 New
- **Ghostscript Bundled In-App** — GS is now packed inside the `.app`. PDF compression works out of the box. No more `brew install ghostscript`.
- Smart GS path detection: bundled → Homebrew → system → `which gs`.

### 🐛 Fixed
- PDF compression cancel button now shows "Cancelling..." state for visual feedback.
- Detailed GS-not-found diagnostics (path-by-path failure reasons).

### 🔧 Changed
- `src/V9.0.2/` — Source code relocated to standardized `src/` directory.
- Build script: `build_app.sh` updated to v9.0.2 label, brew install note removed.
- README: now English-primary with 繁體中文 section at bottom.

### 📦 File Structure
```
SmartFinder/
├── src/V9.0.2/          # Latest source code
├── CHANGELOG.md         # This file
├── README.md            # English-primary docs
├── Preview.png          # Screenshot
└── .gitignore
```

---

## v9.0.1 (2026-05-11)

### 🐛 Fixed
- **PDF compressor no longer freezes UI** — Rewrote from `threading.Thread` to **QProcess** (native PyQt async process management).
- Real-time progress: elapsed seconds + GS stderr output shown on screen.
- Cancel button added: kills GS process immediately.
- Dynamic timeout based on file size (30s base + 3s per MB).

### 🏷 UI
- Bottom-left corner now shows version label `v9.0.1` in light gray.

---

## v9.0.0 (2026-04-29)

### 🗜️ New
- **PDF Compression** — Quality slider with real-time size estimation. Compress PDFs to 15%~70% of original size.
- Ghostscript backend (requires `brew install ghostscript`).
- Save-as dialog for output location.

### 🔧 Changed
- All features from v8.1.0 preserved.
- Original source kept in V9.0.0 directory.

---

## v8.1.0 (2026-04-29)

### 🌐 New
- **Bilingual UI** — Full English / 繁體中文 support for all buttons, labels, dialogs, and column headers.
- **Language slider** at bottom-left corner — instant toggle between `EN ⇄ 中文`, no restart needed.
- **Language persistence** — Choice saved in `Smart_Finder_settings.json`, defaults to English on first launch.

### 🔧 Changed
- All features from v8 preserved.

---

## v8.0.0

### ✨ Features
- Smart fuzzy file search across directories.
- File info display: name, type, date modified, size, relative path.
- Recent directories (last 5 remembered).
- Batch open files / batch open location.
- Batch rename: prefix, suffix, full rename with auto-numbering.
- Batch move with duplicate detection.
- Robust error handling (permission errors, missing files, corrupted JSON).
- Native macOS experience: custom Dock icon, app icon, standalone `.app`.
- Settings persisted in JSON files.