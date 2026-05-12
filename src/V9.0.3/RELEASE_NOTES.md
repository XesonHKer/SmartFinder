# SmartFinder v9.0.3 Release Notes

## v9.0.3 (2026-05-13)

### 🐛 Bugfix: PDF Compression Font Encoding

**Problem:** Ghostscript 10.07 pdfwrite strips `/Encoding` from single-glyph TrueType subsets (e.g. Zapfino used only for `&`), causing characters to appear blank in Apple Preview / Chrome PDF viewer.

**Fix:** New post-compression `fix_font_encoding()` step automatically scans the compressed PDF and injects `/Encoding /WinAnsiEncoding` into all font objects that are missing it. The fix:
- Preserves every character exactly as rendered
- Adds zero size overhead (~0 bytes change on a 2.5 MB PDF)
- Runs automatically — no user interaction needed
- Shows status during compression ("修復字體編碼...")

**Files changed:**
- `pdf_compressor.py` — added `fix_font_encoding()` function, integrated into `_on_process_finished`
- `SmartFinder_v9_0_2.py` → `SmartFinder_v9_0_3.py` — version bumped to 9.0.3
- `SmartFinder.spec` — update entry point and version strings
- `build_app.sh` — version label update

### ⚙️ New Runtime Dependency
- `qpdf` is now required for the font encoding fix (`brew install qpdf`)
- If `qpdf` is not installed, compression still works but font repair is skipped

### 📋 Full Changelog
- ✓ Fix: Missing & and decorative characters in compressed PDFs
- ✓ Zero overhead: font fix doesn't increase file size
- ✓ Graceful fallback: compression works without qpdf
- ✓ Rebuilt from v9.0.2 source base (all features preserved)

---

## Older Versions

### v9.0.2 (2026-05-11)
- Ghostscript bundled inside `.app` — no `brew install` needed
- Cancel with feedback ("Cancelling..." state)
- Real-time progress with quality info
- Detailed GS-not-found diagnostics
- Language persistence across restarts
- Recent directories memory (last 5)

### v9.0.1 (2026-04-25)
- Batch move with duplicate detection
- Batch rename with prefix/suffix/numbering
- Recent directories tracking
- Language persistence
- PDF compression with quality slider and progress

### v9.0.0 (2026-04-15)
- Initial v9 release
- Smart fuzzy file search
- Detailed file info display
- Batch open / open location
- Bilingual UI (English / 繁體中文)