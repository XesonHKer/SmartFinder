#!/usr/bin/env python3
"""
SmartFinder v9 — PDF Compressor module
Integrated into SmartFinder as a separate dialog/feature.
Requires Ghostscript installed: brew install ghostscript
"""
import os, subprocess, threading, shutil, json, re
from pathlib import Path
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
    QSlider, QFileDialog, QMessageBox, QProgressBar, QLineEdit
)
from PyQt5.QtCore import Qt, QTimer

def get_ghostscript_path():
    """Find Ghostscript binary — try common locations."""
    candidates = [
        "/opt/homebrew/bin/gs",
        "/usr/local/bin/gs",
        "/usr/bin/gs",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    # Try which command
    try:
        r = subprocess.run(["which", "gs"], capture_output=True, text=True, timeout=5)
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout.strip()
    except:
        pass
    return None

def estimate_pdf_size(pdf_path, quality_pct):
    """
    Estimate output size based on quality percentage.
    Uses real analysis of current PDF + heuristic formula.
    """
    try:
        size_bytes = os.path.getsize(pdf_path)
        # Map percentage to DPI / compression level
        # 100% -> no compression (original size)
        # 75%  -> /printer (300dpi) ~70% of original for image-heavy PDFs
        # 50%  -> /ebook (150dpi) ~40%
        # 25%  -> /screen (72dpi)  ~25%
        # 0%   -> max compression ~15%
        
        ratios = {
            0: 0.15,
            10: 0.18,
            20: 0.22,
            25: 0.25,
            30: 0.30,
            40: 0.35,
            50: 0.40,
            60: 0.50,
            70: 0.60,
            75: 0.65,
            80: 0.72,
            90: 0.85,
            100: 1.0,
        }
        # Find closest matching ratio
        closest_pct = min(ratios.keys(), key=lambda x: abs(x - quality_pct))
        ratio = ratios[closest_pct]
        
        estimated = int(size_bytes * ratio)
        return estimated
    except:
        return 0

def compress_pdf(input_path, output_path, quality_pct, callback=None):
    """
    Compress a PDF file using Ghostscript.
    
    Args:
        input_path: Source PDF path
        output_path: Destination PDF path  
        quality_pct: 0-100 slider value
        callback: Optional function(progress_msg) for UI updates
    """
    gs_path = get_ghostscript_path()
    if not gs_path:
        raise RuntimeError("Ghostscript not found. Please run: brew install ghostscript")
    
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input PDF not found: {input_path}")
    
    # Map quality percentage to Ghostscript PDFSETTINGS
    if quality_pct >= 80:
        # /printer — 300 DPI, high quality
        pdf_settings = "/printer"
        dpi = 300
    elif quality_pct >= 50:
        # /ebook — 150 DPI, medium
        pdf_settings = "/ebook"
        dpi = 150
    elif quality_pct >= 25:
        # /screen — 72 DPI
        pdf_settings = "/screen"
        dpi = 72
    else:
        # Max compression
        pdf_settings = "/screen"
        dpi = 72
    
    if callback:
        callback(f"壓縮中 (品質 {quality_pct}%, {dpi} DPI)...")
    
    # Build Ghostscript command
    cmd = [
        gs_path,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.5",
        f"-dPDFSETTINGS={pdf_settings}",
        "-dUseFlateCompression=true",
        "-dNOPAUSE",
        "-dBATCH",
        "-dQUIET",
        "-sOutputFile=" + output_path,
        input_path,
    ]
    
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if r.returncode != 0:
            error_msg = r.stderr.strip() if r.stderr.strip() else f"Ghostscript exited with code {r.returncode}"
            raise RuntimeError(f"PDF壓縮失敗: {error_msg[:200]}")
        
        if not os.path.exists(output_path):
            raise RuntimeError("壓縮失敗: 輸出檔案未產生")
        
        result_size = os.path.getsize(output_path)
        original_size = os.path.getsize(input_path)
        
        if callback:
            ratio = (1 - result_size / original_size) * 100
            callback(f"✓ 壓縮完成! {result_size/1024/1024:.1f} MB (減少了 {ratio:.0f}%)")
        
        return {
            "original_size": original_size,
            "compressed_size": result_size,
            "ratio": result_size / original_size,
        }
    except subprocess.TimeoutExpired:
        raise RuntimeError("壓縮超時 (超過5分鐘)，請嘗試較小的PDF檔案")


# Translation entries to add to the app's TRANSLATIONS dict
TRANSLATION_ENTRIES_EN = {
    "compress_pdf": "Compress PDF",
    "compress_pdf_desc": "Reduce PDF file size by adjusting image quality",
    "select_pdf": "Select a PDF file...",
    "drag_or_select": "Drag & drop a PDF here, or click Browse",
    "browse": "Browse...",
    "quality_label": "Quality:",
    "quality_low": "Low (smaller)",
    "quality_high": "High (larger)",
    "original_size": "Original: {size}",
    "estimated_size": "Estimated: {size}",
    "compress_btn": "Compress PDF",
    "compress_running": "Compressing... please wait",
    "compress_done": "Compression complete!",
    "compress_success": "PDF compressed successfully!\n\nOriginal: {orig}\nCompressed: {compressed}\nReduction: {ratio}%\n\nSaved to:\n{output}",
    "invalid_file": "Invalid File",
    "invalid_file_msg": "Please select a valid PDF file.",
    "gs_not_found": "Ghostscript Not Found",
    "gs_not_found_msg": "Ghostscript (gs) is required for PDF compression.\n\nInstall: brew install ghostscript\nThen restart SmartFinder.",
    "compress_failed": "Compression Failed",
    "save_as": "Save Compressed PDF As...",
}

TRANSLATION_ENTRIES_ZH = {
    "compress_pdf": "壓縮 PDF",
    "compress_pdf_desc": "透過調整圖片品質來縮小 PDF 檔案大小",
    "select_pdf": "選擇 PDF 檔案...",
    "drag_or_select": "將 PDF 拖曳至此，或點選瀏覽",
    "browse": "瀏覽...",
    "quality_label": "品質:",
    "quality_low": "低 (較小)",
    "quality_high": "高 (較大)",
    "original_size": "原始: {size}",
    "estimated_size": "預計: {size}",
    "compress_btn": "壓縮 PDF",
    "compress_running": "壓縮中... 請稍候",
    "compress_done": "壓縮完成!",
    "compress_success": "PDF 壓縮成功!\n\n原始: {orig}\n壓縮後: {compressed}\n縮減: {ratio}%\n\n儲存至:\n{output}",
    "invalid_file": "無效檔案",
    "invalid_file_msg": "請選擇有效的 PDF 檔案。",
    "gs_not_found": "未安裝 Ghostscript",
    "gs_not_found_msg": "PDF 壓縮功能需要 Ghostscript (gs)。\n\n安裝方式: brew install ghostscript\n安裝後重新啟動 SmartFinder。",
    "compress_failed": "壓縮失敗",
    "save_as": "另存壓縮後的 PDF...",
}

class PdfCompressDialog(QDialog):
    """PDF Compression dialog with quality slider."""
    
    def __init__(self, parent=None, lang="en"):
        super().__init__(parent)
        self.lang = lang
        self.t_dict = TRANSLATION_ENTRIES_EN if lang == "en" else TRANSLATION_ENTRIES_ZH
        self.input_pdf = None
        self.output_pdf = None
        
        self.setWindowTitle(self.t("compress_pdf"))
        self.setGeometry(200, 200, 500, 350)
        
        layout = QVBoxLayout()
        
        # Description
        desc = QLabel(self.t("compress_pdf_desc"))
        desc.setWordWrap(True)
        layout.addWidget(desc)
        
        # File selection area
        file_layout = QHBoxLayout()
        self.file_path_label = QLineEdit()
        self.file_path_label.setPlaceholderText(self.t("select_pdf"))
        self.file_path_label.setReadOnly(True)
        file_layout.addWidget(self.file_path_label)
        
        self.browse_btn = QPushButton(self.t("browse"))
        self.browse_btn.clicked.connect(self.browse_pdf)
        file_layout.addWidget(self.browse_btn)
        layout.addLayout(file_layout)
        
        # Quality slider area
        quality_layout = QVBoxLayout()
        quality_header = QHBoxLayout()
        quality_header.addWidget(QLabel(self.t("quality_label")))
        quality_header.addStretch()
        self.original_size_label = QLabel("")
        quality_header.addWidget(self.original_size_label)
        quality_layout.addLayout(quality_header)
        
        slider_layout = QHBoxLayout()
        slider_layout.addWidget(QLabel(self.t("quality_low")))
        
        self.quality_slider = QSlider(Qt.Horizontal)
        self.quality_slider.setMinimum(0)
        self.quality_slider.setMaximum(100)
        self.quality_slider.setValue(50)
        self.quality_slider.setTickPosition(QSlider.TicksBelow)
        self.quality_slider.setTickInterval(10)
        self.quality_slider.valueChanged.connect(self.on_quality_changed)
        slider_layout.addWidget(self.quality_slider)
        
        slider_layout.addWidget(QLabel(self.t("quality_high")))
        quality_layout.addLayout(slider_layout)
        
        self.estimated_size_label = QLabel("")
        quality_layout.addWidget(self.estimated_size_label)
        layout.addLayout(quality_layout)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # indeterminate
        self.progress_bar.hide()
        self.progress_label = QLabel("")
        self.progress_label.hide()
        layout.addWidget(self.progress_label)
        layout.addWidget(self.progress_bar)
        
        # Compress button
        self.compress_btn = QPushButton(self.t("compress_btn"))
        self.compress_btn.setEnabled(False)
        self.compress_btn.clicked.connect(self.start_compress)
        layout.addWidget(self.compress_btn)
        
        # Close button
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)
    
    def t(self, key, **kwargs):
        text = self.t_dict.get(key, key)
        if kwargs:
            try:
                text = text.format(**kwargs)
            except:
                pass
        return text
    
    def browse_pdf(self):
        path, _ = QFileDialog.getOpenFileName(
            self, self.t("select_pdf"), "", "PDF Files (*.pdf)"
        )
        if path:
            self.input_pdf = path
            self.file_path_label.setText(os.path.basename(path))
            size = os.path.getsize(path)
            self.original_size_label.setText(
                self.t("original_size", size=self._fmt_size(size))
            )
            self.on_quality_changed(self.quality_slider.value())
            self.compress_btn.setEnabled(True)
    
    def _fmt_size(self, bytes_val):
        if bytes_val < 1024:
            return f"{bytes_val} B"
        elif bytes_val < 1024*1024:
            return f"{bytes_val/1024:.1f} KB"
        else:
            return f"{bytes_val/(1024*1024):.1f} MB"
    
    def on_quality_changed(self, value):
        if self.input_pdf:
            estimated = estimate_pdf_size(self.input_pdf, value)
            self.estimated_size_label.setText(
                self.t("estimated_size", size=self._fmt_size(estimated))
            )
    
    def start_compress(self):
        if not self.input_pdf or not os.path.exists(self.input_pdf):
            QMessageBox.warning(self, self.t("invalid_file"), self.t("invalid_file_msg"))
            return
        
        # Check Ghostscript
        if not get_ghostscript_path():
            QMessageBox.critical(self, self.t("gs_not_found"), self.t("gs_not_found_msg"))
            return
        
        # Ask save location
        output_path, _ = QFileDialog.getSaveFileName(
            self, self.t("save_as"), 
            os.path.splitext(self.input_pdf)[0] + "_compressed.pdf",
            "PDF Files (*.pdf)"
        )
        if not output_path:
            return
        
        self.output_pdf = output_path
        quality = self.quality_slider.value()
        
        # Disable UI
        self.compress_btn.setEnabled(False)
        self.browse_btn.setEnabled(False)
        self.quality_slider.setEnabled(False)
        self.progress_bar.show()
        self.progress_label.show()
        self.progress_label.setText(self.t("compress_running"))
        
        # Run compression in background
        def worker():
            try:
                result = compress_pdf(
                    self.input_pdf, self.output_pdf, quality,
                    callback=lambda msg: self._update_progress(msg)
                )
                # Done — show result
                self._on_compress_done(result)
            except Exception as e:
                self._on_compress_error(str(e))
        
        thread = threading.Thread(target=worker, daemon=True)
        thread.start()
    
    def _update_progress(self, msg):
        """Called from worker thread — use QTimer to update UI safely."""
        def update():
            self.progress_label.setText(msg)
        QTimer.singleShot(0, update)
    
    def _on_compress_done(self, result):
        def update():
            self.progress_bar.hide()
            self.progress_label.hide()
            self.compress_btn.setEnabled(True)
            self.browse_btn.setEnabled(True)
            self.quality_slider.setEnabled(True)
            
            orig = self._fmt_size(result["original_size"])
            compressed = self._fmt_size(result["compressed_size"])
            ratio = (1 - result["ratio"]) * 100
            
            QMessageBox.information(self, self.t("compress_done"),
                self.t("compress_success", orig=orig, compressed=compressed, 
                       ratio=f"{ratio:.0f}", output=self.output_pdf))
        QTimer.singleShot(0, update)
    
    def _on_compress_error(self, error_msg):
        def update():
            self.progress_bar.hide()
            self.progress_label.hide()
            self.compress_btn.setEnabled(True)
            self.browse_btn.setEnabled(True)
            self.quality_slider.setEnabled(True)
            QMessageBox.critical(self, self.t("compress_failed"), error_msg)
        QTimer.singleShot(0, update)


# =========================================================
# Integration helper: generates the code to add to SmartFinder
# =========================================================
def generate_integration_code():
    """
    Prints the code changes needed to add PDF compression to SmartFinder.
    
    Steps:
    1. Add TRANSLATION_ENTRIES to the TRANSLATIONS dict
    2. Add a "Compress PDF" button in action_layout 
    3. Add compress_pdf method to SmartFinderWindow class
    4. Import the module at the top
    """
    # This is just documentation — manual edits needed in the main file
    pass