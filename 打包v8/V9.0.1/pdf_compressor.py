#!/usr/bin/env python3
"""
SmartFinder v9 — PDF Compressor module
Integrated into SmartFinder as a separate dialog/feature.
Requires Ghostscript installed: brew install ghostscript

v2: Uses QProcess instead of threading for proper UI feedback.
"""
import os, subprocess, sys
from pathlib import Path
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
    QSlider, QFileDialog, QMessageBox, QProgressBar, QLineEdit
)
from PyQt5.QtCore import Qt, QProcess, QTimer


def get_ghostscript_path():
    """Find Ghostscript binary — check bundled first, then system paths."""
    # 1. Check if bundled inside the .app
    candidates = [
        os.path.join(os.path.dirname(sys.argv[0]), "gs_bundle", "gs-wrapper.sh"),
        os.path.join(os.path.dirname(sys.argv[0]), "gs_bundle", "gs"),
        # 2. Check system paths
        "/opt/homebrew/bin/gs",
        "/usr/local/bin/gs",
        "/usr/bin/gs",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    # 3. Try which command
    try:
        r = subprocess.run(["which", "gs"], capture_output=True, text=True, timeout=5)
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout.strip()
    except:
        pass
    return None


def estimate_pdf_size(pdf_path, quality_pct):
    """Estimate output size based on quality percentage."""
    try:
        size_bytes = os.path.getsize(pdf_path)
        # Map percentage to expected compression ratio (based on empirical tests)
        # These ratios were measured on a 43MB image-heavy PDF:
        # DPI 200 -> 34%, 150 -> 33%, 100 -> 30%, 85 -> 11%, 72 -> 8%, 60 -> 7%, 45 -> 4%
        if quality_pct >= 90:
            ratio = 0.35
        elif quality_pct >= 75:
            ratio = 0.34
        elif quality_pct >= 60:
            ratio = 0.30
        elif quality_pct >= 40:
            ratio = 0.11
        elif quality_pct >= 20:
            ratio = 0.08
        else:
            ratio = 0.07
        estimated = int(size_bytes * ratio)
        return estimated
    except:
        return 0


def _fmt_size(bytes_val):
    if bytes_val < 1024:
        return f"{bytes_val} B"
    elif bytes_val < 1024*1024:
        return f"{bytes_val/1024:.1f} KB"
    else:
        return f"{bytes_val/(1024*1024):.1f} MB"


TRANSLATION_ENTRIES_EN = {
    "compress_pdf": "Compress PDF",
    "compress_pdf_desc": "Reduce PDF file size by adjusting image quality",
    "select_pdf": "Select a PDF file...",
    "quality_label": "Quality:",
    "quality_low": "Low (smaller)",
    "quality_high": "High (larger)",
    "original_size": "Original: {size}",
    "estimated_size": "Estimated: {size}",
    "compress_btn": "Compress PDF",
    "compress_running": "Compressing... {info}",
    "compress_done": "Compression Complete!",
    "compress_success": "PDF compressed successfully!\n\nOriginal: {orig}\nCompressed: {compressed}\nReduction: {ratio}%\n\nSaved to:\n{output}",
    "invalid_file": "Invalid File",
    "invalid_file_msg": "Please select a valid PDF file.",
    "gs_not_found": "Ghostscript Not Found",
    "gs_not_found_msg": "Ghostscript (gs) is required for PDF compression.\n\nInstall: brew install ghostscript\nThen restart SmartFinder.",
    "compress_failed": "Compression Failed",
    "save_as": "Save Compressed PDF As...",
    "cancel": "Cancel",
    "cancel_compress": "Cancelling...",
    "compress_cancelled": "Compression cancelled.",
}

TRANSLATION_ENTRIES_ZH = {
    "compress_pdf": "壓縮 PDF",
    "compress_pdf_desc": "透過調整圖片品質來縮小 PDF 檔案大小",
    "select_pdf": "選擇 PDF 檔案...",
    "quality_label": "品質:",
    "quality_low": "低 (較小)",
    "quality_high": "高 (較大)",
    "original_size": "原始: {size}",
    "estimated_size": "預計: {size}",
    "compress_btn": "壓縮 PDF",
    "compress_running": "壓縮中... {info}",
    "compress_done": "壓縮完成!",
    "compress_success": "PDF 壓縮成功!\n\n原始: {orig}\n壓縮後: {compressed}\n縮減: {ratio}%\n\n儲存至:\n{output}",
    "invalid_file": "無效檔案",
    "invalid_file_msg": "請選擇有效的 PDF 檔案。",
    "gs_not_found": "未安裝 Ghostscript",
    "gs_not_found_msg": "PDF 壓縮功能需要 Ghostscript (gs)。\n\n安裝方式: brew install ghostscript\n安裝後重新啟動 SmartFinder。",
    "compress_failed": "壓縮失敗",
    "save_as": "另存壓縮後的 PDF...",
    "cancel": "取消",
    "cancel_compress": "取消中...",
    "compress_cancelled": "壓縮已取消。",
}


class PdfCompressDialog(QDialog):
    """PDF Compression dialog with quality slider."""
    
    def __init__(self, parent=None, lang="en"):
        super().__init__(parent)
        self.lang = lang
        self.t_dict = TRANSLATION_ENTRIES_EN if lang == "en" else TRANSLATION_ENTRIES_ZH
        self.input_pdf = None
        self.output_pdf = None
        self.process = None
        self.cancelled = False
        
        self.setWindowTitle(self.t("compress_pdf"))
        self.setGeometry(200, 200, 520, 380)
        
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
        
        # Progress area
        self.progress_label = QLabel("")
        self.progress_label.hide()
        layout.addWidget(self.progress_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # indeterminate
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)
        
        # Button row
        btn_layout = QHBoxLayout()
        self.compress_btn = QPushButton(self.t("compress_btn"))
        self.compress_btn.setEnabled(False)
        self.compress_btn.clicked.connect(self.start_compress)
        btn_layout.addWidget(self.compress_btn)
        
        self.cancel_btn = QPushButton(self.t("cancel"))
        self.cancel_btn.clicked.connect(self.cancel_compress)
        self.cancel_btn.hide()
        btn_layout.addWidget(self.cancel_btn)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)
        
        layout.addLayout(btn_layout)
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
                self.t("original_size", size=_fmt_size(size))
            )
            self.on_quality_changed(self.quality_slider.value())
            self.compress_btn.setEnabled(True)
    
    def on_quality_changed(self, value):
        if self.input_pdf:
            estimated = estimate_pdf_size(self.input_pdf, value)
            self.estimated_size_label.setText(
                self.t("estimated_size", size=_fmt_size(estimated))
            )
    
    def start_compress(self):
        if not self.input_pdf or not os.path.exists(self.input_pdf):
            QMessageBox.warning(self, self.t("invalid_file"), self.t("invalid_file_msg"))
            return
        
        gs_path = get_ghostscript_path()
        if not gs_path:
            QMessageBox.critical(self, self.t("gs_not_found"), self.t("gs_not_found_msg"))
            return
        
        output_path, _ = QFileDialog.getSaveFileName(
            self, self.t("save_as"), 
            os.path.splitext(self.input_pdf)[0] + "_compressed.pdf",
            "PDF Files (*.pdf)"
        )
        if not output_path:
            return
        
        self.output_pdf = output_path
        self.cancelled = False
        quality = self.quality_slider.value()
        
        # Map quality to resolution (NOT using dPDFSETTINGS — that loses images on macOS GS)
        # macOS Ghostscript 10.07 has a broken downsample filter that causes
        # -dPDFSETTINGS to silently drop image content. We use explicit params instead.
        if quality >= 90:
            dpi = 200
        elif quality >= 75:
            dpi = 150
        elif quality >= 60:
            dpi = 100
        elif quality >= 40:
            dpi = 85
        elif quality >= 20:
            dpi = 72
        else:
            dpi = 60
        
        # Build command — explicit image params, NO dPDFSETTINGS
        cmd = [
            gs_path,
            "-sDEVICE=pdfwrite",
            "-dNOPAUSE",
            "-dBATCH",
            "-dCompatibilityLevel=1.5",
            "-dUseFlateCompression=true",
            "-dNOINTERPOLATE",
            f"-dColorImageResolution={dpi}",
            f"-dGrayImageResolution={dpi}",
            f"-dMonoImageResolution={dpi}",
            "-dDownsampleColorImages=true",
            "-dDownsampleGrayImages=true",
            "-dDownsampleMonoImages=true",
            "-dAutoFilterColorImages=true",
            "-dAutoFilterGrayImages=true",
            "-sOutputFile=" + output_path,
            self.input_pdf,
        ]
        
        # Update UI
        self.compress_btn.setEnabled(False)
        self.browse_btn.setEnabled(False)
        self.quality_slider.setEnabled(False)
        self.cancel_btn.show()
        self.progress_bar.show()
        self.progress_label.show()
        self.progress_label.setText(
            self.t("compress_running", info=f"{quality}% / {dpi} DPI")
        )
        
        # Use QProcess for non-blocking execution
        self.process = QProcess(self)
        self.process.setProcessChannelMode(QProcess.ForwardedOutputChannel)
        self.process.readyReadStandardError.connect(self._on_process_stderr)
        self.process.finished.connect(self._on_process_finished)
        self.process.errorOccurred.connect(self._on_process_error)
        
        # Start the timer to show elapsed time
        self.elapsed = 0
        self.elapsed_timer = QTimer(self)
        self.elapsed_timer.timeout.connect(self._update_elapsed)
        self.elapsed_timer.start(2000)  # every 2 seconds
        
        self.process.start(cmd[0], cmd[1:])
    
    def cancel_compress(self):
        if self.process and self.process.state() == QProcess.Running:
            self.cancelled = True
            self.cancel_btn.setEnabled(False)
            self.cancel_btn.setText(self.t("cancel_compress"))
            self.process.kill()
    
    def _on_process_stderr(self):
        """GS outputs progress info on stderr."""
        data = self.process.readAllStandardError().data().decode('utf-8', errors='replace')
        # Look for progress info in GS output
        if data.strip():
            # Show last meaningful line from GS
            lines = [l.strip() for l in data.split('\n') if l.strip()]
            if lines:
                last = lines[-1][:80]
                self.progress_label.setText(
                    self.t("compress_running", info=last)
                )
    
    def _update_elapsed(self):
        """Update elapsed time every 2 seconds."""
        self.elapsed += 2
        if self.process and self.process.state() == QProcess.Running:
            self.progress_label.setText(
                self.t("compress_running", info=f"已過 {self.elapsed} 秒...")
            )
    
    def _on_process_finished(self, exit_code, exit_status):
        self.elapsed_timer.stop()
        self.process = None
        
        self.progress_bar.hide()
        self.cancel_btn.hide()
        self.compress_btn.setEnabled(True)
        self.browse_btn.setEnabled(True)
        self.quality_slider.setEnabled(True)
        
        if self.cancelled:
            self.progress_label.hide()
            QMessageBox.information(self, self.t("compress_cancelled"), "")
            return
        
        if exit_code != 0 or not os.path.exists(self.output_pdf):
            self.progress_label.hide()
            QMessageBox.critical(self, self.t("compress_failed"), 
                f"Ghostscript exited with code {exit_code}\n輸出檔案未產生。")
            return
        
        result_size = os.path.getsize(self.output_pdf)
        original_size = os.path.getsize(self.input_pdf)
        ratio = (1 - result_size / original_size) * 100
        
        self.progress_label.hide()
        QMessageBox.information(self, self.t("compress_done"),
            self.t("compress_success", orig=_fmt_size(original_size),
                   compressed=_fmt_size(result_size), ratio=f"{ratio:.0f}",
                   output=self.output_pdf))
    
    def _on_process_error(self, error):
        self.elapsed_timer.stop()
        self.process = None
        self.progress_bar.hide()
        self.cancel_btn.hide()
        self.compress_btn.setEnabled(True)
        self.browse_btn.setEnabled(True)
        self.quality_slider.setEnabled(True)
        self.progress_label.hide()
        
        error_map = {
            QProcess.FailedToStart: "Ghostscript 無法啟動，請確認安裝狀態",
            QProcess.Crashed: "Ghostscript 意外崩潰",
            QProcess.Timedout: "GHostscript 超時",
            QProcess.WriteError: "寫入錯誤",
            QProcess.ReadError: "讀取錯誤",
        }
        msg = error_map.get(error, f"未知錯誤 ({error})")
        QMessageBox.critical(self, self.t("compress_failed"), msg)