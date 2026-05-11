import os
import sys
import json
from datetime import datetime
import mimetypes
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QTreeWidget, QTreeWidgetItem,
                             QFileDialog, QMessageBox, QComboBox, QRadioButton, QDialog,
                             QSlider)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QIcon, QPixmap, QFontMetrics


# =========================================================
#  Internationalization (i18n)
# =========================================================
TRANSLATIONS = {
    "en": {
        # Window / labels
        "window_title": "Smart Finder",
        "target_dir": "Target Directory:",
        "target_filename": "Target Filename:",
        "search_btn": "Confirm & Search",
        "col_name": "File Name",
        "col_kind": "Kind",
        "col_modified": "Date Modified",
        "col_size": "Size",
        "col_location": "Location",
        "result_count": "Search Result: {count} file(s)",
        "selected_count": "Selected: {count} file(s)",
        "open_files": "Open Files",
        "open_location": "Open Location",
        "rename_files": "Batch Rename",
        "change_save_location": "Change Save Location",
        "close_app": "Close",
        "app_label": "Smart Finder @ Xeson",
        "language_label": "Language:",
        "lang_en": "EN",
        "lang_zh": "中文",
        # Dialogs
        "rename_dialog_title": "Batch Rename",
        "rename_prefix": "Add Prefix",
        "rename_suffix": "Add Suffix",
        "rename_full": "Full Rename",
        "apply": "Apply",
        # Messages
        "missing_info_title": "Missing Information",
        "missing_info_msg": "Please enter a target directory and filename.",
        "open_many_title": "Opening Multiple Files",
        "open_many_msg": "❗️Warning❗️\n\nYou have selected {n} files,\nare you sure you want to open them all at once?\n\nNot responsible for app freezing.",
        "open_many_folder_title": "Opening Multiple Folders",
        "open_many_folder_msg": "❗️Warning❗️\n\nYour selected files are located in\n {n} different folders,\nare you sure you want to open them all at once?\n\nNot responsible for app freezing.",
        "no_selection_title": "No Files Selected",
        "no_selection_open_msg": "Please select files to open first.",
        "no_selection_loc_msg": "Please select files to open their locations.",
        "no_selection_move_msg": "Please select files whose location you want to change.",
        "no_selection_rename_msg": "Please select files to rename first.",
        "select_new_location": "Select New Save Location",
        "move_failed_title": "Move Failed",
        "move_failed_msg": "Failed to move file {name}:\n{err}",
        "move_partial_title": "Some Files Not Moved",
        "move_partial_msg": "The following files were not moved\nbecause duplicates already exist at the destination:\n\n{files}",
        "move_done_title": "Move Complete",
        "move_done_msg": "The following files have been moved\nto the new location:\n\n{files}\n\nNew location:\n{loc}",
        "move_all_failed_title": "Move Failed",
        "move_all_failed_msg": "None of the selected files were moved.\n\nPlease check whether duplicates exist at the destination.",
        "rename_confirm_title": "Confirm Batch Rename",
        "rename_confirm_header": "The following files will be renamed:\n\n",
        "rename_dup_title": "Cannot Rename",
        "rename_dup_msg": "Cannot rename — please choose a non-duplicate filename.",
        "rename_done_title": "Rename Complete",
        "rename_done_msg": "Batch rename completed!",
        "change_loc_q_title": "Change Save Location",
        "change_loc_q_msg": "Do you want to change the save location of the renamed files?",
        "missing_new_name_title": "Missing Information",
        "missing_new_name_msg": "Please enter a new filename.",
    },
    "zh": {
        "window_title": "Smart Finder",
        "target_dir": "目標地址:",
        "target_filename": "目標檔案名稱:",
        "search_btn": "確定及搜索",
        "col_name": "檔案名稱",
        "col_kind": "類型",
        "col_modified": "修改日期",
        "col_size": "大小",
        "col_location": "位置",
        "result_count": "搜尋結果: {count} 個檔案",
        "selected_count": "選擇的檔案: {count} 個",
        "open_files": "打開檔案",
        "open_location": "打開路徑",
        "rename_files": "批量改名",
        "change_save_location": "更改存放位置",
        "close_app": "關閉程式",
        "app_label": "Smart Finder @ Xeson",
        "language_label": "語言:",
        "lang_en": "EN",
        "lang_zh": "中文",
        "rename_dialog_title": "批量改名",
        "rename_prefix": "加入 Prefix",
        "rename_suffix": "加入 Suffix",
        "rename_full": "全名修改",
        "apply": "套用",
        "missing_info_title": "缺少資訊",
        "missing_info_msg": "請輸入目標地址和目標檔案名稱.",
        "open_many_title": "打開多個檔案",
        "open_many_msg": "❗️警告❗️\n\n您選擇了 {n} 個檔案,\n確定要一次打開所有檔案嗎?\n\n卡死不負責。",
        "open_many_folder_title": "打開多個文件夾",
        "open_many_folder_msg": "❗️警告❗️\n\n您選擇的檔案位於\n {n} 個不同的文件夾,\n確定要一次打開所有文件夾?\n\n卡死不負責。",
        "no_selection_title": "未選取檔案",
        "no_selection_open_msg": "請先選取要開啟的檔案。",
        "no_selection_loc_msg": "請先選取要開啟路徑的檔案。",
        "no_selection_move_msg": "請先選取要更改存放位置的檔案。",
        "no_selection_rename_msg": "請先選取要進行批量改名的檔案。",
        "select_new_location": "選擇新的存放位置",
        "move_failed_title": "移動失敗",
        "move_failed_msg": "移動檔案 {name} 失敗:\n{err}",
        "move_partial_title": "移動部分失敗",
        "move_partial_msg": "以下檔案由\n於目標位置已有重複檔名而\n未被移動:\n\n{files}",
        "move_done_title": "移動完成",
        "move_done_msg": "以下檔案已成功移動\n到新的存放位置:\n\n{files}\n\n新的存放位置: \n{loc}",
        "move_all_failed_title": "移動失敗",
        "move_all_failed_msg": "所有選取的檔案\n都未被移動,\n\n請檢查目標位置是否有重複檔名。",
        "rename_confirm_title": "確認批量改名",
        "rename_confirm_header": "以下是將要進行改名的檔案及其新檔名：\n\n",
        "rename_dup_title": "無法改名",
        "rename_dup_msg": "無法改名\n請改一個沒有重複的檔案名稱",
        "rename_done_title": "批量改名完成",
        "rename_done_msg": "檔案批量改名已完成！",
        "change_loc_q_title": "改變存放位置",
        "change_loc_q_msg": "是否要改變已改名檔案的存放位置？",
        "missing_new_name_title": "缺少資訊",
        "missing_new_name_msg": "請輸入新的檔案名稱。",
    },
}

# 設定檔（保存語言偏好）
SETTINGS_FILE = "Smart_Finder_settings.json"
RECENT_DIRS_FILE = "Smart_Finder_recent_directories.json"


def resource_path(relative_path):
    """獲取資源的絕對路徑，相容開發環境與 PyInstaller 打包後的環境"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))

    candidates = [
        os.path.join(base_path, relative_path),
        os.path.join(base_path, "Icon", os.path.basename(relative_path)),
        os.path.join(os.path.abspath("."), relative_path),
        os.path.join(os.path.abspath("."), "Icon", os.path.basename(relative_path)),
        os.path.join(os.path.abspath(os.path.dirname(__file__)), "Icon", os.path.basename(relative_path)),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return os.path.join(base_path, relative_path)


def load_settings():
    """載入設定檔；如果不存在則返回預設值（英文）。"""
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
    except Exception:
        pass
    return {"language": "en"}


def save_settings(settings):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False)
    except Exception:
        pass


def get_file_kind(file_path):
    if os.path.isdir(file_path):
        return "Folder"
    else:
        file_type, _ = mimetypes.guess_type(file_path)
        if file_type:
            file_kind = file_type.split("/")[-1].upper()
            if file_kind == "VND.OPENXMLFORMATS-OFFICEDOCUMENT.WORDPROCESSINGML.DOCUMENT":
                return "Word"
            elif file_kind == "doc":
                return "Word-old"
            elif file_kind == "VND.OPENXMLFORMATS-OFFICEDOCUMENT.SPREADSHEETML.SHEET":
                return "Excel"
            elif file_kind == "VND.MS-EXCEL.SHEET.MACROENABLED.12":
                return "Excel-MacroEnable"
            elif file_kind == "QUICKTIME":
                return "MOV"
            elif file_kind == "MPEG":
                if file_path.lower().endswith(".mp3"):
                    return "Mp3 Audio"
                else:
                    return file_kind
            elif file_kind == "VND.ADOBE.PHOTOSHOP":
                return "Adobe Photoshop"
            elif file_kind == "POSTSCRIPT":
                if file_path.lower().endswith(".ai"):
                    return "Adobe AI"
                else:
                    return file_kind
            else:
                return file_kind
        else:
            _, extension = os.path.splitext(file_path)
            return extension[1:].upper()


def format_file_size(size_bytes):
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    elif size_bytes < 1024 * 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024 * 1024):.1f} TB"


def search_files(directory, target_filename):
    matching_files = []
    target_filename = target_filename.lower()

    try:
        for root, dirs, files in os.walk(directory):
            try:
                for file in files:
                    if target_filename in file.lower():
                        file_path = os.path.join(root, file)
                        matching_files.append(file_path)
            except (PermissionError, OSError):
                continue
    except (PermissionError, OSError):
        pass

    return matching_files


def load_recent_directories():
    try:
        if os.path.exists(RECENT_DIRS_FILE):
            with open(RECENT_DIRS_FILE, "r") as file:
                return json.load(file)
    except (IOError, json.JSONDecodeError, Exception):
        pass
    return []


def save_recent_directories(directories):
    try:
        with open(RECENT_DIRS_FILE, "w") as file:
            json.dump(directories, file)
    except (IOError, Exception):
        pass


def update_recent_directories(directory):
    try:
        recent_directories = load_recent_directories()
        if directory in recent_directories:
            recent_directories.remove(directory)
        recent_directories.insert(0, directory)
        recent_directories = recent_directories[:5]
        save_recent_directories(recent_directories)
    except Exception:
        pass


class RenameDialog(QDialog):
    def __init__(self, parent=None, lang="en"):
        try:
            super().__init__(parent)
            self.lang = lang
            t = TRANSLATIONS[lang]
            self.setWindowTitle(t["rename_dialog_title"])
            self.setGeometry(200, 200, 320, 160)

            layout = QVBoxLayout()

            self.prefix_radio = QRadioButton(t["rename_prefix"])
            self.suffix_radio = QRadioButton(t["rename_suffix"])
            self.full_rename_radio = QRadioButton(t["rename_full"])
            self.prefix_radio.setChecked(True)

            self.rename_entry = QLineEdit()
            self.apply_button = QPushButton(t["apply"])

            layout.addWidget(self.prefix_radio)
            layout.addWidget(self.suffix_radio)
            layout.addWidget(self.full_rename_radio)
            layout.addWidget(self.rename_entry)
            layout.addWidget(self.apply_button)

            self.setLayout(layout)

            self.apply_button.clicked.connect(self.accept)
        except Exception:
            super().__init__(parent)
            self.setWindowTitle("Rename")

    def get_rename_option(self):
        try:
            if self.prefix_radio.isChecked():
                return "prefix"
            elif self.suffix_radio.isChecked():
                return "suffix"
            else:
                return "full"
        except Exception:
            return "prefix"

    def get_new_name(self):
        try:
            return self.rename_entry.text()
        except Exception:
            return ""


class SmartFinderWindow(QMainWindow):
    def __init__(self):
        try:
            super().__init__()

            # 載入語言設定（預設為英文）
            self.settings = load_settings()
            self.lang = self.settings.get("language", "en")
            if self.lang not in TRANSLATIONS:
                self.lang = "en"

            self.setGeometry(100, 100, 1000, 600)

            central_widget = QWidget()
            self.setCentralWidget(central_widget)
            main_layout = QVBoxLayout(central_widget)

            # Window Icon
            try:
                window_icon_path = resource_path("program_icon.png")
                if os.path.exists(window_icon_path):
                    self.setWindowIcon(QIcon(window_icon_path))
            except Exception:
                pass

            # Target directory
            target_dir_layout = QHBoxLayout()
            self.target_dir_label = QLabel()
            self.target_dir_entry = QComboBox()
            self.target_dir_entry.setEditable(True)
            self.populate_recent_directories()
            target_dir_layout.addWidget(self.target_dir_label)
            target_dir_layout.addWidget(self.target_dir_entry)
            main_layout.addLayout(target_dir_layout)

            # Target filename
            target_filename_layout = QHBoxLayout()
            self.target_filename_label = QLabel()
            self.target_filename_entry = QLineEdit()
            target_filename_layout.addWidget(self.target_filename_label)
            target_filename_layout.addWidget(self.target_filename_entry)
            main_layout.addLayout(target_filename_layout)

            # Search button
            self.select_button = QPushButton()
            self.select_button.clicked.connect(self.select_files)
            main_layout.addWidget(self.select_button)

            # Result list
            self.result_listbox = QTreeWidget()
            self.result_listbox.setColumnCount(5)
            self.result_listbox.setColumnWidth(0, 350)
            self.result_listbox.setColumnWidth(1, 100)
            self.result_listbox.setColumnWidth(2, 150)
            self.result_listbox.setColumnWidth(3, 55)
            self.result_listbox.setColumnWidth(4, 350)
            self.result_listbox.setSelectionMode(QTreeWidget.ExtendedSelection)
            self.result_listbox.itemSelectionChanged.connect(self.update_selected_count)
            main_layout.addWidget(self.result_listbox)

            # Count labels
            count_layout = QHBoxLayout()
            self.result_count_label = QLabel()
            self.selected_count_label = QLabel()
            count_layout.addWidget(self.selected_count_label)
            count_layout.addWidget(self.result_count_label)
            main_layout.addLayout(count_layout)

            # Action buttons
            action_layout = QHBoxLayout()
            self.open_button = QPushButton()
            self.open_button.clicked.connect(self.open_selected_files)
            self.open_location_button = QPushButton()
            self.open_location_button.clicked.connect(self.open_file_location)
            self.rename_button = QPushButton()
            self.rename_button.clicked.connect(self.rename_files)
            self.change_save_location_button = QPushButton()
            self.change_save_location_button.clicked.connect(self.change_save_location)
            action_layout.addWidget(self.open_button)
            action_layout.addWidget(self.open_location_button)
            action_layout.addWidget(self.rename_button)
            action_layout.addWidget(self.change_save_location_button)
            main_layout.addLayout(action_layout)

            # Close button
            self.close_button = QPushButton()
            self.close_button.clicked.connect(self.close)
            main_layout.addWidget(self.close_button)

            # ============ Footer: 左下角語言拉桿 + 右下角 App label ============
            footer_layout = QHBoxLayout()
            footer_layout.setContentsMargins(0, 0, 0, 0)
            footer_layout.setSpacing(8)

            # 左下角：語言選擇拉桿（EN <-> 中文）
            self.language_label = QLabel()
            footer_layout.addWidget(self.language_label, 0, Qt.AlignVCenter)

            self.lang_en_label = QLabel()
            footer_layout.addWidget(self.lang_en_label, 0, Qt.AlignVCenter)

            self.language_slider = QSlider(Qt.Horizontal)
            self.language_slider.setMinimum(0)   # 0 = English
            self.language_slider.setMaximum(1)   # 1 = 中文
            self.language_slider.setSingleStep(1)
            self.language_slider.setPageStep(1)
            self.language_slider.setTickPosition(QSlider.TicksBelow)
            self.language_slider.setTickInterval(1)
            self.language_slider.setFixedWidth(60)
            self.language_slider.setValue(0 if self.lang == "en" else 1)
            self.language_slider.valueChanged.connect(self.on_language_changed)
            footer_layout.addWidget(self.language_slider, 0, Qt.AlignVCenter)

            self.lang_zh_label = QLabel()
            footer_layout.addWidget(self.lang_zh_label, 0, Qt.AlignVCenter)

            # spacer
            footer_layout.addStretch()

            # 右下角：App label + icon
            try:
                self.app_icon_label = QLabel()
                self.app_label_text = QLabel()
                fm = QFontMetrics(self.app_label_text.font())
                icon_size = fm.height()
                app_icon_path = resource_path("program_icon.png")
                pix = QPixmap(app_icon_path)
                if not pix.isNull():
                    self.app_icon_label.setPixmap(
                        pix.scaled(QSize(icon_size, icon_size),
                                   Qt.KeepAspectRatio,
                                   Qt.SmoothTransformation)
                    )
                    footer_layout.addWidget(self.app_icon_label, 0, Qt.AlignVCenter)
                footer_layout.addWidget(self.app_label_text, 0, Qt.AlignVCenter)
            except Exception:
                self.app_label_text = QLabel()
                footer_layout.addWidget(self.app_label_text, 0, Qt.AlignVCenter)

            main_layout.addLayout(footer_layout)

            # 套用語言文字
            self.retranslate_ui()
        except Exception:
            super().__init__()
            self.setWindowTitle("SmartFinder - Error")

    # -----------------------------------------------------
    #  Language handling
    # -----------------------------------------------------
    def t(self, key, **kwargs):
        text = TRANSLATIONS.get(self.lang, TRANSLATIONS["en"]).get(
            key, TRANSLATIONS["en"].get(key, key)
        )
        if kwargs:
            try:
                text = text.format(**kwargs)
            except Exception:
                pass
        return text

    def on_language_changed(self, value):
        new_lang = "zh" if value == 1 else "en"
        if new_lang != self.lang:
            self.lang = new_lang
            self.settings["language"] = new_lang
            save_settings(self.settings)
            self.retranslate_ui()

    def retranslate_ui(self):
        try:
            self.setWindowTitle(self.t("window_title"))
            self.target_dir_label.setText(self.t("target_dir"))
            self.target_filename_label.setText(self.t("target_filename"))
            self.select_button.setText(self.t("search_btn"))

            self.result_listbox.setHeaderLabels([
                self.t("col_name"),
                self.t("col_kind"),
                self.t("col_modified"),
                self.t("col_size"),
                self.t("col_location"),
            ])

            # 重新整理計數
            try:
                count = self.result_listbox.invisibleRootItem().childCount()
            except Exception:
                count = 0
            self.result_count_label.setText(self.t("result_count", count=count))

            try:
                selected_count = len(self.result_listbox.selectedItems())
            except Exception:
                selected_count = 0
            self.selected_count_label.setText(self.t("selected_count", count=selected_count))

            self.open_button.setText(self.t("open_files"))
            self.open_location_button.setText(self.t("open_location"))
            self.rename_button.setText(self.t("rename_files"))
            self.change_save_location_button.setText(self.t("change_save_location"))
            self.close_button.setText(self.t("close_app"))

            self.language_label.setText(self.t("language_label"))
            self.lang_en_label.setText(self.t("lang_en"))
            self.lang_zh_label.setText(self.t("lang_zh"))

            self.app_label_text.setText(self.t("app_label"))
        except Exception:
            pass

    # -----------------------------------------------------
    #  Functionality
    # -----------------------------------------------------
    def populate_recent_directories(self):
        recent_directories = load_recent_directories()
        self.target_dir_entry.clear()
        self.target_dir_entry.addItems(recent_directories)

    def select_files(self):
        target_dir = self.target_dir_entry.currentText()
        target_filename = self.target_filename_entry.text()

        if target_dir and target_filename:
            try:
                update_recent_directories(target_dir)
                self.populate_recent_directories()
            except Exception:
                pass

            self.select_button.setEnabled(False)
            matching_files = search_files(target_dir, target_filename)

            self.result_listbox.clear()
            for file_path in matching_files:
                try:
                    if not os.path.exists(file_path):
                        continue

                    file = os.path.basename(file_path)
                    file_stats = os.stat(file_path)
                    file_size = file_stats.st_size
                    file_modified_time = datetime.fromtimestamp(file_stats.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                    file_kind = get_file_kind(file_path)

                    try:
                        relative_path = os.path.relpath(os.path.dirname(file_path), target_dir)
                    except (ValueError, OSError):
                        relative_path = os.path.dirname(file_path)

                    display_path = os.path.join("~", relative_path)
                    item = QTreeWidgetItem([file, file_kind, file_modified_time, format_file_size(file_size), display_path])
                    item.setData(0, Qt.UserRole, file_path)
                    self.result_listbox.addTopLevelItem(item)
                except (OSError, PermissionError, FileNotFoundError, ValueError):
                    continue

            result_count = self.result_listbox.invisibleRootItem().childCount()
            self.result_count_label.setText(self.t("result_count", count=result_count))
            self.select_button.setEnabled(True)
        else:
            try:
                QMessageBox.information(self, self.t("missing_info_title"), self.t("missing_info_msg"))
            except Exception:
                pass

    def update_selected_count(self):
        selected_count = len(self.result_listbox.selectedItems())
        self.selected_count_label.setText(self.t("selected_count", count=selected_count))

    def open_selected_files(self):
        try:
            selected_items = self.result_listbox.selectedItems()
            if selected_items:
                file_paths = [item.data(0, Qt.UserRole) for item in selected_items]
                if len(file_paths) > 5:
                    try:
                        reply = QMessageBox.warning(self, self.t("open_many_title"),
                                                    self.t("open_many_msg", n=len(file_paths)),
                                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                    except Exception:
                        reply = QMessageBox.No

                    if reply == QMessageBox.Yes:
                        for path in file_paths:
                            try:
                                os.system(f'open "{path}"')
                            except Exception:
                                pass
                else:
                    for path in file_paths:
                        try:
                            os.system(f'open "{path}"')
                        except Exception:
                            pass
            else:
                try:
                    QMessageBox.information(self, self.t("no_selection_title"), self.t("no_selection_open_msg"))
                except Exception:
                    pass
        except Exception:
            pass

    def open_file_location(self):
        try:
            selected_items = self.result_listbox.selectedItems()
            if selected_items:
                folder_paths = list(set([os.path.dirname(item.data(0, Qt.UserRole)) for item in selected_items]))
                if len(folder_paths) > 5:
                    try:
                        reply = QMessageBox.warning(self, self.t("open_many_folder_title"),
                                                    self.t("open_many_folder_msg", n=len(folder_paths)),
                                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                    except Exception:
                        reply = QMessageBox.No

                    if reply == QMessageBox.Yes:
                        for path in folder_paths:
                            try:
                                os.system(f'open "{path}"')
                            except Exception:
                                pass
                else:
                    for path in folder_paths:
                        try:
                            os.system(f'open "{path}"')
                        except Exception:
                            pass
            else:
                try:
                    QMessageBox.information(self, self.t("no_selection_title"), self.t("no_selection_loc_msg"))
                except Exception:
                    pass
        except Exception:
            pass

    def change_save_location(self):
        try:
            selected_items = self.result_listbox.selectedItems()
            if selected_items:
                new_location = QFileDialog.getExistingDirectory(self, self.t("select_new_location"))
                if new_location:
                    duplicate_files = []
                    moved_files = []
                    for item in selected_items:
                        try:
                            old_path = item.data(0, Qt.UserRole)
                            file_name = os.path.basename(old_path)
                            new_path = os.path.join(new_location, file_name)
                            if os.path.exists(new_path):
                                duplicate_files.append(file_name)
                            else:
                                try:
                                    os.rename(old_path, new_path)
                                    moved_files.append(file_name)
                                except Exception as e:
                                    try:
                                        QMessageBox.information(self, self.t("move_failed_title"),
                                                                self.t("move_failed_msg", name=file_name, err=str(e)))
                                    except Exception:
                                        pass
                        except Exception:
                            pass

                    if duplicate_files:
                        try:
                            duplicate_files_text = "\n".join(duplicate_files)
                            QMessageBox.information(self, self.t("move_partial_title"),
                                                    self.t("move_partial_msg", files=duplicate_files_text))
                        except Exception:
                            pass

                    if moved_files:
                        try:
                            moved_files_text = "\n".join(moved_files)
                            QMessageBox.information(self, self.t("move_done_title"),
                                                    self.t("move_done_msg", files=moved_files_text, loc=new_location))
                            self.select_files()
                        except Exception:
                            pass
                    else:
                        try:
                            QMessageBox.information(self, self.t("move_all_failed_title"),
                                                    self.t("move_all_failed_msg"))
                        except Exception:
                            pass
            else:
                try:
                    QMessageBox.information(self, self.t("no_selection_title"), self.t("no_selection_move_msg"))
                except Exception:
                    pass
        except Exception:
            pass

    def rename_files(self):
        try:
            selected_items = self.result_listbox.selectedItems()
            if selected_items:
                dialog = RenameDialog(self, lang=self.lang)
                if dialog.exec_():
                    rename_option = dialog.get_rename_option()
                    new_name = dialog.get_new_name()
                    if new_name:
                        try:
                            confirmation_text = self.t("rename_confirm_header")
                            renamed_files = []
                            selected_items_sorted = sorted(selected_items, key=lambda x: os.path.getmtime(x.data(0, Qt.UserRole)))
                            num_files = len(selected_items_sorted)
                            num_digits = len(str(num_files))
                            for index, item in enumerate(selected_items_sorted, start=1):
                                try:
                                    old_path = item.data(0, Qt.UserRole)
                                    old_name = os.path.basename(old_path)
                                    if rename_option == "prefix":
                                        new_file_name = new_name + old_name
                                    elif rename_option == "suffix":
                                        name_parts = old_name.split(".")
                                        if len(name_parts) > 1:
                                            new_file_name = ".".join(name_parts[:-1]) + new_name + "." + name_parts[-1]
                                        else:
                                            new_file_name = old_name + new_name
                                    else:
                                        ext = os.path.splitext(old_name)[1]
                                        new_file_name = f"{new_name}_{index:0{num_digits}}{ext}"
                                    confirmation_text += f"{old_name} -> {new_file_name}\n"
                                    renamed_files.append((old_path, new_file_name))
                                except Exception:
                                    pass

                            try:
                                reply = QMessageBox.question(self, self.t("rename_confirm_title"), confirmation_text,
                                                             QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                            except Exception:
                                reply = QMessageBox.No

                            if reply == QMessageBox.Yes:
                                rename_success = True
                                for old_path, new_file_name in renamed_files:
                                    try:
                                        new_path = os.path.join(os.path.dirname(old_path), new_file_name)
                                        if os.path.exists(new_path):
                                            try:
                                                QMessageBox.information(self, self.t("rename_dup_title"), self.t("rename_dup_msg"))
                                            except Exception:
                                                pass
                                            rename_success = False
                                            break
                                    except Exception:
                                        pass

                                if rename_success:
                                    for old_path, new_file_name in renamed_files:
                                        try:
                                            new_path = os.path.join(os.path.dirname(old_path), new_file_name)
                                            os.rename(old_path, new_path)
                                        except Exception:
                                            pass

                                    try:
                                        QMessageBox.information(self, self.t("rename_done_title"), self.t("rename_done_msg"))
                                    except Exception:
                                        pass

                                    try:
                                        self.select_files()
                                    except Exception:
                                        pass

                                    try:
                                        reply = QMessageBox.question(self, self.t("change_loc_q_title"), self.t("change_loc_q_msg"),
                                                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                                    except Exception:
                                        reply = QMessageBox.No

                                    if reply == QMessageBox.Yes:
                                        new_location = QFileDialog.getExistingDirectory(self, self.t("select_new_location"))
                                        if new_location:
                                            duplicate_files = []
                                            moved_files = []
                                            for old_path, new_file_name in renamed_files:
                                                try:
                                                    old_dir = os.path.dirname(old_path)
                                                    new_path = os.path.join(new_location, new_file_name)
                                                    if os.path.exists(new_path):
                                                        duplicate_files.append(new_file_name)
                                                    else:
                                                        try:
                                                            os.rename(os.path.join(old_dir, new_file_name), new_path)
                                                            moved_files.append(new_file_name)
                                                        except Exception as e:
                                                            try:
                                                                QMessageBox.information(self, self.t("move_failed_title"),
                                                                                        self.t("move_failed_msg", name=new_file_name, err=str(e)))
                                                            except Exception:
                                                                pass
                                                except Exception:
                                                    pass

                                            if duplicate_files:
                                                try:
                                                    duplicate_files_text = "\n".join(duplicate_files)
                                                    QMessageBox.information(self, self.t("move_partial_title"),
                                                                            self.t("move_partial_msg", files=duplicate_files_text))
                                                except Exception:
                                                    pass

                                            if moved_files:
                                                try:
                                                    moved_files_text = "\n".join(moved_files)
                                                    QMessageBox.information(self, self.t("move_done_title"),
                                                                            self.t("move_done_msg", files=moved_files_text, loc=new_location))
                                                except Exception:
                                                    pass
                                                try:
                                                    self.select_files()
                                                except Exception:
                                                    pass
                                            else:
                                                try:
                                                    QMessageBox.information(self, self.t("move_all_failed_title"),
                                                                            self.t("move_all_failed_msg"))
                                                except Exception:
                                                    pass
                        except Exception:
                            pass
                    else:
                        try:
                            QMessageBox.information(self, self.t("missing_new_name_title"), self.t("missing_new_name_msg"))
                        except Exception:
                            pass
            else:
                try:
                    QMessageBox.information(self, self.t("no_selection_title"), self.t("no_selection_rename_msg"))
                except Exception:
                    pass
        except Exception:
            pass


def main():
    try:
        app = QApplication(sys.argv)
        window = SmartFinderWindow()
        window.show()
        sys.exit(app.exec_())
    except Exception:
        try:
            import traceback
            traceback.print_exc()
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()
