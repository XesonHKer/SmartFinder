import os
import sys
import json
from datetime import datetime
import mimetypes
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTreeWidget, QTreeWidgetItem, 
                             QFileDialog, QMessageBox, QComboBox, QRadioButton, QDialog)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QIcon, QPixmap, QFontMetrics

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
                    return "Mp3 音訊"
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
            # 处理 os.walk 中的权限错误
            try:
                for file in files:
                    if target_filename in file.lower():
                        file_path = os.path.join(root, file)
                        matching_files.append(file_path)
            except (PermissionError, OSError):
                # 跳过无权限访问的文件夹
                continue
    except (PermissionError, OSError) as e:
        # 如果目录本身无法访问，返回空列表或显示警告
        pass
    
    return matching_files

def load_recent_directories():
    try:
        if os.path.exists("Smart_Finder_recent_directories.json"):
            with open("Smart_Finder_recent_directories.json", "r") as file:
                return json.load(file)
    except (IOError, json.JSONDecodeError, Exception):
        # 如果JSON文件损坏或无法读取，返回空列表
        pass
    return []

def save_recent_directories(directories):
    try:
        with open("Smart_Finder_recent_directories.json", "w") as file:
            json.dump(directories, file)
    except (IOError, Exception):
        # 如果无法保存，默认忽略错误
        pass

def update_recent_directories(directory):
    try:
        recent_directories = load_recent_directories()
        if directory in recent_directories:
            recent_directories.remove(directory)
        recent_directories.insert(0, directory)
        recent_directories = recent_directories[:5]  # 只保留最近的5個目錄
        save_recent_directories(recent_directories)
    except Exception:
        # 如果更新失败，默认忽略
        pass

class RenameDialog(QDialog):
    def __init__(self, parent=None):
        try:
            super().__init__(parent)
            self.setWindowTitle("批量改名")
            self.setGeometry(200, 200, 300, 150)

            layout = QVBoxLayout()

            self.prefix_radio = QRadioButton("加入 Prefix")
            self.suffix_radio = QRadioButton("加入 Suffix")
            self.full_rename_radio = QRadioButton("全名修改")
            self.prefix_radio.setChecked(True)

            self.rename_entry = QLineEdit()
            self.apply_button = QPushButton("套用")

            layout.addWidget(self.prefix_radio)
            layout.addWidget(self.suffix_radio)
            layout.addWidget(self.full_rename_radio)
            layout.addWidget(self.rename_entry)
            layout.addWidget(self.apply_button)

            self.setLayout(layout)

            self.apply_button.clicked.connect(self.accept)
        except Exception:
            # 如果初始化失败，创建最小化的对话框
            super().__init__(parent)
            self.setWindowTitle("改名")

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
            self.setWindowTitle("Smart Finder")
            self.setGeometry(100, 100, 1000, 600)

            central_widget = QWidget()
            self.setCentralWidget(central_widget)
            main_layout = QVBoxLayout(central_widget)

            # Window Icon (Dock / 視窗左上角)
            try:
                window_icon_path = resource_path("program_icon.png")
                if os.path.exists(window_icon_path):
                    self.setWindowIcon(QIcon(window_icon_path))
            except Exception:
                pass

            # Target directory
            target_dir_layout = QHBoxLayout()
            target_dir_label = QLabel("目標地址:")
            self.target_dir_entry = QComboBox()
            self.target_dir_entry.setEditable(True)
            self.populate_recent_directories()
            target_dir_layout.addWidget(target_dir_label)
            target_dir_layout.addWidget(self.target_dir_entry)
            main_layout.addLayout(target_dir_layout)

            # Target filename
            target_filename_layout = QHBoxLayout()
            target_filename_label = QLabel("目標檔案名稱:")
            self.target_filename_entry = QLineEdit()
            target_filename_layout.addWidget(target_filename_label)
            target_filename_layout.addWidget(self.target_filename_entry)
            main_layout.addLayout(target_filename_layout)

            # Search button
            self.select_button = QPushButton("確定及搜索")
            self.select_button.clicked.connect(self.select_files)
            main_layout.addWidget(self.select_button)

            # Result list
            self.result_listbox = QTreeWidget()
            self.result_listbox.setHeaderLabels(["File Name", "Kind", "Date Modified", "Size", "Location"])
            self.result_listbox.setColumnWidth(0, 350)
            self.result_listbox.setColumnWidth(1, 100)
            self.result_listbox.setColumnWidth(2, 150)
            self.result_listbox.setColumnWidth(3, 55)
            self.result_listbox.setColumnWidth(4, 350)
            self.result_listbox.setSelectionMode(QTreeWidget.ExtendedSelection)  # Allow multiple selections
            self.result_listbox.itemSelectionChanged.connect(self.update_selected_count)
            main_layout.addWidget(self.result_listbox)

            # Count labels
            count_layout = QHBoxLayout()
            self.result_count_label = QLabel("搜尋結果: 0 個檔案")
            self.selected_count_label = QLabel("選擇的檔案: 0 個")
            count_layout.addWidget(self.selected_count_label)
            count_layout.addWidget(self.result_count_label)
            main_layout.addLayout(count_layout)

            # Action buttons
            action_layout = QHBoxLayout()
            open_button = QPushButton("打開檔案")
            open_button.clicked.connect(self.open_selected_files)
            open_location_button = QPushButton("打開路徑")
            open_location_button.clicked.connect(self.open_file_location)
            rename_button = QPushButton("批量改名")
            rename_button.clicked.connect(self.rename_files)
            change_save_location_button = QPushButton("更改存放位置")
            change_save_location_button.clicked.connect(self.change_save_location)
            action_layout.addWidget(open_button)
            action_layout.addWidget(open_location_button)
            action_layout.addWidget(rename_button)
            action_layout.addWidget(change_save_location_button)
            main_layout.addLayout(action_layout)

            # Close button
            close_button = QPushButton("關閉程式")
            close_button.clicked.connect(self.close)
            main_layout.addWidget(close_button)

            # App label (右下角) - 文字左邊加上 icon，icon 大小與字體高度一致
            try:
                app_label_layout = QHBoxLayout()
                app_label_layout.setContentsMargins(0, 0, 0, 0)
                app_label_layout.setSpacing(4)
                app_label_layout.addStretch()

                app_label_text = QLabel("Smart Finder @ Xeson")

                # 取得字體高度作為 icon 大小
                fm = QFontMetrics(app_label_text.font())
                icon_size = fm.height()

                # 嘗試載入 icon
                app_icon_label = QLabel()
                app_icon_path = resource_path("program_icon.png")
                pix = QPixmap(app_icon_path)
                if not pix.isNull():
                    app_icon_label.setPixmap(
                        pix.scaled(QSize(icon_size, icon_size),
                                   Qt.KeepAspectRatio,
                                   Qt.SmoothTransformation)
                    )
                    app_label_layout.addWidget(app_icon_label, 0, Qt.AlignVCenter)

                app_label_layout.addWidget(app_label_text, 0, Qt.AlignVCenter)
                main_layout.addLayout(app_label_layout)
            except Exception:
                # 後備方案：仍顯示原來的文字
                app_label = QLabel("Smart Finder @ Xeson")
                app_label.setAlignment(Qt.AlignRight)
                main_layout.addWidget(app_label)
        except Exception:
            # 如果初始化失败，至少创建一个基本的窗口
            super().__init__()
            self.setWindowTitle("SmartFinder - Error")

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
            except Exception as e:
                pass  # 忽略保存最近目录的错误
            
            self.select_button.setEnabled(False)
            matching_files = search_files(target_dir, target_filename)
            
            self.result_listbox.clear()
            for file_path in matching_files:
                try:
                    # 添加异常处理，防止文件访问错误
                    if not os.path.exists(file_path):
                        continue
                    
                    file = os.path.basename(file_path)
                    file_stats = os.stat(file_path)
                    file_size = file_stats.st_size
                    file_modified_time = datetime.fromtimestamp(file_stats.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                    file_kind = get_file_kind(file_path)
                    
                    # 安全地计算相对路径
                    try:
                        relative_path = os.path.relpath(os.path.dirname(file_path), target_dir)
                    except (ValueError, OSError):
                        # 如果无法计算相对路径，使用绝对路径
                        relative_path = os.path.dirname(file_path)
                    
                    display_path = os.path.join("~", relative_path)
                    item = QTreeWidgetItem([file, file_kind, file_modified_time, format_file_size(file_size), display_path])
                    item.setData(0, Qt.UserRole, file_path)
                    self.result_listbox.addTopLevelItem(item)
                except (OSError, PermissionError, FileNotFoundError, ValueError) as e:
                    # 跳过无法访问或已删除的文件
                    continue
            
            result_count = self.result_listbox.invisibleRootItem().childCount()
            self.result_count_label.setText(f"搜尋結果: {result_count} 個檔案")
            self.select_button.setEnabled(True)
        else:
            try:
                QMessageBox.information(self, "缺少資訊", "請輸入目標地址和目標檔案名稱.")
            except Exception:
                pass  # 如果QMessageBox失败，忽略

    def update_selected_count(self):
        selected_count = len(self.result_listbox.selectedItems())
        self.selected_count_label.setText(f"選擇的檔案: {selected_count} 個")

    def open_selected_files(self):
        try:
            selected_items = self.result_listbox.selectedItems()
            if selected_items:
                file_paths = [item.data(0, Qt.UserRole) for item in selected_items]
                if len(file_paths) > 5:
                    try:
                        reply = QMessageBox.warning(self, "打開多個檔案", 
                                                    f"❗️警告❗️\n\n您選擇了 {len(file_paths)} 個檔案,\n確定要一次打開所有檔案嗎?\n\n卡死不負責。", 
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
                    QMessageBox.information(self, "未選取檔案", "請先選取要開啟的檔案。")
                except Exception:
                    pass
        except Exception:
            pass  # 捕获所有未预期的异常

    def open_file_location(self):
        try:
            selected_items = self.result_listbox.selectedItems()
            if selected_items:
                folder_paths = list(set([os.path.dirname(item.data(0, Qt.UserRole)) for item in selected_items]))
                if len(folder_paths) > 5:
                    try:
                        reply = QMessageBox.warning(self, "打開多個文件夾", 
                                                    f"❗️警告❗️\n\n您選擇的檔案位於\n {len(folder_paths)} 個不同的文件夾,\n確定要一次打開所有文件夾?\n\n卡死不負責。", 
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
                    QMessageBox.information(self, "未選取檔案", "請先選取要開啟路徑的檔案。")
                except Exception:
                    pass
        except Exception:
            pass  # 捕获所有未预期的异常

    def change_save_location(self):
        try:
            selected_items = self.result_listbox.selectedItems()
            if selected_items:
                new_location = QFileDialog.getExistingDirectory(self, "選擇新的存放位置")
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
                                        QMessageBox.information(self, "移動失敗", f"移動檔案 {file_name} 失敗:\n{str(e)}")
                                    except Exception:
                                        pass
                        except Exception:
                            pass

                    if duplicate_files:
                        try:
                            duplicate_files_text = "\n".join(duplicate_files)
                            QMessageBox.information(self, "移動部分失敗", f"以下檔案由\n於目標位置已有重複檔名而\n未被移動:\n\n{duplicate_files_text}")
                        except Exception:
                            pass

                    if moved_files:
                        try:
                            moved_files_text = "\n".join(moved_files)
                            QMessageBox.information(self, "移動完成", f"以下檔案已成功移動\n到新的存放位置:\n\n{moved_files_text}\n\n新的存放位置: \n{new_location}")
                            self.select_files()  # 刷新檔案列表
                        except Exception:
                            pass
                    else:
                        try:
                            QMessageBox.information(self, "移動失敗", "所有選取的檔案\n都未被移動,\n\n請檢查目標位置是否有重複檔名。")
                        except Exception:
                            pass
            else:
                try:
                    QMessageBox.information(self, "未選取檔案", "請先選取要更改存放位置的檔案。")
                except Exception:
                    pass
        except Exception:
            pass  # 捕获所有未预期的异常

    def rename_files(self):
        try:
            selected_items = self.result_listbox.selectedItems()
            if selected_items:
                dialog = RenameDialog(self)
                if dialog.exec_():
                    rename_option = dialog.get_rename_option()
                    new_name = dialog.get_new_name()
                    if new_name:
                        try:
                            confirmation_text = "以下是將要進行改名的檔案及其新檔名：\n\n"
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
                                reply = QMessageBox.question(self, "確認批量改名", confirmation_text, 
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
                                                QMessageBox.information(self, "無法改名", "無法改名\n請改一個沒有重複的檔案名稱")
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
                                        QMessageBox.information(self, "批量改名完成", "檔案批量改名已完成！")
                                    except Exception:
                                        pass
                                    
                                    try:
                                        self.select_files()  # Refresh the file list
                                    except Exception:
                                        pass
                                    
                                    try:
                                        reply = QMessageBox.question(self, "改變存放位置", "是否要改變已改名檔案的存放位置？", 
                                                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                                    except Exception:
                                        reply = QMessageBox.No
                                    
                                    if reply == QMessageBox.Yes:
                                        new_location = QFileDialog.getExistingDirectory(self, "選擇新的存放位置")
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
                                                                QMessageBox.information(self, "移動失敗", f"移動檔案 {new_file_name} 失敗:\n{str(e)}")
                                                            except Exception:
                                                                pass
                                                except Exception:
                                                    pass

                                            if duplicate_files:
                                                try:
                                                    duplicate_files_text = "\n".join(duplicate_files)
                                                    QMessageBox.information(self, "移動部分失敗", f"以下檔案由於目標位置已有重複檔名而未被移動:\n\n{duplicate_files_text}")
                                                except Exception:
                                                    pass

                                            if moved_files:
                                                try:
                                                    moved_files_text = "\n".join(moved_files)
                                                    QMessageBox.information(self, "移動完成", f"以下檔案已成功移動到新的存放位置:\n\n{moved_files_text}\n\n新的存放位置: \n{new_location}")
                                                except Exception:
                                                    pass
                                                try:
                                                    self.select_files()  # 刷新檔案列表
                                                except Exception:
                                                    pass
                                            else:
                                                try:
                                                    QMessageBox.information(self, "移動失敗", "所有選取的檔案都未被移動,請檢查目標位置是否有重複檔名。")
                                                except Exception:
                                                    pass
                        except Exception:
                            pass
                    else:
                        try:
                            QMessageBox.information(self, "缺少資訊", "請輸入新的檔案名稱。")
                        except Exception:
                            pass
            else:
                try:
                    QMessageBox.information(self, "未選取檔案", "請先選取要進行批量改名的檔案。")
                except Exception:
                    pass
        except Exception:
            pass  # 捕获所有未预期的异常

def main():
    try:
        app = QApplication(sys.argv)
        window = SmartFinderWindow()
        window.show()
        sys.exit(app.exec_())
    except Exception:
        # 如果应用初始化失败，尝试优雅地退出
        try:
            import traceback
            traceback.print_exc()
        except:
            pass
        sys.exit(1)

if __name__ == "__main__":
    main()