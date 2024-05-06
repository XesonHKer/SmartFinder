import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
import mimetypes

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
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            if target_filename in file.lower():
                file_path = os.path.join(root, file)
                matching_files.append(file_path)
    
    return matching_files

def select_files():
    target_dir = target_dir_entry.get()
    target_filename = target_filename_entry.get()

    if target_dir and target_filename:
        select_button.config(state=tk.DISABLED)
        matching_files = search_files(target_dir, target_filename)
        
        if matching_files:
            result_listbox.delete(*result_listbox.get_children())  # 清空結果列表框
            for file_path in matching_files:
                file = os.path.basename(file_path)
                file_stats = os.stat(file_path)
                file_size = file_stats.st_size
                file_modified_time = datetime.fromtimestamp(file_stats.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                file_kind = get_file_kind(file_path)
                file_location = os.path.relpath(os.path.dirname(file_path), target_dir)
                result_listbox.insert("", tk.END, values=(file, file_kind, file_modified_time, format_file_size(file_size), file_location))
            
            result_count_label.config(text=f"搜尋結果: {len(matching_files)} 個檔案")
            sort_column("Date Modified", True)  # 默認按照最新修改時間降序排序
        else:
            result_listbox.delete(*result_listbox.get_children())  # 清空結果列表框
            result_count_label.config(text="搜尋結果: 0 個檔案")
            tk.messagebox.showinfo("沒有找到檔案", "在指定的目錄中\n沒有找到符合條件的檔案。")
        
        select_button.config(state=tk.NORMAL)
    else:
        tk.messagebox.showinfo("缺少資訊", "請輸入目標地址和目標檔案名稱。")

def open_selected_files():
    selected_files = result_listbox.selection()
    
    if selected_files:
        file_paths = [os.path.join(target_dir_entry.get(), result_listbox.item(item)['values'][4], result_listbox.item(item)['values'][0]) for item in selected_files]
        
        if len(file_paths) > 5:
            confirm_open = messagebox.askquestion("打開多個檔案", f"❗️警告❗️\n\n您選擇了 {len(file_paths)} 個檔案,\n確定要一次打開所有檔案嗎?\n\n卡死不負責。", icon='warning')
            if confirm_open == 'yes':
                for path in file_paths:
                    os.system(f'open "{path}"')
        else:
            for path in file_paths:
                os.system(f'open "{path}"')
    else:
        tk.messagebox.showinfo("未選取檔案", "請先選取要開啟的檔案。")
        
def open_file_location():
    selected_files = result_listbox.selection()
    
    if selected_files:
        folder_paths = list(set([os.path.dirname(os.path.join(target_dir_entry.get(), result_listbox.item(item)['values'][4], result_listbox.item(item)['values'][0])) for item in selected_files]))
        
        if len(folder_paths) > 5:
            confirm_open = messagebox.askquestion("打開多個文件夾", f"❗️警告❗️\n\n您選擇的檔案位於 {len(folder_paths)} 個不同的文件夾,\n確定要一次打開所有文件夾?\n\n卡死不負責。", icon='warning')
            if confirm_open == 'yes':
                for path in folder_paths:
                    os.system(f'open "{path}"')
        else:
            for path in folder_paths:
                os.system(f'open "{path}"')
    else:
        tk.messagebox.showinfo("未選取檔案", "請先選取要開啟路徑的檔案。")

def update_selected_count(event):
    selected_count = len(result_listbox.selection())
    selected_count_label.config(text=f"選擇的檔案: {selected_count} 個")

def close_program():
    window.destroy()

def sort_column(column, reverse=False):
    sorted_data = sorted(result_listbox.get_children(), key=lambda x: result_listbox.item(x)['values'][result_listbox['columns'].index(column)], reverse=reverse)
    
    for index, item in enumerate(sorted_data):
        result_listbox.move(item, "", index)

# 創建主視窗
window = tk.Tk()
window.title("多重選擇檔案")

# 目標地址標籤和輸入框
target_dir_label = tk.Label(window, text="目標地址:")
target_dir_label.pack()
target_dir_entry = tk.Entry(window, width=50)
target_dir_entry.pack()

# 目標檔案名稱標籤和輸入框
target_filename_label = tk.Label(window, text="目標檔案名稱:")
target_filename_label.pack()
target_filename_entry = tk.Entry(window, width=50)
target_filename_entry.pack()

# 確定及選取按鈕
select_button = tk.Button(window, text="確定及選取", command=select_files)
select_button.pack()

# 結果列表框
result_listbox = ttk.Treeview(window, columns=("File Name", "Kind", "Date Modified", "Size", "Location"), show="headings", selectmode="extended")
result_listbox.heading("File Name", text="File Name", command=lambda: sort_column("File Name"))
result_listbox.heading("Kind", text="Kind", command=lambda: sort_column("Kind"))
result_listbox.heading("Date Modified", text="Date Modified", command=lambda: sort_column("Date Modified"))
result_listbox.heading("Size", text="Size", command=lambda: sort_column("Size"))
result_listbox.heading("Location", text="Location", command=lambda: sort_column("Location"))
result_listbox.pack(fill=tk.BOTH, expand=True)

# 搜尋結果數量標籤
result_count_label = tk.Label(window, text="搜尋結果: 0 個檔案")
result_count_label.pack()

# 選擇的檔案數量標籤
selected_count_label = tk.Label(window, text="選擇的檔案: 0 個")
selected_count_label.pack()

# 打開選取檔案和打開路徑按鈕
open_button_frame = tk.Frame(window)
open_button_frame.pack()

open_button = tk.Button(open_button_frame, text="打開選取檔案", command=open_selected_files)
open_button.pack(side=tk.LEFT, padx=5)

open_location_button = tk.Button(open_button_frame, text="打開路徑", command=open_file_location)
open_location_button.pack(side=tk.LEFT, padx=5)

# 關閉程式按鈕
close_button = tk.Button(window, text="關閉程式", command=close_program)
close_button.pack()

# 綁定選擇事件
result_listbox.bind("<<TreeviewSelect>>", update_selected_count)

# 運行主循環
window.mainloop()