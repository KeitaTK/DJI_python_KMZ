import os
import shutil
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

# 変換後のファイルを格納するフォルダ
OUTPUT_DIR = r"C:\Users\KT\Documents\local\M30_GPS\解凍フォルダ"

def convert_kmz_to_zip(kmz_path):
    if not kmz_path.lower().endswith('.kmz'):
        return False, "拡張子が.kmzではありません"
    zip_path = os.path.splitext(kmz_path)[0] + '.zip'
    try:
        # 元ファイルは残し、コピーで.zipを作成
        shutil.copy(kmz_path, zip_path)
        # 指定フォルダがなければ作成
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        dest_path = os.path.join(OUTPUT_DIR, os.path.basename(zip_path))
        shutil.move(zip_path, dest_path)
        return True, f"変換完了: {dest_path}"
    except Exception as e:
        return False, f"変換失敗: {e}"

def drop_file(event):
    paths = event.data.strip().split()
    if len(paths) != 1:
        messagebox.showwarning("エラー", "ファイルは1つだけドロップしてください。")
        label.config(text="ファイルは1つだけ受け付けます。")
        return
    path = paths[0].strip("{}")
    file_name = os.path.basename(path)
    label.config(text=f"ファイル名: {file_name}\nパス: {path}")
    # .kmz→.zip変換処理
    if path.lower().endswith('.kmz'):
        success, msg = convert_kmz_to_zip(path)
        messagebox.showinfo("変換結果", msg)
    else:
        messagebox.showinfo("情報", ".kmzファイルの場合のみ.zipに変換し、指定フォルダへ格納します。")

root = TkinterDnD.Tk()
root.title("ドラッグ＆ドロップでファイル情報表示（1ファイル限定）")
root.geometry("500x300")

label = tk.Label(root, text="ここに.kmzファイルをドラッグ＆ドロップ", bg="lightgray", width=60, height=10, anchor="nw", justify="left")
label.pack(padx=20, pady=40, fill="both", expand=True)

label.drop_target_register(DND_FILES)
label.dnd_bind('<<Drop>>', drop_file)

root.mainloop()
