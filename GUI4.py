import os
import shutil
import zipfile
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

# 展開先フォルダ
OUTPUT_DIR = r"C:\Users\KT\Documents\local\M30_GPS\解凍フォルダ"

def convert_and_extract_kmz(kmz_path):
    if not kmz_path.lower().endswith('.kmz'):
        return False, "拡張子が.kmzではありません"
    zip_path = os.path.splitext(kmz_path)[0] + '.zip'
    try:
        # 元ファイルは残し、コピーで.zipを作成
        shutil.copy(kmz_path, zip_path)
        # 指定フォルダがなければ作成
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        # zipを展開
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(OUTPUT_DIR)
        # zipファイルは不要なら削除（必要ならコメントアウト）
        os.remove(zip_path)
        return True, f"展開完了：{OUTPUT_DIR}"
    except Exception as e:
        return False, f"変換・展開失敗：{e}"

def drop_file(event):
    paths = event.data.strip().split()
    if len(paths) != 1:
        messagebox.showwarning("エラー", "ファイルは1つだけドロップしてください。", parent=root)
        label.config(text="ファイルは1つだけ受け付けます。")
        return
    path = paths[0].strip("{}")
    file_name = os.path.basename(path)
    label.config(text=f"ファイル名：{file_name}\nパス：{path}")
    # .kmz→.zip変換＆展開処理
    if path.lower().endswith('.kmz'):
        success, msg = convert_and_extract_kmz(path)
        messagebox.showinfo("結果", msg, parent=root)
    else:
        messagebox.showinfo("情報", ".kmzファイルの場合のみ対応します。", parent=root)

root = TkinterDnD.Tk()
root.title("ドラッグ＆ドロップでKMZ展開（1ファイル限定）")
root.geometry("500x300")

# 日本語フォント指定（Meiryo）
LABEL_FONT = ("Meiryo", 12)

label = tk.Label(
    root,
    text="ここに.kmzファイルをドラッグ＆ドロップ",
    bg="lightgray",
    width=60,
    height=10,
    anchor="nw",
    justify="left",
    font=LABEL_FONT
)
label.pack(padx=20, pady=40, fill="both", expand=True)

label.drop_target_register(DND_FILES)
label.dnd_bind('<<Drop>>', drop_file)

root.mainloop()
