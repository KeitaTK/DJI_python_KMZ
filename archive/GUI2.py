import os
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES

def drop_file(event):
    # 複数ファイル対応：空白区切りで複数パスが渡る場合がある
    paths = event.data.strip().split()
    result = ""
    for path in paths:
        # {}で囲まれている場合があるので除去
        clean_path = path.strip("{}")
        file_name = os.path.basename(clean_path)
        result += f"ファイル名: {file_name}\nパス: {clean_path}\n\n"
    label.config(text=result)

root = TkinterDnD.Tk()
root.title("ドラッグ＆ドロップでファイル情報表示")
root.geometry("500x300")

label = tk.Label(root, text="ここにファイルをドラッグ＆ドロップ", bg="lightgray", width=60, height=10, anchor="nw", justify="left")
label.pack(padx=20, pady=40, fill="both", expand=True)

label.drop_target_register(DND_FILES)
label.dnd_bind('<<Drop>>', drop_file)

root.mainloop()
