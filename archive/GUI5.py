import os
import shutil
import zipfile
import glob
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

# 展開先フォルダ
OUTPUT_DIR = r"C:\Users\taki\Desktop\kaitouaaa"
LABEL_FONT = ("Meiryo", 12)

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
        # zipファイルは不要なので削除
        os.remove(zip_path)
        return True, f"展開完了：{OUTPUT_DIR}"
    except Exception as e:
        return False, f"変換・展開失敗：{e}"

def find_template_kml(base_dir):
    # 再帰的にwpmz\template.kmlを検索
    pattern = os.path.join(base_dir, "*", "wpmz", "template.kml")
    files = glob.glob(pattern, recursive=True)
    return files

def edit_kml_file(kml_path):
    # 例: <name>タグの内容を「変換済み」に書き換える
    try:
        tree = ET.parse(kml_path)
        root = tree.getroot()
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for name_elem in root.findall('.//kml:name', ns):
            name_elem.text = "変換済み"
        tree.write(kml_path, encoding="utf-8", xml_declaration=True)
        return True, f"編集・上書き成功: {kml_path}"
    except Exception as e:
        return False, f"編集失敗: {e}"

def process_kmz_and_edit_kml(kmz_path):
    # 1. kmz→zip変換＆展開
    success, msg = convert_and_extract_kmz(kmz_path)
    if not success:
        return False, msg
    # 2. template.kmlを自動検出して編集
    kml_files = find_template_kml(OUTPUT_DIR)
    if not kml_files:
        return False, "template.kmlが見つかりませんでした。"
    results = []
    for kml_path in kml_files:
        s, m = edit_kml_file(kml_path)
        results.append(m)
    return True, "\n".join(results)

def drop_file(event):
    paths = event.data.strip().split()
    if len(paths) != 1:
        messagebox.showwarning("エラー", "ファイルは1つだけドロップしてください。", parent=root)
        label.config(text="ファイルは1つだけ受け付けます。")
        return
    path = paths[0].strip("{}")
    file_name = os.path.basename(path)
    label.config(text=f"ファイル名：{file_name}\nパス：{path}")
    if path.lower().endswith('.kmz'):
        success, msg = process_kmz_and_edit_kml(path)
        messagebox.showinfo("結果", msg, parent=root)
    else:
        messagebox.showinfo("情報", ".kmzファイルの場合のみ対応します。", parent=root)

root = TkinterDnD.Tk()
root.title("ドラッグ＆ドロップでKMZ展開＆KML編集（1ファイル限定）")
root.geometry("500x300")

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
