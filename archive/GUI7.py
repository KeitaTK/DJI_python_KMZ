import os
import shutil
import zipfile
import glob
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

LABEL_FONT = ("Meiryo", 12)

def convert_and_extract_kmz(kmz_path):
    """
    KMZ → ZIP にコピー＆展開。展開先は同フォルダ内の Converted フォルダ。
    """
    if not kmz_path.lower().endswith(".kmz"):
        return False, "拡張子が .kmz ではありません"
    base_dir = os.path.dirname(kmz_path)
    output_dir = os.path.join(base_dir, "Converted")
    zip_path = os.path.splitext(kmz_path)[0] + ".zip"

    try:
        # ZIP にコピー＆展開
        shutil.copy(kmz_path, zip_path)
        # 既存 Converted フォルダは一旦クリア
        if os.path.isdir(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir)
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(output_dir)
        os.remove(zip_path)
        return True, output_dir
    except Exception as e:
        return False, f"展開失敗: {e}"

def find_template_kml(base_dir):
    """
    base_dir 以下を再帰的に検索し、wpmz/template.kml を探す。
    """
    pattern = os.path.join(base_dir, "**", "wpmz", "template.kml")
    return glob.glob(pattern, recursive=True)

def edit_kml_file(kml_path):
    """
    <name> タグを "Converted" に書き換えて上書き保存。
    """
    try:
        tree = ET.parse(kml_path)
        root = tree.getroot()
        ns = {"kml": "http://www.opengis.net/kml/2.2"}
        for elem in root.findall(".//kml:name", ns):
            elem.text = "Converted"
        tree.write(kml_path, encoding="utf-8", xml_declaration=True)
        return True, None
    except Exception as e:
        return False, f"編集失敗: {e}"

def repack_to_kmz_only(folder_path, original_kmz_path):
    """
    Converted フォルダ内の内容を zip 圧縮し、新 KMZ のみ Converted フォルダ内に出力。
    """
    base_dir = folder_path  # Converted フォルダ
    new_kmz = os.path.splitext(original_kmz_path)[0] + "_converted.kmz"
    temp_zip = os.path.splitext(new_kmz)[0] + ".zip"

    # ZIP 圧縮
    with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as z:
        for root_dir, _, files in os.walk(folder_path):
            for f in files:
                full = os.path.join(root_dir, f)
                rel = os.path.relpath(full, folder_path)
                z.write(full, rel)

    # ZIP → KMZ にリネーム
    if os.path.exists(new_kmz):
        os.remove(new_kmz)
    os.rename(temp_zip, os.path.join(base_dir, os.path.basename(new_kmz)))

    # 展開フォルダ内の他ファイルを削除して KMZ のみ残す
    for item in os.listdir(base_dir):
        path = os.path.join(base_dir, item)
        if not item.endswith(".kmz"):
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)

    return os.path.join(base_dir, os.path.basename(new_kmz))

def process_kmz_and_edit_kml(kmz_path):
    # 1. 展開
    ok, res = convert_and_extract_kmz(kmz_path)
    if not ok:
        return False, res
    work_dir = res

    # 2. 編集
    kmls = find_template_kml(work_dir)
    if not kmls:
        shutil.rmtree(work_dir)
        return False, "template.kml が見つかりませんでした。"
    for k in kmls:
        ok, err = edit_kml_file(k)
        if not ok:
            shutil.rmtree(work_dir)
            return False, err

    # 3. 再パック & Converted フォルダには KMZ のみ
    new_kmz = repack_to_kmz_only(work_dir, kmz_path)
    return True, f"新 KMZ 出力: {new_kmz}"

def drop_file(event):
    paths = event.data.strip().split()
    if len(paths) != 1:
        messagebox.showwarning("エラー", "ファイルは1つだけドロップしてください。", parent=root)
        return
    path = paths[0].strip("{}")
    label.config(text=f"ファイル：{os.path.basename(path)}")
    if path.lower().endswith(".kmz"):
        ok, msg = process_kmz_and_edit_kml(path)
        if ok:
            messagebox.showinfo("完了", msg, parent=root)
        else:
            messagebox.showerror("エラー", msg, parent=root)
    else:
        messagebox.showinfo("情報", ".kmz のみ対応", parent=root)

# GUI 初期化
root = TkinterDnD.Tk()
root.title("KMZ 展開・編集・再パック")
root.geometry("500x300")

label = tk.Label(
    root,
    text="ここに .kmz をドロップ",
    bg="lightgray",
    width=60,
    height=10,
    anchor="nw",
    justify="left",
    font=LABEL_FONT,
)
label.pack(padx=20, pady=40, fill="both", expand=True)

label.drop_target_register(DND_FILES)
label.dnd_bind("<<Drop>>", drop_file)

root.mainloop()
