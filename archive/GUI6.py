import os
import shutil
import zipfile
import glob
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import messagebox
# 正しいインポート方法
from tkinterdnd2 import TkinterDnD, DND_FILES

LABEL_FONT = ("Meiryo", 12)


def convert_and_extract_kmz(kmz_path):
    """
    KMZ → ZIP に一時コピーしたうえで展開する。
    展開先は KMZ ファイルと同じフォルダ内の「解凍フォルダ」。
    """
    if not kmz_path.lower().endswith(".kmz"):
        return False, "拡張子が .kmz ではありません"

    base_dir = os.path.dirname(kmz_path)
    output_dir = os.path.join(base_dir, "Converted")
    zip_path = os.path.splitext(kmz_path)[0] + ".zip"

    try:
        shutil.copy(kmz_path, zip_path)
        os.makedirs(output_dir, exist_ok=True)
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(output_dir)
        os.remove(zip_path)
        return True, f"展開完了: {output_dir}"
    except Exception as e:
        return False, f"変換・展開失敗: {e}"


def find_template_kml(base_dir):
    """
    base_dir 以下を再帰的に検索し、wpmz/template.kml を探す。
    """
    pattern = os.path.join(base_dir, "**", "wpmz", "template.kml")
    return glob.glob(pattern, recursive=True)


def edit_kml_file(kml_path):
    """
    例として <name> タグを「Converted」に書き換えて上書き保存。
    """
    try:
        tree = ET.parse(kml_path)
        root = tree.getroot()
        ns = {"kml": "http://www.opengis.net/kml/2.2"}
        for name_elem in root.findall(".//kml:name", ns):
            name_elem.text = "Converted"
        tree.write(kml_path, encoding="utf-8", xml_declaration=True)
        return True, f"編集・上書き成功: {kml_path}"
    except Exception as e:
        return False, f"編集失敗: {e}"


def process_kmz_and_edit_kml(kmz_path):
    """
    1. KMZ → ZIP 変換 & 展開
    2. template.kml を自動検出して編集
    """
    success, msg = convert_and_extract_kmz(kmz_path)
    if not success:
        return False, msg

    output_dir = os.path.join(os.path.dirname(kmz_path), "Converted")
    kml_files = find_template_kml(output_dir)
    if not kml_files:
        return False, "template.kml が見つかりませんでした。"

    results = []
    for kml_path in kml_files:
        s, m = edit_kml_file(kml_path)
        results.append(m)
    return True, "\n".join(results)


def drop_file(event):
    """
    ドラッグ＆ドロップされたファイルのパスを取得し処理。
    """
    paths = event.data.strip().split()
    if len(paths) != 1:
        messagebox.showwarning("エラー", "ファイルは1つだけドロップしてください。", parent=root)
        label.config(text="ファイルは1つだけ受け付けます。")
        return

    path = paths[0].strip("{}")
    file_name = os.path.basename(path)
    label.config(text=f"ファイル名：{file_name}\nパス：{path}")

    if path.lower().endswith(".kmz"):
        success, msg = process_kmz_and_edit_kml(path)
        if success:
            messagebox.showinfo("結果", msg, parent=root)
        else:
            messagebox.showerror("エラー", msg, parent=root)
    else:
        messagebox.showinfo("情報", ".kmzファイルの場合のみ対応します。", parent=root)


# GUI初期化
root = TkinterDnD.Tk()
root.title("ドラッグ＆ドロップで KMZ 展開＆KML 編集（1 ファイル限定）")
root.geometry("500x300")

label = tk.Label(
    root,
    text="ここに .kmz ファイルをドラッグ＆ドロップ",
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
