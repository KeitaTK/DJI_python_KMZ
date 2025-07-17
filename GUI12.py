#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
convert_height_gui_asl.py
---------------------------------
KMZ 内の WPML 高度をオフセット付きで変換し、
<wpml:heightMode> を relativeToStartPoint → EGM96 に置換する GUI ツール。

依存ライブラリ:
    pip install lxml tkinterdnd2
"""

import os
import shutil
import zipfile
import glob
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES
from lxml import etree

# WPML 名前空間
NS = {
    "kml":  "http://www.opengis.net/kml/2.2",
    "wpml": "http://www.dji.com/wpmz/1.0.6"
}

# ---------- KMZ 操作ユーティリティ ----------

def extract_kmz(kmz_path: str, work_dir: str) -> None:
    """KMZ → 作業ディレクトリへ展開"""
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(kmz_path, "r") as zf:
        zf.extractall(work_dir)

def repackage_kmz(work_dir: str, out_kmz: str) -> None:
    """作業ディレクトリ → KMZ 再パック"""
    tmp_zip = out_kmz + ".zip"
    with zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(work_dir):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, work_dir)
                zf.write(full, rel)
    if os.path.exists(out_kmz):
        os.remove(out_kmz)
    os.rename(tmp_zip, out_kmz)
    shutil.rmtree(work_dir)

# ---------- 高度・heightMode 変換 ----------

def convert_heights(tree: etree._ElementTree, offset: float) -> None:
    """高さ要素にオフセットを加算し、heightMode を EGM96 に置換"""
    # 1) 各 Placemark の高さ変換
    for pm in tree.findall(".//kml:Placemark", NS):
        for tag in ("height", "ellipsoidHeight"):
            elem = pm.find(f"wpml:{tag}", NS)
            if elem is not None:
                try:
                    elem.text = str(float(elem.text) + offset)
                except (TypeError, ValueError):
                    pass  # 数値でない場合はスキップ

    # 2) globalHeight も変換（存在する場合）
    gheight = tree.find(".//wpml:globalHeight", NS)
    if gheight is not None:
        try:
            gheight.text = str(float(gheight.text) + offset)
        except (TypeError, ValueError):
            pass

    # 3) heightMode をすべて EGM96 へ
    for hmode in tree.findall(".//wpml:heightMode", NS):
        hmode.text = "EGM96"

# ---------- ファイル単位で処理 ----------

def process_kmz(kmz_path: str, offset: float, log: tk.Text) -> None:
    """指定 KMZ を変換して保存"""
    work = "_kmz_work"
    try:
        log.insert(tk.END, f"Extracting {os.path.basename(kmz_path)}\n")
        extract_kmz(kmz_path, work)

        # template.kml は WPML の標準格納先
        kml_files = glob.glob(os.path.join(work, "**", "template.kml"), recursive=True)
        if not kml_files:
            raise FileNotFoundError("template.kml が見つかりません。KMZ の構造を確認してください。")

        for kml in kml_files:
            log.insert(tk.END, f"Converting {os.path.basename(kml)}\n")
            tree = etree.parse(kml)
            convert_heights(tree, offset)
            tree.write(kml, encoding="utf-8", pretty_print=True, xml_declaration=True)

        # 出力ファイル名に _asl とオフセット値を付加
        base, ext = os.path.splitext(kmz_path)
        sign = "+" if offset >= 0 else "-"
        out_kmz = f"{base}_asl{sign}{abs(offset)}{ext}"
        log.insert(tk.END, f"Repackaging → {os.path.basename(out_kmz)}\n")
        repackage_kmz(work, out_kmz)

        messagebox.showinfo("完了", f"変換完了:\n{out_kmz}")
        log.insert(tk.END, "Finished\n\n")
    except Exception as e:
        messagebox.showerror("エラー", str(e))
        log.insert(tk.END, f"Error: {e}\n\n")
        if os.path.exists(work):
            shutil.rmtree(work)

# ---------- GUI 構築 ----------

def on_drop(event):
    path = event.data.strip("{}")  # Windows の空白対策
    if not path.lower().endswith(".kmz"):
        messagebox.showwarning("警告", "KMZ ファイルのみ処理できます。")
        return
    try:
        offset_val = float(offset_entry.get())
    except ValueError:
        messagebox.showwarning("警告", "オフセットは数値で入力してください。")
        return
    threading.Thread(target=process_kmz,
                     args=(path, offset_val, log_text),
                     daemon=True).start()

root = TkinterDnD.Tk()
root.title("ATL → ASL 高度変換ツール")
root.geometry("640x420")

frame = ttk.Frame(root, padding=10)
frame.pack(fill="both", expand=True)

ttk.Label(frame, text="高度オフセット (m)").grid(row=0, column=0, sticky="w")
offset_entry = ttk.Entry(frame, width=10)
offset_entry.insert(0, "0.0")
offset_entry.grid(row=0, column=1, sticky="w", padx=5)

drop_label = tk.Label(frame, text="ここに .kmz をドラッグ＆ドロップ", bg="#d9d9d9",
                      width=60, height=5, relief=tk.RIDGE)
drop_label.grid(row=1, column=0, columnspan=2, pady=12, sticky="nsew")
drop_label.drop_target_register(DND_FILES)
drop_label.dnd_bind("<<Drop>>", on_drop)

log_frame = ttk.LabelFrame(frame, text="ログ")
log_frame.grid(row=2, column=0, columnspan=2, sticky="nsew")
frame.rowconfigure(2, weight=1)
frame.columnconfigure(1, weight=1)

log_text = scrolledtext.ScrolledText(log_frame, height=12, state="normal")
log_text.pack(fill="both", expand=True)

root.mainloop()
