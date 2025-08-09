#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
wayline_to_google_earth_kmz.py

DJI Wayline-KML（.kmz 内 template.kml）を解析し、
• 各ウェイポイントを高度付きピン
• ウェイポイント間を高度付きラインで結ぶ
Google Earth Pro 用 KMZ を生成する GUIツール
"""

import os
import zipfile
import shutil
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES
from lxml import etree

# pip install simplekml
import simplekml

# KML 名前空間
NS = {
    "kml":   "http://www.opengis.net/kml/2.2",
    "wpml":  "http://www.dji.com/wpmz/1.0.6"
}

def extract_kmz(kmz_path, work_dir="_kmz_work"):
    if os.path.isdir(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(kmz_path, "r") as zf:
        zf.extractall(work_dir)
    # テンプレート KML を検索
    for root, _, files in os.walk(work_dir):
        for f in files:
            if f.lower() == "template.kml":
                return os.path.join(root, f)
    raise FileNotFoundError("template.kml が見つかりません。")

def parse_waypoints(kml_file):
    tree = etree.parse(kml_file)
    pts = []
    for pm in tree.findall(".//kml:Placemark", NS):
        idx = pm.findtext("wpml:index", namespaces=NS)
        coord = pm.findtext("kml:Point/kml:coordinates", namespaces=NS)
        alt = pm.findtext("wpml:height", namespaces=NS) or pm.findtext("wpml:ellipsoidHeight", namespaces=NS) or "0"
        if coord and idx is not None:
            lon, lat = map(float, coord.split(",")[:2])
            pts.append((int(idx), lon, lat, float(alt)))
    # インデックス順にソート
    pts.sort(key=lambda x: x[0])
    return pts

def build_kmz(pts, out_kmz):
    kml = simplekml.Kml()
    # ウェイポイント
    for idx, lon, lat, alt in pts:
        p = kml.newpoint(name=f"WP{idx}", coords=[(lon, lat, alt)])
        p.altitudemode = simplekml.AltitudeMode.absolute
        p.style.iconstyle.icon.href = "http://maps.google.com/mapfiles/kml/paddle/red-circle.png"
    # ルートライン
    coords = [(lon, lat, alt) for _, lon, lat, alt in pts]
    ls = kml.newlinestring(name="Flight Path", coords=coords)
    ls.altitudemode = simplekml.AltitudeMode.absolute
    ls.extrude = 1
    ls.style.linestyle.width = 3
    ls.style.linestyle.color = simplekml.Color.blue
    # KMZ 出力
    kml.savekmz(out_kmz)

def process_file(path, log):
    try:
        log.insert(tk.END, f"処理開始: {path}\n")
        tpl = extract_kmz(path)
        log.insert(tk.END, "template.kml を解析中…\n")
        pts = parse_waypoints(tpl)
        if not pts:
            messagebox.showwarning("警告", "ウェイポイントが見つかりません。")
            return
        out_kmz = os.path.splitext(path)[0] + "_GE.kmz"
        build_kmz(pts, out_kmz)
        log.insert(tk.END, f"出力完了: {out_kmz}\n")
        messagebox.showinfo("完了", f"KMZ を生成しました:\n{out_kmz}")
    except Exception as e:
        log.insert(tk.END, f"エラー: {e}\n")
        messagebox.showerror("エラー", str(e))

class App(ttk.Frame):
    def __init__(self, root):
        super().__init__(root, padding=10)
        root.title("Wayline→GE KMZ 変換ツール")
        root.geometry("600x400")
        self.pack(fill="both", expand=True)
        lbl = tk.Label(self, text=".kmz ファイルをここにドロップ", bg="lightgray", height=5)
        lbl.pack(fill="x", pady=5)
        lbl.drop_target_register(DND_FILES)
        lbl.dnd_bind("<<Drop>>", self.on_drop)
        self.log = scrolledtext.ScrolledText(self, height=10)
        self.log.pack(fill="both", expand=True)
    def on_drop(self, event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", ".kmz ファイルのみ対応")
            return
        threading.Thread(target=process_file, args=(path, self.log), daemon=True).start()

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    App(root)
    root.mainloop()
