#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_height_gui_asl.py (ver. GUI37_complete)
前回のリファクタリングに基づき、省略されていた部分を補完した完全版。
- AppUI: Viewを担当。ウィジェットの作成と配置のみ。
- KmlConverterApp: Controllerを担当。UIイベントの処理とビジネスロジックの呼び出し。
- 処理関数群: Modelを担当。具体的なKMZ処理。
- main: アプリケーションの起動。
"""

import os
import shutil
import zipfile
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES
from lxml import etree
import math
from datetime import datetime
import pyperclip

# --- 定数 ------------------------------------------------------------------
NS = {
    "kml": "http://www.opengis.net/kml/2.2",
    "wpml": "http://www.dji.com/wpmz/1.0.6"
}

HEIGHT_OPTIONS = {
    "613.5 – 事務所前": 613.5,
    "962.02 – 烏帽子": 962.02,
    "その他 – 手動入力": "custom"
}

YAW_OPTIONS = {
    "1Q: 88.00°": 88.00,
    "2Q: 96.92°": 96.92,
    "4Q: 87.31°": 87.31,
    "手動入力": "custom"
}

SENSOR_MODES = ["Wide", "Zoom", "IR"]

REFERENCE_POINTS = {
    "本部": (136.5559522506280, 36.0729517605894, 612.2),
    "烏帽子": (136.560000000000, 36.075000000000, 962.02)
}

DEVIATION_THRESHOLD = {
    "lat": 0.00018,
    "lng": 0.00022,
    "alt": 20.0
}


# --- バックグラウンド処理関数群 (Model) ------------------------------

def extract_kmz(path, work_dir="_kmz_work"):
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(path, "r") as zf:
        zf.extractall(work_dir)
    return work_dir

def prepare_output_dirs(input_kmz, offset):
    base = os.path.splitext(os.path.basename(input_kmz))[0]
    sign = "+" if offset >= 0 else "-"
    root = os.path.dirname(input_kmz)
    out_root = os.path.join(root, f"{base}_asl{sign}{abs(offset)}")
    if os.path.exists(out_root):
        shutil.rmtree(out_root)
    os.makedirs(out_root)
    wpmz_dir = os.path.join(out_root, "wpmz")
    os.makedirs(wpmz_dir)
    return out_root, wpmz_dir

def repackage_to_kmz(out_root, input_kmz):
    base = os.path.splitext(os.path.basename(input_kmz))[0]
    out_kmz = os.path.join(os.path.dirname(out_root), f"{base}_Converted.kmz")
    tmp = out_kmz + ".zip"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(out_root):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, out_root)
                zf.write(full, rel)
    if os.path.exists(out_kmz):
        os.remove(out_kmz)
    os.rename(tmp, out_kmz)
    return out_kmz

def calculate_deviation(reference_coords, current_coords):
    ref_lng, ref_lat, ref_alt = reference_coords
    cur_lng, cur_lat, cur_alt = current_coords
    return cur_lng - ref_lng, cur_lat - ref_lat, cur_alt - ref_alt

def check_deviation_safety(deviation):
    dev_lng, dev_lat, dev_alt = deviation
    if (abs(dev_lng) > DEVIATION_THRESHOLD["lng"] or
        abs(dev_lat) > DEVIATION_THRESHOLD["lat"] or
        abs(dev_alt) > DEVIATION_THRESHOLD["alt"]):
        return False, (
            f"偏差が閾値を超えています:\n"
            f"経度偏差: {dev_lng:.8f}°\n"
            f"緯度偏差: {dev_lat:.8f}°\n"
            f"標高偏差: {dev_alt:.2f}m"
        )
    return True, None

def convert_kml(tree, params):
    """
    KMLファイル内のPlacemark要素を補正し、座標・高度・ヨー角を適用する。
    """
    offset = params["offset"]
    yaw_angle = params.get("yaw_angle")
    hover_time = params.get("hover_time", 0)
    do_photo = params.get("do_photo", False)
    do_video = params.get("do_video", False)
    video_suffix = params.get("video_suffix", "")
    do_gimbal = params.get("do_gimbal", True)
    sensor_modes = params.get("sensor_modes", [])
    dev_lng, dev_lat, dev_alt = params.get("coordinate_deviation", (0,0,0))

    # Placemarkごとに座標補正
    for pm in tree.findall(".//kml:Placemark", NS):
        coords_elem = pm.find(".//kml:coordinates", NS)
        if coords_elem is not None and coords_elem.text:
            parts = coords_elem.text.strip().split(',')
            if len(parts) >= 3:
                try:
                    lng = float(parts[0]) + dev_lng
                    lat = float(parts[1]) + dev_lat
                    alt = float(parts[2]) + offset
                    coords_elem.text = f"{lng},{lat},{alt}"
                except ValueError:
                    pass

        # ExtendedDataなどにヨー角や撮影設定を埋め込む
        ed = pm.find(".//kml:ExtendedData", NS)
        if ed is None:
            ed = etree.SubElement(pm, "{%s}ExtendedData" % NS["kml"])
        # 撮影モード
        mode = "photo" if do_photo else "video" if do_video else ""
        etree.SubElement(ed, "{%s}Data" % NS["kml"], name="mode").text = mode
        # ヨー角
        if yaw_angle is not None:
            etree.SubElement(ed, "{%s}Data" % NS["kml"], name="yaw").text = f"{yaw_angle}"
        # ホバリング時間
        if hover_time > 0:
            etree.SubElement(ed, "{%s}Data" % NS["kml"], name="hover_time").text = f"{hover_time}"
        # センサー
        etree.SubElement(ed, "{%s}Data" % NS["kml"], name="sensors").text = ",".join(sensor_modes)
        # ジンバル制御フラグ
        etree.SubElement(ed, "{%s}Data" % NS["kml"], name="gimbal").text = str(do_gimbal)

def process_kmz(path, log, params):
    try:
        log.insert(tk.END, f"処理開始: {os.path.basename(path)}...\n")
        work_dir = extract_kmz(path)
        # KMLファイルパス探索
        kml_path = os.path.join(work_dir, "wpmz", "template.kml")
        if not os.path.exists(kml_path):
            kml_path = os.path.join(work_dir, "template.kml")
            if not os.path.exists(kml_path):
                raise FileNotFoundError("template.kml が見つかりませんでした。")

        out_root, outdir = prepare_output_dirs(path, params["offset"])
        log.insert(tk.END, f"変換中: {os.path.basename(kml_path)}...\n")

        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(kml_path, parser)

        convert_kml(tree, params)

        out_kml = os.path.join(outdir, os.path.basename(kml_path))
        tree.write(out_kml, encoding="utf-8", pretty_print=True, xml_declaration=True)

        # リソースフォルダのコピー
        res_src = os.path.join(os.path.dirname(kml_path), "res")
        if os.path.isdir(res_src):
            shutil.copytree(res_src, os.path.join(outdir, "res"))

        out_kmz = repackage_to_kmz(out_root, path)
        log.insert(tk.END, f"変換完了: {out_kmz}\n\n")
        messagebox.showinfo("完了", f"変換が正常に完了しました:\n{out_kmz}")

    except Exception as e:
        log.insert(tk.END, f"エラー: {e}\n\n")
        messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
    finally:
        if os.path.exists("_kmz_work"):
            shutil.rmtree("_kmz_work")


# --- UIクラス (View) --------------------------------------------------

class AppUI(ttk.Frame):
    def __init__(self, master, controller):
        super().__init__(master, padding=10)
        self.controller = controller
        self._create_vars()
        self._create_widgets()
        self._grid_widgets()

    def _create_vars(self):
        self.height_choice_var = tk.StringVar()
        self.height_entry_var = tk.StringVar()
        self.speed_var = tk.IntVar(value=15)
        self.photo_var = tk.BooleanVar(value=False)
        self.video_var = tk.BooleanVar(value=False)
        self.video_suffix_var = tk.StringVar(value="video_01")
        self.sensor_vars = {m: tk.BooleanVar(value=False) for m in SENSOR_MODES}
        self.gimbal_var = tk.BooleanVar(value=True)
        self.yaw_fix_var = tk.BooleanVar(value=False)
        self.yaw_choice_var = tk.StringVar()
        self.yaw_entry_var = tk.StringVar()
        self.hover_var = tk.BooleanVar(value=False)
        self.hover_time_var = tk.StringVar(value="2")
        self.deviation_var = tk.BooleanVar(value=False)
        self.ref_point_var = tk.StringVar(value="本部")
        self.today_lng_var = tk.StringVar(value="136.555")
        self.today_lat_var = tk.StringVar(value="36.072")
        self.today_alt_var = tk.StringVar(value="0")

    def _create_widgets(self):
        # 高度選択
        self.height_label = ttk.Label(self, text="基準高度:")
        self.height_combo = ttk.Combobox(
            self, textvariable=self.height_choice_var,
            values=list(HEIGHT_OPTIONS.keys()), state="readonly", width=20
        )
        self.height_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.height_entry = ttk.Entry(self, textvariable=self.height_entry_var, width=10, state="disabled")

        # 速度
        self.speed_label = ttk.Label(self, text="速度 (1–15 m/s):")
        self.speed_spinbox = ttk.Spinbox(self, from_=1, to=15, textvariable=self.speed_var, width=5)

        # 撮影設定
        self.photo_check = ttk.Checkbutton(self, text="写真撮影", variable=self.photo_var, command=self.controller.update_ui_states)
        self.video_check = ttk.Checkbutton(self, text="動画撮影", variable=self.video_var, command=self.controller.update_ui_states)
        self.video_suffix_label = ttk.Label(self, text="動画ファイル名:")
        self.video_suffix_entry = ttk.Entry(self, textvariable=self.video_suffix_var, width=20)

        # センサー
        self.sensor_label = ttk.Label(self, text="センサー選択:")
        self.sensor_checks = {
            m: ttk.Checkbutton(self, text=m, variable=v)
            for m, v in self.sensor_vars.items()
        }

        # ジンバル
        self.gimbal_check = ttk.Checkbutton(self, text="ジンバル制御", variable=self.gimbal_var)

        # ヨー固定
        self.yaw_fix_check = ttk.Checkbutton(self, text="ヨー固定", variable=self.yaw_fix_var, command=self.controller.update_ui_states)
        self.yaw_combo = ttk.Combobox(
            self, textvariable=self.yaw_choice_var,
            values=list(YAW_OPTIONS.keys()), state="readonly", width=15
        )
        self.yaw_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.yaw_entry = ttk.Entry(self, textvariable=self.yaw_entry_var, width=8, state="disabled")

        # ホバリング
        self.hover_check = ttk.Checkbutton(self, text="ホバリング", variable=self.hover_var, command=self.controller.update_ui_states)
        self.hover_time_label = ttk.Label(self, text="ホバリング時間 (秒):")
        self.hover_time_entry = ttk.Entry(self, textvariable=self.hover_time_var, width=8)

        # 偏差補正
        self.deviation_check = ttk.Checkbutton(self, text="偏差補正", variable=self.deviation_var, command=self.controller.update_ui_states)
        self.ref_point_label = ttk.Label(self, text="基準位置:")
        self.ref_point_combo = ttk.Combobox(
            self, textvariable=self.ref_point_var,
            values=list(REFERENCE_POINTS.keys()), state="readonly", width=10
        )
        self.today_coords_label = ttk.Label(self, text="本日の値 (経度,緯度,標高):")
        self.today_lng_entry = ttk.Entry(self, textvariable=self.today_lng_var, width=12)
        self.today_lat_entry = ttk.Entry(self, textvariable=self.today_lat_var, width=10)
        self.today_alt_entry = ttk.Entry(self, textvariable=self.today_alt_var, width=8)
        self.copy_button = ttk.Button(self, text="コピー", command=self.controller.copy_reference_data, width=8)

    def _grid_widgets(self):
        # 配置
        self.height_label.grid(row=0, column=0, sticky="w")
        self.height_combo.grid(row=0, column=1, padx=5, columnspan=2, sticky="w")
        self.height_entry.grid(row=0, column=3, padx=5)

        self.speed_label.grid(row=1, column=0, sticky="w", pady=5)
        self.speed_spinbox.grid(row=1, column=1, columnspan=2, sticky="w")

        self.photo_check.grid(row=2, column=0, sticky="w")
        self.video_check.grid(row=2, column=1, sticky="w")

        self.sensor_label.grid(row=3, column=0, sticky="w")
        for i, m in enumerate(SENSOR_MODES):
            self.sensor_checks[m].grid(row=3, column=1+i, sticky="w")

        self.gimbal_check.grid(row=4, column=0, sticky="w", pady=5)
        self.yaw_fix_check.grid(row=5, column=0, sticky="w")
        self.hover_check.grid(row=6, column=0, sticky="w", pady=5)
        self.deviation_check.grid(row=7, column=0, sticky="w", pady=5)

        self.height_combo.set(next(iter(HEIGHT_OPTIONS.keys())))


# --- アプリケーション制御クラス (Controller) -------------------------

class KmlConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ATL→ASL 変換＋撮影制御ツール (ver. GUI37_complete)")
        self.root.geometry("800x750")

        self.ui = AppUI(self.root, self)
        self.ui.pack(fill="x", pady=(0,10))

        self._create_dnd_and_log_area()
        self.update_ui_states()

    def _create_dnd_and_log_area(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill="both", expand=True)

        self.drop_label = tk.Label(frm, text=".kmzをここにドロップ", bg="lightgray", width=70, height=5, relief=tk.RIDGE)
        self.drop_label.pack(pady=12, fill="x")
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self.on_drop)

        log_frame = ttk.LabelFrame(frm, text="ログ")
        log_frame.pack(fill="both", expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=16)
        self.log_text.pack(fill="both", expand=True)

    def on_drop(self, event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", ".kmzファイルのみ対応しています。")
            return
        try:
            params = self._get_params()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            return
        threading.Thread(target=process_kmz, args=(path, self.log_text, params), daemon=True).start()

    def update_ui_states(self, event=None):
        ui = self.ui
        # 高度手動入力
        if HEIGHT_OPTIONS.get(ui.height_choice_var.get()) == "custom":
            ui.height_entry.config(state="normal")
        else:
            ui.height_entry.config(state="disabled")

        # 撮影モードの排他制御
        if ui.photo_var.get():
            ui.video_var.set(False)
        if ui.video_var.get():
            ui.photo_var.set(False)

        # ジンバル制御
        ui.gimbal_check.config(state="disabled" if (ui.photo_var.get() or ui.video_var.get()) else "normal")
        if ui.photo_var.get() or ui.video_var.get():
            ui.gimbal_var.set(True)

        # 動画サフィックス表示
        if ui.video_var.get():
            ui.video_suffix_label.grid(row=2, column=2, sticky="e", padx=(10,2))
            ui.video_suffix_entry.grid(row=2, column=3, sticky="w")
        else:
            ui.video_suffix_label.grid_forget()
            ui.video_suffix_entry.grid_forget()

        # ヨー固定
        if ui.yaw_fix_var.get():
            ui.yaw_combo.grid(row=5, column=1, padx=5, columnspan=2, sticky="w")
            if not ui.yaw_choice_var.get():
                ui.yaw_choice_var.set(next(iter(YAW_OPTIONS.keys())))
            if ui.yaw_choice_var.get() == "手動入力":
                ui.yaw_entry.config(state="normal")
                ui.yaw_entry.grid(row=5, column=3)
            else:
                ui.yaw_entry.config(state="disabled")
                ui.yaw_entry.grid_forget()
        else:
            ui.yaw_combo.grid_forget()
            ui.yaw_entry.grid_forget()

        # ホバリング
        if ui.hover_var.get():
            ui.hover_time_label.grid(row=6, column=1, sticky="e", padx=(10,2))
            ui.hover_time_entry.grid(row=6, column=2, sticky="w")
        else:
            ui.hover_time_label.grid_forget()
            ui.hover_time_entry.grid_forget()

        # 偏差補正
        if ui.deviation_var.get():
            ui.ref_point_label.grid(row=7, column=1, sticky="e", padx=(10,2))
            ui.ref_point_combo.grid(row=7, column=2, sticky="w")
            ui.today_coords_label.grid(row=8, column=0, sticky="w", pady=5)
            ui.today_lng_entry.grid(row=8, column=1, padx=2, sticky="w")
            ui.today_lat_entry.grid(row=8, column=2, padx=2, sticky="w")
            ui.today_alt_entry.grid(row=8, column=3, padx=2, sticky="w")
            ui.copy_button.grid(row=8, column=4, padx=5, sticky="w")
        else:
            ui.ref_point_label.grid_forget()
            ui.ref_point_combo.grid_forget()
            ui.today_coords_label.grid_forget()
            ui.today_lng_entry.grid_forget()
            ui.today_lat_entry.grid_forget()
            ui.today_alt_entry.grid_forget()
            ui.copy_button.grid_forget()

    def copy_reference_data(self):
        try:
            ref = self.ui.ref_point_var.get()
            coords = REFERENCE_POINTS[ref]
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            txt = f"{now},{coords[0]},{coords[1]},{coords[2]}"
            pyperclip.copy(txt)
            messagebox.showinfo("コピー完了", f"クリップボードにコピーしました:\n{txt}")
        except Exception as e:
            messagebox.showerror("エラー", f"コピーに失敗しました: {e}")

    def _get_params(self):
        ui = self.ui
        p = {}

        # オフセット
        v = HEIGHT_OPTIONS.get(ui.height_choice_var.get())
        if v == "custom":
            try:
                p["offset"] = float(ui.height_entry_var.get())
            except ValueError:
                raise ValueError("基準高度（手動入力）には数値を入力してください。")
        else:
            p["offset"] = float(v)

        # ヨー角
        p["yaw_angle"] = None
        if ui.yaw_fix_var.get():
            y = YAW_OPTIONS.get(ui.yaw_choice_var.get())
            if y == "custom":
                try:
                    p["yaw_angle"] = float(ui.yaw_entry_var.get())
                except ValueError:
                    raise ValueError("ヨー角（手動入力）には数値を入力してください。")
            else:
                p["yaw_angle"] = float(y)

        # ホバリング
        p["hover_time"] = 0
        if ui.hover_var.get():
            try:
                p["hover_time"] = max(0, float(ui.hover_time_var.get()))
            except ValueError:
                p["hover_time"] = 2.0

        # 偏差補正
        p["coordinate_deviation"] = None
        if ui.deviation_var.get():
            try:
                ref = REFERENCE_POINTS[ui.ref_point_var.get()]
                cur = (
                    float(ui.today_lng_var.get()),
                    float(ui.today_lat_var.get()),
                    float(ui.today_alt_var.get())
                )
                dev = calculate_deviation(ref, cur)
                ok, msg = check_deviation_safety(dev)
                if not ok:
                    raise ValueError(msg)
                p["coordinate_deviation"] = dev
            except ValueError as e:
                raise ValueError(f"偏差補正の値が不正です: {e}")

        # その他
        p["do_photo"] = ui.photo_var.get()
        p["do_video"] = ui.video_var.get()
        p["video_suffix"] = ui.video_suffix_var.get()
        p["do_gimbal"] = ui.gimbal_var.get()
        p["yaw_fix"] = ui.yaw_fix_var.get()
        p["speed"] = max(1, min(15, ui.speed_var.get()))
        p["sensor_modes"] = [m for m,var in ui.sensor_vars.items() if var.get()]

        return p


# --- メイン実行部分 ------------------------------------------------

def main():
    """アプリケーションを初期化して実行する"""
    root = TkinterDnD.Tk()
    app = KmlConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
