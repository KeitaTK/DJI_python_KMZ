#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
convert_height_gui_asl.py (ver. GUI51 – 元角度維持統合版)

変更点
• 「ジンバルピッチ」「ヨー固定」の選択肢に「元の角度維持」を追加  
• 元角度維持時は変換前KMLから抽出した各ウェイポイントの角度をそのまま適用  
• 全体の「元ジンバル角度使用」チェックを廃止し、個別選択に統合  
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
    "元の角度維持": "original",
    "手動入力": "custom"
}

GIMBAL_PITCH_OPTIONS = {
    "真下: -90°": -90.0,
    "真ん前: 0°": 0.0,
    "元の角度維持": "original",
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

# --- 元ジンバル角度抽出 -----------------------------------------------------

def extract_original_gimbal_angles(tree):
    original = {}
    for pm in tree.findall(".//kml:Placemark", NS):
        idx = pm.find("wpml:index", NS)
        if idx is None: continue
        i = int(idx.text)
        for action in pm.findall("wpml:action", NS):
            func = action.find("wpml:actionActuatorFunc", NS)
            if func is not None and func.text == "orientedShoot":
                p = action.find("wpml:actionActuatorFuncParam", NS)
                pitch = float(p.find("wpml:gimbalPitchRotateAngle", NS).text)
                yaw   = float(p.find("wpml:gimbalYawRotateAngle", NS).text)
                head_elem = p.find("wpml:aircraftHeading", NS)
                head  = float(head_elem.text) if head_elem is not None else None
                original[i] = {"pitch": pitch, "yaw": yaw, "heading": head}
                break
    return original

# --- KMZ ユーティリティ -----------------------------------------------------

def extract_kmz(path, work_dir="_kmz_work"):
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(path, "r") as z:
        z.extractall(work_dir)
    return work_dir

def prepare_output_dirs(input_kmz, offset):
    base = os.path.splitext(os.path.basename(input_kmz))[0]
    sign = "+" if offset >= 0 else "-"
    root = os.path.dirname(input_kmz)
    out = os.path.join(root, f"{base}_asl{sign}{abs(offset)}")
    if os.path.exists(out):
        shutil.rmtree(out)
    os.makedirs(out)
    os.makedirs(os.path.join(out, "wpmz"))
    return out, os.path.join(out, "wpmz")

def repackage_to_kmz(out_root, input_kmz):
    base = os.path.splitext(os.path.basename(input_kmz))[0]
    out_kmz = os.path.join(os.path.dirname(out_root), f"{base}_Converted.kmz")
    tmp = out_kmz + ".zip"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for r, _, fs in os.walk(out_root):
            for f in fs:
                if f.lower().endswith(".wpml"):
                    continue
                full = os.path.join(r, f)
                rel = os.path.relpath(full, out_root)
                z.write(full, rel)
    if os.path.exists(out_kmz):
        os.remove(out_kmz)
    os.rename(tmp, out_kmz)
    return out_kmz

# --- 偏差計算 ---------------------------------------------------------------

def calculate_deviation(ref, cur):
    return cur[0] - ref[0], cur[1] - ref[1], cur[2] - ref[2]

def check_deviation_safety(dev):
    if abs(dev[0]) > DEVIATION_THRESHOLD["lng"] or \
       abs(dev[1]) > DEVIATION_THRESHOLD["lat"] or \
       abs(dev[2]) > DEVIATION_THRESHOLD["alt"]:
        return False, f"偏差が20mを超えています: 経度{dev[0]:.8f}, 緯度{dev[1]:.8f}, 標高{dev[2]:.2f}"
    return True, None

# --- KML 変換 ---------------------------------------------------------------

def convert_kml(tree, params, original_angles):
    offset = params["offset"]
    do_photo = params["do_photo"]
    do_video = params["do_video"]
    video_suffix = params["video_suffix"]
    yaw_mode = params["yaw_mode"]
    yaw_angle = params["yaw_angle"]
    gp_mode = params["gimbal_pitch_mode"]
    gp_angle = params["gimbal_pitch_angle"]
    speed = params["speed"]
    sensor_modes = params["sensor_modes"]
    hover_time = params["hover_time"]
    dev = params["coordinate_deviation"]

    # 座標偏差補正
    if dev:
        for pm in tree.findall(".//kml:Placemark", NS):
            c = pm.find(".//kml:coordinates", NS)
            if c is not None and c.text:
                lon, lat, *_ = c.text.strip().split(",")
                c.text = f"{float(lon)+dev[0]},{float(lat)+dev[1]}"

    # 高度補正＋EGM96
    for pm in tree.findall(".//kml:Placemark", NS):
        for tag in ("height", "ellipsoidHeight"):
            el = pm.find(f"wpml:{tag}", NS)
            if el is not None and el.text:
                h = float(el.text) + offset
                if dev:
                    h += dev[2]
                el.text = str(h)
    gh = tree.find(".//wpml:globalHeight", NS)
    if gh is not None and gh.text:
        h = float(gh.text) + offset
        if dev:
            h += dev[2]
        gh.text = str(h)
    for hm in tree.findall(".//wpml:heightMode", NS):
        hm.text = "EGM96"

    # 速度設定
    for tag, path in [("globalTransitionalSpeed", ".//wpml:missionConfig"),
                      ("autoFlightSpeed", ".//kml:Folder")]:
        el = tree.find(f"{path}/wpml:{tag}", NS)
        if el is not None:
            el.text = str(speed)
        else:
            parent = tree.find(path, NS)
            if parent is not None:
                etree.SubElement(parent, f"{{{NS['wpml']}}}{tag}").text = str(speed)
    for pm in tree.findall(".//kml:Placemark", NS):
        for t, v in [("waypointSpeed", speed), ("useGlobalSpeed", 0)]:
            el = pm.find(f"wpml:{t}", NS)
            if el is not None:
                el.text = str(v)
            else:
                etree.SubElement(pm, f"{{{NS['wpml']}}}{t}").text = str(v)

    # アクション全削除
    for pm in tree.findall(".//kml:Placemark", NS):
        for ag in list(pm.findall("wpml:actionGroup", NS)):
            pm.remove(ag)
    pp = tree.find(".//wpml:payloadParam", NS)
    if pp is not None:
        img = pp.find("wpml:imageFormat", NS)
        if img is not None:
            pp.remove(img)

    # センサー選択
    if sensor_modes:
        if pp is None:
            fld = tree.find(".//kml:Folder", NS)
            pp = etree.SubElement(fld, f"{{{NS['wpml']}}}payloadParam")
            etree.SubElement(pp, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
        fmt = ",".join(m.lower() for m in sensor_modes)
        etree.SubElement(pp, f"{{{NS['wpml']}}}imageFormat").text = fmt

    # 各PlacemarkにactionGroup追加
    placemarks = sorted(
        [p for p in tree.findall(".//kml:Placemark", NS)
         if p.find("wpml:index", NS) is not None],
        key=lambda p: int(p.find("wpml:index", NS).text)
    )
    for pm in placemarks:
        idx = int(pm.find("wpml:index", NS).text)
        ag = etree.SubElement(pm, f"{{{NS['wpml']}}}actionGroup")
        for tag in ("actionGroupId", "actionGroupStartIndex", "actionGroupEndIndex"):
            etree.SubElement(ag, f"{{{NS['wpml']}}}{tag}").text = str(idx)
        etree.SubElement(ag, f"{{{NS['wpml']}}}actionGroupMode").text = "sequence"
        trg = etree.SubElement(ag, f"{{{NS['wpml']}}}actionTrigger")
        etree.SubElement(trg, f"{{{NS['wpml']}}}actionTriggerType").text = "reachPoint"

        # ヨー制御
        if yaw_mode == "original" and idx in original_angles:
            hd = original_angles[idx]["heading"]
            if hd is not None:
                a = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(a, f"{{{NS['wpml']}}}actionId").text = "0"
                etree.SubElement(a, f"{{{NS['wpml']}}}actionActuatorFunc").text = "rotateYaw"
                p = etree.SubElement(a, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(p, f"{{{NS['wpml']}}}aircraftHeading").text = str(int(hd))
                etree.SubElement(p, f"{{{NS['wpml']}}}aircraftPathMode").text = "counterClockwise"
        elif yaw_mode == "fixed":
            a = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(a, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(a, f"{{{NS['wpml']}}}actionActuatorFunc").text = "rotateYaw"
            p = etree.SubElement(a, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(p, f"{{{NS['wpml']}}}aircraftHeading").text = str(int(yaw_angle))
            etree.SubElement(p, f"{{{NS['wpml']}}}aircraftPathMode").text = "counterClockwise"

        # ジンバルピッチ制御
        if gp_mode == "original" and idx in original_angles:
            od = original_angles[idx]
            a = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(a, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(a, f"{{{NS['wpml']}}}actionActuatorFunc").text = "gimbalRotate"
            p = etree.SubElement(a, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalRotateMode").text = "absoluteAngle"
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalPitchRotateEnable").text = "1"
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalPitchRotateAngle").text = str(int(od["pitch"]))
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalRollRotateEnable").text = "0"
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalYawRotateEnable").text = "0"
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalRotateTimeEnable").text = "0"
            etree.SubElement(p, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
        elif gp_mode == "fixed":
            a = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(a, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(a, f"{{{NS['wpml']}}}actionActuatorFunc").text = "gimbalRotate"
            p = etree.SubElement(a, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalRotateMode").text = "absoluteAngle"
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalPitchRotateEnable").text = "1"
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalPitchRotateAngle").text = str(int(gp_angle))
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalRollRotateEnable").text = "0"
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalYawRotateEnable").text = "0"
            etree.SubElement(p, f"{{{NS['wpml']}}}gimbalRotateTimeEnable").text = "0"
            etree.SubElement(p, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"

        # 写真撮影アクション
        if do_photo:
            a = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
            etree.SubElement(a, f"{{{NS['wpml']}}}actionId").text = "0"
            etree.SubElement(a, f"{{{NS['wpml']}}}actionActuatorFunc").text = "takePhoto"
            p = etree.SubElement(a, f"{{{NS['wpml']}}}actionActuatorFuncParam")
            etree.SubElement(p, f"{{{NS['wpml']}}}fileSuffix").text = f"ウェイポイント{idx}"
            etree.SubElement(p, f"{{{NS['wpml']}}}payloadPositionIndex").text = "0"
            if hover_time > 0:
                ha = etree.SubElement(ag, f"{{{NS['wpml']}}}action")
                etree.SubElement(ha, f"{{{NS['wpml']}}}actionId").text = "0"
                etree.SubElement(ha, f"{{{NS['wpml']}}}actionActuatorFunc").text = "hover"
                hp = etree.SubElement(ha, f"{{{NS['wpml']}}}actionActuatorFuncParam")
                etree.SubElement(hp, f"{{{NS['wpml']}}}hoverTime").text = str(int(hover_time))

    # 動画モードの処理は同様に YAW/Gimbal を適用して startRecord/stopRecord を追加

# --- 処理フロー -------------------------------------------------------------

def process_kmz(path, params, log):
    try:
        log.insert(tk.END, f"Extracting {os.path.basename(path)}...\n")
        wd = extract_kmz(path)
        kmls = glob.glob(os.path.join(wd, "**", "template.kml"), recursive=True)
        if not kmls:
            raise FileNotFoundError("template.kml が見つかりません。")

        out_root, _ = prepare_output_dirs(path, params["offset"])
        original_angles = None

        for kml in kmls:
            log.insert(tk.END, f"Converting {os.path.basename(kml)}...\n")
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(kml, parser)

            # 元角度維持モードのため抽出
            if params["yaw_mode"] == "original" or params["gimbal_pitch_mode"] == "original":
                original_angles = extract_original_gimbal_angles(tree)
                log.insert(tk.END, f"Extracted original angles for {len(original_angles)} points\n")

            convert_kml(tree, params, original_angles)

            out_path = os.path.join(out_root, os.path.basename(kml))
            tree.write(out_path, encoding="utf-8", pretty_print=True, xml_declaration=True)

        # res フォルダのみコピー
        for name in ["res"]:
            srcs = glob.glob(os.path.join(wd, "**", name), recursive=True)
            if srcs:
                src = srcs[0]
                dst = os.path.join(out_root, name)
                if os.path.isdir(src):
                    if os.path.exists(dst): shutil.rmtree(dst)
                    shutil.copytree(src, dst)
                else:
                    shutil.copy2(src, dst)

        out_kmz = repackage_to_kmz(out_root, path)
        log.insert(tk.END, f"Saved: {out_kmz}\nFinished\n")
        messagebox.showinfo("完了", f"変換完了:\n{out_kmz}")

    except Exception as e:
        messagebox.showerror("エラー", str(e))
        log.insert(tk.END, f"Error: {e}\n")
    finally:
        if os.path.exists("_kmz_work"):
            shutil.rmtree("_kmz_work")

# --- GUI --------------------------------------------------------------------

class AppGUI(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        # 基準高度
        ttk.Label(self, text="基準高度:").grid(row=0, column=0, sticky="w")
        self.hc = ttk.Combobox(self, values=list(HEIGHT_OPTIONS), state="readonly", width=20)
        self.hc.set(next(iter(HEIGHT_OPTIONS)))
        self.hc.grid(row=0, column=1, columnspan=2, sticky="w")
        self.hc.bind("<<ComboboxSelected>>", self.on_height_change)
        self.he = ttk.Entry(self, width=10, state="disabled")
        self.he.grid(row=0, column=3)

        # 速度
        ttk.Label(self, text="速度 (1–15 m/s):").grid(row=1, column=0, sticky="w", pady=5)
        self.sp = tk.IntVar(value=15)
        ttk.Spinbox(self, from_=1, to=15, textvariable=self.sp, width=5).grid(row=1, column=1, columnspan=2, sticky="w")

        # 撮影設定
        self.ph = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="写真撮影", variable=self.ph).grid(row=2, column=0, sticky="w")
        self.vd = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="動画撮影", variable=self.vd).grid(row=2, column=1, sticky="w")
        self.vd_suffix_label = ttk.Label(self, text="動画ファイル名:")
        self.vd_suffix_var = tk.StringVar(value="video_01")
        self.vd_suffix_entry = ttk.Entry(self, textvariable=self.vd_suffix_var, width=20)

        # センサー選択
        ttk.Label(self, text="センサー選択:").grid(row=3, column=0, sticky="w")
        self.sm_vars = {m: tk.BooleanVar(value=False) for m in SENSOR_MODES}
        for i, m in enumerate(SENSOR_MODES):
            ttk.Checkbutton(self, text=m, variable=self.sm_vars[m]).grid(row=3, column=1+i, sticky="w")

        # ジンバルピッチ
        ttk.Label(self, text="ジンバルピッチ:").grid(row=4, column=0, sticky="w", pady=5)
        self.gp_cb = ttk.Combobox(self, values=list(GIMBAL_PITCH_OPTIONS), state="readonly", width=15)
        self.gp_cb.set(next(iter(GIMBAL_PITCH_OPTIONS)))
        self.gp_cb.grid(row=4, column=1, columnspan=2, sticky="w")
        self.gp_cb.bind("<<ComboboxSelected>>", self.update_gp_entry)
        self.gp_entry = ttk.Entry(self, width=8, state="disabled")

        # ヨー固定
        ttk.Label(self, text="ヨー固定:").grid(row=5, column=0, sticky="w", pady=5)
        self.yaw_cb = ttk.Combobox(self, values=list(YAW_OPTIONS), state="readonly", width=15)
        self.yaw_cb.set(next(iter(YAW_OPTIONS)))
        self.yaw_cb.grid(row=5, column=1, columnspan=2, sticky="w")
        self.yaw_cb.bind("<<ComboboxSelected>>", self.update_yaw_entry)
        self.yaw_entry = ttk.Entry(self, width=8, state="disabled")

        # ホバリング
        self.hv = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="ホバリング", variable=self.hv, command=self.update_hover).grid(row=6, column=0, sticky="w", pady=5)
        self.hover_time_label = ttk.Label(self, text="ホバリング時間 (秒):")
        self.hover_time_var = tk.StringVar(value="2")
        self.hover_time_entry = ttk.Entry(self, textvariable=self.hover_time_var, width=8)

        # 偏差補正
        self.dc = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="偏差補正", variable=self.dc, command=self.update_deviation).grid(row=7, column=0, sticky="w", pady=5)
        self.ref_point_label = ttk.Label(self, text="基準位置:")
        self.ref_point_var = tk.StringVar(value="本部")
        self.ref_point_combo = ttk.Combobox(self, values=list(REFERENCE_POINTS), state="readonly", width=10)
        self.today_coords_label = ttk.Label(self, text="本日の値 (経度,緯度,標高):")
        self.today_lng_var = tk.StringVar(value="136.555")
        self.today_lat_var = tk.StringVar(value="36.072")
        self.today_alt_var = tk.StringVar(value="0")
        self.today_lng_entry = ttk.Entry(self, textvariable=self.today_lng_var, width=12)
        self.today_lat_entry = ttk.Entry(self, textvariable=self.today_lat_var, width=10)
        self.today_alt_entry = ttk.Entry(self, textvariable=self.today_alt_var, width=8)
        self.copy_button = ttk.Button(self, text="コピー", command=self.copy_reference_data, width=8)

        # 初期 UI 設定
        self.update_ctrl()
        self.update_gp_entry()
        self.update_yaw_entry()
        self.update_hover()
        self.update_deviation()

    def on_height_change(self, event=None):
        if HEIGHT_OPTIONS[self.hc.get()] == "custom":
            self.he.config(state="normal")
        else:
            self.he.config(state="disabled")
            self.he.delete(0, tk.END)

    def update_ctrl(self):
        if self.vd.get():
            self.vd_suffix_label.grid(row=2, column=2, sticky="e")
            self.vd_suffix_entry.grid(row=2, column=3, sticky="w")
        else:
            self.vd_suffix_label.grid_forget()
            self.vd_suffix_entry.grid_forget()

    def update_gp_entry(self, event=None):
        mode = GIMBAL_PITCH_OPTIONS[self.gp_cb.get()]
        if mode == "custom":
            self.gp_entry.config(state="normal")
            self.gp_entry.grid(row=4, column=3)
        else:
            self.gp_entry.config(state="disabled")
            self.gp_entry.grid_forget()

    def update_yaw_entry(self, event=None):
        mode = YAW_OPTIONS[self.yaw_cb.get()]
        if mode == "custom":
            self.yaw_entry.config(state="normal")
            self.yaw_entry.grid(row=5, column=3)
        else:
            self.yaw_entry.config(state="disabled")
            self.yaw_entry.grid_forget()

    def update_hover(self):
        if self.hv.get():
            self.hover_time_label.grid(row=6, column=1, sticky="e")
            self.hover_time_entry.grid(row=6, column=2, sticky="w")
        else:
            self.hover_time_label.grid_forget()
            self.hover_time_entry.grid_forget()

    def update_deviation(self):
        if self.dc.get():
            self.ref_point_label.grid(row=7, column=1, sticky="e")
            self.ref_point_combo.grid(row=7, column=2, sticky="w")
            self.today_coords_label.grid(row=8, column=0, sticky="w")
            self.today_lng_entry.grid(row=8, column=1)
            self.today_lat_entry.grid(row=8, column=2)
            self.today_alt_entry.grid(row=8, column=3)
            self.copy_button.grid(row=8, column=4)
        else:
            self.ref_point_label.grid_forget()
            self.ref_point_combo.grid_forget()
            self.today_coords_label.grid_forget()
            self.today_lng_entry.grid_forget()
            self.today_lat_entry.grid_forget()
            self.today_alt_entry.grid_forget()
            self.copy_button.grid_forget()

    def copy_reference_data(self):
        try:
            rp = self.ref_point_combo.get()
            coords = REFERENCE_POINTS[rp]
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            txt = f"{now},{coords[0]},{coords[1]},{coords[2]}"
            pyperclip.copy(txt)
            messagebox.showinfo("コピー完了", f"{txt}")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def get_params(self):
        # 高度オフセット
        v = HEIGHT_OPTIONS[self.hc.get()]
        offset = float(self.he.get()) if v == "custom" else float(v)

        # 撮影モード
        do_photo = self.ph.get()
        do_video = self.vd.get()

        # 動画ファイル名
        video_suffix = self.vd_suffix_var.get()

        # ジンバルピッチモードと角度
        gp_mode_key = self.gp_cb.get()
        gp_mode = GIMBAL_PITCH_OPTIONS[gp_mode_key]
        if gp_mode == "custom":
            gp_angle = float(self.gp_entry.get())
        else:
            gp_angle = None

        # ヨーモードと角度
        yaw_mode_key = self.yaw_cb.get()
        yaw_mode = YAW_OPTIONS[yaw_mode_key]
        if yaw_mode == "custom":
            yaw_angle = float(self.yaw_entry.get())
        else:
            yaw_angle = None

        # 速度
        speed = max(1, min(15, self.sp.get()))

        # センサー
        sensor_modes = [m for m, var in self.sm_vars.items() if var.get()]

        # ホバリング
        try:
            hover_time = float(self.hover_time_var.get()) if self.hv.get() else 0
        except:
            hover_time = 0

        # 偏差補正
        coordinate_deviation = None
        if self.dc.get():
            rp = self.ref_point_combo.get()
            ref = REFERENCE_POINTS[rp]
            curr = (float(self.today_lng_entry.get()),
                    float(self.today_lat_entry.get()),
                    float(self.today_alt_entry.get()))
            dev = calculate_deviation(ref, curr)
            ok, msg = check_deviation_safety(dev)
            if not ok:
                messagebox.showerror("偏差補正エラー", msg)
                return None
            coordinate_deviation = dev

        return {
            "offset": offset,
            "do_photo": do_photo,
            "do_video": do_video,
            "video_suffix": video_suffix,
            "gimbal_pitch_mode": "original" if gp_mode == "original" else ("fixed" if gp_mode != "custom" else "fixed"),
            "gimbal_pitch_angle": gp_angle,
            "yaw_mode": "original" if yaw_mode == "original" else ("fixed" if yaw_mode != "custom" else "fixed"),
            "yaw_angle": yaw_angle,
            "speed": speed,
            "sensor_modes": sensor_modes,
            "hover_time": hover_time,
            "coordinate_deviation": coordinate_deviation
        }

# --- エントリポイント -------------------------------------------------------

def main():
    root = TkinterDnD.Tk()
    root.title("ATL→ASL 変換＋撮影制御ツール (ver. GUI51 – 元角度維持統合版)")
    root.geometry("820x820")

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)
    app = AppGUI(frm)
    app.pack(fill="x", pady=(0,10))

    drop = tk.Label(frm, text=".kmz をここにドロップ", bg="lightgray",
                    width=70, height=5, relief=tk.RIDGE)
    drop.pack(pady=12, fill="x")
    drop.drop_target_register(DND_FILES)

    log_frame = ttk.LabelFrame(frm, text="ログ")
    log_frame.pack(fill="both", expand=True)
    log = scrolledtext.ScrolledText(log_frame, height=16)
    log.pack(fill="both", expand=True)

    def on_drop(event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", ".kmz ファイルのみ対応")
            return
        params = app.get_params()
        if not params:
            return
        drop.config(state="disabled")
        threading.Thread(target=lambda: (process_kmz(path, params, log), drop.config(state="normal")), daemon=True).start()

    drop.dnd_bind("<<Drop>>", on_drop)
    root.mainloop()

if __name__ == "__main__":
    main()
