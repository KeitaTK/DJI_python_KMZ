#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI44_imageFormat.py

ATL→ASL 変換＋撮影制御ツール v4.4
– グローバルヨー固定
– 写真時ジンバル前方固定
– 機体回転もヨー固定角度に設定
"""

import os, shutil, zipfile, threading, tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES
from lxml import etree
from datetime import datetime
import pyperclip

# -------------------------------------------------------------------
# 定数
# -------------------------------------------------------------------
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
IMAGE_FORMAT_OPTIONS = ["zoom", "wide", "ir"]
REFERENCE_POINTS = {
    "本部": (136.5559522506280, 36.0729517605894, 612.2),
    "烏帽子": (136.560000000000, 36.075000000000, 962.02),
}
DEVIATION_THRESHOLD = {"lat": 0.00018, "lng": 0.00022, "alt": 20.0}

# -------------------------------------------------------------------
# ユーティリティ関数
# -------------------------------------------------------------------
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
    rx, ry, rz = reference_coords
    cx, cy, cz = current_coords
    return cx - rx, cy - ry, cz - rz

def check_deviation_safety(deviation):
    dx, dy, dz = deviation
    if abs(dx) > DEVIATION_THRESHOLD["lng"] or abs(dy) > DEVIATION_THRESHOLD["lat"] or abs(dz) > DEVIATION_THRESHOLD["alt"]:
        return False, (f"偏差が閾値を超えています:\n"
                       f"経度偏差: {dx:.8f}°\n"
                       f"緯度偏差: {dy:.8f}°\n"
                       f"標高偏差: {dz:.2f}m")
    return True, None

def log_conversion_details(log, params, count):
    log.insert(tk.END, "="*60 + "\n")
    log.insert(tk.END, "変換設定詳細:\n")
    log.insert(tk.END, "="*60 + "\n")
    log.insert(tk.END, f"ATL→ASL変換: {'有効' if params['do_asl'] else '無効'}")
    if params['do_asl']:
        log.insert(tk.END, f" (オフセット: {params['offset']}m)")
    log.insert(tk.END, "\n")
    mode = "写真撮影" if params['do_photo'] else "動画撮影" if params['do_video'] else "設定なし"
    log.insert(tk.END, f"撮影モード: {mode}\n")
    fmts = ",".join(params['image_formats'])
    log.insert(tk.END, f"imageFormat: {fmts if fmts else '既存設定を維持'}\n")
    if params['gimbal_pitch_ctrl'] and params['gimbal_pitch_angle'] is not None:
        log.insert(tk.END, f"ジンバルピッチ制御: 有効 (角度: {params['gimbal_pitch_angle']}°)\n")
    else:
        log.insert(tk.END, "ジンバルピッチ制御: 無効\n")
    if params['yaw_angle'] is not None:
        log.insert(tk.END, f"ヨー固定: 有効 ({params['yaw_angle']}°)\n")
    else:
        log.insert(tk.END, "ヨー固定: 無効\n")
    if params['coordinate_deviation']:
        dx, dy, dz = params['coordinate_deviation']
        log.insert(tk.END, f"偏差補正: 有効 (経度: {dx:.8f}°, 緯度: {dy:.8f}°, 標高: {dz:.2f}m)\n")
    else:
        log.insert(tk.END, "偏差補正: 無効\n")
    log.insert(tk.END, f"ウェイポイント数: {count}\n")
    log.insert(tk.END, "="*60 + "\n\n")

# -------------------------------------------------------------------
# アクション生成
# -------------------------------------------------------------------
def create_yaw_rotate_action(yaw):
    a = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
    etree.SubElement(a, "{http://www.dji.com/wpmz/1.0.6}actionId").text = "0"
    etree.SubElement(a, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "rotateYaw"
    p = etree.SubElement(a, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
    etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}aircraftHeading").text = str(int(yaw))
    etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}aircraftPathMode").text = "counterClockwise"
    return a

def create_gimbal_yaw_fix_action():
    a = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
    etree.SubElement(a, "{http://www.dji.com/wpmz/1.0.6}actionId").text = "0"
    etree.SubElement(a, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "gimbalRotate"
    p = etree.SubElement(a, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
    etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalRotateMode").text = "absoluteAngle"
    etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalPitchRotateEnable").text = "0"
    etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalRollRotateEnable").text = "0"
    etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateEnable").text = "1"
    etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateAngle").text = "0"
    etree.SubElement(p, "{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex").text = "0"
    return a

def create_start_record_action():
    action = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionId").text = "0"
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "startRecord"
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
    return action

def create_stop_record_action():
    action = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionId").text = "0"
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "stopRecord"
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
    return action

def update_payload_param(tree, image_formats):
    if not image_formats:
        return
    fmt_str = ",".join(image_formats)
    payload = tree.find(".//wpml:payloadParam", NS)
    if payload is None:
        fld = tree.find(".//kml:Folder", NS)
        if fld is not None:
            payload = etree.SubElement(fld, "{http://www.dji.com/wpmz/1.0.6}payloadParam")
    if payload is not None:
        img = payload.find("wpml:imageFormat", NS)
        if img is None:
            img = etree.SubElement(payload, "{http://www.dji.com/wpmz/1.0.6}imageFormat")
        img.text = fmt_str

# -------------------------------------------------------------------
# アクション再編成
# -------------------------------------------------------------------
def reorganize_actions(ag, params, log, idx, first, last):
    existing = {"rotateYaw": [], "gimbalRotate": [], "orientedShoot": [], 
                "startRecord": [], "stopRecord": [], "other": []}
    for act in ag.findall("wpml:action", NS):
        fn = act.find("wpml:actionActuatorFunc", NS)
        if fn is not None and fn.text in existing:
            existing[fn.text].append(act)
        else:
            existing["other"].append(act)
    for act in ag.findall("wpml:action", NS):
        ag.remove(act)
    new_actions, logs = [], []

    # 1. グローバルヨー固定設定
    if params['yaw_angle'] is not None:
        # グローバル設定を探す
        root_element = ag
        while root_element.getparent() is not None:
            root_element = root_element.getparent()
        
        for gwp in root_element.findall(".//wpml:globalWaypointHeadingParam", NS):
            mode = gwp.find("wpml:waypointHeadingMode", NS)
            ang = gwp.find("wpml:waypointHeadingAngle", NS)
            if mode is not None: 
                mode.text = "fixed"
            if ang is not None: 
                ang.text = str(int(params['yaw_angle']))
        logs.append(f"グローバルヨー固定: {params['yaw_angle']}°")

    # 2. 写真時ジンバル前方固定
    if params['do_photo'] and params['yaw_angle'] is not None:
        gy = create_gimbal_yaw_fix_action()
        new_actions.append(gy)
        logs.append("ジンバルヨー前方固定")

    # 3. 機体回転
    if params['yaw_angle'] is not None:
        ry = create_yaw_rotate_action(params['yaw_angle'])
        new_actions.append(ry)
        logs.append(f"ヨー回転: {params['yaw_angle']}°")

    # 4. 既存撮影アクション
    if params['do_photo']:
        for s in existing['orientedShoot']:
            new_actions.append(s)
            logs.append("写真撮影")
    elif params['do_video']:
        if first:
            sr = existing['startRecord'][0] if existing['startRecord'] else create_start_record_action()
            new_actions.append(sr)
            logs.append("録画開始")
        if last:
            st = existing['stopRecord'][0] if existing['stopRecord'] else create_stop_record_action()
            new_actions.append(st)
            logs.append("録画停止")

    # 5. その他
    for o in existing['other']:
        new_actions.append(o)

    for i, act in enumerate(new_actions):
        aid = act.find("wpml:actionId", NS)
        if aid is not None: 
            aid.text = str(i)
        ag.append(act)

    if logs:
        log.insert(tk.END, f" ウェイポイント{idx}: " + " → ".join(logs) + "\n")

# -------------------------------------------------------------------
# KML変換処理
# -------------------------------------------------------------------
def convert_kml(tree, params, log):
    offset, dev = params['offset'], params['coordinate_deviation']
    do_asl = params['do_asl']
    hme = tree.find("./kml:Document/kml:Folder/wpml:waylineCoordinateSysParam/wpml:heightMode", NS)
    hm = hme.text if hme is not None else ""
    ghe = tree.find(".//wpml:globalHeight", NS)
    gh0 = float(ghe.text) if ghe is not None else 0.0

    if params['image_formats']:
        update_payload_param(tree, params['image_formats'])
        log.insert(tk.END, f"imageFormat更新: {','.join(params['image_formats'])}\n")

    log.insert(tk.END, f"変換開始: globalHeight={gh0}, heightMode={hm}\n")
    pms = tree.findall(".//kml:Placemark", NS)
    log.insert(tk.END, f"ウェイポイント処理: {len(pms)}箇所\n")

    for i, pm in enumerate(pms):
        idx_elem = pm.find("wpml:index", NS)
        idx = int(idx_elem.text) if idx_elem is not None else i
        first, last = (i==0), (i==len(pms)-1)
        coords = pm.find(".//kml:coordinates", NS)
        if not coords or not coords.text: 
            continue
        parts = list(map(float, coords.text.strip().split(",")))
        lng, lat = parts[0], parts[1]
        alt = parts[2] if len(parts)>2 else float(pm.find("wpml:height", NS).text or 0.0)
        orig_alt = alt
        if dev:
            lng += dev[0]; lat += dev[1]; alt += dev[2]
        if do_asl and hm=="relativeToStartPoint":
            alt += offset
        coords.text = f"{lng},{lat},{alt}"
        for tag in ("height","ellipsoidHeight"):
            e = pm.find(f"wpml:{tag}", NS)
            if e is not None: 
                e.text = str(alt)
        ag = pm.find(".//wpml:actionGroup", NS)
        if ag is not None:
            reorganize_actions(ag, params, log, idx, first, last)
        update_extended_data(pm, params)
        log.insert(tk.END, f" ウェイポイント{idx}: 高度 {orig_alt:.1f}→{alt:.1f}m\n")

    gh1 = gh0 + (offset if do_asl else 0) + (dev[2] if dev else 0)
    if ghe is not None: 
        ghe.text = str(gh1)
    if do_asl and hm=="relativeToStartPoint":
        for hme2 in tree.findall(".//wpml:heightMode", NS):
            hme2.text = "EGM96"
    log.insert(tk.END, f"globalHeight: {gh0}→{gh1}\n完了\n")

# -------------------------------------------------------------------
# ExtendedData 更新
# -------------------------------------------------------------------
def update_extended_data(pm, params):
    ed = pm.find(".//kml:ExtendedData", NS)
    if ed is None:
        ed = etree.SubElement(pm, "{http://www.opengis.net/kml/2.2}ExtendedData")
    for d in ed.findall("kml:Data", NS):
        ed.remove(d)
    mode = "photo" if params['do_photo'] else "video" if params['do_video'] else ""
    if mode:
        n = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="mode")
        n.text = mode
    gb = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="gimbal")
    gb.text = str(params['do_gimbal'])
    ya = params['yaw_angle']
    if ya is not None:
        yel = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="yaw")
        yel.text = str(ya)
    if params['hover_time']>0:
        hel = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="hover_time")
        hel.text = str(params['hover_time'])

# -------------------------------------------------------------------
# KMZ 全体処理
# -------------------------------------------------------------------
def process_kmz(path, log, params):
    try:
        log.insert(tk.END, f"処理開始: {os.path.basename(path)}\n")
        wd = extract_kmz(path)
        kml = os.path.join(wd, "wpmz", "template.kml")
        if not os.path.exists(kml):
            kml = os.path.join(wd, "template.kml")
        tree = etree.parse(kml, etree.XMLParser(remove_blank_text=True))
        cnt = len(tree.findall(".//kml:Placemark", NS))
        log_conversion_details(log, params, cnt)
        out_root, outdir = prepare_output_dirs(path, params['offset'])
        convert_kml(tree, params, log)
        out_kml = os.path.join(outdir, os.path.basename(kml))
        tree.write(out_kml, encoding="utf-8", pretty_print=True, xml_declaration=True)
        res = os.path.join(os.path.dirname(kml), "res")
        if os.path.isdir(res):
            shutil.copytree(res, os.path.join(outdir, "res"))
        outf = repackage_to_kmz(out_root, path)
        log.insert(tk.END, f"変換完了: {outf}\n")
        messagebox.showinfo("完了", f"変換完了:\n{outf}")
    except Exception as e:
        log.insert(tk.END, f"エラー: {e}\n")
        messagebox.showerror("エラー", str(e))
    finally:
        if os.path.exists("_kmz_work"):
            shutil.rmtree("_kmz_work")

# -------------------------------------------------------------------
# UI Classes
# -------------------------------------------------------------------
class AppUI(ttk.Frame):
    def __init__(self, master, ctrl):
        super().__init__(master, padding=10)
        self.controller = ctrl
        self._create_vars()
        self._create_widgets()
        self._grid_widgets()

    def _create_vars(self):
        self.height_choice_var = tk.StringVar()
        self.height_entry_var = tk.StringVar()
        self.asl_var = tk.BooleanVar(value=False)
        self.speed_var = tk.IntVar(value=15)
        self.photo_var = tk.BooleanVar(value=False)
        self.video_var = tk.BooleanVar(value=False)
        self.video_suffix_var = tk.StringVar(value="video_01")  # 修正
        self.image_format_vars = {f: tk.BooleanVar() for f in IMAGE_FORMAT_OPTIONS}
        self.gimbal_var = tk.BooleanVar(value=True)
        self.gim_pitch_ctrl_var = tk.BooleanVar(value=False)
        self.gim_pitch_choice_var = tk.StringVar()
        self.gim_pitch_entry_var = tk.StringVar()
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
        # ASL
        self.asl_check = ttk.Checkbutton(self, text="ATL→ASL変換", variable=self.asl_var,
                                         command=self.controller.update_ui_states)
        self.height_label = ttk.Label(self, text="基準高度:")
        self.height_combo = ttk.Combobox(self, textvariable=self.height_choice_var,
                                         values=list(HEIGHT_OPTIONS.keys()), state="readonly", width=20)
        self.height_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.height_entry = ttk.Entry(self, textvariable=self.height_entry_var, width=10)
        # 速度
        self.speed_label = ttk.Label(self, text="速度 (1–15 m/s):")
        self.speed_spin = ttk.Spinbox(self, from_=1, to=15, textvariable=self.speed_var, width=5)
        # 撮影
        self.photo_check = ttk.Checkbutton(self, text="写真撮影", variable=self.photo_var,
                                           command=self.controller.update_ui_states)
        self.video_check = ttk.Checkbutton(self, text="動画撮影", variable=self.video_var,
                                           command=self.controller.update_ui_states)
        self.video_suffix_label = ttk.Label(self, text="動画ファイル名:")
        self.video_suffix_entry = ttk.Entry(self, textvariable=self.video_suffix_var, width=20)
        # imageFormat
        self.image_format_label = ttk.Label(self, text="imageFormat:")
        self.image_format_checks = {
            f: ttk.Checkbutton(self, text=f, variable=self.image_format_vars[f])
            for f in IMAGE_FORMAT_OPTIONS
        }
        # ジンバル
        self.gimbal_check = ttk.Checkbutton(self, text="ジンバル制御", variable=self.gimbal_var)
        self.gim_pitch_ctrl = ttk.Checkbutton(self, text="ジンバルピッチ制御",
                                              variable=self.gim_pitch_ctrl_var,
                                              command=self.controller.update_ui_states)
        self.gim_pitch_combo = ttk.Combobox(self, textvariable=self.gim_pitch_choice_var,
                                            values=["真下 (-90°)", "前 (0°)", "手動入力", "フリー"],
                                            state="readonly", width=15)
        self.gim_pitch_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.gim_pitch_entry = ttk.Entry(self, textvariable=self.gim_pitch_entry_var, width=8)
        # ヨー固定
        self.yaw_fix_check = ttk.Checkbutton(self, text="ヨー固定", variable=self.yaw_fix_var,
                                             command=self.controller.update_ui_states)
        self.yaw_combo = ttk.Combobox(self, textvariable=self.yaw_choice_var,
                                      values=list(YAW_OPTIONS.keys()), state="readonly", width=15)
        self.yaw_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.yaw_entry = ttk.Entry(self, textvariable=self.yaw_entry_var, width=8)
        # ホバリング
        self.hover_check = ttk.Checkbutton(self, text="ホバリング", variable=self.hover_var,
                                           command=self.controller.update_ui_states)
        self.hover_time_label = ttk.Label(self, text="ホバリング時間 (秒):")
        self.hover_time_entry = ttk.Entry(self, textvariable=self.hover_time_var, width=8)
        # 偏差補正
        self.deviation_check = ttk.Checkbutton(self, text="偏差補正", variable=self.deviation_var,
                                               command=self.controller.update_ui_states)
        self.ref_point_label = ttk.Label(self, text="基準位置:")
        self.ref_point_combo = ttk.Combobox(self, textvariable=self.ref_point_var,
                                            values=list(REFERENCE_POINTS.keys()), state="readonly", width=10)
        self.today_coords_label = ttk.Label(self, text="本日の値 (経度,緯度,標高):")
        self.today_lng_entry = ttk.Entry(self, textvariable=self.today_lng_var, width=12)
        self.today_lat_entry = ttk.Entry(self, textvariable=self.today_lat_var, width=10)
        self.today_alt_entry = ttk.Entry(self, textvariable=self.today_alt_var, width=8)
        self.copy_button = ttk.Button(self, text="コピー",
                                      command=self.controller.copy_reference_data, width=8)

    def _grid_widgets(self):
        self.asl_check.grid(row=0, column=0, sticky="w", pady=5)
        self.speed_label.grid(row=1, column=0, sticky="w", pady=5)
        self.speed_spin.grid(row=1, column=1, columnspan=2, sticky="w")
        self.photo_check.grid(row=2, column=0, sticky="w")
        self.video_check.grid(row=2, column=1, sticky="w")
        self.image_format_label.grid(row=3, column=0, sticky="w")
        for i, f in enumerate(IMAGE_FORMAT_OPTIONS):
            self.image_format_checks[f].grid(row=3, column=1+i, sticky="w", padx=5)
        self.gimbal_check.grid(row=4, column=0, sticky="w", pady=5)
        self.gim_pitch_ctrl.grid(row=5, column=0, sticky="w")
        self.yaw_fix_check.grid(row=6, column=0, sticky="w")
        self.hover_check.grid(row=7, column=0, sticky="w")
        self.deviation_check.grid(row=8, column=0, sticky="w")

# -------------------------------------------------------------------
# Controller
# -------------------------------------------------------------------
class KmlConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ATL→ASL 変換＋撮影制御ツール v4.4")
        self.root.geometry("820x800")
        self.ui = AppUI(root, self)
        self.ui.pack(fill="x", pady=(0,10))
        self._create_dnd_and_log_area()
        self.update_ui_states()

    def _create_dnd_and_log_area(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill="both", expand=True)
        self.drop_label = tk.Label(frm, text=".kmz をここにドロップ",
                                    bg="lightgray", width=70, height=5, relief=tk.RIDGE)
        self.drop_label.pack(pady=12, fill="x")
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self.on_drop)
        log_frame = ttk.LabelFrame(frm, text="変換ログ")
        log_frame.pack(fill="both", expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=20)
        self.log_text.pack(fill="both", expand=True)

    def update_ui_states(self, event=None):
        ui = self.ui
        # ASL設定
        if ui.asl_var.get():
            ui.height_label.grid(row=0, column=2, sticky="w")
            ui.height_combo.grid(row=0, column=3, padx=5, columnspan=2, sticky="w")
            if HEIGHT_OPTIONS.get(ui.height_choice_var.get()) == "custom":
                ui.height_entry.config(state="normal")
                ui.height_entry.grid(row=0, column=5, padx=5)
            else:
                ui.height_entry.config(state="disabled") 
                ui.height_entry.grid_forget()
        else:
            ui.height_label.grid_forget()
            ui.height_combo.grid_forget()
            ui.height_entry.grid_forget()
        
        # 撮影モード排他制御
        if ui.photo_var.get():
            ui.video_var.set(False)
        if ui.video_var.get():
            ui.photo_var.set(False)
        
        # 動画ファイル名
        if ui.video_var.get():
            ui.video_suffix_label.grid(row=2, column=2, sticky="e", padx=(10, 2))
            ui.video_suffix_entry.grid(row=2, column=3, sticky="w")
        else:
            ui.video_suffix_label.grid_forget()
            ui.video_suffix_entry.grid_forget()
        
        # ジンバルピッチ制御
        if ui.gim_pitch_ctrl_var.get():
            ui.gim_pitch_combo.grid(row=5, column=1, padx=5, sticky="w")
            if not ui.gim_pitch_choice_var.get():
                ui.gim_pitch_choice_var.set("真下 (-90°)")
            if ui.gim_pitch_choice_var.get() == "手動入力":
                ui.gim_pitch_entry.config(state="normal")
                ui.gim_pitch_entry.grid(row=5, column=2)
            else:
                ui.gim_pitch_entry.config(state="disabled")
                ui.gim_pitch_entry.grid_forget()
        else:
            ui.gim_pitch_combo.grid_forget()
            ui.gim_pitch_entry.grid_forget()
        
        # ヨー固定
        if ui.yaw_fix_var.get():
            ui.yaw_combo.grid(row=6, column=1, padx=5, columnspan=2, sticky="w")
            if not ui.yaw_choice_var.get():
                ui.yaw_choice_var.set(next(iter(YAW_OPTIONS.keys())))
            if ui.yaw_choice_var.get() == "手動入力":
                ui.yaw_entry.config(state="normal")
                ui.yaw_entry.grid(row=6, column=3)
            else:
                ui.yaw_entry.config(state="disabled")
                ui.yaw_entry.grid_forget()
        else:
            ui.yaw_combo.grid_forget()
            ui.yaw_entry.grid_forget()
        
        # ホバリング
        if ui.hover_var.get():
            ui.hover_time_label.grid(row=7, column=1, sticky="e", padx=(10, 2))
            ui.hover_time_entry.grid(row=7, column=2, sticky="w")
        else:
            ui.hover_time_label.grid_forget()
            ui.hover_time_entry.grid_forget()
        
        # 偏差補正
        if ui.deviation_var.get():
            ui.ref_point_label.grid(row=8, column=1, sticky="e", padx=(10, 2))
            ui.ref_point_combo.grid(row=8, column=2, sticky="w")
            ui.today_coords_label.grid(row=9, column=0, sticky="w", pady=5)
            ui.today_lng_entry.grid(row=9, column=1, padx=2, sticky="w")
            ui.today_lat_entry.grid(row=9, column=2, padx=2, sticky="w")
            ui.today_alt_entry.grid(row=9, column=3, padx=2, sticky="w")
            ui.copy_button.grid(row=9, column=4, padx=5, sticky="w")
        else:
            ui.ref_point_label.grid_forget()
            ui.ref_point_combo.grid_forget()
            ui.today_coords_label.grid_forget()
            ui.today_lng_entry.grid_forget()
            ui.today_lat_entry.grid_forget()
            ui.today_alt_entry.grid_forget()
            ui.copy_button.grid_forget()

    def on_drop(self, event):
        path = event.data.strip("{}")
        if not path.lower().endswith(".kmz"):
            messagebox.showwarning("警告", ".kmzファイルのみ対応します")
            return
        try:
            params = self._get_params()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            return
        threading.Thread(target=process_kmz,
                         args=(path, self.log_text, params),
                         daemon=True).start()

    def copy_reference_data(self):
        try:
            ref = self.ui.ref_point_var.get()
            coords = REFERENCE_POINTS[ref]
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            txt = f"{now},{coords[0]},{coords[1]},{coords[2]}"
            pyperclip.copy(txt)
            messagebox.showinfo("コピー完了", f"コピーしました:\n{txt}")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def _get_params(self):
        ui = self.ui
        p = {}
        # ASL
        if ui.asl_var.get():
            v = HEIGHT_OPTIONS.get(ui.height_choice_var.get())
            if v == "custom":
                try:
                    p["offset"] = float(ui.height_entry_var.get())
                except ValueError:
                    raise ValueError("基準高度（手動入力）には数値を入力してください。")
            else:
                p["offset"] = float(v)
        else:
            p["offset"] = 0.0
        p["do_asl"] = ui.asl_var.get()
        
        # imageFormat
        p["image_formats"] = [f for f, var in ui.image_format_vars.items() if var.get()]
        
        # ジンバルピッチ制御
        p["gimbal_pitch_ctrl"] = ui.gim_pitch_ctrl_var.get()
        if ui.gim_pitch_ctrl_var.get():
            choice = ui.gim_pitch_choice_var.get()
            if choice.startswith("真下"):
                p["gimbal_pitch_mode"] = "-90"
                p["gimbal_pitch_angle"] = -90.0
            elif choice.startswith("前"):
                p["gimbal_pitch_mode"] = "0"
                p["gimbal_pitch_angle"] = 0.0
            elif choice == "手動入力":
                p["gimbal_pitch_mode"] = "manual"
                try:
                    p["gimbal_pitch_angle"] = float(ui.gim_pitch_entry_var.get())
                except ValueError:
                    raise ValueError("ジンバルピッチ角（手動入力）には数値を入力してください。")
            else: # フリー
                p["gimbal_pitch_mode"] = "free"
                p["gimbal_pitch_angle"] = None
        else:
            p["gimbal_pitch_ctrl"] = False
            p["gimbal_pitch_mode"] = "free"
            p["gimbal_pitch_angle"] = None
        
        # ヨー固定
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
            # 写真撮影モードかつヨー固定時にジンバルヨーも固定するフラグ
            p["gimbal_yaw_fixed"] = ui.photo_var.get()
        else:
            p["gimbal_yaw_fixed"] = False
        
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
                    float(ui.today_alt_var.get()),
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
        p["speed"] = max(1, min(15, ui.speed_var.get()))
        return p

# -------------------------------------------------------------------
# main
# -------------------------------------------------------------------
def main():
    root = TkinterDnD.Tk()
    app = KmlConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
