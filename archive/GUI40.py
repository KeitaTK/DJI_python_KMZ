#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_height_gui_asl_mod_v2.py  (改良版)

主な改善点：
1) アクション順序：ヨー→ピッチ→撮影に変更
2) actionId連番採番：動的に0,1,2...で自動採番
3) ExtendedData重複排除：同名タグの重複を防止
4) ASL高度計算：基準高度+ATLの正確な計算
5) センサ一意化：順序固定・重複排除
6) 詳細ログ：変換内容、実行順序、設定値を詳細表示
"""

import os
import shutil
import zipfile
import threading
import tkinter as tk
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

SENSOR_MODES = ["Wide", "Zoom", "IR"]

REFERENCE_POINTS = {
    "本部": (136.5559522506280, 36.0729517605894, 612.2),
    "烏帽子": (136.560000000000, 36.075000000000, 962.02)
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
    ref_lng, ref_lat, ref_alt = reference_coords
    cur_lng, cur_lat, cur_alt = current_coords
    return cur_lng - ref_lng, cur_lat - ref_lat, cur_alt - ref_alt

def check_deviation_safety(deviation):
    dev_lng, dev_lat, dev_alt = deviation
    if (
        abs(dev_lng) > DEVIATION_THRESHOLD["lng"]
        or abs(dev_lat) > DEVIATION_THRESHOLD["lat"]
        or abs(dev_alt) > DEVIATION_THRESHOLD["alt"]
    ):
        return False, (
            f"偏差が閾値を超えています:\n"
            f"経度偏差: {dev_lng:.8f}°\n"
            f"緯度偏差: {dev_lat:.8f}°\n"
            f"標高偏差: {dev_alt:.2f}m"
        )
    return True, None

def format_sensor_list(sensor_modes):
    """センサーリストを一意化・順序固定"""
    if not sensor_modes:
        return []
    # 重複排除＋順序固定
    unique_sensors = []
    for sensor in SENSOR_MODES:  # 定義順序で処理
        if sensor in sensor_modes:
            unique_sensors.append(sensor)
    return unique_sensors

def log_conversion_details(log, params, placemark_count):
    """変換内容の詳細ログ出力"""
    log.insert(tk.END, "=" * 60 + "\n")
    log.insert(tk.END, "変換設定詳細:\n")
    log.insert(tk.END, "=" * 60 + "\n")
    
    # ATL→ASL変換
    if params.get("do_asl", False):
        log.insert(tk.END, f"ATL→ASL変換: 有効 (オフセット: {params.get('offset', 0)}m)\n")
    else:
        log.insert(tk.END, "ATL→ASL変換: 無効\n")
    
    # 撮影モード
    if params.get("do_photo"):
        log.insert(tk.END, "撮影モード: 写真撮影\n")
    elif params.get("do_video"):
        log.insert(tk.END, "撮影モード: 動画撮影\n")
    else:
        log.insert(tk.END, "撮影モード: 設定なし\n")
    
    # ジンバルピッチ制御
    if params.get("gimbal_pitch_ctrl"):
        mode = params.get("gimbal_pitch_mode")
        angle = params.get("gimbal_pitch_angle")
        log.insert(tk.END, f"ジンバルピッチ制御: 有効 (モード: {mode}, 角度: {angle}°)\n")
    else:
        log.insert(tk.END, "ジンバルピッチ制御: 無効 (既存設定を維持)\n")
    
    # ヨー固定
    if params.get("yaw_angle") is not None:
        log.insert(tk.END, f"ヨー固定: 有効 ({params.get('yaw_angle')}°)\n")
    else:
        log.insert(tk.END, "ヨー固定: 無効\n")
    
    # センサー選択
    sensors = params.get("sensor_modes", [])
    if sensors:
        log.insert(tk.END, f"センサー選択: {', '.join(sensors)}\n")
    else:
        log.insert(tk.END, "センサー選択: 既存設定を維持\n")
    
    # 偏差補正
    if params.get("coordinate_deviation"):
        dev = params.get("coordinate_deviation")
        log.insert(tk.END, f"偏差補正: 有効 (経度: {dev[0]:.8f}°, 緯度: {dev[1]:.8f}°, 標高: {dev[2]:.2f}m)\n")
    else:
        log.insert(tk.END, "偏差補正: 無効\n")
    
    # アクション実行順序
    log.insert(tk.END, "\nアクション実行順序:\n")
    action_sequence = []
    if params.get("yaw_angle") is not None:
        action_sequence.append("1. ヨー回転 (rotateYaw)")
    if params.get("gimbal_pitch_ctrl") and params.get("gimbal_pitch_mode") != "free":
        action_sequence.append("2. ジンバルピッチ制御 (gimbalRotate)")
    if params.get("do_photo") or params.get("do_video"):
        action_sequence.append("3. 撮影実行 (orientedShoot)")
    
    if action_sequence:
        for seq in action_sequence:
            log.insert(tk.END, f"  {seq}\n")
    else:
        log.insert(tk.END, "  既存アクションを維持\n")
    
    log.insert(tk.END, f"\n処理対象ウェイポイント数: {placemark_count}\n")
    log.insert(tk.END, "=" * 60 + "\n\n")

# -------------------------------------------------------------------
# アクション生成・管理関数
# -------------------------------------------------------------------
def create_yaw_rotate_action(yaw_angle):
    """ヨー回転アクション生成"""
    action = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionId").text = "0"  # 仮ID、後で再採番
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "rotateYaw"
    
    param = etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}aircraftHeading").text = str(yaw_angle)
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}aircraftPathMode").text = "counterClockwise"
    
    return action

def create_gimbal_rotate_action(angle):
    """ジンバル回転アクション生成"""
    action = etree.Element("{http://www.dji.com/wpmz/1.0.6}action")
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionId").text = "0"  # 仮ID、後で再採番
    etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFunc").text = "gimbalRotate"

    param = etree.SubElement(action, "{http://www.dji.com/wpmz/1.0.6}actionActuatorFuncParam")
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalRotateMode").text = "absoluteAngle"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalPitchRotateEnable").text = "1"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalPitchRotateAngle").text = str(angle)
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalRollRotateEnable").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalRollRotateAngle").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateEnable").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalYawRotateAngle").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalRotateTimeEnable").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}gimbalRotateTime").text = "0"
    etree.SubElement(param, "{http://www.dji.com/wpmz/1.0.6}payloadPositionIndex").text = "0"

    return action

def reorganize_actions(action_group, params, log, waypoint_index):
    """アクションを指定順序で再編成"""
    # 既存アクションを分類
    existing_actions = {
        "rotateYaw": [],
        "gimbalRotate": [],
        "orientedShoot": [],
        "other": []
    }
    
    for action in action_group.findall("wpml:action", NS):
        func_elem = action.find("wpml:actionActuatorFunc", NS)
        if func_elem is not None:
            func_name = func_elem.text
            if func_name in existing_actions:
                existing_actions[func_name].append(action)
            else:
                existing_actions["other"].append(action)
    
    # 既存アクションをすべて削除
    for action in action_group.findall("wpml:action", NS):
        action_group.remove(action)
    
    # 新しい順序でアクションを追加
    new_actions = []
    action_log = []
    
    # 1. ヨー回転（必要な場合）
    yaw_angle = params.get("yaw_angle")
    if yaw_angle is not None:
        if existing_actions["rotateYaw"]:
            # 既存のrotateYawを更新
            yaw_action = existing_actions["rotateYaw"][0]
            param = yaw_action.find("wpml:actionActuatorFuncParam", NS)
            if param is not None:
                heading_elem = param.find("wpml:aircraftHeading", NS)
                if heading_elem is not None:
                    heading_elem.text = str(yaw_angle)
        else:
            # 新規作成
            yaw_action = create_yaw_rotate_action(yaw_angle)
        new_actions.append(yaw_action)
        action_log.append(f"ヨー回転: {yaw_angle}°")
    
    # 2. ジンバルピッチ制御（必要な場合）
    if params.get("gimbal_pitch_ctrl") and params.get("gimbal_pitch_mode") != "free":
        gimbal_angle = params.get("gimbal_pitch_angle")
        if gimbal_angle is not None:
            if existing_actions["gimbalRotate"]:
                # 既存のgimbalRotateを更新
                gimbal_action = existing_actions["gimbalRotate"][0]
                param = gimbal_action.find("wpml:actionActuatorFuncParam", NS)
                if param is not None:
                    pitch_elem = param.find("wpml:gimbalPitchRotateAngle", NS)
                    if pitch_elem is not None:
                        pitch_elem.text = str(gimbal_angle)
            else:
                # 新規作成
                gimbal_action = create_gimbal_rotate_action(gimbal_angle)
            new_actions.append(gimbal_action)
            action_log.append(f"ジンバルピッチ: {gimbal_angle}°")
            
            # orientedShootのピッチ角度も統一
            for shoot_action in existing_actions["orientedShoot"]:
                param = shoot_action.find("wpml:actionActuatorFuncParam", NS)
                if param is not None:
                    pitch_elem = param.find("wpml:gimbalPitchRotateAngle", NS)
                    if pitch_elem is not None:
                        pitch_elem.text = str(gimbal_angle)
    
    # 3. 撮影アクション
    for shoot_action in existing_actions["orientedShoot"]:
        new_actions.append(shoot_action)
        action_log.append("写真撮影")
    
    # 4. その他のアクション
    for other_action in existing_actions["other"]:
        new_actions.append(other_action)
        action_log.append("その他アクション")
    
    # アクションを追加してactionIdを再採番
    for idx, action in enumerate(new_actions):
        action_id_elem = action.find("wpml:actionId", NS)
        if action_id_elem is not None:
            action_id_elem.text = str(idx)
        action_group.append(action)
    
    # ログ出力
    if action_log:
        log.insert(tk.END, f"  ウェイポイント{waypoint_index}: {' → '.join(action_log)}\n")

def update_extended_data(placemark, params):
    """ExtendedDataを更新（重複排除）"""
    ed = placemark.find(".//kml:ExtendedData", NS)
    if ed is None:
        ed = etree.SubElement(placemark, "{http://www.opengis.net/kml/2.2}ExtendedData")
    
    # 既存のDataエレメントを辞書化（重複排除）
    existing_data = {}
    for data_elem in ed.findall("kml:Data", NS):
        name = data_elem.get("name")
        if name:
            existing_data[name] = data_elem
    
    # すべてのDataエレメントを削除
    for data_elem in ed.findall("kml:Data", NS):
        ed.remove(data_elem)
    
    # 撮影モード
    mode = "photo" if params.get("do_photo") else "video" if params.get("do_video") else ""
    if mode:
        existing_data["mode"] = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="mode")
        existing_data["mode"].text = mode
    
    # センサー選択（一意化・順序固定）
    sensor_modes = params.get("sensor_modes", [])
    if sensor_modes:
        unique_sensors = format_sensor_list(sensor_modes)
        if unique_sensors:
            existing_data["sensors"] = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="sensors")
            existing_data["sensors"].text = ",".join(unique_sensors)
    
    # ジンバル制御フラグ
    existing_data["gimbal"] = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="gimbal")
    existing_data["gimbal"].text = str(params.get("do_gimbal", True))
    
    # ヨー角
    yaw_angle = params.get("yaw_angle")
    if yaw_angle is not None:
        existing_data["yaw"] = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="yaw")
        existing_data["yaw"].text = str(yaw_angle)
    
    # ホバリング時間
    hover_time = params.get("hover_time", 0)
    if hover_time > 0:
        existing_data["hover_time"] = etree.SubElement(ed, "{http://www.opengis.net/kml/2.2}Data", name="hover_time")
        existing_data["hover_time"].text = str(hover_time)

# -------------------------------------------------------------------
# KML変換処理本体
# -------------------------------------------------------------------
def convert_kml(tree, params, log):
    offset = params.get("offset", 0.0)
    deviation = params.get("coordinate_deviation")
    do_asl = params.get("do_asl", False)
    
    hmode_elem = tree.find("./kml:Document/kml:Folder/wpml:waylineCoordinateSysParam/wpml:heightMode", NS)
    height_mode = hmode_elem.text if hmode_elem is not None else None
    
    # globalHeightの取得
    global_height_elem = tree.find(".//wpml:globalHeight", NS)
    original_global_height = float(global_height_elem.text) if global_height_elem is not None else 100.0
    
    log.insert(tk.END, f"変換処理開始:\n")
    log.insert(tk.END, f"  元のglobalHeight: {original_global_height}m\n")
    log.insert(tk.END, f"  heightMode: {height_mode}\n")
    
    placemark_count = 0
    log.insert(tk.END, f"\nウェイポイント処理:\n")
    
    for pm in tree.findall(".//kml:Placemark", NS):
        placemark_count += 1
        waypoint_index = pm.find("wpml:index", NS)
        wp_idx = int(waypoint_index.text) if waypoint_index is not None else placemark_count - 1
        
        # 座標・高度処理
        coords_elem = pm.find(".//kml:coordinates", NS)
        if coords_elem is None or not coords_elem.text:
            continue
            
        coords_text = coords_elem.text.strip().split(",")
        try:
            if len(coords_text) == 2:
                height_el = pm.find("wpml:height", NS)
                alt = float(height_el.text) if height_el is not None else 0.0
                lng, lat = map(float, coords_text)
            elif len(coords_text) == 3:
                lng, lat, alt = map(float, coords_text)
            else:
                continue
        except (ValueError, IndexError):
            continue
        
        original_alt = alt
        
        # 偏差補正
        if deviation is not None:
            dlng, dlat, dalt = deviation
            lng += dlng
            lat += dlat
            alt += dalt
        
        # ATL→ASL変換
        if do_asl and height_mode == "relativeToStartPoint":
            alt += offset
        
        # 座標更新
        coords_elem.text = f"{lng},{lat},{alt}"
        
        # height・ellipsoidHeight更新
        for tag in ("height", "ellipsoidHeight"):
            el = pm.find(f"wpml:{tag}", NS)
            if el is not None:
                el.text = f"{alt}"
        
        # アクション再編成
        action_group = pm.find(".//wpml:actionGroup", NS)
        if action_group is not None:
            reorganize_actions(action_group, params, log, wp_idx)
        
        # ExtendedData更新
        update_extended_data(pm, params)
        
        log.insert(tk.END, f"  ウェイポイント{wp_idx}: 高度 {original_alt:.1f}m → {alt:.1f}m\n")
    
    # globalHeight更新
    if global_height_elem is not None:
        new_global_height = original_global_height
        if do_asl and height_mode == "relativeToStartPoint":
            new_global_height += offset
        if deviation is not None:
            new_global_height += deviation[2]
        global_height_elem.text = str(new_global_height)
        log.insert(tk.END, f"\nglobalHeight更新: {original_global_height}m → {new_global_height}m\n")
    
    # heightMode更新
    if do_asl and height_mode == "relativeToStartPoint":
        for hm in tree.findall(".//wpml:heightMode", NS):
            hm.text = "EGM96"
        log.insert(tk.END, f"heightMode更新: {height_mode} → EGM96\n")
    
    log.insert(tk.END, f"\n処理完了: {placemark_count}個のウェイポイントを変換\n")
    return placemark_count

def process_kmz(path, log, params):
    try:
        log.insert(tk.END, f"処理開始: {os.path.basename(path)}...\n")
        
        # 変換設定の詳細ログ出力
        work_dir = extract_kmz(path)
        
        kml_path = os.path.join(work_dir, "wpmz", "template.kml")
        if not os.path.exists(kml_path):
            kml_path = os.path.join(work_dir, "template.kml")
        if not os.path.exists(kml_path):
            raise FileNotFoundError("template.kml が見つかりませんでした。")

        # 事前にPlacemark数を取得してログ出力
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(kml_path, parser)
        placemark_count = len(tree.findall(".//kml:Placemark", NS))
        
        log_conversion_details(log, params, placemark_count)
        
        out_root, outdir = prepare_output_dirs(path, params["offset"])
        log.insert(tk.END, f"変換中: {os.path.basename(kml_path)}...\n")

        # KML変換実行
        actual_count = convert_kml(tree, params, log)

        out_kml = os.path.join(outdir, os.path.basename(kml_path))
        tree.write(out_kml, encoding="utf-8", pretty_print=True, xml_declaration=True)

        # リソースフォルダのコピー
        res_src = os.path.join(os.path.dirname(kml_path), "res")
        if os.path.isdir(res_src):
            shutil.copytree(res_src, os.path.join(outdir, "res"))

        out_kmz = repackage_to_kmz(out_root, path)
        log.insert(tk.END, f"\n変換完了: {out_kmz}\n")
        log.insert(tk.END, "=" * 60 + "\n\n")
        messagebox.showinfo("完了", f"変換が正常に完了しました:\n{out_kmz}")
        
    except Exception as e:
        log.insert(tk.END, f"エラー: {e}\n\n")
        messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
    finally:
        if os.path.exists("_kmz_work"):
            shutil.rmtree("_kmz_work")

# -------------------------------------------------------------------
# View
# -------------------------------------------------------------------
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
        self.asl_var = tk.BooleanVar(value=False)

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

        # ジンバルピッチ制御
        self.gim_pitch_ctrl_var = tk.BooleanVar(value=False)
        self.gim_pitch_choice_var = tk.StringVar()
        self.gim_pitch_entry_var = tk.StringVar()

    def _create_widgets(self):
        # ASL
        self.asl_check = ttk.Checkbutton(
            self,
            text="ATL→ASL変換",
            variable=self.asl_var,
            command=self.controller.update_ui_states,
        )
        # 基準高度
        self.height_label = ttk.Label(self, text="基準高度:")
        self.height_combo = ttk.Combobox(
            self,
            textvariable=self.height_choice_var,
            values=list(HEIGHT_OPTIONS.keys()),
            state="readonly",
            width=20,
        )
        self.height_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.height_entry = ttk.Entry(self, textvariable=self.height_entry_var, width=10)

        # 速度
        self.speed_label = ttk.Label(self, text="速度 (1–15 m/s):")
        self.speed_spinbox = ttk.Spinbox(self, from_=1, to=15, textvariable=self.speed_var, width=5)

        # 撮影設定
        self.photo_check = ttk.Checkbutton(
            self, text="写真撮影", variable=self.photo_var, command=self.controller.update_ui_states
        )
        self.video_check = ttk.Checkbutton(
            self, text="動画撮影", variable=self.video_var, command=self.controller.update_ui_states
        )
        self.video_suffix_label = ttk.Label(self, text="動画ファイル名:")
        self.video_suffix_entry = ttk.Entry(self, textvariable=self.video_suffix_var, width=20)

        # センサー
        self.sensor_label = ttk.Label(self, text="センサー選択:")
        self.sensor_checks = {m: ttk.Checkbutton(self, text=m, variable=v) for m, v in self.sensor_vars.items()}

        # ジンバル有無
        self.gimbal_check = ttk.Checkbutton(self, text="ジンバル制御", variable=self.gimbal_var)

        # ジンバルピッチ制御
        self.gim_pitch_ctrl_check = ttk.Checkbutton(
            self, text="ジンバルピッチ制御", variable=self.gim_pitch_ctrl_var, command=self.controller.update_ui_states
        )
        self.gim_pitch_combo = ttk.Combobox(
            self,
            textvariable=self.gim_pitch_choice_var,
            values=["真下 (-90°)", "前 (0°)", "手動入力", "フリー"],
            state="readonly",
            width=15,
        )
        self.gim_pitch_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.gim_pitch_entry = ttk.Entry(self, textvariable=self.gim_pitch_entry_var, width=8)

        # ヨー固定
        self.yaw_fix_check = ttk.Checkbutton(
            self, text="ヨー固定", variable=self.yaw_fix_var, command=self.controller.update_ui_states
        )
        self.yaw_combo = ttk.Combobox(
            self, textvariable=self.yaw_choice_var, values=list(YAW_OPTIONS.keys()), state="readonly", width=15
        )
        self.yaw_combo.bind("<<ComboboxSelected>>", self.controller.update_ui_states)
        self.yaw_entry = ttk.Entry(self, textvariable=self.yaw_entry_var, width=8)

        # ホバリング
        self.hover_check = ttk.Checkbutton(
            self, text="ホバリング", variable=self.hover_var, command=self.controller.update_ui_states
        )
        self.hover_time_label = ttk.Label(self, text="ホバリング時間 (秒):")
        self.hover_time_entry = ttk.Entry(self, textvariable=self.hover_time_var, width=8)

        # 偏差補正
        self.deviation_check = ttk.Checkbutton(
            self, text="偏差補正", variable=self.deviation_var, command=self.controller.update_ui_states
        )
        self.ref_point_label = ttk.Label(self, text="基準位置:")
        self.ref_point_combo = ttk.Combobox(
            self, textvariable=self.ref_point_var, values=list(REFERENCE_POINTS.keys()), state="readonly", width=10
        )
        self.today_coords_label = ttk.Label(self, text="本日の値 (経度,緯度,標高):")
        self.today_lng_entry = ttk.Entry(self, textvariable=self.today_lng_var, width=12)
        self.today_lat_entry = ttk.Entry(self, textvariable=self.today_lat_var, width=10)
        self.today_alt_entry = ttk.Entry(self, textvariable=self.today_alt_var, width=8)
        self.copy_button = ttk.Button(self, text="コピー", command=self.controller.copy_reference_data, width=8)

    def _grid_widgets(self):
        # レイアウトは元のコードと同じ
        self.asl_check.grid(row=0, column=0, sticky="w", pady=5)
        self.speed_label.grid(row=1, column=0, sticky="w", pady=5)
        self.speed_spinbox.grid(row=1, column=1, columnspan=2, sticky="w")
        self.photo_check.grid(row=2, column=0, sticky="w")
        self.video_check.grid(row=2, column=1, sticky="w")
        self.sensor_label.grid(row=3, column=0, sticky="w")
        for i, m in enumerate(SENSOR_MODES):
            self.sensor_checks[m].grid(row=3, column=1 + i, sticky="w")
        self.gimbal_check.grid(row=4, column=0, sticky="w", pady=5)
        self.gim_pitch_ctrl_check.grid(row=5, column=0, sticky="w")
        self.yaw_fix_check.grid(row=6, column=0, sticky="w")
        self.hover_check.grid(row=7, column=0, sticky="w")
        self.deviation_check.grid(row=8, column=0, sticky="w")

# -------------------------------------------------------------------
# Controller
# -------------------------------------------------------------------
class KmlConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ATL→ASL 変換＋撮影制御ツール v2.0 (Enhanced)")
        self.root.geometry("820x800")
        self.ui = AppUI(self.root, self)
        self.ui.pack(fill="x", pady=(0, 10))
        self._create_dnd_and_log_area()
        self.update_ui_states()

    def _create_dnd_and_log_area(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill="both", expand=True)

        self.drop_label = tk.Label(frm, text=".kmz をここにドロップ", bg="lightgray", width=70, height=5, relief=tk.RIDGE)
        self.drop_label.pack(pady=12, fill="x")
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self.on_drop)

        log_frame = ttk.LabelFrame(frm, text="変換ログ")
        log_frame.pack(fill="both", expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=20)
        self.log_text.pack(fill="both", expand=True)

    def update_ui_states(self, event=None):
        ui = self.ui

        # ASL
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

        # 撮影モード排他
        if ui.photo_var.get():
            ui.video_var.set(False)
        if ui.video_var.get():
            ui.photo_var.set(False)

        # ジンバル制御チェック
        ui.gimbal_check.config(state="disabled" if ui.photo_var.get() or ui.video_var.get() else "normal")
        if ui.photo_var.get() or ui.video_var.get():
            ui.gimbal_var.set(True)

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
            messagebox.showwarning("警告", ".kmz ファイルのみ対応しています。")
            return

        try:
            params = self._get_params()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            return

        threading.Thread(target=process_kmz, args=(path, self.log_text, params), daemon=True).start()

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
            else:  # フリー
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
        p["sensor_modes"] = [m for m, var in ui.sensor_vars.items() if var.get()]

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
